'use strict';

/**
 * PDFs protected with a password: the Standard security handler of ISO
 * 32000, read and written with node's own crypto (RC4 is small enough to
 * be written here, and OpenSSL 3 no longer offers it by default).
 *
 *   V1 and V2, revisions 2 and 3: RC4 with a key of 40 to 128 bits
 *   V4, revision 4: crypt filters (StdCF and its kin): RC4 (/V2), AES-128
 *     (/AESV2) or none (/Identity), for strings and streams apart
 *   V5, revisions 5 and 6: AES-256 (/AESV3); revision 6's password hash is
 *     the loop of SHA-256, SHA-384 and SHA-512 over AES-128
 *
 * A password opens a file as its user (the one that opens it at all) or
 * as its owner (the one that lifts its permissions); a file with an empty
 * user password opens without one, its permissions still to be kept.
 * Strings and streams are encrypted object by object: for RC4 and AES-128
 * the key is the file's key with the object's number and generation, for
 * AES-256 the file's key itself. What is never encrypted stays as it is:
 * the Encrypt dictionary, cross-reference streams, a signature's
 * /Contents, and the document's XMP metadata when EncryptMetadata is false.
 *
 * Nothing here keeps the password: it makes the file's key and is let go.
 * The handler keeps the key, for reading the file and for writing into it
 * (an update, or a new file under the same protection, as Save writes).
 * Other handlers (a certificate's, a plug-in's) are recognised and refused.
 */

const crypto = require('crypto');

class SecurityError extends Error {
  constructor(message) {
    super(message);
    this.name = 'PdfSecurityError';
    this.code = 'encrypted';
  }
}

const PAD = Buffer.from('28bf4e5e4e758a4164004e56fffa01082e2e00b6d0683e802f0ca9fe6453697a', 'hex');
const EMPTY = Buffer.alloc(0);

/** RC4, the stream cipher revisions 2 to 4 use. */
function rc4(key, data) {
  const s = new Uint8Array(256);
  for (let i = 0; i < 256; i++) s[i] = i;
  for (let i = 0, j = 0; i < 256; i++) {
    j = (j + s[i] + key[i % key.length]) & 255;
    const t = s[i]; s[i] = s[j]; s[j] = t;
  }
  const out = Buffer.alloc(data.length);
  for (let k = 0, i = 0, j = 0; k < data.length; k++) {
    i = (i + 1) & 255;
    j = (j + s[i]) & 255;
    const t = s[i]; s[i] = s[j]; s[j] = t;
    out[k] = data[k] ^ s[(s[i] + s[j]) & 255];
  }
  return out;
}

const md5 = (...parts) => crypto.createHash('md5').update(Buffer.concat(parts)).digest();
const xorKey = (key, i) => Buffer.from(key.map((b) => b ^ i));
const same = (a, b) => a.length === b.length && crypto.timingSafeEqual(a, b);

/** A password for revisions 2 to 4: PDFDocEncoding, which Latin-1 is for what people type, padded to 32 bytes. */
function padded(password) {
  const bytes = Buffer.from(String(password).replace(/[^\x00-\xff]/g, '?'), 'latin1').subarray(0, 32);
  return Buffer.concat([bytes, PAD]).subarray(0, 32);
}

/** A password for revisions 5 and 6: UTF-8 of its normalised form, at most 127 bytes. */
const utf8 = (password) => Buffer.from(String(password).normalize('NFKC'), 'utf8').subarray(0, 127);

/** The hash revision 5 (SHA-256) and revision 6 (ISO 32000-2, Algorithm 2.B) check a password with. */
function hash5(revision, pw, salt, udata = EMPTY) {
  let k = crypto.createHash('sha256').update(Buffer.concat([pw, salt, udata])).digest();
  if (revision < 6) return k;
  let e = EMPTY;
  for (let i = 0; i < 64 || e[e.length - 1] > i - 32; i++) {
    const block = Buffer.concat([pw, k, udata]);
    const k1 = Buffer.alloc(block.length * 64);
    for (let r = 0; r < 64; r++) block.copy(k1, r * block.length);
    const c = crypto.createCipheriv('aes-128-cbc', k.subarray(0, 16), k.subarray(16, 32));
    c.setAutoPadding(false);
    e = Buffer.concat([c.update(k1), c.final()]);
    // The first 16 bytes of E as a number, modulo 3: 256 is 1 modulo 3, so the sum of the bytes is.
    let sum = 0;
    for (let b = 0; b < 16; b++) sum += e[b];
    k = crypto.createHash(['sha256', 'sha384', 'sha512'][sum % 3]).update(e).digest();
  }
  return k.subarray(0, 32);
}

function aesDecrypt(bits, key, data) {
  // A string or stream too short to hold its starting block is left as it is: a writer's slip, not a secret.
  if (data.length < 16) return data;
  const iv = data.subarray(0, 16);
  const body = data.subarray(16);
  if (!body.length) return EMPTY;
  const whole = body.length - (body.length % 16);
  try {
    const d = crypto.createDecipheriv(`aes-${bits}-cbc`, key, iv);
    return Buffer.concat([d.update(body.subarray(0, whole)), d.final()]);
  } catch {
    // Padding that does not check out: what decrypts, as other readers give it.
    const d = crypto.createDecipheriv(`aes-${bits}-cbc`, key, iv);
    d.setAutoPadding(false);
    return Buffer.concat([d.update(body.subarray(0, whole)), d.final()]);
  }
}

function aesEncrypt(bits, key, data) {
  const iv = crypto.randomBytes(16);
  const c = crypto.createCipheriv(`aes-${bits}-cbc`, key, iv);
  return Buffer.concat([iv, c.update(data), c.final()]);
}

/** The bytes a key-wrapping entry (OE, UE) holds, unwrapped: AES-256 with a zero starting block and no padding. */
function unwrap(key, data) {
  const d = crypto.createDecipheriv('aes-256-cbc', key, Buffer.alloc(16));
  d.setAutoPadding(false);
  return Buffer.concat([d.update(data.subarray(0, 32)), d.final()]);
}

// What each permission bit allows (ISO 32000, table 22), numbered from 1.
const BITS = { print: 3, change: 4, copy: 5, comment: 6, fill: 9, assemble: 11 };

class StandardSecurity {
  /**
   * @param {object} dict the Encrypt dictionary (the reader's Dict, its values resolved)
   * @param {Buffer} id the first element of the file's /ID
   */
  constructor(dict, id) {
    const get = (key) => dict.get(key);
    const name = (v) => (v && typeof v === 'object' && 'name' in v ? v.name : '');
    const bytes = (v) => (v && Buffer.isBuffer(v.bytes) ? v.bytes : EMPTY);
    if (name(get('Filter')) !== 'Standard') {
      throw new SecurityError(`This PDF is protected by ${name(get('Filter')) === 'Adobe.PubSec' ? 'a certificate' : `the ${name(get('Filter')) || 'unnamed'} handler`}, not a password, and cannot be opened here.`);
    }
    this.v = Number(get('V')) || 0;
    this.r = Number(get('R')) || 0;
    if (![1, 2, 4, 5].includes(this.v) || ![2, 3, 4, 5, 6].includes(this.r)) {
      throw new SecurityError(`This PDF is protected in a way this reader does not know (version ${this.v}, revision ${this.r}).`);
    }
    this.dict = dict;
    this.id = id || EMPTY;
    this.O = bytes(get('O'));
    this.U = bytes(get('U'));
    this.OE = bytes(get('OE'));
    this.UE = bytes(get('UE'));
    this.P = Number(get('P')) | 0;
    this.encryptMetadata = get('EncryptMetadata') !== false;
    // How strings and streams are each encrypted: RC4 ('V2'), AES-128, AES-256, or not ('None').
    if (this.v >= 4) {
      const filters = get('CF');
      const method = (which) => {
        const n = name(get(which)) || 'Identity';
        if (n === 'Identity') return { method: 'None' };
        const cf = filters && filters.get ? filters.get(n) : null;
        const cfm = cf && cf.get ? name(cf.get('CFM')) : '';
        if (!cfm || cfm === 'None') return { method: 'None' };
        if (!['V2', 'AESV2', 'AESV3'].includes(cfm)) throw new SecurityError(`This PDF is protected with a cipher this reader does not know (${cfm}).`);
        return { method: cfm, length: cf.get('Length') };
      };
      this.stm = method('StmF');
      this.str = method('StrF');
      this.filters = filters;
      this.method = method;
    } else {
      this.stm = { method: 'V2' };
      this.str = { method: 'V2' };
    }
    // The file key's length in bytes.
    if (this.r >= 5) this.n = 32;
    else if (this.r === 2) this.n = 5;
    else {
      const aes = this.stm.method === 'AESV2' || this.str.method === 'AESV2';
      let bits = Number(get('Length')) || (this.v === 4 ? 128 : 40);
      if (this.v === 4) {
        const cfLength = Number((this.stm.method !== 'None' ? this.stm : this.str).length) || 0;
        if (cfLength) bits = cfLength <= 32 ? cfLength * 8 : cfLength; // in bytes as Acrobat writes it, or in bits as the standard says
      }
      this.n = aes ? 16 : Math.max(5, Math.min(16, Math.floor(bits / 8)));
    }
    this.key = null;
    this.access = null;
  }

  // ── the passwords ───────────────────────────────────────────────────────

  /** Algorithm 2: the file key from a (padded) user password, revisions 2 to 4. */
  keyFor(paddedUser) {
    const p = Buffer.alloc(4);
    p.writeInt32LE(this.P);
    const extra = this.r >= 4 && !this.encryptMetadata ? Buffer.from([0xff, 0xff, 0xff, 0xff]) : EMPTY;
    let hash = md5(paddedUser, this.O.subarray(0, 32), p, this.id, extra);
    if (this.r >= 3) for (let i = 0; i < 50; i++) hash = md5(hash.subarray(0, this.n));
    return hash.subarray(0, this.n);
  }

  /** Algorithms 4 and 5: whether a key made from a user password makes the file's /U. */
  userKey(paddedUser) {
    const key = this.keyFor(paddedUser);
    if (this.r === 2) return same(rc4(key, PAD), this.U.subarray(0, 32)) ? key : null;
    let u = rc4(key, md5(PAD, this.id));
    for (let i = 1; i <= 19; i++) u = rc4(xorKey(key, i), u);
    return same(u, this.U.subarray(0, 16)) ? key : null;
  }

  /** Algorithm 7: the user password the owner password unlocks from /O, padded. */
  userFromOwner(owner) {
    let hash = md5(padded(owner));
    if (this.r >= 3) for (let i = 0; i < 50; i++) hash = md5(hash);
    const key = hash.subarray(0, this.n);
    let user = this.O.subarray(0, 32);
    if (this.r === 2) user = rc4(key, user);
    else for (let i = 19; i >= 0; i--) user = rc4(xorKey(key, i), user);
    return user;
  }

  /**
   * Try a password: 'owner' when it is the owner password, 'user' when it
   * is the user password, null when it is neither. The handler keeps the
   * file key it makes; the password itself is not kept.
   */
  authenticate(password = '') {
    let key = null;
    let access = null;
    if (this.r >= 5) {
      const pw = utf8(password);
      const U48 = this.U.subarray(0, 48);
      if (this.O.length >= 48 && same(hash5(this.r, pw, this.O.subarray(32, 40), U48), this.O.subarray(0, 32))) {
        key = unwrap(hash5(this.r, pw, this.O.subarray(40, 48), U48), this.OE);
        access = 'owner';
      } else if (this.U.length >= 48 && same(hash5(this.r, pw, this.U.subarray(32, 40)), this.U.subarray(0, 32))) {
        key = unwrap(hash5(this.r, pw, this.U.subarray(40, 48)), this.UE);
        access = 'user';
      }
    } else {
      key = this.userKey(this.userFromOwner(password));
      if (key) access = 'owner';
      else {
        key = this.userKey(padded(password));
        if (key) access = 'user';
      }
    }
    if (!key) return null;
    this.key = key;
    this.access = access;
    return access;
  }

  /** What the file allows the person who opened it: everything, for its owner. */
  get permissions() {
    const out = {};
    for (const [what, bit] of Object.entries(BITS)) out[what] = this.access === 'owner' || Boolean(this.P & (1 << (bit - 1)));
    // Revision 2 has no bits of its own for filling forms or assembling: changing the document covers both.
    if (this.access !== 'owner' && this.r === 2) { out.fill = out.change || out.comment; out.assemble = out.change; }
    if (this.access !== 'owner') out.fill = out.fill || out.comment;
    return out;
  }

  // ── objects ─────────────────────────────────────────────────────────────

  /** The key and cipher for one object's strings or streams under a method. */
  cipher(method, num, gen) {
    if (method === 'AESV3' || this.r >= 5) return { bits: 256, key: this.key };
    const extra = Buffer.from([num & 255, (num >> 8) & 255, (num >> 16) & 255, gen & 255, (gen >> 8) & 255]);
    const aes = method === 'AESV2';
    const key = md5(this.key, extra, aes ? Buffer.from('sAlT', 'latin1') : EMPTY).subarray(0, Math.min(this.n + 5, 16));
    return aes ? { bits: 128, key } : { rc4: true, key };
  }

  crypt(method, num, gen, data, encrypt) {
    if (!this.key || method === 'None') return data;
    const c = this.cipher(method, num, gen);
    if (c.rc4) return rc4(c.key, data);
    return encrypt ? aesEncrypt(c.bits, c.key, data) : aesDecrypt(c.bits, c.key, data);
  }

  /** The method a stream is encrypted with: its own crypt filter when it names one, else the file's. */
  streamMethod(dict, resolve = (v) => v) {
    if (dict.type === 'XRef') return 'None';
    if (!this.encryptMetadata && dict.type === 'Metadata') return 'None';
    const filters = resolve(dict.get('Filter'));
    const list = Array.isArray(filters) ? filters : filters ? [filters] : [];
    const at = list.findIndex((f) => resolve(f)?.name === 'Crypt');
    if (at >= 0 && this.method) {
      const parms = resolve(dict.get('DecodeParms'));
      const p = resolve(Array.isArray(parms) ? parms[at] : parms);
      const n = p && p.get ? resolve(p.get('Name'))?.name : '';
      if (!n || n === 'Identity') return 'None';
      const cf = this.filters && this.filters.get ? resolve(this.filters.get(n)) : null;
      return cf && cf.get ? resolve(cf.get('CFM'))?.name || 'None' : 'None';
    }
    return this.stm.method;
  }

  /**
   * A value read from object `num gen`, its strings and its stream's bytes
   * decrypted (`encrypt` false) or encrypted (true). Reading changes the
   * value given, which the reader has just parsed; writing gives a copy,
   * so the reader's own values are never changed.
   */
  transform(value, num, gen, encrypt, resolve) {
    // The reader's own kinds of value, asked for here: the reader needs this module first.
    const { PdfString, Dict, Stream } = require('./reader');
    const str = this.str.method;
    const walk = (v, depth) => {
      if (depth > 256) return v;
      if (v instanceof PdfString) {
        const bytes = this.crypt(str, num, gen, v.bytes, encrypt);
        if (!encrypt) { v.bytes = bytes; return v; }
        return new PdfString(bytes, true);
      }
      if (Array.isArray(v)) {
        if (!encrypt) { for (let i = 0; i < v.length; i++) v[i] = walk(v[i], depth + 1); return v; }
        return v.map((x) => walk(x, depth + 1));
      }
      if (v instanceof Dict) {
        // A signature's /Contents is the signature's bytes as they were signed, never encrypted.
        const signature = v.has('ByteRange') && v.has('Contents');
        const out = encrypt ? new Dict() : v;
        for (const [k, x] of v.entries()) out.set(k, signature && k === 'Contents' ? x : walk(x, depth + 1));
        return out;
      }
      return v;
    };
    if (value instanceof Stream) {
      if (value.dict.type === 'XRef') return value;
      const method = this.streamMethod(value.dict, resolve);
      const dict = walk(value.dict, 0);
      const data = this.crypt(method, num, gen, value.data, encrypt);
      if (!encrypt) { value.data = data; return value; }
      return new Stream(dict, data);
    }
    return walk(value, 0);
  }

  decryptObject(value, num, gen, resolve) { return this.transform(value, num, gen, false, resolve); }

  encryptObject(value, num, gen, resolve) { return this.transform(value, num, gen, true, resolve); }

  /** Whether another file's handler is the same protection, the same key: what a save of pages from both may keep. */
  sameAs(other) {
    return Boolean(other && this.key && other.key && same(this.key, other.key) && same(this.O, other.O) && same(this.U, other.U) && this.P === other.P);
  }
}

module.exports = { StandardSecurity, SecurityError, rc4 };
