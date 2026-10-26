'use strict';

// Protected PDFs for the tests and the window checks, made here as a PDF
// writer protects one (the Standard security handler of ISO 32000): RC4 of
// 40 and 128 bits (revisions 2 and 3), AES-128 (revision 4, crypt filters)
// and AES-256 (revision 6, the hash loop of SHA-256/384/512). Written apart
// from the reader's own decryption (src/security.js) on purpose, so the one
// checks the other, and pdf.js checks both.
//
// A plain file goes in (any the reader opens); every object comes out with
// its strings and streams encrypted, under a new Encrypt dictionary, with
// classic cross-references or, asked, with the objects that can be packed
// into one object stream and a cross-reference stream, as Acrobat writes.

const crypto = require('crypto');
const zlib = require('zlib');
const { openPdf, Dict, Name, Ref, PdfString, Stream } = require('../src/reader');
const { valueText } = require('../src/assemble');

const PAD = Buffer.from('28bf4e5e4e758a4164004e56fffa01082e2e00b6d0683e802f0ca9fe6453697a', 'hex');

function rc4(key, data) {
  const s = Array.from({ length: 256 }, (_, i) => i);
  for (let i = 0, j = 0; i < 256; i++) {
    j = (j + s[i] + key[i % key.length]) & 255;
    [s[i], s[j]] = [s[j], s[i]];
  }
  const out = Buffer.alloc(data.length);
  for (let k = 0, i = 0, j = 0; k < data.length; k++) {
    i = (i + 1) & 255;
    j = (j + s[i]) & 255;
    [s[i], s[j]] = [s[j], s[i]];
    out[k] = data[k] ^ s[(s[i] + s[j]) & 255];
  }
  return out;
}

const md5 = (...parts) => crypto.createHash('md5').update(Buffer.concat(parts)).digest();
const padded = (password) => Buffer.concat([Buffer.from(password, 'latin1').subarray(0, 32), PAD]).subarray(0, 32);
const xorKey = (key, i) => Buffer.from(key.map((b) => b ^ i));

/** ISO 32000-2's hash for revision 6 (Algorithm 2.B). */
function hash6(password, salt, udata = Buffer.alloc(0)) {
  const pw = Buffer.from(password, 'utf8').subarray(0, 127);
  let k = crypto.createHash('sha256').update(Buffer.concat([pw, salt, udata])).digest();
  let e = Buffer.alloc(0);
  for (let i = 0; i < 64 || e[e.length - 1] > i - 32; i++) {
    const k1 = Buffer.concat(Array(64).fill(Buffer.concat([pw, k, udata])));
    const c = crypto.createCipheriv('aes-128-cbc', k.subarray(0, 16), k.subarray(16, 32));
    c.setAutoPadding(false);
    e = Buffer.concat([c.update(k1), c.final()]);
    let sum = 0;
    for (let b = 0; b < 16; b++) sum += e[b];
    k = crypto.createHash(['sha256', 'sha384', 'sha512'][sum % 3]).update(e).digest();
  }
  return k.subarray(0, 32);
}

const aes = (bits, key, data, { iv = crypto.randomBytes(16), pad = true, ecb = false } = {}) => {
  const c = ecb ? crypto.createCipheriv(`aes-${bits}-ecb`, key, null) : crypto.createCipheriv(`aes-${bits}-cbc`, key, iv);
  c.setAutoPadding(pad);
  const body = Buffer.concat([c.update(data), c.final()]);
  return ecb || !pad ? body : Buffer.concat([iv, body]);
};

const KINDS = {
  'rc4-40': { v: 1, r: 2, bits: 40 },
  'rc4-128': { v: 2, r: 3, bits: 128 },
  'aes-128': { v: 4, r: 4, bits: 128 },
  'aes-256': { v: 5, r: 6, bits: 256 },
};

/**
 * The Encrypt dictionary and the means to encrypt each object, for a kind,
 * passwords and permissions. `owner` empty takes the user password, as
 * writers do.
 */
function securityFor({ kind = 'aes-256', user = '', owner = '', permissions = -4, encryptMetadata = true, id }) {
  const { v, r, bits } = KINDS[kind];
  const P = permissions | 0;
  const pBytes = Buffer.alloc(4);
  pBytes.writeInt32LE(P);
  const dict = new Dict([['Filter', new Name('Standard')], ['V', v], ['R', r], ['P', P]]);
  if (r === 6) {
    const key = crypto.randomBytes(32);
    const [uv, uk, ov, ok] = [0, 1, 2, 3].map(() => crypto.randomBytes(8));
    const U = Buffer.concat([hash6(user, uv), uv, uk]);
    const UE = aes(256, hash6(user, uk), key, { iv: Buffer.alloc(16), pad: false });
    const O = Buffer.concat([hash6(owner || user, ov, U), ov, ok]);
    const OE = aes(256, hash6(owner || user, ok, U), key, { iv: Buffer.alloc(16), pad: false });
    const perms = Buffer.concat([pBytes, Buffer.from([0xff, 0xff, 0xff, 0xff]), Buffer.from(encryptMetadata ? 'T' : 'F'), Buffer.from('adb'), crypto.randomBytes(4)]);
    dict.set('Length', 256);
    dict.set('CF', new Dict([['StdCF', new Dict([['AuthEvent', new Name('DocOpen')], ['CFM', new Name('AESV3')], ['Length', 32]])]]));
    dict.set('StmF', new Name('StdCF'));
    dict.set('StrF', new Name('StdCF'));
    for (const [k, b] of [['O', O], ['U', U], ['OE', OE], ['UE', UE], ['Perms', aes(256, key, perms, { ecb: true, pad: false })]]) dict.set(k, new PdfString(b, true));
    if (!encryptMetadata) dict.set('EncryptMetadata', false);
    return { dict, objectKey: () => key, cipher: (k, data) => aes(256, k, data) };
  }
  const n = bits / 8;
  // The owner entry: the user password encrypted with a key from the owner password.
  let ohash = md5(padded(owner || user));
  if (r >= 3) for (let i = 0; i < 50; i++) ohash = md5(ohash);
  const okey = ohash.subarray(0, n);
  let O = rc4(okey, padded(user));
  if (r >= 3) for (let i = 1; i <= 19; i++) O = rc4(xorKey(okey, i), O);
  // The file key, from the user password.
  let hash = md5(padded(user), O, pBytes, id, r >= 4 && !encryptMetadata ? Buffer.from([0xff, 0xff, 0xff, 0xff]) : Buffer.alloc(0));
  if (r >= 3) for (let i = 0; i < 50; i++) hash = md5(hash.subarray(0, n));
  const key = hash.subarray(0, n);
  let U;
  if (r === 2) U = rc4(key, PAD);
  else {
    U = rc4(key, md5(PAD, id));
    for (let i = 1; i <= 19; i++) U = rc4(xorKey(key, i), U);
    U = Buffer.concat([U, crypto.randomBytes(16)]);
  }
  if (r >= 3) dict.set('Length', bits);
  dict.set('O', new PdfString(O, true));
  dict.set('U', new PdfString(U, true));
  const useAes = kind === 'aes-128';
  if (r === 4) {
    dict.set('CF', new Dict([['StdCF', new Dict([['AuthEvent', new Name('DocOpen')], ['CFM', new Name(useAes ? 'AESV2' : 'V2')], ['Length', 16]])]]));
    dict.set('StmF', new Name('StdCF'));
    dict.set('StrF', new Name('StdCF'));
    if (!encryptMetadata) dict.set('EncryptMetadata', false);
  }
  const objectKey = (num, gen) => {
    const extra = Buffer.from([num & 255, (num >> 8) & 255, (num >> 16) & 255, gen & 255, (gen >> 8) & 255]);
    return md5(key, extra, useAes ? Buffer.from('sAlT') : Buffer.alloc(0)).subarray(0, Math.min(n + 5, 16));
  };
  return { dict, objectKey, cipher: (k, data) => (useAes ? aes(128, k, data) : rc4(k, data)) };
}

/** A value with every string in it encrypted for the object it belongs to. */
function sealed(value, k, sec) {
  if (value instanceof PdfString) return new PdfString(sec.cipher(k, value.bytes), true);
  if (Array.isArray(value)) return value.map((x) => sealed(x, k, sec));
  if (value instanceof Dict) {
    const out = new Dict();
    for (const [key, x] of value.entries()) out.set(key, sealed(x, k, sec));
    return out;
  }
  return value;
}

/**
 * The plain file `bytes` protected: { kind: 'rc4-40' | 'rc4-128' |
 * 'aes-128' | 'aes-256', user, owner, permissions, encryptMetadata,
 * objectStreams }. Returns the protected file's bytes.
 */
function protectPdf(bytes, options = {}) {
  const file = openPdf(bytes);
  const id = crypto.createHash('md5').update(bytes).digest();
  const sec = securityFor({ ...options, id });
  const nums = [...file.entries.keys()].filter((num) => file.entries.get(num).type !== 0).sort((a, b) => a - b);
  const values = new Map();
  for (const num of nums) {
    const value = file.getObject(num);
    if (value === null) continue;
    if (value instanceof Stream && (value.dict.type === 'XRef' || value.dict.type === 'ObjStm')) continue;
    values.set(num, value);
  }
  const encryptNum = Math.max(...values.keys()) + 1;
  let next = encryptNum + 1;
  const metadataPlain = options.encryptMetadata === false;
  const seal = (num, value) => {
    const k = sec.objectKey(num, 0);
    if (value instanceof Stream) {
      const keep = metadataPlain && value.dict.type === 'Metadata';
      const dict = sealed(value.dict, k, sec);
      const data = keep ? value.data : sec.cipher(k, value.data);
      return new Stream(dict, data);
    }
    return sealed(value, k, sec);
  };
  const chunks = [];
  let length = 0;
  const push = (data) => {
    const b = Buffer.isBuffer(data) ? data : Buffer.from(data, 'latin1');
    chunks.push(b);
    length += b.length;
  };
  const objectText = (value) => {
    if (!(value instanceof Stream)) return Buffer.from(`${valueText(value)}`, 'latin1');
    const dict = new Dict(value.dict.entries());
    dict.set('Length', value.data.length);
    return Buffer.concat([Buffer.from(`${valueText(dict)}\nstream\n`, 'latin1'), value.data, Buffer.from('\nendstream', 'latin1')]);
  };
  push('%PDF-1.7\n%\xE2\xE3\xCF\xD3\n');
  const offsets = new Map(); // num -> { offset } | { stm, idx }
  const write = (num, body) => {
    offsets.set(num, { offset: length });
    push(`${num} 0 obj\n`);
    push(body);
    push('\nendobj\n');
  };
  const packed = [];
  for (const [num, value] of values) {
    if (options.objectStreams && !(value instanceof Stream)) { packed.push([num, value]); continue; }
    write(num, objectText(seal(num, value)));
  }
  write(encryptNum, objectText(sec.dict));
  const trailer = new Dict([['Root', file.trailer.get('Root')], ['Encrypt', new Ref(encryptNum, 0)], ['ID', [new PdfString(id, true), new PdfString(id, true)]]]);
  if (file.trailer.has('Info')) trailer.set('Info', file.trailer.get('Info'));
  if (options.objectStreams) {
    // The packed objects: plain inside, the object stream encrypted as one stream.
    const stmNum = next++;
    let head = '';
    let body = '';
    packed.forEach(([num, value], idx) => {
      head += `${num} ${body.length} `;
      body += `${valueText(value)}\n`;
      offsets.set(num, { stm: stmNum, idx });
    });
    const data = zlib.deflateSync(Buffer.from(head + body, 'latin1'));
    const stm = new Stream(new Dict([['Type', new Name('ObjStm')], ['N', packed.length], ['First', head.length], ['Filter', new Name('FlateDecode')]]), data);
    write(stmNum, objectText(seal(stmNum, stm)));
    const xrefNum = next++;
    const size = next;
    offsets.set(xrefNum, { offset: length });
    const rows = [];
    for (let num = 0; num < size; num++) {
      const e = offsets.get(num);
      const row = Buffer.alloc(7);
      if (!e) row[0] = 0;
      else if (e.stm) { row[0] = 2; row.writeUInt32BE(e.stm, 1); row.writeUInt16BE(e.idx, 5); } else { row[0] = 1; row.writeUInt32BE(e.offset, 1); }
      rows.push(row);
    }
    const xdata = zlib.deflateSync(Buffer.concat(rows));
    for (const [k, x] of [['Type', new Name('XRef')], ['Size', size], ['W', [1, 4, 2]], ['Filter', new Name('FlateDecode')]]) trailer.set(k, x);
    const xrefAt = length;
    push(`${xrefNum} 0 obj\n`);
    push(objectText(new Stream(trailer, xdata)));
    push(`\nendobj\nstartxref\n${xrefAt}\n%%EOF\n`);
    return Buffer.concat(chunks, length);
  }
  const size = next;
  const xrefAt = length;
  let xref = `xref\n0 ${size}\n0000000000 65535 f \n`;
  for (let num = 1; num < size; num++) {
    const e = offsets.get(num);
    xref += e ? `${String(e.offset).padStart(10, '0')} 00000 n \n` : '0000000000 00001 f \n';
  }
  push(xref);
  trailer.set('Size', size);
  push(`trailer\n${valueText(trailer)}\nstartxref\n${xrefAt}\n%%EOF\n`);
  return Buffer.concat(chunks, length);
}

module.exports = { protectPdf, KINDS };
