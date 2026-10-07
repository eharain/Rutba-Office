// Password-protected Word, Excel and PowerPoint 97-2003 files.
//
// A .doc, .xls or .ppt saved with "Password to open" stays a compound file
// with its streams in place, their bytes encrypted where the format says:
//
//   - RC4 CryptoAPI ([MS-OFFCRYPTO] 2.3.5) — what Office 2003 onwards writes:
//     SHA-1 over a salt and the password, then over a block number, gives
//     each block's RC4 key (128-bit, or 40 bits padded with zeros).
//   - RC4 ([MS-OFFCRYPTO] 2.3.6) — Office 97/2000 compatible: MD5 over the
//     password, the salt and a block number.
//   - XOR obfuscation ([MS-OFFCRYPTO] 2.3.7, method 1) — Excel 5.0/95, and
//     Excel 97's weakest option: a sixteen-byte key from the password.
//
// Where each format keeps the description and what it encrypts:
//
//   - Word ([MS-DOC] 2.2.6.2): the FIB's fEncrypted flag; the encryption
//     header at the start of the table stream; the WordDocument stream (all
//     but its first 68 bytes), the table stream (all but that header) and
//     the Data stream, in 512-byte blocks.
//   - Excel ([MS-XLS] 2.2.10): a FilePass record after the first BOF; every
//     record's data after it, in 1024-byte blocks counted from the start of
//     the stream — but for BOF, FilePass and a few others, and the first
//     four bytes of each BoundSheet8 (where the sheet starts).
//   - PowerPoint ([MS-PPT] 2.3.7): the Current User's header token, and a
//     CryptSession10Container the last edit names; each persist object of
//     the PowerPoint Document stream encrypted whole with its persist id as
//     the block number, and each picture of the Pictures stream.
//
// Decrypted, a file is written back as the compound file it would have been
// without the password — the flags cleared, the FilePass record taken out —
// for the readers to read as any other.
//
// Node's own hashes; this is the engine's side, never a window's.

import { createHash } from 'node:crypto';
import { CompoundFile } from './cfb.js';
import { writeCompoundFile } from './cfb-write.js';
import { EncryptedFileError } from './crypt.js';

const u16 = (b, o) => b[o] | (b[o + 1] << 8);
const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
const md5 = (data) => createHash('md5').update(data).digest();
const sha1 = (data) => createHash('sha1').update(data).digest();
const le32 = (n) => { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0, 0); return b; };

/* ── RC4 ───────────────────────────────────────────────────────────────── */

class RC4 {
  constructor(key) {
    const s = new Uint8Array(256);
    for (let i = 0; i < 256; i++) s[i] = i;
    for (let i = 0, j = 0; i < 256; i++) {
      j = (j + s[i] + key[i % key.length]) & 255;
      const t = s[i]; s[i] = s[j]; s[j] = t;
    }
    this.s = s;
    this.i = 0;
    this.j = 0;
  }

  /** XOR `n` bytes of `buf` from `at` with the keystream, in place. */
  xor(buf, at = 0, n = buf.length - at) {
    const s = this.s;
    let { i, j } = this;
    for (let k = at; k < at + n; k++) {
      i = (i + 1) & 255;
      j = (j + s[i]) & 255;
      const t = s[i]; s[i] = s[j]; s[j] = t;
      buf[k] ^= s[(s[i] + s[j]) & 255];
    }
    this.i = i;
    this.j = j;
    return buf;
  }

  skip(n) {
    this.xor(new Uint8Array(n));
  }
}

/**
 * XOR a range of a stream with the keystream for its place in it: the key
 * changes at every `blockSize` bytes, block numbers counted from the start.
 */
function positional(keyFor, blockSize) {
  let block = -1;
  let offset = 0;
  let rc4 = null;
  return (buf, start, end) => {
    for (let pos = start; pos < end;) {
      const b = Math.floor(pos / blockSize);
      const o = pos % blockSize;
      if (b !== block || o < offset) {
        rc4 = new RC4(keyFor(b));
        block = b;
        offset = 0;
      }
      if (o > offset) { rc4.skip(o - offset); offset = o; }
      const n = Math.min(end - pos, blockSize - o);
      rc4.xor(buf, pos, n);
      offset += n;
      pos += n;
    }
  };
}

/* ── the description, and the keys from a password ─────────────────────── */

/** An RC4 or RC4 CryptoAPI encryption header and verifier, from `at` in `b`. */
export function readEncryptionHeader(b, at = 0) {
  if (b.length < at + 8) throw new EncryptedFileError('damaged', 'its encryption header is cut short');
  const major = u16(b, at);
  const minor = u16(b, at + 2);
  if (major === 1 && minor === 1) {
    if (b.length < at + 52) throw new EncryptedFileError('damaged', 'its encryption header is cut short');
    return { kind: 'rc4', keyBits: 128, salt: b.subarray(at + 4, at + 20), verifier: b.subarray(at + 20, at + 36), verifierHash: b.subarray(at + 36, at + 52) };
  }
  if ((major === 2 || major === 3 || major === 4) && minor === 2) {
    const headerSize = u32(b, at + 8);
    const h = at + 12;
    const algId = u32(b, h + 8);
    const algIdHash = u32(b, h + 12);
    if (algId && algId !== 0x6801) throw new EncryptedFileError('unsupported', `it is encrypted with an algorithm other than RC4 (0x${algId.toString(16)})`);
    if (algIdHash && algIdHash !== 0x8004) throw new EncryptedFileError('unsupported', `its key is hashed with an algorithm other than SHA-1 (0x${algIdHash.toString(16)})`);
    const keyBits = u32(b, h + 16) || 40;
    if (keyBits < 40 || keyBits > 128 || keyBits % 8) throw new EncryptedFileError('damaged', `its key size of ${keyBits} bits is not one RC4 uses`);
    const v = h + headerSize;
    const saltSize = u32(b, v);
    const hashAt = v + 4 + saltSize + 16;
    if (saltSize !== 16 || b.length < hashAt + 4) throw new EncryptedFileError('damaged', 'its password verifier is cut short');
    const hashSize = u32(b, hashAt);
    if (b.length < hashAt + 4 + hashSize) throw new EncryptedFileError('damaged', 'its password verifier is cut short');
    return {
      kind: 'cryptoapi', keyBits,
      salt: b.subarray(v + 4, v + 20),
      verifier: b.subarray(v + 20, v + 36),
      verifierHash: b.subarray(hashAt + 4, hashAt + 4 + hashSize),
    };
  }
  throw new EncryptedFileError('unsupported', `it is protected with a version ${major}.${minor} encryption header`);
}

/**
 * The function from a block number to its RC4 key, once the password is
 * known to be right; a wrong one is refused here.
 */
export function rc4Keys(header, password) {
  const pw = Buffer.from(String(password), 'utf16le');
  let keyFor;
  if (header.kind === 'rc4') {
    // [MS-OFFCRYPTO] 2.3.6.2: MD5 of the password, five bytes of it with the
    // salt sixteen times over, five bytes of that with the block number.
    const h0 = md5(pw).subarray(0, 5);
    const spread = Buffer.alloc(21 * 16);
    for (let i = 0; i < 16; i++) { h0.copy(spread, i * 21); Buffer.from(header.salt).copy(spread, i * 21 + 5); }
    const h1 = md5(spread).subarray(0, 5);
    keyFor = (block) => md5(Buffer.concat([h1, le32(block)]));
  } else {
    // [MS-OFFCRYPTO] 2.3.5.2: SHA-1 of the salt and password, then of that
    // and the block number; a 40-bit key is padded to 128 bits with zeros.
    const h0 = sha1(Buffer.concat([Buffer.from(header.salt), pw]));
    keyFor = (block) => {
      const h = sha1(Buffer.concat([h0, le32(block)]));
      return header.keyBits === 40 ? Buffer.concat([h.subarray(0, 5), Buffer.alloc(11)]) : h.subarray(0, header.keyBits / 8);
    };
  }
  // The verifier and its hash, decrypted with block 0's key as one stream.
  const rc4 = new RC4(keyFor(0));
  const verifier = rc4.xor(Buffer.from(header.verifier));
  const hash = rc4.xor(Buffer.from(header.verifierHash));
  const want = header.kind === 'rc4' ? md5(verifier) : sha1(verifier);
  if (!want.subarray(0, Math.min(want.length, hash.length)).equals(hash.subarray(0, Math.min(want.length, hash.length)))) {
    throw new EncryptedFileError('password', 'the password is not the one this file was saved with');
  }
  return keyFor;
}

/* ── XOR obfuscation ([MS-OFFCRYPTO] 2.3.7, method 1) ──────────────────── */

/** A password as the single bytes method 1 works on: each character's low byte, or its high one when that is zero; at most fifteen. */
const xorBytes = (password) => [...String(password)].slice(0, 15).map((ch) => { const c = ch.charCodeAt(0); return (c & 0xff) || (c >> 8); });

/** The password verifier Excel keeps: [MS-OFFCRYPTO] 2.3.7.1. */
export function xorVerifier(password) {
  const bytes = xorBytes(password);
  let v = 0;
  for (let i = bytes.length - 1; i >= 0; i--) v = ((((v >> 14) & 1) | ((v << 1) & 0x7fff)) ^ bytes[i]) & 0xffff;
  v = ((((v >> 14) & 1) | ((v << 1) & 0x7fff)) ^ bytes.length) & 0xffff;
  return v ^ 0xce4b;
}

const rol16 = (v, n) => ((v << n) | (v >>> (16 - n))) & 0xffff;
const rol8 = (v, n) => ((v << n) | (v >>> (8 - n))) & 0xff;

/** The sixteen-bit key the password's bytes give: [MS-OFFCRYPTO] 2.3.7.2's, made by its shift register rather than its table. */
export function xorKey(password) {
  const bytes = xorBytes(password);
  if (!bytes.length) return 0;
  let key = 0;
  let base = 0x8000;
  let end = 0xffff;
  for (let i = bytes.length - 1; i >= 0; i--) {
    let c = bytes[i] & 0x7f;
    for (let bit = 0; bit < 8; bit++) {
      base = rol16(base, 1);
      if (base & 1) base ^= 0x1020;
      if (c & 1) key ^= base;
      c >>= 1;
      end = rol16(end, 1);
      if (end & 1) end ^= 0x1020;
    }
  }
  return (key ^ end) & 0xffff;
}

const XOR_PAD = [0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0x00, 0xbe, 0x0f, 0x00, 0xbf, 0x0f, 0x00];

/** The sixteen key bytes Excel's obfuscation XORs with: the password padded, each byte XORed with the key and turned two bits. */
function xorArray(password, key) {
  const bytes = xorBytes(password);
  const out = new Uint8Array(16);
  for (let i = 0; i < 16; i++) {
    const b = i < bytes.length ? bytes[i] : XOR_PAD[i - bytes.length] ?? 0;
    out[i] = rol8(b ^ (i & 1 ? key >> 8 : key & 0xff), 2);
  }
  return out;
}

/* ── which kind, if any ────────────────────────────────────────────────── */

/**
 * How a 97-2003 file is protected: { app: 'doc' | 'xls' | 'ppt', method:
 * 'rc4' | 'cryptoapi' | 'xor' } — or null for one that is not.
 */
export function binaryEncryption(bytes) {
  if (!CompoundFile.is(bytes)) return null;
  let cfb;
  try { cfb = new CompoundFile(bytes); } catch { return null; }
  const app = cfb.application();
  try {
    if (app === 'doc') {
      const wd = cfb.read(cfb.find(['WordDocument']));
      if (wd.length < 0x44) return null;
      const flags = u16(wd, 0x0a);
      if (!(flags & 0x0100)) return null;
      if (flags & 0x8000) return { app, method: 'xor' };
      const table = cfb.read(cfb.find([flags & 0x0200 ? '1Table' : '0Table']));
      return { app, method: table.length >= 4 && u16(table, 0) === 1 ? 'rc4' : 'cryptoapi' };
    }
    if (app === 'xls') {
      const s = cfb.read(cfb.find(['Workbook']) || cfb.find(['Book']));
      const pass = filePass(s);
      if (!pass) return null;
      return { app, method: pass.method };
    }
    if (app === 'ppt') {
      const cu = cfb.read(cfb.find(['Current User']));
      if (cu.length >= 16 && u32(cu, 12) === 0xf3d1c4df) return { app, method: 'cryptoapi' };
      return null;
    }
  } catch {
    return null;
  }
  return null;
}

/* ── Word ──────────────────────────────────────────────────────────────── */

function decryptDoc(cfb, password) {
  const wdEntry = cfb.find(['WordDocument']);
  const wd = Buffer.from(cfb.read(wdEntry));
  const flags = u16(wd, 0x0a);
  if (flags & 0x8000) throw new EncryptedFileError('unsupported', "it is hidden with Word's old XOR obfuscation, which this version does not undo");
  const tableName = flags & 0x0200 ? '1Table' : '0Table';
  const tableEntry = cfb.find([tableName]);
  if (!tableEntry) throw new EncryptedFileError('damaged', `it has no ${tableName} stream`);
  const table = Buffer.from(cfb.read(tableEntry));
  const lKey = Math.min(u32(wd, 0x0e), table.length);
  const keyFor = rc4Keys(readEncryptionHeader(table, 0), password);
  const out = new Map();
  // The whole stream through the keystream, then what was never encrypted put back.
  const plainWd = Buffer.from(wd);
  positional(keyFor, 0x200)(plainWd, 0, plainWd.length);
  wd.copy(plainWd, 0, 0, 0x44);
  plainWd.writeUInt16LE(flags & ~0x0100, 0x0a);
  plainWd.writeUInt32LE(0, 0x0e);
  out.set('WordDocument', plainWd);
  const plainTable = Buffer.from(table);
  positional(keyFor, 0x200)(plainTable, 0, plainTable.length);
  table.copy(plainTable, 0, 0, lKey);
  out.set(tableName, plainTable);
  const dataEntry = cfb.find(['Data']);
  if (dataEntry) {
    const data = Buffer.from(cfb.read(dataEntry));
    positional(keyFor, 0x200)(data, 0, data.length);
    out.set('Data', data);
  }
  return out;
}

/* ── Excel ─────────────────────────────────────────────────────────────── */

const BOFS = new Set([0x0009, 0x0209, 0x0409, 0x0809]);
/** [MS-XLS] 2.2.10: the records never encrypted — BOF, FilePass, UsrExcl, FileLock, InterfaceHdr, RRDInfo, RRDHead. */
const NEVER = new Set([...BOFS, 0x002f, 0x0194, 0x0195, 0x00e1, 0x0196, 0x0138]);

function records(s) {
  const out = [];
  for (let p = 0; p + 4 <= s.length;) {
    const size = u16(s, p + 2);
    out.push({ id: u16(s, p), at: p, size });
    p += 4 + size;
  }
  return out;
}

/** The FilePass record of a workbook stream, and the method it names; null when there is none. */
function filePass(s) {
  const recs = records(s);
  const rec = recs.find((r) => r.id === 0x002f);
  if (!rec) return null;
  const data = s.subarray(rec.at + 4, rec.at + 4 + rec.size);
  const biff8 = recs[0]?.id === 0x0809 && u16(s, 4) === 0x0600;
  if (biff8 && data.length >= 6 && u16(data, 0) === 1) return { rec, recs, data, biff8, method: u16(data, 2) === 1 && u16(data, 4) === 1 ? 'rc4' : 'cryptoapi' };
  return { rec, recs, data, biff8, method: 'xor' };
}

function decryptXls(cfb, password) {
  const entry = cfb.find(['Workbook']) || cfb.find(['Book']);
  const s = Buffer.from(cfb.read(entry));
  const pass = filePass(s);
  if (!pass) return new Map();
  const { rec, recs, data, biff8 } = pass;
  const out = Buffer.from(s);
  if (pass.method !== 'xor') {
    const crypt = positional(rc4Keys(readEncryptionHeader(data, 2), password), 0x400);
    for (const r of recs) {
      if (r.at <= rec.at || NEVER.has(r.id)) continue;
      crypt(out, r.at + 4 + (r.id === 0x0085 ? 4 : 0), r.at + 4 + r.size);
    }
  } else {
    const at = biff8 ? 2 : 0;
    const fileKey = u16(data, at);
    if (xorVerifier(password) !== u16(data, at + 2)) throw new EncryptedFileError('password', 'the password is not the one this file was saved with');
    const key = xorArray(password, fileKey || xorKey(password));
    for (const r of recs) {
      if (r.at <= rec.at || NEVER.has(r.id)) continue;
      // Each byte's key byte by its place in the stream, offset by the record's size, as Excel counts it.
      for (let pos = r.at + 4 + (r.id === 0x0085 ? 4 : 0); pos < r.at + 4 + r.size; pos++) out[pos] = rol8(s[pos], 3) ^ key[(pos + r.size) & 15];
    }
  }
  // The FilePass record taken out, and each sheet's start moved up by as much.
  const cut = 4 + rec.size;
  const plain = Buffer.concat([out.subarray(0, rec.at), out.subarray(rec.at + cut)]);
  for (const r of records(plain)) {
    if (r.id !== 0x0085 || r.size < 4) continue;
    const pos = u32(plain, r.at + 4);
    if (pos > rec.at) plain.writeUInt32LE(pos - cut, r.at + 4);
  }
  return new Map([[entry.name, plain]]);
}

/* ── PowerPoint ────────────────────────────────────────────────────────── */

function decryptPpt(cfb, password) {
  const cuEntry = cfb.find(['Current User']);
  const docEntry = cfb.find(['PowerPoint Document']);
  if (!cuEntry || !docEntry) throw new EncryptedFileError('damaged', 'it is missing the streams a presentation has');
  const cu = Buffer.from(cfb.read(cuEntry));
  const doc = Buffer.from(cfb.read(docEntry));
  // Every edit back to the first, each with its persist directory.
  const edits = [];
  const objects = [];
  let session = 0;
  for (let at = u32(cu, 16), guard = 0; guard < 512 && at + 8 <= doc.length; guard++) {
    if (u16(doc, at + 2) !== 0x0ff5) break;
    const len = u32(doc, at + 4);
    const body = at + 8;
    edits.push({ at, len });
    if (!session && len >= 32) session = u32(doc, body + 28);
    const dirAt = u32(doc, body + 12);
    if (dirAt + 8 <= doc.length && u16(doc, dirAt + 2) === 0x1772) {
      const end = Math.min(doc.length, dirAt + 8 + u32(doc, dirAt + 4));
      for (let p = dirAt + 8; p + 4 <= end;) {
        const word = u32(doc, p);
        const first = word & 0xfffff;
        const count = word >>> 20;
        p += 4;
        for (let k = 0; k < count && p + 4 <= end; k++, p += 4) objects.push({ id: first + k, offset: u32(doc, p) });
      }
    }
    const previous = u32(doc, body + 8);
    if (!previous || previous === at) break;
    at = previous;
  }
  const sessionAt = objects.find((o) => o.id === session)?.offset;
  if (!session || sessionAt == null || u16(doc, sessionAt + 2) !== 0x2f14) throw new EncryptedFileError('damaged', 'its encryption session cannot be found');
  const keyFor = rc4Keys(readEncryptionHeader(doc, sessionAt + 8), password);
  const out = Buffer.from(doc);
  const done = new Set();
  for (const { id, offset } of objects) {
    if (id === session || done.has(offset) || offset + 8 > doc.length) continue;
    done.add(offset);
    // The whole record, header and all, is one stream under its persist id's key.
    const head = new RC4(keyFor(id)).xor(Buffer.from(doc.subarray(offset, offset + 8)));
    const end = Math.min(doc.length, offset + 8 + u32(head, 4));
    const plain = new RC4(keyFor(id)).xor(Buffer.from(doc.subarray(offset, end)));
    plain.copy(out, offset);
  }
  // No longer encrypted: the edits name no session, the Current User says so.
  for (const { at, len } of edits) if (len >= 32) out.writeUInt32LE(0, at + 8 + 28);
  const cuOut = Buffer.from(cu);
  cuOut.writeUInt32LE(0xe391c05f, 12);
  const result = new Map([[cuEntry.name, cuOut], [docEntry.name, out]]);
  // The pictures: each record's fields — its header, its one or two ids,
  // the tag of a bitmap or the header of a metafile, and the picture
  // itself — each encrypted by itself under block 0's key.
  const picEntry = cfb.find(['Pictures']);
  if (picEntry) {
    const pics = Buffer.from(cfb.read(picEntry));
    const plain = Buffer.from(pics);
    const field = (from, to) => { if (to > from) new RC4(keyFor(0)).xor(plain, from, Math.min(to, plain.length) - from); };
    for (let p = 0; p + 8 <= pics.length;) {
      field(p, p + 8);
      const instance = u16(plain, p) >> 4;
      const type = u16(plain, p + 2);
      const len = u32(plain, p + 4);
      if (type < 0xf018 || type > 0xf117 || p + 8 + len > pics.length) { pics.copy(plain, p, p, p + 8); break; }
      let q = p + 8;
      const ids = instance & 1 ? 2 : 1;
      for (let k = 0; k < ids; k++, q += 16) field(q, q + 16);
      const head = type >= 0xf01a && type <= 0xf01c ? 34 : 1;
      field(q, q + head);
      field(q + head, p + 8 + len);
      p += 8 + len;
    }
    result.set(picEntry.name, plain);
  }
  return result;
}

/* ── the file ──────────────────────────────────────────────────────────── */

/**
 * A password-protected 97-2003 file, decrypted: the compound file it would
 * be without the password, every other stream as it was. A wrong password
 * throws EncryptedFileError('password'); a protection this does not undo,
 * 'unsupported'; a file that cannot be followed, 'damaged'.
 */
export function decryptBinary(bytes, password) {
  const cfb = new CompoundFile(bytes);
  const app = cfb.application();
  const replaced = app === 'doc' ? decryptDoc(cfb, password)
    : app === 'xls' ? decryptXls(cfb, password)
    : app === 'ppt' ? decryptPpt(cfb, password)
    : null;
  if (!replaced) throw new EncryptedFileError('unsupported', 'it is not a Word, Excel or PowerPoint 97-2003 file');
  const streams = [];
  const walk = (entry, path) => {
    for (const child of cfb.childrenOf(entry)) {
      const at = [...path, child.name];
      if (child.type === 2) streams.push({ path: at, data: path.length === 0 && replaced.has(child.name) ? replaced.get(child.name) : cfb.read(child) });
      else if (child.type === 1) walk(child, at);
    }
  };
  walk(cfb.root, []);
  return writeCompoundFile(streams);
}
