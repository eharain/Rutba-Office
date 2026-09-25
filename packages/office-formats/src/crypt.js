// Password-protected Office files — [MS-OFFCRYPTO].
//
// A .docx, .xlsx or .pptx with "Encrypt with Password" set is not a zip: it
// is a compound file holding the zip encrypted (the EncryptedPackage stream),
// a description of how (EncryptionInfo), and a \x06DataSpaces storage naming
// the transform. Two kinds are read here, which between them cover every
// file Office 2007 onwards writes with a password:
//
//   - Agile Encryption (EncryptionInfo 4.4, XML). Office 2010 wrote it with
//     SHA-1 and AES-128, Office 2013 onwards with SHA-512 and AES-256. The
//     password is hashed with a salt and spun (100,000 rounds) into keys
//     that unlock a random key; that key encrypts the package in 4096-byte
//     segments, each with an IV of its own, and an HMAC over the encrypted
//     stream (dataIntegrity) says whether a byte of it was changed.
//
//   - Standard Encryption (EncryptionInfo 3.2 or 4.2, binary). Office 2007:
//     SHA-1 spun 50,000 times, AES-128 in ECB over the whole package, and a
//     verifier to tell a wrong password from a right one. It carries no HMAC.
//
// Written: Agile, as Office 2013 onwards writes it — AES-256-CBC, SHA-512,
// 100,000 rounds, fresh random salts and a fresh key on every save, and the
// HMAC.
//
// Node's own crypto only; this is the engine's side, never a window's.

import * as nodeCrypto from 'node:crypto';
import { CompoundFile } from './cfb.js';
import { writeCompoundFile } from './cfb-write.js';

const { createCipheriv, createDecipheriv, createHash, createHmac, randomBytes, timingSafeEqual } = nodeCrypto;

/**
 * One hash of one buffer. `crypto.hash` (Node 21.7 onwards) skips making a
 * hash object per round, which is a third of the time 100,000 rounds take.
 */
const oneShot = typeof nodeCrypto.hash === 'function' ? (alg, data) => nodeCrypto.hash(alg, data, 'buffer') : (alg, data) => createHash(alg).update(data).digest();

/** `rounds` more rounds of H(iterator + h), the iterator four bytes little-endian. */
function spin(alg, h, rounds) {
  let buf = Buffer.alloc(4 + h.length);
  for (let i = 0; i < rounds; i++) {
    if (buf.length !== 4 + h.length) buf = Buffer.alloc(4 + h.length);
    buf.writeUInt32LE(i, 0);
    h.copy(buf, 4);
    h = oneShot(alg, buf);
  }
  return h;
}

/** Why a file did not open, for the window to say in a sentence. */
export class EncryptedFileError extends Error {
  /** @param {'password' | 'tampered' | 'unsupported' | 'damaged'} code */
  constructor(code, message) {
    super(message);
    this.name = 'EncryptedFileError';
    this.code = code;
  }
}

/* ── the constants the spec fixes ──────────────────────────────────────── */

/** [MS-OFFCRYPTO] 2.3.4.11 and 2.3.4.14: the block keys each derived key is made with. */
export const BLOCK_KEYS = {
  verifierHashInput: Buffer.from([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
  verifierHashValue: Buffer.from([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
  keyValue: Buffer.from([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]),
  integrityKey: Buffer.from([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
  integrityValue: Buffer.from([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]),
};

const SEGMENT = 4096;
const NS_ENCRYPTION = 'http://schemas.microsoft.com/office/2006/encryption';
const NS_PASSWORD = 'http://schemas.microsoft.com/office/2006/keyEncryptor/password';
const NS_CERTIFICATE = 'http://schemas.microsoft.com/office/2006/keyEncryptor/certificate';

const HASHES = { SHA1: 'sha1', 'SHA-1': 'sha1', SHA256: 'sha256', 'SHA-256': 'sha256', SHA384: 'sha384', 'SHA-384': 'sha384', SHA512: 'sha512', 'SHA-512': 'sha512', MD5: 'md5' };

/* ── recognising one ───────────────────────────────────────────────────── */

/** Is this a compound file holding an encrypted package? */
export function isEncryptedPackage(bytes) {
  if (!CompoundFile.is(bytes)) return false;
  try {
    return new CompoundFile(bytes).application() === 'encrypted';
  } catch {
    return false;
  }
}

function streams(bytes) {
  let cfb;
  try {
    cfb = new CompoundFile(bytes);
  } catch {
    throw new EncryptedFileError('damaged', 'the file is not a readable compound file');
  }
  const info = cfb.find(['EncryptionInfo']);
  const pkg = cfb.find(['EncryptedPackage']);
  if (!info || !pkg) throw new EncryptedFileError('damaged', 'the file has no encrypted package in it');
  return { info: Buffer.from(cfb.read(info)), pkg: Buffer.from(cfb.read(pkg)) };
}

/**
 * How a file is encrypted, read from its EncryptionInfo stream.
 * @returns {{ kind: 'agile', keyData, dataIntegrity, password } | { kind: 'standard', ... }}
 */
export function readEncryptionInfo(info) {
  const buf = Buffer.from(info);
  if (buf.length < 8) throw new EncryptedFileError('damaged', 'the description of the encryption is cut short');
  const major = buf.readUInt16LE(0);
  const minor = buf.readUInt16LE(2);
  if (major === 4 && minor === 4) return readAgile(buf.subarray(8).toString('utf8'));
  if ((major === 2 || major === 3 || major === 4) && minor === 2) return readStandard(buf);
  if ((major === 3 || major === 4) && minor === 3) {
    throw new EncryptedFileError('unsupported', 'it is protected by rights management (Extensible Encryption), which only the service that set it can open');
  }
  throw new EncryptedFileError('unsupported', `its encryption (version ${major}.${minor}) is not one Office writes for a password`);
}

function attrs(tag) {
  const out = {};
  for (const m of tag.matchAll(/([\w:]+)\s*=\s*"([^"]*)"/g)) out[m[1].replace(/^\w+:/, '')] = m[2];
  return out;
}

function readAgile(xml) {
  const keyData = /<(?:\w+:)?keyData\b([^>]*)\/?>/.exec(xml);
  const integrity = /<(?:\w+:)?dataIntegrity\b([^>]*)\/?>/.exec(xml);
  if (!keyData) throw new EncryptedFileError('damaged', 'the description of the encryption has no keyData');
  // The password's key encryptor. A file can also carry certificate
  // encryptors (a key for a smart card); only the password one opens here.
  let password = null;
  for (const m of xml.matchAll(/<(?:\w+:)?keyEncryptor\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?keyEncryptor>/g)) {
    if (attrs(m[1]).uri !== NS_PASSWORD) continue;
    const key = /<(?:\w+:)?encryptedKey\b([^>]*)\/?>/.exec(m[2]);
    if (key) password = attrs(key[1]);
  }
  if (!password) {
    throw new EncryptedFileError('unsupported', xml.includes(NS_CERTIFICATE) ? 'it is locked to a certificate, not a password' : 'it carries no password key');
  }
  return { kind: 'agile', keyData: attrs(keyData[1]), dataIntegrity: integrity ? attrs(integrity[1]) : null, password };
}

const ALG_AES = { 0x660e: 128, 0x660f: 192, 0x6610: 256 };

function readStandard(buf) {
  // [MS-OFFCRYPTO] 2.3.4.5: flags, header size, the EncryptionHeader, then
  // the EncryptionVerifier.
  const flags = buf.readUInt32LE(4);
  const headerSize = buf.readUInt32LE(8);
  const h = 12;
  if (buf.length < h + headerSize + 4) throw new EncryptedFileError('damaged', 'the description of the encryption is cut short');
  const algId = buf.readUInt32LE(h + 8);
  const algIdHash = buf.readUInt32LE(h + 12);
  const keyBits = buf.readUInt32LE(h + 16);
  if (flags & 0x10) throw new EncryptedFileError('unsupported', 'it is protected by rights management, which only the service that set it can open');
  if (!(flags & 0x20)) throw new EncryptedFileError('unsupported', 'its encryption is the RC4 kind, which Office never used for these files');
  if (algId !== 0 && !ALG_AES[algId]) throw new EncryptedFileError('unsupported', 'its cipher is not AES');
  if (algIdHash !== 0 && algIdHash !== 0x8004) throw new EncryptedFileError('unsupported', 'its hash is not SHA-1');
  const v = h + headerSize;
  const saltSize = buf.readUInt32LE(v);
  const salt = buf.subarray(v + 4, v + 4 + saltSize);
  const encryptedVerifier = buf.subarray(v + 4 + saltSize, v + 4 + saltSize + 16);
  const hashSize = buf.readUInt32LE(v + 20 + saltSize);
  const encryptedVerifierHash = buf.subarray(v + 24 + saltSize, v + 24 + saltSize + 32);
  return { kind: 'standard', keyBits: keyBits || ALG_AES[algId] || 128, salt, encryptedVerifier, verifierHashSize: hashSize, encryptedVerifierHash };
}

/* ── the pieces of the algorithm ───────────────────────────────────────── */

function hashName(name) {
  const h = HASHES[String(name || '').toUpperCase()];
  if (!h) throw new EncryptedFileError('unsupported', `its hash (${name}) is not one this suite reads`);
  return h;
}

const H = (alg, ...parts) => {
  const h = createHash(alg);
  for (const p of parts) h.update(p);
  return h.digest();
};

const le32 = (n) => {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(n >>> 0, 0);
  return b;
};

/** A value cut or padded to `size`, padding with 0x36 as the spec asks. */
function fit(value, size) {
  if (value.length >= size) return value.subarray(0, size);
  return Buffer.concat([value, Buffer.alloc(size - value.length, 0x36)]);
}

/**
 * [MS-OFFCRYPTO] 2.3.4.11: the password hashed with its salt, then spun —
 * each round the hash of the round's number (four bytes, little-endian)
 * followed by the last round.
 */
export function spinPassword(password, salt, spinCount, alg = 'sha512') {
  return spin(alg, H(alg, salt, Buffer.from(String(password), 'utf16le')), spinCount);
}

/** A key for one purpose from the spun hash: H(spun + blockKey), cut or padded to the key's length. */
export function agileKey(spun, blockKey, keyBits, alg = 'sha512') {
  return fit(H(alg, spun, blockKey), keyBits / 8);
}

function cipherName(keyBits, chaining) {
  if (![128, 192, 256].includes(keyBits)) throw new EncryptedFileError('unsupported', `its key length (${keyBits} bits) is not AES's`);
  if (chaining === 'ChainingModeCFB') return `aes-${keyBits}-cfb8`;
  if (chaining && chaining !== 'ChainingModeCBC') throw new EncryptedFileError('unsupported', `its chaining (${chaining}) is not one Office writes`);
  return `aes-${keyBits}-cbc`;
}

function aes(encrypt, name, key, iv, data) {
  const c = (encrypt ? createCipheriv : createDecipheriv)(name, key, iv);
  c.setAutoPadding(false);
  return Buffer.concat([c.update(data), c.final()]);
}

/** Data padded with zeros to a whole number of blocks. */
function blocks(data, size = 16) {
  const rest = data.length % size;
  return rest ? Buffer.concat([data, Buffer.alloc(size - rest)]) : data;
}

function aesEcb(encrypt, key, data) {
  const c = (encrypt ? createCipheriv : createDecipheriv)(`aes-${key.length * 8}-ecb`, key, null);
  c.setAutoPadding(false);
  return Buffer.concat([c.update(data), c.final()]);
}

/* ── reading ───────────────────────────────────────────────────────────── */

function agileSpec(a, fallback = {}) {
  return {
    alg: hashName(a.hashAlgorithm || fallback.hashAlgorithm),
    keyBits: Number(a.keyBits || fallback.keyBits),
    blockSize: Number(a.blockSize || fallback.blockSize || 16),
    hashSize: Number(a.hashSize || fallback.hashSize),
    chaining: a.cipherChaining || fallback.cipherChaining || 'ChainingModeCBC',
    salt: Buffer.from(a.saltValue || '', 'base64'),
    cipher: a.cipherAlgorithm || fallback.cipherAlgorithm || 'AES',
  };
}

/** The key that encrypts the package, or null when the password is wrong. */
function agileSecretKey(info, password) {
  const p = info.password;
  const s = agileSpec(p);
  if (s.cipher !== 'AES') throw new EncryptedFileError('unsupported', `its cipher (${s.cipher}) is not AES`);
  const name = cipherName(s.keyBits, s.chaining);
  const spun = spinPassword(password, s.salt, Number(p.spinCount) || 0, s.alg);
  const iv = fit(s.salt, s.blockSize);
  const input = aes(false, name, agileKey(spun, BLOCK_KEYS.verifierHashInput, s.keyBits, s.alg), iv, Buffer.from(p.encryptedVerifierHashInput || '', 'base64'));
  const value = aes(false, name, agileKey(spun, BLOCK_KEYS.verifierHashValue, s.keyBits, s.alg), iv, Buffer.from(p.encryptedVerifierHashValue || '', 'base64'));
  const expected = H(s.alg, input.subarray(0, s.salt.length || input.length));
  const got = value.subarray(0, s.hashSize);
  if (got.length !== expected.length || !timingSafeEqual(got, expected)) return null;
  const encryptedKey = Buffer.from(p.encryptedKeyValue || '', 'base64');
  const key = aes(false, name, agileKey(spun, BLOCK_KEYS.keyValue, s.keyBits, s.alg), iv, encryptedKey);
  const kd = agileSpec(info.keyData);
  return key.subarray(0, kd.keyBits / 8);
}

/** [MS-OFFCRYPTO] 2.3.4.14: the HMAC over the whole EncryptedPackage stream. */
function agileIntegrityHolds(info, secretKey, pkg) {
  const d = info.dataIntegrity;
  if (!d?.encryptedHmacKey || !d?.encryptedHmacValue) return false;
  const kd = agileSpec(info.keyData);
  const name = cipherName(kd.keyBits, kd.chaining);
  const ivKey = fit(H(kd.alg, kd.salt, BLOCK_KEYS.integrityKey), kd.blockSize);
  const ivValue = fit(H(kd.alg, kd.salt, BLOCK_KEYS.integrityValue), kd.blockSize);
  const hmacKey = aes(false, name, secretKey, ivKey, Buffer.from(d.encryptedHmacKey, 'base64')).subarray(0, kd.hashSize);
  const hmacValue = aes(false, name, secretKey, ivValue, Buffer.from(d.encryptedHmacValue, 'base64')).subarray(0, kd.hashSize);
  const actual = createHmac(kd.alg, hmacKey).update(pkg).digest();
  return actual.length === hmacValue.length && timingSafeEqual(actual, hmacValue);
}

function packageSize(pkg) {
  if (pkg.length < 8) throw new EncryptedFileError('damaged', 'the encrypted package is cut short');
  return pkg.readUInt32LE(0) + pkg.readUInt32LE(4) * 2 ** 32;
}

function agileDecryptPackage(info, secretKey, pkg) {
  const kd = agileSpec(info.keyData);
  const name = cipherName(kd.keyBits, kd.chaining);
  const size = packageSize(pkg);
  const body = pkg.subarray(8);
  const out = [];
  for (let i = 0, at = 0; at < body.length; i++, at += SEGMENT) {
    let seg = body.subarray(at, Math.min(body.length, at + SEGMENT));
    // A stream that stops mid-block is damaged at its last few bytes: decrypt
    // what is whole and let the zip reader judge the rest.
    const whole = seg.length - (seg.length % kd.blockSize);
    if (!whole) break;
    seg = seg.subarray(0, whole);
    const iv = fit(H(kd.alg, kd.salt, le32(i)), kd.blockSize);
    out.push(aes(false, name, secretKey, iv, seg));
  }
  const plain = Buffer.concat(out);
  if (plain.length < size) throw new EncryptedFileError('damaged', 'the encrypted package is shorter than it says');
  return plain.subarray(0, size);
}

/** [MS-OFFCRYPTO] 2.3.4.7: Standard Encryption's key from a password. */
export function standardKey(password, salt, keyBits = 128) {
  const h = spin('sha1', H('sha1', salt, Buffer.from(String(password), 'utf16le')), 50000);
  const final = H('sha1', h, le32(0));
  const x = (fill) => {
    const buf = Buffer.alloc(64, fill);
    for (let i = 0; i < final.length; i++) buf[i] ^= final[i];
    return H('sha1', buf);
  };
  return Buffer.concat([x(0x36), x(0x5c)]).subarray(0, keyBits / 8);
}

function standardDecrypt(info, password, pkg) {
  const key = standardKey(password, info.salt, info.keyBits);
  const verifier = aesEcb(false, key, info.encryptedVerifier);
  const hash = aesEcb(false, key, blocks(info.encryptedVerifierHash)).subarray(0, info.verifierHashSize || 20);
  const expected = H('sha1', verifier);
  if (hash.length !== expected.length || !timingSafeEqual(hash, expected)) return null;
  const size = packageSize(pkg);
  const body = pkg.subarray(8);
  const plain = aesEcb(false, key, body.subarray(0, body.length - (body.length % 16)));
  if (plain.length < size) throw new EncryptedFileError('damaged', 'the encrypted package is shorter than it says');
  return plain.subarray(0, size);
}

/**
 * The package inside a password-protected file.
 *
 * Throws EncryptedFileError: `password` when the password does not open it,
 * `tampered` when the HMAC says the encrypted bytes were changed after they
 * were written, `unsupported` for encryption Office does not use for a
 * password (rights management, a certificate), `damaged` for a file cut short.
 *
 * @returns {{ bytes: Buffer, kind: 'agile' | 'standard', integrity: boolean }}
 */
export function decryptPackage(bytes, password) {
  const { info: infoBytes, pkg } = streams(bytes);
  const info = readEncryptionInfo(infoBytes);
  let plain;
  let integrity = false;
  if (info.kind === 'agile') {
    const secretKey = agileSecretKey(info, password ?? '');
    if (!secretKey) throw new EncryptedFileError('password', 'the password is not right');
    // Checked before anything is decrypted: a changed byte is refused, not
    // opened as whatever it now decrypts to.
    if (info.dataIntegrity) {
      if (!agileIntegrityHolds(info, secretKey, pkg)) throw new EncryptedFileError('tampered', 'the file has been changed since it was encrypted');
      integrity = true;
    }
    plain = agileDecryptPackage(info, secretKey, pkg);
  } else {
    plain = standardDecrypt(info, password ?? '', pkg);
    if (!plain) throw new EncryptedFileError('password', 'the password is not right');
  }
  if (!(plain[0] === 0x50 && plain[1] === 0x4b)) {
    throw new EncryptedFileError(info.kind === 'standard' ? 'tampered' : 'damaged', 'what the password unlocked is not a document');
  }
  return { bytes: plain, kind: info.kind, integrity };
}

/* ── writing ───────────────────────────────────────────────────────────── */

/** What Office 2013 onwards writes. */
export const AGILE_DEFAULTS = { alg: 'sha512', hashAlgorithm: 'SHA512', keyBits: 256, hashSize: 64, blockSize: 16, saltSize: 16, spinCount: 100000 };

/**
 * The spun password hash for a salt, kept so that saving an encrypted
 * document every half minute does not spin 100,000 rounds each time. The
 * cache holds one entry, in memory, for the document it belongs to.
 */
function spunFor(password, spinCount, cache) {
  if (cache && cache.password === password && cache.spinCount === spinCount && cache.salt && cache.spun) return cache;
  const salt = randomBytes(AGILE_DEFAULTS.saltSize);
  const spun = spinPassword(password, salt, spinCount, AGILE_DEFAULTS.alg);
  const entry = { password, spinCount, salt, spun };
  if (cache) Object.assign(cache, entry);
  return entry;
}

/** A UNICODE-LP-P4 string: its length in bytes, UTF-16LE, padded to four. */
function lpp4(text) {
  const s = Buffer.from(text, 'utf16le');
  const len = Buffer.alloc(4);
  len.writeUInt32LE(s.length, 0);
  const pad = (4 - (s.length % 4)) % 4;
  return Buffer.concat([len, s, Buffer.alloc(pad)]);
}

const u32 = (...values) => Buffer.concat(values.map(le32));
const version10 = () => Buffer.from([1, 0, 0, 0]);

/**
 * The \x06DataSpaces storage ([MS-OFFCRYPTO] 2.1): the encrypted package is
 * mapped to a data space whose one transform is the encryption transform.
 */
export function dataSpacesStreams() {
  const DS = '\u0006DataSpaces';
  const version = Buffer.concat([lpp4('Microsoft.Container.DataSpaces'), version10(), version10(), version10()]);
  const entry = Buffer.concat([u32(1, 0), lpp4('EncryptedPackage'), lpp4('StrongEncryptionDataSpace')]);
  const map = Buffer.concat([u32(8, 1, entry.length + 4), entry]);
  const info = Buffer.concat([u32(8, 1), lpp4('StrongEncryptionTransform')]);
  const id = lpp4('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
  const primary = Buffer.concat([
    u32(8 + id.length, 1),
    id,
    lpp4('Microsoft.Container.EncryptionTransform'),
    version10(),
    version10(),
    version10(),
    // EncryptionTransformInfo: no name, block size and cipher mode 0 (the
    // EncryptionInfo stream says both), reserved 4.
    u32(0, 0, 0, 4),
  ]);
  return [
    { path: [DS, 'Version'], data: version },
    { path: [DS, 'DataSpaceMap'], data: map },
    { path: [DS, 'DataSpaceInfo', 'StrongEncryptionDataSpace'], data: info },
    { path: [DS, 'TransformInfo', 'StrongEncryptionTransform', '\u0006Primary'], data: primary },
  ];
}

/**
 * A document's package, encrypted with a password as Office 2013 onwards
 * encrypts one, in the compound file Word, Excel and PowerPoint open.
 *
 * @param {Uint8Array} plain the .docx/.xlsx/.pptx bytes
 * @param {string} password
 * @param {{ spinCount?: number, cache?: object }} [options] `cache`, an
 *   object kept beside the open document, spares re-spinning the password
 */
export function encryptPackage(plain, password, { spinCount = AGILE_DEFAULTS.spinCount, cache = null } = {}) {
  if (!password) throw new Error('a password is needed to encrypt a file');
  const d = AGILE_DEFAULTS;
  const name = cipherName(d.keyBits, 'ChainingModeCBC');
  const { salt: passwordSalt, spun } = spunFor(String(password), spinCount, cache);
  const iv = passwordSalt;

  // Fresh on every save: the key that encrypts the package, its salt, and
  // the verifier.
  const secretKey = randomBytes(d.keyBits / 8);
  const keyDataSalt = randomBytes(d.saltSize);
  const verifierInput = randomBytes(d.saltSize);

  const encryptedVerifierHashInput = aes(true, name, agileKey(spun, BLOCK_KEYS.verifierHashInput, d.keyBits, d.alg), iv, blocks(verifierInput, d.blockSize));
  const encryptedVerifierHashValue = aes(true, name, agileKey(spun, BLOCK_KEYS.verifierHashValue, d.keyBits, d.alg), iv, blocks(H(d.alg, verifierInput), d.blockSize));
  const encryptedKeyValue = aes(true, name, agileKey(spun, BLOCK_KEYS.keyValue, d.keyBits, d.alg), iv, blocks(secretKey, d.blockSize));

  // The package, in 4096-byte segments, each with the IV its index makes.
  const source = Buffer.from(plain.buffer, plain.byteOffset, plain.byteLength);
  const parts = [Buffer.alloc(8)];
  parts[0].writeUInt32LE(source.length % 2 ** 32, 0);
  parts[0].writeUInt32LE(Math.floor(source.length / 2 ** 32), 4);
  for (let i = 0, at = 0; at < source.length; i++, at += SEGMENT) {
    const seg = blocks(source.subarray(at, Math.min(source.length, at + SEGMENT)), d.blockSize);
    parts.push(aes(true, name, secretKey, fit(H(d.alg, keyDataSalt, le32(i)), d.blockSize), seg));
  }
  const pkg = Buffer.concat(parts);

  // dataIntegrity: an HMAC over the encrypted stream, its key and value
  // themselves encrypted with the package key.
  const hmacKey = randomBytes(d.hashSize);
  const hmacValue = createHmac(d.alg, hmacKey).update(pkg).digest();
  const encryptedHmacKey = aes(true, name, secretKey, fit(H(d.alg, keyDataSalt, BLOCK_KEYS.integrityKey), d.blockSize), blocks(hmacKey, d.blockSize));
  const encryptedHmacValue = aes(true, name, secretKey, fit(H(d.alg, keyDataSalt, BLOCK_KEYS.integrityValue), d.blockSize), blocks(hmacValue, d.blockSize));

  const common = `saltSize="${d.saltSize}" blockSize="${d.blockSize}" keyBits="${d.keyBits}" hashSize="${d.hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${d.hashAlgorithm}"`;
  const b64 = (b) => b.toString('base64');
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<encryption xmlns="${NS_ENCRYPTION}" xmlns:p="${NS_PASSWORD}" xmlns:c="${NS_CERTIFICATE}">` +
    `<keyData ${common} saltValue="${b64(keyDataSalt)}"/>` +
    `<dataIntegrity encryptedHmacKey="${b64(encryptedHmacKey)}" encryptedHmacValue="${b64(encryptedHmacValue)}"/>` +
    `<keyEncryptors><keyEncryptor uri="${NS_PASSWORD}">` +
    `<p:encryptedKey spinCount="${spinCount}" ${common} saltValue="${b64(passwordSalt)}" encryptedVerifierHashInput="${b64(encryptedVerifierHashInput)}" encryptedVerifierHashValue="${b64(encryptedVerifierHashValue)}" encryptedKeyValue="${b64(encryptedKeyValue)}"/>` +
    '</keyEncryptor></keyEncryptors></encryption>';
  // Version 4.4 and the reserved flag 0x40, then the XML.
  const info = Buffer.concat([Buffer.from([4, 0, 4, 0, 0x40, 0, 0, 0]), Buffer.from(xml, 'utf8')]);

  return writeCompoundFile([
    { path: ['EncryptionInfo'], data: info },
    { path: ['EncryptedPackage'], data: pkg },
    ...dataSpacesStreams(),
  ]);
}
