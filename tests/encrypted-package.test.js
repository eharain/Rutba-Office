/**
 * Password-protected Office files, [MS-OFFCRYPTO]: the engine.
 *
 * A file another program encrypted opens with its password and is refused
 * without it; what this suite encrypts decrypts to the same bytes; a byte
 * changed in the encrypted stream is caught by the HMAC; and Office 2007's
 * Standard Encryption — built here from the spec's own steps, written out
 * independently of the engine — opens too.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createCipheriv, createHash, randomBytes } from 'node:crypto';
import { decryptPackage, encryptPackage, isEncryptedPackage, readEncryptionInfo, EncryptedFileError, spinPassword, agileKey, BLOCK_KEYS } from '@rutba/office-formats/crypt';
import { CompoundFile } from '@rutba/office-formats/cfb';
import { writeCompoundFile } from '@rutba/office-formats/cfb-write';
import { readZip } from '@rutba/ooxml/zip';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';

const FIXTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures');

const code = (fn) => {
  try {
    fn();
    return null;
  } catch (err) {
    assert.ok(err instanceof EncryptedFileError, `an EncryptedFileError, not ${err}`);
    return err.code;
  }
};
const names = (zip) => readZip(Buffer.from(zip)).entries.map((e) => e.name);
const infoOf = (bytes) => {
  const cfb = new CompoundFile(bytes);
  return Buffer.from(cfb.read(cfb.find(['EncryptionInfo'])));
};

test('a file another program encrypted opens with its password and only with it', () => {
  // The sample site's password-protected document (password 123): Agile
  // Encryption with SHA-1 and AES-128, as Office 2010 writes it. The
  // verifier passing is the known answer for the spun key derivation, and
  // the HMAC passing the known answer for dataIntegrity.
  const bytes = fs.readFileSync(path.join(FIXTURES, 'encrypted.docx'));
  assert.equal(isEncryptedPackage(bytes), true);
  const info = readEncryptionInfo(infoOf(bytes));
  assert.equal(info.kind, 'agile');
  assert.equal(info.password.hashAlgorithm, 'SHA1');
  assert.equal(info.password.spinCount, '100000');

  const opened = decryptPackage(bytes, '123');
  assert.equal(opened.kind, 'agile');
  assert.equal(opened.integrity, true, 'the HMAC was checked and holds');
  assert.ok(names(opened.bytes).includes('word/document.xml'));
  assert.ok(names(opened.bytes).includes('[Content_Types].xml'));

  assert.equal(code(() => decryptPackage(bytes, '1234')), 'password');
  assert.equal(code(() => decryptPackage(bytes, '')), 'password');
  assert.equal(code(() => decryptPackage(bytes, '12 3')), 'password');
});

test('Agile key derivation follows the spec step by step', () => {
  // [MS-OFFCRYPTO] 2.3.4.11: H0 = H(salt + password), Hn = H(iterator + Hn-1),
  // then H(Hfinal + blockKey) cut to the key's length. Written out here
  // with nothing from the engine but the answer.
  const salt = Buffer.from('000102030405060708090a0b0c0d0e0f', 'hex');
  let h = createHash('sha512').update(salt).update(Buffer.from('Pässword', 'utf16le')).digest();
  for (let i = 0; i < 1000; i++) {
    const it = Buffer.alloc(4);
    it.writeUInt32LE(i, 0);
    h = createHash('sha512').update(it).update(h).digest();
  }
  assert.deepEqual(spinPassword('Pässword', salt, 1000, 'sha512'), h);
  const key = createHash('sha512').update(h).update(BLOCK_KEYS.keyValue).digest().subarray(0, 32);
  assert.deepEqual(agileKey(h, BLOCK_KEYS.keyValue, 256, 'sha512'), key);
  // A key longer than the hash is padded with 0x36, as the spec says.
  const short = agileKey(Buffer.alloc(20), BLOCK_KEYS.keyValue, 256, 'sha1');
  assert.equal(short.length, 32);
  assert.deepEqual(short.subarray(20), Buffer.alloc(12, 0x36));
  // The block keys are the spec's constants.
  assert.equal(BLOCK_KEYS.verifierHashInput.toString('hex'), 'fea7d2763b4b9e79');
  assert.equal(BLOCK_KEYS.verifierHashValue.toString('hex'), 'd7aa0f6d3061344e');
  assert.equal(BLOCK_KEYS.keyValue.toString('hex'), '146e0be7abacd0d6');
  assert.equal(BLOCK_KEYS.integrityKey.toString('hex'), '5fb2ad010cb9e1f6');
  assert.equal(BLOCK_KEYS.integrityValue.toString('hex'), 'a0677f02b22c8433');
});

test('a document, a workbook and a presentation round-trip through Encrypt with Password', () => {
  const cache = {};
  const packages = {
    docx: buildDocx({ paragraphs: ['Quarterly figures, for the board only.', 'Second paragraph.'], styles: true }),
    xlsx: buildXlsx({ sheets: [{ name: 'Budget', rows: [['Item', 'Cost'], ['Rent', 1200], ['Total', '=SUM(B2:B2)']] }] }),
    pptx: buildPptx({ title: 'Plan', slides: [{ layout: 'title', title: 'Plan', body: 'Not for circulation' }] }),
  };
  for (const [kind, plain] of Object.entries(packages)) {
    const encrypted = encryptPackage(plain, 'correct horse', { cache });
    assert.equal(isEncryptedPackage(encrypted), true, `${kind} is a compound file holding the encrypted package`);
    assert.equal(Buffer.from(encrypted).includes(Buffer.from(plain).subarray(30, 60)), false, `${kind}: none of the package in the clear`);
    const back = decryptPackage(encrypted, 'correct horse');
    assert.equal(back.integrity, true);
    assert.deepEqual(Buffer.from(back.bytes), Buffer.from(plain), `${kind} decrypts to the same bytes`);
    assert.equal(code(() => decryptPackage(encrypted, 'Correct horse')), 'password', `${kind}: a wrong password is refused`);
  }
});

test('what is written is Agile Encryption as Office 2013 onwards writes it', () => {
  const plain = buildDocx({ paragraphs: ['x'] });
  const one = encryptPackage(plain, 'pw');
  const two = encryptPackage(plain, 'pw');
  const xml = infoOf(one).subarray(8).toString('utf8');
  assert.deepEqual([...infoOf(one).subarray(0, 8)], [4, 0, 4, 0, 0x40, 0, 0, 0], 'version 4.4, reserved 0x40');
  assert.match(xml, /^<\?xml version="1\.0" encoding="UTF-8" standalone="yes"\?>\r\n<encryption /);
  assert.match(xml, /<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="/);
  assert.match(xml, /<dataIntegrity encryptedHmacKey="[^"]+" encryptedHmacValue="[^"]+"\/>/);
  assert.match(xml, /<p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64"/);
  // Fresh salts and a fresh key on every save: two saves share nothing.
  const salts = (bytes) => [...infoOf(bytes).subarray(8).toString('utf8').matchAll(/saltValue="([^"]+)"/g)].map((m) => m[1]);
  const [a, b] = [salts(one), salts(two)];
  assert.notEqual(a[0], b[0]);
  assert.notEqual(a[1], b[1]);
  // The \x06DataSpaces storage Office writes beside the package.
  const cfb = new CompoundFile(one);
  for (const p of [['\u0006DataSpaces', 'Version'], ['\u0006DataSpaces', 'DataSpaceMap'], ['\u0006DataSpaces', 'DataSpaceInfo', 'StrongEncryptionDataSpace'], ['\u0006DataSpaces', 'TransformInfo', 'StrongEncryptionTransform', '\u0006Primary']]) {
    assert.ok(cfb.find(p), `${p.join('/')} is there`);
  }
});

test('a byte changed in the encrypted package is refused by the HMAC', () => {
  const plain = buildXlsx({ sheets: [{ name: 'S', rows: Array.from({ length: 400 }, (_, i) => ['Row ' + i, i * 3.25]) }] });
  const encrypted = Buffer.from(encryptPackage(plain, 'pw', { spinCount: 1000 }));
  const cfb = new CompoundFile(encrypted);
  const pkg = cfb.find(['EncryptedPackage']);
  assert.ok(pkg.size >= 4096, 'the package is in regular sectors');
  const at = (pkg.start + 1) * 512 + 700;
  for (const flip of [0x01, 0x80]) {
    const tampered = Buffer.from(encrypted);
    tampered[at] ^= flip;
    assert.equal(code(() => decryptPackage(tampered, 'pw')), 'tampered');
  }
  // The size in front of the package is inside the HMAC too.
  const resized = Buffer.from(encrypted);
  resized[(pkg.start + 1) * 512] ^= 0x02;
  assert.equal(code(() => decryptPackage(resized, 'pw')), 'tampered');
  // And the wrong password is still the wrong password, not a tampered file.
  assert.equal(code(() => decryptPackage(encrypted, 'nope')), 'password');
  assert.equal(decryptPackage(encrypted, 'pw').integrity, true);
});

/* ── Standard Encryption, built from [MS-OFFCRYPTO] 2.3.4.5–2.3.4.9 ───── */

const sha1 = (...parts) => {
  const h = createHash('sha1');
  for (const p of parts) h.update(p);
  return h.digest();
};
const u32 = (n) => {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(n, 0);
  return b;
};
const ecb = (key, data) => {
  const c = createCipheriv('aes-128-ecb', key, null);
  c.setAutoPadding(false);
  return Buffer.concat([c.update(data), c.final()]);
};

function standardEncrypt(plain, password) {
  const salt = randomBytes(16);
  // 2.3.4.7: H0 = SHA-1(salt + password); Hn = SHA-1(iterator + Hn-1), 50,000
  // times; Hfinal = SHA-1(Hn + block 0); X1/X2 from 0x36 and 0x5C buffers.
  let h = sha1(salt, Buffer.from(password, 'utf16le'));
  for (let i = 0; i < 50000; i++) h = sha1(u32(i), h);
  const final = sha1(h, u32(0));
  const derive = (fill) => {
    const buf = Buffer.alloc(64, fill);
    for (let i = 0; i < 20; i++) buf[i] ^= final[i];
    return sha1(buf);
  };
  const key = Buffer.concat([derive(0x36), derive(0x5c)]).subarray(0, 16);
  // 2.3.4.8: the verifier, and its SHA-1 padded to 32 bytes, both encrypted.
  const verifier = randomBytes(16);
  const verifierHash = Buffer.concat([sha1(verifier), Buffer.alloc(12)]);
  // 2.3.2: the EncryptionHeader — fCryptoAPI | fAES, AES-128, SHA-1, the
  // AES provider type and its name.
  const csp = Buffer.from('Microsoft Enhanced RSA and AES Cryptographic Provider\0', 'utf16le');
  const header = Buffer.concat([u32(0x24), u32(0), u32(0x660e), u32(0x8004), u32(128), u32(0x18), u32(0), u32(0), csp]);
  const info = Buffer.concat([
    Buffer.from([3, 0, 2, 0]),
    u32(0x24),
    u32(header.length),
    header,
    u32(16),
    salt,
    ecb(key, verifier),
    u32(20),
    ecb(key, verifierHash),
  ]);
  const padded = Buffer.concat([plain, Buffer.alloc((16 - (plain.length % 16)) % 16)]);
  const size = Buffer.alloc(8);
  size.writeUInt32LE(plain.length, 0);
  const pkg = Buffer.concat([size, ecb(key, padded)]);
  return writeCompoundFile([
    { path: ['EncryptionInfo'], data: info },
    { path: ['EncryptedPackage'], data: pkg },
  ]);
}

test('Standard Encryption (Office 2007) opens with its password and only with it', () => {
  const plain = Buffer.from(buildDocx({ paragraphs: ['Encrypted the 2007 way.'] }));
  const file = standardEncrypt(plain, 'Secret 2007');
  assert.equal(isEncryptedPackage(file), true);
  assert.equal(readEncryptionInfo(infoOf(file)).kind, 'standard');
  const opened = decryptPackage(file, 'Secret 2007');
  assert.equal(opened.kind, 'standard');
  assert.equal(opened.integrity, false, 'Standard Encryption carries no HMAC');
  assert.deepEqual(Buffer.from(opened.bytes), plain);
  assert.equal(code(() => decryptPackage(file, 'secret 2007')), 'password');
});

test('encryption a password cannot open is refused by name', () => {
  const plain = buildDocx({ paragraphs: ['x'] });
  const encrypted = encryptPackage(plain, 'pw', { spinCount: 10 });
  const cfb = new CompoundFile(encrypted);
  const info = Buffer.from(cfb.read(cfb.find(['EncryptionInfo'])));
  const pkg = Buffer.from(cfb.read(cfb.find(['EncryptedPackage'])));
  const rebuilt = (infoBytes) => writeCompoundFile([{ path: ['EncryptionInfo'], data: infoBytes }, { path: ['EncryptedPackage'], data: pkg }]);
  // Extensible Encryption (rights management) is version 4.3.
  const irm = Buffer.from(info);
  irm.writeUInt16LE(3, 2);
  assert.equal(code(() => decryptPackage(rebuilt(irm), 'pw')), 'unsupported');
  // A certificate key encryptor with no password one.
  const cert = Buffer.from(info.toString('utf8').replace('keyEncryptor/password"><p:encryptedKey', 'keyEncryptor/certificate"><c:encryptedKey').replace('</p:encryptedKey>', ''), 'utf8');
  assert.equal(code(() => decryptPackage(rebuilt(cert), 'pw')), 'unsupported');
});
