/**
 * Password-protected Word, Excel and PowerPoint 97-2003 files, opened with
 * their password.
 *
 * locked.doc, locked.xls, locked-95.xls and locked.ppt were saved with the
 * password "Rutba-1" by Word, Excel and PowerPoint themselves
 * (tools/make-locked-fixtures.ps1): RC4 CryptoAPI, 128-bit, for the 97-2003
 * formats, and XOR obfuscation for Excel 5.0/95. The Office here writes no
 * other kind, so the Office 97/2000 RC4 and the 40-bit CryptoAPI ones are
 * made in this file from the plain workbook and document, keyed as
 * [MS-OFFCRYPTO] says — written here on their own, not with the reader's code.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { binaryEncryption, decryptBinary, xorVerifier, xorKey } from '@rutba/office-formats/crypt-binary';
import { EncryptedFileError } from '@rutba/office-formats/crypt';
import { CompoundFile } from '@rutba/office-formats/cfb';
import { writeCompoundFile } from '@rutba/office-formats/cfb-write';
import { readXls } from '@rutba/office-formats/msxls';
import { xlsModelToXlsx } from '@rutba/office-formats/msxls-xlsx';
import { readPpt } from '@rutba/office-formats/msppt';
import { readWordDocument } from '@rutba/office-formats/msword';
import { SheetView } from '@rutba/sheet-view';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const fixture = (...p) => new Uint8Array(fs.readFileSync(path.join(HERE, 'fixtures', ...p)));
const PASSWORD = 'Rutba-1';

/** Every cell's text on the first sheet of a binary workbook, as the suite shows it. */
function cellsOf(xlsBytes) {
  const view = SheetView.open(Buffer.from(xlsModelToXlsx(readXls(xlsBytes))), { viewportWidth: 2000, viewportHeight: 2000 });
  return view.render().cells.filter((c) => c.text !== '').map((c) => `${c.ref}=${c.text}`).sort();
}
const wordsOf = (docBytes) => readWordDocument(docBytes).body.filter((b) => b.type === 'paragraph').map((p) => p.runs.map((r) => r.text).join(''));

const refuses = (fn, code) => assert.throws(fn, (err) => err instanceof EncryptedFileError && err.code === code);

test('each file Office locked says how, refuses a wrong password, and decrypted is the file without one', () => {
  for (const [name, app, method] of [['locked.doc', 'doc', 'cryptoapi'], ['locked.xls', 'xls', 'cryptoapi'], ['locked-95.xls', 'xls', 'xor'], ['locked.ppt', 'ppt', 'cryptoapi']]) {
    const bytes = fixture('binary', name);
    assert.deepEqual(binaryEncryption(bytes), { app, method }, name);
    refuses(() => decryptBinary(bytes, 'rutba-1'), 'password');
    const plain = decryptBinary(bytes, PASSWORD);
    assert.equal(binaryEncryption(plain), null, `${name} decrypted is not locked`);
    assert.equal(new CompoundFile(plain).application(), app);
  }
  assert.equal(binaryEncryption(fixture('binary', 'text.doc')), null, 'a file with no password is not');
  assert.equal(binaryEncryption(fixture('binary', 'showcase.xls')), null);
});

test('a locked Word document opens to its words, its bold kept', () => {
  const plain = decryptBinary(fixture('binary', 'locked.doc'), PASSWORD);
  assert.deepEqual(wordsOf(plain), ['A locked document', 'Its words are kept from anyone without the password, and open with it.', 'A second paragraph, in bold.']);
  const third = readWordDocument(plain).body.filter((b) => b.type === 'paragraph')[2];
  assert.ok(third.runs.every((r) => r.chp?.bold), 'the third paragraph in bold');
});

test('a locked workbook opens to the cells it was saved with — RC4 CryptoAPI and Excel 95\'s XOR alike', () => {
  const before = cellsOf(fixture('binary', 'conditions.xls'));
  assert.ok(before.length > 10, 'the plain workbook has cells to compare');
  assert.deepEqual(cellsOf(decryptBinary(fixture('binary', 'locked.xls'), PASSWORD)), before);
  const book = readXls(decryptBinary(fixture('binary', 'locked-95.xls'), PASSWORD));
  assert.equal(book.biff, 5);
  assert.deepEqual(book.sheets.map((s) => s.name), readXls(fixture('binary', 'conditions.xls')).sheets.map((s) => s.name), 'the sheet names, whose bytes start a record that is half plain');
  assert.deepEqual(cellsOf(decryptBinary(fixture('binary', 'locked-95.xls'), PASSWORD)), before);
});

test('XOR obfuscation: the verifier and key Excel 95 kept for the password', () => {
  const s = new CompoundFile(fixture('binary', 'locked-95.xls'));
  const book = s.read(s.find(['Book']));
  const at = 4 + (book[2] | (book[3] << 8)) + 4; // the FILEPASS record's data, after the BOF
  assert.equal(xorKey(PASSWORD), book[at] | (book[at + 1] << 8));
  assert.equal(xorVerifier(PASSWORD), book[at + 2] | (book[at + 3] << 8));
});

test('a locked presentation opens to its slides and its picture, byte for byte', () => {
  const model = readPpt(decryptBinary(fixture('binary', 'locked.ppt'), PASSWORD));
  assert.equal(model.slides.length, 2);
  const words = model.slides.map((s) => s.shapes.map((sh) => sh.words).filter(Boolean));
  assert.deepEqual(words[0], ['A locked deck', 'Opened with its password']);
  assert.equal(words[1][0], 'Second slide');
  assert.equal(model.images.length, 1);
  assert.ok(Buffer.from(model.images[0].bytes).equals(fs.readFileSync(path.join(HERE, 'fixtures', 'rich', 'picture.png'))), 'the picture as it was put in');
});

/* ── the kinds this Office does not write, made from the spec ─────────────── */

const md5 = (b) => crypto.createHash('md5').update(b).digest();
const sha1 = (b) => crypto.createHash('sha1').update(b).digest();
const le16 = (n) => { const b = Buffer.alloc(2); b.writeUInt16LE(n); return b; };
const le32 = (n) => { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0); return b; };

function rc4(key, data) {
  const s = [...Array(256).keys()];
  for (let i = 0, j = 0; i < 256; i++) { j = (j + s[i] + key[i % key.length]) & 255; [s[i], s[j]] = [s[j], s[i]]; }
  const out = Buffer.from(data);
  for (let k = 0, i = 0, j = 0; k < out.length; k++) {
    i = (i + 1) & 255; j = (j + s[i]) & 255; [s[i], s[j]] = [s[j], s[i]];
    out[k] ^= s[(s[i] + s[j]) & 255];
  }
  return out;
}

/** An encryption header for a password, and the key for each block: [MS-OFFCRYPTO] 2.3.6 (rc4) or 2.3.5 (cryptoapi). */
function lockHeader(kind, password, keyBits = 128) {
  const pw = Buffer.from(password, 'utf16le');
  const salt = crypto.randomBytes(16);
  let keyFor;
  if (kind === 'rc4') {
    const h1 = md5(Buffer.concat(Array.from({ length: 16 }, () => Buffer.concat([md5(pw).subarray(0, 5), salt])))).subarray(0, 5);
    keyFor = (b) => md5(Buffer.concat([h1, le32(b)]));
  } else {
    const h0 = sha1(Buffer.concat([salt, pw]));
    keyFor = (b) => { const h = sha1(Buffer.concat([h0, le32(b)])); return keyBits === 40 ? Buffer.concat([h.subarray(0, 5), Buffer.alloc(11)]) : h.subarray(0, keyBits / 8); };
  }
  const verifier = crypto.randomBytes(16);
  const sealed = rc4(keyFor(0), Buffer.concat([verifier, kind === 'rc4' ? md5(verifier) : sha1(verifier)]));
  if (kind === 'rc4') return { keyFor, bytes: Buffer.concat([le16(1), le16(1), salt, sealed]) };
  const csp = Buffer.from('Microsoft Base Cryptographic Provider v1.0\0', 'utf16le');
  const head = Buffer.concat([le32(0x04), le32(0), le32(0x6801), le32(0x8004), le32(keyBits), le32(1), le32(0), le32(0), csp]);
  return { keyFor, bytes: Buffer.concat([le16(2), le16(2), le32(0x04), le32(head.length), head, le32(16), salt, sealed.subarray(0, 16), le32(20), sealed.subarray(16)]) };
}

/** XOR a stream's range with the keystream for its place, a new key every `size` bytes. */
function sealRange(buf, keyFor, size, from, to) {
  for (let p = from; p < to;) {
    const block = Math.floor(p / size);
    const end = Math.min(to, (block + 1) * size);
    const stream = rc4(keyFor(block), Buffer.alloc(end - block * size));
    for (let k = p; k < end; k++) buf[k] ^= stream[k - block * size];
    p = end;
  }
}

/** A compound file with some of its root streams replaced. */
function rewrite(bytes, replaced) {
  const cfb = new CompoundFile(bytes);
  const streams = [];
  const walk = (entry, at) => {
    for (const c of cfb.childrenOf(entry)) {
      if (c.type === 2) streams.push({ path: [...at, c.name], data: !at.length && replaced[c.name] ? replaced[c.name] : cfb.read(c) });
      else if (c.type === 1) walk(c, [...at, c.name]);
    }
  };
  walk(cfb.root, []);
  return writeCompoundFile(streams);
}

/** A plain BIFF8 workbook locked: a FilePass after the BOF, the sheets' starts moved, every record after it sealed but those never sealed. */
function lockXls(plain, kind, keyBits) {
  const cfb = new CompoundFile(plain);
  const s = Buffer.from(cfb.read(cfb.find(['Workbook'])));
  const { keyFor, bytes } = lockHeader(kind, PASSWORD, keyBits);
  const pass = Buffer.concat([le16(0x2f), le16(2 + bytes.length), le16(1), bytes]);
  const bofEnd = 4 + s.readUInt16LE(2);
  const out = Buffer.concat([s.subarray(0, bofEnd), pass, s.subarray(bofEnd)]);
  const never = new Set([0x0809, 0x002f, 0x0194, 0x0195, 0x00e1, 0x0196, 0x0138]);
  for (let p = 0; p + 4 <= out.length;) {
    const id = out.readUInt16LE(p);
    const size = out.readUInt16LE(p + 2);
    if (id === 0x85) out.writeUInt32LE(out.readUInt32LE(p + 4) + pass.length, p + 4);
    if (p > bofEnd && !never.has(id)) sealRange(out, keyFor, 1024, p + 4 + (id === 0x85 ? 4 : 0), p + 4 + size);
    p += 4 + size;
  }
  return rewrite(plain, { Workbook: out });
}

test('Office 97/2000\'s RC4 and the 40-bit CryptoAPI open too', () => {
  const plain = decryptBinary(fixture('binary', 'locked.xls'), PASSWORD);
  const before = cellsOf(plain);
  for (const [kind, bits] of [['rc4', 128], ['cryptoapi', 40], ['cryptoapi', 56]]) {
    const locked = lockXls(plain, kind, bits);
    assert.deepEqual(binaryEncryption(locked), { app: 'xls', method: kind }, `${kind} ${bits}`);
    refuses(() => decryptBinary(locked, 'wrong'), 'password');
    assert.deepEqual(cellsOf(decryptBinary(locked, PASSWORD)), before, `${kind} ${bits}`);
  }
  // A Word document under Office 97/2000's RC4: the header at the start of the table stream, the streams sealed in 512-byte blocks.
  const doc = decryptBinary(fixture('binary', 'locked.doc'), PASSWORD);
  const cfb = new CompoundFile(doc);
  const wd = Buffer.from(cfb.read(cfb.find(['WordDocument'])));
  const tableName = wd.readUInt16LE(0x0a) & 0x0200 ? '1Table' : '0Table';
  const table = Buffer.from(cfb.read(cfb.find([tableName])));
  const { keyFor, bytes } = lockHeader('rc4', PASSWORD);
  bytes.copy(table, 0);
  wd.writeUInt16LE(wd.readUInt16LE(0x0a) | 0x0100, 0x0a);
  wd.writeUInt32LE(bytes.length, 0x0e);
  const head = Buffer.from(wd.subarray(0, 0x44));
  sealRange(wd, keyFor, 512, 0, wd.length);
  head.copy(wd, 0);
  sealRange(table, keyFor, 512, bytes.length, table.length);
  const lockedDoc = rewrite(doc, { WordDocument: wd, [tableName]: table });
  assert.deepEqual(binaryEncryption(lockedDoc), { app: 'doc', method: 'rc4' });
  assert.deepEqual(wordsOf(decryptBinary(lockedDoc, PASSWORD)), wordsOf(doc));
});

/* ── through the suite ───────────────────────────────────────────────────── */

test('the suite asks for a locked .doc, .xls or .ppt\'s password, opens it with it, and keeps it for the save', () => {
  const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-locked-'));
  for (const [name, kind, from] of [['locked.doc', 'doc', 'doc'], ['locked.xls', 'sheet', 'xls'], ['locked-95.xls', 'sheet', 'xls'], ['locked.ppt', 'deck', 'ppt']]) {
    const file = path.join(dir, name);
    fs.writeFileSync(file, fixture('binary', name));
    assert.deepEqual(service.open({ path: file }), { locked: true, name, wrong: false, path: file });
    assert.equal(service.open({ path: file, password: 'nope' }).wrong, true);
    const opened = service.open({ path: file, password: PASSWORD });
    assert.equal(opened.kind, kind, name);
    assert.equal(opened.converted.from, from);
    assert.ok(!opened.converted.partial, `${name} read in full`);
    assert.equal(opened.encrypted, true, `${name} is saved under its password`);
    service.close({ id: opened.id });
  }
});
