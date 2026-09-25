/**
 * Encrypt with Password, through the document service: what a window's
 * open, Info and Save do to a protected file, with no Electron.
 *
 * A protected file answers `locked` until its password is given; File →
 * Info's password makes every save an encrypted one, and clearing it goes
 * back to a plain save; the autosave copy of a protected document is
 * encrypted with the same password; and the password is never written
 * anywhere in the clear.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx } from '@rutba/presentation';
import { CompoundFile } from '@rutba/office-formats/cfb';
import { isEncryptedPackage } from '@rutba/office-formats/crypt';

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-encrypted-'));
const recoveryDir = path.join(dir, 'recovery');
const doc = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }), recoveryDir });
const write = (name, bytes) => {
  const p = path.join(dir, name);
  fs.writeFileSync(p, bytes);
  return p;
};
const textOf = (model) => (model.blocks || []).map((b) => (b.runs || []).map((r) => r.text).join('') || b.text || '').join('\n');
const isZip = (bytes) => bytes[0] === 0x50 && bytes[1] === 0x4b;
const PASSWORD = 'Tr0ub4dor&3';
/** Is the password anywhere in these bytes, as UTF-8 or UTF-16? */
const holdsPassword = (bytes) => Buffer.from(bytes).includes(Buffer.from(PASSWORD, 'utf8')) || Buffer.from(bytes).includes(Buffer.from(PASSWORD, 'utf16le'));

test('Encrypt with Password makes Save write an encrypted file, and clearing it a plain one', () => {
  const file = write('board.docx', buildDocx({ paragraphs: ['For the board only.'] }));
  const opened = doc.open({ path: file, kind: 'doc' });
  assert.equal(opened.encrypted, false);

  const meta = doc.setPassword({ id: opened.id, password: PASSWORD });
  assert.equal(meta.encrypted, true);
  assert.equal(meta.dirty, true, 'setting a password is a change to save');
  doc.save({ id: opened.id });
  const saved = fs.readFileSync(file);
  assert.equal(isEncryptedPackage(saved), true, 'the file on disk is encrypted');
  assert.equal(holdsPassword(saved), false);
  assert.equal(saved.includes(Buffer.from('For the board only.')), false);

  // Saved again after an edit: still encrypted, with the same password.
  doc.apply({ id: opened.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'Draft: ' }] });
  doc.save({ id: opened.id });
  assert.equal(isEncryptedPackage(fs.readFileSync(file)), true);
  doc.close({ id: opened.id });

  // Opening it asks, refuses the wrong password, and opens with the right one.
  assert.equal(doc.open({ path: file, kind: 'doc' }).locked, true);
  const wrong = doc.open({ path: file, kind: 'doc', password: 'tr0ub4dor&3' });
  assert.equal(wrong.locked, true);
  assert.equal(wrong.wrong, true);
  const again = doc.open({ path: file, kind: 'doc', password: PASSWORD });
  assert.equal(again.encrypted, true, 'opened with a password, saved with it');
  assert.match(textOf(again.model), /Draft: For the board only\./);

  // Clearing the password goes back to a plain save.
  assert.equal(doc.setPassword({ id: again.id, password: '' }).encrypted, false);
  doc.save({ id: again.id });
  assert.equal(isZip(fs.readFileSync(file)), true, 'a plain .docx again');
  doc.close({ id: again.id });
  const plain = doc.open({ path: file, kind: 'doc' });
  assert.equal(plain.locked, undefined);
  assert.equal(plain.encrypted, false);
  doc.close({ id: plain.id });
});

test('a workbook and a presentation are saved encrypted and open again', () => {
  const cases = [
    ['figures.xlsx', 'sheet', buildXlsx({ sheets: [{ name: 'S', rows: [['Salary', 91000]] }] })],
    ['plan.pptx', 'deck', buildPptx({ title: 'Plan', slides: [{ layout: 'title', title: 'Plan', body: 'Secret' }] })],
  ];
  for (const [name, kind, bytes] of cases) {
    const file = write(name, bytes);
    const opened = doc.open({ path: file, kind });
    doc.setPassword({ id: opened.id, password: PASSWORD });
    const saved = doc.save({ id: opened.id });
    assert.equal(saved.encrypted, true);
    doc.close({ id: opened.id });
    const disk = fs.readFileSync(file);
    assert.equal(isEncryptedPackage(disk), true, `${name} is encrypted on disk`);
    assert.ok(new CompoundFile(disk).find(['\u0006DataSpaces', 'TransformInfo', 'StrongEncryptionTransform', '\u0006Primary']), `${name} carries \\x06DataSpaces`);
    assert.equal(doc.open({ path: file, kind }).locked, true);
    const back = doc.open({ path: file, kind, password: PASSWORD });
    assert.equal(back.kind, kind);
    assert.equal(back.encrypted, true);
    doc.close({ id: back.id });
  }
});

test('the autosave copy of a protected document is encrypted with its password', () => {
  const file = write('memo.docx', buildDocx({ paragraphs: ['Private memo.'] }));
  const opened = doc.open({ path: file, kind: 'doc' });
  doc.setPassword({ id: opened.id, password: PASSWORD });
  doc.apply({ id: opened.id, ops: [{ op: 'setSelection', anchor: { block: 0, offset: 0 }, focus: { block: 0, offset: 0 } }, { op: 'insertText', text: 'Unsaved ' }] });
  const [entry] = doc.autosave().filter((e) => e.name === 'memo.docx');
  assert.ok(entry, 'a copy was written');
  assert.equal(entry.encrypted, true);
  const copy = fs.readFileSync(path.join(recoveryDir, entry.file));
  assert.equal(isEncryptedPackage(copy), true, 'the copy is encrypted, never the document in the clear');
  assert.equal(copy.includes(Buffer.from('Unsaved')), false);
  for (const name of fs.readdirSync(recoveryDir)) {
    assert.equal(holdsPassword(fs.readFileSync(path.join(recoveryDir, name))), false, `${name} does not hold the password`);
  }

  // After a crash: a new service, the copy offered, and the password asked for.
  const after = createDocumentService({ holdBlob: () => ({}), recoveryDir });
  const offered = after.recoverable().find((e) => e.file === entry.file);
  assert.ok(offered);
  assert.equal(after.recover({ file: entry.file }).locked, true);
  assert.equal(after.recover({ file: entry.file, password: 'nope' }).wrong, true);
  const recovered = after.recover({ file: entry.file, password: PASSWORD });
  assert.equal(recovered.encrypted, true, 'and it saves encrypted again');
  assert.equal(recovered.dirty, true);
  assert.match(textOf(recovered.model), /Unsaved Private memo\./);
  after.close({ id: recovered.id });
  doc.close({ id: opened.id });
});

test('a protected file changed after it was encrypted is refused with a sentence', () => {
  const file = write('tampered.xlsx', buildXlsx({ sheets: [{ name: 'S', rows: Array.from({ length: 400 }, (_, i) => ['Row ' + i, i * 3.25]) }] }));
  const opened = doc.open({ path: file, kind: 'sheet' });
  doc.setPassword({ id: opened.id, password: PASSWORD });
  doc.save({ id: opened.id });
  doc.close({ id: opened.id });
  const bytes = fs.readFileSync(file);
  const cfb = new CompoundFile(bytes);
  const pkg = cfb.find(['EncryptedPackage']);
  bytes[(pkg.start + 1) * 512 + 64] ^= 0x10;
  fs.writeFileSync(file, bytes);
  assert.throws(() => doc.open({ path: file, kind: 'sheet', password: PASSWORD }), /tampered\.xlsx was not opened: it has been changed since its password was set/);
  // Without the password it only asks: nothing is judged before it is given.
  assert.equal(doc.open({ path: file, kind: 'sheet' }).locked, true);
});
