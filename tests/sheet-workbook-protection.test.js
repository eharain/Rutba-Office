/**
 * Review → Protect Workbook and Allow Edit Ranges — the engine's half.
 *
 * A password is hashed the way Excel hashes one: SHA-512 over the salt and
 * the password as UTF-16LE, then 100,000 rounds each over the last hash and
 * the round's number as four little-endian bytes. The expected value below
 * was worked out with a second, separate implementation (.NET's SHA512 in
 * PowerShell) from the same documented steps, so the two agree only if both
 * follow the algorithm. The legacy 16-bit hash is checked against the
 * values Excel writes for "password" (83AF) and "secret" (DAA7).
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { hashPassword, legacyPasswordHash, checkPassword } from '@rutba/ooxml/protection';
import { SheetView } from '@rutba/sheet-view';
import { WRONG_PASSWORD, WORKBOOK_LOCKED } from '../packages/sheet-view/src/view.js';

const SALT = 'AAECAwQFBgcICQoLDA0ODw==';
const RUTBA_HASH = 'VuujxQaYHeJ14+FUG9kIYQMcmdZUeAqNzfZavjHt5PbCYp6Sc2+KgimXS6nhmw/lRB0EHrODhQMGCqig3NYTYQ==';

const book = () => buildXlsx({ sheets: [
  { name: 'Data', rows: [['Item', 'Qty'], ['Pens', 4], ['Ink', 2]] },
  { name: 'Summary', rows: [['=SUM(Data!B2:B3)']] },
  { name: 'Notes', rows: [['n']] },
] });

/** A workbook whose workbook.xml carries `el` before its sheets, as Excel would write it. */
function withWorkbookXml(bytes, el) {
  const pkg = OoxmlPackage.read(bytes);
  pkg.write_('xl/workbook.xml', pkg.text('xl/workbook.xml').replace('<sheets>', el + '<sheets>'));
  return pkg.write();
}

test('the SHA-512 password hash is the one the documented algorithm gives, round for round', () => {
  assert.equal(hashPassword('Rutba', { salt: SALT, spinCount: 100000 }), RUTBA_HASH);
  assert.notEqual(hashPassword('rutba', { salt: SALT, spinCount: 100000 }), RUTBA_HASH, 'capitals count');
  assert.equal(checkPassword('Rutba', { algorithmName: 'SHA-512', hashValue: RUTBA_HASH, saltValue: SALT, spinCount: '100000' }), true);
});

test('the legacy 16-bit hash is the one Excel writes, and still opens an old file', () => {
  assert.equal(legacyPasswordHash('password'), '83AF');
  assert.equal(legacyPasswordHash('secret'), 'DAA7');
  const view = new SheetView(withWorkbookXml(book(), '<workbookProtection workbookPassword="83AF" lockStructure="1"/>'));
  assert.deepEqual(view.workbookProtection(), { structure: true, windows: false, hasPassword: true });
  assert.throws(() => view.unprotectWorkbook({ password: 'Password' }), (e) => e.message === WRONG_PASSWORD);
  view.unprotectWorkbook({ password: 'password' });
  assert.equal(view.workbookProtection().structure, false);
  assert.doesNotMatch(OoxmlPackage.read(view.save()).text('xl/workbook.xml'), /workbookProtection/);
});

test('a workbook Excel locked with a modern password is unlocked with it and refused without', () => {
  const el = `<workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="${RUTBA_HASH}" workbookSaltValue="${SALT}" workbookSpinCount="100000" lockStructure="1"/>`;
  const view = new SheetView(withWorkbookXml(book(), el));
  assert.equal(view.workbookProtection().hasPassword, true);
  assert.throws(() => view.unprotectWorkbook(), (e) => e.needsPassword === 'workbook', 'no password: asked for');
  assert.throws(() => view.unprotectWorkbook({ password: 'rutba' }), (e) => e.message === WRONG_PASSWORD);
  view.unprotectWorkbook({ password: 'Rutba' });
  assert.equal(view.workbookProtection().structure, false);
});

test('Protect Workbook writes workbookProtection as Excel does — before the sheets, SHA-512, a salt, 100,000 rounds', () => {
  const view = new SheetView(book());
  view.protectWorkbook({ password: 'open sesame' });
  const xml = OoxmlPackage.read(view.save()).text('xl/workbook.xml');
  const m = /<workbookProtection\b([^>]*)\/>/.exec(xml);
  assert.ok(m, 'the element');
  assert.ok(xml.indexOf('<workbookProtection') < xml.indexOf('<sheets>'), 'before the sheets, where the schema puts it');
  assert.match(m[1], /workbookAlgorithmName="SHA-512"/);
  assert.match(m[1], /workbookSpinCount="100000"/);
  assert.match(m[1], /lockStructure="1"/);
  const salt = /workbookSaltValue="([^"]+)"/.exec(m[1])[1];
  const hash = /workbookHashValue="([^"]+)"/.exec(m[1])[1];
  assert.equal(Buffer.from(salt, 'base64').length, 16, 'a 16-byte salt');
  assert.equal(hash, hashPassword('open sesame', { salt, spinCount: 100000 }), 'the hash of the password with that salt');

  const reopened = new SheetView(view.save());
  assert.deepEqual(reopened.workbookProtection(), { structure: true, windows: false, hasPassword: true });
  assert.throws(() => reopened.unprotectWorkbook({ password: 'open' }), (e) => e.message === WRONG_PASSWORD);
  reopened.unprotectWorkbook({ password: 'open sesame' });
  assert.equal(reopened.workbookProtection().structure, false);
});

test('while the structure is locked no sheet is added, deleted, renamed, moved, hidden or shown again', () => {
  const view = new SheetView(book());
  view.hideSheet('Notes');
  view.protectWorkbook();
  const refused = (fn) => assert.throws(fn, (e) => e.message === WORKBOOK_LOCKED && e.protection === true);
  refused(() => view.addSheet());
  refused(() => view.removeSheet('Summary'));
  refused(() => view.renameSheet('Data', 'Figures'));
  refused(() => view.moveSheet('Data', 1));
  refused(() => view.hideSheet('Summary'));
  refused(() => view.unhideSheet('Notes'));
  assert.deepEqual(view.sheetNames(), ['Data', 'Summary', 'Notes'], 'nothing changed');
  assert.deepEqual(view.hiddenSheets(), ['Notes']);
  // Cells are still the sheet's business.
  view.setCell(1, 1, 9);
  assert.equal(view.displayValue(1, 1).text, '9');
  view.unprotectWorkbook();
  view.renameSheet('Data', 'Figures');
  assert.deepEqual(view.sheetNames(), ['Figures', 'Summary', 'Notes']);
});

test('Hide, Unhide and Move: the sheet entry state, the order, names scoped by position follow, one visible sheet kept', () => {
  const view = new SheetView(book());
  view.selectSheet('Summary');
  view.workbook.setDefinedName('_xlnm.Print_Area', 'Summary!$A$1:$A$1', { localSheetId: 1 });
  view.moveSheet('Summary', 0);
  assert.deepEqual(view.sheetNames(), ['Summary', 'Data', 'Notes']);
  assert.equal(view.activeSheet, 'Summary');
  let xml = OoxmlPackage.read(view.save()).text('xl/workbook.xml');
  assert.match(xml, /<definedName name="_xlnm.Print_Area" localSheetId="0">/, 'the print area went with its sheet');
  assert.equal(new SheetView(view.save()).sheetNames()[0], 'Summary', 'kept in the file');
  view.undo();
  assert.deepEqual(view.sheetNames(), ['Data', 'Summary', 'Notes'], 'one undo step');

  view.selectSheet('Notes');
  view.hideSheet('Notes');
  assert.notEqual(view.activeSheet, 'Notes', 'the view moved off the hidden sheet');
  xml = OoxmlPackage.read(view.save()).text('xl/workbook.xml');
  assert.match(xml, /<sheet name="Notes"[^>]*state="hidden"/);
  view.hideSheet('Summary');
  assert.throws(() => view.hideSheet('Data'), /at least one visible worksheet/);
  assert.deepEqual(new SheetView(view.save()).hiddenSheets(), ['Summary', 'Notes']);
  view.unhideSheet('Notes');
  assert.equal(view.activeSheet, 'Notes', 'Unhide goes to the sheet');
  assert.doesNotMatch(OoxmlPackage.read(view.save()).text('xl/workbook.xml'), /<sheet name="Notes"[^>]*state=/);
});

test('Allow Edit Ranges: a locked cell in a range stays editable on a protected sheet, others are refused as before', () => {
  const view = new SheetView(book());
  view.setEditRange({ title: 'Quantities', ref: '$B$2:$B$3' });
  assert.deepEqual(view.editRanges(), [{ title: 'Quantities', ref: 'B2:B3', hasPassword: false }]);
  const tail = OoxmlPackage.read(view.save()).text('xl/worksheets/sheet1.xml');
  assert.match(tail, /<protectedRanges><protectedRange sqref="B2:B3" name="Quantities"\/><\/protectedRanges>/);
  view.protect();
  const sheetXml = OoxmlPackage.read(view.save()).text('xl/worksheets/sheet1.xml');
  assert.ok(sheetXml.indexOf('<sheetProtection') < sheetXml.indexOf('<protectedRanges'), 'protection before the ranges, in schema order');
  view.setCell(1, 1, 10);
  assert.equal(view.displayValue(1, 1).text, '10', 'inside the range: taken');
  assert.throws(() => view.setCell(1, 0, 'Pencils'), /is locked/, 'outside it: refused as today');
  assert.throws(() => view.setEditRange({ title: 'More', ref: 'A1' }), /unprotect it/, 'the ranges are fixed while the sheet is protected');
  assert.throws(() => view.deleteEditRange('Quantities'), /unprotect it/);
  view.unprotect();
  view.deleteEditRange('Quantities');
  assert.deepEqual(view.editRanges(), []);
});

test('a range with a password is unlocked once for the session; a sheet password is asked for and checked', () => {
  const view = new SheetView(book());
  view.setEditRange({ title: 'Budget', ref: 'B2:B3', password: 'ledger' });
  const block = /<protectedRange\b[^>]*\/>/.exec(OoxmlPackage.read(view.save()).text('xl/worksheets/sheet1.xml'))[0];
  assert.match(block, /algorithmName="SHA-512"/);
  assert.match(block, /spinCount="100000"/);
  const salt = /saltValue="([^"]+)"/.exec(block)[1];
  assert.equal(/hashValue="([^"]+)"/.exec(block)[1], hashPassword('ledger', { salt, spinCount: 100000 }));

  view.protect({ password: 'boss' });
  assert.equal(view.protection().hasPassword, true);
  view.select(1, 1);
  assert.equal(view.rangeLockAt(), 'Budget', 'the window asks before an edit starts');
  assert.throws(() => view.setCell(1, 1, 5), (e) => e.range === 'Budget' && /protected by a password/.test(e.message));
  assert.throws(() => view.unlockRange('Budget', 'Ledger'), (e) => e.message === WRONG_PASSWORD);
  view.unlockRange('Budget', 'ledger');
  assert.equal(view.rangeLockAt(), null);
  view.setCell(1, 1, 5);
  assert.equal(view.displayValue(1, 1).text, '5');
  view.setCell(2, 1, 6);
  assert.equal(view.displayValue(2, 1).text, '6', 'once for the whole range');

  assert.throws(() => view.unprotect(), (e) => e.needsPassword === 'sheet');
  assert.throws(() => view.unprotect({ password: 'Boss' }), (e) => e.message === WRONG_PASSWORD);
  view.unprotect({ password: 'boss' });
  assert.equal(view.protection().sheet, false);
  // Modify keeps the range's password when none is typed.
  view.setEditRange({ was: 'Budget', title: 'Budget 2', ref: 'B2:B4' });
  assert.deepEqual(view.editRanges(), [{ title: 'Budget 2', ref: 'B2:B4', hasPassword: true }]);
});
