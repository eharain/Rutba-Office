// A window's Save compresses the changed parts on Node's thread pool
// (OoxmlPackage saveAsync), so a large workbook's save no longer holds the
// main thread: the same bytes as a plain save, the event loop turning while
// it runs, and an edit made meanwhile kept out of the file and unsaved.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { buildXlsx, OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const bigBook = (n = 40000) => {
  const rows = [['Item', 'Region', 'Units']];
  for (let i = 1; i <= n; i++) rows.push([`Item ${i}`, ['East', 'West'][i % 2], i % 97]);
  return buildXlsx({ sheets: [{ name: 'Data', rows }] });
};

test('saveAsync writes the same bytes as save, with the event loop turning while it compresses', async () => {
  const bytes = bigBook();
  const a = new SheetView(bytes);
  const b = new SheetView(bytes);
  a.select(5, 1); a.setCell(5, 1, 'North');
  b.select(5, 1); b.setCell(5, 1, 'North');
  const plain = Buffer.from(a.save());
  let turns = 0;
  const tick = setInterval(() => { turns += 1; }, 0);
  const held = Buffer.from(await b.pkg.saveAsync(() => b.save()));
  clearInterval(tick);
  assert.ok(plain.equals(held), 'byte for byte');
  assert.ok(turns >= 1, `the loop turned ${turns} times while the parts were deflated`);
  assert.equal(OoxmlPackage.read(held).has('xl/worksheets/sheet1.xml'), true);
});

test('a window\'s Save goes through saveAsync: written, and an edit made while it runs is kept out and left unsaved', async () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-save-')), 'big.xlsx');
  fs.writeFileSync(file, bigBook(20000));
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = await docs.open({ path: file });
  docs.apply({ id, ops: [{ op: 'select', row: 1, col: 0 }, { op: 'setCell', row: 1, col: 0, value: 'Before' }] });
  const saving = docs.saveAsync({ id });
  // Meanwhile, before the compression is done.
  docs.apply({ id, ops: [{ op: 'select', row: 2, col: 0 }, { op: 'setCell', row: 2, col: 0, value: 'During' }] });
  const saved = await saving;
  assert.equal(saved.dirty, true, 'the edit made during the save is not saved');
  const again = new SheetView(fs.readFileSync(file));
  assert.equal(again.displayValue(1, 0).text, 'Before');
  assert.notEqual(again.displayValue(2, 0).text, 'During', 'the file holds what was flushed');
  const done = await docs.saveAsync({ id });
  assert.equal(done.dirty, false);
  assert.equal(new SheetView(fs.readFileSync(file)).displayValue(2, 0).text, 'During');
});
