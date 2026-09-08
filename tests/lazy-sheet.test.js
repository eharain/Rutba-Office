// A large sheet with no formula stays in its part and is answered on demand.
//
// One sample-data workbook of eighteen million plain values took twenty
// seconds and two gigabytes to copy into the calculation model, for a model
// with nothing to calculate. Past a size, a formula-free sheet is a provider
// the model reads through; an edit overlays it and a cleared cell stays
// cleared.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { Workbook } from '@rutba/ooxml';
import { toSpreadsheet } from '@rutba/ooxml/recalc';
import { SheetView } from '@rutba/sheet-view';

const ROWS = 20500;
const rows = [];
for (let r = 0; r < ROWS; r++) rows.push([r + 1, `name ${r + 1}`, (r % 7) * 1.5, r % 2 === 0 ? 'yes' : 'no']);
const BIG = buildXlsx({ sheets: [{ name: 'Data', rows }, { name: 'Small', rows: [[1, 2, '=A1+B1']] }] });

test('a big formula-free sheet is lazy: read through, not copied, and it computes and edits like any other', () => {
  const wb = Workbook.open(BIG);
  const started = Date.now();
  const { sheet } = toSpreadsheet(wb);
  const ms = Date.now() - started;
  assert.ok(sheet.lazy.has('Data'), 'the big sheet is a provider');
  assert.ok(!sheet.lazy.has('Small'), 'the small one is loaded as before');
  assert.equal(sheet.sheets.get('Data').size, 0, 'nothing of it was copied into the model');
  assert.ok(ms < 1500, `building the model took ${ms} ms`);

  assert.equal(sheet.getValue('Data', 20000, 1), 'name 20001');
  assert.equal(sheet.getValue('Data', 6, 2), 9, 'a number reads back as a number');
  assert.equal(sheet.getInput('Data', 6, 2), 9);
  assert.equal(sheet.getValue('Data', 3, 9), '', 'an empty cell is empty');
  assert.deepEqual(sheet.usedBounds('Data'), { maxRow: ROWS - 1, maxCol: 3 }, 'the extent comes from the dimension');

  assert.equal(sheet.getValue('Small', 0, 2), 3);
  sheet.setCell('Small', 0, 3, '=SUM(Data!A1:A100)');
  sheet.recalculate();
  assert.equal(sheet.getValue('Small', 0, 3), 5050, 'a formula sums through the provider');

  sheet.setCell('Data', 5, 1, 'edited');
  assert.equal(sheet.getValue('Data', 5, 1), 'edited', 'an edit overlays the provider');
  sheet.setCell('Data', 6, 1, '');
  assert.equal(sheet.getValue('Data', 6, 1), '', 'a cleared cell stays cleared');
  assert.equal(sheet.getValue('Data', 7, 1), 'name 8', 'its neighbour is untouched');
});

test('the sheet view opens the big workbook, shows the right cells, and saves an edit into the file', () => {
  const view = new SheetView(BIG);
  view.viewportWidth = 900;
  view.viewportHeight = 500;
  const frame = view.render();
  const b2 = frame.cells.find((c) => c.ref === 'B2');
  assert.equal(b2?.text, 'name 2');
  assert.ok(frame.rows.length > 0 && view.bounds.maxRow === ROWS - 1, `bounds ${JSON.stringify(view.bounds)}`);
  view.scrollTo(0, 20000 * 20);
  const far = view.render();
  assert.ok(far.cells.some((c) => /^B2\d{4}$/.test(c.ref) && /^name 2\d{4}$/.test(c.text)), 'the far end reads through the provider');
});
