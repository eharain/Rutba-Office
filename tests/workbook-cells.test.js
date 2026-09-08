/**
 * Walking a sheet's cells, and the row cache behind it.
 *
 * `Workbook.cells(sheet)` is what the calc model is built from. It used to
 * be `getCell` per cell, and `getCell` re-parsed the whole row per call — so
 * loading was quadratic in a sheet's width, and a tender workbook of 2.4
 * million mostly-blank formatted cells took eleven minutes to open. The
 * cache is keyed on the row's XML string, so an edit is seen at once.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

const book = () => Workbook.open(buildXlsx({
  sheets: [{ name: 'S', rows: [[1, 'two', null, '=A1*2', 3.5], [], ['last']] }],
}));

test('cells() yields every cell that holds something, decoded, in row order', () => {
  const wb = book();
  const cells = [...wb.cells('S')];
  assert.deepEqual(cells.map((c) => c.ref), ['A1', 'B1', 'D1', 'E1', 'A3']);
  assert.deepEqual(cells.map((c) => [c.row, c.col]), [[0, 0], [0, 1], [0, 3], [0, 4], [2, 0]]);
  assert.deepEqual(cells.map((c) => c.value), [1, 'two', '=A1*2', 3.5, 'last']);
  // The same answers `getCell` gives, one at a time.
  for (const c of cells) assert.equal(wb.getCell('S', c.ref), c.value);
  assert.equal(wb.getCell('S', 'C1'), null, 'the gap is a gap');
});

test('a row is parsed once per version of its XML, and an edit refreshes it', () => {
  const wb = book();
  const { part } = wb._sheetPart('S');
  const row = part.rows[0];
  const first = part.constructor.cellsByCol(row);
  assert.strictEqual(part.constructor.cellsByCol(row), first, 'the same map while the XML stands');
  wb.setCell('S', 'C1', 'filled');
  assert.notStrictEqual(part.constructor.cellsByCol(row), first, 'a new map after the row changed');
  assert.equal(wb.getCell('S', 'C1'), 'filled');
  assert.deepEqual([...wb.cells('S')].map((c) => c.ref), ['A1', 'B1', 'C1', 'D1', 'E1', 'A3']);
});

test('a wide sheet loads in linear time, not quadratic', () => {
  // 400 columns by 60 rows, every cell present: 24,000 cells. Quadratic
  // parsing made this take seconds; it should be well under one.
  const rows = Array.from({ length: 60 }, (_, r) => Array.from({ length: 400 }, (_, c) => (c % 7 === 0 ? r * 400 + c : `t${c}`)));
  const wb = Workbook.open(buildXlsx({ sheets: [{ name: 'Wide', rows }] }));
  const started = Date.now();
  const count = [...wb.cells('Wide')].length;
  const ms = Date.now() - started;
  assert.equal(count, 24000);
  assert.ok(ms < 1500, `24,000 cells walked in ${ms} ms`);
});
