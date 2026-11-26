// A table is known to formulas from the moment it is made.
//
// Format as Table wrote the table part and told the formula engine nothing:
// =[@Qty]*[@Price] in the new table gave #VALUE! and =SUM(Table1[Total])
// gave #NAME? until the workbook was saved and opened again. A table that
// grew by a typed row was not told of either. The engine's tables are now
// made the workbook's own again after every edit, undo and redo.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const BOOK = () => buildXlsx({ sheets: [{ name: 'Orders', rows: [
  ['Item', 'Qty', 'Price', 'Total'],
  ['Pens', 3, 2, ''],
  ['Ink', 5, 4, ''],
] }] });

const put = (view, row, col, value) => { view.select(row, col); view.setCell(row, col, value); };
const shown = (view, row, col) => view.displayValue(row, col).text;

test('formulas see a table made with Format as Table straight away', () => {
  const view = new SheetView(BOOK());
  view.select(0, 0);
  view.formatAsTable();
  const name = view.workbook.tables()[0].name;

  put(view, 1, 3, '=[@Qty]*[@Price]');
  put(view, 2, 3, '=[@Qty]*[@Price]');
  assert.deepEqual([shown(view, 1, 3), shown(view, 2, 3)], ['6', '20'], 'this row, by column name');
  put(view, 0, 6, '=SUM(' + name + '[Total])');
  assert.equal(shown(view, 0, 6), '26', 'a whole column, by the table\'s name');
  put(view, 1, 6, '=COUNTA(' + name + '[Item])');
  assert.equal(shown(view, 1, 6), '2');
});

test('a formula written before the table existed gets its answer when the table is made', () => {
  const view = new SheetView(BOOK());
  put(view, 0, 6, '=SUM(Table1[Qty])');
  assert.match(shown(view, 0, 6), /^#/, 'no table yet');
  view.select(0, 0);
  view.formatAsTable();
  assert.equal(view.workbook.tables()[0].name, 'Table1');
  assert.equal(shown(view, 0, 6), '8');
});

test('a row typed under the table is in the table for formulas too', () => {
  const view = new SheetView(BOOK());
  view.select(0, 0);
  view.formatAsTable();
  put(view, 0, 6, '=SUM(Table1[Qty])');
  put(view, 1, 6, '=COUNTA(Table1[Item])');
  put(view, 3, 0, 'Clips');
  assert.equal(view.workbook.tables()[0].ref, 'A1:D4', 'the table took the row');
  put(view, 3, 1, 10);
  assert.deepEqual([shown(view, 0, 6), shown(view, 1, 6)], ['18', '3']);
  put(view, 3, 2, 1);
  put(view, 3, 3, '=[@Qty]*[@Price]');
  assert.equal(shown(view, 3, 3), '10', 'this row works in the new row');
});

test('undo takes the table away from formulas, redo gives it back, and the saved file agrees', () => {
  const view = new SheetView(BOOK());
  put(view, 0, 6, '=SUM(Table1[Qty])');
  view.select(0, 0);
  view.formatAsTable();
  assert.equal(shown(view, 0, 6), '8');
  view.undo();
  assert.equal(view.workbook.tables().length, 0);
  assert.match(shown(view, 0, 6), /^#/, 'the table is gone, and formulas know it');
  view.redo();
  assert.equal(shown(view, 0, 6), '8');

  const again = SheetView.open(view.save());
  assert.equal(shown(again, 0, 6), '8', 'and the saved file says the same');
});
