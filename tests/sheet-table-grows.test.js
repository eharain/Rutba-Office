/**
 * A table grows to take a row typed straight under it, as Excel's
 * AutoExpansion does. Before 2.18 the row stayed outside the table, unbanded
 * and out of its filter.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

const tabled = () => {
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows: [['Item', 'Qty'], ['Ink', 4], ['Paper', 6]] }] }));
  view.select(0, 0);
  view.select(2, 1, { extend: true });
  view.formatAsTable();
  return view;
};
const tableRef = (view) => view.workbook.tables()[0].ref;

test('a value typed under a table makes the table a row longer, its filter too, and undo takes it back', () => {
  const view = tabled();
  assert.equal(tableRef(view), 'A1:B3');
  view.select(3, 0);
  view.setCell(3, 0, 'Pens');
  assert.equal(tableRef(view), 'A1:B4');
  assert.equal(view.calc.getValue('Data', 3, 0), 'Pens');
  const part = Workbook.open(view.save()).pkg.text(view.workbook.tables()[0].part);
  assert.match(part, /<table\b[^>]*\sref="A1:B4"/);
  assert.doesNotMatch(part, /<autoFilter\b[^>]*\sref="A1:B3"/, 'the filter covers the new row');

  view.setCell(3, 1, '2');
  assert.equal(tableRef(view), 'A1:B4', 'the next cell of the new row is in the table already');
  view.undo();
  view.undo();
  assert.equal(tableRef(view), 'A1:B3', 'one undo step takes the row and the growth back');
  assert.equal(view.calc.getValue('Data', 3, 0), '');
});

test('a table does not take a row typed beside it, two rows under it, or a cleared cell', () => {
  const view = tabled();
  view.setCell(3, 2, 'beside');
  view.setCell(4, 0, 'two under');
  view.setCell(3, 0, '');
  assert.equal(tableRef(view), 'A1:B3');
});
