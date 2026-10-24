/**
 * Sort reads the list's own heading row: the Sort dialog names the columns
 * from it, and no sort moves it into the data. Before 2.18 the dialog named
 * columns from the row above the active cell, and a selected list was
 * sorted heading and all.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { buildXlsx } from '@rutba/ooxml/build';

const list = (rows, styles) => SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows, ...(styles ? { styles } : {}) }] }));
const col = (view, c, from, to) => { const out = []; for (let r = from; r <= to; r++) out.push(view.displayValue(r, c).text); return out; };

test('the Sort dialog names the columns from the list\'s heading row, wherever the cursor is in it', () => {
  const view = list([['Name', 'Qty'], ['Pear', 5], ['Apple', 2], ['Fig', 9]]);
  view.select(3, 1);
  const plan = view.sortPlan();
  assert.deepEqual(plan.columns, [{ col: 0, letter: 'A', heading: 'Name' }, { col: 1, letter: 'B', heading: 'Qty' }]);
  assert.equal(plan.header, 0);
  assert.deepEqual(plan.range, { top: 1, bottom: 3, left: 0, right: 1 });
  assert.equal(plan.keyCol, 1, 'A to Z sorts by the active cell\'s column');
});

test('a list selected with its heading row is sorted below it', () => {
  const view = list([['Name', 'Qty'], ['Pear', 5], ['Apple', 2], ['Fig', 9]]);
  view.select(0, 0);
  view.select(3, 1, { extend: true });
  view.sortSelection({ keys: [{ col: 0, ascending: true }] });
  assert.deepEqual(col(view, 0, 0, 3), ['Name', 'Apple', 'Fig', 'Pear']);
  assert.deepEqual(col(view, 1, 0, 3), ['Qty', '2', '9', '5']);
});

test('words over words read as a heading row when it is bold over plain', () => {
  const view = list([['Name', 'Town'], ['Pear', 'York'], ['Apple', 'Bath']], { 'A1:B1': { bold: true } });
  view.select(1, 0);
  assert.equal(view.sortPlan().header, 0);
  view.sortSelection({ ascending: true });
  assert.deepEqual(col(view, 0, 0, 2), ['Name', 'Apple', 'Pear']);

  // Without the bold the top row is data, as Excel reads it.
  const plain = list([['Pear', 'York'], ['Apple', 'Bath']]);
  plain.select(0, 0);
  assert.equal(plain.sortPlan().header, null);
  assert.deepEqual(plain.sortPlan().columns.map((c) => c.heading), [null, null]);
});

test('rows of data selected alone are sorted as they are', () => {
  const view = list([['Name', 'Qty'], ['Pear', 5], ['Apple', 2], ['Fig', 9]]);
  view.select(1, 0);
  view.select(3, 1, { extend: true });
  assert.equal(view.sortPlan().header, null);
  view.sortSelection({ ascending: false });
  assert.deepEqual(col(view, 0, 0, 3), ['Name', 'Pear', 'Fig', 'Apple']);
});

test('a table\'s header row stays at the top, whatever it holds', () => {
  const view = list([['Item', 'Code'], ['b', 'x'], ['a', 'y']]);
  view.select(0, 0);
  view.select(2, 1, { extend: true });
  view.formatAsTable();
  assert.equal(view.workbook.tables().length, 1);
  view.select(2, 0);
  assert.equal(view.sortPlan().header, 0);
  view.sortSelection({ ascending: true });
  assert.deepEqual(col(view, 0, 0, 2), ['Item', 'a', 'b']);
});
