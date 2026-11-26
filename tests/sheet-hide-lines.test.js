/**
 * Rows and columns hidden and shown again. Before 2.28 the heading menus
 * held only Insert and Delete and Home's Format menu had no Hide or Unhide:
 * a row could be hidden by a filter and by nothing else.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

const book = () => buildXlsx({
  sheets: [{ name: 'Data', rows: [['Item', 'Cost', 'Price', 'Sold'], ['Ink', 2, 5, 40], ['Paper', 3, 6, 12], ['Pens', 1, 4, 30], ['Tape', 2, 3, 9]] }],
});
const open = () => SheetView.open(book());
const sheetXml = (view) => Workbook.open(view.save()).pkg.text('xl/worksheets/sheet1.xml');
const rowsShown = (view) => [0, 1, 2, 3, 4].filter((r) => view.geo.rowHeight(r) > 0);
const colsShown = (view) => [0, 1, 2, 3].filter((c) => view.geo.colWidth(c) > 0);

test('rows are hidden, written as Excel writes them, and shown again by the row beside them', () => {
  const view = open();
  view.select(1, 0);
  view.select(2, 0, { extend: true });
  view.hideLines('row');
  assert.deepEqual(rowsShown(view), [0, 3, 4]);
  const xml = sheetXml(view);
  assert.match(xml, /<row[^>]*r="2"[^>]*hidden="1"/);
  assert.match(xml, /<row[^>]*r="3"[^>]*hidden="1"/);
  assert.equal(view.calc.getValue('Data', 1, 0), 'Ink', 'a hidden row keeps what it holds');
  assert.deepEqual(rowsShown(SheetView.open(view.save())), [0, 3, 4], 'and is hidden when the file is opened again');

  // The row under the hidden ones is enough to bring them back.
  view.select(3, 0);
  view.unhideLines('row');
  assert.deepEqual(rowsShown(view), [0, 1, 2, 3, 4]);
  assert.doesNotMatch(sheetXml(view), /hidden=/);
});

test('columns are hidden and shown again, and come back at their own width', () => {
  const view = open();
  view.setColWidth(1, 140);
  const wide = view.geo.colWidth(1);
  const usual = view.geo.colWidth(2);
  view.select(0, 1);
  view.select(0, 2, { extend: true });
  view.hideLines('col');
  assert.deepEqual(colsShown(view), [0, 3]);
  const xml = sheetXml(view);
  assert.match(xml, /<col min="2" max="2"[^>]*hidden="1"/, 'the column with a width of its own keeps it and gains hidden');
  assert.match(xml, /<col min="3" max="3" width="0" hidden="1" customWidth="1"\/>/, 'a plain column is hidden as Excel writes one');
  assert.deepEqual(colsShown(SheetView.open(view.save())), [0, 3]);

  // Selecting across the hidden columns, as Excel asks.
  view.select(0, 0);
  view.select(0, 3, { extend: true });
  view.unhideLines('col');
  assert.deepEqual(colsShown(view), [0, 1, 2, 3]);
  assert.equal(view.geo.colWidth(1), wide, 'its own width is back');
  assert.equal(view.geo.colWidth(2), usual, 'the plain column is back at the usual width');
  assert.doesNotMatch(sheetXml(view), /hidden=/);
});

test('one undo shows what was hidden, and redo hides it again', () => {
  const view = open();
  view.select(2, 1);
  view.hideLines('row');
  view.hideLines('col');
  assert.deepEqual(rowsShown(view), [0, 1, 3, 4]);
  assert.deepEqual(colsShown(view), [0, 2, 3]);
  view.undo();
  assert.deepEqual(colsShown(view), [0, 1, 2, 3]);
  view.undo();
  assert.deepEqual(rowsShown(view), [0, 1, 2, 3, 4]);
  view.redo();
  assert.deepEqual(rowsShown(view), [0, 1, 3, 4]);
});

test('the first row and column, which cannot be selected across, come back with the one beside them or with all', () => {
  const view = open();
  view.select(0, 0);
  view.hideLines('row');
  view.hideLines('col');
  assert.deepEqual(rowsShown(view), [1, 2, 3, 4]);
  view.select(1, 1);
  view.unhideLines('row');
  view.unhideLines('col');
  assert.deepEqual(rowsShown(view), [0, 1, 2, 3, 4]);
  assert.deepEqual(colsShown(view), [0, 1, 2, 3]);

  view.select(1, 0);
  view.hideLines('row');
  view.select(4, 0);
  view.hideLines('row');
  view.select(2, 0);
  view.unhideLines('row', { all: true });
  assert.deepEqual(rowsShown(view), [0, 1, 2, 3, 4], 'every hidden row of the sheet');
});

test('nothing hidden near the selection: unhide does nothing and makes no undo step', () => {
  const view = open();
  view.select(2, 2);
  view.unhideLines('row');
  view.unhideLines('col');
  assert.equal(view.canUndo, false);
});

test('a protected sheet keeps its rows and columns as they are', () => {
  const view = open();
  view.select(1, 1);
  view.protect();
  assert.throws(() => view.hideLines('row'), /protected/);
  assert.throws(() => view.unhideLines('col'), /protected/);
  assert.deepEqual(rowsShown(view), [0, 1, 2, 3, 4]);
});

test('a formula still counts a hidden row, as Excel\'s SUM does', () => {
  const view = open();
  view.setCell(5, 3, '=SUM(D2:D5)');
  view.select(1, 0);
  view.hideLines('row');
  assert.equal(view.calc.getValue('Data', 5, 3), 91);
});

test('every hidden row and column comes back in one step, and one undo hides them again', () => {
  const view = open();
  view.select(1, 1);
  view.hideLines('row');
  view.hideLines('col');
  view.select(3, 3);
  view.unhideAllLines();
  assert.deepEqual(rowsShown(view), [0, 1, 2, 3, 4]);
  assert.deepEqual(colsShown(view), [0, 1, 2, 3]);
  view.undo();
  assert.deepEqual(rowsShown(view), [0, 2, 3, 4]);
  assert.deepEqual(colsShown(view), [0, 2, 3]);
});
