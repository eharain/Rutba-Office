/**
 * A table's total row, its style changed after it is made, and a calculated
 * column that carries its formula into a row typed under the table. Before
 * 2.28 nothing switched a total row on, Format as Table on a table said the
 * range already had one, and the new row's formula cell stayed empty.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

const open = () => {
  const view = SheetView.open(buildXlsx({
    sheets: [{ name: 'Data', rows: [['Item', 'Qty', 'Price'], ['Ink', 4, 2.5], ['Paper', 6, 3], ['Pens', 10, 1]] }],
  }));
  view.select(1, 1);
  view.formatAsTable();
  return view;
};
const tableXml = (view) => {
  const wb = Workbook.open(view.save());
  return wb.pkg.text(wb.tables()[0].part);
};
const look = (view, cell) => view.render().cells.find((c) => c.ref === cell)?.style ?? null;

// ---- the total row ------------------------------------------------------------

test('the total row: Total in the first column, the last column summed as Excel writes it', () => {
  const view = open();
  assert.equal(view.activeTable().totalRow, false);
  view.setTotalRow();
  assert.equal(view.calc.getValue('Data', 4, 0), 'Total');
  assert.equal(view.editValue(4, 2), '=SUBTOTAL(109,Table1[Price])');
  assert.equal(view.calc.getValue('Data', 4, 2), 6.5);
  assert.equal(view.editValue(4, 1), '', 'a column in the middle totals nothing until asked');
  const xml = tableXml(view);
  assert.match(xml, /<table\b[^>]*\bref="A1:C5"[^>]*\btotalsRowCount="1"/);
  assert.doesNotMatch(xml, /totalsRowShown/);
  assert.match(xml, /<autoFilter ref="A1:C4"\/>/, 'the filter stays over the data');
  assert.match(xml, /<tableColumn id="1" name="Item" totalsRowLabel="Total"\/>/);
  assert.match(xml, /<tableColumn id="3" name="Price" totalsRowFunction="sum"\/>/);
  assert.equal(view.activeTable().totalRow, true);
  const again = SheetView.open(view.save());
  assert.equal(again.calc.getValue('Data', 4, 2), 6.5, 'and totals the same when the file is opened again');
});

test('a last column of words is counted, not summed', () => {
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows: [['Qty', 'Item'], [4, 'Ink'], [6, 'Paper']] }] }));
  view.select(1, 0);
  view.formatAsTable();
  view.setTotalRow(true);
  assert.equal(view.editValue(3, 1), '=SUBTOTAL(103,Table1[Item])');
  assert.equal(view.calc.getValue('Data', 3, 1), 2);
});

test('a total cell takes another function, or none, and the table part says which', () => {
  const view = open();
  view.setTotalRow(true);
  view.select(4, 1);
  assert.deepEqual(view.activeTable().total, { column: 'Qty', fn: null });
  view.setTotalFunction('average');
  assert.equal(view.editValue(4, 1), '=SUBTOTAL(101,Table1[Qty])');
  assert.equal(Math.round(view.calc.getValue('Data', 4, 1) * 100) / 100, 6.67);
  assert.deepEqual(view.activeTable().total, { column: 'Qty', fn: 'average' });
  assert.match(tableXml(view), /name="Qty" totalsRowFunction="average"/);
  view.setTotalFunction('max');
  assert.equal(view.calc.getValue('Data', 4, 1), 10);
  view.setTotalFunction(null);
  assert.equal(view.editValue(4, 1), '');
  assert.doesNotMatch(tableXml(view), /name="Qty" totalsRowFunction/);
  assert.throws(() => view.setTotalFunction('median'), /not a total/);
  view.select(2, 1);
  assert.throws(() => view.setTotalFunction('sum'), /total row/);
});

test('the total counts what a filter shows', () => {
  const view = open();
  view.setTotalRow(true);
  view.applyFilterRule('Table1', 'Qty', { kind: 'number', op: 'greater', value: 5 });
  assert.equal(view.calc.getValue('Data', 4, 2), 4, 'Paper and Pens: 3 + 1');
});

test('switched off, the row is emptied and leaves the table; switched on again, the totals it had are back', () => {
  const view = open();
  view.setTotalRow(true);
  view.select(4, 1);
  view.setTotalFunction('max');
  view.setTotalRow(false);
  assert.equal(view.isFilled(4, 0), false);
  assert.equal(view.isFilled(4, 2), false);
  const off = tableXml(view);
  assert.match(off, /<table\b[^>]*\bref="A1:C4"/);
  assert.match(off, /totalsRowShown="0"/);
  assert.doesNotMatch(off, /totalsRowCount/);
  view.select(2, 1);
  view.setTotalRow(true);
  assert.equal(view.calc.getValue('Data', 4, 0), 'Total');
  assert.equal(view.calc.getValue('Data', 4, 1), 10, 'the maximum it had');
  assert.equal(view.calc.getValue('Data', 4, 2), 6.5);
});

test('one undo takes the total row away again, and redo brings it back', () => {
  const view = open();
  view.setTotalRow(true);
  view.undo();
  assert.equal(view.isFilled(4, 0), false);
  assert.equal(view.activeTable().totalRow, false);
  assert.match(tableXml(view), /ref="A1:C4"/);
  view.redo();
  assert.equal(view.calc.getValue('Data', 4, 2), 6.5);
  assert.equal(view.activeTable().totalRow, true);
});

test('a row in use under the table is not taken, a cell outside a table is told so, a protected sheet refuses', () => {
  const view = SheetView.open(buildXlsx({
    sheets: [{ name: 'Data', rows: [['Item', 'Qty'], ['Ink', 4], ['Paper', 6], ['a note under the list', '']] }],
  }));
  view.select(0, 0);
  view.select(2, 1, { extend: true });
  view.formatAsTable();
  view.select(1, 1);
  assert.throws(() => view.setTotalRow(true), /row under the table is in use/);
  view.select(8, 8);
  assert.throws(() => view.setTotalRow(true), /Select a cell in a table/);
  const locked = open();
  locked.protect();
  assert.throws(() => locked.setTotalRow(true), /protected/);
  assert.throws(() => locked.setTableStyle({ style: 'TableStyleDark1' }), /protected/);
});

// ---- the style, changed later ---------------------------------------------------

test('Format as Table on a table changes its style where it used to refuse', () => {
  const view = open();
  assert.equal(view.activeTable().style, 'TableStyleMedium2');
  view.formatAsTable({ style: 'TableStyleDark1', stripes: true });
  assert.equal(view.activeTable().style, 'TableStyleDark1');
  assert.equal(Workbook.open(view.save()).tables().length, 1, 'still one table');
  assert.match(tableXml(view), /<tableStyleInfo name="TableStyleDark1" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"\/>/);
  view.setTableStyle({ stripes: false });
  assert.equal(view.activeTable().stripes, false);
  assert.match(tableXml(view), /name="TableStyleDark1"[^>]*showRowStripes="0"/, 'the style kept, the bands off');
  view.undo();
  view.undo();
  assert.equal(view.activeTable().style, 'TableStyleMedium2');
});

test('the three families look different, so a style picked later shows', () => {
  const view = open();
  const medium = { head: look(view, 'A1'), body: look(view, 'A2') };
  view.setTableStyle({ style: 'TableStyleLight9' });
  const light = { head: look(view, 'A1'), body: look(view, 'A2') };
  view.setTableStyle({ style: 'TableStyleDark1' });
  const dark = { head: look(view, 'A1'), body: look(view, 'A2') };
  const fill = (s) => s?.fill?.colour ?? s?.fill ?? null;
  assert.notDeepEqual(fill(light.head), fill(medium.head));
  assert.notDeepEqual(fill(dark.head), fill(medium.head));
  assert.notDeepEqual(fill(dark.body), fill(medium.body), 'a dark table colours its rows');
});

// ---- the calculated column, carried into a new row ------------------------------

test('a row typed under a table takes the calculated column\'s formula', () => {
  const view = SheetView.open(buildXlsx({
    sheets: [{ name: 'Data', rows: [['Item', 'Qty', 'Price', 'Total'], ['Ink', 4, 2.5, '=B2*C2'], ['Paper', 6, 3, '=B3*C3']] }],
  }));
  view.select(1, 0);
  view.formatAsTable();
  view.setCell(3, 0, 'Pens');
  assert.equal(view.editValue(3, 3), '=B4*C4', 'the formula moved to the new row');
  assert.match(tableXml(view), /ref="A1:D4"/, 'and the row joined the table');
  view.setCell(3, 1, 10);
  view.setCell(3, 2, 1);
  assert.equal(view.calc.getValue('Data', 3, 3), 10);
  // One undo for the typing, the growth and the formula together.
  const fresh = SheetView.open(buildXlsx({
    sheets: [{ name: 'Data', rows: [['Item', 'Qty', 'Twice'], ['Ink', 4, '=B2*2'], ['Paper', 6, '=B3*2']] }],
  }));
  fresh.select(1, 0);
  fresh.formatAsTable();
  fresh.setCell(3, 1, 9);
  assert.equal(fresh.calc.getValue('Data', 3, 2), 18);
  fresh.undo();
  assert.equal(fresh.isFilled(3, 1), false);
  assert.equal(fresh.isFilled(3, 2), false);
  assert.match(tableXml(fresh), /ref="A1:C3"/);
});

test('a column whose rows hold different formulas, or plain numbers, is left empty in the new row', () => {
  const view = SheetView.open(buildXlsx({
    sheets: [{ name: 'Data', rows: [['Item', 'Qty', 'Mixed', 'Plain'], ['Ink', 4, '=B2*2', 7], ['Paper', 6, '=B3+100', 8]] }],
  }));
  view.select(1, 0);
  view.formatAsTable();
  view.setCell(3, 0, 'Pens');
  assert.equal(view.isFilled(3, 2), false);
  assert.equal(view.isFilled(3, 3), false);
});
