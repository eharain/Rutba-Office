/**
 * A PivotTable sorted, and its values shown as a share or a difference.
 * Before 2.28 a pivot's items were always down the alphabet and a value was
 * only ever the plain total: a file Excel had sorted by value, or shown as a
 * percentage of the grand total, came back in alphabetical order with raw sums.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';

// Region totals: East 16, North 7, South 8. Product totals: Ink 12, Paper 19.
const ROWS = [
  ['Region', 'Product', 'Qty'],
  ['North', 'Ink', 4],
  ['South', 'Ink', 6],
  ['North', 'Paper', 3],
  ['South', 'Paper', 2],
  ['East', 'Ink', 2],
  ['East', 'Paper', 14],
];
const open = (extra = {}) => {
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows: ROWS }] }));
  view.createPivot({ name: 'Sales', source: 'Data!A1:C7', target: { sheet: 'Data', row: 0, col: 6 }, rowFields: ['Region'], dataFields: [{ field: 'Qty', subtotal: 'sum' }], ...extra });
  return view;
};
const LEFT = 6;
const at = (view, r, c) => view.calc.getValue('Data', r, c);
/** The pivot's first column, top to bottom, down to its Grand Total. */
const labels = (view) => {
  const out = [];
  for (let r = 0; r < 20; r++) {
    const v = at(view, r, LEFT);
    if (v === 'Grand Total') break;
    if (v !== '' && v !== null && v !== undefined && v !== 'Row Labels' && v !== 'Region') out.push(v);
  }
  return out;
};
const rowOf = (view, label) => {
  for (let r = 0; r < 20; r++) if (at(view, r, LEFT) === label) return r;
  throw new Error('no row ' + label);
};
const colOf = (view, label) => {
  for (let r = 0; r < 4; r++) for (let c = LEFT; c < LEFT + 8; c++) if (at(view, r, c) === label) return c;
  throw new Error('no column ' + label);
};
const frame = (view) => view.render().pivots.find((p) => p.name === 'Sales');
const text = (view, r, c) => view.render().cells.find((cell) => cell.row === r && cell.col === c)?.text ?? null;
const definition = (view) => {
  const part = view.pivots().find((p) => p.name === 'Sales').part;
  return Workbook.open(view.save()).pkg.text(part);
};
const round = (n) => Math.round(n * 10000) / 10000;

// ---- sorted --------------------------------------------------------------------

test('a row field sorted Z to A, written as Excel writes it, kept when the file is opened again', () => {
  const view = open();
  assert.deepEqual(labels(view), ['East', 'North', 'South']);
  assert.deepEqual(frame(view).sorts, {});
  view.sortPivotField('Sales', 'Region', { descending: true });
  assert.deepEqual(labels(view), ['South', 'North', 'East']);
  assert.equal(at(view, rowOf(view, 'East'), LEFT + 1), 16, 'each total moved with its label');
  assert.deepEqual(frame(view).sorts, { Region: { by: 'label', descending: true } });
  const xml = definition(view);
  assert.match(xml, /<pivotField\b[^>]*\baxis="axisRow"[^>]*\bsortType="descending"/);
  assert.doesNotMatch(xml, /autoSortScope/);
  const again = SheetView.open(view.save());
  again.refreshPivot('Sales');
  assert.deepEqual(labels(again), ['South', 'North', 'East']);
});

test('sorted by a value, largest first or smallest first, and back to A to Z', () => {
  const view = open();
  view.sortPivotField('Sales', 'Region', { by: 'Qty', descending: true });
  assert.deepEqual(labels(view), ['East', 'South', 'North']);
  assert.deepEqual(frame(view).sorts, { Region: { by: 0, descending: true } });
  const xml = definition(view);
  assert.match(xml, /sortType="descending"/);
  assert.match(xml, /<\/items><autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="0"\/><\/reference><\/references><\/pivotArea><\/autoSortScope>/);
  const again = SheetView.open(view.save());
  again.refreshPivot('Sales');
  assert.deepEqual(labels(again), ['East', 'South', 'North'], 'the file says it, so a refresh keeps it');

  view.sortPivotField('Sales', 'Region', { by: 0, descending: false });
  assert.deepEqual(labels(view), ['North', 'South', 'East']);
  view.sortPivotField('Sales', 'Region');
  assert.deepEqual(labels(view), ['East', 'North', 'South']);
  const plain = definition(view);
  assert.match(plain, /sortType="ascending"/);
  assert.doesNotMatch(plain, /autoSortScope/);
});

test('one undo puts the order back, and redo sorts again', () => {
  const view = open();
  view.sortPivotField('Sales', 'Region', { by: 'Qty', descending: true });
  view.undo();
  assert.deepEqual(labels(view), ['East', 'North', 'South']);
  assert.deepEqual(frame(view).sorts, {});
  assert.doesNotMatch(definition(view), /sortType/);
  view.redo();
  assert.deepEqual(labels(view), ['East', 'South', 'North']);
});

test('a new source row is sorted into its place on refresh', () => {
  const view = open();
  view.sortPivotField('Sales', 'Region', { by: 'Qty', descending: true });
  view.setCell(1, 2, 40);
  view.refreshPivot('Sales');
  assert.deepEqual(labels(view), ['North', 'East', 'South'], 'North is 43 now');
});

test('a sort outlives a change of layout, a column field sorts too, and each level sorts inside the one above', () => {
  const view = open();
  view.sortPivotField('Sales', 'Region', { by: 'Qty', descending: true });
  view.setPivotLayout('Sales', { rows: ['Region'], cols: ['Product'], values: [{ field: 'Qty', subtotal: 'sum' }] });
  assert.deepEqual(labels(view).filter((v) => ['East', 'North', 'South'].includes(v)), ['East', 'South', 'North']);
  assert.ok(colOf(view, 'Ink') < colOf(view, 'Paper'));
  view.sortPivotField('Sales', 'Product', { descending: true });
  assert.ok(colOf(view, 'Paper') < colOf(view, 'Ink'), 'the columns Z to A');
  assert.equal(at(view, rowOf(view, 'East'), colOf(view, 'Paper')), 14);

  const two = open();
  two.setPivotLayout('Sales', { rows: ['Region', 'Product'], values: [{ field: 'Qty', subtotal: 'sum' }] });
  two.sortPivotField('Sales', 'Product', { by: 'Qty', descending: true });
  // Two row fields sit side by side, one row for each pair.
  const pairs = [];
  for (let r = 1; at(two, r, LEFT) !== 'Grand Total'; r++) pairs.push(at(two, r, LEFT) + ' ' + at(two, r, LEFT + 1));
  assert.deepEqual(pairs, ['East Paper', 'East Ink', 'North Ink', 'North Paper', 'South Ink', 'South Paper'],
    'each region lists its larger product first, and the regions are still down the alphabet');
});

test('a field that is not on the rows or columns is not sorted, nor by a value the pivot has not', () => {
  const view = open();
  assert.throws(() => view.sortPivotField('Sales', 'Product'), /Rows or Columns/);
  assert.throws(() => view.sortPivotField('Sales', 'Region', { by: 'Product' }), /no value/);
  assert.throws(() => view.sortPivotField('Sales', 'Region', { by: 3 }), /no value/);
  assert.throws(() => view.sortPivotField('Nothing', 'Region'), /no pivot table/);
});

// ---- Show values as ---------------------------------------------------------------

test('% of grand total: fractions shown as percentages, written as Excel writes it', () => {
  const view = open();
  view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty', subtotal: 'sum', showAs: 'percentOfTotal' }] });
  const c = LEFT + 1;
  assert.equal(round(at(view, rowOf(view, 'East'), c)), round(16 / 31));
  assert.equal(round(at(view, rowOf(view, 'North'), c)), round(7 / 31));
  assert.equal(at(view, rowOf(view, 'Grand Total'), c), 1);
  assert.equal(text(view, rowOf(view, 'East'), c), '51.61%');
  assert.equal(text(view, rowOf(view, 'Grand Total'), c), '100.00%');
  assert.equal(frame(view).valueFields[0].showAs, 'percentOfTotal');
  assert.match(definition(view), /<dataField\b[^>]*\bshowDataAs="percentOfTotal"[^>]*\bnumFmtId="10"/);

  const again = SheetView.open(view.save());
  again.refreshPivot('Sales');
  assert.equal(round(at(again, rowOf(again, 'South'), c)), round(8 / 31), 'read back and worked out the same');
  assert.equal(text(again, rowOf(again, 'South'), c), '25.81%');
});

test('back to no calculation the totals return and the percent format goes; undo does the same', () => {
  const view = open();
  const c = LEFT + 1;
  view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty', subtotal: 'sum', showAs: 'percentOfTotal' }] });
  view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty', subtotal: 'sum', showAs: 'normal' }] });
  assert.equal(at(view, rowOf(view, 'East'), c), 16);
  assert.equal(text(view, rowOf(view, 'East'), c), '16');
  assert.doesNotMatch(definition(view), /showDataAs/);
  view.undo();
  assert.equal(text(view, rowOf(view, 'East'), c), '51.61%');
  view.undo();
  assert.equal(at(view, rowOf(view, 'East'), c), 16);
  assert.equal(text(view, rowOf(view, 'East'), c), '16');
});

test('% of row total and % of column total', () => {
  const view = open();
  view.setPivotLayout('Sales', { rows: ['Region'], cols: ['Product'], values: [{ field: 'Qty', subtotal: 'sum', showAs: 'percentOfRow' }] });
  const east = rowOf(view, 'East');
  assert.equal(at(view, east, colOf(view, 'Ink')), 2 / 16);
  assert.equal(at(view, east, colOf(view, 'Paper')), 14 / 16);
  assert.equal(at(view, east, colOf(view, 'Grand Total')), 1);
  assert.equal(text(view, east, colOf(view, 'Paper')), '87.50%');
  assert.equal(round(at(view, rowOf(view, 'Grand Total'), colOf(view, 'Ink'))), round(12 / 31));

  view.setPivotLayout('Sales', { rows: ['Region'], cols: ['Product'], values: [{ field: 'Qty', subtotal: 'sum', showAs: 'percentOfCol' }] });
  assert.equal(round(at(view, rowOf(view, 'East'), colOf(view, 'Ink'))), round(2 / 12));
  assert.equal(round(at(view, rowOf(view, 'South'), colOf(view, 'Paper'))), round(2 / 19));
  assert.equal(at(view, rowOf(view, 'Grand Total'), colOf(view, 'Paper')), 1);
  assert.equal(round(at(view, rowOf(view, 'East'), colOf(view, 'Grand Total'))), round(16 / 31));
});

test('the difference from the item before: the first row and the total stay empty', () => {
  const view = open();
  view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty', subtotal: 'sum', showAs: 'difference' }] });
  const c = LEFT + 1;
  assert.equal(view.isFilled(rowOf(view, 'East'), c), false);
  assert.equal(at(view, rowOf(view, 'North'), c), -9);
  assert.equal(at(view, rowOf(view, 'South'), c), 1);
  assert.equal(view.isFilled(rowOf(view, 'Grand Total'), c), false);
  assert.match(definition(view), /<dataField\b[^>]*\bshowDataAs="difference" baseField="0" baseItem="1048828"/);
  const again = SheetView.open(view.save());
  again.refreshPivot('Sales');
  assert.equal(at(again, rowOf(again, 'North'), c), -9);
  // With the order changed, "before" is the row above as it is shown.
  view.sortPivotField('Sales', 'Region', { descending: true });
  assert.equal(view.isFilled(rowOf(view, 'South'), c), false);
  assert.equal(at(view, rowOf(view, 'North'), c), -1);
  assert.equal(at(view, rowOf(view, 'East'), c), 9);
});

test('a layout that does not say keeps what the value was shown as; a wrong word is refused', () => {
  const view = open();
  view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty', subtotal: 'sum', showAs: 'percentOfTotal' }] });
  view.setPivotLayout('Sales', { rows: ['Region'], cols: ['Product'], values: [{ field: 'Qty', subtotal: 'sum' }] });
  assert.equal(frame(view).valueFields[0].showAs, 'percentOfTotal');
  assert.equal(at(view, rowOf(view, 'Grand Total'), colOf(view, 'Grand Total')), 1);
  assert.equal(text(view, rowOf(view, 'East'), colOf(view, 'Paper')), '45.16%');
  assert.throws(() => view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty', subtotal: 'sum', showAs: 'rank' }] }), /not a way to show/);
});

test('a number format put on the pivot\'s plain totals outlives a refresh', () => {
  const view = open();
  const r = rowOf(view, 'East');
  view.select(r, LEFT + 1);
  view.setFormat({ numberFormat: '0.0' });
  view.refreshPivot('Sales');
  assert.equal(text(view, r, LEFT + 1), '16.0');
  view.sortPivotField('Sales', 'Region', { descending: true });
  assert.equal(view.formatFor(r, LEFT + 1), '0.0', 'the format belongs to the cell, as in Excel');
});
