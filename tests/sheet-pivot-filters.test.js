/**
 * A PivotTable's Filters area, its months, and the pivot shown when it is
 * made. Before 2.18 a pivot made under a long list was out of sight, the
 * Fields pane had no Filters area, and a date field listed every date.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';
import { dateToSerial } from '@rutba/formula';

const d = (y, m, day) => dateToSerial(new Date(Date.UTC(y, m - 1, day)));

const sales = () => {
  const rows = [
    ['Date', 'Region', 'Product', 'Qty'],
    [d(2026, 1, 5), 'North', 'Ink', 4],
    [d(2026, 1, 20), 'South', 'Ink', 6],
    [d(2026, 2, 3), 'North', 'Paper', 3],
    [d(2026, 2, 14), 'South', 'Paper', 1],
    [d(2026, 3, 9), 'North', 'Ink', 2],
  ];
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows }] }));
  view.select(1, 0);
  view.select(5, 0, { extend: true });
  view.setFormat({ numberFormat: 'yyyy-mm-dd' });
  view.createPivot({ name: 'Sales', source: 'Data!A1:D6', target: { sheet: 'Data', row: 0, col: 6 }, rowFields: ['Region'], dataFields: [{ field: 'Qty', subtotal: 'sum' }] });
  return view;
};

const column = (view, col, from, to) => {
  const out = [];
  for (let r = from; r <= to; r++) out.push(view.calc.getValue('Data', r, col));
  return out;
};
const frame = (view) => view.render().pivots.find((p) => p.name === 'Sales');

test('a new pivot made under a long list is picked and scrolled to', () => {
  const rows = [['Name', 'Qty']];
  for (let i = 1; i <= 120; i++) rows.push(['Item ' + (i % 4), i]);
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows }] }), { viewportWidth: 1000, viewportHeight: 500 });
  view.select(0, 0);
  const made = view.createPivot({ rowFields: ['Name'], dataFields: [{ field: 'Qty' }] });
  assert.equal(made.top, 122, 'made under the list, a line between');
  assert.deepEqual(view.selection.active, { row: 122, col: 0 }, 'its first cell is picked');
  assert.ok(view.scrollY > 0, 'and the sheet scrolled down to it');
  assert.ok(view.render().cells.some((c) => c.row === 122 && c.col === 0), 'the pivot is in the window');
});

test('the Filters area shows one item, written as Excel writes it, and undo takes it away', () => {
  const view = sales();
  view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty', subtotal: 'sum' }], filters: [{ field: 'Product', item: null }] });
  assert.deepEqual(frame(view).filters, [{ field: 'Product', item: null, items: ['Ink', 'Paper'] }]);
  assert.deepEqual(column(view, 6, 0, 5), ['Product', '', 'Region', 'North', 'South', 'Grand Total']);
  assert.equal(view.calc.getValue('Data', 0, 7), '(All)');
  assert.deepEqual(column(view, 7, 3, 5), [9, 7, 16]);

  view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty', subtotal: 'sum' }], filters: [{ field: 'Product', item: 'Ink' }] });
  assert.equal(frame(view).filters[0].item, 'Ink');
  assert.equal(view.calc.getValue('Data', 0, 7), 'Ink');
  assert.deepEqual(column(view, 7, 3, 5), [6, 6, 12], 'only the ink is counted');

  const saved = view.save();
  const xml = Workbook.open(saved).pkg.text('xl/pivotTables/pivotTable1.xml');
  assert.match(xml, /<pivotField axis="axisPage" showAll="0"><items count="3">/);
  assert.match(xml, /<\/colItems><pageFields count="1"><pageField fld="2" item="0" hier="-1"\/><\/pageFields><dataFields/);
  assert.match(xml, /<location ref="G3:H6" firstHeaderRow="0" firstDataRow="1" firstDataCol="1" rowPageCount="1" colPageCount="1"\/>/);

  // Opened again and refreshed, it is laid out where it was, still showing the ink.
  const again = SheetView.open(saved);
  assert.equal(frame(again).filters[0].item, 'Ink');
  again.refreshPivot('Sales');
  assert.deepEqual(column(again, 7, 0, 5), ['Ink', '', 'Sum of Qty', 6, 6, 12]);

  view.undo();
  view.undo();
  assert.deepEqual(frame(view).filters, []);
  assert.deepEqual(column(view, 6, 0, 3), ['Region', 'North', 'South', 'Grand Total'], 'back as it was');
  assert.throws(() => view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty' }], filters: [{ field: 'Product', item: 'Glue' }] }), /no item "Glue"/);
  assert.throws(() => view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty' }], filters: [{ field: 'Region', item: null }] }), /filter and on an axis/);
});

test('a date field shows by month, in calendar order, and keeps its months in the file', () => {
  const view = sales();
  assert.deepEqual(frame(view).dates, [true, false, false, false]);
  view.setPivotLayout('Sales', { rows: ['Date'], values: [{ field: 'Qty', subtotal: 'sum' }] });
  assert.equal(view.calc.getValue('Data', 6, 6), 'Grand Total', 'five dates, one line each');
  view.groupPivotField('Sales', 'Date', 'months');
  assert.deepEqual(frame(view).months, [true, false, false, false]);
  assert.deepEqual(column(view, 6, 0, 4), ['Date', 'Jan', 'Feb', 'Mar', 'Grand Total']);
  assert.deepEqual(column(view, 7, 1, 4), [10, 4, 2, 16]);
  assert.equal(view.calc.getValue('Data', 5, 6), '', 'the lines the dates took are cleared');

  const saved = view.save();
  const cache = Workbook.open(saved).pkg.text('xl/pivotCache/pivotCacheDefinition1.xml');
  assert.match(cache, /<sharedItems [^>]*containsDate="1"[^>]*><d v="2026-01-05T00:00:00"\/>/);
  assert.match(cache, /<fieldGroup base="0"><rangePr groupBy="months" startDate="2026-01-05T00:00:00" endDate="2026-03-10T00:00:00"\/><groupItems count="14"><s v="&lt;05\/01\/2026"\/><s v="Jan"\/>/);
  const table = Workbook.open(saved).pkg.text('xl/pivotTables/pivotTable1.xml');
  assert.match(table, /<rowItems count="13"><i><x v="1"\/><\/i>/);

  const again = SheetView.open(saved);
  assert.ok(!frame(again).unsupported, 'a pivot grouped by months can still be refreshed here');
  again.setCell(4, 3, 11);
  again.refreshPivot('Sales');
  assert.deepEqual(column(again, 7, 1, 4), [10, 14, 2, 26], 'February follows its edited figure');

  // A month as a filter.
  view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty', subtotal: 'sum' }], filters: [{ field: 'Date', item: 'Feb' }] });
  assert.deepEqual(frame(view).filters, [{ field: 'Date', item: 'Feb', items: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'] }]);
  assert.deepEqual(column(view, 7, 3, 5), [3, 1, 4]);

  // Shown by date again.
  view.setPivotLayout('Sales', { rows: ['Date'], values: [{ field: 'Qty', subtotal: 'sum' }], filters: [] });
  view.groupPivotField('Sales', 'Date', null);
  assert.deepEqual(frame(view).months, [false, false, false, false]);
  assert.equal(view.calc.getValue('Data', 6, 6), 'Grand Total');
  view.undo();
  assert.deepEqual(column(view, 6, 0, 4), ['Date', 'Jan', 'Feb', 'Mar', 'Grand Total'], 'undo shows the months again');
  assert.throws(() => view.groupPivotField('Sales', 'Region', 'months'), /only when it holds dates/);
});

test('a date as a filter item reads as the sheet shows it', () => {
  const view = sales();
  view.setPivotLayout('Sales', { rows: ['Region'], values: [{ field: 'Qty', subtotal: 'sum' }], filters: [{ field: 'Date', item: '2026-02-03' }] });
  const f = frame(view).filters[0];
  assert.equal(f.item, '2026-02-03');
  assert.deepEqual(f.items.slice(0, 2), ['2026-01-05', '2026-01-20']);
  assert.deepEqual(column(view, 7, 3, 4), [3, 3], 'North only, on the third of February');
});
