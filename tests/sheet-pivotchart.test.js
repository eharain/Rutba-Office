// Insert → PivotChart: a chart bound to a pivot table.
//
// Written as Excel writes one — `c:pivotSource` naming the pivot, `c:pivotFmts`
// in the chart, series whose references point into the pivot's rectangle and
// whose caches hold its current figures — and drawn by the chart kit from
// them. A refresh of the pivot, or a slicer's filter, redraws the chart in the
// same undo step. On data, PivotChart & PivotTable makes both at once.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { readPivotChart, parsePivotSource } from '@rutba/ooxml/pivot';
import { SheetView } from '@rutba/sheet-view';

const ROWS = [
  ['Region', 'Rep', 'Units'],
  ['North', 'Amira', 10],
  ['South', 'Ben', 20],
  ['North', 'Chen', 5],
  ['East', 'Dana', 7],
  ['South', 'Amira', 3],
  ['East', 'Ben', 12],
];
const BOOK = () => buildXlsx({ sheets: [{ name: 'Sales', rows: ROWS }] });

const chartParts = (view) => view.pkg.partNames().filter((p) => /^xl\/charts\/chart\d+\.xml$/.test(p));
const cached = (xml, tag) => {
  const block = new RegExp('<c:' + tag + '>([\\s\\S]*?)</c:' + tag + '>').exec(xml)?.[1] ?? '';
  return [...block.matchAll(/<c:v>([^<]*)<\/c:v>/g)].map((m) => m[1]);
};

test('PivotChart & PivotTable on data: the pivot, and a chart bound to it as Excel writes one', () => {
  const view = new SheetView(BOOK());
  view.select(1, 0);
  const made = view.createPivot({
    source: 'Sales!A1:C7', rowFields: ['Region'], dataFields: [{ name: 'Units', fn: 'SUM' }],
    chart: { kind: 'column' }, fileName: 'Report.xlsx',
  });
  assert.equal(made.name, 'PivotTable1', 'a pivot is named as Excel names its first');
  const [part] = chartParts(view);
  assert.ok(part, 'a chart part');
  const xml = view.pkg.text(part);
  const info = readPivotChart(xml);
  assert.deepEqual([info.name, info.sheet, info.kind], ['PivotTable1', 'Sales', 'column']);
  assert.match(xml, /<c:pivotSource><c:name>\[Report\.xlsx\]Sales!PivotTable1<\/c:name><c:fmtId val="0"\/><\/c:pivotSource><c:chart>/, 'c:pivotSource before c:chart');
  assert.match(xml, /<c:pivotFmts><c:pivotFmt><c:idx val="0"\/><\/c:pivotFmt><\/c:pivotFmts><c:plotArea>/, 'c:pivotFmts before the plot area');
  // The pivot sits at A9 (two rows under the data): header at row 9, East/North/South at 10-12.
  assert.match(xml, /<c:cat><c:strRef><c:f>Sales!\$A\$10:\$A\$12<\/c:f>/, 'categories are the row labels');
  assert.match(xml, /<c:val><c:numRef><c:f>Sales!\$B\$10:\$B\$12<\/c:f>/, 'values are the pivot column');
  assert.deepEqual(cached(xml, 'cat'), ['East', 'North', 'South']);
  assert.deepEqual(cached(xml, 'val'), ['19', '15', '23']);

  const frame = view.render();
  const chart = frame.drawings.find((d) => d.kind === 'chart');
  assert.equal(chart.pivot, 'PivotTable1', 'the frame knows it is a PivotChart');
  assert.match(chart.svg, /<rect\b/, 'drawn by the chart kit');
  const saved = SheetView.open(view.save());
  assert.equal(saved.render().drawings.find((d) => d.kind === 'chart')?.pivot, 'PivotTable1', 'and so does the reopened file');
});

test('a refresh redraws the PivotChart from the pivot, in the same undo step', () => {
  const view = new SheetView(BOOK());
  view.createPivot({ source: 'Sales!A1:C7', rowFields: ['Region'], dataFields: [{ field: 'Units' }], chart: { kind: 'bar', title: 'Units by region' } });
  const [part] = chartParts(view);
  view.setCell(1, 2, 100); // North, Amira: 10 -> 100
  view.refreshPivot('PivotTable1');
  let xml = view.pkg.text(part);
  assert.deepEqual(cached(xml, 'val'), ['19', '105', '23'], 'North follows the source');
  assert.equal(readPivotChart(xml).kind, 'bar', 'kind kept');
  assert.equal(readPivotChart(xml).title, 'Units by region', 'title kept');
  const svg = view.render().drawings.find((d) => d.kind === 'chart').svg;
  assert.match(svg, /105/, 'the drawing shows the new figure');
  view.undo();
  xml = view.pkg.text(part);
  assert.deepEqual(cached(xml, 'val'), ['19', '15', '23'], 'undo puts the chart back with the pivot');
  // Refresh All: with no name, every pivot on the sheet.
  // (the source still says 100: only the refresh was undone)
  view.setCell(3, 2, 50);
  view.refreshPivot();
  assert.deepEqual(cached(view.pkg.text(part), 'val'), ['19', '150', '23']);
});

test('Insert → PivotChart on a pivot: a chart for the pivot the cursor is in; elsewhere it says why', () => {
  const view = new SheetView(BOOK());
  view.createPivot({ source: 'Sales!A1:C7', rowFields: ['Region'], colFields: [], dataFields: [{ field: 'Units' }] });
  view.select(0, 5);
  assert.throws(() => view.insertPivotChart({ kind: 'line' }), /cursor in a pivot table/);
  view.select(9, 1);
  view.insertPivotChart({ kind: 'line', fileName: 'a.xlsx' });
  const parts = chartParts(view);
  assert.equal(parts.length, 1);
  assert.equal(readPivotChart(view.pkg.text(parts[0])).kind, 'line');
  assert.throws(() => view.insertPivotChart({ kind: 'scatter' }), /not a PivotChart kind/);
  view.undo();
  assert.equal(chartParts(view).length, 0, 'undo takes the chart and its part away');
});

test('a pivot with a column field charts one series per column item; a hidden item leaves chart and grid', async () => {
  const view = new SheetView(BOOK());
  view.createPivot({ source: 'Sales!A1:C7', rowFields: ['Region'], colFields: ['Rep'], dataFields: [{ field: 'Units' }], chart: { kind: 'column' } });
  const [part] = chartParts(view);
  let xml = view.pkg.text(part);
  const names = [...xml.matchAll(/<c:tx><c:strRef><c:f>[^<]*<\/c:f><c:strCache><c:ptCount val="1"\/><c:pt idx="0"><c:v>([^<]*)<\/c:v>/g)].map((m) => m[1]);
  assert.deepEqual(names, ['Amira', 'Ben', 'Chen', 'Dana'], 'a series per rep, grand total left out');
  assert.match(xml, /<c:legend>/, 'several series get a legend');

  // A slicer's choice (hidden items in the field) through the engine's own path.
  view.select(9, 1);
  view.insertSlicers({ fields: ['Region'] });
  view.setSlicerSelection({ name: 'Region', values: ['North', 'South'] });
  xml = view.pkg.text(part);
  assert.deepEqual(cached(xml, 'cat'), ['North', 'South'], 'East is filtered out of the chart');
  const grid = view.pivots().find((p) => p.name === 'PivotTable1');
  assert.ok(grid.hidden.get(0)?.size === 1, 'one item of Region is hidden in the pivot');
});

test('the pivot source a chart names is read whatever the file or sheet is called', () => {
  assert.deepEqual(parsePivotSource("[Book 1.xlsx]'Sales 2026'!PivotTable3"), { sheet: 'Sales 2026', name: 'PivotTable3' });
  assert.deepEqual(parsePivotSource('Sheet2!PivotTable1'), { sheet: 'Sheet2', name: 'PivotTable1' });
});
