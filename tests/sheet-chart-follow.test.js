/**
 * A chart follows its cells: an edited figure redraws it at once, and the
 * file keeps the values it was drawn from. Before 2.18 a chart kept the
 * values it was opened with until the file was opened again elsewhere.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx } from '@rutba/ooxml/build';
import { refreshChartCaches, chartRefCells } from '../packages/sheet-view/src/chart-cache.js';

const book = () => buildXlsx({ sheets: [{ name: 'Data', rows: [['Month', 'Takings'], ['Jan', 120], ['Feb', 80], ['Mar', 95]] }] });

const charted = () => {
  const view = SheetView.open(book());
  view.select(1, 1);
  view.insertChart({ kind: 'column', title: 'Takings' });
  return view;
};

const chartSvg = (view) => view.render().drawings.find((d) => d.svg)?.svg ?? '';

test('a chart reference reads as its sheet and cells, quoted names and all', () => {
  assert.deepEqual(chartRefCells('Data!$B$2:$B$4'), { sheet: 'Data', cells: [{ row: 1, col: 1 }, { row: 2, col: 1 }, { row: 3, col: 1 }] });
  assert.equal(chartRefCells("'Jo''s sheet'!A1").sheet, "Jo's sheet");
  assert.equal(chartRefCells('MyName'), null, 'a name keeps its cache');
  assert.equal(chartRefCells('(Data!$A$1,Data!$A$3)'), null, 'several areas keep theirs');
});

test('a cache that agrees with its cells is left byte for byte', () => {
  const xml = '<c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f><c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val="2"/>'
    + '<c:pt idx="0"><c:v>120.0</c:v></c:pt><c:pt idx="1"><c:v>80</c:v></c:pt></c:numCache></c:numRef></c:val>';
  assert.equal(refreshChartCaches(xml, () => [120, 80]), xml);
  const moved = refreshChartCaches(xml, () => [120, 81]);
  assert.match(moved, /<c:formatCode>0\.0<\/c:formatCode><c:ptCount val="2"\/><c:pt idx="0"><c:v>120<\/c:v><\/c:pt><c:pt idx="1"><c:v>81<\/c:v><\/c:pt>/, 'its number format kept');
  assert.equal(refreshChartCaches(xml, () => null), xml, 'an unreadable reference keeps its cache');
});

test('an edited figure redraws the chart at once, and undo draws it back', () => {
  const view = charted();
  const before = chartSvg(view);
  assert.ok(before, 'the chart paints');
  view.select(2, 1);
  view.setCell(2, 1, '400');
  const after = chartSvg(view);
  assert.ok(after !== before && after.includes('>400</text>'), 'the Feb column changed with its cell');
  view.undo();
  assert.ok(chartSvg(view) === before, 'undo draws the old figure again');
});

test('the saved chart keeps the values it was drawn from, its labels as the cells read', () => {
  const view = charted();
  view.setCell(2, 1, '400');
  view.setCell(3, 0, 'March');
  const chart = Workbook.open(view.save()).pkg.text('xl/charts/chart1.xml');
  assert.match(chart, /<c:pt idx="1"><c:v>400<\/c:v><\/c:pt>/);
  assert.doesNotMatch(chart, /<c:v>80<\/c:v>/, 'the old figure is gone from the cache');
  assert.match(chart, /<c:v>March<\/c:v>/);

  // Opened again, the chart is drawn from the saved cache: the same picture.
  const again = SheetView.open(view.save());
  assert.ok(chartSvg(again) === chartSvg(view), 'the reopened chart is drawn as the saved one was');
  assert.ok(chartSvg(view).includes('>March</text>'), 'and still from the edited cells after the save');
});

test('an untouched workbook with a chart saves the chart part as it was', () => {
  const view = charted();
  const first = Workbook.open(view.save()).pkg.text('xl/charts/chart1.xml');
  const reopened = SheetView.open(view.save());
  assert.equal(Workbook.open(reopened.save()).pkg.text('xl/charts/chart1.xml'), first);
});
