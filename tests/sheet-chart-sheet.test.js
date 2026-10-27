/**
 * A chart sheet is a tab that holds one chart and no grid. Workbooks made by
 * Excel carry them often, and one used to stop the whole file opening. And
 * a chart read from a file is drawn as the file has it: the design rules a
 * chart made here follows (three scatter series, one pie series) are not a
 * reason to refuse someone else's.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx, OoxmlPackage } from '@rutba/ooxml';
import { Workbook } from '@rutba/ooxml/workbook';
import { buildChart, normaliseSpec } from '@rutba/drawing/chart';

const CHARTSHEET = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  + '<sheetPr/><sheetViews><sheetView zoomScale="94" workbookViewId="0"/></sheetViews>'
  + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
  + '<drawing r:id="rId1"/></chartsheet>';

function withChartSheet() {
  const pkg = OoxmlPackage.read(buildXlsx({ sheets: [{ name: 'Data', rows: [[1, 2]] }, { name: 'Chart1', rows: [[]] }] }));
  pkg.write_('xl/worksheets/sheet2.xml', CHARTSHEET);
  return pkg.write();
}

test('a workbook with a chart sheet opens, and the chart sheet reads as an empty grid', () => {
  const wb = Workbook.open(withChartSheet());
  assert.deepEqual(wb.sheetNames(), ['Data', 'Chart1']);
  assert.equal(wb.getCell('Data', 'A1'), 1);
  assert.ok([null, undefined, ''].includes(wb.getCell('Chart1', 'A1')));
});

test('saving a workbook keeps its chart sheet exactly as it was', () => {
  const wb = Workbook.open(withChartSheet());
  wb.setCell('Data', 'B1', 42);
  const again = OoxmlPackage.read(wb.save());
  assert.equal(again.text('xl/worksheets/sheet2.xml'), CHARTSHEET);
  assert.equal(Workbook.open(again).getCell('Data', 'B1'), 42);
});

const series = (n) => Array.from({ length: n }, (_, i) => ({ name: 'S' + i, values: [i, i + 1, i + 2], x: [1, 2, 3] }));

test('a scatter of more than three series is refused when made here, drawn when read from a file', () => {
  assert.throws(() => buildChart({ type: 'scatter', series: series(6) }), /capped at 3/);
  assert.ok(buildChart({ type: 'scatter', series: series(6), asWritten: true }));
  assert.equal(normaliseSpec({ type: 'scatter', series: series(6), asWritten: true }).series.length, 6, 'every series is kept');
});

test('a pie of several series is refused when made here, drawn from its first when read from a file', () => {
  const spec = { type: 'pie', categories: ['a', 'b', 'c'], series: series(2) };
  assert.throws(() => buildChart(spec), /one series/);
  assert.ok(buildChart({ ...spec, asWritten: true }));
  assert.deepEqual(normaliseSpec({ ...spec, asWritten: true }).series.map((s) => s.name), ['S0'], 'as Excel does, from the first');
});
