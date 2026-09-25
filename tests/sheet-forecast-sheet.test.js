/**
 * Data → Forecast Sheet: a new sheet, in front of the data's, with Excel's
 * table — the history, FORECAST.ETS past its last point and the
 * confidence bounds from FORECAST.ETS.CONFINT — and a chart of it, the
 * series in Excel's colours without markers, the bounds drawn as one band.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { buildChart, renderSvg } from '@rutba/drawing';
import { parseChartXml } from '../packages/drawing/src/ooxml.js';

const serial = (y, m) => Date.UTC(y, m, 1) / 86400000 + 25569;
const HISTORY = Array.from({ length: 36 }, (_, i) => [serial(2023, i), Math.round(1000 + 12 * i + 180 * Math.sin((2 * Math.PI * i) / 12) + 25 * Math.sin(i * 1.7))]);
const open = () => new SheetView(buildXlsx({ sheets: [{
  name: 'Sales',
  rows: [['Month', 'Revenue'], ...HISTORY],
  styles: { 'A2:A37': { numFmt: 'mmm yyyy' }, 'B2:B37': { numFmt: '#,##0' } },
}] }));

test('the dialog opens on the data round the cell: the timeline, the values, and an end a third as far again', () => {
  const view = open();
  view.select(4, 1);
  const info = view.forecastInfo();
  assert.equal(info.timeline, 'Sales!$A$2:$A$37');
  assert.equal(info.values, 'Sales!$B$2:$B$37');
  assert.equal(info.header, true);
  assert.equal(info.isDate, true);
  assert.ok(info.end > info.last);
});

test('Forecast Sheet makes Excel\'s table on a new sheet in front of the data: history, FORECAST.ETS and the bounds', () => {
  const view = open();
  const made = view.forecastSheet({ timeline: 'Sales!$A$2:$A$37', values: 'Sales!$B$2:$B$37', end: serial(2026, 11) });
  assert.equal(made.sheet, 'Sheet2');
  assert.deepEqual(view.sheetNames(), ['Sheet2', 'Sales'], 'in front of the data, as Excel puts it');
  assert.equal(view.activeSheet, 'Sheet2');
  const input = (r, c) => view.calc.getInput('Sheet2', r, c);
  const value = (r, c) => view.calc.getValue('Sheet2', r, c);
  assert.deepEqual([0, 1, 2, 3, 4].map((c) => input(0, c)), ['Month', 'Revenue', 'Forecast(Revenue)', 'Lower Confidence Bound(Revenue)', 'Upper Confidence Bound(Revenue)']);
  assert.equal(value(36, 1), HISTORY[35][1], 'the history copied');
  assert.equal(input(36, 2), '=B37', 'the last actual point repeated, so the lines meet');
  assert.equal(made.forecast, 12, 'twelve months, to December 2026');
  assert.equal(value(37, 0), serial(2026, 0), 'month-start dates stay month starts');
  assert.equal(value(48, 0), serial(2026, 11));
  assert.equal(input(37, 2), '=FORECAST.ETS(A38,$B$2:$B$37,$A$2:$A$37,1,1)');
  assert.equal(input(37, 3), '=C38-FORECAST.ETS.CONFINT(A38,$B$2:$B$37,$A$2:$A$37,0.95,1,1)');
  const f = value(37, 2);
  const lo = value(37, 3);
  const hi = value(37, 4);
  assert.ok(typeof f === 'number' && lo < f && f < hi, `${lo} < ${f} < ${hi}`);
  assert.ok(Math.abs(f - (1000 + 12 * 36 + 180 * Math.sin(0))) < 150, 'a sensible January');
  assert.ok(value(48, 4) - value(48, 3) > hi - lo, 'wider a year out');
  const pkg = OoxmlPackage.read(view.save());
  const tablePart = pkg.partNames().find((p) => /^xl\/tables\/table\d+\.xml$/.test(p));
  assert.match(pkg.text(tablePart), /ref="A1:E49"/);
  assert.match(pkg.text(tablePart), /<tableColumn id="3" name="Forecast\(Revenue\)"\/>/);
});

test('the chart: four series in Excel\'s colours without markers, drawn with the bounds as one band', () => {
  const view = open();
  view.forecastSheet({ timeline: 'Sales!$A$2:$A$37', values: 'Sales!$B$2:$B$37', end: serial(2026, 11) });
  const pkg = OoxmlPackage.read(view.save());
  const chart = pkg.text(pkg.partNames().find((p) => /^xl\/charts\/chart\d+\.xml$/.test(p)));
  assert.match(chart, /<c:lineChart>/);
  assert.equal((chart.match(/<c:ser>/g) || []).length, 4);
  assert.match(chart, /<c:f>Sheet2!\$C\$2:\$C\$49<\/c:f>/);
  assert.match(chart, /<a:srgbClr val="4472C4"\/>/);
  assert.match(chart, /<a:srgbClr val="ED7D31"\/>/);
  assert.equal((chart.match(/<c:symbol val="none"\/>/g) || []).length, 4);
  const spec = parseChartXml(chart, { width: 640, height: 360 });
  assert.equal(spec.series[0].markers, false);
  const svg = renderSvg(buildChart(spec));
  assert.match(svg, /class="band"/, 'the confidence band');
  assert.doesNotMatch(svg, /<ellipse/, 'no markers');
  const drawn = view.render().drawings.find((d) => d.kind === 'chart');
  assert.ok(drawn && /class="band"/.test(drawn.svg), 'drawn on the new sheet');
});

test('columns without the bounds, no confidence interval, a season set by hand; one undo takes it all away', () => {
  const view = open();
  view.forecastSheet({ timeline: 'Sales!$A$2:$A$37', values: 'Sales!$B$2:$B$37', end: serial(2026, 5), kind: 'column', confidence: null, seasonality: 12, aggregation: 7 });
  assert.equal(view.calc.getInput('Sheet2', 0, 3), '', 'no bound columns');
  assert.equal(view.calc.getInput('Sheet2', 37, 2), '=FORECAST.ETS(A38,$B$2:$B$37,$A$2:$A$37,12,1,7)');
  const chart = OoxmlPackage.read(view.save()).partNames().find((p) => /^xl\/charts\//.test(p));
  assert.match(OoxmlPackage.read(view.save()).text(chart), /<c:barChart>/);
  view.undo();
  assert.deepEqual(view.sheetNames(), ['Sales']);
  assert.ok(!OoxmlPackage.read(view.save()).partNames().some((p) => /^xl\/(charts|tables)\//.test(p)), 'the chart and table parts gone too');
});

test('refusals: an end before the last point, an uneven timeline, a locked workbook', () => {
  const view = open();
  assert.throws(() => view.forecastSheet({ timeline: 'Sales!$A$2:$A$37', values: 'Sales!$B$2:$B$37', end: serial(2024, 0) }), /after the last point/);
  view.setCell(5, 0, serial(2023, 3) + 15);
  assert.throws(() => view.forecastSheet({ timeline: 'Sales!$A$2:$A$37', values: 'Sales!$B$2:$B$37', end: serial(2026, 11) }), /constant step/);
  const locked = open();
  locked.protectWorkbook();
  assert.throws(() => locked.forecastSheet({ timeline: 'Sales!$A$2:$A$37', values: 'Sales!$B$2:$B$37', end: serial(2026, 11) }), /Workbook is protected/);
});
