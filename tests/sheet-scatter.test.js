/**
 * Insert → Scatter: XY charts, markers only, with straight lines, or with
 * smooth lines — written as Excel writes `c:scatterChart` (xVal and yVal
 * number references, scatterStyle, each series' line, `c:smooth`, two
 * value axes), read back from Excel's own parts, and drawn by the chart kit
 * with a numeric X axis.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx, chartPartXml } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { buildChart, renderSvg, parseChartXml } from '@rutba/drawing';

const ROWS = [
  ['Hours', 'Score', 'Target'],
  [1, 52, 50],
  [2, 58, 55],
  [4, 71, 60],
  [7, 83, 65],
  [9, 90, 70],
];
const open = (rows = ROWS) => new SheetView(buildXlsx({ sheets: [{ name: 'Study', rows }] }));
const chartXml = (view) => {
  const name = view.pkg.partNames().find((n) => /^xl\/charts\/chart\d+\.xml$/.test(n));
  return name ? view.pkg.text(name) : null;
};

test('the writer puts a scatter chart the way Excel does, for each of its three looks', () => {
  const base = {
    kind: 'scatter',
    categories: { ref: 'S!$A$2:$A$4', values: [1, 2, 4] },
    series: [{ name: 'Score', nameRef: 'S!$B$1', ref: 'S!$B$2:$B$4', values: [52, 58, 71] }],
  };
  const markers = chartPartXml({ ...base, scatterStyle: 'markers' });
  assert.match(markers, /<c:scatterChart><c:scatterStyle val="lineMarker"\/><c:varyColors val="0"\/>/);
  assert.match(markers, /<c:spPr><a:ln w="19050" cap="rnd"><a:noFill\/><a:round\/><\/a:ln><\/c:spPr><c:marker>/, 'markers only: the series line is switched off');
  assert.match(markers, /<c:xVal><c:numRef><c:f>S!\$A\$2:\$A\$4<\/c:f><c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="3"\/><c:pt idx="0"><c:v>1<\/c:v><\/c:pt>/);
  assert.match(markers, /<c:yVal><c:numRef><c:f>S!\$B\$2:\$B\$4<\/c:f>/);
  assert.match(markers, /<\/c:yVal><c:smooth val="0"\/><\/c:ser>/);
  assert.equal((markers.match(/<c:valAx>/g) || []).length, 2, 'two value axes, no category axis');
  assert.doesNotMatch(markers, /<c:catAx>/);
  assert.match(markers, /<c:axPos val="b"\/>[\s\S]*<c:axPos val="l"\/>/);

  const lines = chartPartXml({ ...base, scatterStyle: 'lines' });
  assert.match(lines, /scatterStyle val="lineMarker"/);
  assert.doesNotMatch(lines, /<a:noFill\/>/);
  const smooth = chartPartXml({ ...base, scatterStyle: 'smooth' });
  assert.match(smooth, /scatterStyle val="smoothMarker"/);
  assert.match(smooth, /<c:smooth val="1"\/>/);

  // And read back as the look it was written as.
  assert.equal(parseChartXml(markers).scatterStyle, 'markers');
  assert.equal(parseChartXml(lines).scatterStyle, 'lines');
  assert.equal(parseChartXml(smooth).scatterStyle, 'smooth');
  assert.deepEqual(parseChartXml(markers).series[0].x, [1, 2, 4]);
});

test('Insert → Scatter takes the first column as X and the others as Y series', () => {
  const view = open();
  view.select(2, 1);
  view.insertChart({ kind: 'scatter', scatterStyle: 'lines' });
  const xml = chartXml(view);
  assert.match(xml, /<c:xVal><c:numRef><c:f>Study!\$A\$2:\$A\$6<\/c:f>/);
  assert.match(xml, /<c:tx><c:strRef><c:f>Study!\$B\$1<\/c:f>/);
  assert.match(xml, /<c:yVal><c:numRef><c:f>Study!\$C\$2:\$C\$6<\/c:f>/);
  const d = view.drawings.get('Study')[0];
  assert.equal(d.kind, 'chart');
  assert.equal(d.spec.type, 'scatter');
  assert.equal(d.spec.scatterStyle, 'lines');
  assert.deepEqual(d.spec.series.map((s) => s.x), [[1, 2, 4, 7, 9], [1, 2, 4, 7, 9]]);
  const frame = view.render();
  const drawn = frame.drawings.find((x) => x.kind === 'chart');
  assert.ok(drawn.svg && /<svg/.test(drawn.svg), 'drawn');
  assert.equal((drawn.svg.match(/<polyline/g) || []).length >= 2, true, 'a straight line per series');
  view.undo();
  assert.equal(chartXml(view), null, 'undo takes the chart away');
});

test('the X axis is numeric: points sit by their X values, not evenly spaced', () => {
  const spec = {
    type: 'scatter', scatterStyle: 'markers', width: 400, height: 300,
    series: [{ name: 'S', values: [10, 20, 30], x: [0, 1, 10] }],
  };
  const svg = renderSvg(buildChart(spec));
  const cx = [...svg.matchAll(/<ellipse[^>]*cx="([\d.]+)"/g)].map((m) => Number(m[1]));
  assert.equal(cx.length, 3);
  const gapA = cx[1] - cx[0];
  const gapB = cx[2] - cx[1];
  assert.ok(gapB > gapA * 8, `the third point is nine units along, not one: ${cx.join(', ')}`);
  assert.doesNotMatch(svg, /<polyline[^>]*stroke="#/, 'markers only draws no line');
  // Tick labels along the foot are numbers.
  assert.match(svg, />10</);
});

test('smooth lines are curves through the points; straight ones are polylines', () => {
  const series = [{ name: 'S', values: [1, 4, 2, 5], x: [1, 2, 3, 4] }];
  const smooth = renderSvg(buildChart({ type: 'scatter', scatterStyle: 'smooth', series }));
  assert.match(smooth, /<path d="M[\d.]+ [\d.]+C/);
  const lines = renderSvg(buildChart({ type: 'scatter', scatterStyle: 'lines', series }));
  assert.match(lines, /<polyline/);
});

test('a first column of words is labels: no X values, the points go 1, 2, 3', () => {
  const view = open([['Name', 'A', 'B'], ['x', 1, 2], ['y', 3, 4], ['z', 5, 6]]);
  view.select(1, 1);
  view.insertChart({ kind: 'scatter' });
  const xml = chartXml(view);
  assert.doesNotMatch(xml, /<c:xVal>/);
  const spec = view.drawings.get('Study')[0].spec;
  assert.equal(spec.series[0].x, undefined);
  assert.ok(renderSvg(buildChart(spec)).includes('<ellipse'));
});

test('an Excel-authored scatter reads: text X values plot 1, 2, 3; its lines and style come through', () => {
  const excel = '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>'
    + '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>'
    + '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Speed</c:v></c:tx><c:spPr><a:ln w="19050"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></c:spPr>'
    + '<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:marker>'
    + '<c:xVal><c:strRef><c:f>S!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:strCache></c:strRef></c:xVal>'
    + '<c:yVal><c:numRef><c:f>S!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:yVal><c:smooth val="0"/></c:ser>'
    + '</c:scatterChart></c:plotArea></c:chart></c:chartSpace>';
  const spec = parseChartXml(excel);
  assert.equal(spec.type, 'scatter');
  assert.equal(spec.scatterStyle, 'lines', 'the marker\'s own noFill is not the series line');
  assert.equal(spec.series[0].x, undefined);
  assert.deepEqual(spec.series[0].values, [3, 5]);
});

test('more than three Y columns is refused with a reason, and the look must be one of the three', () => {
  const view = open([['X', 'a', 'b', 'c', 'd'], [1, 1, 2, 3, 4], [2, 2, 3, 4, 5]]);
  view.select(1, 1);
  assert.throws(() => view.insertChart({ kind: 'scatter' }), /up to three series/);
  assert.throws(() => view.insertChart({ kind: 'scatter', scatterStyle: 'bubbles' }), /not a scatter look/);
});
