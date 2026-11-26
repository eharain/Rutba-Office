/**
 * Chart Design: axis titles, data labels, the legend's place and a series'
 * own colour. Before 2.28 the tab held the chart's title and nothing else,
 * so a chart could not be made clear. Each choice is written into the chart
 * part the way Excel writes it, drawn at once, and read back from the file.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { Workbook } from '@rutba/ooxml';
import { buildXlsx, chartPartXml } from '@rutba/ooxml/build';
import {
  buildChart, renderSvg, parseChartXml, chartShows,
  chartLook, setAxisTitle, setDataLabels, setLegend, setSeriesColour,
} from '@rutba/drawing';

const part = (kind = 'column', series = 2, extra = {}) => chartPartXml({
  kind,
  title: 'Takings',
  categories: { ref: 'Data!$A$2:$A$4', values: kind === 'scatter' ? [1, 2, 3] : ['Jan', 'Feb', 'Mar'] },
  series: [
    { name: 'Shop', ref: 'Data!$B$2:$B$4', values: [120, 80, 95] },
    { name: 'Online', ref: 'Data!$C$2:$C$4', values: [60, 75, 110] },
  ].slice(0, series),
  ...extra,
});

const svgOf = (spec) => renderSvg(buildChart(spec));
const spec = (over = {}) => ({
  type: 'column',
  title: 'Takings',
  categories: ['Jan', 'Feb', 'Mar'],
  series: [{ name: 'Shop', values: [120, 80, 95] }, { name: 'Online', values: [60, 75, 110] }],
  ...over,
});

// ---- the chart part --------------------------------------------------------

test('a chart made here leaves its legend and its labels open, and names no axis', () => {
  const look = chartLook(part());
  assert.deepEqual(look, {
    axes: true, categoryAxisTitle: null, valueAxisTitle: null,
    dataLabels: null, legend: null, slices: false, seriesColours: [null, null],
  });
  assert.equal(chartLook(part('pie', 1)).axes, false);
  assert.equal(chartLook(part('pie', 1)).slices, true);
});

test('an axis title goes where the schema has it, upright beside a vertical axis, and comes away again', () => {
  let xml = setAxisTitle(part(), 'value', 'Sales & tips (£)');
  assert.equal(chartLook(xml).valueAxisTitle, 'Sales & tips (£)');
  assert.equal(chartLook(xml).categoryAxisTitle, null);
  assert.match(xml, /<c:valAx>[\s\S]*<c:axPos val="l"\/><c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"\/>[\s\S]*<a:t>Sales &amp; tips \(£\)<\/a:t>[\s\S]*<\/c:title><c:crossAx/);
  xml = setAxisTitle(xml, 'category', 'Month');
  assert.match(xml, /<c:catAx>[\s\S]*<c:axPos val="b"\/><c:title><c:tx><c:rich><a:bodyPr\/>[\s\S]*<a:t>Month<\/a:t>[\s\S]*<\/c:title><c:crossAx/);
  assert.equal(chartLook(xml).categoryAxisTitle, 'Month');
  // The chart's own title is not an axis's, and stays.
  assert.equal(parseChartXml(xml).title, 'Takings');
  assert.equal(parseChartXml(xml).valueAxis.title, 'Sales & tips (£)');
  assert.equal(parseChartXml(xml).categoryAxis.title, 'Month');
  // Named again, an axis has one title, not two.
  xml = setAxisTitle(xml, 'value', 'Pounds');
  assert.equal((xml.match(/<c:title>/g) || []).length, 3, 'the chart, and each axis once');
  xml = setAxisTitle(setAxisTitle(xml, 'value', ''), 'category', '  ');
  assert.equal((xml.match(/<c:title>/g) || []).length, 1);
  // A scatter's X axis is its category axis; a pie has none to name.
  const xy = setAxisTitle(setAxisTitle(part('scatter'), 'category', 'Week'), 'value', 'Height');
  assert.match(xy, /<c:valAx>[\s\S]*<c:axPos val="b"\/><c:title>[\s\S]*<a:t>Week<\/a:t>[\s\S]*<\/c:valAx><c:valAx>[\s\S]*<c:majorGridlines\/><c:title>[\s\S]*<a:t>Height<\/a:t>/);
  assert.throws(() => setAxisTitle(part('pie', 1), 'value', 'x'), /no axes/);
});

test('the first choice writes the chart down as it stood, so nothing else on it moves', () => {
  // Two series of columns are drawn with a legend at the foot and their values on them.
  const moved = setLegend(part(), 'right');
  assert.match(moved, /<\/c:plotArea><c:legend><c:legendPos val="r"\/><c:overlay val="0"\/><\/c:legend><c:plotVisOnly val="1"\/><c:dispBlanksAs val="gap"\/><\/c:chart>/);
  assert.equal(chartLook(moved).legend, 'right');
  assert.equal(chartLook(moved).dataLabels, true, 'the labels it was drawn with are written down');
  assert.match(moved, /<\/c:ser><c:dLbls>[\s\S]*<c:showVal val="1"\/>[\s\S]*<\/c:dLbls><c:axId/);

  const bare = setDataLabels(part(), false);
  assert.equal(chartLook(bare).dataLabels, false);
  assert.equal(chartLook(bare).legend, 'bottom', 'the legend it was drawn with is written down');
  assert.doesNotMatch(bare, /<c:dLbls>/);

  // One series has no legend by the house rule, and none is written for it.
  assert.equal(chartLook(setDataLabels(part('column', 1), false)).legend, 'none');
  // A line's values were never written on it.
  assert.equal(chartLook(setLegend(part('line'), 'top')).dataLabels, false);
  // A pie shows shares, and a legend for its slices.
  const pie = setLegend(part('pie', 1), 'left');
  assert.match(pie, /<\/c:ser><c:dLbls>[\s\S]*<c:showVal val="0"\/>[\s\S]*<c:showPercent val="1"\/>[\s\S]*<\/c:dLbls><c:firstSliceAng/);
  assert.equal(chartLook(pie).dataLabels, true);
});

test('the legend moves, goes and comes back; data labels go on and off', () => {
  let xml = part();
  for (const [place, letter] of [['right', 'r'], ['top', 't'], ['left', 'l'], ['bottom', 'b']]) {
    xml = setLegend(xml, place);
    assert.equal(chartLook(xml).legend, place);
    assert.ok(xml.includes('<c:legendPos val="' + letter + '"/>'));
    assert.equal((xml.match(/<c:legend>/g) || []).length, 1);
  }
  xml = setLegend(xml, 'none');
  assert.doesNotMatch(xml, /<c:legend>/);
  assert.equal(chartLook(xml).legend, 'none', 'taken away, it is not drawn by rule again');
  assert.equal(parseChartXml(xml).legend, 'none');
  assert.equal(chartLook(setLegend(xml, 'top')).legend, 'top');
  assert.throws(() => setLegend(xml, 'middle'), /right, the top, the left or the bottom/);

  xml = setDataLabels(xml, false);
  assert.equal(parseChartXml(xml).dataLabels, false);
  xml = setDataLabels(setDataLabels(xml, true), true);
  assert.equal((xml.match(/<c:dLbls>/g) || []).length, 1, 'once, however often it is switched on');
  assert.equal(parseChartXml(xml).dataLabels, true);

  // A legend Excel wrote keeps what it holds, and loses only its place.
  const excel = part().replace('<c:plotVisOnly', '<c:legend><c:legendPos val="r"/><c:layout><c:manualLayout/></c:layout><c:overlay val="0"/><c:txPr><a:bodyPr/></c:txPr></c:legend><c:plotVisOnly');
  assert.equal(chartLook(excel).legend, 'right');
  assert.match(setLegend(excel, 'bottom'), /<c:legend><c:legendPos val="b"\/><c:overlay val="0"\/><c:txPr><a:bodyPr\/><\/c:txPr><\/c:legend>/);
});

test('a series takes a colour of its own, and gives it back', () => {
  let xml = setSeriesColour(part(), 1, '#c00000');
  assert.deepEqual(chartLook(xml).seriesColours, [null, '#C00000']);
  assert.match(xml, /<c:v>Online<\/c:v><\/c:tx><c:spPr><a:solidFill><a:srgbClr val="C00000"\/><\/a:solidFill><\/c:spPr><c:cat>/);
  assert.deepEqual(parseChartXml(xml).series.map((s) => s.colour ?? null), [null, '#C00000']);
  xml = setSeriesColour(xml, 1, '#1F6FB2');
  assert.equal((xml.match(/<c:spPr>/g) || []).length, 1, 'coloured again, one shape, not two');
  assert.deepEqual(chartLook(setSeriesColour(xml, 1, null)).seriesColours, [null, null]);
  assert.equal(setSeriesColour(setSeriesColour(part(), 0, '#123456'), 0, null), part(), 'given back, the part is as it was');

  // A line is drawn in the colour, and its markers wear it.
  const line = setSeriesColour(part('line'), 0, '#008000');
  assert.match(line, /<c:v>Shop<\/c:v><\/c:tx><c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="008000"\/><\/a:solidFill><a:round\/><\/a:ln><\/c:spPr><c:marker><c:spPr><a:solidFill><a:srgbClr val="008000"\/>[\s\S]*<\/c:marker><c:cat>/);
  assert.deepEqual(chartLook(line).seriesColours, ['#008000', null]);
  assert.equal(setSeriesColour(line, 0, null), part('line'));
  // A scatter of markers alone keeps its line off, and colours the markers.
  const dots = setSeriesColour(part('scatter'), 0, '#7030A0');
  assert.match(dots, /<a:ln w="19050" cap="rnd"><a:noFill\/><a:round\/><\/a:ln><\/c:spPr><c:marker><c:symbol val="circle"\/><c:size val="5"\/><c:spPr><a:solidFill><a:srgbClr val="7030A0"\/>/);
  assert.equal(chartLook(dots).seriesColours[0], '#7030A0');
  assert.equal(parseChartXml(dots).scatterStyle, 'markers');

  assert.throws(() => setSeriesColour(part(), 5, '#000000'), /no series 6/);
  assert.throws(() => setSeriesColour(part(), 0, 'red'), /#1F6FB2/);
  assert.throws(() => setSeriesColour(part('pie', 1), 0, '#000000'), /slice/);
});

test('a chart from Excel is read as it is written: no legend there, none here', () => {
  // Excel always says how blanks are shown; a part that does has made its choices.
  const excel = part().replace('<c:plotVisOnly val="1"/>', '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>');
  assert.equal(chartLook(excel).legend, 'none');
  assert.equal(chartLook(excel).dataLabels, false);
  assert.equal(setLegend(excel, 'none'), excel, 'nothing to write down');
  assert.deepEqual(chartShows(parseChartXml(excel)), { legend: 'none', dataLabels: false });
  assert.deepEqual(chartShows(parseChartXml(part())), { legend: 'bottom', dataLabels: true });
  assert.deepEqual(chartShows(parseChartXml(part('line', 1))), { legend: 'none', dataLabels: false });
  assert.deepEqual(chartShows(parseChartXml(part('pie', 1))), { legend: 'bottom', dataLabels: true });
});

// ---- the drawing -----------------------------------------------------------

test('left open, a chart is drawn as it always was, and the choices written down draw the same', () => {
  assert.equal(svgOf(spec({ legend: 'bottom', dataLabels: true })), svgOf(spec()));
  assert.equal(svgOf(spec({ type: 'line', legend: 'bottom', dataLabels: false })), svgOf(spec({ type: 'line' })));
  const pie = { type: 'pie', categories: ['a', 'b', 'c'], series: [{ name: 'S', values: [3, 2, 1] }] };
  assert.equal(svgOf({ ...pie, legend: 'bottom', dataLabels: true }), svgOf(pie));
});

test('the legend is drawn where it was put, or not at all', () => {
  const xOf = (svg, word) => Number(new RegExp('<text[^>]*\\bx="([\\d.]+)"[^>]*>' + word + '<').exec(svg)?.[1]);
  const yOf = (svg, word) => Number(new RegExp('<text[^>]*\\by="([\\d.]+)"[^>]*>' + word + '<').exec(svg)?.[1]);
  const none = svgOf(spec({ legend: 'none' }));
  assert.doesNotMatch(none, /class="legend"/);
  const foot = svgOf(spec());
  const right = svgOf(spec({ legend: 'right' }));
  const left = svgOf(spec({ legend: 'left' }));
  const top = svgOf(spec({ legend: 'top' }));
  for (const svg of [foot, right, left, top]) assert.match(svg, /class="legend"/);
  assert.ok(xOf(right, 'Online') > 480 * 0.6, 'at the right');
  assert.ok(xOf(left, 'Online') < 480 * 0.2, 'at the left');
  assert.ok(yOf(right, 'Online') > yOf(right, 'Shop'), 'one name under another');
  assert.ok(yOf(top, 'Shop') < 80 && yOf(foot, 'Shop') > 250, 'at the top, and at the foot');
  // One series has no legend by rule, and has one when it is asked for.
  const one = spec({ series: [{ name: 'Shop', values: [120, 80, 95] }] });
  assert.doesNotMatch(svgOf(one), /class="legend"/);
  assert.match(svgOf({ ...one, legend: 'right' }), /class="legend"/);
  // A pie's legend names its slices, wherever it is.
  const pie = { type: 'pie', categories: ['Tea', 'Coffee'], series: [{ name: 'S', values: [3, 2] }] };
  assert.ok(xOf(svgOf({ ...pie, legend: 'right' }), 'Coffee') > 480 * 0.6);
  assert.doesNotMatch(svgOf({ ...pie, legend: 'none' }), /class="legend"/);
});

test('axis titles are drawn: under the foot, and reading upward at the side', () => {
  const svg = svgOf(spec({ valueAxis: { title: 'Pounds' }, categoryAxis: { title: 'Month' } }));
  assert.match(svg, /<g transform="rotate\(-90 [\d.]+ [\d.]+\)" class="axis-title"><text[^>]*>Pounds</);
  assert.match(svg, /<g class="axis-title"><text[^>]*>Month</);
  // A bar chart lies on its side: its categories stand at the left.
  const bar = svgOf(spec({ type: 'bar', valueAxis: { title: 'Pounds' }, categoryAxis: { title: 'Month' } }));
  assert.match(bar, /<g transform="rotate\(-90 [\d.]+ [\d.]+\)" class="axis-title"><text[^>]*>Month</);
  assert.match(bar, /<g class="axis-title"><text[^>]*>Pounds</);
  const xy = svgOf({ type: 'scatter', categories: [1, 2, 3], series: [{ name: 'H', values: [4, 5, 9], x: [1, 2, 3] }], valueAxis: { title: 'Height' }, categoryAxis: { title: 'Week' } });
  assert.match(xy, /rotate\(-90[^>]*><text[^>]*>Height</);
  assert.match(xy, /<g class="axis-title"><text[^>]*>Week</);
  assert.doesNotMatch(svgOf(spec()), /axis-title/);
});

test('data labels are drawn on, and left off, as the chart says', () => {
  const values = (svg) => ['120', '80', '95'].filter((v) => new RegExp('>' + v + '<').test(svg)).length;
  assert.equal(values(svgOf(spec({ series: [{ name: 'Shop', values: [120, 80, 95] }] }))), 3, 'a few bars carry their values by rule');
  assert.equal(values(svgOf(spec({ dataLabels: false, series: [{ name: 'Shop', values: [120, 80, 95] }] }))), 0);
  assert.equal(values(svgOf(spec({ type: 'line', series: [{ name: 'Shop', values: [120, 80, 95] }] }))), 0, 'a line carries none by rule');
  assert.equal(values(svgOf(spec({ type: 'line', dataLabels: true, series: [{ name: 'Shop', values: [120, 80, 95] }] }))), 3);
  assert.equal(values(svgOf(spec({ stacked: true }))), 0, 'a stack carries none by rule');
  assert.equal(values(svgOf(spec({ stacked: true, dataLabels: true }))), 3, 'and each piece its own when asked');
  const xy = { type: 'scatter', categories: [1, 2, 3], series: [{ name: 'H', values: [120, 80, 95], x: [1, 2, 3] }] };
  assert.equal(values(svgOf(xy)), 0);
  assert.equal(values(svgOf({ ...xy, dataLabels: true })), 3, 'each point its Y value');
  const pie = { type: 'pie', categories: ['a', 'b'], series: [{ name: 'S', values: [3, 1] }] };
  assert.match(svgOf(pie), />75%</);
  assert.doesNotMatch(svgOf({ ...pie, dataLabels: false }), />75%</);
});

test('a series is drawn in its own colour, in the legend too', () => {
  const svg = svgOf(spec({ series: [{ name: 'Shop', values: [120, 80, 95] }, { name: 'Online', values: [60, 75, 110], colour: '#C00000' }] }));
  assert.equal((svg.match(/fill="#C00000"/g) || []).length, 4, 'three columns and the legend');
  const line = svgOf(spec({ type: 'line', series: [{ name: 'Shop', values: [120, 80, 95], colour: '#008000' }] }));
  assert.match(line, /stroke="#008000"/);
  assert.doesNotMatch(svgOf(spec({ series: [{ name: 'Shop', values: [1, 2, 3], colour: 'red' }] })), /fill="red"/, 'only a colour written as #RRGGBB is taken');
});

// ---- the sheet -------------------------------------------------------------

const book = () => buildXlsx({ sheets: [{ name: 'Data', rows: [['Month', 'Shop', 'Online'], ['Jan', 120, 60], ['Feb', 80, 75], ['Mar', 95, 110]] }] });
const charted = (kind = 'column') => {
  const view = SheetView.open(book());
  view.select(1, 1);
  view.insertChart({ kind, title: 'Takings' });
  return view;
};
const chartOf = (view) => view.render().objects.find((o) => o.kind === 'chart');
const chartSvg = (view) => view.render().drawings.find((d) => d.svg)?.svg ?? '';
const partOf = (view) => Workbook.open(view.save()).pkg.text('xl/charts/chart1.xml');

test('Chart Design tells the window what the chart shows', () => {
  const view = charted();
  assert.deepEqual(chartOf(view).chart, {
    type: 'column', axes: true, categoryAxisTitle: '', valueAxisTitle: '', dataLabels: true, legend: 'bottom',
    series: [{ name: 'Shop', colour: null }, { name: 'Online', colour: null }],
  });
  assert.equal(chartOf(charted('pie')).chart.axes, false);
});

test('each choice is drawn at once, saved, read back, and undone in one step', () => {
  const view = charted();
  const { id } = chartOf(view);
  const before = chartSvg(view);
  const first = partOf(view);

  view.setChartAxisTitle({ id, axis: 'value', title: 'Pounds' });
  view.setChartAxisTitle({ id, axis: 'category', title: 'Month of 2026' });
  view.setChartLegend({ id, place: 'right' });
  view.setChartDataLabels({ id, on: false });
  view.setChartSeriesColour({ id, series: 1, colour: '#C00000' });

  const facts = chartOf(view).chart;
  assert.equal(facts.valueAxisTitle, 'Pounds');
  assert.equal(facts.categoryAxisTitle, 'Month of 2026');
  assert.equal(facts.legend, 'right');
  assert.equal(facts.dataLabels, false);
  assert.deepEqual(facts.series, [{ name: 'Shop', colour: null }, { name: 'Online', colour: '#C00000' }]);
  const svg = chartSvg(view);
  assert.match(svg, />Pounds</);
  assert.match(svg, />Month of 2026</);
  assert.match(svg, /fill="#C00000"/);
  assert.doesNotMatch(svg, />120</, 'no value on a column');

  // Saved and opened again, in this Office and as Excel orders a chart part.
  const again = SheetView.open(view.save());
  assert.deepEqual(chartOf(again).chart, facts);
  assert.equal(chartSvg(again), svg);
  const xml = partOf(view);
  assert.match(xml, /<\/c:plotArea><c:legend><c:legendPos val="r"\/>/);
  assert.match(xml, /<c:dispBlanksAs val="gap"\/><\/c:chart>/);

  // A changed figure still redraws a formatted chart, with its look kept.
  view.setCell(1, 2, 999);
  assert.match(chartSvg(view), /fill="#C00000"/);
  assert.match(chartSvg(view), />Pounds</);
  view.undo();

  // Five choices, five steps back to the chart as it was made.
  for (let i = 0; i < 5; i++) view.undo();
  assert.equal(chartSvg(view), before);
  assert.equal(partOf(view), first);
  view.redo();
  assert.equal(chartOf(view).chart.valueAxisTitle, 'Pounds');
});

test('what cannot be done says so', () => {
  const view = charted();
  const { id } = chartOf(view);
  assert.throws(() => view.setChartLegend({ id: 'nope', place: 'top' }));
  assert.throws(() => view.setChartLegend({ id, place: 'middle' }), /right, the top, the left or the bottom/);
  assert.throws(() => view.setChartAxisTitle({ id, axis: 'depth', title: 'x' }), /category axis or the value axis/);
  assert.throws(() => view.setChartSeriesColour({ id, series: 9, colour: '#000000' }), /no series 10/);
  const pie = charted('pie');
  assert.throws(() => pie.setChartAxisTitle({ id: chartOf(pie).id, axis: 'value', title: 'x' }), /no axes/);
  // Asking for what the chart already has is no step at all.
  const steps = view.history.describe();
  view.setChartAxisTitle({ id, axis: 'value', title: '' });
  assert.deepEqual(view.history.describe(), steps);
});
