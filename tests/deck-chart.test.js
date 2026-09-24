// A chart on a slide — Insert → Chart.
//
// Written the way PowerPoint writes one it has just drawn (a chart part, a
// relationship from the slide, a graphic frame naming the DrawingML chart
// namespace), minus the one thing only PowerPoint's own gallery adds: an
// embedded workbook. `chartPartXml` — the writer Word and Worksheets already
// use — bakes the values into the part as a cache instead, which the reader
// and the renderer both draw from directly.
import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const RICH = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'rich');

const DECK = buildPptx({
  title: 'Charts',
  slides: [
    { layout: 'title', title: 'Cover', body: 'A deck with a chart' },
    { layout: 'obj', title: 'Numbers', body: ['One'] },
  ],
});

const SAMPLE = {
  type: 'column',
  title: 'Sales',
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Sales', values: [12, 18, 15, 22] },
    { name: 'Costs', values: [8, 9, 10, 11] },
  ],
};

test('addChart writes the part, the content-type override, the relationship and the frame exactly', () => {
  const deck = Deck.open(DECK);
  const id = deck.addChart(1, SAMPLE);
  const pkg = OoxmlPackage.read(deck.save());

  assert.ok(pkg.has('ppt/charts/chart1.xml'), 'the chart part must exist');
  assert.match(pkg.text('[Content_Types].xml'),
    /PartName="\/ppt\/charts\/chart1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.drawingml\.chart\+xml"/,
    'an undeclared part is a corrupt package, not a missing feature');

  const slideRels = pkg.rels('ppt/slides/slide2.xml');
  const chartRel = slideRels.find((r) => r.Type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart');
  assert.ok(chartRel, 'the slide must point at its chart');
  assert.equal(OoxmlPackage.resolveTarget('ppt/slides/slide2.xml', chartRel.Target), 'ppt/charts/chart1.xml');

  const xml = pkg.text('ppt/slides/slide2.xml');
  assert.match(xml, new RegExp(`<p:cNvPr id="${id}" name="Chart ${id}"/>`));
  assert.match(xml, /<a:graphicData uri="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/chart">/);
  assert.match(xml, new RegExp(`<c:chart xmlns:c="http://schemas\\.openxmlformats\\.org/drawingml/2006/chart" xmlns:r="http://schemas\\.openxmlformats\\.org/officeDocument/2006/relationships" r:id="${chartRel.Id}"/>`));

  const chartXml = pkg.text('ppt/charts/chart1.xml');
  assert.match(chartXml, /<c:barChart><c:barDir val="col"\/>/, 'a column chart is a bar chart with barDir=col');
  assert.deepEqual([...chartXml.matchAll(/<c:cat><c:strRef>[\s\S]*?<\/c:strRef><\/c:cat>/g)].length > 0, true);
  assert.match(chartXml, /<c:strCache><c:ptCount val="4"\/>(?:<c:pt idx="\d"><c:v>Q[1-4]<\/c:v><\/c:pt>){4}<\/c:strCache>/);
  assert.match(chartXml, /<c:numCache>[\s\S]*<c:v>12<\/c:v>/, 'the first series\' values are cached');
  assert.match(chartXml, /<c:numCache>[\s\S]*<c:v>8<\/c:v>/, 'the second series\' values are cached');
});

test('the reader lists it as kind: chart, and chartData reads it back equal to what was given', () => {
  const deck = Deck.open(DECK);
  const id = deck.addChart(1, SAMPLE);
  const reopened = Deck.open(deck.save());
  const shape = reopened.slide(1).shapes.find((s) => s.kind === 'chart');
  assert.ok(shape, 'a chart shape');
  assert.equal(String(shape.id), String(id));

  const data = reopened.chartData(1, id);
  assert.equal(data.type, 'column');
  assert.equal(data.title, 'Sales');
  assert.deepEqual(data.categories, SAMPLE.categories);
  assert.deepEqual(data.series, SAMPLE.series);
});

test('renderSlide draws bars for a column chart', () => {
  const deck = Deck.open(DECK);
  deck.addChart(1, SAMPLE);
  const svg = renderSlide(deck.slide(1));
  const marks = /<g class="marks">([\s\S]*?)<\/g>/.exec(svg)?.[1] || '';
  const bars = marks.match(/<path/g) || [];
  assert.ok(bars.length >= 8, `two series of four categories should draw at least eight bars, got ${bars.length}`);
});

test('setChartData changes the values and the type, the frame kept byte-identical', () => {
  const deck = Deck.open(DECK);
  const id = deck.addChart(1, SAMPLE);
  const before = deck.pkg.text(deck.slideParts[1].part);

  deck.setChartData(1, id, { categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'Sales', values: [30, 18, 15, 22] }, { name: 'Costs', values: [8, 9, 10, 11] }] });
  const afterValues = deck.pkg.text(deck.slideParts[1].part);
  assert.equal(afterValues, before, 'the frame is not touched by a data edit');
  assert.equal(deck.chartData(1, id).series[0].values[0], 30);

  deck.setChartData(1, id, { type: 'line', categories: ['A', 'B'], series: [{ name: 'X', values: [1, 2] }] });
  assert.equal(deck.pkg.text(deck.slideParts[1].part), before, 'the frame is not touched by a type change either');
  const line = deck.chartData(1, id);
  assert.equal(line.type, 'line');
  assert.deepEqual(line.categories, ['A', 'B']);

  deck.setChartData(1, id, { type: 'pie', categories: ['A', 'B'], series: [{ name: 'X', values: [1, 2] }] });
  assert.equal(deck.chartData(1, id).type, 'pie');

  const reopened = Deck.open(deck.save());
  assert.equal(reopened.chartData(1, id).type, 'pie');
});

test('adding and editing a chart leaves the other slide byte-identical', () => {
  const before = OoxmlPackage.read(DECK);
  const deck = Deck.open(DECK);
  const id = deck.addChart(1, SAMPLE);
  deck.setChartData(1, id, { categories: ['A'], series: [{ name: 'X', values: [1] }] });
  const after = OoxmlPackage.read(deck.save());
  assert.deepEqual(after.read('ppt/slides/slide1.xml'), before.read('ppt/slides/slide1.xml'), 'slide 1 was not touched');
  assert.notDeepEqual(after.read('ppt/slides/slide2.xml'), before.read('ppt/slides/slide2.xml'));
});

test('a chart survives Deck.open(deck.save())', () => {
  const deck = Deck.open(DECK);
  const id = deck.addChart(1, SAMPLE);
  const reopened = Deck.open(deck.save());
  assert.deepEqual(reopened.chartData(1, id), reopened.chartData(1, id));
  assert.deepEqual(reopened.chartData(1, id).series, SAMPLE.series);
});

test('two charts on the same slide get distinct part numbers and ids', () => {
  const deck = Deck.open(DECK);
  const id1 = deck.addChart(1, SAMPLE);
  const id2 = deck.addChart(1, SAMPLE);
  assert.notEqual(id1, id2);
  const parts = deck.pkg.partNames().filter((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p));
  assert.deepEqual(parts.sort(), ['ppt/charts/chart1.xml', 'ppt/charts/chart2.xml']);
});

test('a bad slide or shape throws', () => {
  const deck = Deck.open(DECK);
  const id = deck.addChart(1, SAMPLE);
  assert.throws(() => deck.addChart(9, SAMPLE), /no slide/);
  assert.throws(() => deck.chartData(9, id), /no slide/);
  assert.throws(() => deck.setChartData(9, id, SAMPLE), /no slide/);
  assert.throws(() => deck.chartData(1, 999999), /not found/);
  assert.throws(() => deck.setChartData(1, 999999, SAMPLE), /not found/);
  assert.throws(() => deck.addChart(1, { ...SAMPLE, type: 'scatter' }), /unknown chart type/);
  assert.throws(() => deck.addChart(1, { ...SAMPLE, series: [] }), /at least one series/);
});

const showcaseFile = join(RICH, 'showcase.pptx');
test('the showcase deck\'s chart — written by PowerPoint itself — still reads as before', { skip: !existsSync(showcaseFile) && 'fixture not generated' }, () => {
  const deck = Deck.open(readFileSync(showcaseFile));
  const shapes = deck.slide(4).shapes.filter((s) => s.kind === 'chart');
  assert.equal(shapes.length, 2, 'two charts on the slide');
  const column = deck.chartData(4, shapes[0].id);
  assert.equal(column.type, 'column');
  assert.equal(column.title, 'Columns');
  assert.deepEqual(column.series[0].values, [4.3, 2.5, 3.5, 4.5]);
  const pie = deck.chartData(4, shapes[1].id);
  assert.equal(pie.type, 'pie');
  assert.equal(pie.title, 'A pie');
});
