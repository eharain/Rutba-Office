// OpenDocument charts, shapes and gradients, drawn as they were made.
//
// showcase.ods and showcase.odp are showcase.xlsx and showcase.pptx as Excel
// and PowerPoint themselves saved them in ODF (tools/make-rich-fixtures.ps1):
// the workbook's five charts are embedded chart objects plotting its
// cells, its shapes sit in cells, and the deck's shapes and first slide are
// filled with gradients. An .ods now keeps its charts — plotting the same
// cells, at the same place and size — and its shapes, picture and text box;
// an .odp keeps its gradients, its backgrounds and its charts.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { readZip, writeZip } from '@rutba/ooxml/zip';
import { readOdf, odfRange, rangeRef, angleDeg } from '../packages/office-formats/src/odf.js';
import { writeOdp } from '../packages/office-formats/src/odf-write.js';
import { Deck, renderSlide } from '@rutba/presentation';
import { createDocumentService } from '../apps/desktop/main/documents.js';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const fixture = (name) => fs.readFileSync(path.join(HERE, 'fixtures', 'rich', name));

/** A file opened through the suite, as the window opens one, and its converted package's parts. */
function converted(bytes, name) {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-odf-')), name);
  fs.writeFileSync(file, bytes);
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const s = docs.open({ path: file });
  const out = path.join(path.dirname(file), name.replace(/\.od.$/, name.endsWith('.ods') ? '.xlsx' : '.pptx'));
  docs.save({ id: s.id, path: out });
  const parts = new Map(readZip(fs.readFileSync(out)).entries.map((e) => [e.name, e.data.toString('utf8')]));
  return { docs, id: s.id, parts, out };
}

test('ODF ranges and angles read as the format writes them', () => {
  assert.deepEqual(odfRange('Sales.$B$2:.$B$13'), { sheet: 'Sales', top: 1, left: 1, bottom: 12, right: 1 });
  assert.deepEqual(odfRange("'My data'.A1"), { sheet: 'My data', top: 0, left: 0, bottom: 0, right: 0 });
  assert.equal(rangeRef(odfRange("'It''s'.$C$3:.$D$4")), "'It''s'!$C$3:$D$4");
  assert.equal(odfRange('nonsense'), null);
  assert.equal(angleDeg('3150'), 315, 'tenths of a degree, ODF 1.2');
  assert.equal(angleDeg('45deg'), 45);
});

test('an .ods reads its charts with the cells they plot, and its shapes, picture and text box in their cells', () => {
  const odf = readOdf(fixture('showcase.ods'));
  const summary = odf.sheets.find((s) => s.name === 'Summary');
  const pie = summary.drawings.find((d) => d.type === 'chart');
  assert.deepEqual(pie.anchor, { row: 1, col: 3 });
  assert.equal(pie.chart.kind, 'pie');
  assert.equal(pie.chart.title, 'Share by region');
  assert.deepEqual(pie.chart.categories.values, ['North', 'South', 'East', 'West']);
  assert.equal(pie.chart.series[0].ref, 'Summary!$A$20:$D$20');
  const charts = odf.sheets.find((s) => s.name === 'Charts').drawings;
  assert.deepEqual(charts.filter((d) => d.type === 'chart').map((d) => d.chart.kind), ['column', 'line', 'bar', 'area'], "Excel's own: columns, lines, horizontal bars, areas");
  const columns = charts.find((d) => d.type === 'chart').chart;
  assert.deepEqual(columns.series.map((s) => s.name), ['North', 'South', 'East', 'West']);
  assert.equal(columns.series[0].ref, 'Sales!$B$2:$B$13');
  assert.equal(columns.series[0].values[0], 1200);
  assert.equal(columns.categories.ref, 'Sales!$A$2:$A$13');
  const shapes = charts.filter((d) => d.type === 'shape');
  assert.ok(shapes.length >= 14, `${shapes.length} shapes`);
  assert.equal(shapes.find((s) => s.name === 'Shape Rectangle').fill, '#c00000');
  assert.equal(shapes.find((s) => s.name === 'Shape Rounded').fill.gradient.start, '#ED7D31');
  assert.ok(charts.some((d) => d.type === 'image' && /\.png$/.test(d.href)));
  assert.deepEqual(charts.find((d) => d.type === 'text').paragraphs.slice(0, 1), ['A text box beside the picture.']);
});

test('opened, an .ods keeps its charts plotting the same cells, at their place and size, and its shapes and picture', () => {
  const { docs, id, parts } = converted(fixture('showcase.ods'), 'showcase.ods');
  const chartParts = [...parts.keys()].filter((n) => /^xl\/charts\/chart\d+\.xml$/.test(n));
  assert.equal(chartParts.length, 5, 'the five charts');
  const all = chartParts.map((n) => parts.get(n)).join('');
  for (const ref of ['Summary!$A$20:$D$20', 'Sales!$B$2:$B$13', 'Sales!$F$2:$F$13', 'Sales!$A$2:$A$13']) assert.ok(all.includes(`<c:f>${ref}</c:f>`), ref);
  assert.ok(/<c:pieChart>/.test(all) && /<c:lineChart>/.test(all) && /<c:areaChart>/.test(all) && /<c:barDir val="bar"\/>/.test(all) && /<c:barDir val="col"\/>/.test(all));
  docs.apply({ id, ops: [{ op: 'sheet', name: 'Charts' }, { op: 'viewport', x: 0, y: 0, width: 4000, height: 4000 }] });
  const drawings = docs.model({ id }).drawings || [];
  const first = drawings.find((d) => d.name === 'Chart 1');
  assert.deepEqual([Math.round(first.width), Math.round(first.height)], [560, 320], 'the size it had');
  assert.deepEqual([Math.round(first.x), Math.round(first.y)], [13, 13], 'and its place');
  assert.ok(drawings.filter((d) => d.kind === 'shape').length >= 15, 'the shapes and the text box');
  assert.ok(drawings.some((d) => d.kind === 'image' || d.kind === 'picture'), 'the picture');
  assert.equal(Math.round(drawings.find((d) => d.name === 'Shape Rectangle').width), 160);
});

test('opened, an .odp keeps its gradients, its first slide\'s gradient background and its charts', () => {
  const { parts } = converted(fixture('showcase.odp'), 'showcase.odp');
  const slides = [...parts.keys()].filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => Number(/(\d+)\.xml/.exec(a)[1]) - Number(/(\d+)\.xml/.exec(b)[1]));
  const first = parts.get(slides[0]);
  assert.match(first, /<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="1F4E79"\/><\/a:gs><a:gs pos="100000"><a:srgbClr val="C00000"\/><\/a:gs><\/a:gsLst><a:lin ang="8100000"/, 'as the deck had it: 135°, navy to red');
  const all = slides.map((n) => parts.get(n)).join('');
  assert.match(all, /name="Shape Rounded"[\s\S]*?<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="ED7D31"\/><\/a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"\/><\/a:gs><\/a:gsLst><a:lin ang="5400000"/, 'top to bottom, orange to white');
  assert.ok(slides.some((n) => /<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFF2CC"\/>/.test(parts.get(n))), 'a plain background too');
  assert.equal([...parts.keys()].filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).length, 2, 'its two charts');
  const charts = [...parts.keys()].filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n)).map((n) => parts.get(n)).join('');
  assert.match(charts, /<c:pieChart>/);
  assert.match(charts, /Category 1/);
});

test('radial and axial gradients come through as DrawingML draws them, and a radial one is drawn round its centre', () => {
  const base = writeOdp({ title: 'Glow', size: { width: 1280, height: 720 }, slides: [{ title: 'Glow', body: [] }] });
  const { entries } = readZip(Buffer.from(base));
  for (const e of entries) {
    if (e.name === 'styles.xml') {
      e.data = Buffer.from(e.data.toString('utf8').replace(/<office:styles>|<office:styles\/>/, () => '<office:styles>'
        + '<draw:gradient draw:name="glow" draw:style="radial" draw:cx="30%" draw:cy="40%" draw:start-color="#000080" draw:end-color="#ffffff" draw:border="20%"/>'
        + '<draw:gradient draw:name="band" draw:style="axial" draw:angle="900" draw:start-color="#ff0000" draw:end-color="#ffff00"/>'
        + (/<office:styles\/>/.test(e.data.toString('utf8')) ? '</office:styles>' : '')), 'utf8');
    }
    if (e.name === 'content.xml') {
      e.data = Buffer.from(e.data.toString('utf8')
        .replace(/<office:automatic-styles\/>|<office:automatic-styles>/, () => '<office:automatic-styles><style:style style:name="g1" style:family="graphic"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="glow"/></style:style><style:style style:name="g2" style:family="graphic"><style:graphic-properties draw:fill="gradient" draw:fill-gradient-name="band"/></style:style>' + (/<office:automatic-styles\/>/.test(e.data.toString('utf8')) ? '</office:automatic-styles>' : ''))
        .replace(/(<draw:page\b[^>]*>)/, (m) => m + '<draw:ellipse draw:name="Sun" draw:style-name="g1" svg:x="2cm" svg:y="2cm" svg:width="6cm" svg:height="6cm"/><draw:rect draw:name="Band" draw:style-name="g2" svg:x="10cm" svg:y="2cm" svg:width="6cm" svg:height="3cm"/>'), 'utf8');
    }
  }
  const { parts, out } = converted(writeZip(entries), 'glow.odp');
  const slide = parts.get('ppt/slides/slide1.xml');
  assert.match(slide, /name="Sun"[\s\S]*?<a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"\/><\/a:gs><a:gs pos="80000"><a:srgbClr val="000080"\/><\/a:gs><a:gs pos="100000"><a:srgbClr val="000080"\/><\/a:gs><\/a:gsLst><a:path path="circle"><a:fillToRect l="30000" t="40000" r="70000" b="60000"\/><\/a:path>/, 'white at the centre, navy out to the border');
  assert.match(slide, /name="Band"[\s\S]*?<a:gs pos="0"><a:srgbClr val="FF0000"\/><\/a:gs><a:gs pos="50000"><a:srgbClr val="FFFF00"\/><\/a:gs><a:gs pos="100000"><a:srgbClr val="FF0000"\/><\/a:gs><\/a:gsLst><a:lin ang="0"/, 'red, yellow, red, turned a quarter');
  const svg = renderSlide(Deck.open(fs.readFileSync(out)).slide(0), { width: 960 });
  assert.match(svg, /<radialGradient\b[^>]*cx="0\.3000" cy="0\.4000"/);
});

test('saved back as .ods, a workbook keeps its charts plotting the same cells, its shapes in their outlines and its picture', () => {
  const { docs, id, out } = converted(fixture('showcase.ods'), 'showcase.ods');
  const back = out.replace(/\.xlsx$/, '-again.ods');
  docs.save({ id, path: back });
  const odf = readOdf(fs.readFileSync(back));
  const charts = odf.sheets.flatMap((s) => s.drawings.filter((d) => d.type === 'chart').map((d) => d.chart));
  assert.deepEqual(charts.map((c) => c.kind).sort(), ['area', 'bar', 'column', 'line', 'pie']);
  assert.ok(charts.some((c) => c.series[0].ref === 'Sales!$B$2:$B$13' && c.categories.ref === 'Sales!$A$2:$A$13'), 'its ranges, as ODF writes them and reads them back');
  const drawn = odf.sheets.find((s) => s.name === 'Charts').drawings;
  assert.ok(drawn.filter((d) => d.type === 'shape' && d.figures).length >= 14, 'the shapes as their own outlines');
  assert.equal(drawn.find((d) => d.name === 'Shape Smile').figures.length, 4, "the smile's eyes and mouth");
  assert.ok(drawn.some((d) => d.type === 'image'), 'the picture');
  assert.equal(drawn.find((d) => d.name === 'Rectangle 23').rotation, 30, 'turned as it was');
});

test('a chart with no cells behind it is written with a table of its own data, and a preset shape by its ODF name', async () => {
  const { writeOds } = await import('../packages/office-formats/src/odf-write.js');
  const bytes = writeOds({
    sheets: [{
      name: 'Data', rows: [[{ value: 1, text: '1' }]],
      drawings: [
        { kind: 'chart', name: 'Pie', x: 10, y: 10, w: 300, h: 200, chart: { kind: 'pie', title: 'Shares', categories: { values: ['A', 'B'] }, series: [{ name: 'Share', values: [3, 7] }] } },
        { kind: 'shape', name: 'Disc', geometry: 'ellipse', fill: '#FF0000', stroke: '#000000', x: 400, y: 10, w: 100, h: 100, text: 'Hi' },
      ],
    }],
  });
  const odf = readOdf(bytes);
  const [pie, disc] = odf.sheets[0].drawings;
  assert.equal(pie.chart.kind, 'pie');
  assert.deepEqual(pie.chart.categories.values, ['A', 'B']);
  assert.deepEqual(pie.chart.series[0].values, [3, 7]);
  assert.equal(pie.chart.series[0].name, 'Share');
  assert.equal(disc.geometry, 'ellipse');
  assert.equal(disc.fill, '#FF0000');
  assert.deepEqual(disc.paragraphs, ['Hi']);
});
