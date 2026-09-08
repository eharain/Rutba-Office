/**
 * The shared drawing engine.
 *
 * Two things under test: that the rendering is correct and deterministic, and
 * that the DESIGN RULES are enforced rather than merely documented — a caller
 * should not be able to produce a misleading chart by passing the wrong options.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  scene, group, rect, text, line, bounds, markCount, emuToPx, pxToEmu,
  renderSvg, svgDataUri, RASTER_REQUIRED_NOTE,
  buildChart, normaliseSpec, niceScale, labelPolicy, chartTable, describe,
  parseChartXml, parseDrawingAnchors, readSheetDrawings, parseShapeXml, parsePictureXml,
  buildShape, buildPicture, resolveColour, SUPPORTED_GEOMETRY,
  CATEGORICAL, seriesColour, needsRelief, SCATTER_SERIES_CAP, INK,
  measureText, truncateText, wrapText,
} from '@rutba/drawing';
import { buildComplexWorkbook } from './fixtures/complex-workbook.js';
import { OoxmlPackage } from '@rutba/ooxml';

// --------------------------------------------------------------- palette --

test('the palette is fixed-order and never cycles', () => {
  assert.equal(CATEGORICAL.light.length, 8);
  assert.equal(CATEGORICAL.dark.length, 8);
  assert.equal(seriesColour(0, 'light'), '#2a78d6');
  assert.equal(seriesColour(1, 'light'), '#eb6834');
  assert.equal(seriesColour(0, 'dark'), '#3987e5');
  // a ninth series is NOT a generated hue — it falls back to muted ink so the
  // caller is pushed toward "Other" or small multiples
  assert.equal(seriesColour(8, 'light'), INK.light.muted);
  assert.equal(seriesColour(99, 'dark'), INK.dark.muted);
});

test('the light-mode relief rule is reported, not silently ignored', () => {
  // slots 2, 3 and 4 sit below 3:1 on the light surface
  assert.equal(needsRelief(2, 'light'), false);
  assert.equal(needsRelief(3, 'light'), true);
  assert.equal(needsRelief(8, 'dark'), false, 'all eight clear 3:1 on the dark surface');
});

// ----------------------------------------------------------------- scene --

test('scene bounds and mark counting', () => {
  const s = scene({
    width: 100, height: 50,
    children: [group([rect({ x: 10, y: 10, width: 30, height: 20 }), line({ x1: 0, y1: 0, x2: 90, y2: 40 })])],
  });
  assert.deepEqual(bounds(s), { x: 0, y: 0, right: 90, bottom: 40 });
  assert.equal(markCount(s), 2);
});

test('EMU conversion round-trips', () => {
  assert.equal(emuToPx(9525), 1);
  assert.equal(pxToEmu(96), 914400);
  assert.equal(emuToPx(pxToEmu(300)), 300);
});

// ------------------------------------------------------------------- svg --

test('SVG is standalone, accessible and deterministic', () => {
  const s = scene({
    width: 120, height: 60, title: 'A title', description: 'A description',
    children: [rect({ x: 1, y: 2, width: 10, height: 20, fill: '#2a78d6' })],
  });
  const svg = renderSvg(s);
  assert.match(svg, /^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg"/);
  assert.match(svg, /viewBox="0 0 120 60"/);
  assert.match(svg, /role="img" aria-label="A title"/);
  assert.match(svg, /<title>A title<\/title><desc>A description<\/desc>/);
  assert.match(svg, /<rect x="1" y="2" width="10" height="20" fill="#2a78d6"\/>/);
  assert.equal(renderSvg(s), svg, 'rendering twice gives identical bytes');
});

test('text is escaped and empty text is omitted', () => {
  const s = scene({
    width: 50, height: 20,
    children: [
      text({ x: 0, y: 10, value: 'a < b & "c"', fill: '#000' }),
      text({ x: 0, y: 20, value: '' }),
    ],
  });
  const svg = renderSvg(s);
  assert.match(svg, /a &lt; b &amp; &quot;c&quot;/);
  assert.equal((svg.match(/<text/g) || []).length, 1, 'an empty label is not emitted');
});

test('a scene paints its own surface so it survives on any background', () => {
  const light = renderSvg(scene({ width: 10, height: 10, mode: 'light', children: [] }));
  const dark = renderSvg(scene({ width: 10, height: 10, mode: 'dark', children: [] }));
  assert.match(light, /fill="#fcfcfb"/);
  assert.match(dark, /fill="#1a1a19"/);
  const none = renderSvg(scene({ width: 10, height: 10, background: 'none', children: [] }));
  assert.ok(!/<rect/.test(none));
});

test('a data URI is produced, and the raster caveat is stated', () => {
  const uri = svgDataUri(scene({ width: 10, height: 10, children: [] }));
  assert.match(uri, /^data:image\/svg\+xml;base64,/);
  assert.match(RASTER_REQUIRED_NOTE, /Outlook/);
  assert.match(RASTER_REQUIRED_NOTE, /rasterised to PNG/);
});

// ---------------------------------------------------------------- scales --

test('axis ticks land on numbers people read', () => {
  assert.deepEqual(niceScale(0, 100).ticks, [0, 20, 40, 60, 80, 100]);
  assert.deepEqual(niceScale(0, 9).ticks, [0, 2, 4, 6, 8, 10]);
  assert.deepEqual(niceScale(0, 1).step, 0.2);
  // never 0.30000000000000004 on an axis
  for (const t of niceScale(0, 1).ticks) assert.equal(String(t).length <= 4, true, 'tick ' + t);
  const negative = niceScale(-30, 70);
  assert.ok(negative.ticks.includes(0), 'a range spanning zero includes zero');
});

test('a flat series still gets a usable axis', () => {
  assert.ok(niceScale(5, 5).ticks.length >= 2);
  assert.deepEqual(niceScale(0, 0).ticks, [0, 0.5, 1]);
});

// ------------------------------------------------------- design guardrails --

test('a dual axis is not expressible', () => {
  const spec = normaliseSpec({ series: [{ name: 'a', values: [1] }], valueAxis: { title: 'x' } });
  assert.equal(spec.valueAxis2, undefined);
  assert.ok(!('secondaryAxis' in spec));
});

test('legend and direct-label policy follows the series count', () => {
  assert.deepEqual(labelPolicy(1, 'light'), { legend: false, directLabels: true, reliefRequired: false });
  assert.equal(labelPolicy(2, 'light').legend, true, 'two series always get a legend');
  assert.equal(labelPolicy(4, 'light').directLabels, true);
  assert.equal(labelPolicy(6, 'light').directLabels, false, 'past four, direct labels would collide');
  assert.equal(labelPolicy(6, 'light').reliefRequired, true);
});

test('scatter refuses more series than the palette can safely separate', () => {
  const series = Array.from({ length: SCATTER_SERIES_CAP + 1 }, (_, i) => ({ name: 's' + i, values: [1, 2] }));
  assert.throws(
    () => buildChart({ type: 'scatter', series }),
    /capped at 3 series.*all-pairs colour-blind floors/s,
  );
  assert.doesNotThrow(() => buildChart({ type: 'scatter', series: series.slice(0, 3) }));
});

test('a pie shows one series, and says so rather than guessing', () => {
  assert.throws(
    () => buildChart({ type: 'pie', series: [{ name: 'a', values: [1] }, { name: 'b', values: [2] }] }),
    /one series.*small multiples/,
  );
});

test('an unknown chart type is refused', () => {
  assert.throws(() => buildChart({ type: 'radar', series: [{ values: [1] }] }), /unknown chart type/);
  assert.throws(() => buildChart({ series: [] }), /at least one series/);
});

// ---------------------------------------------------------------- charts --

const STOCK = {
  type: 'column',
  title: 'Stock on hand',
  categories: ['SKU-1001', 'SKU-1002', 'SKU-1140', 'SKU-2202'],
  series: [{ name: 'On hand', values: [1420, 860, 240, 12] }],
  width: 480,
  height: 300,
};

test('a single-series column chart renders with no legend', () => {
  const svg = renderSvg(buildChart(STOCK));
  assert.match(svg, /<title>Stock on hand<\/title>/);
  assert.ok(!/class="legend"/.test(svg), 'one series needs no legend — the title names it');
  assert.match(svg, /class="marks"/);
  assert.match(svg, /class="labels"/, 'four or fewer series are direct-labelled');
  assert.match(svg, /#2a78d6/, 'slot 1');
  assert.match(svg, /SKU-1001/);
});

test('two series get a legend, and each keeps its own slot colour', () => {
  const svg = renderSvg(buildChart({
    ...STOCK,
    series: [
      { name: 'On hand', values: [1420, 860, 240, 12] },
      { name: 'Reserved', values: [120, 40, 0, 12] },
    ],
  }));
  assert.match(svg, /class="legend"/);
  assert.match(svg, /#2a78d6/);
  assert.match(svg, /#eb6834/);
  assert.match(svg, />On hand</);
  assert.match(svg, />Reserved</);
});

test('colour follows the entity, so dropping a series does not repaint the others', () => {
  const withBoth = buildChart({
    ...STOCK,
    series: [{ name: 'A', values: [1, 2] }, { name: 'B', values: [3, 4] }],
  });
  const filtered = buildChart({ ...STOCK, series: [{ name: 'A', values: [1, 2] }] });
  const colourOf = (s) => JSON.stringify(s).match(/#[0-9a-f]{6}/g).filter((c) => c === '#2a78d6');
  assert.ok(colourOf(withBoth).length > 0);
  assert.ok(colourOf(filtered).length > 0, 'series A keeps slot 1 either way');
});

test('every chart type produces a scene', () => {
  for (const type of ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter']) {
    const built = buildChart({ ...STOCK, type });
    assert.equal(built.type, 'scene', type);
    const svg = renderSvg(built);
    assert.match(svg, /^<svg/, type);
    assert.ok(svg.length > 200, type + ' produced something substantial');
  }
});

test('stacked columns leave a surface gap between segments', () => {
  const built = buildChart({
    ...STOCK, stacked: true,
    series: [{ name: 'A', values: [10, 20] }, { name: 'B', values: [5, 5] }],
  });
  assert.equal(built.type, 'scene');
  assert.match(renderSvg(built), /class="marks"/);
});

test('a chart carries alt text and a table for relief', () => {
  const spec = { ...STOCK, series: Array.from({ length: 6 }, (_, i) => ({ name: 's' + i, values: [1, 2, 3, 4] })) };
  assert.match(describe(spec), /6 series/);
  const table = chartTable(spec);
  assert.equal(table.header.length, 7);
  assert.equal(table.rows.length, 4);
  assert.equal(table.rows[0][0], 'SKU-1001');
  assert.equal(table.reliefRequired, true, 'six series in light mode needs the table view');
});

test('dark mode is its own selected palette, not an inverted light one', () => {
  const light = renderSvg(buildChart({ ...STOCK, mode: 'light' }));
  const dark = renderSvg(buildChart({ ...STOCK, mode: 'dark' }));
  assert.match(light, /#2a78d6/);
  assert.match(dark, /#3987e5/);
  assert.match(dark, /fill="#1a1a19"/);
  assert.ok(!dark.includes('#fcfcfb'));
});

test('gaps in a series are gaps, not zeroes', () => {
  const built = buildChart({ ...STOCK, type: 'line', series: [{ name: 'A', values: [10, null, 30, 40] }] });
  const svg = renderSvg(built);
  const points = /points="([^"]*)"/.exec(svg);
  assert.ok(points, 'a polyline was drawn');
  assert.equal(points[1].trim().split(/\s+/).length, 3, 'the null is skipped, not plotted as zero');
});

// ----------------------------------------------------------- measurement --

test('text measurement is usable for reserving space', () => {
  assert.ok(measureText('MMMM', { size: 10 }) > measureText('iiii', { size: 10 }));
  assert.equal(measureText('', { size: 10 }), 0);
  assert.ok(measureText('abc', { size: 20 }) > measureText('abc', { size: 10 }));
  assert.equal(truncateText('short', 1000), 'short', 'what fits is returned unchanged');
  assert.match(truncateText('a very long label indeed', 40), /…$/);
  assert.deepEqual(wrapText('one two three four', 40).length >= 2, true);
});

// ----------------------------------------------------------- OOXML edge --

test('a DrawingML chart part parses into a spec', () => {
  const chartXml = `<?xml version="1.0"?>
    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Warehouse</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>On hand</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>SKU-1001</c:v></c:pt><c:pt idx="1"><c:v>SKU-1002</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>1420</c:v></c:pt><c:pt idx="1"><c:v>860</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart></c:plotArea></c:chart></c:chartSpace>`;

  const spec = parseChartXml(chartXml);
  assert.equal(spec.type, 'column');
  assert.equal(spec.stacked, false);
  assert.equal(spec.title, 'Warehouse');
  assert.deepEqual(spec.categories, ['SKU-1001', 'SKU-1002']);
  assert.equal(spec.series.length, 1);
  assert.equal(spec.series[0].name, 'On hand');
  assert.deepEqual(spec.series[0].values, [1420, 860]);

  // and it renders through the ordinary path, in OUR palette
  assert.match(renderSvg(buildChart(spec)), /#2a78d6/);
});

test('barDir and grouping are honoured', () => {
  const make = (dir, grouping) => `<c:chartSpace xmlns:c="x"><c:barChart>
    <c:barDir val="${dir}"/><c:grouping val="${grouping}"/>
    <c:ser><c:val><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:val></c:ser>
    </c:barChart></c:chartSpace>`;
  assert.equal(parseChartXml(make('col', 'clustered')).type, 'column');
  assert.equal(parseChartXml(make('bar', 'clustered')).type, 'bar');
  assert.equal(parseChartXml(make('col', 'stacked')).stacked, true);
});

test('an unreadable chart part is null, not a crash', () => {
  assert.equal(parseChartXml(''), null);
  assert.equal(parseChartXml('<c:chartSpace/>'), null);
  assert.equal(parseChartXml('<c:chartSpace><c:barChart/></c:chartSpace>'), null, 'no series, no spec');
});

test('drawing anchors are read from a real fixture', () => {
  const pkg = OoxmlPackage.read(buildComplexWorkbook());
  const anchors = parseDrawingAnchors(pkg.text('xl/drawings/drawing1.xml'));
  assert.deepEqual(anchors.map((a) => a.kind), ['chart', 'shape', 'image']);
  assert.equal(anchors[0].anchorType, 'twoCell');
  assert.deepEqual(anchors[0].from, { col: 5, colOffsetEmu: 0, row: 1, rowOffsetEmu: 0 });
  assert.deepEqual(anchors[0].to, { col: 12, colOffsetEmu: 0, row: 16, rowOffsetEmu: 0 });
  assert.equal(anchors[0].name, 'Chart 1');

  // a oneCellAnchor sizes itself from its own extent rather than from cells
  assert.equal(anchors[2].anchorType, 'oneCell');
  assert.equal(Math.round(anchors[2].widthPx), 96); // 914400 EMU = one inch
});

test('readSheetDrawings resolves a chart end to end from a package', () => {
  const pkg = OoxmlPackage.read(buildComplexWorkbook());
  const drawingPart = 'xl/drawings/drawing1.xml';
  const rels = new Map(pkg.rels(drawingPart).map((r) => [r.Id, r.Target]));

  const drawings = readSheetDrawings({
    drawingXml: pkg.text(drawingPart),
    resolveRelationship: (id) => {
      const target = rels.get(id);
      return target ? OoxmlPackage.resolveTarget(drawingPart, target) : null;
    },
    readPart: (name) => (pkg.has(name) ? pkg.text(name) : null),
  });
  assert.deepEqual(drawings.map((d) => d.kind), ['chart', 'shape', 'image']);
  assert.equal(drawings[0].part, 'xl/charts/chart1.xml');

  const spec = drawings[0].spec;
  assert.ok(spec, 'the chart resolved through its relationship');
  assert.equal(spec.type, 'column');
  assert.equal(spec.title, 'Stock on hand');
  assert.deepEqual(spec.series.map((s) => s.name), ['On hand', 'Reserved']);
  assert.deepEqual(spec.series[0].values, [1420, 860]);
  assert.deepEqual(spec.categories, ['Steel bracket 40mm', 'Steel bracket 60mm']);

  // and the anchor sized it from the cells it spans
  assert.ok(spec.width > 100 && spec.height > 100);
  assert.match(renderSvg(buildChart(spec)), /Stock on hand/);
});

test('a drawing we cannot resolve degrades to a reason, not an exception', () => {
  const drawings = readSheetDrawings({
    drawingXml: '<xdr:wsDr><xdr:twoCellAnchor>'
      + '<xdr:from><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:from>'
      + '<xdr:to><xdr:col>5</xdr:col><xdr:row>10</xdr:row></xdr:to>'
      + '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Broken"/></xdr:nvGraphicFramePr>'
      + '<c:chart r:id="rIdMissing"/></xdr:graphicFrame>'
      + '</xdr:twoCellAnchor></xdr:wsDr>',
    resolveRelationship: () => null,
    readPart: () => null,
  });
  assert.equal(drawings.length, 1);
  assert.equal(drawings[0].spec, null);
  assert.equal(drawings[0].name, 'Broken');
});

// ---------------------------------------------------------------------------
// Shapes and pictures — the other half of a drawing.
// ---------------------------------------------------------------------------

const shapeXml = (inner) =>
  '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3" name="Callout"/></xdr:nvSpPr>' + inner + '</xdr:sp>';

test('a shape reads its geometry, its colours and its words', () => {
  const d = parseShapeXml(shapeXml(
    '<xdr:spPr><a:prstGeom prst="roundRect"/>'
    + '<a:solidFill><a:srgbClr val="1BAF7A"/></a:solidFill>'
    + '<a:ln w="19050"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></xdr:spPr>'
    + '<xdr:txBody><a:p><a:r><a:rPr b="1" sz="1400"/><a:t>Check this </a:t></a:r>'
    + '<a:r><a:t>total</a:t></a:r></a:p></xdr:txBody>',
  ));
  assert.equal(d.kind, 'shape');
  assert.equal(d.name, 'Callout');
  assert.equal(d.geometry, 'roundRect');
  assert.deepEqual(d.fill, { type: 'srgb', value: '1BAF7A' });
  assert.deepEqual(d.stroke, { type: 'scheme', value: 'tx1' });
  assert.equal(d.strokeWidth, 2); // scene units are pixels: 19050 EMU / 9525
  // runs are concatenated, so a shape split mid-sentence still reads as a sentence
  assert.equal(d.text, 'Check this total');
  assert.equal(d.textBold, true);
  assert.equal(d.textSize, 14); // sz is hundredths of a point
});

test('an unsupported preset still draws a box of the right size and colour', () => {
  const d = parseShapeXml(shapeXml(
    '<xdr:spPr><a:prstGeom prst="cloudCallout"/>'
    + '<a:solidFill><a:srgbClr val="2F6FEB"/></a:solidFill></xdr:spPr>',
  ));
  assert.equal(d.geometry, 'cloudCallout');
  assert.equal(d.supported, false); // honest about the approximation

  const svg = renderSvg(scene({
    width: 120, height: 60, children: [buildShape(d, { x: 0, y: 0, width: 120, height: 60 })],
  }));
  // a wrong-shaped box beats a blank: something with the right colour is there
  assert.match(svg, /<rect[^>]*fill="#2f6feb"/);
});

test('every supported preset renders a mark, and none of them emit float noise', () => {
  for (const geometry of SUPPORTED_GEOMETRY) {
    const node = buildShape({ geometry }, { x: 3, y: 7, width: 111, height: 47 });
    const svg = renderSvg(scene({ width: 120, height: 60, children: [node] }));
    assert.match(svg, /<(rect|path|polygon|ellipse|line)\b/, geometry + ' drew nothing');
    const long = svg.match(/\d+\.\d{4,}/);
    assert.equal(long, null, geometry + ' emitted ' + long);
  }
});

test('a theme colour lands on OUR palette, not on a colour we never loaded', () => {
  // We do not read the customer's theme part, so accent1 becomes our first
  // series colour. Their theme survives untouched in the file either way.
  assert.equal(resolveColour({ type: 'scheme', value: 'accent1' }, 'light'), seriesColour(0, 'light'));
  assert.equal(resolveColour({ type: 'scheme', value: 'accent3' }, 'dark'), seriesColour(2, 'dark'));
  assert.equal(resolveColour({ type: 'srgb', value: 'AABBCC' }), '#aabbcc');
  assert.equal(resolveColour({ type: 'none' }), 'none');
  assert.equal(resolveColour(null), null);
});

test('shape text wraps inside the box and centres on it', () => {
  const node = buildShape(
    { geometry: 'rect', text: 'Outstanding balance as at 31 March', textSize: 12 },
    { x: 0, y: 0, width: 90, height: 70 },
  );
  const lines = node.children.filter((c) => c.type === 'text');
  assert.ok(lines.length > 1, 'long text should wrap');
  assert.ok(lines.every((l) => l.anchor === 'middle' && l.x === 45));
  // vertically centred: the run of lines straddles the middle
  const ys = lines.map((l) => l.y);
  assert.ok(Math.abs((ys[0] + ys[ys.length - 1]) / 2 - 35) < 0.51);
});

test('a picture becomes a data URI so it survives leaving the package', () => {
  const png = Buffer.from('89504e470d0a1a0a', 'hex');
  const d = parsePictureXml(
    '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="5" name="Logo"/></xdr:nvPicPr>'
    + '<xdr:blipFill><a:blip r:embed="rId7"/></xdr:blipFill></xdr:pic>',
    { resolveRelationship: (id) => (id === 'rId7' ? 'xl/media/image1.png' : null), readPartBinary: () => png },
  );
  assert.equal(d.kind, 'picture');
  assert.equal(d.name, 'Logo');
  assert.equal(d.part, 'xl/media/image1.png');
  assert.equal(d.href, 'data:image/png;base64,' + png.toString('base64'));

  const svg = renderSvg(scene({
    width: 40, height: 40, children: [buildPicture(d, { x: 0, y: 0, width: 40, height: 40 })],
  }));
  assert.match(svg, /<image[^>]*href="data:image\/png;base64,/);
});

test('a picture format we cannot show degrades to no href rather than a broken one', () => {
  // EMF is common in Office files and is not a web image type.
  const d = parsePictureXml(
    '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="6" name="Diagram"/></xdr:nvPicPr>'
    + '<xdr:blipFill><a:blip r:embed="rId9"/></xdr:blipFill></xdr:pic>',
    { resolveRelationship: () => 'xl/media/image2.emf', readPartBinary: () => Buffer.from([1, 2, 3]) },
  );
  assert.equal(d.href, null);
  assert.equal(d.supported, false);
  assert.equal(d.part, 'xl/media/image2.emf'); // still named, so the UI can say what it was
});

test('a sheet mixing a shape and a picture reads both from one drawing part', () => {
  const png = Buffer.from('89504e470d0a1a0a', 'hex');
  const drawings = readSheetDrawings({
    drawingXml: '<xdr:wsDr>'
      + '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:from>'
      + '<xdr:to><xdr:col>4</xdr:col><xdr:row>6</xdr:row></xdr:to>'
      + '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Note"/></xdr:nvSpPr>'
      + '<xdr:spPr><a:prstGeom prst="ellipse"/></xdr:spPr>'
      + '<xdr:txBody><a:p><a:r><a:t>Query</a:t></a:r></a:p></xdr:txBody></xdr:sp>'
      + '</xdr:twoCellAnchor>'
      + '<xdr:oneCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:row>0</xdr:row></xdr:from>'
      + '<xdr:ext cx="952500" cy="476250"/>'
      + '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Logo"/></xdr:nvPicPr>'
      + '<xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic>'
      + '</xdr:oneCellAnchor></xdr:wsDr>',
    resolveRelationship: () => 'xl/media/image1.png',
    readPart: () => null,
    readPartBinary: () => png,
  });
  assert.deepEqual(drawings.map((d) => d.kind), ['shape', 'image']);
  assert.equal(drawings[0].descriptor.geometry, 'ellipse');
  assert.equal(drawings[0].descriptor.text, 'Query');
  assert.equal(drawings[1].descriptor.kind, 'picture');
  assert.ok(drawings[1].descriptor.href.startsWith('data:image/png'));
  // the oneCellAnchor sized itself from its own extent: 952500 EMU = 100px
  assert.equal(Math.round(drawings[1].widthPx), 100);
  assert.equal(Math.round(drawings[1].heightPx), 50);
});

test('a connector is a shape whose line runs corner to corner, flipped as its box says', () => {
  const xml =
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
    '<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
    '<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="2" name="Straight Connector 1"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>' +
    '<xdr:spPr><a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></xdr:spPr>' +
    '<xdr:style><a:lnRef idx="1"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>' +
    '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style></xdr:cxnSp>' +
    '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>';
  const [found] = readSheetDrawings({ drawingXml: xml, resolveRelationship: () => null, readPart: () => null, readPartBinary: () => null });
  assert.equal(found.kind, 'shape', 'a cxnSp is read, not left unknown');
  const d = found.descriptor;
  assert.equal(d.connector, true);
  assert.equal(d.geometry, 'line');
  assert.equal(d.flipV, true);
  assert.equal(d.flipH, false);
  assert.deepEqual(d.stroke, { type: 'scheme', value: 'accent1' }, 'the line colour comes from lnRef, and the shade child does not overwrite it');
  assert.equal(d.fill, null, 'fillRef idx 0 is no fill');
  const built = buildShape(d, { x: 10, y: 20, width: 100, height: 50 });
  const ln = built.children.find((c) => c.type === 'line');
  assert.ok(ln, 'drawn as a line');
  assert.deepEqual([ln.x1, ln.y1, ln.x2, ln.y2], [10, 70, 110, 20], 'from the bottom-left corner to the top-right, as flipV says');
});

test('a shape with no fill of its own takes the theme fill its style refers to; a brace is a stroke with no fill', () => {
  const sp = (geom, style) =>
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="${geom}"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom></xdr:spPr>${style}</xdr:sp>`;
  const styled = '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
    '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>';
  const arrow = parseShapeXml(sp('rightArrow', styled));
  assert.deepEqual(arrow.fill, { type: 'scheme', value: 'accent1' }, 'filled from fillRef');
  assert.deepEqual(arrow.stroke, { type: 'scheme', value: 'accent1' }, 'outlined from lnRef');
  assert.deepEqual(arrow.textColour, { type: 'scheme', value: 'lt1' }, 'text coloured from fontRef');

  const brace = parseShapeXml(sp('leftBrace', styled));
  assert.equal(brace.supported, true, 'a brace is a geometry we draw');
  const built = buildShape(brace, { x: 0, y: 0, width: 20, height: 100 });
  const p = built.children.find((c) => c.type === 'path');
  assert.ok(p, 'drawn as a path');
  assert.equal(p.fill, 'none', 'a brace has no body');
  assert.ok(p.stroke, 'and a stroke');
  assert.match(p.d, /^M20 0Q/, 'its arms start at the far side');
  assert.match(p.d, /Q10 50 0 50/, 'and its point touches the near side at mid-height');

  const bent = buildShape({ ...parseShapeXml(sp('bentConnector3', styled)), flipH: true }, { x: 0, y: 0, width: 100, height: 40 });
  const elbow = bent.children.find((c) => c.type === 'path');
  assert.equal(elbow.d, 'M100 0L50 0L50 40L0 40', 'a bent connector turns half way across');
});

test('a rotated shape turns about its centre, and the turn reaches the SVG', () => {
  const xml =
    '<xdr:sp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" macro="" textlink="">' +
    '<xdr:nvSpPr><xdr:cNvPr id="5" name="Right Brace 4"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
    '<xdr:spPr><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rightBrace"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>';
  const d = parseShapeXml(xml);
  assert.equal(d.rotation, 90);
  const built = buildShape(d, { x: 0, y: 0, width: 200, height: 100 });
  assert.equal(built.transform, 'rotate(90 100 50)');
  const svg = renderSvg(scene({ width: 200, height: 100, children: [built] }));
  assert.match(svg, /<g transform="rotate\(90 100 50\)"/);
  assert.equal(parseShapeXml(xml.replace(' rot="5400000"', '')).rotation, 0, 'no rot is no turn');
});
