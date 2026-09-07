/**
 * The creator's chart — built from nothing, painted by the same reader that
 * paints a customer's.
 *
 * The build-golden suite pins the BYTES of a charted workbook; this file is
 * what makes a failed digest diagnosable, and it closes the loop the samples
 * depend on: a chart the creator writes must be one `readSheetDrawings` can
 * parse and `SheetView` can paint, because /try's budget sample ships one and
 * a guest's first impression must not be a "could not be read" placeholder.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx, OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

const charted = () => buildXlsx({
  sheets: [{
    name: 'Data',
    rows: [['Jan', 'Feb', 'Mar'], [120, 80, 95]],
    drawings: [{
      kind: 'chart', name: 'Chart 1', title: 'Widgets shipped',
      categories: { ref: 'Data!$A$1:$C$1', values: ['Jan', 'Feb', 'Mar'] },
      series: [{ name: 'Shipped', ref: 'Data!$A$2:$C$2', values: [120, 80, 95] }],
      from: { col: 0, row: 3 }, to: { col: 7, row: 18 },
    }],
  }],
});

test('a built chart is a declared part wired through real relationships', () => {
  const pkg = OoxmlPackage.read(charted());

  assert.ok(pkg.has('xl/charts/chart1.xml'), 'the chart part must exist');
  assert.match(pkg.text('[Content_Types].xml'),
    /PartName="\/xl\/charts\/chart1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.drawingml\.chart\+xml"/,
    'an undeclared part is a corrupt package, not a missing feature');

  // sheet -> drawing -> chart, each hop a real relationship
  const sheetRels = pkg.rels('xl/worksheets/sheet1.xml');
  const drawingRel = sheetRels.find((r) => String(r.Type).endsWith('/drawing'));
  assert.ok(drawingRel, 'the sheet must point at its drawing');
  const drawingRels = pkg.rels('xl/drawings/drawing1.xml');
  const chartRel = drawingRels.find((r) => String(r.Type).endsWith('/chart'));
  assert.ok(chartRel, 'the drawing must point at the chart');
  assert.equal(OoxmlPackage.resolveTarget('xl/drawings/drawing1.xml', chartRel.Target), 'xl/charts/chart1.xml');

  // The formulas that keep the chart live in Excel, and the caches that every
  // reader paints before recalculating.
  const chartXml = pkg.text('xl/charts/chart1.xml');
  assert.match(chartXml, /<c:f>Data!\$A\$2:\$C\$2<\/c:f>/, 'values must reference the sheet');
  assert.match(chartXml, /<c:f>Data!\$A\$1:\$C\$1<\/c:f>/, 'categories must reference the sheet');
  assert.match(chartXml, /<c:numCache>[\s\S]*<c:v>120<\/c:v>/, 'the value cache must be populated');
  // ECMA-376: a bar chart without its two axes is a file Excel repairs.
  assert.match(chartXml, /<c:catAx>[\s\S]*<c:valAx>/, 'both axes must be present');
});

test('the sheet view paints the built chart, not a placeholder', () => {
  const view = SheetView.open(charted(), { viewportWidth: 1600, viewportHeight: 1400 });
  const drawings = view.render().drawings;
  const chart = drawings.find((d) => d.kind === 'chart');
  assert.ok(chart, 'the chart anchor must be found');
  assert.ok(chart.svg && !chart.unsupported, chart.unsupported ?? 'no svg was produced');
  assert.match(chart.svg, /Widgets shipped/, 'the title must be painted');
  assert.match(chart.svg, /Feb/, 'category labels must be painted');
});

test('a workbook without charts writes no chart part', () => {
  // Invariant 2: only write what was set. The chartless goldens hold the bytes
  // still; this states the same rule where a reader will look for it.
  const pkg = OoxmlPackage.read(buildXlsx({ sheets: [{ name: 'Plain', rows: [[1]] }] }));
  assert.ok(!pkg.partNames().some((n) => n.startsWith('xl/charts/')));
  assert.ok(!pkg.text('[Content_Types].xml').includes('chart'));
});

test('an edit elsewhere leaves the chart part untouched on save', () => {
  const view = SheetView.open(charted(), { viewportWidth: 1600, viewportHeight: 1400 });
  view.setCell(1, 0, '999');
  const saved = OoxmlPackage.read(view.save());
  assert.deepEqual(saved.read('xl/charts/chart1.xml'), OoxmlPackage.read(charted()).read('xl/charts/chart1.xml'));
});
