/**
 * Data → Consolidate: ranges on several sheets summarised at the active
 * cell — by position, or by the labels in their top row and left column —
 * with Excel's eleven functions; with "Create links to source data" the
 * result is formulas over hidden detail rows, one per source, grouped under
 * an outline as Excel lays it out. What was consolidated is kept in the
 * sheet's `<dataConsolidate>`, as Excel keeps it.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

const book = () => new SheetView(buildXlsx({ sheets: [
  { name: 'Summary', rows: [['']] },
  { name: 'East', rows: [['Item', 'Q1', 'Q2'], ['Pens', 10, 20], ['Ink', 5, 7]] },
  { name: 'West', rows: [['Item', 'Q2', 'Q3'], ['Ink', 3, 4], ['Paper', 8, 9], ['Pens', 1, 1]] },
  { name: 'North sales', rows: [[1, 2], [3, 4]] },
] }));
const cell = (view, sheet, r, c) => view.calc.getValue(sheet, r, c);
const grid = (view, sheet, rows, cols) => Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => cell(view, sheet, r, c)));

test('by position: the first cell of every range into the first cell of the result, as large as the largest', () => {
  const view = book();
  view.select(0, 0);
  const out = view.consolidate({ fn: 'sum', refs: ['East!$B$2:$C$3', "'North sales'!$A$1:$B$2", 'West!B2:C4'] });
  assert.equal(out.ref, 'A1:B3');
  assert.deepEqual(grid(view, 'Summary', 3, 2), [[10 + 1 + 3, 20 + 2 + 4], [5 + 3 + 8, 7 + 4 + 9], [1, 1]]);
  view.undo();
  assert.equal(cell(view, 'Summary', 0, 0), '', 'one undo step');
});

test('by labels: like labels meet whatever their place, in the order first met, the labels written round the result', () => {
  const view = book();
  view.select(1, 1);
  view.consolidate({ fn: 'sum', refs: ['East!$A$1:$C$3', 'West!$A$1:$C$4'], topRow: true, leftCol: true });
  assert.deepEqual(grid(view, 'Summary', 5, 5).map((r) => r.slice(1)), [
    ['', '', '', ''],
    ['', 'Q1', 'Q2', 'Q3'],
    ['Pens', 10, 21, 1],
    ['Ink', 5, 10, 4],
    ['Paper', '', 8, 9],
  ]);
});

test('each of Excel\'s functions', () => {
  const run = (fn) => {
    const view = book();
    view.select(0, 0);
    view.consolidate({ fn, refs: ['East!$B$2:$C$3', 'West!$B$2:$C$3'] });
    return grid(view, 'Summary', 1, 2)[0];
  };
  assert.deepEqual(run('average'), [(10 + 3) / 2, (20 + 4) / 2]);
  assert.deepEqual(run('count'), [2, 2]);
  assert.deepEqual(run('countNumbers'), [2, 2]);
  assert.deepEqual(run('max'), [10, 20]);
  assert.deepEqual(run('min'), [3, 4]);
  assert.deepEqual(run('product'), [30, 80]);
  const sd = run('stdDev');
  assert.ok(Math.abs(sd[0] - Math.sqrt(((10 - 6.5) ** 2 + (3 - 6.5) ** 2) / 1)) < 1e-12);
  assert.ok(Math.abs(run('stdDevp')[0] - 3.5) < 1e-12);
  assert.ok(Math.abs(run('var')[0] - 24.5) < 1e-12);
  assert.ok(Math.abs(run('varp')[0] - 12.25) < 1e-12);
  assert.throws(() => run('mode'), /no function/);
});

test('Create links to source data: formulas over hidden detail rows, one per source, under an outline — and the totals follow the sources', () => {
  const view = book();
  view.select(0, 0);
  view.consolidate({ fn: 'sum', refs: ['East!$A$1:$C$3', 'West!$A$1:$C$4'], topRow: true, leftCol: true, links: true });
  // Header, then Pens: East, West, total; Ink: East, West, total; Paper: West, total.
  assert.equal(view.calc.getInput('Summary', 1, 1), 'East', 'the detail row names its source');
  assert.equal(view.calc.getInput('Summary', 1, 2), '=East!$B$2');
  assert.equal(view.calc.getInput('Summary', 2, 3), '=West!$B$4', 'West\'s Pens row, its Q2 column');
  assert.equal(view.calc.getInput('Summary', 3, 0), 'Pens');
  assert.equal(view.calc.getInput('Summary', 3, 3), '=SUM(D2:D3)');
  assert.equal(cell(view, 'Summary', 3, 3), 21);
  assert.equal(cell(view, 'Summary', 8, 4), 9, 'Paper\'s Q3');
  assert.deepEqual([...view.geo.hiddenRows].sort((a, b) => a - b), [1, 2, 4, 5, 7], 'the detail rows hidden');
  const xml = OoxmlPackage.read(view.save()).text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<row r="2"[^>]*outlineLevel="1"[^>]*hidden="1"|<row r="2"[^>]*hidden="1"[^>]*outlineLevel="1"/);
  assert.match(xml, /<row r="4"[^>]*collapsed="1"/, 'the total row carries the fold');
  assert.match(xml, /outlineLevelRow="1"/);
  // The totals follow the sources.
  view.selectSheet('East');
  view.setCell(1, 2, 100);
  assert.equal(cell(view, 'Summary', 3, 3), 101);
  // Excel's refusal: no links into a source's own sheet.
  view.selectSheet('East');
  view.select(10, 0);
  assert.throws(() => view.consolidate({ refs: ['East!$A$1:$C$3'], links: true }), /same sheet/);
});

test('the dialog\'s memory is Excel\'s <dataConsolidate>; a result over its own source is refused', () => {
  const view = book();
  view.select(0, 0);
  view.consolidate({ fn: 'average', refs: ['East!$A$1:$C$3', "'North sales'!$A$1:$B$2"], topRow: true, leftCol: true });
  const xml = OoxmlPackage.read(view.save()).text('xl/worksheets/sheet1.xml');
  assert.match(xml, /<dataConsolidate function="average" leftLabels="1" topLabels="1"><dataRefs count="2"><dataRef ref="A1:C3" sheet="East"\/><dataRef ref="A1:B2" sheet="North sales"\/><\/dataRefs><\/dataConsolidate>/);
  const again = new SheetView(view.save());
  assert.deepEqual(again.consolidateInfo().last, {
    fn: 'average', topRow: true, leftCol: true, links: false, refs: ['East!$A$1:$C$3', "'North sales'!$A$1:$B$2"],
  });
  view.selectSheet('East');
  view.select(1, 1);
  assert.throws(() => view.consolidate({ refs: ['East!$A$1:$C$3'] }), /overlap destination/);
  assert.throws(() => view.consolidate({ refs: ['Nowhere!A1:B2'] }), /no sheet called/);
  assert.throws(() => view.consolidate({ refs: [] }), /at least one reference/);
});
