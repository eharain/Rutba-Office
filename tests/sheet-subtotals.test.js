/**
 * Data → Subtotal and Remove All, in the engine — and SUBTOTAL leaving out
 * the subtotals inside its own range, which is what makes a Grand Total
 * over a subtotalled list count each row once.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const SALES = [
  ['Region', 'Rep', 'Units', 'Sales'],
  ['East', 'Kim', 3, 300],
  ['East', 'Lee', 2, 250],
  ['North', 'Ann', 5, 400],
  ['West', 'Bo', 1, 90],
  ['West', 'Cy', 4, 210],
  ['West', 'Di', 2, 120],
];
const open = (rows = SALES) => new SheetView(buildXlsx({ sheets: [{ name: 'S', rows }] }));
const text = (view, row, col) => view.displayValue(row, col).text;
const input = (view, row, col) => view.editValue(row, col);
const column = (view, col, from, to) => Array.from({ length: to - from + 1 }, (_, i) => text(view, from + i, col));
const partXml = (view) => {
  const part = view.workbook.partNameFor(view.activeSheet);
  return view.workbook.snapshotParts([part])[part];
};

test('Subtotal adds a SUBTOTAL row under each group, a bold label, a Grand Total and a three-level outline', () => {
  const view = open();
  view.select(2, 1);
  const done = view.subtotal({ groupBy: 0, fn: 'sum', columns: [2, 3] });
  assert.deepEqual(done, { groups: 3, rows: 4, range: 'A1:D11', levels: 2 });
  assert.deepEqual(column(view, 0, 0, 10), [
    'Region', 'East', 'East', 'East Total', 'North', 'North Total', 'West', 'West', 'West', 'West Total', 'Grand Total',
  ]);
  assert.equal(input(view, 3, 3), '=SUBTOTAL(9,D2:D3)');
  assert.equal(input(view, 5, 2), '=SUBTOTAL(9,C5:C5)');
  assert.equal(input(view, 10, 3), '=SUBTOTAL(9,D2:D10)', 'the grand total runs over the subtotal rows too');
  assert.deepEqual([3, 5, 9, 10].map((r) => text(view, r, 3)), ['550', '400', '420', '1370'], 'the grand total counts each row once');
  assert.deepEqual([3, 5, 9, 10].map((r) => text(view, r, 2)), ['5', '5', '7', '17']);
  assert.equal(view.styleFor(3, 0)?.font?.bold, true, 'the label is bold');
  assert.equal(view.styleFor(10, 0)?.font?.bold, true);
  assert.notEqual(view.styleFor(3, 3)?.font?.bold, true, 'the figure is not');

  const xml = partXml(view);
  assert.match(xml, /<sheetFormatPr[^>]*outlineLevelRow="2"/);
  const level = (r) => Number(/outlineLevel="(\d)"/.exec((new RegExp('<row\\b[^>]*\\br="' + r + '"[^>]*>').exec(xml) ?? [''])[0])?.[1] ?? 0);
  assert.deepEqual([2, 3, 4, 5, 6, 10, 11].map(level), [2, 2, 1, 2, 1, 1, 0], 'detail at 2, subtotals at 1, the grand total at 0');
  const outline = view.render().outline.rows;
  assert.equal(outline.levels, 2);
  assert.deepEqual(outline.groups.map((g) => [g.level, g.start, g.end, g.summary]), [
    [1, 1, 9, 10], [2, 1, 2, 3], [2, 4, 4, 5], [2, 6, 8, 9],
  ]);
});

test('SUBTOTAL leaves out the SUBTOTAL cells in its range, and nothing else', () => {
  const view = open([
    ['n'], [10], [20], ['=SUBTOTAL(9,A2:A3)'], [5], ['=SUM(A2:A3)'], ['=SUBTOTAL(9,A2:A6)'], ['=SUM(A2:A7)'], ['=SUBTOTAL(1,A2:A4)'],
  ]);
  assert.equal(text(view, 3, 0), '30');
  assert.equal(text(view, 5, 0), '30', 'a SUM is data to SUBTOTAL');
  assert.equal(text(view, 6, 0), '65', '10 + 20 + 5 + the SUM\'s 30; the inner SUBTOTAL left out');
  assert.equal(text(view, 7, 0), '160', 'SUM itself does not leave anything out');
  assert.equal(text(view, 8, 0), '15', 'AVERAGE of 10 and 20 — the SUBTOTAL cell is not a blank zero, it is not there');
});

test('Replace current subtotals takes the old ones away first; without it the new groups nest inside', () => {
  const view = open();
  view.select(1, 0);
  view.subtotal({ groupBy: 0, fn: 'sum', columns: [3] });
  view.select(1, 0);
  view.subtotal({ groupBy: 0, fn: 'count', columns: [3], replace: true });
  assert.deepEqual(column(view, 0, 0, 10), [
    'Region', 'East', 'East', 'East Count', 'North', 'North Count', 'West', 'West', 'West', 'West Count', 'Grand Count',
  ]);
  assert.equal(input(view, 3, 3), '=SUBTOTAL(3,D2:D3)');
  assert.equal(text(view, 10, 3), '6');
  assert.equal(text(view, 11, 0), '', 'no row of the first run is left');

  // Nested: sums by region, then counts by rep inside them.
  const nested = open([
    ['Region', 'Rep', 'Sales'], ['East', 'Kim', 1], ['East', 'Kim', 2], ['East', 'Lee', 3], ['West', 'Bo', 4],
  ]);
  nested.select(1, 0);
  nested.subtotal({ groupBy: 0, fn: 'sum', columns: [2] });
  nested.select(1, 0);
  nested.subtotal({ groupBy: 1, fn: 'sum', columns: [2], replace: false });
  assert.deepEqual(column(nested, 1, 0, 10).slice(0, 10), ['Rep', 'Kim', 'Kim', 'Kim Total', 'Lee', 'Lee Total', '', 'Bo', 'Bo Total', '']);
  assert.deepEqual(column(nested, 0, 6, 11), ['East Total', 'West', '', 'West Total', '', 'Grand Total'], 'the new grand row sits above the old one');
  assert.equal(text(nested, 6, 2), '6', 'the outer subtotal still counts its rows once');
  assert.equal(nested.render().outline.rows.levels, 3, 'one level deeper');
});

test('Remove All takes the subtotal rows and the outline away, back to the list as it was', () => {
  const view = open();
  const before = partXml(view);
  view.select(1, 0);
  view.subtotal({ groupBy: 0, fn: 'sum', columns: [2, 3], pageBreaks: true });
  assert.match(partXml(view), /<rowBreaks count="2" manualBreakCount="2"><brk id="4" max="16383" man="1"\/><brk id="6"/, 'a break after each group but the last');
  view.select(2, 2);
  assert.deepEqual(view.removeSubtotals(), { removed: 4, range: 'A1:D7' });
  assert.deepEqual(column(view, 0, 0, 7), ['Region', 'East', 'East', 'North', 'West', 'West', 'West', '']);
  assert.equal(text(view, 6, 3), '120');
  const xml = partXml(view);
  assert.doesNotMatch(xml, /outlineLevel|SUBTOTAL|rowBreaks/);
  assert.equal(view.render().outline, null);
  assert.throws(() => view.removeSubtotals(), /no subtotals/);
  assert.equal(xml.replace(/<sheetFormatPr[^>]*\/>/, ''), before.replace(/<sheetFormatPr[^>]*\/>/, ''), 'the rows are the list as it was');
});

test('Subtotal is one undo step, and redo puts it back', () => {
  const view = open();
  view.select(1, 0);
  view.subtotal({ groupBy: 0, fn: 'average', columns: [3] });
  assert.equal(text(view, 3, 0), 'East Average');
  assert.equal(text(view, 3, 3), '275');
  view.undo();
  assert.deepEqual(column(view, 0, 0, 7), ['Region', 'East', 'East', 'North', 'West', 'West', 'West', '']);
  assert.equal(view.render().outline, null);
  view.redo();
  assert.equal(text(view, 10, 0), 'Grand Average');
  assert.match(text(view, 10, 3), /^228.33/);
  assert.equal(view.render().outline.rows.levels, 2);
});

test('Summaries above the data put the Grand Total first and each subtotal over its group', () => {
  const view = open();
  view.select(1, 0);
  view.subtotal({ groupBy: 0, fn: 'max', columns: [3], summaryBelow: false });
  assert.deepEqual(column(view, 0, 0, 10), [
    'Region', 'Grand Max', 'East Max', 'East', 'East', 'North Max', 'North', 'West Max', 'West', 'West', 'West',
  ]);
  assert.equal(input(view, 2, 3), '=SUBTOTAL(4,D4:D5)');
  assert.equal(input(view, 1, 3), '=SUBTOTAL(4,D3:D11)');
  assert.equal(text(view, 1, 3), '400');
  assert.match(partXml(view), /<outlinePr summaryBelow="0"\/>/);
  const groups = view.render().outline.rows;
  assert.equal(groups.below, false);
  assert.deepEqual(groups.groups.filter((g) => g.level === 2).map((g) => [g.start, g.end, g.summary]), [[3, 4, 2], [6, 6, 5], [8, 10, 7]]);
  assert.throws(() => view.subtotal({ groupBy: 0, fn: 'median', columns: [3] }), /no function/);
  assert.throws(() => view.subtotal({ groupBy: 0, fn: 'sum', columns: [] }), /at least one column/);
});
