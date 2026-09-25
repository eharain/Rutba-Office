/**
 * Data → Group and Ungroup, the outline's buttons, Show and Hide Detail,
 * Clear Outline — in the engine.
 *
 * An outline is `outlineLevel` on rows and columns, `hidden` on the ones
 * folded away and `collapsed` on the summary of a folded group, with the
 * deepest levels in `sheetFormatPr` and the summaries' side in
 * `sheetPr/outlinePr`. These read what the engine writes, what it reads
 * back, and what an Excel-authored file carries.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

const rows = (n) => Array.from({ length: n }, (_, r) => [r === 0 ? 'Item' : 'Row ' + (r + 1), r === 0 ? 'Qty' : r]);
const open = (n = 12) => new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: rows(n) }] }));
const partXml = (view) => {
  const part = view.workbook.partNameFor(view.activeSheet);
  return view.workbook.snapshotParts([part])[part];
};
const rowTag = (xml, r) => (new RegExp('<row\\b[^>]*\\br="' + r + '"[^>]*>').exec(xml) ?? [''])[0];
const selectRows = (view, a, b) => { view.select(a, 0); view.select(b, 1, { extend: true }); };

test('Group writes outlineLevel on the rows and the depth on sheetFormatPr; Ungroup takes both off', () => {
  const view = open();
  selectRows(view, 1, 3);
  const done = view.group();
  assert.deepEqual(done, { axis: 'row', from: 1, to: 3, levels: 1 });
  const xml = partXml(view);
  for (const r of [2, 3, 4]) assert.match(rowTag(xml, r), /outlineLevel="1"/, 'row ' + r);
  assert.doesNotMatch(rowTag(xml, 5), /outlineLevel/, 'the summary row is not in the group');
  assert.match(xml, /<sheetFormatPr\b[^>]*outlineLevelRow="1"/);
  const frame = view.render().outline;
  assert.equal(frame.rows.levels, 1);
  assert.deepEqual(frame.rows.groups.map((g) => [g.level, g.start, g.end, g.summary, g.collapsed]), [[1, 1, 3, 4, false]]);
  assert.equal(frame.cols, null, 'no columns are grouped');

  view.ungroup();
  const after = partXml(view);
  assert.doesNotMatch(after, /outlineLevel="/, 'no row keeps a level');
  assert.doesNotMatch(after, /outlineLevelRow/, 'nor the sheet its depth');
  assert.equal(view.render().outline, null);
  assert.throws(() => view.ungroup(), /not in a group/);
});

test('Groups nest to three levels and more, and an eighth level is refused', () => {
  const view = open(20);
  selectRows(view, 1, 10);
  view.group();
  selectRows(view, 2, 6);
  view.group();
  selectRows(view, 3, 4);
  view.group();
  const g = view.render().outline.rows;
  assert.equal(g.levels, 3);
  assert.deepEqual(g.groups.map((x) => [x.level, x.start, x.end]), [[1, 1, 10], [2, 2, 6], [3, 3, 4]]);
  assert.match(partXml(view), /outlineLevelRow="3"/);
  assert.match(rowTag(partXml(view), 4), /outlineLevel="3"/);

  selectRows(view, 3, 3);
  for (let i = 0; i < 4; i++) view.group();
  assert.equal(view.geo.rowLevels.get(3), 7, 'seven levels, Excel\'s limit');
  assert.throws(() => view.group(), /seven levels/);
  assert.equal(view.geo.rowLevels.get(3), 7, 'the refused press changed nothing');
});

test('Collapsing a group hides its rows and marks the summary collapsed; expanding keeps an inner fold', () => {
  const view = open(20);
  selectRows(view, 1, 8);
  view.group();
  selectRows(view, 2, 4);
  view.group();
  // The inner group (rows 3-5 in the file) folded first.
  view.toggleOutlineGroup({ axis: 'row', level: 2, start: 2 });
  let xml = partXml(view);
  for (const r of [3, 4, 5]) assert.match(rowTag(xml, r), /hidden="1"/, 'row ' + r + ' hidden');
  assert.match(rowTag(xml, 6), /collapsed="1"/, 'the inner summary marked');
  assert.equal(view.geo.rowHeight(3), 0, 'a folded row takes no space');

  // Then the outer, then the outer open again: the inner stays folded.
  view.toggleOutlineGroup({ axis: 'row', level: 1, start: 1 });
  assert.equal(view.geo.rowHeight(1), 0);
  assert.match(rowTag(partXml(view), 10), /collapsed="1"/);
  view.toggleOutlineGroup({ axis: 'row', level: 1, start: 1 });
  xml = partXml(view);
  assert.doesNotMatch(rowTag(xml, 2), /hidden/, 'the outer group\'s own rows are back');
  assert.match(rowTag(xml, 4), /hidden="1"/, 'the inner fold is kept, as Excel keeps it');
  assert.doesNotMatch(rowTag(xml, 10), /collapsed/, 'the outer summary is open');
  const frame = view.render();
  assert.ok(!frame.rows.some((r) => r.index === 3), 'the viewport skips the folded rows');
  assert.ok(frame.rows.some((r) => r.index === 5), 'and draws the inner summary');

  // One undo step each: the last unfold undone folds the outer again.
  view.undo();
  assert.equal(view.geo.rowHeight(1), 0, 'undo re-reads the part and the fold is back');
  view.undo();
  view.undo();
  assert.ok(view.geo.rowHeight(3) > 0, 'three undos: nothing folded');
});

test('The level buttons show the rows above a level and fold the rest; the last shows all', () => {
  const view = open(20);
  selectRows(view, 1, 8);
  view.group();
  selectRows(view, 2, 4);
  view.group();
  view.showOutlineLevel({ axis: 'row', level: 1 });
  for (let r = 1; r <= 8; r++) assert.equal(view.geo.rowHeight(r), 0, 'row ' + r + ' folded at level 1');
  assert.ok(view.geo.rowHeight(9) > 0, 'the summary shows');
  view.showOutlineLevel({ axis: 'row', level: 2 });
  assert.ok(view.geo.rowHeight(1) > 0);
  assert.equal(view.geo.rowHeight(2), 0, 'level 2 rows fold at button 2');
  assert.ok(view.geo.rowHeight(5) > 0, 'the inner summary shows');
  const groups = view.render().outline.rows.groups;
  assert.deepEqual(groups.map((g) => g.collapsed), [false, true]);
  view.showOutlineLevel({ axis: 'row', level: 3 });
  for (let r = 1; r <= 9; r++) assert.ok(view.geo.rowHeight(r) > 0, 'row ' + r + ' shown at the last button');
  assert.doesNotMatch(partXml(view), /collapsed=|hidden=/);
});

test('Show Detail and Hide Detail act on the group at the active cell; Clear Outline takes every level away', () => {
  const view = open(20);
  selectRows(view, 1, 4);
  view.group();
  view.select(2, 0);
  view.hideDetail();
  assert.equal(view.geo.rowHeight(2), 0);
  view.select(5, 0);
  view.showDetail();
  assert.ok(view.geo.rowHeight(2) > 0, 'Show Detail on the summary row opens its group');
  view.select(2, 0);
  view.hideDetail();
  view.clearOutline();
  const xml = partXml(view);
  assert.doesNotMatch(xml, /outlineLevel|collapsed=|hidden=/);
  assert.ok(view.geo.rowHeight(2) > 0, 'rows folded by the outline come back');
  assert.throws(() => view.clearOutline(), /no outline/);
  assert.throws(() => view.showDetail(), /no folded group/);
});

test('Whole columns selected group as columns, on <col> records, and fold the same way', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: [['a', 'b', 'c', 'd', 'e'], [1, 2, 3, 4, 5]] }] }));
  view.selectColumn(1);
  view.selectColumn(2, { extend: true });
  assert.equal(view.group().axis, 'col');
  let xml = partXml(view);
  assert.match(xml, /<col min="2" max="2" outlineLevel="1"\/>/);
  assert.match(xml, /<col min="3" max="3" outlineLevel="1"\/>/);
  assert.match(xml, /outlineLevelCol="1"/);
  view.toggleOutlineGroup({ axis: 'col', level: 1, start: 1 });
  xml = partXml(view);
  assert.match(xml, /<col min="2" max="2" outlineLevel="1" hidden="1"\/>/);
  assert.match(xml, /<col min="4" max="4" collapsed="1"\/>/);
  assert.equal(view.geo.colWidth(2), 0);
  const frame = view.render();
  assert.equal(frame.outline.cols.levels, 1);
  assert.deepEqual(frame.outline.cols.groups.map((g) => [g.start, g.end, g.summary, g.collapsed]), [[1, 2, 3, true]]);
  assert.ok(!frame.columns.some((c) => c.index === 1), 'a folded column is not drawn');
});

test('An outline survives save and reopen, summaries above included', () => {
  const view = open(12);
  selectRows(view, 2, 5);
  view.group();
  view.toggleOutlineGroup({ axis: 'row', level: 1, start: 2 });
  view.workbook._sheetPart('S').part.setOutlineProps({ summaryBelow: false });
  const again = new SheetView(view.save());
  assert.equal(again.geo.rowLevels.get(3), 1);
  assert.ok(again.geo.hiddenRows.has(4));
  assert.equal(again.geo.summaryBelow, false, 'outlinePr summaryBelow="0" read back');
  const g = again.render().outline.rows;
  assert.equal(g.below, false);
  assert.deepEqual(g.groups.map((x) => [x.start, x.end, x.summary]), [[2, 5, 1]], 'the summary is the row above');
  const xml = again.pkg.text(again.workbook.partNameFor('S'));
  assert.match(xml, /<sheetPr><outlinePr summaryBelow="0"\/><\/sheetPr>/);
  assert.match(xml, /<sheetFormatPr[^>]*outlineLevelRow="1"/);
});

test('An outline Excel wrote reads as groups: levels, a folded group, grouped columns', () => {
  const pkg = OoxmlPackage.read(buildXlsx({ sheets: [{ name: 'S', rows: rows(10) }] }));
  const part = 'xl/worksheets/sheet1.xml';
  // As Excel 365 writes a sheet with rows 2-4 folded under row 5, rows
  // 6-8 open under row 9 inside a level-1 group of 2-9, and columns B:C
  // grouped.
  let xml = pkg.text(part);
  xml = xml.replace(/<sheetData>[\s\S]*<\/sheetData>/, '<sheetData>'
    + '<row r="1"><c r="A1" t="inlineStr"><is><t>Head</t></is></c></row>'
    + '<row r="2" hidden="1" outlineLevel="2"><c r="A2"><v>1</v></c></row>'
    + '<row r="3" hidden="1" outlineLevel="2"><c r="A3"><v>2</v></c></row>'
    + '<row r="4" hidden="1" outlineLevel="2"><c r="A4"><v>3</v></c></row>'
    + '<row r="5" collapsed="1" outlineLevel="1"><c r="A5"><v>6</v></c></row>'
    + '<row r="6" outlineLevel="2"><c r="A6"><v>1</v></c></row>'
    + '<row r="7" outlineLevel="2"><c r="A7"><v>1</v></c></row>'
    + '<row r="8" outlineLevel="2"><c r="A8"><v>1</v></c></row>'
    + '<row r="9" outlineLevel="1"><c r="A9"><v>3</v></c></row>'
    + '<row r="10"><c r="A10"><v>9</v></c></row>'
    + '</sheetData>');
  xml = xml.replace('<sheetData>', '<sheetFormatPr defaultRowHeight="15" outlineLevelRow="2" outlineLevelCol="1" x14ac:dyDescent="0.25"/>'
    + '<cols><col min="2" max="3" width="9.140625" outlineLevel="1"/></cols><sheetData>');
  pkg.write_(part, xml);
  const view = new SheetView(pkg.write());
  const rowsOutline = view.render().outline.rows;
  assert.equal(rowsOutline.levels, 2);
  assert.deepEqual(rowsOutline.groups.map((g) => [g.level, g.start, g.end, g.summary, g.collapsed]), [
    [1, 1, 8, 9, false], [2, 1, 3, 4, true], [2, 5, 7, 8, false],
  ]);
  assert.equal(view.geo.rowHeight(2), 0);
  const cols = view.render().outline.cols;
  assert.deepEqual(cols.groups.map((g) => [g.level, g.start, g.end, g.summary]), [[1, 1, 2, 3]]);
  // Pressing its + opens the folded group and writes only what changed.
  view.toggleOutlineGroup({ axis: 'row', level: 2, start: 1 });
  const out = partXml(view);
  assert.match(rowTag(out, 2), /^<row r="2" outlineLevel="2">$/);
  assert.match(rowTag(out, 5), /^<row r="5" outlineLevel="1">$/);
  assert.match(out, /x14ac:dyDescent="0.25"/, 'the rest of sheetFormatPr is left as it was');
});
