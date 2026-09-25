/**
 * View → Split: the window in four (or two) panes at the active cell, the
 * top and left ones scrolling on their own. Written as Excel writes a split
 * — a `<pane>` with xSplit and ySplit (twentieths of a point from the
 * window's edge) and no frozen state, `topLeftCell` for the bottom-right
 * pane and the sheet view's own for the top-left — and told apart from
 * Freeze Panes, which it replaces and which replaces it.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { OoxmlPackage } from '@rutba/ooxml';
import { SheetView } from '@rutba/sheet-view';

const grid = (rows, cols) => Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => r * 100 + c));
const open = (rows = 200, cols = 30) => new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: grid(rows, cols) }] }), { viewportWidth: 1000, viewportHeight: 600 });
const sheetXml = (view) => {
  const part = view.workbook.partNameFor(view.activeSheet);
  return view.workbook.snapshotParts([part])[part];
};
const paneTag = (view) => (/<pane\b[^>]*\/>/.exec(sheetXml(view)) ?? [''])[0];

test('Split at a cell writes the pane Excel writes, with no frozen state', () => {
  const view = open();
  view.select(5, 3);
  view.toggleSplit();
  const pane = paneTag(view);
  const geo = view.geo;
  assert.match(pane, new RegExp('xSplit="' + (geo.colOffset(3) + geo.headerWidth) * 15 + '"'));
  assert.match(pane, new RegExp('ySplit="' + (geo.rowOffset(5) + geo.headerHeight) * 15 + '"'));
  assert.match(pane, /topLeftCell="D6"/);
  assert.match(pane, /activePane="bottomRight"/);
  assert.doesNotMatch(pane, /state=/);
  assert.match(sheetXml(view), /<sheetView\b[^>]*topLeftCell="A1"/);
  assert.equal(view.frozenPane().rows, 0, 'a split is not a freeze');
  const s = view.splitPane();
  assert.deepEqual(s, { width: geo.colOffset(3), height: geo.rowOffset(5), top: 0, left: 0 });
  const frame = view.render();
  assert.deepEqual(frame.split.rows, [0, 1, 2, 3, 4]);
  assert.deepEqual(frame.split.cols, [0, 1, 2]);
  view.toggleSplit();
  assert.equal(paneTag(view), '');
  assert.equal(view.render().split, null);
});

test('in the first row or column showing the window splits one way; at the top-left cell, in four at the middle', () => {
  const view = open();
  view.select(0, 4);
  view.splitAt();
  let s = view.splitPane();
  assert.equal(s.height, 0);
  assert.equal(s.width, view.geo.colOffset(4));
  assert.match(paneTag(view), /activePane="topRight"/);
  view.removeSplit();
  view.select(0, 0);
  view.splitAt();
  s = view.splitPane();
  assert.ok(s.width > 400 && s.width < 600, 'about half the window across: ' + s.width);
  assert.ok(s.height > 250 && s.height < 350, 'about half the window down: ' + s.height);
  assert.equal(s.height % view.geo.defaultRowHeight, 0, 'on a row edge');
});

test('the bars snap to the nearest edge, and dragged to the edge they go', () => {
  const view = open();
  view.select(5, 3);
  view.splitAt();
  view.setSplit({ height: 107, width: 100 });
  const s = view.splitPane();
  assert.equal(s.height, 100, '107 px is nearest the fifth row edge');
  assert.equal(s.width, 128, '100 px is nearest the second column edge');
  view.setSplit({ height: 0, width: 128 });
  assert.equal(view.splitPane().height, 0);
  assert.match(paneTag(view), /activePane="topRight"/);
  view.setSplit({ height: 0, width: 0 });
  assert.equal(view.splitPane(), null);
});

test('the top pane scrolls on its own and draws the rows it has scrolled to; the file is not touched', () => {
  const view = open();
  view.select(5, 3);
  view.splitAt();
  const before = sheetXml(view);
  view.scrollSplit({ rows: 40, cols: 2 });
  const s = view.splitPane();
  assert.equal(s.top, 40);
  assert.equal(s.left, 2);
  const frame = view.render();
  assert.equal(frame.split.rows[0], 40);
  assert.equal(frame.split.topY, view.geo.rowOffset(40));
  assert.ok(frame.cells.some((c) => c.row === 40 && c.col === 5), 'the top pane\'s cells are in the frame');
  assert.ok(frame.cells.some((c) => c.row === 20 && c.col === 2), 'the left pane\'s cells beside the main rows');
  assert.ok(frame.cells.some((c) => c.row === 41 && c.col === 2), 'and the corner\'s');
  assert.equal(sheetXml(view), before);
  view.scrollSplit({ rows: -100 });
  assert.equal(view.splitPane().top, 0, 'never before the first row');
});

test('Freeze Panes replaces a split, and a split replaces a freeze', () => {
  const view = open();
  view.select(5, 3);
  view.splitAt();
  view.freezePanes(1, 0);
  assert.deepEqual(view.frozenPane(), { rows: 1, cols: 0 });
  assert.equal(view.splitPane(), null);
  assert.match(paneTag(view), /state="frozen"/);
  view.select(4, 2);
  view.splitAt();
  assert.equal(view.frozenPane().rows, 0);
  assert.doesNotMatch(paneTag(view), /state=/);
  assert.equal((sheetXml(view).match(/<pane\b/g) || []).length, 1);
});

test('a split written by Excel is read, the top-left pane where its sheet view says', () => {
  const pkg = OoxmlPackage.read(buildXlsx({ sheets: [{ name: 'S', rows: grid(50, 10) }] }));
  const part = 'xl/worksheets/sheet1.xml';
  pkg.write_(part, pkg.text(part).replace(/<sheetData/, '<sheetViews><sheetView tabSelected="1" topLeftCell="A11" workbookViewId="0"><pane xSplit="2280" ySplit="2385" topLeftCell="D22" activePane="bottomRight"/><selection pane="topRight"/><selection pane="bottomRight" activeCell="D22" sqref="D22"/></sheetView></sheetViews><sheetData'));
  const view = new SheetView(pkg.write());
  const s = view.splitPane();
  assert.equal(s.width, Math.round(2280 / 15) - view.geo.headerWidth);
  assert.equal(s.height, Math.round(2385 / 15) - view.geo.headerHeight);
  assert.equal(s.top, 10);
  assert.equal(s.left, 0);
  assert.equal(view.frozenPane().rows, 0);
});
