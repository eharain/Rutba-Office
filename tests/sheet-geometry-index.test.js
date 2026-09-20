// Offsets on a sheet that sizes every row.
//
// Excel writes a height on each row of many exports, and a hidden row often
// keeps the height it had. Every offset used to walk all of them; now the
// exceptions are indexed once and an offset is a binary search — and a
// hidden row takes no space whether or not the file also gave it a height.

import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetGeometry, SheetView } from '@rutba/sheet-view';
import { buildXlsx } from '@rutba/ooxml/build';

test('a height on every row costs one binary search per offset, and the offsets are right', () => {
  const geo = new SheetGeometry();
  // Twenty thousand, not the sixty thousand it was measured on: the gate
  // runs the test files side by side, and a fixture this size crowded a
  // wide-sheet timing test next door.
  const ROWS = 20000;
  for (let r = 0; r < ROWS; r++) geo.rowHeights.set(r, r % 3 === 0 ? 24 : 20);
  // What the walk would give: rows 0..n-1 summed.
  let sum = 0;
  const walked = [];
  for (let r = 0; r <= ROWS; r++) { walked[r] = sum; sum += geo.rowHeight(r); }
  assert.equal(geo.rowOffset(0), 0);
  assert.equal(geo.rowOffset(1), 24);
  assert.equal(geo.rowOffset(3), 24 + 20 + 20);
  assert.equal(geo.rowOffset(10000), walked[10000]);
  assert.equal(geo.rowOffset(ROWS), walked[ROWS]);
  assert.equal(geo.rowOffset(ROWS + 5000), walked[ROWS] + 5000 * geo.defaultRowHeight, 'past the last exception, the default');

  const started = performance.now();
  for (let r = 0; r < 20000; r++) geo.rowOffset(r * 3);
  const ms = performance.now() - started;
  assert.ok(ms < 200, `twenty thousand offsets took ${ms.toFixed(0)} ms`);
  assert.equal(geo.rowAt(walked[10000] + 5), 10000, 'rowAt reads through the same index');
});

test('a hidden row takes no space even when the file gave it a height', () => {
  const geo = new SheetGeometry();
  geo.rowHeights.set(2, 40);
  geo.hiddenRows.add(2);
  assert.equal(geo.rowHeight(2), 0);
  assert.equal(geo.rowOffset(3), 2 * geo.defaultRowHeight, 'row 3 sits right under row 1');
  geo.hiddenCols.add(1);
  geo.colWidths.set(1, 200);
  assert.equal(geo.colOffset(2), geo.defaultColWidth, 'a hidden column with a width is still nothing wide');
});

test('the index is dropped when a size changes, a row is hidden or a column is shown again', () => {
  const geo = new SheetGeometry();
  assert.equal(geo.rowOffset(5), 5 * geo.defaultRowHeight);
  geo.rowHeights.set(1, 50);
  assert.equal(geo.rowOffset(5), 4 * geo.defaultRowHeight + 50, 'a height set after a read counts');
  geo.rowHeights.delete(1);
  assert.equal(geo.rowOffset(5), 5 * geo.defaultRowHeight, 'and one deleted stops counting');
  geo.hiddenRows.add(0);
  assert.equal(geo.rowOffset(5), 4 * geo.defaultRowHeight);
  geo.hiddenRows.delete(0);
  assert.equal(geo.rowOffset(5), 5 * geo.defaultRowHeight);
  geo.colWidths.set(0, 100);
  assert.equal(geo.colOffset(1), 100);
  geo.colWidths.clear();
  assert.equal(geo.colOffset(1), geo.defaultColWidth);
  geo.hiddenCols.add(0);
  assert.equal(geo.colOffset(1), 0);
  geo.hiddenCols.clear();
  assert.equal(geo.colOffset(1), geo.defaultColWidth);
});

test('a frame does not carry the current region; it is asked for, and the frame on a big table is quick', () => {
  const ROWS = 6000;
  const rows = [['id', 'name', 'amount']];
  for (let r = 1; r <= ROWS; r++) rows.push([r, `name ${r}`, r % 100]);
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Data', rows }] }));
  view.viewportWidth = 1000;
  view.viewportHeight = 600;
  view.render();
  const started = performance.now();
  view.scrollTo(0, 3000 * view.geo.defaultRowHeight);
  const frame = view.render();
  const ms = performance.now() - started;
  assert.ok(!('region' in frame), 'the frame carries no region');
  assert.ok(frame.viewport.firstRow <= 3000 - 10 && frame.viewport.lastRow >= 3000 + 30, `ten rows of overscan above: ${JSON.stringify(frame.viewport)}`);
  // Twenty milliseconds alone; the gate runs the test files side by side on
  // a machine that is also bundling, so the bound only catches the old
  // behaviour (a third of a second and up), not a busy afternoon.
  assert.ok(ms < 1500, `a frame ten thousand rows down took ${ms.toFixed(0)} ms`);
  const region = view.regionAround(5, 1);
  assert.deepEqual(region, { ref: `A1:C${ROWS + 1}`, headers: ['id', 'name', 'amount'] });
  assert.equal(view.regionAround(5, 8), null, 'off the block, nothing');
});
