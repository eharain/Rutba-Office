// A press on the grid is for the cell under the pointer.
//
// Words longer than their column are drawn over the empty cells beside them,
// so the cell's box on the screen is wider than the cell. A click on the
// empty G1, where F1's "Gift wrapped" lay over it, was taken by F1's box and
// selected F1: what was typed next replaced the heading. The press is now
// for the cell under the pointer.

import test from 'node:test';
import assert from 'node:assert/strict';
import { cellUnder, pressedCell } from '../apps/desktop/renderer/apps/sheets/press.js';

// Columns E to H, 64 wide from x = 256; rows 1 to 3, 20 high.
const MODEL = {
  columns: [4, 5, 6, 7].map((index, i) => ({ index, x: 256 + i * 64, width: 64 })),
  rows: [0, 1, 2].map((index) => ({ index, y: index * 20, height: 20 })),
};
const F1 = { row: 0, col: 5, x: 320, y: 0, width: 64, height: 20, ref: 'F1' };

test('the frame says which cell is under a point', () => {
  assert.deepEqual(cellUnder(MODEL, 330, 5), { row: 0, col: 5, x: 320 });
  assert.deepEqual(cellUnder(MODEL, 384, 25), { row: 1, col: 6, x: 384 }, 'a column starts on its own left edge');
  assert.equal(cellUnder(MODEL, 2000, 5), null, 'off the columns in the frame');
  assert.equal(cellUnder(null, 1, 1), null);
});

test('a press on words lying over the next cell is for that cell', () => {
  // F1's box is drawn from 320 to 448: its own column and G's.
  assert.deepEqual(pressedCell(F1, cellUnder(MODEL, 400, 8)), { row: 0, col: 6 }, 'G1, not F1');
  assert.deepEqual(pressedCell(F1, cellUnder(MODEL, 350, 8)), { row: 0, col: 5 }, 'over its own column it is F1');
  // Right-aligned words lie over the cell on the left.
  assert.deepEqual(pressedCell(F1, cellUnder(MODEL, 300, 8)), { row: 0, col: 4 }, 'E1, for words lying to the left');
});

test('a merged cell is its own cell all the way across, and a point off the frame changes nothing', () => {
  const merged = { ...F1, width: 128 }; // F1:G1 merged
  assert.deepEqual(pressedCell(merged, cellUnder(MODEL, 400, 8)), { row: 0, col: 5 });
  assert.deepEqual(pressedCell(F1, null), { row: 0, col: 5 });
  // The row is the cell's own: words lie over neighbours across, never down.
  assert.deepEqual(pressedCell({ ...F1, height: 40 }, cellUnder(MODEL, 400, 30)), { row: 0, col: 6 });
});
