/**
 * The Data tab's two everyday tools, in the engine.
 *
 * Text to Columns splits one column's cells on a delimiter into the cells to
 * the right; Remove Duplicates drops the rows of a block that repeat an
 * earlier one and closes the rest up. Both are one undo step, and both
 * refuse or do nothing rather than guess.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const open = (rows) => new SheetView(buildXlsx({ sheets: [{ name: 'S', rows }] }));
const text = (view, row, col) => view.displayValue(row, col).text;

test('Text to Columns splits one column on the delimiter into the cells to the right', () => {
  const view = open([['Name', 'Note'], ['Kim, Lee, Park', 'x'], ['Ann;Bo', 'y'], ['plain', 'z']]);
  view.select(1, 0);
  view.select(3, 0, { extend: true });
  const done = view.textToColumns({ delimiter: 'comma' });
  assert.deepEqual(done, { rows: 1, columns: 3 }, 'one row had commas; three pieces');
  assert.equal(text(view, 1, 0), 'Kim');
  assert.equal(text(view, 1, 1), 'Lee', 'the piece lands over the Note, as Excel does');
  assert.equal(text(view, 1, 2), 'Park');
  assert.equal(text(view, 2, 0), 'Ann;Bo', 'a semicolon is not a comma');
  assert.equal(text(view, 3, 0), 'plain');

  view.undo();
  assert.equal(text(view, 1, 0), 'Kim, Lee, Park', 'one undo step');
  assert.equal(text(view, 1, 1), 'x');

  view.select(0, 0);
  view.select(1, 1, { extend: true });
  assert.throws(() => view.textToColumns({ delimiter: 'comma' }), /one column at a time/);
});

test('Remove Duplicates drops repeated rows of the block round the cell, keeping the header, and closes up', () => {
  const view = open([
    ['Region', 'Q1'], ['North', 1], ['South', 2], ['north ', 1], ['East', 3], ['South', 2], ['West', 4],
  ]);
  view.select(1, 0);
  const done = view.removeDuplicates();
  assert.equal(done.removed, 2, 'the repeated North (case and a space aside) and the repeated South');
  assert.equal(done.kept, 4);
  assert.equal(done.range, 'A1:B7');
  assert.deepEqual([1, 2, 3, 4, 5, 6].map((r) => text(view, r, 0)), ['North', 'South', 'East', 'West', '', ''], 'the rest closed up, the tail cleared');
  assert.equal(text(view, 0, 0), 'Region', 'the header stayed');
  assert.equal(text(view, 5, 1), '', 'the tail is clear in every column');

  view.undo();
  assert.equal(text(view, 5, 0), 'South', 'one undo step puts every row back');
  assert.equal(text(view, 6, 0), 'West');
});

test('Remove Duplicates on rows selected by hand treats them all as data, and says when there are none', () => {
  const view = open([['a', 1], ['a', 1], ['b', 2]]);
  view.select(0, 0);
  view.select(2, 1, { extend: true });
  const done = view.removeDuplicates();
  assert.equal(done.removed, 1, 'the first row is data when the rows were chosen by hand');
  assert.deepEqual([0, 1, 2].map((r) => text(view, r, 0)), ['a', 'b', '']);
  assert.deepEqual(view.removeDuplicates(), { removed: 0, kept: 2, range: 'A1:B3' }, 'nothing to remove reads as nothing done');
});
