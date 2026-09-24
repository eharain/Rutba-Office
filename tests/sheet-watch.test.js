/**
 * Formulas → Watch Window, in the engine: the plain list logic behind the
 * pane (add, remove, keep order) and reading a watch's live value and
 * formula text off a workbook — the same cell lookup Error Checking reads,
 * just aimed at the cells someone chose rather than every error on a sheet.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView, addWatches, removeWatch, resolveWatches } from '@rutba/sheet-view';

const view = () => new SheetView(buildXlsx({ sheets: [
  { name: 'S', rows: [
    [1, 2, '=A1+B1'],
    [4, 5, '=1/0'],
  ] },
  { name: 'Other', rows: [[9, '=A1*2']] },
] }));

test('addWatches dedupes: the same cell twice, or already on the list, leaves it once', () => {
  const once = addWatches([], [{ sheet: 'S', row: 0, col: 0 }, { sheet: 'S', row: 0, col: 0 }]);
  assert.equal(once.length, 1);
  const again = addWatches(once, [{ sheet: 'S', row: 0, col: 0 }]);
  assert.equal(again.length, 1);
  assert.deepEqual(again, once);
});

test('addWatches keeps order: new refs land after what was already there, in the order given', () => {
  const list = addWatches([], [{ sheet: 'S', row: 0, col: 0 }, { sheet: 'S', row: 0, col: 1 }]);
  const grown = addWatches(list, [{ sheet: 'Other', row: 0, col: 0 }, { sheet: 'S', row: 1, col: 2 }]);
  assert.deepEqual(grown, [
    { sheet: 'S', row: 0, col: 0 },
    { sheet: 'S', row: 0, col: 1 },
    { sheet: 'Other', row: 0, col: 0 },
    { sheet: 'S', row: 1, col: 2 },
  ]);
});

test('removeWatch drops one watch and leaves the rest, in the order they were', () => {
  const list = addWatches([], [
    { sheet: 'S', row: 0, col: 0 },
    { sheet: 'S', row: 0, col: 1 },
    { sheet: 'S', row: 1, col: 2 },
  ]);
  const after = removeWatch(list, { sheet: 'S', row: 0, col: 1 });
  assert.deepEqual(after, [
    { sheet: 'S', row: 0, col: 0 },
    { sheet: 'S', row: 1, col: 2 },
  ]);
  // A ref nobody is watching changes nothing.
  assert.deepEqual(removeWatch(after, { sheet: 'S', row: 5, col: 5 }), after);
});

test('resolveWatches reads a watch\'s live value and its formula text', () => {
  const v = view();
  const list = addWatches([], [{ sheet: 'S', row: 0, col: 2 }, { sheet: 'S', row: 0, col: 0 }]);
  const rows = resolveWatches(list, v);
  assert.deepEqual(rows, [
    { sheet: 'S', row: 0, col: 2, ref: 'C1', value: '3', formula: '=A1+B1' },
    { sheet: 'S', row: 0, col: 0, ref: 'A1', value: '1', formula: null },
  ]);
});

test('resolveWatches keeps reading live: a precedent edited changes the watched formula\'s value', () => {
  const v = view();
  const list = addWatches([], [{ sheet: 'S', row: 0, col: 2 }]);
  assert.equal(resolveWatches(list, v)[0].value, '3');
  v.setCell(0, 0, 10); // A1: was 1, now 10 — C1 (=A1+B1) follows on the next read
  assert.equal(resolveWatches(list, v)[0].value, '12');
});

test('a removed sheet\'s watches drop out rather than throwing', () => {
  const v = view();
  const list = addWatches([], [{ sheet: 'S', row: 0, col: 0 }, { sheet: 'Other', row: 0, col: 0 }]);
  v.removeSheet('Other');
  const rows = resolveWatches(list, v);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].sheet, 'S');
});
