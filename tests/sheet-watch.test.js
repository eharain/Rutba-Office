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

// The pane showed 6651.75999999999 for a cell shown as £6,651.76, and 0.0569
// for 5.7%: the value was read without the cell's number format.
test('a watch shows its value as the cell shows it, on any sheet', () => {
  const v = view();
  v.setCell(2, 0, '=6651.76-0.00000000001');
  v.select(2, 0);
  v.setFormat({ numberFormat: '"£"#,##0.00' });
  v.setCell(2, 1, 0.0569);
  v.select(2, 1);
  v.setFormat({ numberFormat: '0.0%' });
  v.setCell(2, 2, 46356);
  v.select(2, 2);
  v.setFormat({ numberFormat: 'dd/mm/yyyy' });
  const list = addWatches([], [{ sheet: 'S', row: 2, col: 0 }, { sheet: 'S', row: 2, col: 1 }, { sheet: 'S', row: 2, col: 2 }, { sheet: 'S', row: 0, col: 0 }]);
  assert.deepEqual(resolveWatches(list, v).map((w) => w.value), ['£6,651.76', '5.7%', '30/11/2026', '1']);

  // Read from another sheet, the watch still wears its own cell's format.
  v.selectSheet('Other');
  assert.deepEqual(resolveWatches(list, v).map((w) => w.value), ['£6,651.76', '5.7%', '30/11/2026', '1']);
  v.select(0, 0);
  v.setFormat({ numberFormat: '0.00' });
  assert.equal(resolveWatches([{ sheet: 'Other', row: 0, col: 0 }], v)[0].value, '9.00');
  assert.equal(resolveWatches([{ sheet: 'S', row: 0, col: 0 }], v)[0].value, '1', 'and not the format of the cell at the same place on the sheet in view');
});

// The pane lay over the bottom of the grid, so the rows under it could not be
// clicked. It now has room of its own between the grid and the sheet tabs.
test('the Watch Window is docked under the grid, not laid over it', async () => {
  const { readFile } = await import('node:fs/promises');
  const src = await readFile(new URL('../apps/desktop/renderer/apps/sheets.js', import.meta.url), 'utf8');
  const rule = src.slice(src.indexOf('\n.sh-watch {'), src.indexOf('}', src.indexOf('\n.sh-watch {')));
  assert.ok(rule.length > 20, 'the rule is there');
  assert.ok(!rule.includes('absolute'), 'it takes space of its own');
  const body = src.indexOf('<div className="sh-body">');
  const pane = src.indexOf('{watchPane()}');
  const tabs = src.indexOf('<div className={`sh-tabs');
  assert.ok(body > 0 && pane > body && pane < tabs, 'drawn between the grid and the sheet tabs');
  assert.ok(src.slice(body, pane).includes('{commentsPane()}'), 'and outside the grid\'s own box, which ends after its side panes');
});
