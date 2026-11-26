// Ctrl+D, Ctrl+R and Home → Fill → Down, Right, Up, Left copy.
//
// They ran the fill handle's rules, so a lone date became the next days, Mon
// became Tue and Item 1 became Item 2, where Excel's Fill Down copies the
// top cell as it is. A formula's references still move, as a copy's do.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { readDatesAs } from '@rutba/formula';
import { fillOps } from '../apps/desktop/renderer/apps/sheets/fill-keys.js';

test.beforeEach(() => readDatesAs({ dayFirst: true, year: 2026 }));
test.after(() => readDatesAs());

const BOOK = () => buildXlsx({ sheets: [{ name: 'S', rows: [['Mon', 'Item 1', 5, '=C1*2'], [], [], []] }] });
const type = (view, row, col, text) => {
  view.select(row, col);
  view.beginEdit({ replace: true, initial: '' });
  view.updateDraft(text);
  view.commitEdit({ move: 'none' });
};
// The window's dispatch, as the operations table runs each one.
const run = (view, ops) => {
  for (const o of ops) {
    if (o.op === 'select') view.select(o.row, o.col, { extend: Boolean(o.extend) });
    else if (o.op === 'fill') view.fill(o.target, { mode: o.mode || 'auto' });
    else throw new Error('an operation the fill never sends: ' + o.op);
  }
};
const fillFrom = (view, direction, range) => {
  view.select(range.top, range.left);
  view.select(range.bottom, range.right, { extend: true });
  const ops = fillOps(direction, view.selection.range, { row: view.selection.active.row, col: view.selection.active.col });
  run(view, ops);
};
const input = (view, row, col) => view.calc.getInput('S', row, col);

test('Fill Down over a range copies the top row: a date, a day name and a numbered word stay as they are', () => {
  const view = new SheetView(BOOK());
  type(view, 0, 4, '30/11/2026');
  const day = input(view, 0, 4);
  assert.equal(typeof day, 'number');

  fillFrom(view, 'down', { top: 0, left: 0, bottom: 3, right: 4 });
  for (const row of [1, 2, 3]) {
    assert.equal(input(view, row, 0), 'Mon', 'row ' + row + ': the day name is copied');
    assert.equal(input(view, row, 1), 'Item 1', 'row ' + row + ': the numbered word is copied');
    assert.equal(input(view, row, 2), 5);
    assert.equal(input(view, row, 3), `=C${row + 1}*2`, 'row ' + row + ': the formula moves with its cell');
    assert.equal(input(view, row, 4), day, 'row ' + row + ': the date is that date');
    assert.equal(view.displayValue(row, 4).text, '30/11/2026', 'and wears the date format');
  }
  const r = view.selection.range;
  assert.deepEqual([r.top, r.left, r.bottom, r.right], [0, 0, 3, 4], 'the selection is what it was');
  view.undo();
  assert.ok(!input(view, 1, 0), 'one undo step');
});

test('Ctrl+D in one cell copies the cell above; Ctrl+R the cell to the left', () => {
  const view = new SheetView(BOOK());
  type(view, 0, 4, '30/11/2026');
  fillFrom(view, 'down', { top: 1, left: 4, bottom: 1, right: 4 });
  assert.equal(input(view, 1, 4), input(view, 0, 4), 'the date is copied, not the next day');
  assert.deepEqual([view.selection.active.row, view.selection.active.col], [1, 4]);

  fillFrom(view, 'down', { top: 1, left: 0, bottom: 1, right: 0 });
  assert.equal(input(view, 1, 0), 'Mon');

  type(view, 5, 0, 'Week 1');
  fillFrom(view, 'right', { top: 5, left: 1, bottom: 5, right: 1 });
  assert.equal(input(view, 5, 1), 'Week 1');
  fillFrom(view, 'right', { top: 5, left: 0, bottom: 5, right: 3 });
  assert.deepEqual([1, 2, 3].map((c) => input(view, 5, c)), ['Week 1', 'Week 1', 'Week 1']);
});

test('Fill Up and Left copy from the last row and column; a cell at the edge has nothing to fill from', () => {
  const view = new SheetView(BOOK());
  type(view, 3, 0, 'Tue');
  fillFrom(view, 'up', { top: 1, left: 0, bottom: 3, right: 0 });
  assert.deepEqual([1, 2].map((r) => input(view, r, 0)), ['Tue', 'Tue']);
  type(view, 6, 3, 'Q1');
  fillFrom(view, 'left', { top: 6, left: 1, bottom: 6, right: 3 });
  assert.deepEqual([1, 2].map((c) => input(view, 6, c)), ['Q1', 'Q1']);

  assert.equal(fillOps('down', { top: 0, left: 2, bottom: 0, right: 2 }, { row: 0, col: 2 }), null);
  assert.equal(fillOps('right', { top: 4, left: 0, bottom: 4, right: 0 }, { row: 4, col: 0 }), null);
});

test('the fill handle still makes a series', () => {
  const view = new SheetView(BOOK());
  view.select(0, 0);
  view.fill({ top: 0, left: 0, bottom: 2, right: 0 });
  assert.deepEqual([1, 2].map((r) => input(view, r, 0)), ['Tue', 'Wed']);
});
