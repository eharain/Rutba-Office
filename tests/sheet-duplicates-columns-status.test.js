/**
 * Remove Duplicates by the columns ticked, with a count of what went and
 * what stays, and the status bar for cells picked one by one with Ctrl.
 * Before 2.28 every column of the block was compared, nothing said how many
 * rows went, and a Ctrl selection showed no Count, Sum or Average.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { buildXlsx } from '@rutba/ooxml/build';

const ROWS = [
  ['Name', 'City', 'Amount'],
  ['Ann', 'Leeds', 5],
  ['Bob', 'York', 7],
  ['Ann', 'Leeds', 9],
  ['Ann', 'Hull', 5],
  ['Bob', 'York', 7],
];
const open = () => {
  const view = SheetView.open(buildXlsx({ sheets: [{ name: 'Data', rows: ROWS }] }));
  view.select(1, 0);
  return view;
};
const names = (view) => {
  const out = [];
  for (let r = 1; r <= 5; r++) out.push(view.calc.getValue('Data', r, 0) + ':' + view.calc.getValue('Data', r, 1) + ':' + view.calc.getValue('Data', r, 2));
  return out;
};

test('the plan for the dialog: the block, its heading row and its columns', () => {
  const plan = open().duplicatesPlan();
  assert.equal(plan.range, 'A1:C6');
  assert.equal(plan.header, true);
  assert.deepEqual(plan.columns, [
    { col: 0, letter: 'A', heading: 'Name' },
    { col: 1, letter: 'B', heading: 'City' },
    { col: 2, letter: 'C', heading: 'Amount' },
  ]);
});

test('every column compared, as before: only the whole-row repeat goes, and the count says so', () => {
  const view = open();
  const said = view.removeDuplicates();
  assert.deepEqual({ removed: said.removed, kept: said.kept }, { removed: 1, kept: 4 });
  assert.deepEqual(names(view), ['Ann:Leeds:5', 'Bob:York:7', 'Ann:Leeds:9', 'Ann:Hull:5', '::']);
});

test('two columns ticked: a row repeats when Name and City match, and the first of each stays whole', () => {
  const view = open();
  const said = view.removeDuplicates({ columns: [0, 1] });
  assert.deepEqual({ removed: said.removed, kept: said.kept }, { removed: 2, kept: 3 });
  assert.deepEqual(names(view), ['Ann:Leeds:5', 'Bob:York:7', 'Ann:Hull:5', '::', '::']);
  assert.equal(view.calc.getValue('Data', 0, 0), 'Name', 'the headings stay');
  view.undo();
  assert.deepEqual(names(view), ['Ann:Leeds:5', 'Bob:York:7', 'Ann:Leeds:9', 'Ann:Hull:5', 'Bob:York:7'], 'one undo brings them back');
});

test('one column ticked, and the first row said to be data', () => {
  const view = open();
  const said = view.removeDuplicates({ columns: [0] });
  assert.deepEqual({ removed: said.removed, kept: said.kept }, { removed: 3, kept: 2 });
  const plain = open();
  // Without headings the first row is a row like any other: 6 rows, 2 names and the word Name.
  const all = plain.removeDuplicates({ columns: [0], header: false });
  assert.deepEqual({ removed: all.removed, kept: all.kept }, { removed: 3, kept: 3 });
});

test('no column ticked is refused; nothing to remove changes nothing', () => {
  const view = open();
  assert.throws(() => view.removeDuplicates({ columns: [] }), /at least one column/);
  assert.throws(() => view.removeDuplicates({ columns: [9] }), /at least one column/);
  const said = view.removeDuplicates({ columns: [2, 0, 1] });
  assert.equal(said.removed, 1);
  // The list stays selected after a removal; back on one cell, the block and its headings are worked out again.
  view.select(1, 0);
  const again = view.removeDuplicates();
  assert.deepEqual({ removed: again.removed, kept: again.kept }, { removed: 0, kept: 4 });
});

// ---- the status bar ---------------------------------------------------------------

test('cells picked one by one with Ctrl are counted, summed and averaged', () => {
  const view = open();
  view.select(1, 2);
  assert.equal(view.statusLine(), null, 'one cell alone says nothing');
  view.select(4, 2, { add: true });
  assert.deepEqual(view.statusLine(), { count: 2, numeric: 2, sum: 10, average: 5, min: 5, max: 5 });
  view.select(2, 2, { add: true });
  view.select(3, 2, { extend: true });
  const line = view.statusLine();
  assert.equal(line.count, 4);
  assert.equal(line.sum, 26);
  // A cell picked twice is counted once.
  view.select(1, 2, { add: true });
  assert.equal(view.statusLine().sum, 26);
  // A plain click ends the Ctrl selection.
  view.select(1, 2);
  assert.equal(view.statusLine(), null);
});
