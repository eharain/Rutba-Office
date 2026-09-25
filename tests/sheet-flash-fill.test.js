/**
 * Flash Fill: the transformation inferred from an example or two, and the
 * column filled by it — or nothing at all when no transformation makes
 * every example.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { inferProgram, runProgram } from '@rutba/sheet-view/flash-fill';

/** Rows of source text, one example output per given row. */
const fill = (sources, outputs) => {
  const examples = outputs.map((output, i) => ({ sources: sources[i].map((t) => (t === null ? null : { text: t })), output }));
  const program = inferProgram(examples);
  if (!program) return null;
  return sources.map((row) => runProgram(program, row.map((t) => (t === null ? null : { text: t }))));
};

test('A first name out of "Last, First", from one example', () => {
  const rows = [['Lovelace, Ada'], ['Hopper, Grace'], ['Smith, John Paul'], ['Turing, Alan']];
  assert.deepEqual(fill(rows, ['Ada']), ['Ada', 'Grace', 'John Paul', 'Alan'], 'the part after the comma, not the second word');
});

test('Initials, from one example — of every word, or first and last from "Last, First"', () => {
  assert.deepEqual(fill([['Ada Lovelace'], ['Grace Brewster Hopper'], ['Alan Turing']], ['AL']), ['AL', 'GBH', 'AT']);
  assert.deepEqual(fill([['Ada Lovelace'], ['Grace Hopper']], ['A.L.']), ['A.L.', 'G.H.']);
  assert.deepEqual(fill([['Lovelace, Ada'], ['Hopper, Grace']], ['AL']), ['AL', 'GH']);
});

test('A change of case, and a part cut by a delimiter', () => {
  assert.deepEqual(fill([['ada lovelace'], ['grace hopper']], ['Ada Lovelace']), ['Ada Lovelace', 'Grace Hopper']);
  assert.deepEqual(fill([['Ada Lovelace'], ['Grace Hopper']], ['LOVELACE']), ['LOVELACE', 'HOPPER']);
  assert.deepEqual(fill([['ada@analytical.org'], ['grace@navy.mil'], ['alan@bletchley.uk']], ['analytical.org']), ['analytical.org', 'navy.mil', 'bletchley.uk']);
  assert.deepEqual(fill([['Order 1042 shipped'], ['Order 77 held'], ['Order 5 shipped']], ['1042']), ['1042', '77', '5']);
});

test('Parts of two columns joined with words of its own', () => {
  const rows = [['Ada', 'Lovelace'], ['Grace', 'Hopper'], ['Alan', 'Turing']];
  assert.deepEqual(fill(rows, ['Lovelace, A.']), ['Lovelace, A.', 'Hopper, G.', 'Turing, A.']);
  assert.deepEqual(fill(rows, ['Dr Ada LOVELACE']), ['Dr Ada LOVELACE', 'Dr Grace HOPPER', 'Dr Alan TURING']);
});

test('A second example settles what one leaves open', () => {
  const rows = [['2024-03-15'], ['2023-11-02'], ['2021-07-30']];
  // "03" alone could be the second token or the first two digits of the day; with "11" it is the month.
  assert.deepEqual(fill(rows, ['03', '11']), ['03', '11', '07']);
  // One example that no single rule makes on the second row: nothing.
  assert.equal(fill([['Ada Lovelace'], ['Grace Hopper']], ['Ada', 'Navy']), null);
  assert.equal(inferProgram([{ sources: [{ text: 'abc' }], output: 'xyz' }]), null, 'words that come from nowhere are not a pattern');
});

test('A date\'s parts, read from the value under its format', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: [['When', 'Month'], [45366, 'March'], [45597, null], [45300, null]], styles: { 'A2:A4': { numFmt: 'dd/mm/yyyy' } } }] }));
  view.select(1, 1);
  assert.deepEqual(view.flashFill(), { filled: 2, examples: 1 });
  assert.deepEqual([2, 3].map((r) => view.displayValue(r, 1).text), ['November', 'January']);
});

test('In the sheet, one column after another: first names, then initials from the names and the first names', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: [
    ['Name', 'First', 'Initials'], ['Lovelace, Ada', 'Ada', 'AL'], ['Hopper, Grace', null, null], ['Turing, Alan', null, null],
  ] }] }));
  view.select(1, 1);
  view.flashFill();
  view.select(1, 2);
  assert.deepEqual(view.flashFill(), { filled: 2, examples: 1 });
  assert.deepEqual([2, 3].map((r) => view.displayValue(r, 1).text + ' ' + view.displayValue(r, 2).text), ['Grace GH', 'Alan AT']);
});

test('In the sheet: Flash Fill fills the column from the examples typed, skips the heading, and is one undo step', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: [
    ['Name', 'First'], ['Lovelace, Ada', 'Ada'], ['Hopper, Grace', null], ['Turing, Alan', null], ['Babbage, Charles', null],
  ] }] }));
  view.select(1, 1);
  assert.deepEqual(view.flashFill(), { filled: 3, examples: 1 }, 'the heading "First" is not an example');
  assert.deepEqual([2, 3, 4].map((r) => view.displayValue(r, 1).text), ['Grace', 'Alan', 'Charles']);
  view.undo();
  assert.deepEqual([2, 3, 4].map((r) => view.displayValue(r, 1).text), ['', '', '']);

  const none = new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: [['Lovelace, Ada', 'Navy'], ['Hopper, Grace', null]] }] }));
  none.select(0, 1);
  assert.throws(() => none.flashFill(), /no pattern/);
  assert.equal(none.displayValue(1, 1).text, '', 'nothing filled');
  assert.equal(none.canUndo, false, 'and no undo step left behind');
  none.select(1, 1);
  const empty = new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: [['Lovelace, Ada', null], ['Hopper, Grace', null]] }] }));
  empty.select(0, 1);
  assert.throws(() => empty.flashFill(), /Type an example/);
});
