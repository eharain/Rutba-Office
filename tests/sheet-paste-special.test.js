/**
 * Worksheets: Cut and paste moving the cells, Paste Special — what is
 * pasted, an operation over what is there, skipping blanks, transposing —
 * the fill handle's Auto Fill Options and Home → Fill → Series.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const open = (rows) => new SheetView(buildXlsx({ sheets: [{ name: 'S', rows }] }));
const text = (view, row, col) => view.displayValue(row, col).text;
const input = (view, row, col) => view.editValue(row, col);
const pick = (view, top, left, bottom = top, right = left) => {
  view.select(top, left);
  if (bottom !== top || right !== left) view.select(bottom, right, { extend: true });
};

test('Cut and paste moves the cells: inputs as typed, a formula among them reading the moved cells where they went, the old place left empty, in one undo step', () => {
  const view = open([[1, 2, '=A1+B1'], [], []]);
  pick(view, 0, 0, 0, 2);
  view.markClipboard({ cut: true });
  pick(view, 2, 1);
  view.pasteText(view.clipboard.text);
  assert.deepEqual([input(view, 2, 1), input(view, 2, 2), input(view, 2, 3)], ['1', '2', '=B3+C3'], 'the formula reads the cells it read, where they went');
  assert.deepEqual([0, 1, 2].map((c) => input(view, 0, c)), ['', '', '']);
  assert.equal(view.clipboard, null, 'a cut is pasted once');
  view.undo();
  assert.deepEqual([0, 1, 2].map((c) => input(view, 0, c)), ['1', '2', '=A1+B1']);
  assert.equal(input(view, 2, 1), '');
});

test('A cut pasted re-points every reference to the moved cells, on any sheet and in the names, as Excel does; a cut pastes on another sheet; one undo puts it all back', () => {
  const book = buildXlsx({ sheets: [{ name: 'S', rows: [[1, 2, '=A1+B1'], [], [], ['=SUM(A1:B1)', '=A1*10', '=A1:A2']] }, { name: 'Other sheet', rows: [['=S!A1+S!B1', '=S!$A$1']] }], definedNames: [{ name: 'Pair', ref: 'S!$A$1:$B$1' }] });
  const view = new SheetView(book);
  pick(view, 0, 0, 0, 1);
  view.markClipboard({ cut: true });
  pick(view, 5, 3);
  view.pasteText(view.clipboard.text);
  assert.deepEqual([input(view, 5, 3), input(view, 5, 4)], ['1', '2']);
  assert.equal(input(view, 0, 2), '=D6+E6', 'a formula beside the cells follows them');
  assert.equal(input(view, 3, 0), '=SUM(D6:E6)', 'a range of only moved cells follows them');
  assert.equal(input(view, 3, 1), '=D6*10');
  assert.equal(input(view, 3, 2), '=A1:A2', 'a range only partly moved stays');
  assert.equal(text(view, 0, 2), '3', 'and still works out');
  view.selectSheet('Other sheet');
  assert.equal(input(view, 0, 0), '=S!D6+S!E6', 'a formula on another sheet follows them');
  assert.equal(input(view, 0, 1), '=S!$D$6', 'dollar signs and all');
  assert.ok(view.pkg.text(view.workbook.mainPart).includes('<definedName name="Pair">S!$D$6:$E$6</definedName>'), 'the name follows them');
  view.undo();
  assert.equal(view.activeSheet, 'S', 'undo goes back to the sheet the cells were moved on');
  assert.deepEqual([input(view, 0, 0), input(view, 0, 2), input(view, 3, 1)], ['1', '=A1+B1', '=A1*10']);
  view.selectSheet('Other sheet');
  assert.equal(input(view, 0, 0), '=S!A1+S!B1', 'one undo puts the other sheet back too');
  view.selectSheet('S');
  assert.ok(view.pkg.text(view.workbook.mainPart).includes('S!$A$1:$B$1'), 'and the name');
  // Onto another sheet: the cells go there, what reads them names that sheet.
  pick(view, 0, 0, 0, 1);
  view.markClipboard({ cut: true });
  view.selectSheet('Other sheet');
  view.select(4, 0);
  view.pasteText(view.clipboard.text);
  assert.deepEqual([input(view, 4, 0), input(view, 4, 1), input(view, 0, 0)], ['1', '2', "='Other sheet'!A5+'Other sheet'!B5"]);
  view.selectSheet('S');
  assert.deepEqual([input(view, 0, 0), input(view, 0, 2)], ['', "='Other sheet'!A5+'Other sheet'!B5"]);
  assert.equal(text(view, 0, 2), '3');
});

test('Paste Special: formulas, values or formats, and transposed — rows become columns', () => {
  const view = open([['Item', 'Qty'], ['Ink', 4], ['Paper', '=B2*2'], [], [], []]);
  pick(view, 0, 0, 2, 1);
  view.markClipboard();
  pick(view, 0, 3);
  view.pasteSpecial({ what: 'values' });
  assert.equal(input(view, 2, 4), '8', 'the value the formula showed, not the formula');
  pick(view, 0, 6);
  view.pasteSpecial({ what: 'formulas' });
  assert.equal(input(view, 2, 7), '=H2*2', 'the formula moved with its cell, six columns over');
  pick(view, 4, 0);
  view.pasteSpecial({ transpose: true });
  assert.deepEqual([0, 1, 2].map((c) => text(view, 4, c)), ['Item', 'Ink', 'Paper'], 'the first column is now the first row');
  assert.deepEqual([0, 1].map((c) => text(view, 5, c)), ['Qty', '4']);
  assert.equal(input(view, 5, 2), '=B6*2', 'the formula turned with the block: the cell above it is now the cell to its left');
  assert.deepEqual(view.selection.range, { top: 4, left: 0, bottom: 5, right: 2 }, 'the pasted block selected, two rows by three');
});

test('Paste Special operations work out the copy against what is there: numbers, formulas, words and blanks, as Excel does', () => {
  const view = open([[10, 'x', '', 3, 5], [], [100, 100, 100, '=5*2', 'words'], [100, 100, 100, 7]]);
  pick(view, 0, 0, 0, 4);
  view.markClipboard();
  pick(view, 2, 0);
  view.pasteSpecial({ what: 'values', operation: 'add' });
  assert.deepEqual([0, 1, 2, 3, 4].map((c) => input(view, 2, c)), ['110', 'x', '100', '=(5*2)+3', 'words'], 'a number added, copied words pasted, a blank adds nothing, a formula joined, words under a number left as they are');
  assert.equal(text(view, 2, 3), '13');
  pick(view, 3, 0);
  view.pasteSpecial({ what: 'values', operation: 'multiply', skipBlanks: true });
  assert.deepEqual([0, 2, 3].map((c) => input(view, 3, c)), ['1000', '100', '21'], 'skipping blanks leaves the cell under one');
  view.undo();
  pick(view, 3, 0);
  view.pasteSpecial({ what: 'values', operation: 'multiply' });
  assert.equal(input(view, 3, 2), '0', 'a blank counts as nought when not skipped');
  pick(view, 3, 0);
  view.markClipboard();
  pick(view, 0, 0);
  view.pasteSpecial({ what: 'values', operation: 'divide', text: view.clipboard.text });
  assert.equal(input(view, 0, 0), '0.01', '10 / 1000');
  assert.throws(() => view.pasteSpecial({ what: 'everything' }), /all, formulas, values or formats/);
});

test('Paste Special from another application pastes its values, transposed when asked; a cut refuses it', () => {
  const view = open([[]]);
  view.select(0, 0);
  view.pasteSpecial({ text: 'a\tb\tc\n1\t2\t3', transpose: true });
  assert.deepEqual([0, 1, 2].map((r) => text(view, r, 0) + text(view, r, 1)), ['a1', 'b2', 'c3']);
  view.markClipboard({ cut: true });
  assert.throws(() => view.pasteSpecial({}), /moved with Paste/);
  view.clipboard = null;
  assert.throws(() => view.pasteSpecial({}), /nothing copied/);
});

test('The fill handle\'s Auto Fill Options: a series where it copied, a copy where it made a series, the formatting or the contents alone', () => {
  const view = open([[1], [], [], []]);
  pick(view, 0, 0);
  view.fill({ top: 0, left: 0, bottom: 3, right: 0 });
  assert.deepEqual([1, 2, 3].map((r) => input(view, r, 0)), ['1', '1', '1'], 'one number is copied, as in Excel');
  view.refill('series');
  assert.deepEqual([1, 2, 3].map((r) => input(view, r, 0)), ['2', '3', '4'], 'Fill Series counts up by one');
  view.refill('copy');
  assert.deepEqual([1, 2, 3].map((r) => input(view, r, 0)), ['1', '1', '1']);
  view.undo();
  assert.deepEqual([1, 2, 3].map((r) => input(view, r, 0)), ['', '', ''], 'the choice replaced the fill: one undo takes it all away');
  const two = open([[1], [2], [], []]);
  pick(two, 0, 0, 1, 0);
  two.fill({ top: 0, left: 0, bottom: 3, right: 0 }, { mode: 'toggle' });
  assert.deepEqual([2, 3].map((r) => input(two, r, 0)), ['1', '2'], 'Ctrl held: a series is copied instead');
  const one = open([[5], [], []]);
  pick(one, 0, 0);
  one.fill({ top: 0, left: 0, bottom: 2, right: 0 }, { mode: 'toggle' });
  assert.deepEqual([1, 2].map((r) => input(one, r, 0)), ['6', '7'], 'Ctrl held on one number: a series');
  one.refill('formats');
  assert.deepEqual([1, 2].map((r) => input(one, r, 0)), ['', ''], 'the formatting alone leaves the cells empty');
  assert.throws(() => open([[1]]).refill('copy'), /drag the fill handle first/);
});

test('Home → Fill → Series: linear, growth and dates by weekday and month down the selection, or until a stop value', () => {
  const view = open([[2, 1, 45658], [], [], [], []]);
  pick(view, 0, 0, 4, 1);
  view.fillSeries({ direction: 'columns', type: 'linear', step: 3 });
  assert.deepEqual([1, 2, 3, 4].map((r) => input(view, r, 0)), ['5', '8', '11', '14']);
  assert.deepEqual([1, 2, 3, 4].map((r) => input(view, r, 1)), ['4', '7', '10', '13'], 'each column from its own first cell');
  pick(view, 0, 1, 4, 1);
  view.fillSeries({ type: 'growth', step: 2 });
  assert.deepEqual([1, 2, 3, 4].map((r) => input(view, r, 1)), ['2', '4', '8', '16']);
  // 45658 is Wednesday 1 January 2025: the weekdays run Thursday, Friday, then Monday.
  pick(view, 0, 2, 3, 2);
  view.fillSeries({ type: 'date', unit: 'weekday', step: 1 });
  assert.deepEqual([1, 2, 3].map((r) => input(view, r, 2)), ['45659', '45660', '45663']);
  pick(view, 0, 2, 2, 2);
  view.fillSeries({ type: 'date', unit: 'month', step: 1 });
  assert.deepEqual([1, 2].map((r) => input(view, r, 2)), ['45689', '45717'], '1 February and 1 March');
  const stop = open([[10]]);
  stop.select(0, 0);
  stop.fillSeries({ direction: 'rows', step: 10, stop: 45 });
  assert.deepEqual([1, 2, 3, 4].map((c) => input(stop, 0, c)), ['20', '30', '40', ''], 'up to the stop and no further');
  assert.deepEqual(stop.selection.range, { top: 0, left: 0, bottom: 0, right: 3 });
  stop.select(0, 0);
  assert.throws(() => stop.fillSeries({}), /give a stop value/);
});

test('Fill Series refuses a step that never reaches its stop, and passes over the cells a merge hides', () => {
  const still = open([[10]]);
  still.select(0, 0);
  assert.throws(() => still.fillSeries({ step: 0, stop: 50 }), /never reaches the stop/);
  assert.equal(input(still, 1, 0), '', 'nothing written');
  const merged = open([[1], [], [], [], []]);
  pick(merged, 2, 0, 3, 0);
  merged.mergeSelection();
  pick(merged, 0, 0, 4, 0);
  merged.fillSeries({ step: 1 });
  assert.deepEqual([1, 2, 3, 4].map((r) => input(merged, r, 0)), ['2', '3', '', '5'], 'the hidden cell of the merge is left alone');
});

test('the fill handle takes a merge along, a whole one at a time, and undo takes it back', () => {
  const view = open([['Region', 1], [null, 2], [], [], [], [], []]);
  pick(view, 0, 0, 1, 0);
  view.mergeSelection();
  pick(view, 0, 0, 1, 1);
  view.fill({ top: 0, left: 0, bottom: 6, right: 1 });
  const merged = (row) => { const m = view.mergeAt(row, 0); return m ? `${m.top}-${m.bottom}` : null; };
  assert.deepEqual([0, 2, 4, 6].map(merged), ['0-1', '2-3', '4-5', null], 'whole merges only: the last row has no room for one');
  assert.equal(view.mergeAt(2, 1), null, 'the merge is the merge, not the column beside it');
  view.undo();
  assert.equal(view.mergeAt(2, 0), null, 'undone');
  assert.equal(merged(0), '0-1', 'the first one stays');
  pick(view, 0, 0, 1, 1);
  view.fill({ top: 0, left: 0, bottom: 3, right: 1 }, { mode: 'values' });
  assert.equal(view.mergeAt(2, 0), null, 'Fill Without Formatting takes no merge');
});
