// The references in a formula being typed, and where a click puts another.
//
// The owner's report: typing "=" and then pointing at cells — =A1*A2,
// =SUM(B3:B12) — should show each reference as Excel does, its own colour
// in the formula and round its cells, and a click should go where the
// formula's caret is. This is the reading that drives it: what the
// references are, where they sit in the text, their colours, and whether a
// click at the caret inserts, replaces or leaves the cell.
import test from 'node:test';
import assert from 'node:assert/strict';
import { formulaReferences, pointSpan, rangeText, rangeOf, cycleAbsolute, colIndex, colLetters, REF_COLOURS } from '@rutba/sheet-view/formula-refs';

const brief = (f) => formulaReferences(f).map((r) => [r.text, r.start, r.end, r.colour]);

test('each reference, where it is in the text, and its colour', () => {
  assert.deepEqual(brief('=A1*A2'), [['A1', 1, 3, 0], ['A2', 4, 6, 1]]);
  assert.deepEqual(brief('=SUM(B3:B12)'), [['B3:B12', 5, 11, 0]]);
  assert.deepEqual(formulaReferences('=SUM(B3:B12)')[0].range, { top: 2, left: 1, bottom: 11, right: 1 });
  // The same reference twice is one colour; a different one the next.
  assert.deepEqual(brief('=A1+$A$1+B2').map((r) => r[3]), [0, 0, 1]);
  assert.equal(REF_COLOURS.length >= 6, true);
});

test('functions, names, strings and numbers are not references', () => {
  assert.deepEqual(brief('=LOG10(A1)'), [['A1', 7, 9, 0]]);
  assert.deepEqual(brief('=IF(A1="B2",C3,"D4")'), [['A1', 4, 6, 0], ['C3', 12, 14, 1]]);
  assert.deepEqual(brief('=Total*1.5+A1B'), []);
  assert.deepEqual(brief('A1+B2'), [], 'not a formula');
  assert.deepEqual(brief('="unclosed A1'), [], 'inside an unclosed string');
});

test('whole columns, whole rows, other sheets and a half-typed range', () => {
  const cols = formulaReferences('=SUM(A:C)')[0];
  assert.deepEqual([cols.text, cols.range.left, cols.range.right, cols.range.top], ['A:C', 0, 2, 0]);
  const rows = formulaReferences('=SUM(2:4)')[0];
  assert.deepEqual([rows.text, rows.range.top, rows.range.bottom, rows.range.left], ['2:4', 1, 3, 0]);
  const other = formulaReferences("=Sheet2!A1+'My data'!B2:C3");
  assert.deepEqual(other.map((r) => [r.text, r.sheet]), [['Sheet2!A1', 'Sheet2'], ["'My data'!B2:C3", 'My data']]);
  assert.deepEqual(brief('=SUM(A1:'), [['A1', 5, 7, 0]], 'a range still being typed keeps its first cell');
});

test('a click goes in after "=", a bracket, a comma or an operator, and replaces a reference just put in', () => {
  assert.deepEqual(pointSpan('=', 1), { start: 1, end: 1 });
  assert.deepEqual(pointSpan('=SUM(', 5), { start: 5, end: 5 });
  assert.deepEqual(pointSpan('=A1*', 4), { start: 4, end: 4 });
  assert.deepEqual(pointSpan('=A1*A2', 6), { start: 4, end: 6 }, 'the reference before the caret is replaced');
  assert.deepEqual(pointSpan('=SUM(B3:B12', 11), { start: 5, end: 11 });
  assert.deepEqual(pointSpan('=SUM(A1:', 8), { start: 8, end: 8 }, 'after a colon: the range\'s far end');
  assert.deepEqual(pointSpan('=SUM(A1, ', 9), { start: 9, end: 9 });
  assert.equal(pointSpan('=SUM', 4), null, 'in the middle of a name, a click leaves the cell');
  assert.equal(pointSpan('=12', 3), null);
  assert.equal(pointSpan('="a', 3), null, 'inside a string');
  assert.equal(pointSpan('hello', 5), null, 'not a formula');
  assert.equal(pointSpan('=A1+', 1), null, 'right after "=", with a reference running on');
});

test('ranges as Excel writes them, and F4\'s turns of the dollar signs', () => {
  assert.equal(rangeText(rangeOf({ row: 11, col: 1 }, { row: 2, col: 1 })), 'B3:B12', 'top-left first, however it was dragged');
  assert.equal(rangeText(rangeOf({ row: 0, col: 0 })), 'A1');
  assert.equal(colLetters(27), 'AB');
  assert.equal(colIndex('AB'), 27);
  assert.equal(cycleAbsolute('A1'), '$A$1');
  assert.equal(cycleAbsolute('$A$1'), 'A$1');
  assert.equal(cycleAbsolute('A$1'), '$A1');
  assert.equal(cycleAbsolute('$A1'), 'A1');
  assert.equal(cycleAbsolute('B3:B12'), '$B$3:$B$12');
  assert.equal(cycleAbsolute('Sheet2!C4'), 'Sheet2!$C$4');
});
