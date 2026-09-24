/**
 * Formulas → Error Checking, in the engine: every cell whose calculated value
 * is an error, plus every cell the recalc engine settled as part of a
 * circular reference (see Spreadsheet#recalculate, which marks those
 * `#CIRCULAR!` rather than leaving them unevaluated).
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const view = () => new SheetView(buildXlsx({ sheets: [
  { name: 'S', rows: [
    [1, 2, '=1/0', '=NOSUCHFN(1)'],
    [4, 5, '=VLOOKUP(9,A1:B2,2,FALSE)', '=SQRT(-1)'],
    ['=B3', '=A3'],
  ] },
  { name: 'Other', rows: [[1, '=1/0']] },
] }));

test('errorCells lists every error cell, in row-major order, with its kind, value and reason', () => {
  const v = view();
  assert.deepEqual(v.errorCells(), [
    { ref: 'C1', row: 0, col: 2, sheet: 'S', formula: '=1/0', value: '#DIV/0!', kind: 'div0', reason: 'Divides by zero' },
    { ref: 'D1', row: 0, col: 3, sheet: 'S', formula: '=NOSUCHFN(1)', value: '#NAME?', kind: 'name', reason: 'Uses a name that is not defined' },
    { ref: 'C2', row: 1, col: 2, sheet: 'S', formula: '=VLOOKUP(9,A1:B2,2,FALSE)', value: '#N/A', kind: 'na', reason: 'A lookup found nothing' },
    { ref: 'D2', row: 1, col: 3, sheet: 'S', formula: '=SQRT(-1)', value: '#NUM!', kind: 'num', reason: 'A number that cannot be represented' },
    { ref: 'A3', row: 2, col: 0, sheet: 'S', formula: '=B3', value: '#CIRCULAR!', kind: 'circular', reason: 'Refers to itself, directly or through other cells' },
    { ref: 'B3', row: 2, col: 1, sheet: 'S', formula: '=A3', value: '#CIRCULAR!', kind: 'circular', reason: 'Refers to itself, directly or through other cells' },
  ]);
});

test('all: true walks every sheet, in workbook order, appending the other sheet\'s errors', () => {
  const v = view();
  const all = v.errorCells({ all: true });
  assert.equal(all.length, 7);
  assert.deepEqual(all.slice(0, 6), v.errorCells());
  assert.deepEqual(all[6], { ref: 'B1', row: 0, col: 1, sheet: 'Other', formula: '=1/0', value: '#DIV/0!', kind: 'div0', reason: 'Divides by zero' });
});

test('fixing a cell drops it from the list on the next call', () => {
  const v = view();
  assert.equal(v.errorCells().length, 6);
  v.setCell(0, 2, 5); // C1: was =1/0, now a plain number
  const after = v.errorCells();
  assert.equal(after.length, 5);
  assert.ok(!after.some((e) => e.ref === 'C1'), 'C1 is fixed and gone');
});

test('a sheet with no errors gives []', () => {
  const buf = buildXlsx({ sheets: [{ name: 'Clean', rows: [[1, 2, '=A1+B1']] }] });
  const v = new SheetView(buf);
  assert.deepEqual(v.errorCells(), []);
  assert.deepEqual(v.errorCells({ all: true }), []);
});
