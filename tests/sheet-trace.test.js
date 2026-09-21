/**
 * Formulas → Trace Precedents and Trace Dependents, in the engine: the cells a
 * formula reads and the formulas that read a cell, as ranges for the grid to
 * draw arrows between. Names and references to other sheets are counted, not
 * drawn.
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';

const view = () => new SheetView(buildXlsx({ sheets: [
  { name: 'S', rows: [[1, '=A1+A2', '=B1*2', '=SUM(A1:A2)', '=Other!A1+A1'], [2, '', '', '', '']] },
  { name: 'Other', rows: [[10]] },
] }));
const box = (a) => [a.top, a.left, a.bottom, a.right].join(',');

test('precedents: each cell a formula reads, a range as one box, other sheets counted', () => {
  const v = view();
  const b1 = v.traceOf(0, 1, 'precedents');
  assert.equal(b1.kind, 'precedents');
  assert.deepEqual(b1.at, { row: 0, col: 1 });
  assert.deepEqual(b1.arrows.map(box), ['0,0,0,0', '1,0,1,0'], 'A1 and A2, one arrow each');
  assert.equal(b1.elsewhere, 0);

  const d1 = v.traceOf(0, 3, 'precedents');
  assert.deepEqual(d1.arrows.map(box), ['0,0,1,0'], 'SUM(A1:A2) is one box round the range');

  const e1 = v.traceOf(0, 4, 'precedents');
  assert.deepEqual(e1.arrows.map(box), ['0,0,0,0'], 'the reference on this sheet is drawn');
  assert.equal(e1.elsewhere, 1, 'the one on the other sheet is counted');

  assert.deepEqual(v.traceOf(0, 0, 'precedents').arrows, [], 'a plain number reads nothing');
});

test('dependents: every formula on the sheet that reads the cell, through a range too', () => {
  const v = view();
  const a1 = v.traceOf(0, 0, 'dependents');
  assert.equal(a1.kind, 'dependents');
  assert.deepEqual(a1.arrows.map(box).sort(), ['0,1,0,1', '0,3,0,3', '0,4,0,4'].sort(), 'B1, D1 (through A1:A2) and E1 read A1');
  assert.deepEqual(v.traceOf(1, 0, 'dependents').arrows.map(box).sort(), ['0,1,0,1', '0,3,0,3'], 'B1 and D1 read A2');
  assert.deepEqual(v.traceOf(0, 1, 'dependents').arrows.map(box), ['0,2,0,2'], 'C1 reads B1');
  assert.deepEqual(v.traceOf(0, 2, 'dependents').arrows, [], 'nothing reads C1');
});
