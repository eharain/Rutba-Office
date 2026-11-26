// Data → What-If Analysis from the window.
//
// The Data Table dialog sends the range as the selection's reference
// (D3:H9); the engine wanted its four edges, threw "bad data table range:
// top", and the grid stayed empty. Goal Seek found the right value and the
// dialog then said "B4 is B4": the cell's name where the value belongs.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { goalSeekSaid } from '../apps/desktop/renderer/apps/sheets/what-if.js';

// A loan: B1 the amount, B2 the rate, B3 the interest for a year.
const BOOK = () => buildXlsx({ sheets: [{ name: 'Loan', rows: [
  ['Amount', 1000],
  ['Rate', 0.1],
  ['Interest', '=B1*B2'],
] }] });

test('a Data Table fills from the range as the dialog sends it', () => {
  const view = new SheetView(BOOK());
  // Rates down D2:D4, the formula at E1, as film 603 lays it out.
  view.setCell(0, 4, '=B3');
  view.setCell(1, 3, 0.05);
  view.setCell(2, 3, 0.2);
  view.setCell(3, 3, 0.5);
  const done = view.dataTable({ range: 'D1:E4', colInput: 'B2' });
  assert.equal(done.cells, 3);
  assert.deepEqual([1, 2, 3].map((r) => view.calc.getValue('Loan', r, 4)), [50, 200, 500]);
  assert.equal(view.calc.getValue('Loan', 2, 1), 100, 'the input cell is as it was');

  // Two inputs: amounts across G1:I1, rates down F2:F3, the formula in the corner.
  view.setCell(0, 5, '=B3');
  view.setCell(0, 6, 2000);
  view.setCell(0, 7, 3000);
  view.setCell(1, 5, 0.1);
  view.setCell(2, 5, 0.2);
  view.dataTable({ range: '$F$1:$H$3', rowInput: 'B1', colInput: '$B$2' });
  assert.deepEqual([[1, 6], [1, 7], [2, 6], [2, 7]].map(([r, c]) => view.calc.getValue('Loan', r, c)), [200, 300, 400, 600]);

  assert.throws(() => view.dataTable({ range: 'D1', colInput: 'B2' }), /needs a range of at least two rows and two columns/);
  assert.throws(() => view.dataTable({ range: 'somewhere', colInput: 'B2' }), /needs a range/);
  assert.throws(() => view.dataTable({ range: 'D1:D4', colInput: 'B2' }), /at least two rows and two columns/);
});

test('Goal Seek says the value it found, not the name of the cell', () => {
  const view = new SheetView(BOOK());
  const found = view.goalSeek({ set: 'B3', to: '42', by: 'B1' });
  assert.equal(found.value, 420);
  // What the operation sends back to the window, and what the dialog then says.
  const sent = JSON.stringify({ goalSeek: { converged: true, value: found.value, reached: 42 } });
  assert.deepEqual(goalSeekSaid(sent, '42'), { converged: true, value: '420', reached: '42' });
  assert.deepEqual(goalSeekSaid(JSON.stringify({ goalSeek: { converged: true, value: 247.0588235294, reached: 0 } }), '0'), { converged: true, value: '247.06', reached: '0' });
  assert.equal(goalSeekSaid(JSON.stringify({ goalSeek: { converged: true, value: 1234567.891, reached: 5 } }), '5').value, (1234567.89).toLocaleString(undefined, { maximumFractionDigits: 2 }));
  assert.deepEqual(goalSeekSaid(undefined, '42'), { converged: false, reached: '42' }, 'nothing came back: it did not converge');
  assert.deepEqual(goalSeekSaid('not json', '42'), { converged: false, reached: '42' });
});
