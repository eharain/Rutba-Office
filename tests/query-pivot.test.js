// Power Query's Pivot Column and Add Conditional Column.
import test from 'node:test';
import assert from 'node:assert/strict';
import { applyStep, describeStep } from '@rutba/sheet-view/queries';

const long = { columns: ['Product', 'Month', 'Units'], rows: [['Pens', 'Jan', 10], ['Pens', 'Feb', 12], ['Ink', 'Jan', 5], ['Ink', 'Mar', 7], ['Pens', 'Jan', 3]] };

test('Pivot Column turns each month into a column, the units added up for each product', () => {
  const wide = applyStep(long, { kind: 'pivotColumn', column: 'Month', values: 'Units', fn: 'sum' });
  assert.deepEqual(wide.columns, ['Product', 'Jan', 'Feb', 'Mar']);
  assert.deepEqual(wide.rows, [['Pens', 13, 12, null], ['Ink', 5, null, 7]]);
  assert.deepEqual(wide.types, { Jan: 'number', Feb: 'number', Mar: 'number' });
  assert.deepEqual(applyStep(long, { kind: 'pivotColumn', column: 'Month', values: 'Units', fn: 'count' }).rows[0], ['Pens', 2, 1, null]);
  assert.throws(() => applyStep(long, { kind: 'pivotColumn', column: 'Month', values: 'Units', fn: 'none' }), /one value a cell/, 'two Pens in January cannot share a cell without adding up');
  assert.match(describeStep({ kind: 'pivotColumn', column: 'Month', values: 'Units' }), /"Month" pivoted, its values from "Units"/);
});

test('Add Conditional Column gives each row the output of the first rule it meets, or the otherwise value', () => {
  const out = applyStep(long, { kind: 'conditionalColumn', name: 'Size', rules: [{ column: 'Units', op: 'greaterOrEqual', value: 10, output: 'Big' }, { column: 'Units', op: 'greaterOrEqual', value: 5, output: 'Middling' }], otherwise: 'Small' });
  assert.deepEqual(out.columns, ['Product', 'Month', 'Units', 'Size']);
  assert.deepEqual(out.rows.map((r) => r[3]), ['Big', 'Big', 'Middling', 'Middling', 'Small']);
  assert.throws(() => applyStep(long, { kind: 'conditionalColumn', rules: [] }), /needs a rule/);
  assert.throws(() => applyStep(long, { kind: 'conditionalColumn', rules: [{ column: 'Units', op: 'sometimes', value: 1, output: 1 }] }), /not a condition/);
});
