/**
 * Power Query → Append Queries and Merge Queries: one table's rows after
 * another's, matched by column name; each row joined to the rows of another
 * table whose key matches — left outer, inner, left anti, full outer — and
 * a query that reads another query's result, never itself.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { runSteps, applyStep, describeStep } from '@rutba/sheet-view/queries';

const ORDERS = { columns: ['Order', 'Customer', 'Total'], rows: [[1, 'C1', 100], [2, 'C2', 250], [3, 'C9', 40], [4, 'c1', 75]] };
const CUSTOMERS = { columns: ['Id', 'Name', 'Total'], rows: [['C1', 'Kim', 9], ['C2', 'Lee', 8], ['C3', 'Ann', 7]] };
const MORE = { columns: ['Order', 'Total', 'Note'], rows: [[5, 60, 'late']] };
const tables = { customers: CUSTOMERS, more: MORE };
const table = (src) => tables[src.name];

test('Append puts the other table\'s rows after these, columns matched by name, new ones added', () => {
  const t = applyStep(ORDERS, { kind: 'appendQuery', with: { kind: 'query', name: 'more' } }, { table });
  assert.deepEqual(t.columns, ['Order', 'Customer', 'Total', 'Note']);
  assert.deepEqual(t.rows[4], [5, null, 60, 'late']);
  assert.deepEqual(t.rows[0], [1, 'C1', 100, null]);
  assert.equal(describeStep({ kind: 'appendQuery', with: { kind: 'table', table: 'More' } }), 'Appended table More');
});

test('Merge joins each row to the matching rows of the other table, in four kinds, keys matched without regard to case', () => {
  const step = (how) => ({ kind: 'mergeQueries', with: { kind: 'query', name: 'customers' }, on: 'Customer', withOn: 'Id', how });
  const left = applyStep(ORDERS, step('left'), { table });
  assert.deepEqual(left.columns, ['Order', 'Customer', 'Total', 'Name', 'Merged.Total'], 'a clashing column named apart');
  assert.deepEqual(left.rows.map((r) => [r[0], r[3]]), [[1, 'Kim'], [2, 'Lee'], [3, null], [4, 'Kim']]);
  assert.deepEqual(applyStep(ORDERS, step('inner'), { table }).rows.map((r) => r[0]), [1, 2, 4]);
  assert.deepEqual(applyStep(ORDERS, step('leftAnti'), { table }).rows, [[3, 'C9', 40]]);
  const full = applyStep(ORDERS, step('full'), { table });
  assert.deepEqual(full.rows[full.rows.length - 1], [null, 'C3', null, 'Ann', 7], 'the unmatched customer, its key in the key column');
  assert.throws(() => applyStep(ORDERS, { ...step('left'), withOn: 'Nope' }, { table }), /no column "Nope"/);
  assert.match(describeStep(step('inner')), /^Merged with query customers on "Customer" = "Id" \(inner\)$/);
});

test('in a workbook a query merges a table, and another query reads its result; one that reads itself is refused', () => {
  const view = new SheetView(buildXlsx({ sheets: [
    { name: 'Orders', rows: [ORDERS.columns, ...ORDERS.rows] },
    { name: 'People', rows: [CUSTOMERS.columns, ...CUSTOMERS.rows] },
  ] }));
  const people = view.addQuery({ name: 'People list', source: { kind: 'range', sheet: 'People', ref: 'A1:C4' }, steps: [{ kind: 'removeColumns', columns: ['Total'] }] });
  const merged = view.addQuery({ name: 'Orders with names', source: { kind: 'range', sheet: 'Orders', ref: 'A1:C5' }, steps: [{ kind: 'mergeQueries', with: { kind: 'query', id: people.id, name: 'People list' }, on: 'Customer', withOn: 'Id', how: 'inner' }] });
  assert.equal(merged.rows, 3);
  const text = (r, c) => view.displayValue(r, c).text;
  assert.deepEqual([0, 1, 2, 3].map((c) => text(0, c)), ['Order', 'Customer', 'Total', 'Name']);
  assert.deepEqual([1, 2, 3].map((r) => text(r, 3)), ['Kim', 'Lee', 'Kim']);
  // A change to the first query's source flows through on Refresh All.
  view.selectSheet('People');
  view.setCell(1, 1, 'Kimberly');
  view.refreshQueries();
  view.selectSheet('Orders with names');
  assert.equal(text(1, 3), 'Kimberly');
  assert.throws(() => view.editQuery({ id: people.id, steps: [{ kind: 'appendQuery', with: { kind: 'query', id: people.id } }] }), /reads itself/);
});
