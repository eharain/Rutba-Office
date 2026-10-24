// Data Files' report Totals (2.18, Part A2 of the plan): a report's totals
// chosen for each number shown, whole numbers and a query's amounts too.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Base, totalFields, withTotal } from '../packages/database/src/index.js';

const shop = () => {
  const b = Base.open(':memory:');
  b.createTable({ name: 'Products', fields: [{ name: 'Name', type: 'text' }, { name: 'Price', type: 'currency' }, { name: 'In Stock', type: 'integer' }, { name: 'Notes', type: 'longtext' }, { name: 'Sold Out', type: 'yesno' }] });
  b.createTable({ name: 'Sales', fields: [{ name: 'Sold On', type: 'date' }, { name: 'Payment', type: 'text' }] });
  b.createTable({ name: 'Sale Lines', fields: [{ name: 'Sale', type: 'integer' }, { name: 'Product', type: 'integer' }, { name: 'Quantity', type: 'integer' }] });
  b.addRelationship({ from: { table: 'Sale Lines', field: 'Product' }, to: { table: 'Products', field: 'ID' } });
  b.addRelationship({ from: { table: 'Sale Lines', field: 'Sale' }, to: { table: 'Sales', field: 'ID' } });
  for (const [n, p, s] of [['Milk', 1.1, 3], ['Bread', 1.4, 10], ['Tea', 3, 4]]) b.insert('Products', { Name: n, Price: p, 'In Stock': s });
  for (const d of ['2026-10-21', '2026-10-22']) b.insert('Sales', { 'Sold On': d, Payment: 'Cash' });
  for (const [s, p, q] of [[1, 1, 2], [1, 2, 1], [2, 3, 5]]) b.insert('Sale Lines', { Sale: s, Product: p, Quantity: q });
  return b;
};

test("a report's totals are chosen for each number: whole numbers too, and a query's amounts; one each, in the columns' order", () => {
  const b = shop();
  const report = b.makeReport({ table: 'Products' });
  // The wizard sums the amounts; a whole number is offered, not summed by itself.
  assert.deepEqual(report.totals, [{ field: 'Price', fn: 'sum' }]);
  const data = b.reportData({ name: 'Products' });
  assert.deepEqual(totalFields(report, data.fields).map((f) => [f.field, f.type]), [['Price', 'currency'], ['In Stock', 'integer']]);
  let l = withTotal(report, 'In Stock', 'avg');
  l = withTotal(l, 'Price', 'count');
  assert.deepEqual(l.totals, [{ field: 'Price', fn: 'count' }, { field: 'In Stock', fn: 'avg' }]);
  b.saveReport(l);
  const totals = b.reportData({ name: 'Products' }).totals;
  assert.equal(totals['count:Price'], 3);
  assert.ok(Math.abs(totals['avg:In Stock'] - 17 / 3) < 1e-9);
  // None takes the total off.
  assert.deepEqual(withTotal(l, 'Price', '').totals, [{ field: 'In Stock', fn: 'avg' }]);
  assert.throws(() => withTotal(l, 'Price', 'median'), /Sum, Average or Count/);
  // The field grouped by, and one pointing at another table's record, have no total.
  const lines = b.makeReport({ table: 'Sale Lines' });
  assert.deepEqual(totalFields(lines, b.reportData({ name: 'Sale Lines' }).fields).map((f) => f.field), ['Quantity']);
  // A report of a query: its columns by their types.
  b.saveQuery('Stock', { tables: ['Products'], columns: [{ table: 'Products', field: 'Name' }, { table: 'Products', field: 'In Stock' }, { table: 'Products', field: 'Price' }] });
  const q = b.makeReport({ query: 'Stock' });
  assert.deepEqual(totalFields(q, b.reportData({ name: 'Stock' }).fields).map((f) => [f.field, f.type]), [['In Stock', 'integer'], ['Price', 'currency']]);
  b.saveReport(withTotal(q, 'In Stock', 'sum'));
  assert.equal(b.reportData({ name: 'Stock' }).totals['sum:In Stock'], 17);
});
