// Data Files' reports (packages/database): the report wizard's layout for a
// table and for a query, grouped by the field that points at another table;
// the records grouped, sorted and totalled, empty values left out of a total;
// a pointing field shown by the related record's words; and reports kept in
// the file, read back after it closes, a version 2 file given their table.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { Base, wizardReport, reportData, total, DESIGN_VERSION } from '../packages/database/src/index.js';

function shop(file = ':memory:') {
  const b = Base.open(file);
  b.createTable({ name: 'Customers', fields: [{ name: 'Name', type: 'text' }, { name: 'City', type: 'text' }] });
  b.createTable({ name: 'Orders', fields: [{ name: 'Customer', type: 'integer' }, { name: 'Item', type: 'text' }, { name: 'Amount', type: 'currency' }] });
  b.addRelationship({ from: { table: 'Orders', field: 'Customer' }, to: { table: 'Customers', field: 'ID' } });
  b.insert('Customers', { Name: 'Ben', City: 'London' });
  b.insert('Customers', { Name: 'Amira', City: 'Lahore' });
  for (const [c, i, a] of [[1, 'pen', 2], [2, 'ink', 5], [2, 'pad', 3], [1, 'cap', null]]) b.insert('Orders', { Customer: c, Item: i, Amount: a });
  return b;
}

test('a total leaves empty values out: an average of 10 and nothing is 10, a count counts what is there', () => {
  assert.equal(total('sum', [2, null, 3, '']), 5);
  assert.equal(total('avg', [10, null]), 10);
  assert.equal(total('count', [10, null, '', 0]), 2);
  assert.equal(total('min', [4, 2, null]), 2);
  assert.equal(total('max', [4, 2, null]), 4);
  assert.equal(total('sum', [null, '']), null);
});

test('the report wizard: grouped by the field that points at another table, amounts on the right and summed', () => {
  const b = shop();
  try {
    const l = wizardReport({ table: 'Orders', fields: b.fields('Orders'), relationships: b.relationships() });
    assert.equal(l.group, 'Customer');
    assert.deepEqual(l.columns.map((c) => [c.field, c.align]), [['Item', 'left'], ['Amount', 'right']]);
    assert.deepEqual(l.totals, [{ field: 'Amount', fn: 'sum' }]);
    assert.deepEqual(l.source, { table: 'Orders' });
    const q = wizardReport({ query: 'Big', columns: ['Item', 'Amount'] });
    assert.equal(q.group, null);
    assert.deepEqual(q.source, { query: 'Big' });
    assert.throws(() => wizardReport({ query: 'None', columns: [] }), /nothing to put/);
  } finally {
    b.close();
  }
});

test("a report's data: groups in order by the related record's words, each with its totals and the whole's", () => {
  const b = shop();
  try {
    b.makeReport({ table: 'Orders' });
    const d = b.reportData({ name: 'Orders' });
    assert.deepEqual(d.groups.map((g) => g.label), ['Amira', 'Ben']);
    assert.deepEqual(d.groups.map((g) => g.rows.map((r) => r.Item)), [['ink', 'pad'], ['pen', 'cap']]);
    assert.deepEqual(d.groups.map((g) => g.totals['sum:Amount']), [8, 2]);
    assert.equal(d.groups[0].rows[0].Customer, 'Amira');
    assert.deepEqual(d.totals, { 'sum:Amount': 10 });
    assert.equal(d.count, 4);
    // Sorted within each group, largest first.
    const sorted = b.reportData({ layout: { ...d.layout, sort: 'Amount', desc: true } });
    assert.deepEqual(sorted.groups.map((g) => g.rows.map((r) => r.Item)), [['ink', 'pad'], ['pen', 'cap']]);
    const flat = reportData({ columns: [], group: null, totals: [{ field: 'n', fn: 'avg' }], sort: 'n' }, [{ n: 3 }, { n: 1 }, { n: null }]);
    assert.deepEqual(flat.groups[0].rows.map((r) => r.n), [1, 3, null]);
    assert.equal(flat.totals['avg:n'], 2);
  } finally {
    b.close();
  }
});

test('a report on a saved query, and a report whose source is not there refused', () => {
  const b = shop();
  try {
    b.saveQuery('Big', { tables: ['Orders'], columns: [{ table: 'Orders', field: 'Item' }, { table: 'Orders', field: 'Amount', criteria: ['> 2'] }] });
    const l = b.makeReport({ query: 'Big' });
    assert.deepEqual(b.reportData({ name: l.name }).groups[0].rows.map((r) => r.Item).sort(), ['ink', 'pad']);
    assert.throws(() => b.saveReport({ name: 'Lost', source: { table: 'Nowhere' }, columns: [] }), /not in this database/);
    assert.throws(() => b.reportData({ name: 'Missing' }), /no report called/);
  } finally {
    b.close();
  }
});

test('reports are kept in the file under names of their own, removed by name, and read back after it closes', () => {
  const file = path.join(os.tmpdir(), `reports-${process.pid}.rbase`);
  try {
    const b = shop(file);
    assert.equal(b.makeReport({ table: 'Orders' }).name, 'Orders');
    assert.equal(b.makeReport({ table: 'Orders' }).name, 'Orders 2');
    b.saveReport({ ...b.reports()[0].layout, title: 'Orders by customer' });
    assert.ok(b.removeReport('Orders 2'));
    b.close();
    const again = Base.open(file);
    assert.deepEqual(again.reports().map((r) => [r.name, r.layout.title]), [['Orders', 'Orders by customer']]);
    again.close();
  } finally {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  }
});

test('a file made before reports opens with their table, and its version brought up', () => {
  const file = path.join(os.tmpdir(), `v2-${process.pid}.rbase`);
  try {
    const old = new DatabaseSync(file);
    old.exec("CREATE TABLE rutba_meta (key TEXT PRIMARY KEY, value TEXT); INSERT INTO rutba_meta VALUES ('version', '2');");
    old.close();
    const b = Base.open(file);
    assert.equal(b.meta('version'), String(DESIGN_VERSION));
    assert.ok(DESIGN_VERSION >= 3);
    assert.deepEqual(b.reports(), []);
    b.close();
  } finally {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  }
});
