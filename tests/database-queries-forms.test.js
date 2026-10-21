// Data Files' queries and forms (packages/database): a query built in the
// designer as the SQL it stands for, its criteria in Access's words and its
// values always parameters; saved queries and forms kept in the file and
// read back; the form wizard's layout; the lookup lists; and a file made
// before queries and forms opening with their tables added.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { Base, designToSql, criterionSql, wizardForm, labelFor, DESIGN_VERSION } from '../packages/database/src/index.js';

function shop() {
  const b = Base.open(':memory:');
  b.createTable({ name: 'Customers', fields: [{ name: 'Name', type: 'text' }, { name: 'City', type: 'text' }] });
  b.createTable({ name: 'Orders', fields: [{ name: 'Customer', type: 'integer' }, { name: 'Amount', type: 'currency' }, { name: 'When', type: 'date' }, { name: 'Paid', type: 'yesno' }] });
  b.addRelationship({ from: { table: 'Orders', field: 'Customer' }, to: { table: 'Customers', field: 'ID' } });
  b.insert('Customers', { Name: 'Amira', City: 'Lahore' });
  b.insert('Customers', { Name: 'Ben', City: 'London' });
  b.insert('Customers', { Name: "O'Brien", City: 'Leeds' });
  b.insert('Orders', { Customer: 1, Amount: 100, When: '2026-01-05' });
  b.insert('Orders', { Customer: 1, Amount: 250, When: '2026-02-10' });
  b.insert('Orders', { Customer: 2, Amount: 40, When: '2026-03-01' });
  b.insert('Orders', { Customer: 3, Amount: 75, When: '2026-03-02' });
  return b;
}

test('a designed query: criteria in Access words, a join from the relationship, sort and totals, values kept as values', () => {
  const b = shop();
  try {
    const run = (design) => b.runQuery({ design }).rows;
    assert.deepEqual(run({ tables: ['Customers'], columns: [{ table: 'Customers', field: 'Name', sort: 'asc' }, { table: 'Customers', field: 'City', criteria: ['Like "L*"'] }] }).map((r) => r.Name), ['Amira', 'Ben', "O'Brien"]);
    assert.deepEqual(run({ tables: ['Customers'], columns: [{ table: 'Customers', field: 'Name' }, { table: 'Customers', field: 'City', criteria: ['London', 'Leeds'] }] }).map((r) => r.Name).sort(), ['Ben', "O'Brien"]);
    assert.deepEqual(run({ tables: ['Customers', 'Orders'], columns: [{ table: 'Customers', field: 'Name', criteria: ['"O\'Brien"'] }, { table: 'Orders', field: 'Amount' }] }), [{ Name: "O'Brien", Amount: 75 }]);
    assert.deepEqual(run({ tables: ['Orders', 'Customers'], columns: [{ table: 'Customers', field: 'Name' }, { table: 'Orders', field: 'When', criteria: ['Between #2026-02-01# And #2026-03-01#'], sort: 'asc' }] }).map((r) => r.Name), ['Amira', 'Ben']);
    assert.deepEqual(run({ tables: ['Customers', 'Orders'], columns: [{ table: 'Customers', field: 'Name', total: 'group', sort: 'asc' }, { table: 'Orders', field: 'Amount', total: 'sum', criteria: ['> 50'] }] }), [{ Name: 'Amira', 'Sum of Amount': 350 }, { Name: "O'Brien", 'Sum of Amount': 75 }]);
    assert.equal(run({ tables: ['Customers'], columns: [{ table: 'Customers', field: 'Name', criteria: ['In ("Ben", "Amira")'] }] }).length, 2);
    assert.equal(run({ tables: ['Customers'], columns: [{ table: 'Customers', field: 'City', criteria: ['Is Not Null'] }] }).length, 3);
    assert.equal(run({ tables: ['Customers'], columns: [{ table: 'Customers', field: 'Name', criteria: ['Not "Ben"'] }] }).length, 2);
    assert.equal(run({ tables: ['Customers'], columns: [{ table: 'Customers', field: 'Name', criteria: ["x' OR 1=1 --"] }] }).length, 0);
    b.createTable({ name: 'Notes', fields: [{ name: 'Text', type: 'text' }] });
    assert.throws(() => designToSql({ tables: ['Customers', 'Notes'], columns: [{ table: 'Notes', field: 'Text' }] }, b.relationships()), /not related/);
    assert.equal(criterionSql('"x"', 'Like "50%*"').params[0], '50\\%%');
    assert.match(b.sqlOf({ tables: ['Customers'], columns: [{ table: 'Customers', field: 'Name', criteria: ['"O\'Brien"'] }] }), /= 'O''Brien'/);
  } finally {
    b.close();
  }
});

test('saved queries and forms are kept in the file, run by name, and read back after it closes', () => {
  const file = path.join(os.tmpdir(), `qf-${process.pid}.rbase`);
  try {
    let b = shop();
    b.close();
    b = Base.open(file);
    b.createTable({ name: 'Customers', fields: [{ name: 'Name', type: 'text' }] });
    b.insert('Customers', { Name: 'Amira' });
    b.saveQuery('Everyone', { tables: ['Customers'], columns: [{ table: 'Customers', field: 'Name' }] });
    assert.throws(() => b.saveQuery('Customers', { tables: ['Customers'], columns: [{ table: 'Customers', field: 'Name' }] }), /table's name/);
    const form = b.makeForm('Customers');
    assert.equal(form.name, 'Customers');
    assert.equal(b.makeForm('Customers').name, 'Customers 2');
    b.close();
    b = Base.open(file);
    assert.deepEqual(b.queries().map((x) => x.name), ['Everyone']);
    assert.deepEqual(b.runQuery({ name: 'Everyone' }).rows, [{ Name: 'Amira' }]);
    assert.deepEqual(b.forms().map((f) => f.name), ['Customers', 'Customers 2']);
    assert.ok(b.steps().some((s) => /Query Everyone kept/.test(s.label)));
    b.removeQuery('Everyone');
    assert.equal(b.queries().length, 0);
    b.close();
  } finally {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  }
});

test('the form wizard: a control for each field by its type, a lookup as a list of the related records, related records in a subform', () => {
  const b = shop();
  try {
    const orders = wizardForm('Orders', (t) => b.fields(t), b.relationships());
    assert.deepEqual(orders.controls.map((c) => [c.kind, c.label]), [['auto', 'ID'], ['list', 'Customer'], ['currency', 'Amount'], ['date', 'When'], ['check', 'Paid']]);
    assert.deepEqual(orders.controls[1].lookup, { table: 'Customers', field: 'ID', show: 'Name' });
    assert.deepEqual(b.lookup(orders.controls[1].lookup).map((x) => x.text), ['Amira', 'Ben', "O'Brien"]);
    const customers = wizardForm('Customers', (t) => b.fields(t), b.relationships());
    assert.deepEqual(customers.subforms.map((s) => [s.table, s.field, s.to]), [['Orders', 'Customer', 'ID']]);
    assert.ok(customers.controls.every((c, i, all) => i === 0 || c.y > all[i - 1].y), 'one below the other');
    assert.deepEqual([labelFor('ID'), labelFor('first_name'), labelFor('CustomerID'), labelFor('VAT_number')], ['ID', 'First name', 'Customer ID', 'VAT number']);
  } finally {
    b.close();
  }
});

test('a file made before saved queries and forms opens with their tables, and its version brought up', () => {
  const file = path.join(os.tmpdir(), `v1-${process.pid}.rbase`);
  try {
    const old = new DatabaseSync(file);
    old.exec("CREATE TABLE rutba_meta (key TEXT PRIMARY KEY, value TEXT); INSERT INTO rutba_meta VALUES ('version', '1');");
    old.close();
    const b = Base.open(file);
    assert.equal(b.meta('version'), String(DESIGN_VERSION));
    assert.deepEqual(b.queries(), []);
    assert.deepEqual(b.forms(), []);
    b.close();
  } finally {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  }
});
