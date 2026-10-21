// The database app's file (the 2.9 plan): tables with field types the file
// itself enforces, every design change a numbered step applied in place with
// the data kept, relationships as real foreign keys, the person's own SQL run
// read-only, and imports from rows and from the old files' sheets.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Base, importRows, importSheets, inferType, columnSql, DESIGN_VERSION } from '../packages/database/src/index.js';

const fresh = () => Base.open(':memory:');

test('a table gets its AutoNumber, and the file itself holds each field to its type', () => {
  const b = fresh();
  b.createTable({ name: 'Customers', fields: [{ name: 'Name', type: 'text', size: 20, required: true }, { name: 'Joined', type: 'date' }, { name: 'Credit', type: 'currency' }, { name: 'Active', type: 'yesno', default: true }] });
  assert.deepEqual(b.fields('Customers').map((f) => [f.name, f.type]), [['ID', 'autonumber'], ['Name', 'text'], ['Joined', 'date'], ['Credit', 'currency'], ['Active', 'yesno']]);
  const id = b.insert('Customers', { Name: 'Amira', Joined: '2024-01-15', Credit: '1,250.5', Active: 'yes' });
  assert.equal(id, 1);
  assert.deepEqual(b.rows('Customers')[0], { _rowid: 1, ID: 1, Name: 'Amira', Joined: '2024-01-15', Credit: 1250.5, Active: 1 });
  assert.throws(() => b.insert('Customers', { Name: 'x'.repeat(21) }), /does not fit its field/);
  assert.throws(() => b.insert('Customers', { Joined: '2024-01-15' }), /Name needs a value/);
  assert.throws(() => b.insert('Customers', { Name: 'Ben', Joined: '15/01/2024' }), /takes a date, as 2026-10-21/);
  // Even a tool that writes past the app is held to the type.
  assert.throws(() => b.db.prepare('INSERT INTO "Customers" ("Name", "Credit") VALUES (?, ?)').run('Cal', 'lots'), /CHECK constraint failed/);
  assert.match(columnSql({ name: 'A', type: 'yesno' }), /CHECK \("A" IN \(0, 1\)\)/);
});

test('every design change is a numbered step, and a type change keeps the data', () => {
  const b = fresh();
  b.createTable({ name: 'Items', fields: [{ name: 'Name', type: 'text' }, { name: 'Price', type: 'text' }] });
  b.insert('Items', { Name: 'Tea', Price: '3.5' });
  b.insert('Items', { Name: 'Cake', Price: '12' });
  b.changeField('Items', 'Price', { type: 'currency' });
  assert.deepEqual(b.rows('Items').map((r) => r.Price), [3.5, 12]);
  b.changeField('Items', 'Name', { name: 'Item' });
  b.addField('Items', { name: 'Stock', type: 'integer' });
  b.removeField('Items', 'Stock');
  b.renameTable('Items', 'Products');
  assert.deepEqual(b.steps().map((s) => [s.n, s.label]), [
    [1, 'Create table Items'], [2, 'Change field Price in Items'], [3, 'Rename field Name to Item in Items'],
    [4, 'Add field Stock to Items'], [5, 'Remove field Stock from Items'], [6, 'Rename table Items to Products'],
  ]);
  assert.deepEqual(b.rows('Products').map((r) => [r.ID, r.Item, r.Price]), [[1, 'Tea', 3.5], [2, 'Cake', 12]]);
  // A change the data does not fit is undone whole.
  b.insert('Products', { Item: 'Gift card', Price: 0.5 });
  assert.throws(() => b.changeField('Products', 'Item', { size: 3 }));
  assert.equal(b.fields('Products').find((f) => f.name === 'Item').size, undefined);
  assert.equal(b.steps().length, 6);
});

test('an AutoNumber is never renumbered: a deleted record\'s number is not given again, through design changes too', () => {
  const b = fresh();
  b.createTable({ name: 'T', fields: [{ name: 'A', type: 'text' }] });
  b.insert('T', { A: 'one' });
  const two = b.insert('T', { A: 'two' });
  b.remove('T', [two]);
  b.changeField('T', 'A', { type: 'longtext' });
  assert.equal(b.insert('T', { A: 'three' }), 3);
});

test('a relationship is a real foreign key: it refuses records pointing at nothing, and cascades as asked', () => {
  const b = fresh();
  b.createTable({ name: 'Customers', fields: [{ name: 'Name', type: 'text' }] });
  b.createTable({ name: 'Orders', fields: [{ name: 'Customer', type: 'integer' }, { name: 'Total', type: 'currency' }] });
  const amira = b.insert('Customers', { Name: 'Amira' });
  b.insert('Orders', { Customer: amira, Total: 10 });
  const id = b.addRelationship({ from: { table: 'Orders', field: 'Customer' }, to: { table: 'Customers', field: 'ID' }, cascadeDelete: true });
  assert.equal(b.relationships()[0].id, id);
  assert.throws(() => b.insert('Orders', { Customer: 99, Total: 1 }), /related to others/);
  assert.throws(() => b.dropTable('Customers'), /Another table points at Customers/);
  b.remove('Customers', [amira]);
  assert.equal(b.count('Orders'), 0);
  // A relationship the records already break is refused, and nothing changes.
  b.insert('Orders', { Customer: null, Total: 2 });
  b.removeRelationship(id);
  b.insert('Orders', { Customer: 42, Total: 3 });
  assert.throws(() => b.addRelationship({ from: { table: 'Orders', field: 'Customer' }, to: { table: 'Customers', field: 'ID' } }), /pointing at nothing/);
  assert.equal(b.relationships().length, 0);
});

test('the person\'s own SQL runs read-only, one SELECT at a time', () => {
  const b = fresh();
  b.createTable({ name: 'Sales', fields: [{ name: 'Region', type: 'text' }, { name: 'Amount', type: 'number' }] });
  for (const [r, a] of [['North', 10], ['South', 5], ['North', 7]]) b.insert('Sales', { Region: r, Amount: a });
  const { columns, rows } = b.query('SELECT Region, sum(Amount) AS Total FROM Sales GROUP BY Region ORDER BY Region;');
  assert.deepEqual(columns, ['Region', 'Total']);
  assert.deepEqual(rows, [{ Region: 'North', Total: 17 }, { Region: 'South', Total: 5 }]);
  assert.throws(() => b.query('DELETE FROM Sales'), /starts with SELECT/);
  assert.throws(() => b.query('SELECT 1; DELETE FROM Sales'), /One statement/);
  assert.throws(() => b.query("WITH x AS (SELECT 1) INSERT INTO Sales (Region) SELECT 'x'"));
  assert.equal(b.count('Sales'), 3);
  assert.deepEqual(b.rows('Sales', { order: '-Amount', limit: 1 }).map((r) => r.Amount), [10]);
  assert.equal(b.count('Sales', { Region: 'North' }), 2);
});

test('a table imported from rows takes the types its values call for', () => {
  assert.deepEqual(inferType([1, 2, '3']), { type: 'integer' });
  assert.deepEqual(inferType(['1.5', 2]), { type: 'number' });
  assert.deepEqual(inferType(['yes', 'No', true]), { type: 'yesno' });
  assert.deepEqual(inferType(['2024-01-01', '', null]), { type: 'date' });
  assert.deepEqual(inferType([45306, 45307], 'date'), { type: 'date' });
  const b = fresh();
  const made = importRows(b, 'People', [['ID', 'Name', 'Age', 'Member', 'Joined'], [7, 'Amira', 34, 'yes', '2020-05-01'], [9, 'Ben', '41', 'no', '2021-07-09'], [null, '', '', '', '']]);
  assert.equal(made.count, 2);
  assert.deepEqual(made.fields.map((f) => [f.name, f.type]), [['ID (imported)', 'integer'], ['Name', 'text'], ['Age', 'integer'], ['Member', 'yesno'], ['Joined', 'date']]);
  assert.deepEqual(b.rows('People').map((r) => [r.ID, r['ID (imported)'], r.Name, r.Age, r.Member, r.Joined]), [[1, 7, 'Amira', 34, 1, '2020-05-01'], [2, 9, 'Ben', 41, 0, '2021-07-09']]);
  // A second import of the same name is its own table.
  assert.equal(importRows(b, 'People', [['A'], [1]]).table, 'People 2');
});

test('the old files\' sheets import as tables, date-formatted columns as dates', () => {
  const b = fresh();
  const model = { sheets: [{ name: 'TEAS', rows: [['NAME', 'PRICE', 'SOLD'], ['Tea', 3.5, 45306], ['Cake', 12, null]], styles: { A1: { bold: true }, B2: { numFmt: '0.00' }, C2: { numFmt: 'yyyy-mm-dd' } } }], names: [], notes: [] };
  const [made] = importSheets(b, model);
  assert.deepEqual(made.fields.map((f) => [f.name, f.type]), [['NAME', 'text'], ['PRICE', 'number'], ['SOLD', 'date']]);
  assert.deepEqual(b.rows('TEAS').map((r) => [r.NAME, r.PRICE, r.SOLD]), [['Tea', 3.5, '2024-01-15'], ['Cake', 12, null]]);
});

test('the file on disk is one SQLite file any tool reads, its design kept beside the data', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-base-'));
  const file = path.join(dir, 'shop.rbase');
  const b = Base.open(file);
  b.createTable({ name: 'Items', fields: [{ name: 'Name', type: 'text' }] });
  b.insert('Items', { Name: 'Tea' });
  b.close();
  const again = Base.open(file);
  assert.deepEqual(again.tables(), [{ name: 'Items', caption: null }]);
  assert.equal(again.rows('Items')[0].Name, 'Tea');
  assert.equal(again.meta('version'), String(DESIGN_VERSION));
  again.close();
  fs.rmSync(dir, { recursive: true, force: true });
});
