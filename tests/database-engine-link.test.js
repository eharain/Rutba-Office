// Tables linked to a database engine (D-2026-10-22-13): the engine holds the
// records, Data Files reads them into the window's own copy and writes each
// change to the engine. Run against a real PostgreSQL named by
// OFFICE_PG_TEST_URL (postgres://user:password@host:port/database, a
// database the account owns); skipped without one.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import pg from 'pg';
import { Base } from '../packages/database/src/index.js';
import { createBaseService } from '../apps/desktop/main/base.js';
import { fieldType, fromEngine } from '../apps/desktop/main/engines/postgres.js';

const URL_ = process.env.OFFICE_PG_TEST_URL;
const skip = URL_ ? false : 'no OFFICE_PG_TEST_URL';

test('PostgreSQL types read as Data Files fields, and values as it keeps them', () => {
  assert.equal(fieldType({ type: 'integer', identity: true }), 'autonumber');
  assert.equal(fieldType({ type: 'bigint' }), 'integer');
  assert.equal(fieldType({ type: 'numeric' }), 'number');
  assert.equal(fieldType({ type: 'money' }), 'currency');
  assert.equal(fieldType({ type: 'boolean' }), 'yesno');
  assert.equal(fieldType({ type: 'timestamp without time zone' }), 'datetime');
  assert.equal(fieldType({ type: 'character varying' }), 'text');
  assert.equal(fieldType({ type: 'text' }), 'longtext');
  assert.equal(fromEngine('number', '12.50'), 12.5);
  assert.equal(fromEngine('currency', '$1,234.50'), 1234.5);
  assert.equal(fromEngine('yesno', true), 1);
  assert.equal(fromEngine('date', '2026-10-22'), '2026-10-22');
  assert.equal(fromEngine('datetime', '2026-10-22 09:30:00'), '2026-10-22T09:30:00');
  assert.equal(fromEngine('datetime', '2026-10-22 09:30:00.5+05'), '2026-10-22T09:30:00');
});

const settingsOf = (u) => ({ host: u.hostname, port: Number(u.port) || 5432, database: u.pathname.slice(1), user: decodeURIComponent(u.username), ssl: false });

test('a PostgreSQL table linked live: read into the window, changed in the engine, joined with a local table, unlinked', { skip }, async () => {
  const u = new URL(URL_);
  const password = decodeURIComponent(u.password);
  const pgc = new pg.Client({ connectionString: URL_ });
  await pgc.connect();
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'engine-link-'));
  const file = path.join(dir, 'shop.rbase');
  const kept = new Map();
  const secrets = { available: () => true, get: (k) => kept.get(k) ?? null, set: (k, v) => kept.set(k, v), delete: (k) => kept.delete(k) };
  const svc = createBaseService({ secrets });
  try {
    await pgc.query('DROP SCHEMA IF EXISTS datafiles_link_test CASCADE');
    await pgc.query('CREATE SCHEMA datafiles_link_test');
    await pgc.query(`CREATE TABLE datafiles_link_test.products (id serial PRIMARY KEY, name varchar(80) NOT NULL, price numeric(10,2), active boolean DEFAULT true, added date)`);
    await pgc.query(`INSERT INTO datafiles_link_test.products (name, price, active, added) VALUES ('Bread', 2.50, true, '2026-10-01'), ('Milk', 1.20, false, '2026-10-02')`);

    // A local table of sales that points at the engine's products by number.
    const b = Base.open(file);
    b.createTable({ name: 'Sales', fields: [{ name: 'Product', type: 'integer' }, { name: 'Quantity', type: 'integer' }] });
    b.insert('Sales', { Product: 1, Quantity: 3 });
    b.insert('Sales', { Product: 2, Quantity: 1 });
    b.close();

    const opened = await svc.open({ file });
    const id = opened.id;
    const found = await svc.engineTables({ engine: 'postgres', settings: settingsOf(u), password });
    assert.ok(found.tables.some((t) => t.schema === 'datafiles_link_test' && t.name === 'products'));
    await assert.rejects(svc.engineTables({ engine: 'postgres', settings: settingsOf(u), password: 'not-it' }), /did not take that account/);

    const { link } = await svc.saveLink({ id, name: 'Head office', engine: 'postgres', settings: settingsOf(u), password });
    assert.equal(kept.size, 1);
    // The password is the computer's to keep, not the file's.
    const raw = fs.readFileSync(file);
    assert.ok(!raw.includes(Buffer.from(password)));

    const linked = await svc.linkTables({ id, link, remotes: [{ schema: 'datafiles_link_test', name: 'products', as: 'Products' }] });
    const products = linked.summary.tables.find((t) => t.name === 'Products');
    assert.equal(products.linked, link);
    assert.equal(products.count, 2);
    assert.equal(products.key, 'id');
    assert.deepEqual(svc.fields({ id, table: 'Products' }).map((f) => [f.name, f.type]), [['id', 'autonumber'], ['name', 'text'], ['price', 'number'], ['active', 'yesno'], ['added', 'date']]);
    assert.deepEqual(svc.rows({ id, table: 'Products' }).rows.map((r) => [r.name, r.price, r.active, r.added]), [['Bread', 2.5, 1, '2026-10-01'], ['Milk', 1.2, 0, '2026-10-02']]);

    // Changes go to the engine, and come back.
    const added = await svc.insert({ id, table: 'Products', values: { name: 'Eggs', price: '3.10', active: true, added: '2026-10-22' } });
    assert.ok(added.rowid);
    await svc.update({ id, table: 'Products', rowid: 1, values: { price: '2.75' } });
    await svc.remove({ id, table: 'Products', rowids: [2] });
    const there = (await pgc.query('SELECT name, price::float AS price, active FROM datafiles_link_test.products ORDER BY id')).rows;
    assert.deepEqual(there, [{ name: 'Bread', price: 2.75, active: true }, { name: 'Eggs', price: 3.1, active: true }]);
    assert.deepEqual(svc.rows({ id, table: 'Products' }).rows.map((r) => r.name), ['Bread', 'Eggs']);

    // Changed in the engine by someone else: read again.
    await pgc.query(`UPDATE datafiles_link_test.products SET name = 'Loaf' WHERE id = 1`);
    await svc.refreshLinked({ id });
    assert.equal(svc.rows({ id, table: 'Products' }).rows[0].name, 'Loaf');

    // A query joins the engine's table with the file's own, as any two tables.
    svc.addRelationship({ id, from: { table: 'Sales', field: 'Product' }, to: { table: 'Products', field: 'id' } });
    const out = svc.runQuery({ id, design: { tables: ['Sales', 'Products'], columns: [{ table: 'Products', field: 'name' }, { table: 'Sales', field: 'Quantity' }] } });
    assert.ok(out.rows.length >= 1);
    assert.equal(out.rows.find((r) => r.name === 'Loaf')?.Quantity, 3);

    // Its design is the engine's.
    assert.throws(() => svc.addField({ id, table: 'Products', field: { name: 'Note', type: 'text' } }), /engine/);

    // Opened again: linked, and read again from the engine.
    svc.close({ id });
    const again = await svc.open({ file });
    assert.equal(again.tables.find((t) => t.name === 'Products').count, 2);
    assert.equal(again.tables.find((t) => t.name === 'Products').linkError, null);
    // The file itself keeps no records of the engine's.
    const file2 = Base.read(file);
    try {
      assert.equal(file2.db.prepare("SELECT count(*) AS n FROM sqlite_master WHERE name = 'Products'").get().n, 0);
    } finally { file2.close(); }

    // Unlinked: the engine keeps its records.
    const after = svc.unlinkTable({ id: again.id, table: 'Products' });
    assert.ok(!after.tables.some((t) => t.name === 'Products'));
    assert.equal((await pgc.query('SELECT count(*)::int AS n FROM datafiles_link_test.products')).rows[0].n, 2);
    svc.close({ id: again.id });
  } finally {
    svc.closeAll();
    await pgc.query('DROP SCHEMA IF EXISTS datafiles_link_test CASCADE').catch(() => {});
    await pgc.end();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('an engine that cannot be reached: the table is listed with why, and the rest of the database opens', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'engine-down-'));
  const file = path.join(dir, 'shop.rbase');
  let tried = 0;
  const svc = createBaseService({ secrets: null, connect: async () => { tried++; throw new Error('The engine at db.example:5432 could not be reached.'); } });
  try {
    const b = Base.open(file);
    b.createTable({ name: 'Notes', fields: [{ name: 'Text', type: 'text' }] });
    const link = b.saveLink({ name: 'Away', engine: 'postgres', settings: { host: 'db.example', database: 'shop', user: 'office' } });
    b.linkTable({ table: 'Stock', link, remote: { schema: 'public', name: 'stock' }, key: 'id', fields: [{ name: 'id', type: 'autonumber' }, { name: 'name', type: 'text' }] });
    b.close();
    // Made elsewhere: nothing is connected to until it is allowed here (A2.18).
    const opened = await svc.open({ file });
    assert.equal(tried, 0);
    assert.deepEqual(opened.ask.map((a) => [a.place, a.host]), [['engine', 'db.example']]);
    const s = await svc.allowSources({ id: opened.id, addresses: opened.ask.map((a) => a.address) });
    assert.deepEqual(s.tables.map((t) => [t.name, t.count, t.linkError]), [['Notes', 0, undefined], ['Stock', 0, 'The engine at db.example:5432 could not be reached.']]);
    assert.equal(s.links[0].password, false);
  } finally {
    svc.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

const MYSQL_URL = process.env.OFFICE_MYSQL_TEST_URL;

test('a MySQL table linked live: read in, changed in the engine, read again (OFFICE_MYSQL_TEST_URL)', { skip: MYSQL_URL ? false : 'no OFFICE_MYSQL_TEST_URL' }, async () => {
  const { createRequire } = await import('node:module');
  const mysql = createRequire(import.meta.url)('mysql2/promise');
  const u = new URL(MYSQL_URL);
  const settings = { host: u.hostname, port: Number(u.port) || 3306, database: u.pathname.slice(1), user: decodeURIComponent(u.username), ssl: false };
  const password = decodeURIComponent(u.password);
  const my = await mysql.createConnection({ ...settings, password });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'engine-mysql-'));
  const file = path.join(dir, 'shop.rbase');
  const kept = new Map();
  const svc = createBaseService({ secrets: { available: () => true, get: (k) => kept.get(k) ?? null, set: (k, v) => kept.set(k, v), delete: (k) => kept.delete(k) } });
  try {
    await my.query('DROP TABLE IF EXISTS datafiles_link_test');
    await my.query('CREATE TABLE datafiles_link_test (id INT AUTO_INCREMENT PRIMARY KEY, name VARCHAR(80) NOT NULL, price DECIMAL(10,2), active TINYINT(1) DEFAULT 1, added DATE, seen DATETIME)');
    await my.query(`INSERT INTO datafiles_link_test (name, price, active, added, seen) VALUES ('Bread', 2.50, 1, '2026-10-01', '2026-10-01 09:30:00'), ('Milk', 1.20, 0, '2026-10-02', NULL)`);
    Base.open(file).close();
    const { id } = await svc.open({ file });
    const found = await svc.engineTables({ engine: 'mysql', settings, password });
    assert.ok(found.tables.some((t) => t.name === 'datafiles_link_test'));
    const { link } = await svc.saveLink({ id, name: 'Shop MySQL', engine: 'mysql', settings, password });
    const linked = await svc.linkTables({ id, link, remotes: [{ schema: settings.database, name: 'datafiles_link_test', as: 'Products' }] });
    assert.equal(linked.summary.tables.find((t) => t.name === 'Products').count, 2);
    assert.deepEqual(svc.fields({ id, table: 'Products' }).map((f) => [f.name, f.type]), [['id', 'autonumber'], ['name', 'text'], ['price', 'number'], ['active', 'yesno'], ['added', 'date'], ['seen', 'datetime']]);
    assert.deepEqual(svc.rows({ id, table: 'Products' }).rows.map((r) => [r.name, r.price, r.active, r.added, r.seen]), [['Bread', 2.5, 1, '2026-10-01', '2026-10-01T09:30:00'], ['Milk', 1.2, 0, '2026-10-02', null]]);
    const added = await svc.insert({ id, table: 'Products', values: { name: 'Eggs', price: '3.10', active: true, added: '2026-10-22', seen: '2026-10-22T17:24' } });
    assert.ok(added.rowid);
    await svc.update({ id, table: 'Products', rowid: 1, values: { price: '2.75' } });
    await svc.remove({ id, table: 'Products', rowids: [2] });
    const [there] = await my.query('SELECT name, CAST(price AS DOUBLE) AS price, active, CAST(seen AS CHAR) AS seen FROM datafiles_link_test ORDER BY id');
    assert.deepEqual(there.map((r) => ({ ...r })), [{ name: 'Bread', price: 2.75, active: 1, seen: '2026-10-01 09:30:00' }, { name: 'Eggs', price: 3.1, active: 1, seen: '2026-10-22 17:24:00' }]);
    await my.query(`UPDATE datafiles_link_test SET name = 'Loaf' WHERE id = 1`);
    await svc.refreshLinked({ id });
    assert.equal(svc.rows({ id, table: 'Products' }).rows[0].name, 'Loaf');
  } finally {
    svc.closeAll();
    await my.query('DROP TABLE IF EXISTS datafiles_link_test').catch(() => {});
    await my.end();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
