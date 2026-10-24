// Synced copies (D-2026-10-22-13, step 4): a table of a database engine kept
// as a copy in the file. Edits made with the engine out of reach wait and
// go when it is back; a record changed here and in the engine is a clash
// the engine wins, the version from here kept; deletes on either side
// follow; a table with no key is copied read only; switching between Live
// and Copy keeps the records. First against a stand-in engine, then
// against a real PostgreSQL and MySQL (OFFICE_PG_TEST_URL,
// OFFICE_MYSQL_TEST_URL), reached through a relay in the test that can be
// stopped and started as the engine would be.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { Base } from '../packages/database/src/index.js';
import { createBaseService } from '../apps/desktop/main/base.js';

const require = createRequire(import.meta.url);
const keptSecrets = () => {
  const kept = new Map();
  return { kept, secrets: { available: () => true, get: (k) => kept.get(k) ?? null, set: (k, v) => kept.set(k, v), delete: (k) => kept.delete(k) } };
};
const tmp = (name) => fs.mkdtempSync(path.join(os.tmpdir(), name));

/** A stand-in engine: one table of products, key id, reached or not as `down` says. */
function standIn() {
  const tables = {
    products: { key: 'id', fields: [{ name: 'id', type: 'autonumber', required: false }, { name: 'name', type: 'text', required: true }, { name: 'price', type: 'number', required: false }], rows: [], next: 1 },
    plain: { key: null, fields: [{ name: 'word', type: 'text', required: false }], rows: [{ word: 'one' }, { word: 'two' }], next: 1 },
    stamped: { key: 'id', fields: [{ name: 'id', type: 'autonumber', required: false }, { name: 'name', type: 'text', required: false }, { name: 'updated', type: 'datetime', required: false }],
      rows: [{ id: 1, name: 'A', updated: '2026-10-01T09:00:00' }, { id: 2, name: 'B', updated: '2026-10-02T09:00:00' }, { id: 3, name: 'C', updated: '2026-10-03T09:00:00' }], next: 4 },
  };
  const world = { down: false, tables, connects: 0 };
  const gate = () => { if (world.down) throw Object.assign(new Error('connect ECONNREFUSED 127.0.0.1:5999'), { code: 'ECONNREFUSED' }); };
  world.add = (name, price) => { const t = tables.products; const id = t.next++; t.rows.push({ id, name, price }); return id; };
  world.connect = async () => {
    gate();
    world.connects++;
    return {
      async tables() { gate(); return Object.keys(tables).map((name) => ({ schema: 'public', name, kind: 'table' })); },
      async describe(remote) { gate(); const t = tables[remote.name]; return { fields: t.fields.map((f) => ({ ...f })), key: t.key }; },
      async read(remote, { since = null } = {}) {
        gate();
        const t = tables[remote.name];
        world.lastSince = since;
        return t.rows.filter((r) => !since || r[since.field] >= since.value).map((r) => ({ ...r }));
      },
      async keys(remote) { gate(); return tables[remote.name].rows.map((r) => r.id); },
      async insert(remote, { key }, values) {
        gate();
        const t = tables[remote.name];
        if (values.name === 'REFUSE') throw new Error('new row violates check constraint "name_ok"');
        const id = t.next++;
        t.rows.push({ id, name: null, price: null, ...values, id });
        return key ? id : null;
      },
      async update(remote, { key }, k, values) { gate(); const r = tables[remote.name].rows.find((x) => x[key] === k); if (r) Object.assign(r, values); return r ? 1 : 0; },
      async remove(remote, { key }, ks) { gate(); const t = tables[remote.name]; const before = t.rows.length; t.rows = t.rows.filter((r) => !ks.includes(r[key])); return before - t.rows.length; },
      async close() {},
    };
  };
  return world;
}

const names = (svc, id, table) => svc.rows({ id, table, limit: 100 }).rows.map((r) => r.name);
const tableOf = (s, name) => s.tables.find((t) => t.name === name);

test('a copy kept in the file: worked on with the engine out of reach, caught up when it is back, a clash kept, deletes followed', async () => {
  const dir = tmp('sync-standin-');
  const file = path.join(dir, 'shop.rbase');
  const world = standIn();
  world.add('Bread', 2.5);
  world.add('Milk', 1.2);
  world.add('Eggs', 3.1);
  const { secrets } = keptSecrets();
  const svc = createBaseService({ secrets, connect: world.connect });
  try {
    Base.open(file).close();
    let s = await svc.open({ file });
    const id = s.id;
    const { link } = await svc.saveLink({ id, name: 'Shop', engine: 'postgres', settings: { host: 'h', database: 'd', user: 'u' }, password: 'p' });
    s = (await svc.linkTables({ id, link, remotes: [{ schema: 'public', name: 'products', as: 'Products', mode: 'copy' }] })).summary;
    const t = tableOf(s, 'Products');
    assert.equal(t.mode, 'copy');
    assert.equal(t.count, 3);
    assert.equal(t.waiting.count, 0);
    assert.ok(t.synced);
    // The records are a real table of the file.
    const raw = Base.read(file);
    try { assert.equal(raw.db.prepare("SELECT count(*) AS n FROM sqlite_master WHERE type = 'table' AND name = 'Products'").get().n, 1); } finally { raw.close(); }

    // The engine stops: changes are made here all the same, and wait.
    world.down = true;
    const added = await svc.insert({ id, table: 'Products', values: { name: 'Butter', price: '4.00' } });
    assert.ok(added.rowid);
    await svc.update({ id, table: 'Products', rowid: 1, values: { price: '2.75' } });
    await svc.remove({ id, table: 'Products', rowids: [2] });
    s = await svc.refreshLinked({ id, table: 'Products' });
    assert.equal(tableOf(s, 'Products').waiting.count, 3);
    assert.ok(tableOf(s, 'Products').waiting.since);
    assert.match(tableOf(s, 'Products').linkError, /ECONNREFUSED/);
    assert.equal(tableOf(s, 'Products').out, true);
    assert.deepEqual(names(svc, id, 'Products'), ['Bread', 'Eggs', 'Butter']);
    // Its key is the engine's to keep.
    await assert.rejects(svc.update({ id, table: 'Products', rowid: 1, values: { id: 99 } }), /engine's/);

    // Opened again with the engine still away: the copy is there, and so is what waits.
    svc.close({ id });
    s = await svc.open({ file });
    const id2 = s.id;
    assert.equal(tableOf(s, 'Products').count, 3);
    assert.equal(tableOf(s, 'Products').waiting.count, 3);

    // Back: what waited goes to the engine, by key, and the engine's own change comes here.
    world.down = false;
    world.tables.products.rows.find((r) => r.id === 3).name = 'Free-range eggs';
    s = await svc.refreshLinked({ id: id2, table: 'Products' });
    assert.equal(tableOf(s, 'Products').linkError, null);
    assert.equal(tableOf(s, 'Products').waiting.count, 0);
    assert.deepEqual(world.tables.products.rows.map((r) => [r.id, r.name, r.price]), [[1, 'Bread', 2.75], [3, 'Free-range eggs', 3.1], [4, 'Butter', 4]]);
    assert.deepEqual(svc.rows({ id: id2, table: 'Products' }).rows.map((r) => [r.id, r.name]), [[1, 'Bread'], [3, 'Free-range eggs'], [4, 'Butter']]);

    // A clash: changed here and in the engine since the last catch-up. The engine wins; the version from here is kept.
    await svc.update({ id: id2, table: 'Products', rowid: 1, values: { name: 'White bread' } });
    world.tables.products.rows.find((r) => r.id === 1).name = 'Brown bread';
    s = await svc.refreshLinked({ id: id2, table: 'Products' });
    assert.equal(tableOf(s, 'Products').clashes, 1);
    assert.equal(names(svc, id2, 'Products')[0], 'Brown bread');
    assert.equal(world.tables.products.rows[0].name, 'Brown bread');
    const [clash] = svc.clashes({ id: id2, table: 'Products' });
    assert.equal(clash.action, 'changed');
    assert.equal(clash.key, 1);
    assert.equal(clash.mine.name, 'White bread');
    assert.equal(clash.theirs.name, 'Brown bread');
    // Mine put back: it goes to the engine at the next catch-up.
    s = svc.settleClash({ id: id2, table: 'Products', clash: clash.id, keep: 'mine' });
    assert.equal(tableOf(s, 'Products').clashes, 0);
    await svc.refreshLinked({ id: id2, table: 'Products' });
    assert.equal(world.tables.products.rows[0].name, 'White bread');

    // Deletes on either side.
    world.tables.products.rows = world.tables.products.rows.filter((r) => r.id !== 3);
    await svc.remove({ id: id2, table: 'Products', rowids: [4] });
    s = await svc.refreshLinked({ id: id2, table: 'Products' });
    assert.deepEqual(names(svc, id2, 'Products'), ['White bread']);
    assert.deepEqual(world.tables.products.rows.map((r) => r.id), [1]);
    // Deleted here, changed there: it comes back, and says so.
    world.add('Jam', 3);
    await svc.refreshLinked({ id: id2, table: 'Products' });
    const jam = svc.rows({ id: id2, table: 'Products' }).rows.find((r) => r.name === 'Jam');
    await svc.remove({ id: id2, table: 'Products', rowids: [jam._rowid] });
    world.tables.products.rows.find((r) => r.name === 'Jam').price = 3.5;
    s = await svc.refreshLinked({ id: id2, table: 'Products' });
    assert.ok(names(svc, id2, 'Products').includes('Jam'));
    assert.equal(svc.clashes({ id: id2, table: 'Products' })[0].action, 'deleted');
    // Changed here, deleted there: it goes, what was changed kept to see.
    await svc.update({ id: id2, table: 'Products', rowid: jam._rowid, values: { price: '9' } });
    world.tables.products.rows = world.tables.products.rows.filter((r) => r.name !== 'Jam');
    await svc.refreshLinked({ id: id2, table: 'Products' });
    assert.ok(!names(svc, id2, 'Products').includes('Jam'));
    assert.equal(svc.clashes({ id: id2, table: 'Products' })[0].action, 'gone');
    assert.equal(svc.clashes({ id: id2, table: 'Products' })[0].mine.price, 9);
    // A record the engine refuses: the engine wins, the one from here kept with why.
    await svc.insert({ id: id2, table: 'Products', values: { name: 'REFUSE' } });
    s = await svc.refreshLinked({ id: id2, table: 'Products' });
    assert.ok(!names(svc, id2, 'Products').includes('REFUSE'));
    const why = svc.clashes({ id: id2, table: 'Products' })[0];
    assert.equal(why.action, 'refused');
    assert.match(why.why, /check constraint/);

    // Two added here while away, numbered here, given the engine's numbers when they go; another computer's comes too.
    world.down = true;
    await svc.insert({ id: id2, table: 'Products', values: { name: 'Tea' } });
    await svc.insert({ id: id2, table: 'Products', values: { name: 'Coffee' } });
    world.down = false;
    const other = world.add('Sugar', 1);
    await svc.refreshLinked({ id: id2, table: 'Products' });
    const after = svc.rows({ id: id2, table: 'Products', limit: 100 }).rows.map((r) => [r.id, r.name]).sort((a, b) => a[0] - b[0]);
    assert.deepEqual(after, world.tables.products.rows.map((r) => [r.id, r.name]).sort((a, b) => a[0] - b[0]));
    assert.ok(after.some(([k, n]) => k === other && n === 'Sugar'));
    assert.ok(after.some(([, n]) => n === 'Tea') && after.some(([, n]) => n === 'Coffee'));

    // Switched to Live and back, the records are kept.
    s = await svc.setLinkMode({ id: id2, table: 'Products', mode: 'live' });
    assert.equal(tableOf(s, 'Products').mode, 'live');
    assert.equal(tableOf(s, 'Products').count, world.tables.products.rows.length);
    const fileNow = Base.read(file);
    try { assert.equal(fileNow.db.prepare("SELECT count(*) AS n FROM sqlite_master WHERE name = 'Products'").get().n, 0); } finally { fileNow.close(); }
    world.down = true;
    s = await svc.setLinkMode({ id: id2, table: 'Products', mode: 'copy' });
    assert.equal(tableOf(s, 'Products').mode, 'copy');
    assert.equal(tableOf(s, 'Products').count, world.tables.products.rows.length);
    // A copy with changes waiting does not go live until they have gone.
    await svc.insert({ id: id2, table: 'Products', values: { name: 'Honey' } });
    await assert.rejects(svc.setLinkMode({ id: id2, table: 'Products', mode: 'live' }), /wait/);
    assert.throws(() => svc.unlinkTable({ id: id2, table: 'Products' }), /wait/);
    world.down = false;
    await svc.refreshLinked({ id: id2, table: 'Products' });
    assert.ok(world.tables.products.rows.some((r) => r.name === 'Honey'));

    // A table with no key: copied, read, not changed.
    s = (await svc.linkTables({ id: id2, link, remotes: [{ schema: 'public', name: 'plain', as: 'Plain', mode: 'copy' }] })).summary;
    assert.equal(tableOf(s, 'Plain').count, 2);
    assert.equal(tableOf(s, 'Plain').readOnly, true);
    assert.ok(svc.fields({ id: id2, table: 'Plain' }).every((f) => f.readOnly));
    await assert.rejects(svc.insert({ id: id2, table: 'Plain', values: { word: 'three' } }), /no primary key/);
    world.tables.plain.rows.push({ word: 'three' });
    await svc.refreshLinked({ id: id2, table: 'Plain' });
    assert.deepEqual(svc.rows({ id: id2, table: 'Plain' }).rows.map((r) => r.word), ['one', 'two', 'three']);
  } finally {
    svc.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a large copy read by its updated-at field: only what changed is read, and what was deleted there still goes', async () => {
  const dir = tmp('sync-stamp-');
  const file = path.join(dir, 'shop.rbase');
  const world = standIn();
  const svc = createBaseService({ secrets: keptSecrets().secrets, connect: world.connect });
  try {
    Base.open(file).close();
    const { id } = await svc.open({ file });
    const { link } = await svc.saveLink({ id, name: 'Shop', engine: 'postgres', settings: { host: 'h', database: 'd', user: 'u' }, password: 'p' });
    await svc.linkTables({ id, link, remotes: [{ schema: 'public', name: 'stamped', as: 'Stamped', mode: 'copy' }] });
    let s = svc.setLinkStamp({ id, table: 'Stamped', field: 'updated' });
    assert.equal(tableOf(s, 'Stamped').stamp, 'updated');
    const rows = world.tables.stamped.rows;
    rows[0].name = 'A2'; rows[0].updated = '2026-10-05T09:00:00';
    world.tables.stamped.rows = rows.filter((r) => r.id !== 2);
    world.tables.stamped.rows.push({ id: 4, name: 'D', updated: '2026-10-06T09:00:00' });
    s = await svc.refreshLinked({ id, table: 'Stamped' });
    assert.deepEqual(world.lastSince, { field: 'updated', value: '2026-10-03T09:00:00' });
    assert.deepEqual(svc.rows({ id, table: 'Stamped' }).rows.map((r) => [r.id, r.name]), [[1, 'A2'], [3, 'C'], [4, 'D']]);
    assert.throws(() => svc.setLinkStamp({ id, table: 'Stamped', field: 'nothing' }), /no field/);
  } finally {
    svc.closeAll();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a file with no links is as it was: no sync tables, and its tables work as before', () => {
  const dir = tmp('sync-none-');
  try {
    const b = Base.open(path.join(dir, 'plain.rbase'));
    b.createTable({ name: 'Notes', fields: [{ name: 'Text', type: 'text' }] });
    b.insert('Notes', { Text: 'hello' });
    assert.equal(b.hasDesignTable('rutba_sync'), false);
    assert.equal(b.hasDesignTable('rutba_clashes'), false);
    assert.deepEqual(b.db.prepare('PRAGMA table_info(rutba_linked)').all().map((c) => c.name), ['tbl', 'link', 'remote', 'key', 'caption', 'position']);
    assert.equal(b.rows('Notes')[0].Text, 'hello');
    b.close();
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

/**
 * A relay to an engine on this computer, standing in for its network: stop()
 * cuts every connection and refuses new ones, as an engine stopped would;
 * start() listens again on the same port.
 */
async function relay(host, port) {
  const sockets = new Set();
  let server = null;
  const listen = (at = 0) => new Promise((resolve) => {
    server = net.createServer((client) => {
      const up = net.connect(port, host);
      sockets.add(client); sockets.add(up);
      client.on('error', () => {}); up.on('error', () => {});
      client.on('close', () => { sockets.delete(client); up.destroy(); });
      up.on('close', () => { sockets.delete(up); client.destroy(); });
      client.pipe(up); up.pipe(client);
    });
    server.listen(at, '127.0.0.1', () => resolve(server.address().port));
  });
  const at = await listen();
  return {
    port: at,
    stop: () => new Promise((resolve) => { for (const s of sockets) s.destroy(); sockets.clear(); server.close(() => resolve()); }),
    start: () => listen(at),
  };
}

/** The same story against a real engine: `run` makes its table and answers how to read it. */
async function engineStory({ engine, settings, password, sql, remote }) {
  const dir = tmp(`sync-${engine}-`);
  const file = path.join(dir, 'shop.rbase');
  const r = await relay(settings.host, settings.port);
  const { secrets } = keptSecrets();
  const svc = createBaseService({ secrets });
  try {
    Base.open(file).close();
    let s = await svc.open({ file });
    const id = s.id;
    const { link } = await svc.saveLink({ id, name: 'Shop', engine, settings: { ...settings, host: '127.0.0.1', port: r.port }, password });
    s = (await svc.linkTables({ id, link, remotes: [{ ...remote, as: 'Products', mode: 'copy' }, { ...remote, name: `${remote.name}_plain`, as: 'Plain', mode: 'copy' }] })).summary;
    assert.equal(tableOf(s, 'Products').mode, 'copy');
    assert.equal(tableOf(s, 'Products').count, 3);
    assert.equal(tableOf(s, 'Plain').readOnly, true);
    await assert.rejects(svc.insert({ id, table: 'Plain', values: { word: 'x' } }), /no primary key/);

    // The engine stopped: edits made here wait.
    await r.stop();
    await svc.update({ id, table: 'Products', rowid: 1, values: { price: '2.75' } });
    await svc.insert({ id, table: 'Products', values: { name: 'Butter', price: '4', active: true } });
    await svc.remove({ id, table: 'Products', rowids: [2] });
    s = await svc.refreshLinked({ id, table: 'Products' });
    assert.equal(tableOf(s, 'Products').waiting.count, 3);
    assert.ok(tableOf(s, 'Products').linkError);

    // Started again: caught up.
    await r.start();
    s = await svc.refreshLinked({ id, table: 'Products' });
    assert.equal(tableOf(s, 'Products').linkError, null);
    assert.equal(tableOf(s, 'Products').waiting.count, 0);
    assert.deepEqual(await sql.list(), [['Bread', 2.75], ['Eggs', 3.1], ['Butter', 4]]);
    assert.deepEqual(names(svc, id, 'Products'), ['Bread', 'Eggs', 'Butter']);

    // A clash: the engine wins, the version from here kept.
    await svc.update({ id, table: 'Products', rowid: 1, values: { name: 'White bread' } });
    await sql.rename(1, 'Brown bread');
    s = await svc.refreshLinked({ id, table: 'Products' });
    assert.equal(tableOf(s, 'Products').clashes, 1);
    assert.equal(names(svc, id, 'Products')[0], 'Brown bread');
    assert.equal(svc.clashes({ id, table: 'Products' })[0].mine.name, 'White bread');

    // Deletes on either side.
    await sql.remove(3);
    const butter = svc.rows({ id, table: 'Products' }).rows.find((x) => x.name === 'Butter');
    await svc.remove({ id, table: 'Products', rowids: [butter._rowid] });
    await svc.refreshLinked({ id, table: 'Products' });
    assert.deepEqual(names(svc, id, 'Products'), ['Brown bread']);
    assert.deepEqual((await sql.list()).map((x) => x[0]), ['Brown bread']);

    // Opened with the engine stopped, the copy is there to work on.
    svc.close({ id });
    await r.stop();
    s = await svc.open({ file });
    assert.equal(tableOf(s, 'Products').count, 1);
    await svc.insert({ id: s.id, table: 'Products', values: { name: 'Tea', price: '1.5', active: true } });
    await r.start();
    await svc.refreshLinked({ id: s.id, table: 'Products' });
    assert.deepEqual((await sql.list()).map((x) => x[0]), ['Brown bread', 'Tea']);
  } finally {
    svc.closeAll();
    await r.stop().catch(() => {});
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

const PG = process.env.OFFICE_PG_TEST_URL;
test('a PostgreSQL table kept as a copy: the engine stopped and started, a clash, deletes, a key-less table read only (OFFICE_PG_TEST_URL)', { skip: PG ? false : 'no OFFICE_PG_TEST_URL' }, async () => {
  const pg = require('pg');
  const u = new URL(PG);
  const c = new pg.Client({ connectionString: PG });
  await c.connect();
  try {
    await c.query('DROP SCHEMA IF EXISTS datafiles_sync_test CASCADE');
    await c.query('CREATE SCHEMA datafiles_sync_test');
    await c.query('CREATE TABLE datafiles_sync_test.products (id serial PRIMARY KEY, name varchar(80) NOT NULL, price numeric(10,2), active boolean DEFAULT true)');
    await c.query(`INSERT INTO datafiles_sync_test.products (name, price) VALUES ('Bread', 2.50), ('Milk', 1.20), ('Eggs', 3.10)`);
    await c.query('CREATE TABLE datafiles_sync_test.products_plain (word text)');
    await c.query(`INSERT INTO datafiles_sync_test.products_plain VALUES ('one'), ('two')`);
    await engineStory({
      engine: 'postgres',
      settings: { host: u.hostname, port: Number(u.port) || 5432, database: u.pathname.slice(1), user: decodeURIComponent(u.username), ssl: false },
      password: decodeURIComponent(u.password),
      remote: { schema: 'datafiles_sync_test', name: 'products' },
      sql: {
        list: async () => (await c.query('SELECT name, price::float AS price FROM datafiles_sync_test.products ORDER BY id')).rows.map((x) => [x.name, x.price]),
        rename: (k, name) => c.query('UPDATE datafiles_sync_test.products SET name = $1 WHERE id = $2', [name, k]),
        remove: (k) => c.query('DELETE FROM datafiles_sync_test.products WHERE id = $1', [k]),
      },
    });
  } finally {
    await c.query('DROP SCHEMA IF EXISTS datafiles_sync_test CASCADE').catch(() => {});
    await c.end();
  }
});

const MY = process.env.OFFICE_MYSQL_TEST_URL;
test('a MySQL table kept as a copy: the engine stopped and started, a clash, deletes, a key-less table read only (OFFICE_MYSQL_TEST_URL)', { skip: MY ? false : 'no OFFICE_MYSQL_TEST_URL' }, async () => {
  const mysql = require('mysql2/promise');
  const u = new URL(MY);
  const settings = { host: u.hostname, port: Number(u.port) || 3306, database: u.pathname.slice(1), user: decodeURIComponent(u.username), ssl: false };
  const password = decodeURIComponent(u.password);
  const my = await mysql.createConnection({ ...settings, password });
  try {
    await my.query('DROP TABLE IF EXISTS datafiles_sync_test, datafiles_sync_test_plain');
    await my.query('CREATE TABLE datafiles_sync_test (id INT AUTO_INCREMENT PRIMARY KEY, name VARCHAR(80) NOT NULL, price DECIMAL(10,2), active TINYINT(1) DEFAULT 1)');
    await my.query(`INSERT INTO datafiles_sync_test (name, price) VALUES ('Bread', 2.50), ('Milk', 1.20), ('Eggs', 3.10)`);
    await my.query('CREATE TABLE datafiles_sync_test_plain (word TEXT)');
    await my.query(`INSERT INTO datafiles_sync_test_plain VALUES ('one'), ('two')`);
    await engineStory({
      engine: 'mysql',
      settings,
      password,
      remote: { schema: settings.database, name: 'datafiles_sync_test' },
      sql: {
        list: async () => (await my.query('SELECT name, CAST(price AS DOUBLE) AS price FROM datafiles_sync_test ORDER BY id'))[0].map((x) => [x.name, x.price]),
        rename: (k, name) => my.query('UPDATE datafiles_sync_test SET name = ? WHERE id = ?', [name, k]),
        remove: (k) => my.query('DELETE FROM datafiles_sync_test WHERE id = ?', [k]),
      },
    });
  } finally {
    await my.query('DROP TABLE IF EXISTS datafiles_sync_test, datafiles_sync_test_plain').catch(() => {});
    await my.end();
  }
});
