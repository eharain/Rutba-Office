// The Data Files format, published (D-2026-10-26-01): docs/DATA-FILES-FORMAT.md
// says what a .rbase file holds, tools/check-rbase.mjs says whether a file
// follows it, and docs/data-files-format/samples/ holds three files made by
// the database code itself. These tests hold the three together: the
// samples and every ready-made database follow the text, a file that departs
// from it is told so, and the design tables in packages/database cannot gain
// or lose a table or a column without the text naming it.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { DatabaseSync } from 'node:sqlite';
import {
  Base, DESIGN_VERSION, FIELD_TYPES, TEMPLATE_IDS, SAMPLE_IDS, buildTemplate, buildSample,
  ensureSync, reconcile, replaceRows, markSynced, copyForeign, designerDesign, unknownKeys,
} from '../packages/database/src/index.js';
import { checkRbase, SPEC_VERSION, DESIGN_TABLES, FIELD_TYPES as SPEC_FIELD_TYPES } from '../tools/check-rbase.mjs';
import { makeSamples, SAMPLES_DIR, EXAMPLE_PEOPLE } from '../tools/make-rbase-samples.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SPEC = fs.readFileSync(path.join(ROOT, 'docs', 'DATA-FILES-FORMAT.md'), 'utf8');
const SAMPLES = ['plain-tables.rbase', 'query-form-report.rbase', 'protected.rbase', 'copies-and-sources.rbase'];
const tmp = (name) => fs.mkdtempSync(path.join(os.tmpdir(), name));

/** A copy of a sample to spoil, changed by `spoil(db)` with plain SQLite. */
function spoiled(sample, spoil) {
  const file = path.join(tmp('rbase-spoiled-'), sample);
  fs.copyFileSync(path.join(SAMPLES_DIR, sample), file);
  const db = new DatabaseSync(file);
  try { spoil(db); } finally { db.close(); }
  return file;
}

const said = (r, re) => r.problems.some((p) => re.test(p));

test('the three sample files are there, small, and follow the format', () => {
  for (const name of SAMPLES) {
    const file = path.join(SAMPLES_DIR, name);
    assert.ok(fs.existsSync(file), `${name} is committed`);
    assert.ok(fs.statSync(file).size < 200 * 1024, `${name} is under 200 KB`);
    assert.ok(!fs.existsSync(`${file}-wal`) && !fs.existsSync(`${file}-shm`), `${name} has no -wal or -shm beside it`);
    const r = checkRbase(file);
    assert.deepEqual(r.problems, [], name);
    assert.equal(r.version, DESIGN_VERSION);
  }
  // Read-only means read-only: checking left nothing beside the file.
  assert.ok(!fs.readdirSync(SAMPLES_DIR).some((f) => /-(wal|shm|journal)$/.test(f)));
});

test('the sample script makes the samples again with the real database code, and the protected one opens with its example passwords', () => {
  const dir = tmp('rbase-samples-');
  const files = makeSamples(dir);
  assert.deepEqual(files.map((f) => path.basename(f)), SAMPLES);
  for (const f of files) assert.deepEqual(checkRbase(f).problems, [], path.basename(f));
  const base = Base.open(files[2]);
  try {
    assert.equal(base.isProtected(), true);
    assert.deepEqual(base.tables(), [], 'nothing shows before sign-in');
    for (const p of EXAMPLE_PEOPLE) {
      assert.equal(base.signIn(p.name, p.password).role, p.role);
      base.signOut();
    }
    assert.throws(() => base.signIn('Owner', 'not-the-password'), /do not match/);
  } finally {
    base.close();
  }
  // The text gives the same example sign-ins.
  for (const p of EXAMPLE_PEOPLE) assert.ok(SPEC.includes(`\`${p.name}\``) && SPEC.includes(`\`${p.password}\``), `the spec names ${p.name}'s example password`);
});

test('every ready-made database Office makes follows the format', () => {
  const dir = tmp('rbase-ready-made-');
  const made = [
    ...TEMPLATE_IDS.map((id) => ['template', id, (b) => buildTemplate(b, id)]),
    ...SAMPLE_IDS.map((id) => ['sample', id, (b) => buildSample(b, id, { withData: false })]),
  ];
  assert.ok(made.length > 10);
  for (const [kind, id, build] of made) {
    const file = path.join(dir, `${kind}-${id}.rbase`);
    // Made in memory and written out whole, as Save As writes a copy: one write, not one for each step.
    const base = Base.open(':memory:');
    try {
      build(base);
      base.db.exec(`VACUUM INTO '${file.replace(/'/g, "''")}'`);
    } finally {
      base.close();
    }
    assert.deepEqual(checkRbase(file).problems, [], `${kind} ${id}`);
  }
});

test('a file with tables kept as synced copies and read from a source follows the format', () => {
  const file = path.join(tmp('rbase-copies-'), 'copies.rbase');
  const base = Base.open(file);
  try {
    // A source: a CSV file's rows kept as a copy, read only.
    const csv = base.saveLink({ name: 'Price list', engine: 'file', settings: { where: '/data/prices.csv', sheet: null, header: 1, every: 15 } });
    const priceFields = [{ name: 'Item', type: 'text' }, { name: 'Price', type: 'currency' }];
    base.linkTable({ table: 'Prices', link: csv, remote: { sheet: null }, key: null, fields: priceFields, mode: 'copy' });
    replaceRows(base, 'Prices', [{ Item: 'Tea', Price: 2.5 }, { Item: 'Cake', Price: 3.25 }], priceFields);
    markSynced(base, 'Prices');
    // An engine's table kept as a copy, with a change made here waiting to go.
    const pg = base.saveLink({ name: 'Head office', engine: 'postgres', settings: { host: 'db.example.com', port: 5432, database: 'shop', user: 'reader', ssl: true } });
    const stockFields = [{ name: 'id', type: 'integer' }, { name: 'Item', type: 'text' }, { name: 'Count', type: 'integer' }];
    base.linkTable({ table: 'Stock', link: pg, remote: { schema: 'public', name: 'stock' }, key: 'id', fields: stockFields, mode: 'copy' });
    reconcile(base, 'Stock', { rows: [{ id: 1, Item: 'Tea', Count: 40 }, { id: 2, Item: 'Cake', Count: 12 }] }, { fields: stockFields, key: 'id' });
    base.update('Stock', 2, { Count: 11 });
    // And one read live: nothing of its records is in the file.
    base.linkTable({ table: 'Orders', link: pg, remote: { schema: 'public', name: 'orders' }, key: 'id', fields: [{ name: 'id', type: 'integer' }, { name: 'Total', type: 'currency' }] });
  } finally {
    base.close();
  }
  const r = checkRbase(file);
  assert.deepEqual(r.problems, []);
  const db = new DatabaseSync(file, { readOnly: true });
  try {
    assert.equal(db.prepare("SELECT state FROM rutba_sync WHERE tbl = 'Stock' AND rid = 2").get().state, 'changed');
    assert.equal(db.prepare("SELECT 1 FROM sqlite_master WHERE name = 'Orders'").get(), undefined);
  } finally {
    db.close();
  }
});

test('a file that departs from the format is told so, and how', () => {
  const cases = [
    ['a form layout that is not JSON', 'query-form-report.rbase', (db) => db.exec("UPDATE rutba_forms SET layout = '{\"name\": \"Customers\", controls: [' WHERE name = 'Customers'"), /Form Customers's layout is not JSON/],
    ['a form layout without its table', 'query-form-report.rbase', (db) => db.exec("UPDATE rutba_forms SET layout = json_remove(layout, '$.table') WHERE name = 'Customers'"), /Form Customers shows undefined/],
    ['a control of a kind there is none of', 'query-form-report.rbase', (db) => db.exec("UPDATE rutba_forms SET layout = json_set(layout, '$.controls[1].kind', 'slider') WHERE name = 'Customers'"), /kind is "slider"/],
    ['a query column from a field that is not there', 'query-form-report.rbase', (db) => db.exec("UPDATE rutba_queries SET design = json_set(design, '$.columns[0].field', 'Nickname') WHERE name = 'Order Lines'"), /Customers has no field Nickname/],
    ['a report total of no kind', 'query-form-report.rbase', (db) => db.exec("UPDATE rutba_reports SET layout = json_set(layout, '$.totals[0].fn', 'median') WHERE name = 'Orders by Customer'"), /total 1, is not a field with one of/],
    ['a print template block of no kind', 'query-form-report.rbase', (db) => db.exec("UPDATE rutba_documents SET layout = json_set(layout, '$.blocks[0].type', 'barcode') WHERE name = 'Statement'"), /type is "barcode"/],
    ['an unknown field type', 'plain-tables.rbase', (db) => db.exec("UPDATE rutba_fields SET type = 'picture' WHERE tbl = 'Customers' AND name = 'Notes'"), /has the type "picture"/],
    ['a field with no column', 'plain-tables.rbase', (db) => db.exec("INSERT INTO rutba_fields (tbl, name, type, position) VALUES ('Customers', 'Phone', 'text', 9)"), /Customers\.Phone has no column/],
    ['a relationship pointing at a field that is not there', 'plain-tables.rbase', (db) => db.exec("UPDATE rutba_relations SET to_field = 'Code'"), /points at a field that is not there/],
    ['a design version newer than the text', 'plain-tables.rbase', (db) => db.exec(`UPDATE rutba_meta SET value = '${DESIGN_VERSION + 1}' WHERE key = 'version'`), /newer than this specification/],
    ['a version 7 file without the copies table', 'plain-tables.rbase', (db) => db.exec('DROP TABLE rutba_sync'), /rutba_sync is missing/],
    ['a version 7 file without a column of copies', 'plain-tables.rbase', (db) => db.exec('ALTER TABLE rutba_linked DROP COLUMN stamp'), /rutba_linked has no column stamp: a file of design version 7/],
    ['a design table missing', 'plain-tables.rbase', (db) => db.exec('DROP TABLE rutba_documents'), /rutba_documents is missing/],
    ['a password kept in a link', 'plain-tables.rbase', (db) => db.exec("INSERT INTO rutba_links (id, name, engine, settings) VALUES ('link-x', 'Engine', 'postgres', '{\"host\":\"h\",\"password\":\"hunter2\"}')"), /keeps a password/],
    ['a password kept as itself', 'protected.rbase', (db) => db.exec("UPDATE rutba_people SET hash = 'example-password' WHERE name = 'Owner'"), /Person Owner's password is not kept as scrypt/],
  ];
  for (const [what, sample, spoil, says] of cases) {
    const r = checkRbase(spoiled(sample, spoil));
    assert.equal(r.ok, false, what);
    assert.ok(said(r, says), `${what}: ${r.problems.join(' | ')}`);
  }
  // What a reader ignores and keeps is a note, not a problem.
  const kept = checkRbase(spoiled('plain-tables.rbase', (db) => {
    db.exec('CREATE TABLE rutba_future (x TEXT)');
    db.exec("INSERT INTO rutba_meta (key, value) VALUES ('future', '1')");
  }));
  assert.deepEqual(kept.problems, []);
  assert.ok(kept.notes.some((n) => /rutba_future/.test(n)) && kept.notes.some((n) => /future/.test(n)));
  // Not a database at all.
  const text = path.join(tmp('rbase-text-'), 'note.rbase');
  fs.writeFileSync(text, 'not a database');
  assert.ok(said(checkRbase(text), /not an SQLite database/));
});

test('the command exits 0 for a file that follows the format and 1 with the problems for one that does not', () => {
  const run = (...args) => spawnSync(process.execPath, [path.join(ROOT, 'tools', 'check-rbase.mjs'), ...args], { encoding: 'utf8' });
  const good = run(path.join(SAMPLES_DIR, 'plain-tables.rbase'));
  assert.equal(good.status, 0, good.stdout + good.stderr);
  assert.match(good.stdout, new RegExp(`follows the Data Files format, design version ${DESIGN_VERSION}`));
  const bad = run(spoiled('plain-tables.rbase', (db) => db.exec("UPDATE rutba_fields SET type = 'picture' WHERE name = 'Notes'")));
  assert.equal(bad.status, 1);
  assert.match(bad.stdout, /does not follow the Data Files format \(1 problem\):\n {2}- Field Customers\.Notes has the type "picture"/);
  assert.equal(run().status, 2);
});

/* ── files made elsewhere (A2.20, section 11) ─────────────────────────── */

/** Another program's database: two related tables, a picture column, a WITHOUT ROWID table, and dates as that program writes them. */
function otherProgram() {
  const file = path.join(tmp('rbase-other-'), 'shop.sqlite');
  const db = new DatabaseSync(file);
  try {
    db.exec(`CREATE TABLE customers (id INTEGER PRIMARY KEY, name VARCHAR(40) NOT NULL, joined DATE, vip BOOLEAN, notes TEXT, photo BLOB);
      CREATE TABLE orders (id INTEGER PRIMARY KEY AUTOINCREMENT, customer INTEGER REFERENCES customers(id), total DECIMAL(10,2), at DATETIME);
      CREATE TABLE settings (k TEXT PRIMARY KEY, v TEXT) WITHOUT ROWID;
      INSERT INTO customers VALUES (1, 'Ann', '2024-01-02', 1, 'Pays early', x'89504e47'), (2, 'Bo', '2024-02-03', 0, NULL, NULL);
      INSERT INTO orders (customer, total, at) VALUES (1, 9.5, '2024-01-02 10:00:00'), (2, 3, '2024-01-03 11:00:00');
      INSERT INTO settings VALUES ('theme', 'dark');`);
  } finally {
    db.close();
  }
  return file;
}

test('another program\'s database opens read only, its tables shown, and its file is left exactly as it was', () => {
  const file = otherProgram();
  const before = fs.readFileSync(file);
  assert.deepEqual(Base.inspect(file), { kind: 'other', version: null, newer: null });
  const b = Base.foreign(file);
  try {
    assert.deepEqual(b.tables().map((t) => [t.name, t.readOnly]), [['customers', true], ['orders', true]]);
    assert.deepEqual(b.foreign.left, [{ table: 'settings', why: 'without-rowid' }]);
    const types = (t) => Object.fromEntries(b.fields(t).map((f) => [f.name, f.type]));
    assert.deepEqual(types('customers'), { id: 'integer', name: 'text', joined: 'date', vip: 'yesno', notes: 'longtext', photo: 'longtext' });
    // A time written with a space is not a Data Files time: shown as the text it is.
    assert.deepEqual(types('orders'), { id: 'autonumber', customer: 'integer', total: 'number', at: 'longtext' });
    assert.ok(b.fields('customers').every((f) => f.readOnly));
    assert.deepEqual(b.relationships().map((r) => `${r.from.table}.${r.from.field}>${r.to.table}.${r.to.field}`), ['orders.customer>customers.id']);
    // Read as any database is read: records, a query across the relationship, SQL of one's own.
    assert.equal(b.count('orders'), 2);
    assert.deepEqual(b.runQuery({ design: { tables: ['orders', 'customers'], columns: [{ table: 'customers', field: 'name' }, { table: 'orders', field: 'total' }] } }).rows, [{ name: 'Ann', total: 9.5 }, { name: 'Bo', total: 3 }]);
    assert.deepEqual(b.query('SELECT count(*) AS n FROM settings').rows, [{ n: 1 }]);
    // Nothing written: records and design both refused, in words.
    assert.throws(() => b.insert('orders', { total: 1 }), /made by another program and is open read only/);
    assert.throws(() => b.createTable({ name: 'Notes' }), /its design is not changed here/);
    assert.throws(() => b.saveQuery('Spend', { tables: ['orders'], columns: [{ table: 'orders', field: 'total' }] }), /its design is not changed here/);
  } finally {
    b.close();
  }
  assert.ok(Buffer.from(fs.readFileSync(file)).equals(before), 'the file is byte for byte as it was');
  assert.deepEqual(fs.readdirSync(path.dirname(file)), ['shop.sqlite'], 'no -wal, -shm or journal beside it');
  // Opened as one of ours, it is refused rather than given design tables.
  assert.throws(() => Base.open(file), (err) => err.foreign === true && /made by another program/.test(err.message));
  assert.ok(Buffer.from(fs.readFileSync(file)).equals(before));
});

test('with changes allowed, another program\'s records change in place and still nothing of ours goes into its file', () => {
  const file = otherProgram();
  const b = Base.foreign(file, { write: true });
  try {
    assert.ok(!b.tables()[0].readOnly);
    assert.equal(b.fields('customers').find((f) => f.name === 'photo').readOnly, true, 'binary values stay read only');
    b.update('customers', 2, { name: 'Bob' });
    b.insert('orders', { customer: 2, total: 4.25 });
    assert.throws(() => b.insert('orders', { customer: 99, total: 1 }), /no record|points at/i, 'its foreign keys are kept');
    assert.throws(() => b.addField('customers', { name: 'Phone', type: 'text' }), /its design is not changed here/);
  } finally {
    b.close();
  }
  const db = new DatabaseSync(file, { readOnly: true });
  try {
    assert.equal(db.prepare('SELECT name FROM customers WHERE id = 2').get().name, 'Bob');
    assert.equal(db.prepare('SELECT count(*) AS n FROM orders').get().n, 3);
    assert.deepEqual(db.prepare("SELECT name FROM sqlite_master WHERE name LIKE 'rutba%'").all(), []);
    assert.equal(db.prepare('PRAGMA journal_mode').get().journal_mode, 'delete');
  } finally {
    db.close();
  }
});

test('a Data Files copy of another program\'s database keeps its tables, types, numbers and relationship, and follows the format', () => {
  const file = otherProgram();
  const from = Base.foreign(file);
  const copyFile = path.join(path.dirname(file), 'shop.rbase');
  const to = Base.open(copyFile);
  let made;
  try {
    made = copyForeign(from, to);
    assert.deepEqual(made.tables.map((x) => [x.table, x.count, x.skipped]), [['customers', 2, 0], ['orders', 2, 0]]);
    assert.deepEqual(made.left, ['customers.photo']);
    assert.equal(made.relationships, 1);
    assert.deepEqual(to.fields('customers').map((f) => [f.name, f.type]), [['id', 'autonumber'], ['name', 'text'], ['joined', 'date'], ['vip', 'yesno'], ['notes', 'longtext']]);
    assert.deepEqual(to.rows('orders').map((r) => [r.id, r.customer, r.total]), [[1, 1, 9.5], [2, 2, 3]]);
    // Now a database of ours: its design changes.
    to.addField('customers', { name: 'Phone', type: 'text' });
  } finally {
    to.close();
    from.close();
  }
  assert.deepEqual(checkRbase(copyFile).problems, []);
});

test('a database of a newer design version is refused in plain words and left unchanged; one of a version that is not a number too', () => {
  for (const [version, says] of [[String(DESIGN_VERSION + 1), new RegExp(`made by a newer version of Rutba Office \\(design version ${DESIGN_VERSION + 1}; this one knows up to ${DESIGN_VERSION}\\)\\. Update Rutba Office to open it\\. Nothing in it was changed\\.`)], ['7b', /says its design version is "7b", which no version of Rutba Office writes/]]) {
    const file = spoiled('plain-tables.rbase', (db) => db.prepare("UPDATE rutba_meta SET value = ? WHERE key = 'version'").run(version));
    const before = fs.readFileSync(file);
    assert.match(Base.inspect(file).newer, says);
    assert.throws(() => Base.open(file), says);
    assert.ok(Buffer.from(fs.readFileSync(file)).equals(before), `version ${version}: unchanged`);
    assert.deepEqual(fs.readdirSync(path.dirname(file)), ['plain-tables.rbase']);
  }
  // The version this build writes, and older, open.
  assert.equal(Base.inspect(path.join(SAMPLES_DIR, 'plain-tables.rbase')).newer, null);
});

test('a version 6 file is brought to 7: the copies\' tables and columns added, nothing that is there changed', () => {
  const file = spoiled('plain-tables.rbase', (db) => {
    db.exec("DROP TABLE rutba_sync; DROP TABLE rutba_clashes; ALTER TABLE rutba_linked DROP COLUMN mode; ALTER TABLE rutba_linked DROP COLUMN stamp; ALTER TABLE rutba_linked DROP COLUMN synced; UPDATE rutba_meta SET value = '6' WHERE key = 'version'");
  });
  const six = checkRbase(file);
  assert.deepEqual(six.problems, [], 'a version 6 file without them follows the text');
  assert.equal(six.version, 6);
  const b = Base.open(file);
  try {
    assert.equal(b.meta('version'), String(DESIGN_VERSION));
    assert.equal(b.count('Orders'), 5);
  } finally {
    b.close();
  }
  const r = checkRbase(file);
  assert.deepEqual(r.problems, []);
  assert.equal(r.version, 7);
});

test('what a newer version keeps that this one does not know is kept: rutba_fields columns, query keys, role rights, link settings', () => {
  const file = spoiled('query-form-report.rbase', (db) => {
    db.exec("ALTER TABLE rutba_fields ADD COLUMN format TEXT; UPDATE rutba_fields SET format = 'upper' WHERE tbl = 'Customers' AND name = 'Name'; UPDATE rutba_fields SET format = 'lower' WHERE tbl = 'Customers' AND name = 'Email'");
    db.exec("UPDATE rutba_queries SET design = json_set(design, '$.layout', json('{\"grid\":\"wide\"}'), '$.columns[0].width', 140, '$.columns[0].as', 'Customer') WHERE name = 'Order Lines'");
  });
  const b = Base.open(file);
  try {
    // A field renamed, another added (in place), one made required (the table made again): the unknown column stays.
    b.changeField('Customers', 'Name', { name: 'Full Name' });
    b.addField('Customers', { name: 'Phone', type: 'text' });
    b.changeField('Customers', 'Email', { required: true });
    const format = (n) => b.db.prepare("SELECT format FROM rutba_fields WHERE tbl = 'Customers' AND name = ?").get(n)?.format ?? null;
    assert.equal(format('Full Name'), 'upper');
    assert.equal(format('Email'), 'lower');
    assert.equal(format('Phone'), null);
    // The query saved again as the designer saves it: its own keys and each column's kept.
    const saved = b.queries().find((x) => x.name === 'Order Lines').design;
    const held = saved.columns.map((c) => ({ ...c, exprText: c.expr ? (c.as ? `${c.as}: ${c.expr}` : c.expr) : null, criteria: c.criteria || [] }));
    held[2].criteria = ['> 1'];
    const again = designerDesign(saved.tables, held, false, saved);
    assert.deepEqual(again.layout, { grid: 'wide' });
    assert.equal(again.columns[0].width, 140);
    assert.equal(again.columns[0].as, 'Customer', 'a field column\'s heading, which the designer does not set, is kept');
    assert.equal(again.columns[3].as, 'Amount');
    assert.equal(again.columns[3].expr, 'Round([Quantity] * [Unit Price], 2)');
    assert.ok(!('exprText' in again.columns[0]), 'the designer\'s own state is not written');
    b.saveQuery('Order Lines', again);
    assert.deepEqual(b.queries().find((x) => x.name === 'Order Lines').design.layout, { grid: 'wide' });
    // A role's rights and an engine link's settings saved again keep the keys this version does not name.
    b.protect({ name: 'Owner', password: 'example-password' });
    b.saveRole('Viewer', { tables: { default: 'read' } });
    b.db.exec("UPDATE rutba_roles SET rights = json_set(rights, '$.expires', '2027-01-01', '$.tables.colour', 'blue') WHERE name = 'Viewer'");
    b.saveRole('Viewer', { tables: { default: 'change' } });
    const rights = JSON.parse(b.db.prepare("SELECT rights FROM rutba_roles WHERE name = 'Viewer'").get().rights);
    assert.equal(rights.expires, '2027-01-01');
    assert.equal(rights.tables.colour, 'blue');
    assert.equal(rights.tables.default, 'change');
    const link = b.saveLink({ name: 'Head office', engine: 'postgres', settings: { host: 'a.example.com', port: 5432, database: 'shop', user: 'reader', ssl: true } });
    b.db.prepare("UPDATE rutba_links SET settings = json_set(settings, '$.timeout', 30, '$.password', 'not-kept') WHERE id = ?").run(link);
    b.saveLink({ id: link, name: 'Head office', engine: 'postgres', settings: { host: 'b.example.com', port: 5432, database: 'shop', user: 'reader', ssl: true } });
    const settings = b.links().find((l) => l.id === link).settings;
    assert.equal(settings.timeout, 30);
    assert.equal(settings.host, 'b.example.com');
    assert.ok(!('password' in settings), 'a secret is never carried into the file');
  } finally {
    b.close();
  }
});

test('unknownKeys answers what a list of known keys leaves out', () => {
  assert.deepEqual(unknownKeys({ a: 1, b: 2, c: 3 }, ['a']), { b: 2, c: 3 });
  assert.deepEqual(unknownKeys(null, ['a']), {});
  assert.deepEqual(unknownKeys([1, 2], []), {});
});

/* ── the guard: the design tables and the text move together ─────────── */

/** Every `CREATE TABLE IF NOT EXISTS rutba_*` in packages/database/src, as { table: [columns] }. */
function createdInCode() {
  const dir = path.join(ROOT, 'packages', 'database', 'src');
  const out = {};
  for (const f of fs.readdirSync(dir).filter((x) => x.endsWith('.js'))) {
    const src = fs.readFileSync(path.join(dir, f), 'utf8');
    for (const m of src.matchAll(/CREATE TABLE IF NOT EXISTS (rutba_\w+) \(([\s\S]*?)\);/g)) {
      const parts = [];
      let depth = 0;
      let cur = '';
      for (const ch of m[2]) {
        if (ch === '(') depth++;
        if (ch === ')') depth--;
        if (ch === ',' && depth === 0) { parts.push(cur); cur = ''; } else cur += ch;
      }
      parts.push(cur);
      out[m[1]] = parts.map((p) => p.trim()).filter((p) => !/^(PRIMARY|UNIQUE|FOREIGN|CHECK|CONSTRAINT)\b/i.test(p)).map((p) => p.split(/\s+/)[0]);
    }
  }
  return out;
}

/** The design tables of a database the code makes, with every table and column it can add. */
function madeByCode() {
  const base = Base.open(':memory:');
  try {
    ensureSync(base);
    const names = base.db.prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name LIKE 'rutba\\_%' ESCAPE '\\'").all().map((r) => r.name);
    return Object.fromEntries(names.map((n) => [n, base.db.prepare(`PRAGMA table_info(${n})`).all().map((c) => c.name)]));
  } finally {
    base.close();
  }
}

/** The tables the text describes, as { table: [columns] }: each `### rutba_*` section's first table of `| \`column\` |` rows. */
function describedInSpec() {
  const out = {};
  let table = null;
  for (const line of SPEC.split(/\r?\n/)) {
    const h = /^(#+)\s+(.*)$/.exec(line);
    if (h) {
      table = h[1] === '###' && /^rutba_\w+$/.test(h[2].trim()) ? h[2].trim() : null;
      if (table) out[table] = [];
      continue;
    }
    const row = table && /^\|\s*`([^`]+)`\s*\|/.exec(line);
    if (row) out[table].push(row[1]);
  }
  return out;
}

test('every design table and column in packages/database is named in docs/DATA-FILES-FORMAT.md, and nothing more', () => {
  const code = createdInCode();
  const made = madeByCode();
  const spec = describedInSpec();
  // What the code can make: the CREATE TABLE statements, and the columns it adds to them later (ALTER TABLE).
  assert.deepEqual(Object.keys(made).sort(), Object.keys(code).sort(), 'a database the code makes has exactly the tables its CREATE TABLE statements name');
  const all = Object.fromEntries(Object.keys(code).map((t) => [t, [...new Set([...code[t], ...made[t]])]]));
  assert.deepEqual(Object.keys(spec).sort(), Object.keys(all).sort(), 'the spec has a ### section for each rutba_* table the code makes, and none other');
  for (const [t, columns] of Object.entries(all)) {
    assert.deepEqual([...spec[t]].sort(), [...columns].sort(), `the spec's columns of ${t} are the code's`);
    // Each section quotes the statement the code runs.
    const stmt = new RegExp(`CREATE TABLE IF NOT EXISTS ${t} \\(`);
    assert.match(SPEC, stmt, `the spec quotes ${t}'s CREATE TABLE`);
  }
  // The checker knows the same tables and columns.
  assert.deepEqual(Object.keys(DESIGN_TABLES).sort(), Object.keys(all).sort());
  for (const [t, d] of Object.entries(DESIGN_TABLES)) assert.deepEqual([...d.columns, ...(d.later || [])].sort(), [...all[t]].sort(), `the checker's columns of ${t}`);
});

test('the text, the checker and the code agree on the design version and the field types', () => {
  assert.equal(SPEC_VERSION, DESIGN_VERSION, 'the checker knows the design version the code writes');
  assert.match(SPEC, new RegExp(`\\*\\*Status\\.\\*\\* Specification version ${DESIGN_VERSION}\\.`), 'the spec says which version it is');
  assert.match(SPEC, new RegExp(`\\| ${DESIGN_VERSION} \\| Rutba Office`), 'the spec says what the version added');
  assert.deepEqual(SPEC_FIELD_TYPES, FIELD_TYPES);
  for (const t of FIELD_TYPES) assert.match(SPEC, new RegExp(`\\| \`${t}\` \\| \`"F"`), `the spec gives the column of a ${t} field`);
});
