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
  ensureSync, reconcile, replaceRows, markSynced,
} from '../packages/database/src/index.js';
import { checkRbase, SPEC_VERSION, DESIGN_TABLES, FIELD_TYPES as SPEC_FIELD_TYPES } from '../tools/check-rbase.mjs';
import { makeSamples, SAMPLES_DIR, EXAMPLE_PEOPLE } from '../tools/make-rbase-samples.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SPEC = fs.readFileSync(path.join(ROOT, 'docs', 'DATA-FILES-FORMAT.md'), 'utf8');
const SAMPLES = ['plain-tables.rbase', 'query-form-report.rbase', 'protected.rbase'];
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
    ['a design version newer than the text', 'plain-tables.rbase', (db) => db.exec("UPDATE rutba_meta SET value = '7' WHERE key = 'version'"), /newer than this specification/],
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
  assert.match(good.stdout, /follows the Data Files format, design version 6/);
  const bad = run(spoiled('plain-tables.rbase', (db) => db.exec("UPDATE rutba_fields SET type = 'picture' WHERE name = 'Notes'")));
  assert.equal(bad.status, 1);
  assert.match(bad.stdout, /does not follow the Data Files format \(1 problem\):\n {2}- Field Customers\.Notes has the type "picture"/);
  assert.equal(run().status, 2);
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
