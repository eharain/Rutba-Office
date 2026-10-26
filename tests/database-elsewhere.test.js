// Databases made elsewhere, as the Data Files window opens them (A2.20):
// another program's SQLite database opens read only and is left as it is,
// changes to its records only once the person allows them, and a Data Files
// copy of it is a database of ours; one made by a newer Rutba Office is
// refused before anything in it is changed. The engine's side is in
// tests/data-files-format.test.js.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { DESIGN_VERSION } from '../packages/database/src/index.js';
import { checkRbase } from '../tools/check-rbase.mjs';
import { createBaseService } from '../apps/desktop/main/base.js';

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'base-elsewhere-'));

function otherProgram(dir) {
  const file = path.join(dir, 'club.db');
  const db = new DatabaseSync(file);
  try {
    db.exec(`CREATE TABLE members (id INTEGER PRIMARY KEY, name TEXT NOT NULL, joined DATE);
      CREATE TABLE fees (id INTEGER PRIMARY KEY, member INTEGER REFERENCES members(id), amount REAL);
      CREATE VIRTUAL TABLE IF NOT EXISTS notes USING fts5(body);
      INSERT INTO members VALUES (1, 'Asha', '2025-04-01'), (2, 'Ben', '2026-01-15');
      INSERT INTO fees (member, amount) VALUES (1, 20), (2, 25.5);`);
  } catch {
    // An SQLite without full-text search: the rest of the file is the same.
    db.exec(`CREATE TABLE IF NOT EXISTS members (id INTEGER PRIMARY KEY, name TEXT NOT NULL, joined DATE);
      CREATE TABLE IF NOT EXISTS fees (id INTEGER PRIMARY KEY, member INTEGER REFERENCES members(id), amount REAL);`);
  } finally {
    db.close();
  }
  return file;
}

test('another program\'s database opens read only, allows changes when asked, and becomes a Data Files copy on Save As', async () => {
  const dir = tmp();
  const file = otherProgram(dir);
  const before = fs.readFileSync(file);
  const svc = createBaseService({});
  try {
    const s = await svc.open({ file });
    assert.equal(s.foreign.write, false);
    assert.deepEqual(s.tables.map((t) => [t.name, t.count, t.readOnly]), [['members', 2, true], ['fees', 2, true]]);
    assert.equal(s.relationships.length, 1);
    assert.ok(svc.fields({ id: s.id, table: 'members' }).every((f) => f.readOnly));
    await assert.rejects(() => svc.insert({ id: s.id, table: 'members', values: { name: 'Cal' } }), /open read only/);
    assert.throws(() => svc.createTable({ id: s.id, name: 'Rooms', fields: [] }), /its design is not changed here/);
    svc.close({ id: s.id });
    assert.ok(Buffer.from(fs.readFileSync(file)).equals(before), 'opened and closed, the file is as it was');

    const again = await svc.open({ file });
    const w = svc.allowChanges({ id: again.id });
    assert.equal(w.foreign.write, true);
    assert.ok(!w.tables[0].readOnly);
    await svc.update({ id: again.id, table: 'members', rowid: 2, values: { name: 'Benedict' } });
    const copyFile = path.join(dir, 'club.rbase');
    const copy = svc.saveAs({ id: again.id, file: copyFile });
    assert.equal(copy.foreign, null, 'the copy is a database of ours');
    assert.deepEqual(copy.copied.left, []);
    assert.equal(copy.copied.relationships, 1);
    assert.deepEqual(copy.tables.map((t) => [t.name, t.count]), [['members', 2], ['fees', 2]]);
    assert.equal(svc.rows({ id: copy.id, table: 'members' }).rows[1].name, 'Benedict');
    svc.close({ id: copy.id });
    assert.deepEqual(checkRbase(copyFile).problems, []);
    const raw = new DatabaseSync(file, { readOnly: true });
    try {
      assert.equal(raw.prepare('SELECT name FROM members WHERE id = 2').get().name, 'Benedict', 'the change went into its own file');
      assert.deepEqual(raw.prepare("SELECT name FROM sqlite_master WHERE name LIKE 'rutba%'").all(), []);
    } finally {
      raw.close();
    }
  } finally {
    svc.closeAll?.();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a database made by a newer Rutba Office is refused as it opens, nothing in it changed', async () => {
  const dir = tmp();
  const file = path.join(dir, 'newer.rbase');
  const db = new DatabaseSync(file);
  db.exec(`CREATE TABLE rutba_meta (key TEXT PRIMARY KEY, value TEXT); CREATE TABLE rutba_tables (name TEXT PRIMARY KEY, caption TEXT, position INTEGER);
    CREATE TABLE rutba_fields (tbl TEXT NOT NULL, name TEXT NOT NULL, type TEXT NOT NULL, size INTEGER, required INTEGER NOT NULL DEFAULT 0, uniq INTEGER NOT NULL DEFAULT 0, dflt TEXT, caption TEXT, position INTEGER NOT NULL, PRIMARY KEY (tbl, name));
    INSERT INTO rutba_meta VALUES ('version', '${DESIGN_VERSION + 2}');`);
  db.close();
  const before = fs.readFileSync(file);
  const svc = createBaseService({});
  try {
    await assert.rejects(() => svc.open({ file }), new RegExp(`made by a newer version of Rutba Office \\(design version ${DESIGN_VERSION + 2}`));
    assert.ok(Buffer.from(fs.readFileSync(file)).equals(before));
    assert.deepEqual(fs.readdirSync(dir), ['newer.rbase']);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
