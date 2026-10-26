// Another program's SQLite database (A2.20): a file with tables of its own
// and no design tables. Data Files reads it as it is, read only unless the
// person chooses otherwise, and puts nothing of its own into it: the design
// it shows (each table, its fields and their types, the relationships its
// foreign keys make) is worked out from the file's own schema and kept in
// TEMP tables on this connection only, gone when it closes. A copy made as a
// Data Files database (import.js, `copyForeign`) is how its design comes to
// be changed.

import { q, fitTo } from './types.js';

const RESERVED = /^(rutba_|sqlite_)/i;
/** How many of a column's values are read to see whether they fit the type its declaration names. */
const SAMPLE = 2000;

/**
 * What a file holds, read from its schema: { kind, version }. `kind` is
 * 'data-files' (rutba_meta names a version), 'other' (tables of another
 * program's and no design) or 'empty' (nothing of anyone's yet); `version`
 * the text rutba_meta holds, or null.
 */
export function fileKind(db) {
  const names = db.prepare("SELECT name FROM main.sqlite_master WHERE type = 'table'").all().map((r) => r.name);
  if (names.includes('rutba_meta')) {
    const v = db.prepare("SELECT value FROM main.rutba_meta WHERE key = 'version'").get()?.value;
    if (v !== undefined && v !== null) return { kind: 'data-files', version: String(v) };
  }
  return names.some((n) => !RESERVED.test(n)) ? { kind: 'other', version: null } : { kind: 'empty', version: null };
}

/**
 * The field type a column's declaration names, by SQLite's own rules of
 * affinity with the common names of dates and yes or no: { type, size }.
 */
function declared(col, autoincrement) {
  const d = String(col.type || '').toUpperCase();
  if (col.pk === 1 && d === 'INTEGER' && autoincrement) return { type: 'autonumber' };
  if (/BOOL/.test(d)) return { type: 'yesno' };
  if (/DATETIME|TIMESTAMP/.test(d)) return { type: 'datetime' };
  if (/DATE/.test(d)) return { type: 'date' };
  if (/INT/.test(d)) return { type: 'integer' };
  if (/MONEY|CURRENCY/.test(d)) return { type: 'currency' };
  if (/CHAR|CLOB|TEXT/.test(d)) {
    const size = Number(/\(\s*(\d+)/.exec(d)?.[1]) || 0;
    return size > 0 && size <= 4000 ? { type: 'text', size } : { type: 'longtext' };
  }
  if (/REAL|FLOA|DOUB|DEC|NUM/.test(d)) return { type: 'number' };
  return { type: 'longtext' };
}

/** Whether a value read from the file is one the field keeps as it is: a date as YYYY-MM-DD, a Yes/No as 0 or 1. */
function keeps(field, v) {
  if (v === null || v === undefined) return true;
  if (v instanceof Uint8Array) return false;
  if (field.type === 'longtext') return typeof v === 'string' || typeof v === 'number';
  if (field.type === 'yesno') return v === 0 || v === 1;
  const fit = fitTo(field, v);
  return fit !== undefined && (fit === v || (typeof v === 'number' && Number(fit) === v));
}

/**
 * The tables of another program's database as Data Files shows them: each
 * rowid table whose name a person could give, its fields typed by their
 * declarations where the values agree and as long text where they do not,
 * and the relationships its single-column foreign keys make. Written into
 * the TEMP design tables of `db`; answers { fixed, left }: the fields that
 * hold what a field cannot (pictures and other binary values), read only,
 * as "table\u0000field", and the tables left out, with why.
 */
export function foreignDesign(db) {
  const fixed = new Set();
  const left = [];
  const keys = new Map();
  const tables = db.prepare("SELECT name, sql FROM main.sqlite_master WHERE type = 'table' ORDER BY rowid").all();
  const seq = new Set((() => { try { return db.prepare('SELECT name FROM main.sqlite_sequence').all().map((r) => r.name); } catch { return []; } })());
  // A virtual table's own tables (full-text search keeps several) are part of it, not tables of their own.
  const virtual = tables.filter((t) => /^\s*CREATE\s+VIRTUAL/i.test(String(t.sql || ''))).map((t) => `${t.name}_`.toLowerCase());
  const shadow = new Set((() => { try { return db.prepare('PRAGMA main.table_list').all().filter((r) => r.type === 'shadow').map((r) => r.name); } catch { return []; } })());
  const shown = [];
  for (const t of tables) {
    if (RESERVED.test(t.name) || shadow.has(t.name) || virtual.some((v) => t.name.toLowerCase().startsWith(v))) continue;
    const sql = String(t.sql || '');
    if (/^\s*CREATE\s+VIRTUAL/i.test(sql)) { left.push({ table: t.name, why: 'virtual' }); continue; }
    if (/\bWITHOUT\s+ROWID\s*$/i.test(sql.trim())) { left.push({ table: t.name, why: 'without-rowid' }); continue; }
    if (!t.name.trim() || t.name.length > 64 || /[\u0000-\u001f]/.test(t.name)) { left.push({ table: t.name, why: 'name' }); continue; }
    shown.push(t);
  }
  const putTable = db.prepare('INSERT INTO temp.rutba_tables (name, caption, position) VALUES (?, NULL, ?)');
  const putField = db.prepare('INSERT INTO temp.rutba_fields (tbl, name, type, size, required, uniq, dflt, caption, position) VALUES (?, ?, ?, ?, ?, ?, NULL, NULL, ?)');
  const fieldsOf = new Map();
  shown.forEach((t, position) => {
    const cols = db.prepare(`PRAGMA main.table_info(${q(t.name)})`).all();
    const autoincrement = seq.has(t.name) || /AUTOINCREMENT/i.test(String(t.sql));
    // A column that is unique on its own: its primary key, or a unique index of that column alone.
    const single = new Set(cols.filter((c) => c.pk === 1 && cols.filter((x) => x.pk > 0).length === 1).map((c) => c.name));
    for (const ix of db.prepare(`PRAGMA main.index_list(${q(t.name)})`).all()) {
      if (!ix.unique) continue;
      const of = db.prepare(`PRAGMA main.index_info(${q(ix.name)})`).all();
      if (of.length === 1 && of[0].name) single.add(of[0].name);
    }
    // Its INTEGER PRIMARY KEY, the rowid by another name: a copy keeps its numbers as the AutoNumber.
    const pk = cols.filter((c) => c.pk > 0);
    if (pk.length === 1 && String(pk[0].type).toUpperCase() === 'INTEGER') keys.set(t.name, pk[0].name);
    putTable.run(t.name, position);
    const names = [];
    cols.forEach((c, i) => {
      let f = declared(c, autoincrement);
      const values = db.prepare(`SELECT ${q(c.name)} AS v FROM main.${q(t.name)} WHERE ${q(c.name)} IS NOT NULL LIMIT ${SAMPLE}`).all().map((r) => r.v);
      if (values.some((v) => v instanceof Uint8Array)) fixed.add(`${t.name}\u0000${c.name}`);
      else if (f.type !== 'autonumber' && !values.every((v) => keeps(f, v))) f = { type: 'longtext' };
      putField.run(t.name, c.name, f.type, f.size ?? null, c.notnull && c.dflt_value === null && f.type !== 'autonumber' ? 1 : 0, single.has(c.name) ? 1 : 0, i);
      names.push(c.name);
    });
    fieldsOf.set(t.name, { names, single });
  });
  // Its foreign keys of one column, between tables shown, as relationships, so queries can join them.
  const putRelation = db.prepare('INSERT INTO temp.rutba_relations (from_tbl, from_field, to_tbl, to_field, cascade_update, cascade_delete) VALUES (?, ?, ?, ?, ?, ?)');
  for (const t of shown) {
    const fks = db.prepare(`PRAGMA main.foreign_key_list(${q(t.name)})`).all();
    for (const id of new Set(fks.map((k) => k.id))) {
      const parts = fks.filter((k) => k.id === id);
      if (parts.length !== 1) continue;
      const k = parts[0];
      const to = fieldsOf.get(k.table);
      if (!to) continue;
      const toField = k.to || db.prepare(`PRAGMA main.table_info(${q(k.table)})`).all().find((c) => c.pk === 1)?.name;
      if (!toField || !to.names.includes(toField) || !fieldsOf.get(t.name).names.includes(k.from)) continue;
      putRelation.run(t.name, k.from, k.table, toField, k.on_update === 'CASCADE' ? 1 : 0, k.on_delete === 'CASCADE' ? 1 : 0);
    }
  }
  return { fixed, left, keys };
}
