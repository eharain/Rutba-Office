// A database of the database app: one SQLite file, its data in ordinary
// tables any SQLite tool can read, and its design (captions, field types,
// relationships, the history of its changes) in tables of its own named
// rutba_*. Every change to the design is a numbered step, applied to the
// data in place in one transaction: nothing is copied over anything, and an
// AutoNumber is never renumbered.

import { DatabaseSync } from 'node:sqlite';
import { FIELD_TYPES, q, columnSql, toStored } from './types.js';
import { designToSql } from './query-design.js';
import { wizardForm, shownField } from './form-design.js';
import { wizardReport, reportData } from './report-design.js';
import { wizardDocument, documentData, PAPERS, BLOCKS } from './document-design.js';
import { filterSql, sniffType } from './filter.js';
import { ACCESS_DESIGN, ADMIN, STARTING_ROLES, normalRights, tableRight, allows, viewSql, hashPassword, checkPassword } from './access.js';

const DESIGN = `
CREATE TABLE IF NOT EXISTS rutba_meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS rutba_tables (name TEXT PRIMARY KEY, caption TEXT, position INTEGER);
CREATE TABLE IF NOT EXISTS rutba_fields (tbl TEXT NOT NULL, name TEXT NOT NULL, type TEXT NOT NULL, size INTEGER, required INTEGER NOT NULL DEFAULT 0,
  uniq INTEGER NOT NULL DEFAULT 0, dflt TEXT, caption TEXT, position INTEGER NOT NULL, PRIMARY KEY (tbl, name));
CREATE TABLE IF NOT EXISTS rutba_relations (id INTEGER PRIMARY KEY, from_tbl TEXT NOT NULL, from_field TEXT NOT NULL, to_tbl TEXT NOT NULL,
  to_field TEXT NOT NULL, cascade_update INTEGER NOT NULL DEFAULT 1, cascade_delete INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS rutba_steps (n INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT NOT NULL, label TEXT NOT NULL, detail TEXT);
CREATE TABLE IF NOT EXISTS rutba_queries (name TEXT PRIMARY KEY, design TEXT NOT NULL, position INTEGER);
CREATE TABLE IF NOT EXISTS rutba_forms (name TEXT PRIMARY KEY, layout TEXT NOT NULL, position INTEGER);
CREATE TABLE IF NOT EXISTS rutba_reports (name TEXT PRIMARY KEY, layout TEXT NOT NULL, position INTEGER);
CREATE TABLE IF NOT EXISTS rutba_documents (name TEXT PRIMARY KEY, layout TEXT NOT NULL, position INTEGER);
${ACCESS_DESIGN}`;

/**
 * The version of the design tables this build writes: 2 keeps saved queries
 * and forms, 3 reports, 4 print templates, 5 people and roles. An older file
 * is given the tables it lacks as it opens.
 */
export const DESIGN_VERSION = 5;

const RESERVED = /^(rutba_|sqlite_)/i;

/** A table or field name a person may give: words, not empty, not one of the file's own. */
function checkName(name, what) {
  const n = String(name ?? '').trim();
  if (!n) throw new Error(`A ${what} needs a name.`);
  if (n.length > 64) throw new Error(`A ${what}'s name is at most 64 characters.`);
  if (RESERVED.test(n)) throw new Error(`${n} is a name the database keeps for itself.`);
  if (/["\u0000-\u001f]/.test(n)) throw new Error(`A ${what}'s name cannot hold quotation marks or control characters.`);
  return n;
}

// A sum of money in binary floating point comes out as 31.240000000000002;
// fifteen significant digits, all a double holds for certain, make it 31.24.
const clean = (v) => (typeof v === 'number' && !Number.isInteger(v) ? Number(v.toPrecision(15)) : v);
function cleanRow(r) {
  for (const k of Object.keys(r)) r[k] = clean(r[k]);
  return r;
}

export class Base {
  /** Open (or make) the database at `file`; ':memory:' for one in memory. */
  static open(file) {
    const db = new DatabaseSync(file);
    if (file !== ':memory:') db.exec('PRAGMA journal_mode = WAL');
    db.exec('PRAGMA foreign_keys = ON');
    db.exec(DESIGN);
    const base = new Base(db);
    if (Number(base.meta('version') || 0) < DESIGN_VERSION) base.setMeta('version', String(DESIGN_VERSION));
    return base;
  }

  /**
   * The database at `file` to read and nothing else, as a mail merge reads
   * it: its design left as it is, and any change refused by SQLite itself.
   */
  static read(file) {
    const db = new DatabaseSync(file);
    db.exec('PRAGMA query_only = ON');
    return new Base(db);
  }

  /** Whether the file has the design table `name` (a file made before queries has no rutba_queries). */
  hasDesignTable(name) {
    return Boolean(this.db.prepare("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?").get(name));
  }

  constructor(db) {
    this.db = db;
    // The day Date() and Now() answer in this database's expressions; null is today.
    this.today = null;
    // Who is using it (D-2026-10-22-11): null while it is not protected; locked until someone signs in.
    this.access = null;
    try { if (this.meta('protected') === '1') this.access = { locked: true }; } catch { /* a file with no design tables */ }
  }

  close() {
    try { this.db.exec('PRAGMA wal_checkpoint(TRUNCATE)'); } catch { /* in memory */ }
    this.db.close();
  }

  meta(key) {
    return this.db.prepare('SELECT value FROM rutba_meta WHERE key = ?').get(key)?.value ?? null;
  }

  setMeta(key, value) {
    this.db.prepare('INSERT INTO rutba_meta (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value').run(key, value);
  }

  /* ── design ─────────────────────────────────────────────────────── */

  /** The tables, in their order: [{ name, caption }]; those the person's role hides left out. */
  tables() {
    if (this.access?.locked) return [];
    const all = this._allTables();
    const rights = this._rights();
    return rights ? all.filter((t) => tableRight(rights, t.name).access !== 'hidden') : all;
  }

  _allTables() {
    return this.db.prepare('SELECT name, caption FROM rutba_tables ORDER BY position, name').all().map((r) => ({ ...r }));
  }

  /**
   * A table's fields, in their order: [{ name, type, size, required, unique,
   * default, caption }]; those the role hides left out, and those it may only
   * read marked readOnly.
   */
  fields(table) {
    if (this.access?.locked) return [];
    const all = this._allFields(table);
    const rights = this._rights();
    if (!rights) return all;
    const right = tableRight(rights, table);
    if (right.access === 'hidden') return [];
    const readOnly = right.access === 'read';
    return all.filter((f) => right.fields[f.name] !== 'hidden').map((f) => (readOnly || right.fields[f.name] === 'read' ? { ...f, readOnly: true } : f));
  }

  _allFields(table) {
    return this.db.prepare('SELECT * FROM rutba_fields WHERE tbl = ? ORDER BY position').all(table).map((f) => ({
      name: f.name, type: f.type, size: f.size ?? undefined, required: Boolean(f.required), unique: Boolean(f.uniq),
      default: f.dflt == null ? undefined : JSON.parse(f.dflt), caption: f.caption ?? undefined,
    }));
  }

  /** The relationships: [{ id, from: { table, field }, to: { table, field }, cascadeUpdate, cascadeDelete }]. */
  relationships() {
    return this.db.prepare('SELECT * FROM rutba_relations ORDER BY id').all().map((r) => ({
      id: r.id, from: { table: r.from_tbl, field: r.from_field }, to: { table: r.to_tbl, field: r.to_field },
      cascadeUpdate: Boolean(r.cascade_update), cascadeDelete: Boolean(r.cascade_delete),
    }));
  }

  /** The design's history: [{ n, at, label }], the latest last. */
  steps() {
    return this.db.prepare('SELECT n, at, label FROM rutba_steps ORDER BY n').all().map((r) => ({ ...r }));
  }

  /** The CREATE TABLE for `table` with `fields` and the relationships from it. */
  _tableSql(table, fields, relations = this.relationships()) {
    const cols = fields.map(columnSql);
    for (const r of relations.filter((x) => x.from.table === table)) {
      cols.push(`FOREIGN KEY (${q(r.from.field)}) REFERENCES ${q(r.to.table)} (${q(r.to.field)})`
        + ` ON UPDATE ${r.cascadeUpdate ? 'CASCADE' : 'RESTRICT'} ON DELETE ${r.cascadeDelete ? 'CASCADE' : 'RESTRICT'}`);
    }
    return `CREATE TABLE ${q(table)} (${cols.join(', ')})`;
  }

  /** One change to the design, as one transaction and one numbered step. */
  _step(label, detail, run) {
    this._mayDesign();
    this.db.exec('BEGIN IMMEDIATE');
    try {
      const out = run();
      this.db.prepare('INSERT INTO rutba_steps (at, label, detail) VALUES (?, ?, ?)').run(new Date().toISOString(), label, detail == null ? null : JSON.stringify(detail));
      this.db.exec('COMMIT');
      return out;
    } catch (err) {
      this.db.exec('ROLLBACK');
      throw err;
    }
  }

  _writeFields(table, fields) {
    this.db.prepare('DELETE FROM rutba_fields WHERE tbl = ?').run(table);
    const put = this.db.prepare('INSERT INTO rutba_fields (tbl, name, type, size, required, uniq, dflt, caption, position) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)');
    fields.forEach((f, i) => put.run(table, f.name, f.type, f.size ?? null, f.required ? 1 : 0, f.unique ? 1 : 0, f.default === undefined ? null : JSON.stringify(f.default), f.caption ?? null, i));
  }

  /** A field as the design keeps it, its name and type checked. */
  _field(f) {
    const name = checkName(f.name, 'field');
    const type = FIELD_TYPES.includes(f.type) ? f.type : 'text';
    return { ...f, name, type };
  }

  /** Make a table: { name, fields: [field], caption }. Without an AutoNumber, one called ID is put first. */
  createTable({ name, fields = [], caption }) {
    const table = checkName(name, 'table');
    if (this.tables().some((t) => t.name.toLowerCase() === table.toLowerCase())) throw new Error(`There is a table called ${table} already.`);
    let list = fields.map((f) => this._field(f));
    if (new Set(list.map((f) => f.name.toLowerCase())).size !== list.length) throw new Error('Two fields cannot have the same name.');
    if (list.filter((f) => f.type === 'autonumber').length > 1) throw new Error('A table has one AutoNumber at most.');
    if (!list.some((f) => f.type === 'autonumber')) list = [{ name: list.some((f) => f.name.toLowerCase() === 'id') ? `${table} ID` : 'ID', type: 'autonumber' }, ...list];
    return this._step(`Create table ${table}`, { table, fields: list }, () => {
      this.db.exec(this._tableSql(table, list, []));
      const position = this.db.prepare('SELECT coalesce(max(position), -1) + 1 AS p FROM rutba_tables').get().p;
      this.db.prepare('INSERT INTO rutba_tables (name, caption, position) VALUES (?, ?, ?)').run(table, caption ?? null, position);
      this._writeFields(table, list);
      return table;
    });
  }

  /**
   * The table made again with `fields` (each with `from`, the field it was,
   * when it was one), its rows copied across in place: how SQLite changes a
   * column's type, removes one, or adds a relationship.
   */
  _rebuild(table, fields, relations) {
    const tmp = `rutba_rebuild_${Date.now()}`;
    const keep = fields.filter((f) => f.from);
    this.db.exec(this._tableSql(tmp, fields.map(({ from, ...f }) => f), relations.map((r) => (r.from.table === table ? { ...r, from: { ...r.from, table: tmp } } : r))));
    if (keep.length) {
      const cast = (f) => (['number', 'currency'].includes(f.type) ? `CAST(${q(f.from)} AS REAL)` : f.type === 'integer' ? `CAST(${q(f.from)} AS INTEGER)` : q(f.from));
      this.db.exec(`INSERT INTO ${q(tmp)} (${keep.map((f) => q(f.name)).join(', ')}) SELECT ${keep.map(cast).join(', ')} FROM ${q(table)}`);
    }
    // The AutoNumber's counter is the table's, not its rows': a number given once is never given again.
    const seq = this.db.prepare('SELECT seq FROM sqlite_sequence WHERE name = ?').get(table)?.seq ?? null;
    this.db.exec(`DROP TABLE ${q(table)}`);
    this.db.exec(`ALTER TABLE ${q(tmp)} RENAME TO ${q(table)}`);
    if (seq != null && fields.some((f) => f.type === 'autonumber')) {
      const now = this.db.prepare('SELECT seq FROM sqlite_sequence WHERE name = ?').get(table);
      if (now) this.db.prepare('UPDATE sqlite_sequence SET seq = max(seq, ?) WHERE name = ?').run(seq, table);
      else this.db.prepare('INSERT INTO sqlite_sequence (name, seq) VALUES (?, ?)').run(table, seq);
    }
  }

  /** Run a rebuild with foreign keys checked at the end, not row by row. */
  _withRebuild(label, detail, run) {
    this.db.exec('PRAGMA foreign_keys = OFF');
    try {
      return this._step(label, detail, () => {
        const out = run();
        const broken = this.db.prepare('PRAGMA foreign_key_check').all();
        if (broken.length) throw new Error(`That would leave ${broken.length} record${broken.length === 1 ? '' : 's'} in ${broken[0].table} pointing at nothing.`);
        return out;
      });
    } finally {
      this.db.exec('PRAGMA foreign_keys = ON');
    }
  }

  /** Add a field at the end of a table. */
  addField(table, field) {
    const f = this._field(field);
    const fields = this.fields(table);
    if (!fields.length) throw new Error(`There is no table called ${table}.`);
    if (fields.some((x) => x.name.toLowerCase() === f.name.toLowerCase())) throw new Error(`${table} has a field called ${f.name} already.`);
    if (f.type === 'autonumber') throw new Error('An AutoNumber is made with the table.');
    const next = [...fields, f];
    // A plain column ALTERs in place; one that is required or unique has the table made again.
    if (!f.required && !f.unique) {
      return this._step(`Add field ${f.name} to ${table}`, { table, field: f }, () => {
        this.db.exec(`ALTER TABLE ${q(table)} ADD COLUMN ${columnSql(f)}`);
        this._writeFields(table, next);
      });
    }
    return this._withRebuild(`Add field ${f.name} to ${table}`, { table, field: f }, () => {
      this._rebuild(table, [...fields.map((x) => ({ ...x, from: x.name })), f], this.relationships());
      this._writeFields(table, next);
    });
  }

  /** Change a field: its name, type, size, required, unique, default or caption. Its data comes with it. */
  changeField(table, name, change) {
    const fields = this.fields(table);
    const i = fields.findIndex((f) => f.name === name);
    if (i < 0) throw new Error(`${table} has no field called ${name}.`);
    const was = fields[i];
    const now = this._field({ ...was, ...change });
    if (now.name !== was.name && fields.some((f, k) => k !== i && f.name.toLowerCase() === now.name.toLowerCase())) throw new Error(`${table} has a field called ${now.name} already.`);
    if ((was.type === 'autonumber') !== (now.type === 'autonumber')) throw new Error('A table keeps its AutoNumber.');
    const next = fields.map((f, k) => (k === i ? now : f));
    const onlyWords = ['name', 'type', 'size', 'required', 'unique', 'default'].every((k) => k === 'name' || JSON.stringify(was[k]) === JSON.stringify(now[k]));
    const relations = this.relationships().map((r) => ({
      ...r,
      from: r.from.table === table && r.from.field === was.name ? { ...r.from, field: now.name } : r.from,
      to: r.to.table === table && r.to.field === was.name ? { ...r.to, field: now.name } : r.to,
    }));
    const saveRelations = () => {
      const put = this.db.prepare('UPDATE rutba_relations SET from_field = ?, to_field = ? WHERE id = ?');
      for (const r of relations) put.run(r.from.field, r.to.field, r.id);
    };
    if (onlyWords) {
      return this._step(now.name !== was.name ? `Rename field ${was.name} to ${now.name} in ${table}` : `Change field ${was.name} in ${table}`, { table, from: was, to: now }, () => {
        if (now.name !== was.name) this.db.exec(`ALTER TABLE ${q(table)} RENAME COLUMN ${q(was.name)} TO ${q(now.name)}`);
        this._writeFields(table, next);
        saveRelations();
      });
    }
    return this._withRebuild(`Change field ${was.name} in ${table}`, { table, from: was, to: now }, () => {
      this._rebuild(table, next.map((f, k) => ({ ...f, from: k === i ? was.name : f.name })), relations);
      this._writeFields(table, next);
      saveRelations();
    });
  }

  /** Take a field away, and its data. One a relationship uses stays until the relationship goes. */
  removeField(table, name) {
    const fields = this.fields(table);
    const f = fields.find((x) => x.name === name);
    if (!f) throw new Error(`${table} has no field called ${name}.`);
    if (f.type === 'autonumber') throw new Error('A table keeps its AutoNumber.');
    if (this.relationships().some((r) => (r.from.table === table && r.from.field === name) || (r.to.table === table && r.to.field === name))) {
      throw new Error(`${name} is in a relationship; take the relationship away first.`);
    }
    const next = fields.filter((x) => x !== f);
    return this._withRebuild(`Remove field ${name} from ${table}`, { table, field: f }, () => {
      this._rebuild(table, next.map((x) => ({ ...x, from: x.name })), this.relationships());
      this._writeFields(table, next);
    });
  }

  /** Rename a table; its relationships follow it. */
  renameTable(table, to) {
    const name = checkName(to, 'table');
    if (this.tables().some((t) => t.name.toLowerCase() === name.toLowerCase() && t.name !== table)) throw new Error(`There is a table called ${name} already.`);
    return this._step(`Rename table ${table} to ${name}`, { from: table, to: name }, () => {
      this.db.exec(`ALTER TABLE ${q(table)} RENAME TO ${q(name)}`);
      this.db.prepare('UPDATE rutba_tables SET name = ? WHERE name = ?').run(name, table);
      this.db.prepare('UPDATE rutba_fields SET tbl = ? WHERE tbl = ?').run(name, table);
      this.db.prepare('UPDATE rutba_relations SET from_tbl = ? WHERE from_tbl = ?').run(name, table);
      this.db.prepare('UPDATE rutba_relations SET to_tbl = ? WHERE to_tbl = ?').run(name, table);
    });
  }

  /** Delete a table and its records. One another table points at stays until that relationship goes. */
  dropTable(table) {
    if (this.relationships().some((r) => r.to.table === table && r.from.table !== table)) throw new Error(`Another table points at ${table}; take that relationship away first.`);
    return this._step(`Delete table ${table}`, { table }, () => {
      this.db.prepare('DELETE FROM rutba_relations WHERE from_tbl = ?').run(table);
      this.db.exec(`DROP TABLE ${q(table)}`);
      this.db.prepare('DELETE FROM rutba_tables WHERE name = ?').run(table);
      this.db.prepare('DELETE FROM rutba_fields WHERE tbl = ?').run(table);
    });
  }

  /**
   * A relationship: many records of `from.table` to one of `to.table`, kept
   * by a real foreign key. `to.field` must be unique there (its AutoNumber,
   * usually). Records that already point at nothing refuse it.
   */
  addRelationship({ from, to, cascadeUpdate = true, cascadeDelete = false }) {
    const fromField = this.fields(from.table).find((f) => f.name === from.field);
    const toField = this.fields(to.table).find((f) => f.name === to.field);
    if (!fromField || !toField) throw new Error('A relationship joins two fields that are there.');
    if (!(toField.type === 'autonumber' || toField.unique)) throw new Error(`${to.field} in ${to.table} must be unique for others to point at it.`);
    const relations = [...this.relationships(), { id: 0, from, to, cascadeUpdate, cascadeDelete }];
    return this._withRebuild(`Relate ${from.table}.${from.field} to ${to.table}.${to.field}`, { from, to, cascadeUpdate, cascadeDelete }, () => {
      this._rebuild(from.table, this.fields(from.table).map((f) => ({ ...f, from: f.name })), relations);
      return Number(this.db.prepare('INSERT INTO rutba_relations (from_tbl, from_field, to_tbl, to_field, cascade_update, cascade_delete) VALUES (?, ?, ?, ?, ?, ?)')
        .run(from.table, from.field, to.table, to.field, cascadeUpdate ? 1 : 0, cascadeDelete ? 1 : 0).lastInsertRowid);
    });
  }

  removeRelationship(id) {
    const r = this.relationships().find((x) => x.id === id);
    if (!r) return;
    const rest = this.relationships().filter((x) => x.id !== id);
    return this._withRebuild(`Unrelate ${r.from.table}.${r.from.field} from ${r.to.table}.${r.to.field}`, { id }, () => {
      this._rebuild(r.from.table, this.fields(r.from.table).map((f) => ({ ...f, from: f.name })), rest);
      this.db.prepare('DELETE FROM rutba_relations WHERE id = ?').run(id);
    });
  }

  /* ── data ───────────────────────────────────────────────────────── */

  _known(table) {
    this._unlocked();
    const fields = this.fields(table);
    if (!fields.length) throw new Error(`There is no table called ${table}.`);
    return fields;
  }

  /** How many records, with `where` ({ field: value }) narrowing them. */
  count(table, where = {}, { search = '', filters = [] } = {}) {
    const fields = this._known(table);
    const keys = Object.keys(where);
    const f = filterSql(fields.map((x) => ({ name: x.name, type: x.type })), { search, filters });
    const conds = [...keys.map((k) => `${q(k)} IS ?`), ...(f.sql ? [f.sql] : [])];
    return this.db.prepare(`SELECT count(*) AS n FROM ${q(table)}${conds.length ? ` WHERE ${conds.join(' AND ')}` : ''}`).get(...keys.map((k) => where[k]), ...f.params).n;
  }

  /** A page of records: [{ _rowid, ...fields }], `order` a field (prefixed with - for descending). */
  rows(table, { offset = 0, limit = 200, order = null, where = {}, search = '', filters = [] } = {}) {
    const fields = this._known(table);
    const keys = Object.keys(where);
    const by = order ? fields.find((f) => f.name === String(order).replace(/^-/, '')) : null;
    // Searched and filtered as a query's results are (filter.js).
    const f = filterSql(fields.map((x) => ({ name: x.name, type: x.type })), { search, filters });
    const conds = [...keys.map((k) => `${q(k)} IS ?`), ...(f.sql ? [f.sql] : [])];
    const sql = `SELECT rowid AS _rowid, * FROM ${q(table)}`
      + (conds.length ? ` WHERE ${conds.join(' AND ')}` : '')
      + ` ORDER BY ${by ? `${q(by.name)} ${String(order).startsWith('-') ? 'DESC' : 'ASC'}, ` : ''}rowid LIMIT ? OFFSET ?`;
    const strip = Boolean(this._rights());
    return this.db.prepare(sql).all(...keys.map((k) => where[k]), ...f.params, Math.max(0, limit | 0), Math.max(0, offset | 0)).map((r) => {
      const row = { ...r };
      // Read through a role's view, whose rowid is a column of its own.
      if (strip) delete row.rowid;
      return row;
    });
  }

  /** Add a record; answers its rowid. Values as a person typed them are put into each field's type. */
  insert(table, values) {
    this._mayWrite(table, 'add');
    const fields = this._known(table).filter((f) => f.type !== 'autonumber' || values[f.name] != null);
    // A field not given takes its default from the table itself: a Yes/No that defaults to Yes is Yes.
    const set = fields.filter((f) => Object.hasOwn(values, f.name) && values[f.name] !== undefined);
    this._noHiddenValues(table, values);
    for (const f of set) if (f.readOnly && values[f.name] != null && values[f.name] !== '') throw new Error(`${f.name} is read only for you.`);
    const given = set.filter((f) => !f.readOnly);
    const sql = given.length
      ? `INSERT INTO ${this._target(table)} (${given.map((f) => q(f.name)).join(', ')}) VALUES (${given.map(() => '?').join(', ')})`
      : `INSERT INTO ${this._target(table)} DEFAULT VALUES`;
    const rowid = Number(this._run(sql, given.map((f) => toStored(f, values[f.name]))).lastInsertRowid);
    this._trail(table, [rowid], 'add');
    return rowid;
  }

  /** Change a record's fields. */
  update(table, rowid, values) {
    this._mayWrite(table, 'change', [rowid]);
    let fields = this._known(table).filter((f) => Object.hasOwn(values, f.name));
    this._noHiddenValues(table, values);
    // A field the role may only read stays as it is: given as it is, it is left out; changed, it is refused.
    const fixed = fields.filter((f) => f.readOnly);
    if (fixed.length) {
      const now = this.db.prepare(`SELECT * FROM main.${q(table)} WHERE rowid = ?`).get(rowid) || {};
      for (const f of fixed) if (toStored(f, values[f.name]) !== now[f.name]) throw new Error(`${f.name} is read only for you.`);
      fields = fields.filter((f) => !f.readOnly);
    }
    if (!fields.length) return 0;
    const changes = Number(this._run(`UPDATE ${this._target(table)} SET ${fields.map((f) => `${q(f.name)} = ?`).join(', ')} WHERE rowid = ?`, [...fields.map((f) => toStored(f, values[f.name])), rowid]).changes);
    if (changes) this._trail(table, [rowid], 'change', { fields: fields.map((f) => f.name) });
    return changes;
  }

  /** Delete records by rowid. */
  remove(table, rowids) {
    const ids = [...rowids];
    this._mayWrite(table, 'delete', ids);
    this._known(table);
    if (!ids.length) return 0;
    const changes = Number(this._run(`DELETE FROM ${this._target(table)} WHERE rowid IN (${ids.map(() => '?').join(', ')})`, ids).changes);
    this._trail(table, ids, 'delete');
    return changes;
  }

  /** A statement, its failures said in words a person can act on. */
  _run(sql, params) {
    try {
      return this.db.prepare(sql).run(...params);
    } catch (err) {
      const m = String(err.message);
      if (/FOREIGN KEY constraint failed/.test(m)) throw new Error('That record is related to others: it points at one that is not there, or others point at it.');
      if (/UNIQUE constraint failed: [^.]+\.(.+)$/.test(m)) throw new Error(`Another record has that ${RegExp.$1} already.`);
      if (/NOT NULL constraint failed: [^.]+\.(.+)$/.test(m)) throw new Error(`${RegExp.$1} needs a value.`);
      if (/CHECK constraint failed/.test(m)) throw new Error('A value does not fit its field: too long, or not the kind the field takes.');
      throw err;
    }
  }

  /* ── saved queries and forms ───────────────────────────────────── */

  /** The saved queries: [{ name, design }], in their order. */
  queries() {
    if (this.access?.locked) return [];
    const all = this.db.prepare('SELECT name, design FROM rutba_queries ORDER BY position, name').all().map((r) => ({ name: r.name, design: JSON.parse(r.design) }));
    return this._shown('queries', all, (x) => x.design.tables || []);
  }

  /** A query kept under `name` (made, or replaced), after its design is checked by building its SQL. */
  saveQuery(name, design) {
    const n = checkName(name, 'query');
    if (this.tables().some((t) => t.name.toLowerCase() === n.toLowerCase())) throw new Error(`${n} is a table's name already. Give the query another.`);
    designToSql(design, this.relationships(), this._designFields());
    const at = this.db.prepare('SELECT COALESCE(MAX(position), -1) + 1 AS p FROM rutba_queries').get().p;
    return this._step(`Query ${n} kept`, { query: n }, () => {
      this.db.prepare('INSERT INTO rutba_queries (name, design, position) VALUES (?, ?, ?) ON CONFLICT(name) DO UPDATE SET design = excluded.design').run(n, JSON.stringify(design), at);
      return n;
    });
  }

  removeQuery(name) {
    return this._step(`Query ${name} removed`, { query: name }, () => this.db.prepare('DELETE FROM rutba_queries WHERE name = ?').run(name).changes > 0);
  }

  /**
   * A query's answer: a saved one by `name`, or a `design` as it stands in
   * the designer. Answers { columns, rows, count, sql }, at most `limit` rows.
   */
  runQuery({ name = null, design = null, limit = 5000, today = this.today } = {}) {
    const d = design || this.queries().find((x) => x.name === name)?.design;
    if (!d) throw new Error(`There is no query called ${name}.`);
    const { sql, params, columns } = designToSql(d, this.relationships(), { ...this._designFields(), today });
    const out = this._select(sql, params);
    const rows = out.rows.slice(0, limit).map(cleanRow);
    return { columns: out.columns.length ? out.columns : columns, rows, count: out.rows.length, sql };
  }

  /**
   * A page of a query's results, searched, filtered and sorted in the
   * database as a table's records are (D-2026-10-22-07): { columns: [{ name,
   * type }], rows, count }, `order` a column's name (- before it for
   * descending), each column typed by the field it shows, or by its values.
   */
  queryRows({ name = null, design = null, offset = 0, limit = 200, order = null, search = '', filters = [], today = this.today } = {}) {
    const d = design || this.queries().find((x) => x.name === name)?.design;
    if (!d) throw new Error(`There is no query called ${name}.`);
    const { sql, params, columns: names, sources } = designToSql(d, this.relationships(), { ...this._designFields(), today });
    const inner = `(${sql}) AS r`;
    // Typed by the field shown; a calculation or a total by its values, read from a few of them.
    const sniffed = new Map();
    const typeOf = (n, i) => {
      const s = sources[i];
      if (s) { const fld = this.fields(s.table).find((x) => x.name === s.field); if (fld) return fld.type; }
      if (!sniffed.has(n)) sniffed.set(n, sniffType(this.db.prepare(`SELECT ${q(n)} AS v FROM ${inner} WHERE ${q(n)} IS NOT NULL LIMIT 50`).all(...params).map((r) => r.v)));
      return sniffed.get(n);
    };
    const columns = names.map((n, i) => ({ name: n, type: typeOf(n, i) }));
    const f = filterSql(columns, { search, filters });
    const by = order ? columns.find((c) => c.name === String(order).replace(/^-/, '')) : null;
    const where = f.sql ? ` WHERE ${f.sql}` : '';
    const count = this.db.prepare(`SELECT count(*) AS n FROM ${inner}${where}`).get(...params, ...f.params).n;
    const rows = this.db.prepare(`SELECT * FROM ${inner}${where}${by ? ` ORDER BY ${q(by.name)} ${String(order).startsWith('-') ? 'DESC' : 'ASC'}` : ''} LIMIT ? OFFSET ?`)
      .all(...params, ...f.params, Math.max(0, limit | 0), Math.max(0, offset | 0)).map(cleanRow);
    return { columns, rows, count };
  }

  /** The SQL a design stands for, its values shown in place of the parameters, for the SQL view. */
  sqlOf(design, { today = this.today } = {}) {
    const { sql, params } = designToSql(design, this.relationships(), { ...this._designFields(), today });
    let i = 0;
    // A ? inside a quoted name or text (an expression's "Is it?") is not a parameter.
    return sql.replace(/'(?:[^']|'')*'|"(?:[^"]|"")*"|\?/g, (m) => {
      if (m !== '?') return m;
      const v = params[i++];
      return typeof v === 'number' ? String(v) : `'${String(v).replace(/'/g, "''")}'`;
    });
  }

  /** A query's tables' fields, for the [Field] names in its expressions. */
  _designFields() {
    const cache = new Map();
    return { fieldsOf: (t) => { if (!cache.has(t)) cache.set(t, this.fields(t).map((f) => f.name)); return cache.get(t); } };
  }

  /** The forms: [{ name, layout }], in their order. */
  forms() {
    if (this.access?.locked) return [];
    const all = this.db.prepare('SELECT name, layout FROM rutba_forms ORDER BY position, name').all().map((r) => ({ name: r.name, layout: JSON.parse(r.layout) }));
    return this._shown('forms', all, (x) => [x.layout.table]);
  }

  /** A form kept (made, or replaced), its table checked. */
  saveForm(layout) {
    const n = checkName(layout?.name, 'form');
    if (!this.tables().some((t) => t.name === layout.table)) throw new Error(`The form's table, ${layout.table}, is not in this database.`);
    const at = this.db.prepare('SELECT COALESCE(MAX(position), -1) + 1 AS p FROM rutba_forms').get().p;
    return this._step(`Form ${n} kept`, { form: n }, () => {
      this.db.prepare('INSERT INTO rutba_forms (name, layout, position) VALUES (?, ?, ?) ON CONFLICT(name) DO UPDATE SET layout = excluded.layout').run(n, JSON.stringify({ ...layout, name: n }), at);
      return n;
    });
  }

  removeForm(name) {
    return this._step(`Form ${name} removed`, { form: name }, () => this.db.prepare('DELETE FROM rutba_forms WHERE name = ?').run(name).changes > 0);
  }

  /** A form for `table` as the wizard makes it, kept under a name of its own. */
  makeForm(table, { name = null } = {}) {
    const taken = new Set(this.forms().map((f) => f.name.toLowerCase()));
    let n = name || table;
    for (let i = 2; taken.has(n.toLowerCase()); i++) n = `${name || table} ${i}`;
    const layout = wizardForm(table, (t) => this.fields(t), this.relationships(), { name: n });
    this.saveForm(layout);
    return layout;
  }

  /** The reports: [{ name, layout }], in their order. */
  reports() {
    if (this.access?.locked) return [];
    const all = this.db.prepare('SELECT name, layout FROM rutba_reports ORDER BY position, name').all().map((r) => ({ name: r.name, layout: JSON.parse(r.layout) }));
    return this._shown('reports', all, (x) => (x.layout.source?.table ? [x.layout.source.table] : []), (x) => x.layout.source?.query);
  }

  /** A report kept (made, or replaced), its source checked. */
  saveReport(layout) {
    const n = checkName(layout?.name, 'report');
    const src = layout.source || {};
    if (src.table && !this.tables().some((t) => t.name === src.table)) throw new Error(`The report's table, ${src.table}, is not in this database.`);
    if (src.query && !this.queries().some((x) => x.name === src.query)) throw new Error(`The report's query, ${src.query}, is not in this database.`);
    const at = this.db.prepare('SELECT COALESCE(MAX(position), -1) + 1 AS p FROM rutba_reports').get().p;
    return this._step(`Report ${n} kept`, { report: n }, () => {
      this.db.prepare('INSERT INTO rutba_reports (name, layout, position) VALUES (?, ?, ?) ON CONFLICT(name) DO UPDATE SET layout = excluded.layout').run(n, JSON.stringify({ ...layout, name: n }), at);
      return n;
    });
  }

  removeReport(name) {
    return this._step(`Report ${name} removed`, { report: name }, () => this.db.prepare('DELETE FROM rutba_reports WHERE name = ?').run(name).changes > 0);
  }

  /** A report for a table or a saved query as the wizard makes it, kept under a name of its own. */
  makeReport({ table = null, query = null, name = null } = {}) {
    const taken = new Set(this.reports().map((r) => r.name.toLowerCase()));
    const want = name || table || query;
    let n = want;
    for (let i = 2; taken.has(n.toLowerCase()); i++) n = `${want} ${i}`;
    const layout = table
      ? wizardReport({ table, fields: this.fields(table), relationships: this.relationships(), name: n })
      : wizardReport({ query, columns: this.runQuery({ name: query, limit: 1 }).columns, name: n });
    this.saveReport(layout);
    return layout;
  }

  /**
   * A report's data, ready to draw: its records grouped with their totals,
   * a field that points at another table's record shown as that record's
   * words. `layout` is a report's, or `name` names a kept one. `fields` are
   * the source's fields with their types, for grouping, sorting and showing.
   */
  reportData({ name = null, layout = null, today = this.today } = {}) {
    const l = layout || this.reports().find((r) => r.name === name)?.layout;
    if (!l) throw new Error(`There is no report called ${name}.`);
    let records;
    let fields;
    const shown = {};
    if (l.source?.table) {
      fields = this.fields(l.source.table).map((f) => ({ name: f.name, type: f.type }));
      records = this.rows(l.source.table, { limit: 1_000_000 });
      for (const r of this.relationships().filter((x) => x.from.table === l.source.table)) {
        const show = shownField(this.fields(r.to.table));
        shown[r.from.field] = new Map(this.lookup({ table: r.to.table, field: r.to.field, show, limit: 1_000_000 }).map((o) => [o.value, o.text]));
      }
    } else {
      const out = this.runQuery({ name: l.source?.query, limit: 1_000_000, today });
      records = out.rows;
      fields = out.columns.map((c) => ({ name: c, type: null }));
    }
    const shownAs = (field, value) => (shown[field]?.has(value) ? shown[field].get(value) : value);
    // A pointing field is shown as words, not as the number it holds.
    for (const f of fields) if (shown[f.name]) f.type = 'text';
    return { layout: l, fields, ...reportData(l, records, shownAs) };
  }

  /** The print templates (an invoice, a receipt, a payslip): [{ name, layout }], in their order. */
  documents() {
    if (!this.hasDesignTable('rutba_documents') || this.access?.locked) return [];
    const all = this.db.prepare('SELECT name, layout FROM rutba_documents ORDER BY position, name').all().map((r) => ({ name: r.name, layout: JSON.parse(r.layout) }));
    return this._shown('documents', all, (x) => [x.layout.table, x.layout.letterhead?.table, ...(x.layout.blocks || []).map((bl) => bl.table)].filter(Boolean));
  }

  /** A print template kept (made, or replaced), its table, paper and blocks checked. */
  saveDocument(layout) {
    const n = checkName(layout?.name, 'print template');
    if (!this.tables().some((t) => t.name === layout.table)) throw new Error(`The template's table, ${layout.table}, is not in this database.`);
    if (layout.paper && !PAPERS.includes(layout.paper)) throw new Error(`A template prints on ${PAPERS.join(', ')}.`);
    for (const b of layout.blocks || []) if (!BLOCKS.includes(b.type)) throw new Error(`A template has no block called ${b.type}.`);
    if (layout.letterhead?.table && !this.tables().some((t) => t.name === layout.letterhead.table)) throw new Error(`The letterhead's table, ${layout.letterhead.table}, is not in this database.`);
    const at = this.db.prepare('SELECT COALESCE(MAX(position), -1) + 1 AS p FROM rutba_documents').get().p;
    return this._step(`Print template ${n} kept`, { document: n }, () => {
      this.db.prepare('INSERT INTO rutba_documents (name, layout, position) VALUES (?, ?, ?) ON CONFLICT(name) DO UPDATE SET layout = excluded.layout').run(n, JSON.stringify({ ...layout, kind: 'document', name: n }), at);
      return n;
    });
  }

  removeDocument(name) {
    return this._step(`Print template ${name} removed`, { document: name }, () => this.db.prepare('DELETE FROM rutba_documents WHERE name = ?').run(name).changes > 0);
  }

  /** A print template for `table` as the wizard makes it, kept under a name of its own. */
  makeDocument({ table, name = null } = {}) {
    const taken = new Set(this.documents().map((d) => d.name.toLowerCase()));
    const want = name || table.replace(/s$/, '');
    let n = want;
    for (let i = 2; taken.has(n.toLowerCase()); i++) n = `${want} ${i}`;
    const layout = wizardDocument({ table, fieldsOf: (t) => this.fields(t), relationships: this.relationships(), tables: this.tables().map((t) => t.name), name: n });
    this.saveDocument(layout);
    return layout;
  }

  /**
   * A print template's documents, ready to draw: one for each record of its
   * table, or each of `keys` (the one a person is looking at, or those chosen).
   */
  documentData({ name = null, layout = null, keys = null, limit = 1000 } = {}) {
    const l = layout || this.documents().find((d) => d.name === name)?.layout;
    if (!l) throw new Error(`There is no print template called ${name}.`);
    return documentData(this, l, { keys, limit });
  }

  /** The choices for a lookup: each record of `table` by its `field`, shown as `show` (a field, or several), at most `limit`. */
  lookup({ table, field, show, limit = 2000 }) {
    // A role sees the choices of a table it may read, by the fields it may see.
    let wanted = Array.isArray(show) && show.length ? show : [show || field];
    if (this._rights()) {
      const seen = new Set(this.fields(table).map((f) => f.name));
      if (!seen.has(field)) return [];
      wanted = wanted.filter((f) => seen.has(f));
      if (!wanted.length) wanted = [field];
    }
    // Two fields or more (a first and a last name) are shown together, a space between, and listed by the last.
    const shown = wanted;
    const text = shown.length > 1 ? `TRIM(${shown.map((f) => `COALESCE(${q(f)}, '')`).join(" || ' ' || ")})` : q(shown[0]);
    const order = shown.length > 1 ? `${q(shown[shown.length - 1])}, ${text}` : text;
    const rows = this.db.prepare(`SELECT ${q(field)} AS value, ${text} AS text FROM ${q(table)} ORDER BY ${order} LIMIT ?`).all(limit);
    return rows.map((r) => ({ value: r.value, text: r.text == null ? String(r.value) : String(r.text) }));
  }

  /**
   * A query of the person's own, in SQL: SELECT (or WITH ... SELECT) only,
   * one statement, run with the database read-only. Answers { columns, rows }.
   */
  query(sql, params = []) {
    // A person's own SQL could name main.<table> and read past their role: it is for those who may change the design.
    this._mayDesign();
    return this._select(sql, params);
  }

  _select(sql, params = []) {
    const text = String(sql).trim().replace(/;\s*$/, '');
    if (!/^(select|with)\b/i.test(text)) throw new Error('A query here starts with SELECT.');
    if (/;/.test(text.replace(/'(?:[^']|'')*'|"(?:[^"]|"")*"/g, ''))) throw new Error('One statement at a time.');
    this.db.exec('PRAGMA query_only = ON');
    try {
      const stmt = this.db.prepare(text);
      const rows = stmt.all(...params).map((r) => ({ ...r }));
      const columns = stmt.columns().map((c) => c.name);
      return { columns, rows };
    } finally {
      this.db.exec('PRAGMA query_only = OFF');
    }
  }

  /* ── people and roles (D-2026-10-22-11) ─────────────────────────── */

  /** Whether the database asks who you are as it opens. */
  isProtected() {
    return this.meta('protected') === '1';
  }

  /** The rights that narrow what this person sees, or null when nothing does. */
  _rights() {
    const a = this.access;
    return a && !a.locked && !a.full ? a.rights : null;
  }

  _unlocked() {
    if (this.access?.locked) throw new Error('Sign in to open this database.');
  }

  _mayDesign() {
    this._unlocked();
    if (this.access?.seeing) throw new Error(`You are seeing the database as ${this.access.seeing} sees it. Go back to your own view to change it.`);
    if (this._rights()) throw new Error('Only someone whose role may change the design can do that.');
  }

  /** May this person `want` ('add', 'change', 'delete') in `table`, on these rows? */
  _mayWrite(table, want, rowids = []) {
    this._unlocked();
    if (this.access?.seeing) throw new Error(`You are seeing the database as ${this.access.seeing} sees it. Go back to your own view to change it.`);
    const rights = this._rights();
    if (!rights) return;
    const words = { add: 'add records to', change: 'change records in', delete: 'delete records from' };
    if (!allows(rights, table, want)) throw new Error(`Your role may not ${words[want]} ${table}.`);
    // A record the role cannot see (another person's, for one that sees its own only) is not theirs to change.
    for (const id of rowids) {
      if (!this.db.prepare(`SELECT 1 FROM ${q(table)} WHERE rowid = ?`).get(id)) throw new Error(`Your role may not ${words[want]} that record of ${table}.`);
    }
  }

  /** A value given for a field the role hides is refused, not quietly dropped. */
  _noHiddenValues(table, values) {
    const rights = this._rights();
    if (!rights) return;
    const right = tableRight(rights, table);
    for (const f of Object.keys(values)) if (right.fields[f] === 'hidden') throw new Error(`There is no field called ${f} in ${table}.`);
  }

  /** Where a write goes: past a role's view, to the table itself. */
  _target(table) {
    return this._rights() ? `main.${q(table)}` : q(table);
  }

  /** Who added, changed or deleted which records, kept while the database is protected. */
  _trail(table, rowids, action, detail = null) {
    if (!this.access || this.access.locked) return;
    const put = this.db.prepare('INSERT INTO rutba_trail (at, person, tbl, row, action, detail) VALUES (?, ?, ?, ?, ?, ?)');
    const at = new Date().toISOString();
    for (const id of rowids) put.run(at, this.access.person?.id ?? null, table, id, action, detail == null ? null : JSON.stringify(detail));
  }

  /** Saved items the role shows: not hidden by name, and none over a table it hides. */
  _shown(kind, all, tablesOf, queryOf = () => null) {
    const rights = this._rights();
    if (!rights) return all;
    const hidden = new Set(rights.hidden[kind]);
    const hiddenQueries = new Set(rights.hidden.queries);
    return all.filter((x) => !hidden.has(x.name)
      && tablesOf(x).every((t) => tableRight(rights, t).access !== 'hidden')
      && !(queryOf(x) && hiddenQueries.has(queryOf(x))));
  }

  /** The views a restricted person reads through, made again (after sign-in, a role changed, See as). */
  _applyViews() {
    for (const v of this.db.prepare("SELECT name FROM sqlite_temp_master WHERE type = 'view'").all()) this.db.exec(`DROP VIEW temp.${q(v.name)}`);
    const rights = this._rights();
    if (!rights) return;
    for (const t of this._allTables()) this.db.exec(viewSql(t.name, this._allFields(t.name), rights, this.access.person?.id));
  }

  /** The person signed in: { id, name, role, admin, seeing }, or null. */
  person() {
    const a = this.access;
    if (!a || a.locked) return null;
    return { ...a.person, admin: Boolean(a.full), seeing: a.seeing || null };
  }

  /** The rights in force, for the app to show what is read only: null when nothing is narrowed. */
  rights() {
    return this._rights();
  }

  /** The names that may sign in, for the sign-in box: nothing else about them. */
  signInNames() {
    if (!this.isProtected()) return [];
    return this.db.prepare('SELECT name FROM rutba_people ORDER BY name').all().map((r) => r.name);
  }

  /**
   * Protect the database: from now on it asks who you are. The person who
   * protects it is its first Admin; two roles to start from are made beside it.
   */
  protect({ name, password }) {
    this._mayDesign();
    if (this.isProtected()) throw new Error('This database is protected already.');
    const n = checkName(name, 'person');
    const hash = hashPassword(password);
    const at = new Date().toISOString();
    this._step('Protect the database', { admin: n }, () => {
      const putRole = this.db.prepare('INSERT INTO rutba_roles (name, rights, position) VALUES (?, ?, ?) ON CONFLICT(name) DO NOTHING');
      putRole.run(ADMIN, JSON.stringify({ design: true }), 0);
      STARTING_ROLES.forEach((r, i) => putRole.run(r.name, JSON.stringify(normalRights(r.rights)), i + 1));
      this.db.prepare('DELETE FROM rutba_people').run();
      this.db.prepare('INSERT INTO rutba_people (name, role, hash, created) VALUES (?, ?, ?, ?)').run(n, ADMIN, hash, at);
      this.setMeta('protected', '1');
    });
    this.access = { locked: true };
    return this.signIn(n, password);
  }

  /** Take the protection away: an Admin, with their password. The people are removed; the trail is kept. */
  unprotect({ password }) {
    const me = this.person();
    if (!me?.admin || me.seeing) throw new Error('Only an Admin can take the protection away.');
    const row = this.db.prepare('SELECT hash FROM rutba_people WHERE id = ?').get(me.id);
    if (!row || !checkPassword(password, row.hash)) throw new Error('That password is not right.');
    this._step('Take the protection away', null, () => {
      this.db.prepare('DELETE FROM rutba_people').run();
      this.setMeta('protected', '0');
    });
    this.access = null;
    this._applyViews();
  }

  /** Sign in: the database opens as the person's role allows. */
  signIn(name, password) {
    if (!this.isProtected()) throw new Error('This database is not protected.');
    const p = this.db.prepare('SELECT id, name, role, hash FROM rutba_people WHERE name = ?').get(String(name ?? '').trim());
    if (!p || !checkPassword(password, p.hash)) throw new Error('That name and password do not match.');
    const role = this.db.prepare('SELECT name, rights FROM rutba_roles WHERE name = ?').get(p.role);
    const rights = normalRights(role ? JSON.parse(role.rights) : {});
    this.access = { person: { id: p.id, name: p.name, role: p.role }, rights, full: p.role === ADMIN || rights.design, seeing: null };
    this._applyViews();
    return this.person();
  }

  /** Sign out: the database is locked again until someone signs in. */
  signOut() {
    if (!this.isProtected()) return;
    this.access = { locked: true };
    this._applyViews();
  }

  /** An Admin sees the database as `role` would (null to go back): reading only. */
  seeAs(role) {
    const a = this.access;
    if (!a || a.locked || !a.person || a.person.role !== ADMIN) throw new Error('Only an Admin can see the database as another role does.');
    if (role == null || role === ADMIN) {
      this.access = { person: a.person, rights: normalRights({ design: true }), full: true, seeing: null };
    } else {
      const r = this.db.prepare('SELECT name, rights FROM rutba_roles WHERE name = ?').get(role);
      if (!r) throw new Error(`There is no role called ${role}.`);
      const rights = normalRights(JSON.parse(r.rights));
      this.access = { person: a.person, rights, full: rights.design, seeing: r.name };
    }
    this._applyViews();
    return this.person();
  }

  _admin() {
    const me = this.person();
    if (!me || me.role !== ADMIN || me.seeing) throw new Error('Only an Admin can change people and roles.');
    return me;
  }

  /** The roles: [{ name, rights, builtIn }], Admin first. */
  roles() {
    this._admin();
    return this.db.prepare('SELECT name, rights FROM rutba_roles ORDER BY position, name').all().map((r) => ({ name: r.name, rights: normalRights(JSON.parse(r.rights)), builtIn: r.name === ADMIN }));
  }

  /** A role kept (made, or its rights replaced). Admin's are its own. */
  saveRole(name, rights) {
    this._admin();
    const n = checkName(name, 'role');
    if (n.toLowerCase() === ADMIN.toLowerCase()) throw new Error('Admin may do everything; its rights are not changed.');
    const at = this.db.prepare('SELECT COALESCE(MAX(position), -1) + 1 AS p FROM rutba_roles').get().p;
    this._step(`Role ${n} kept`, { role: n }, () => {
      this.db.prepare('INSERT INTO rutba_roles (name, rights, position) VALUES (?, ?, ?) ON CONFLICT(name) DO UPDATE SET rights = excluded.rights')
        .run(n, JSON.stringify(normalRights(rights)), at);
    });
    return n;
  }

  /** A role taken away; one that people have stays until they are given another. */
  removeRole(name) {
    this._admin();
    if (name === ADMIN) throw new Error('Admin stays.');
    const n = this.db.prepare('SELECT count(*) AS n FROM rutba_people WHERE role = ?').get(name).n;
    if (n) throw new Error(`${n} ${n === 1 ? 'person has' : 'people have'} the role ${name}. Give them another first.`);
    this._step(`Role ${name} removed`, { role: name }, () => this.db.prepare('DELETE FROM rutba_roles WHERE name = ?').run(name));
  }

  /** The people: [{ id, name, role, created }]. */
  people() {
    this._admin();
    return this.db.prepare('SELECT id, name, role, created FROM rutba_people ORDER BY name').all().map((r) => ({ ...r }));
  }

  _roleThere(role) {
    if (!this.db.prepare('SELECT 1 FROM rutba_roles WHERE name = ?').get(role)) throw new Error(`There is no role called ${role}.`);
  }

  /** Someone added, with a role and a password of their own. */
  addPerson({ name, role, password }) {
    this._admin();
    const n = checkName(name, 'person');
    this._roleThere(role);
    if (this.db.prepare('SELECT 1 FROM rutba_people WHERE name = ?').get(n)) throw new Error(`There is someone called ${n} already.`);
    const hash = hashPassword(password);
    return this._step(`Person ${n} added`, { person: n, role }, () => Number(this.db.prepare('INSERT INTO rutba_people (name, role, hash, created) VALUES (?, ?, ?, ?)')
      .run(n, role, hash, new Date().toISOString()).lastInsertRowid));
  }

  /** A person's name or role changed. The last Admin stays an Admin. */
  changePerson(id, { name, role }) {
    const me = this._admin();
    const p = this.db.prepare('SELECT id, name, role FROM rutba_people WHERE id = ?').get(id);
    if (!p) throw new Error('There is no such person.');
    const n = name == null ? p.name : checkName(name, 'person');
    const r = role == null ? p.role : role;
    this._roleThere(r);
    if (p.role === ADMIN && r !== ADMIN && this._admins() === 1) throw new Error('The database keeps at least one Admin.');
    this._step(`Person ${n} changed`, { person: n, role: r }, () => this.db.prepare('UPDATE rutba_people SET name = ?, role = ? WHERE id = ?').run(n, r, id));
    if (id === me.id) this.access.person = { ...this.access.person, name: n };
  }

  _admins() {
    return this.db.prepare('SELECT count(*) AS n FROM rutba_people WHERE role = ?').get(ADMIN).n;
  }

  /** A person taken away. The last Admin, and yourself, stay. */
  removePerson(id) {
    const me = this._admin();
    if (id === me.id) throw new Error('You cannot remove yourself.');
    const p = this.db.prepare('SELECT name, role FROM rutba_people WHERE id = ?').get(id);
    if (!p) return;
    if (p.role === ADMIN && this._admins() === 1) throw new Error('The database keeps at least one Admin.');
    this._step(`Person ${p.name} removed`, { person: p.name }, () => this.db.prepare('DELETE FROM rutba_people WHERE id = ?').run(id));
  }

  /** A new password: your own, with the one you have now; anyone's, for an Admin. */
  setPassword(id, password, { current = null } = {}) {
    const me = this.person();
    if (!me) throw new Error('Sign in to open this database.');
    const self = me.id === id && !me.seeing;
    if (!self) this._admin();
    if (self) {
      const row = this.db.prepare('SELECT hash FROM rutba_people WHERE id = ?').get(id);
      if (!row || !checkPassword(current, row.hash)) throw new Error('Your password now is not right.');
    }
    const hash = hashPassword(password);
    // Not a design step: anyone may change their own, and a password is not part of the design's history.
    this.db.prepare('UPDATE rutba_people SET hash = ? WHERE id = ?').run(hash, id);
  }

  /** Who added and changed a record: [{ at, person, action, fields }], for a record the person can see. */
  trailOf(table, rowid) {
    this._known(table);
    if (!this.db.prepare(`SELECT 1 FROM ${q(table)} WHERE rowid = ?`).get(rowid)) return [];
    return this.db.prepare('SELECT t.at, p.name AS person, t.action, t.detail FROM rutba_trail t LEFT JOIN rutba_people p ON p.id = t.person WHERE t.tbl = ? AND t.row = ? ORDER BY t.n')
      .all(table, rowid).map((r) => ({ at: r.at, person: r.person ?? null, action: r.action, fields: r.detail ? JSON.parse(r.detail).fields || null : null }));
  }

}
