#!/usr/bin/env node
// Whether a file follows the Data Files format (docs/DATA-FILES-FORMAT.md):
//
//   node tools/check-rbase.mjs <file.rbase>
//
// The file is read and never changed: it is opened read-only (immutable when
// nothing else has it open, so no -wal or -shm file is made beside it). It
// stands alone, with node:sqlite and nothing else, and checks the file
// against the text rather than against Rutba Office's own code, so that
// what it accepts is what the specification says. Exit 0 when the file
// follows it; 1 with a list of the problems when it does not; 2 when it is
// used wrongly.

import { DatabaseSync } from 'node:sqlite';
import { pathToFileURL, fileURLToPath } from 'node:url';
import fs from 'node:fs';
import path from 'node:path';

/** The design version this checker knows: the specification's version. */
export const SPEC_VERSION = 6;

/** The field types of section 2.2. */
export const FIELD_TYPES = ['autonumber', 'text', 'longtext', 'number', 'integer', 'currency', 'date', 'datetime', 'yesno'];

/**
 * The design tables of section 3: the version that added each, whether it is
 * made only when first needed, and its columns (those a file must have; the
 * `later` ones were added within the version and may be absent).
 */
export const DESIGN_TABLES = {
  rutba_meta: { since: 1, columns: ['key', 'value'] },
  rutba_tables: { since: 1, columns: ['name', 'caption', 'position'] },
  rutba_fields: { since: 1, columns: ['tbl', 'name', 'type', 'size', 'required', 'uniq', 'dflt', 'caption', 'position'] },
  rutba_relations: { since: 1, columns: ['id', 'from_tbl', 'from_field', 'to_tbl', 'to_field', 'cascade_update', 'cascade_delete'] },
  rutba_steps: { since: 1, columns: ['n', 'at', 'label', 'detail'] },
  rutba_queries: { since: 2, columns: ['name', 'design', 'position'] },
  rutba_forms: { since: 2, columns: ['name', 'layout', 'position'] },
  rutba_reports: { since: 3, columns: ['name', 'layout', 'position'] },
  rutba_documents: { since: 4, columns: ['name', 'layout', 'position'] },
  rutba_roles: { since: 5, columns: ['name', 'rights', 'position'] },
  rutba_people: { since: 5, columns: ['id', 'name', 'role', 'hash', 'account', 'created'] },
  rutba_trail: { since: 5, columns: ['n', 'at', 'person', 'tbl', 'row', 'action', 'detail'] },
  rutba_links: { since: 6, columns: ['id', 'name', 'engine', 'settings', 'position'] },
  rutba_linked: { since: 6, columns: ['tbl', 'link', 'remote', 'key', 'caption', 'position'], later: ['mode', 'stamp', 'synced'] },
  rutba_sync: { since: 6, whenNeeded: true, columns: ['tbl', 'rid', 'kj', 'hash', 'state', 'n', 'since'] },
  rutba_clashes: { since: 6, whenNeeded: true, columns: ['id', 'tbl', 'kj', 'at', 'action', 'mine', 'theirs', 'why'] },
};

const ENGINES = ['postgres', 'mysql'];
const SOURCES = ['file', 'rbase', 'json'];
const CONTROL_KINDS = ['text', 'longtext', 'number', 'currency', 'date', 'datetime', 'check', 'list', 'auto'];
const QUERY_TOTALS = ['group', 'sum', 'avg', 'count', 'min', 'max', 'where'];
const REPORT_FNS = ['sum', 'avg', 'count', 'min', 'max'];
const PAPERS = ['a4', 'letter', 'a5', 'receipt-80', 'receipt-58'];
const BLOCKS = ['heading', 'fields', 'lines', 'totals', 'text', 'signatures', 'spacer'];
const LEVELS = ['hidden', 'read', 'add', 'change', 'delete'];
const PIN_KINDS = ['table', 'query', 'form', 'report', 'document'];
/** A column's declared type for each field type: own tables (section 2.2) and copies (section 9.2). */
const AFFINITY = { autonumber: 'INTEGER', integer: 'INTEGER', yesno: 'INTEGER', number: 'REAL', currency: 'REAL', text: 'TEXT', longtext: 'TEXT', date: 'TEXT', datetime: 'TEXT' };

const q = (name) => `"${String(name).replace(/"/g, '""')}"`;
const isObject = (v) => v !== null && typeof v === 'object' && !Array.isArray(v);
const isText = (v) => typeof v === 'string';
const isWhole = (v) => typeof v === 'number' && Number.isInteger(v);
const isNumber = (v) => typeof v === 'number' && Number.isFinite(v);
const lower = (s) => String(s).toLowerCase();

/** Section 2.1: a name a person may give. Null when it is one; else what is wrong with it. */
function nameFault(name) {
  if (!isText(name)) return 'is not text';
  if (name !== name.trim()) return 'has spaces at an end';
  if (!name.length) return 'is empty';
  if (name.length > 64) return 'is longer than 64 characters';
  if (/^(rutba_|sqlite_)/i.test(name)) return 'starts with a name the file keeps for itself';
  if (/["\u0000-\u001f]/.test(name)) return 'holds a quotation mark or a control character';
  return null;
}

/** The file opened read-only: as it is, when another program has it open (a -wal beside it); else immutable, touching nothing. */
function openReadOnly(file) {
  const abs = path.resolve(file);
  if (fs.existsSync(`${abs}-wal`)) return new DatabaseSync(abs, { readOnly: true });
  return new DatabaseSync(`${pathToFileURL(abs).href}?immutable=1`, { readOnly: true });
}

/**
 * Check `file` against the specification. Answers { ok, version, problems,
 * notes }: `problems` are where it departs from the text, `notes` what a
 * reader would ignore and keep (unknown tables, columns and keys).
 */
export function checkRbase(file) {
  const problems = [];
  const notes = [];
  const problem = (s) => problems.push(s);
  const note = (s) => notes.push(s);
  const out = (version = null) => ({ ok: problems.length === 0, version, problems, notes });

  if (!fs.existsSync(file) || !fs.statSync(file).isFile()) { problem(`${file} is not there.`); return out(); }
  const head = Buffer.alloc(16);
  const fd = fs.openSync(file, 'r');
  try { fs.readSync(fd, head, 0, 16, 0); } finally { fs.closeSync(fd); }
  if (head.toString('latin1') !== 'SQLite format 3\u0000') { problem('It is not an SQLite database: its first 16 bytes are not "SQLite format 3" and a zero byte.'); return out(); }

  let db;
  try {
    db = openReadOnly(file);
  } catch (err) {
    problem(`SQLite could not open it: ${err.message}`);
    return out();
  }
  try {
    return check(db, problem, note, out);
  } catch (err) {
    problem(`It could not be read to the end: ${err.message}`);
    return out();
  } finally {
    db.close();
  }
}

function check(db, problem, note, out) {
  const all = (sql, ...p) => db.prepare(sql).all(...p).map((r) => ({ ...r }));
  const sound = all('PRAGMA quick_check');
  if (!(sound.length === 1 && Object.values(sound[0])[0] === 'ok')) problem(`SQLite finds the file damaged: ${sound.map((r) => Object.values(r)[0]).slice(0, 5).join('; ')}`);

  const objects = all("SELECT type, name, tbl_name, sql FROM sqlite_master");
  const tableNames = new Set(objects.filter((o) => o.type === 'table').map((o) => o.name));
  const hasTable = (name) => tableNames.has(name);
  const columnsOf = (table) => all(`PRAGMA table_info(${q(table)})`);
  for (const o of objects) {
    if (o.type === 'view' || o.type === 'trigger') problem(`The file holds a ${o.type}, ${o.name}: a Data Files database holds tables and indexes only.`);
  }

  // A Data Files database at all (section 1).
  const missing = ['rutba_meta', 'rutba_tables', 'rutba_fields'].filter((t) => !hasTable(t));
  for (const t of missing) problem(`There is no ${t} table, so it is not a Data Files database.`);
  if (missing.length) return out();

  const metaRows = all('SELECT key, value FROM rutba_meta');
  const meta = new Map(metaRows.map((r) => [r.key, r.value]));
  const versionText = meta.get('version');
  let version = null;
  if (versionText == null) problem('rutba_meta has no version.');
  else if (!/^[1-9][0-9]*$/.test(String(versionText))) problem(`rutba_meta's version, "${versionText}", is not a whole number.`);
  else {
    version = Number(versionText);
    if (version > SPEC_VERSION) problem(`Its design version, ${version}, is newer than this specification (${SPEC_VERSION}): a reader must not change it.`);
  }
  const v = version ?? SPEC_VERSION;

  // The design tables and their columns (section 3).
  const cols = {};
  for (const [name, t] of Object.entries(DESIGN_TABLES)) {
    if (!hasTable(name)) {
      if (t.since <= v && !t.whenNeeded) problem(`${name} is missing: a file of design version ${v} has it.`);
      continue;
    }
    const have = columnsOf(name).map((c) => c.name);
    cols[name] = new Set(have);
    for (const c of t.columns) if (!cols[name].has(c)) problem(`${name} has no column ${c}.`);
    for (const c of have) if (!t.columns.includes(c) && !(t.later || []).includes(c)) note(`${name} has a column this specification does not name, ${c}: a reader keeps it.`);
  }
  for (const name of tableNames) {
    if (/^rutba_/i.test(name) && !DESIGN_TABLES[name]) note(`${name} is a design table this specification does not name: a reader ignores it and keeps it.`);
  }
  if (hasTable('rutba_sync') && cols.rutba_linked) {
    for (const c of DESIGN_TABLES.rutba_linked.later) if (!cols.rutba_linked.has(c)) problem(`rutba_linked has no column ${c}, which a file with rutba_sync has.`);
  }
  const read = (table, order = '') => (hasTable(table) ? all(`SELECT * FROM ${table}${order}`) : []);

  // JSON, parsed where a column holds it.
  const json = (text, where, { object = true, nullable = false } = {}) => {
    if (text === null || text === undefined) {
      if (!nullable) problem(`${where} is empty.`);
      return undefined;
    }
    if (!isText(text)) { problem(`${where} is not text.`); return undefined; }
    let value;
    try { value = JSON.parse(text); } catch (err) { problem(`${where} is not JSON: ${err.message}`); return undefined; }
    if (object && !isObject(value)) { problem(`${where} is not a JSON object.`); return undefined; }
    return value;
  };
  const bool01 = (value, where) => { if (value !== 0 && value !== 1) problem(`${where} is ${JSON.stringify(value)}, not 0 or 1.`); };
  const time = (value, where, nullable = false) => {
    if (value === null && nullable) return;
    if (!isText(value) || Number.isNaN(Date.parse(value))) problem(`${where}, ${JSON.stringify(value)}, is not a time.`);
  };

  // rutba_meta's keys.
  if (meta.has('protected') && !['0', '1'].includes(meta.get('protected'))) problem(`rutba_meta's protected is "${meta.get('protected')}", not 0 or 1.`);
  const known = new Set(['version', 'protected', 'shown', 'pinned', 'folded', 'origin', 'origin.design']);
  for (const k of meta.keys()) if (!known.has(k)) note(`rutba_meta has a key this specification does not name, ${k}: a reader keeps it.`);

  // Tables and fields (sections 2 and 3).
  const tables = read('rutba_tables');
  const linked = read('rutba_linked');
  const own = new Map();
  for (const t of tables) {
    const fault = nameFault(t.name);
    if (fault) problem(`The table name ${JSON.stringify(t.name)} ${fault}.`);
    if (!hasTable(t.name)) problem(`rutba_tables names ${t.name}, which is not a table of the file.`);
    if (t.position !== null && !isWhole(t.position)) problem(`rutba_tables: ${t.name}'s position is not a whole number.`);
    if ([...own.keys()].some((n) => lower(n) === lower(t.name))) problem(`Two tables are called ${t.name}, without regard to case.`);
    own.set(t.name, true);
  }
  const linkedNames = new Map(linked.map((l) => [l.tbl, l]));
  for (const name of tableNames) {
    if (!/^(rutba_|sqlite_)/i.test(name) && !own.has(name) && !linkedNames.has(name)) note(`The table ${name} is not in rutba_tables or rutba_linked, so it is not part of the design.`);
  }
  for (const l of linked) if ([...own.keys()].some((n) => lower(n) === lower(l.tbl))) problem(`${l.tbl} is both a table of the file and a linked table.`);
  const isTable = (table) => own.has(table) || linkedNames.has(table);

  const fieldRows = read('rutba_fields', ' ORDER BY tbl, position');
  const fields = new Map();
  for (const f of fieldRows) {
    const at = `Field ${f.tbl}.${f.name}`;
    if (!isTable(f.tbl)) { problem(`${at} belongs to ${f.tbl}, which is neither in rutba_tables nor in rutba_linked.`); continue; }
    const fault = nameFault(f.name);
    if (fault) problem(`${at}: its name ${fault}.`);
    if (!FIELD_TYPES.includes(f.type)) problem(`${at} has the type ${JSON.stringify(f.type)}, which is not one of ${FIELD_TYPES.join(', ')}.`);
    if (f.size !== null && !(isWhole(f.size) && f.size > 0)) problem(`${at}'s size is ${JSON.stringify(f.size)}, not a whole number above 0.`);
    bool01(f.required, `${at}'s required`);
    bool01(f.uniq, `${at}'s uniq`);
    if (f.dflt !== null) json(f.dflt, `${at}'s default`, { object: false });
    if (!isWhole(f.position)) problem(`${at}'s position is not a whole number.`);
    if (!fields.has(f.tbl)) fields.set(f.tbl, []);
    fields.get(f.tbl).push(f);
  }
  const fieldOf = (table, name) => (fields.get(table) || []).find((f) => f.name === name) || null;
  const fieldNames = (table) => new Set((fields.get(table) || []).map((f) => f.name));

  for (const [table, list] of fields) {
    const names = list.map((f) => lower(f.name));
    if (new Set(names).size !== names.length) problem(`Two fields of ${table} have the same name, without regard to case.`);
    if (list.filter((f) => f.type === 'autonumber').length > 1) problem(`${table} has more than one AutoNumber.`);
  }

  // Each own table's columns, as its fields say (section 2.2); a copy's as section 9.2 says.
  const copyOf = new Map(linked.filter((l) => (l.mode ?? 'live') === 'copy').map((l) => [l.tbl, l]));
  for (const table of [...own.keys(), ...copyOf.keys()]) {
    if (!hasTable(table)) { if (copyOf.has(table)) problem(`${table} is kept as a copy, but the file has no table ${table}.`); continue; }
    const columns = columnsOf(table);
    const described = fields.get(table) || [];
    if (!described.length) problem(`${table} has no fields in rutba_fields.`);
    for (const c of columns) if (!described.some((f) => f.name === c.name)) problem(`Column ${c.name} of ${table} is not in rutba_fields.`);
    for (const f of described) {
      const c = columns.find((x) => x.name === f.name);
      if (!c) { problem(`Field ${table}.${f.name} has no column in the table.`); continue; }
      const want = AFFINITY[f.type];
      if (want && String(c.type).toUpperCase() !== want) problem(`Column ${table}.${f.name} is declared ${c.type || 'with no type'}, not ${want} as a ${f.type} field is.`);
      if (f.type === 'autonumber' && !copyOf.has(table) && c.pk !== 1) problem(`The AutoNumber ${table}.${f.name} is not the table's INTEGER PRIMARY KEY.`);
    }
  }

  // Relationships (section 5).
  for (const r of read('rutba_relations', ' ORDER BY id')) {
    const at = `Relationship ${r.id} (${r.from_tbl}.${r.from_field} to ${r.to_tbl}.${r.to_field})`;
    const from = fieldOf(r.from_tbl, r.from_field);
    const to = fieldOf(r.to_tbl, r.to_field);
    if (!isTable(r.from_tbl) || !from) problem(`${at} points from a field that is not there.`);
    if (!isTable(r.to_tbl) || !to) problem(`${at} points at a field that is not there.`);
    bool01(r.cascade_update, `${at}'s cascade_update`);
    bool01(r.cascade_delete, `${at}'s cascade_delete`);
    if (!from || !to) continue;
    const withLinked = linkedNames.has(r.from_tbl) || linkedNames.has(r.to_tbl);
    if (withLinked) {
      if (r.cascade_update !== 0 || r.cascade_delete !== 0) problem(`${at} is with a linked table, so its cascade_update and cascade_delete are 0.`);
      continue;
    }
    if (!(to.type === 'autonumber' || to.uniq === 1)) problem(`${at} points at ${r.to_tbl}.${r.to_field}, which is neither its AutoNumber nor unique.`);
    if (!hasTable(r.from_tbl)) continue;
    const keys = all(`PRAGMA foreign_key_list(${q(r.from_tbl)})`);
    const fk = keys.find((k) => k.from === r.from_field && k.table === r.to_tbl && k.to === r.to_field);
    if (!fk) problem(`${at} has no FOREIGN KEY in ${r.from_tbl}.`);
    else {
      if (fk.on_update !== (r.cascade_update ? 'CASCADE' : 'RESTRICT')) problem(`${at}: its FOREIGN KEY says ON UPDATE ${fk.on_update}, the design says cascade_update ${r.cascade_update}.`);
      if (fk.on_delete !== (r.cascade_delete ? 'CASCADE' : 'RESTRICT')) problem(`${at}: its FOREIGN KEY says ON DELETE ${fk.on_delete}, the design says cascade_delete ${r.cascade_delete}.`);
    }
  }
  const broken = all('PRAGMA foreign_key_check');
  if (broken.length) problem(`${broken.length} record${broken.length === 1 ? '' : 's'} point at nothing (first in ${broken[0].table}).`);

  // The history (section 6).
  for (const s of read('rutba_steps', ' ORDER BY n')) {
    const at = `Step ${s.n}`;
    time(s.at, `${at}'s time`);
    if (!isText(s.label) || !s.label.trim()) problem(`${at} has no label.`);
    if (s.detail !== null) json(s.detail, `${at}'s detail`);
  }

  // Saved queries (section 4.1).
  const queries = new Map();
  for (const row of read('rutba_queries')) {
    const at = `Query ${row.name}`;
    const fault = nameFault(row.name);
    if (fault) problem(`${at}: its name ${fault}.`);
    if (isTable(row.name) || [...own.keys(), ...linkedNames.keys()].some((t) => lower(t) === lower(row.name))) problem(`${at} has the name of a table.`);
    const d = json(row.design, `${at}'s design`);
    if (!d) continue;
    queries.set(row.name, d);
    if ('sql' in d) {
      if (!isText(d.sql) || !/^\s*(select|with)\b/i.test(d.sql)) problem(`${at} is written in SQL, but its sql is not a SELECT.`);
      continue;
    }
    if (!Array.isArray(d.tables) || !d.tables.length || !d.tables.every(isText)) { problem(`${at}'s design has no list of tables.`); continue; }
    for (const t of d.tables) if (!isTable(t)) problem(`${at} reads ${t}, which is not a table of the database.`);
    if (!Array.isArray(d.columns)) { problem(`${at}'s design has no list of columns.`); continue; }
    d.columns.forEach((c, i) => {
      const where = `${at}, column ${i + 1}`;
      if (!isObject(c)) { problem(`${where} is not an object.`); return; }
      if (isText(c.expr)) {
        if (!c.expr.trim()) problem(`${where}'s expression is empty.`);
      } else if (isText(c.table) && isText(c.field)) {
        if (!d.tables.includes(c.table)) problem(`${where} is from ${c.table}, which is not one of the query's tables.`);
        else if (isTable(c.table) && !fieldNames(c.table).has(c.field)) problem(`${where}: ${c.table} has no field ${c.field}.`);
      } else problem(`${where} has neither a table and a field nor an expression.`);
      if (c.as != null && !isText(c.as)) problem(`${where}'s heading is not text.`);
      if (c.show !== undefined && typeof c.show !== 'boolean') problem(`${where}'s show is not true or false.`);
      if (c.sort != null && !['asc', 'desc'].includes(c.sort)) problem(`${where}'s sort is ${JSON.stringify(c.sort)}, not asc, desc or null.`);
      if (c.total != null && !QUERY_TOTALS.includes(c.total)) problem(`${where}'s total is ${JSON.stringify(c.total)}, not one of ${QUERY_TOTALS.join(', ')}.`);
      if (c.criteria !== undefined && !(Array.isArray(c.criteria) && c.criteria.every(isText))) problem(`${where}'s criteria are not a list of text.`);
      if (c.decimals != null && !(isWhole(c.decimals) && c.decimals >= 0 && c.decimals <= 6)) problem(`${where}'s decimals is ${JSON.stringify(c.decimals)}, not a whole number 0 to 6.`);
    });
  }

  const layoutRow = (row, column, what) => {
    const at = `${what} ${row.name}`;
    const fault = nameFault(row.name);
    if (fault) problem(`${at}: its name ${fault}.`);
    const l = json(row[column], `${at}'s ${column}`);
    if (l && l.name !== row.name) problem(`${at}'s layout is named ${JSON.stringify(l.name)}, not ${JSON.stringify(row.name)}.`);
    return l;
  };
  const box = (o, where) => { for (const k of ['x', 'y', 'w', 'h']) if (!isNumber(o[k])) problem(`${where} has no ${k}.`); };

  // Forms (section 4.2).
  for (const row of read('rutba_forms')) {
    const l = layoutRow(row, 'layout', 'Form');
    if (!l) continue;
    const at = `Form ${row.name}`;
    if (!isText(l.table) || !isTable(l.table)) { problem(`${at} shows ${JSON.stringify(l.table)}, which is not a table of the database.`); continue; }
    const have = fieldNames(l.table);
    if (!Array.isArray(l.controls)) problem(`${at} has no list of controls.`);
    else l.controls.forEach((c, i) => {
      const where = `${at}, control ${i + 1}`;
      if (!isObject(c)) { problem(`${where} is not an object.`); return; }
      if (!CONTROL_KINDS.includes(c.kind)) problem(`${where}'s kind is ${JSON.stringify(c.kind)}, not one of ${CONTROL_KINDS.join(', ')}.`);
      if (!isText(c.field) || !have.has(c.field)) problem(`${where} shows ${JSON.stringify(c.field)}, which is not a field of ${l.table}.`);
      box(c, where);
      if (c.lookup !== undefined) {
        if (!isObject(c.lookup) || !isTable(c.lookup.table) || !fieldNames(c.lookup.table).has(c.lookup.field)) problem(`${where}'s lookup does not name a table and a field that are there.`);
      }
    });
    if (l.subforms !== undefined) {
      if (!Array.isArray(l.subforms)) problem(`${at}'s subforms are not a list.`);
      else l.subforms.forEach((s, i) => {
        const where = `${at}, subform ${i + 1}`;
        if (!isObject(s)) { problem(`${where} is not an object.`); return; }
        if (!isTable(s.table)) { problem(`${where} shows ${JSON.stringify(s.table)}, which is not a table of the database.`); return; }
        if (!fieldNames(s.table).has(s.field)) problem(`${where}: ${s.table} has no field ${JSON.stringify(s.field)}.`);
        if (s.to !== undefined && !have.has(s.to)) problem(`${where} points at ${JSON.stringify(s.to)}, which is not a field of ${l.table}.`);
        if (!Array.isArray(s.columns) || !s.columns.every((f) => fieldNames(s.table).has(f))) problem(`${where}'s columns are not fields of ${s.table}.`);
        box(s, where);
      });
    }
  }

  // Reports (section 4.3).
  for (const row of read('rutba_reports')) {
    const l = layoutRow(row, 'layout', 'Report');
    if (!l) continue;
    const at = `Report ${row.name}`;
    const src = l.source;
    let have = null;
    if (!isObject(src) || (isText(src.table) === isText(src.query))) problem(`${at}'s source is not a table or a query.`);
    else if (isText(src.table)) {
      if (!isTable(src.table)) problem(`${at} reads ${src.table}, which is not a table of the database.`);
      else have = fieldNames(src.table);
    } else if (!queries.has(src.query)) problem(`${at} reads the query ${src.query}, which is not a saved query.`);
    const inSource = (f, where) => { if (have && !have.has(f)) problem(`${where}: ${src.table} has no field ${JSON.stringify(f)}.`); };
    if (!Array.isArray(l.columns)) problem(`${at} has no list of columns.`);
    else l.columns.forEach((c, i) => {
      const where = `${at}, column ${i + 1}`;
      if (!isObject(c) || !isText(c.field)) { problem(`${where} names no field.`); return; }
      inSource(c.field, where);
      if (c.align !== undefined && !['left', 'right'].includes(c.align)) problem(`${where}'s align is ${JSON.stringify(c.align)}, not left or right.`);
      if (c.type != null && !FIELD_TYPES.includes(c.type)) problem(`${where}'s type is ${JSON.stringify(c.type)}, not a field type.`);
    });
    if (l.group != null) { if (!isText(l.group)) problem(`${at}'s group is not a field.`); else inSource(l.group, `${at}'s group`); }
    if (l.sort != null) { if (!isText(l.sort)) problem(`${at}'s sort is not a field.`); else inSource(l.sort, `${at}'s sort`); }
    if (l.desc !== undefined && typeof l.desc !== 'boolean') problem(`${at}'s desc is not true or false.`);
    if (l.totals !== undefined) {
      if (!Array.isArray(l.totals)) problem(`${at}'s totals are not a list.`);
      else l.totals.forEach((t, i) => {
        if (!isObject(t) || !isText(t.field) || !REPORT_FNS.includes(t.fn)) problem(`${at}, total ${i + 1}, is not a field with one of ${REPORT_FNS.join(', ')}.`);
      });
    }
  }

  // Print templates (section 4.4).
  for (const row of read('rutba_documents')) {
    const l = layoutRow(row, 'layout', 'Print template');
    if (!l) continue;
    const at = `Print template ${row.name}`;
    if (l.kind !== 'document') problem(`${at}'s kind is ${JSON.stringify(l.kind)}, not "document".`);
    if (!isText(l.table) || !isTable(l.table)) problem(`${at} prints ${JSON.stringify(l.table)}, which is not a table of the database.`);
    if (l.paper !== undefined && !PAPERS.includes(l.paper)) problem(`${at}'s paper is ${JSON.stringify(l.paper)}, not one of ${PAPERS.join(', ')}.`);
    if (l.letterhead !== undefined && !(isObject(l.letterhead) && isTable(l.letterhead.table))) problem(`${at}'s letterhead does not name a table of the database.`);
    if (l.where !== undefined && !isText(l.where)) problem(`${at}'s where is not an expression.`);
    if (!Array.isArray(l.blocks)) { problem(`${at} has no list of blocks.`); continue; }
    l.blocks.forEach((b, i) => {
      const where = `${at}, block ${i + 1}`;
      if (!isObject(b)) { problem(`${where} is not an object.`); return; }
      if (!BLOCKS.includes(b.type)) { problem(`${where}'s type is ${JSON.stringify(b.type)}, not one of ${BLOCKS.join(', ')}.`); return; }
      if (b.if !== undefined && !isText(b.if)) problem(`${where}'s if is not an expression.`);
      if (b.type === 'fields' && !(Array.isArray(b.items) && b.items.every((it) => isObject(it) && isText(it.field)))) problem(`${where}'s items do not each name a field.`);
      if (b.type === 'lines') {
        if (!isTable(b.table)) problem(`${where} lists ${JSON.stringify(b.table)}, which is not a table of the database.`);
        else if (b.link !== undefined && !fieldNames(b.table).has(b.link)) problem(`${where} links by ${JSON.stringify(b.link)}, which is not a field of ${b.table}.`);
        if (!Array.isArray(b.columns) || !b.columns.every((c) => isObject(c) && (isText(c.field) || isText(c.expr) || isText(c.running)))) problem(`${where}'s columns do not each have a field, an expression or a running balance.`);
        if (b.on !== undefined && !isObject(b.on)) problem(`${where}'s on is not an object.`);
      }
      if (b.type === 'totals' && !(Array.isArray(b.rows) && b.rows.every((r) => isObject(r) && isText(r.label) && isText(r.expr)))) problem(`${where}'s rows do not each have a label and an expression.`);
      if ((b.type === 'heading' || b.type === 'text') && b.text !== undefined && !isText(b.text)) problem(`${where}'s text is not text.`);
      if (b.type === 'signatures' && b.labels !== undefined && !(Array.isArray(b.labels) && b.labels.every(isText))) problem(`${where}'s labels are not a list of text.`);
    });
  }

  // People, roles and the trail (section 8).
  const roles = new Map();
  for (const r of read('rutba_roles')) {
    const rights = json(r.rights, `Role ${r.name}'s rights`);
    if (!rights) continue;
    roles.set(lower(r.name), rights);
    if (rights.design !== undefined && typeof rights.design !== 'boolean') problem(`Role ${r.name}: design is not true or false.`);
    const t = rights.tables;
    if (t !== undefined) {
      if (!isObject(t)) problem(`Role ${r.name}: tables is not an object.`);
      else {
        if (t.default !== undefined && !LEVELS.includes(t.default)) problem(`Role ${r.name}: the default level ${JSON.stringify(t.default)} is not one of ${LEVELS.join(', ')}.`);
        for (const [name, e] of Object.entries(t.each || {})) {
          const access = isText(e) ? e : e?.access;
          if (!LEVELS.includes(access)) problem(`Role ${r.name}: ${name}'s level ${JSON.stringify(access)} is not one of ${LEVELS.join(', ')}.`);
          for (const [f, x] of Object.entries((isObject(e) && e.fields) || {})) if (!['hidden', 'read'].includes(x)) problem(`Role ${r.name}: ${name}.${f} is ${JSON.stringify(x)}, not hidden or read.`);
        }
      }
    }
  }
  const people = read('rutba_people');
  const isProtected = meta.get('protected') === '1';
  for (const p of people) {
    const at = `Person ${p.name}`;
    if (!roles.has(lower(p.role))) problem(`${at} has the role ${p.role}, which is not in rutba_roles.`);
    const parts = String(p.hash ?? '').split('$');
    const b64 = /^[A-Za-z0-9+/]+={0,2}$/;
    if (parts.length !== 6 || parts[0] !== 'scrypt' || ![1, 2, 3].every((i) => /^[1-9][0-9]*$/.test(parts[i])) || !b64.test(parts[4]) || !b64.test(parts[5])) {
      problem(`${at}'s password is not kept as scrypt$N$r$p$salt$key.`);
    } else if ((Number(parts[1]) & (Number(parts[1]) - 1)) !== 0 || Number(parts[1]) < 2) problem(`${at}'s scrypt N, ${parts[1]}, is not a power of two.`);
    time(p.created, `${at}'s created`);
  }
  if (isProtected) {
    if (!roles.has('admin')) problem('The database is protected, but rutba_roles has no Admin.');
    if (!people.some((p) => lower(p.role) === 'admin')) problem('The database is protected, but nobody has the role Admin.');
  } else if (people.length) problem(`rutba_people has ${people.length} ${people.length === 1 ? 'person' : 'people'}, but the database is not protected.`);
  for (const t of read('rutba_trail')) {
    if (!['add', 'change', 'delete'].includes(t.action)) problem(`Trail entry ${t.n}'s action is ${JSON.stringify(t.action)}, not add, change or delete.`);
    if (!isWhole(t.row)) problem(`Trail entry ${t.n}'s row is not a whole number.`);
    if (t.detail !== null) json(t.detail, `Trail entry ${t.n}'s detail`);
  }

  // Links, linked tables and copies (section 9).
  const links = new Map();
  for (const l of read('rutba_links')) {
    const at = `Link ${l.name}`;
    links.set(l.id, l);
    if (![...ENGINES, ...SOURCES].includes(l.engine)) problem(`${at} reaches ${JSON.stringify(l.engine)}, which is not one of ${[...ENGINES, ...SOURCES].join(', ')}.`);
    const s = json(l.settings, `${at}'s settings`);
    if (!s) continue;
    const secret = (o, where) => {
      for (const [k, x] of Object.entries(o)) {
        if (/^(password|passwd|secret|token|api[_-]?key)$/i.test(k)) problem(`${where} keeps a ${k}: secrets are not kept in the file.`);
        else if (isObject(x)) secret(x, where);
      }
    };
    secret(s, at);
    if (ENGINES.includes(l.engine)) {
      for (const k of Object.keys(s)) if (!['host', 'port', 'database', 'user', 'ssl'].includes(k)) note(`${at}'s settings have a key this specification does not name, ${k}.`);
    }
    if (l.engine === 'json' && s.headers !== undefined) {
      if (!Array.isArray(s.headers)) problem(`${at}'s headers are not a list.`);
      else for (const h of s.headers) if (isObject(h) && h.secret && h.value !== undefined) problem(`${at}: the secret header ${h.name} keeps its value in the file.`);
    }
    if (s.every !== undefined && ![5, 15, 60, 0].includes(s.every)) problem(`${at}'s every is ${JSON.stringify(s.every)}, not 5, 15, 60 or 0.`);
  }
  for (const l of linked) {
    const at = `Linked table ${l.tbl}`;
    const fault = nameFault(l.tbl);
    if (fault) problem(`${at}: its name ${fault}.`);
    const link = links.get(l.link);
    if (!link) problem(`${at} is linked through ${l.link}, which is not in rutba_links.`);
    json(l.remote, `${at}'s remote`);
    const mode = l.mode ?? 'live';
    if (!['live', 'copy'].includes(mode)) problem(`${at}'s mode is ${JSON.stringify(mode)}, not live or copy.`);
    if (link && SOURCES.includes(link.engine) && mode !== 'copy') problem(`${at} is read from a source, so it is kept as a copy.`);
    if (!(fields.get(l.tbl) || []).length) problem(`${at} has no fields in rutba_fields.`);
    if (l.key !== null && !fieldNames(l.tbl).has(l.key)) problem(`${at}'s key, ${l.key}, is not one of its fields.`);
    if (l.stamp != null && !fieldNames(l.tbl).has(l.stamp)) problem(`${at}'s stamp, ${l.stamp}, is not one of its fields.`);
    if (l.synced != null) time(l.synced, `${at}'s synced`, true);
    if (mode === 'copy' && !hasTable('rutba_sync')) problem(`${at} is kept as a copy, but the file has no rutba_sync.`);
  }
  for (const s of read('rutba_sync')) {
    const at = `rutba_sync entry ${s.tbl}/${s.rid}`;
    if (!copyOf.has(s.tbl)) problem(`${at} is for ${s.tbl}, which is not kept as a copy.`);
    if (!isWhole(s.rid)) problem(`${at}'s rid is not a whole number.`);
    if (s.state !== null && !['added', 'changed', 'deleted'].includes(s.state)) problem(`${at}'s state is ${JSON.stringify(s.state)}, not added, changed, deleted or empty.`);
    if (s.kj !== null) json(s.kj, `${at}'s key`, { object: false });
    if (s.hash !== null && !/^[0-9a-f]{40}$/.test(String(s.hash))) problem(`${at}'s hash is not a SHA-1 in hexadecimal.`);
  }
  for (const c of read('rutba_clashes')) {
    const at = `Clash ${c.id}`;
    if (!['changed', 'gone', 'deleted', 'refused'].includes(c.action)) problem(`${at}'s action is ${JSON.stringify(c.action)}, not changed, gone, deleted or refused.`);
    time(c.at, `${at}'s time`);
    if (c.mine !== null) json(c.mine, `${at}'s mine`);
    if (c.theirs !== null) json(c.theirs, `${at}'s theirs`);
  }

  // rutba_meta's JSON keys, now the tables are known.
  if (meta.has('shown')) {
    const shown = json(meta.get('shown'), "rutba_meta's shown");
    for (const [table, list] of Object.entries(shown || {})) {
      if (!isTable(table)) { problem(`rutba_meta's shown names ${table}, which is not a table of the database.`); continue; }
      for (const f of [].concat(list)) if (!fieldNames(table).has(f)) problem(`rutba_meta's shown names ${table}.${f}, which is not a field.`);
    }
  }
  if (meta.has('pinned')) {
    const pinned = json(meta.get('pinned'), "rutba_meta's pinned", { object: false });
    if (pinned !== undefined && !(Array.isArray(pinned) && pinned.every((p) => isObject(p) && PIN_KINDS.includes(p.kind) && isText(p.name)))) problem("rutba_meta's pinned is not a list of { kind, name }.");
  }
  if (meta.has('folded')) {
    const folded = json(meta.get('folded'), "rutba_meta's folded", { object: false });
    if (folded !== undefined && !(Array.isArray(folded) && folded.every(isText))) problem("rutba_meta's folded is not a list of text.");
  }
  if (meta.has('origin')) {
    const o = json(meta.get('origin'), "rutba_meta's origin");
    if (o && !(['template', 'sample'].includes(o.kind) && isText(o.id))) problem("rutba_meta's origin has no kind (template or sample) and id.");
  }
  if (meta.has('origin.design')) json(meta.get('origin.design'), "rutba_meta's origin.design");

  return out(version);
}

/** The command: check each file named, say what was found, exit 0 or 1. */
function main(argv) {
  const files = argv.filter((a) => !a.startsWith('--'));
  if (!files.length || argv.includes('--help')) {
    process.stderr.write('Usage: node tools/check-rbase.mjs <file.rbase> [more.rbase ...]\n');
    return 2;
  }
  let failed = false;
  for (const file of files) {
    const r = checkRbase(file);
    if (r.ok) process.stdout.write(`${file}: follows the Data Files format, design version ${r.version}.\n`);
    else {
      failed = true;
      process.stdout.write(`${file}: does not follow the Data Files format (${r.problems.length} problem${r.problems.length === 1 ? '' : 's'}):\n`);
      for (const p of r.problems) process.stdout.write(`  - ${p}\n`);
    }
    for (const n of r.notes) process.stdout.write(`  note: ${n}\n`);
  }
  return failed ? 1 : 0;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) process.exitCode = main(process.argv.slice(2));
