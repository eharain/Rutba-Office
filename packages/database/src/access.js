// People, roles and what each role may see and do (D-2026-10-22-11): a
// database that is protected asks who you are as it opens, and gives each
// person what their role allows, table by table, field by field, and query,
// form, report and print template by name.
//
// What a role may read is enforced beneath everything that reads: each table
// is shadowed, for that connection only, by a TEMP VIEW of the same name that
// holds the records and fields the role may see. SQLite looks in temp before
// main, so a query, a report, a print template and a lookup all read the
// view without knowing it. Writes go to main.<table>, after the role is
// checked here. A role that may change the design sees everything: there is
// nothing for it to be shadowed from.

import crypto from 'node:crypto';
import { q } from './types.js';

/** What a role may do with a table, each level including the ones before it. */
export const LEVELS = ['hidden', 'read', 'add', 'change', 'delete'];

export const ADMIN = 'Admin';

const levelOf = (access) => Math.max(0, LEVELS.indexOf(access));

/** The tables people and roles are kept in, made with the rest of the design. */
export const ACCESS_DESIGN = `
CREATE TABLE IF NOT EXISTS rutba_roles (name TEXT PRIMARY KEY COLLATE NOCASE, rights TEXT NOT NULL, position INTEGER);
CREATE TABLE IF NOT EXISTS rutba_people (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE COLLATE NOCASE, role TEXT NOT NULL,
  hash TEXT, account TEXT, created TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS rutba_trail (n INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT NOT NULL, person INTEGER, tbl TEXT NOT NULL, row INTEGER NOT NULL,
  action TEXT NOT NULL, detail TEXT);
CREATE INDEX IF NOT EXISTS rutba_trail_row ON rutba_trail (tbl, row);
`;

/* ── passwords ─────────────────────────────────────────────────────── */

const SCRYPT = { N: 16384, r: 8, p: 1 };

/** A password kept as scrypt's answer and its salt, never as itself. */
export function hashPassword(password) {
  const pw = String(password ?? '');
  if (pw.length < 4) throw new Error('A password has at least 4 characters.');
  const salt = crypto.randomBytes(16);
  const hash = crypto.scryptSync(pw, salt, 32, SCRYPT);
  return `scrypt$${SCRYPT.N}$${SCRYPT.r}$${SCRYPT.p}$${salt.toString('base64')}$${hash.toString('base64')}`;
}

/** Whether `password` is the one `kept` was made from. */
export function checkPassword(password, kept) {
  const parts = String(kept || '').split('$');
  if (parts.length !== 6 || parts[0] !== 'scrypt') return false;
  const [, N, r, p, salt, hash] = parts;
  const want = Buffer.from(hash, 'base64');
  const got = crypto.scryptSync(String(password ?? ''), Buffer.from(salt, 'base64'), want.length, { N: Number(N), r: Number(r), p: Number(p) });
  return crypto.timingSafeEqual(want, got);
}

/* ── rights ────────────────────────────────────────────────────────── */

/**
 * A role's rights, every part there: { design, tables: { default, each:
 * { [table]: { access, own, fields: { [field]: 'hidden' | 'read' } } } },
 * hidden: { queries, forms, reports, documents } }. A table's entry may be
 * given as its access alone ('read').
 */
export function normalRights(rights = {}) {
  const tables = rights.tables || {};
  const each = {};
  for (const [name, t] of Object.entries(tables.each || {})) {
    const entry = typeof t === 'string' ? { access: t } : t || {};
    const fields = {};
    for (const [f, v] of Object.entries(entry.fields || {})) if (v === 'hidden' || v === 'read') fields[f] = v;
    each[name] = { access: LEVELS.includes(entry.access) ? entry.access : 'hidden', own: Boolean(entry.own), fields };
  }
  const list = (x) => [...new Set((Array.isArray(x) ? x : []).map(String))];
  const hidden = rights.hidden || {};
  return {
    design: Boolean(rights.design),
    tables: { default: LEVELS.includes(tables.default) ? tables.default : 'read', each },
    hidden: { queries: list(hidden.queries), forms: list(hidden.forms), reports: list(hidden.reports), documents: list(hidden.documents) },
  };
}

/** What a role may do with one table: { access, own, fields }. */
export function tableRight(rights, table) {
  return rights.tables.each[table] || { access: rights.tables.default, own: false, fields: {} };
}

/** Whether `rights` allow `want` ('read', 'add', 'change', 'delete') on `table`. */
export function allows(rights, table, want) {
  return levelOf(tableRight(rights, table).access) >= levelOf(want);
}

/** The roles a new protected database starts with, besides Admin: one that reads, one that works. */
export const STARTING_ROLES = [
  { name: 'Staff', rights: { tables: { default: 'change' } } },
  { name: 'Read only', rights: { tables: { default: 'read' } } },
];

/* ── the views a role reads through ────────────────────────────────── */

const lit = (v) => (typeof v === 'number' ? String(v) : `'${String(v).replace(/'/g, "''")}'`);

/**
 * The TEMP VIEW that stands in for `table` for a person of `rights`: its
 * rowid kept as a column (a view has none of its own), the fields the role may
 * see, and, for a role that sees only its own records, those its person added.
 */
export function viewSql(table, fields, rights, personId) {
  const right = tableRight(rights, table);
  if (right.access === 'hidden') return `CREATE TEMP VIEW ${q(table)} AS SELECT rowid AS rowid FROM main.${q(table)} WHERE 0`;
  const shown = fields.filter((f) => right.fields[f.name] !== 'hidden').map((f) => q(f.name));
  const own = right.own
    ? ` WHERE rowid IN (SELECT row FROM main.rutba_trail WHERE tbl = ${lit(table)} AND action = 'add' AND person = ${lit(personId ?? -1)})`
    : '';
  return `CREATE TEMP VIEW ${q(table)} AS SELECT rowid AS rowid${shown.length ? `, ${shown.join(', ')}` : ''} FROM main.${q(table)}${own}`;
}
