// PostgreSQL, for Data Files' tables linked to a database engine
// (D-2026-10-22-13): the engine holds the records; this reads its tables'
// designs and records, and writes each change to it at once.

import pg from 'pg';

const q = (name) => `"${String(name).replace(/"/g, '""')}"`;

/** schema.table, each part quoted. */
const qualified = ({ schema, name }) => (schema ? `${q(schema)}.${q(name)}` : q(name));

/** A PostgreSQL column's type as a Data Files field type. */
export function fieldType(col) {
  const t = String(col.type || '').toLowerCase();
  if (col.identity) return 'autonumber';
  if (['smallint', 'integer', 'bigint'].includes(t)) return 'integer';
  if (t === 'money') return 'currency';
  if (['numeric', 'decimal', 'real', 'double precision'].includes(t)) return 'number';
  if (t === 'boolean') return 'yesno';
  if (t === 'date') return 'date';
  if (t.startsWith('timestamp')) return 'datetime';
  if (t === 'text' || t === 'json' || t === 'jsonb' || t === 'xml') return 'longtext';
  return 'text';
}

/** A value read from PostgreSQL as Data Files keeps it: numbers as numbers, dates as ISO text, yes/no as 1 or 0. */
export function fromEngine(type, v) {
  if (v === null || v === undefined) return null;
  // Dates and times come as the engine's own text (see RAW_TIMES), so no time zone moves them.
  if ((type === 'date' || type === 'datetime') && typeof v === 'string') {
    const m = /^(\d{4}-\d{2}-\d{2})(?:[ T](\d{2}:\d{2}(?::\d{2})?))?/.exec(v);
    if (m) return type === 'date' || !m[2] ? m[1] : `${m[1]}T${m[2].length === 5 ? `${m[2]}:00` : m[2]}`;
  }
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (['integer', 'number', 'currency', 'autonumber'].includes(type) && typeof v === 'string') {
    const n = Number(v.replace(/[^0-9.eE+-]/g, ''));
    return Number.isFinite(n) ? n : v;
  }
  if (typeof v === 'object') return JSON.stringify(v);
  return v;
}

/** A value as PostgreSQL takes it for a column of `type`. */
function toEngine(type, v) {
  if (v === null || v === undefined || v === '') return null;
  if (type === 'yesno') return v === true || v === 1 || v === '1' || /^(yes|true)$/i.test(String(v));
  return v;
}

/** date, timestamp and timestamptz kept as the engine writes them, not made into a JavaScript Date. */
const RAW_TIMES = new Set([1082, 1114, 1184]);
const types = { getTypeParser: (oid, format) => (RAW_TIMES.has(oid) ? (v) => v : pg.types.getTypeParser(oid, format)) };

/** Connected to one PostgreSQL database: { tables, describe, read, insert, update, remove, close }. */
export async function connect({ host, port = 5432, database, user, ssl = false }, password) {
  const client = new pg.Client({ types, host, port: Number(port) || 5432, database, user, password, ssl: ssl ? { rejectUnauthorized: false } : false, connectionTimeoutMillis: 8000, statement_timeout: 30000 });
  // A refused or broken connection says so in words a person can act on.
  try {
    await client.connect();
  } catch (err) {
    const m = String(err.message || err);
    if (/password authentication failed|no password supplied/i.test(m)) throw new Error('The engine did not take that account and password.');
    if (/ECONNREFUSED|ENOTFOUND|ETIMEDOUT|timeout/i.test(m)) throw new Error(`The engine at ${host}:${port} could not be reached.`);
    if (/does not exist/i.test(m)) throw new Error(`There is no database called ${database} there.`);
    throw err;
  }
  client.on('error', () => {});
  const typeOf = new Map();

  return {
    engine: 'postgres',

    /** The tables and views a person may link: [{ schema, name, kind }]. */
    async tables() {
      const r = await client.query(`SELECT table_schema AS schema, table_name AS name, table_type AS kind FROM information_schema.tables
        WHERE table_schema NOT IN ('pg_catalog', 'information_schema') ORDER BY table_schema = 'public' DESC, table_schema, table_name`);
      return r.rows.map((x) => ({ schema: x.schema, name: x.name, kind: x.kind === 'VIEW' ? 'view' : 'table' }));
    },

    /** A table's design: { fields: [{ name, type, required }], key } with the key its primary key's one column, or null. */
    async describe(remote) {
      const cols = await client.query(`SELECT column_name AS name, data_type AS type, is_nullable = 'NO' AS required,
          is_identity = 'YES' OR coalesce(column_default, '') LIKE 'nextval(%' AS identity
        FROM information_schema.columns WHERE table_schema = $1 AND table_name = $2 ORDER BY ordinal_position`, [remote.schema, remote.name]);
      if (!cols.rows.length) throw new Error(`There is no table called ${remote.name} in the engine.`);
      const keys = await client.query(`SELECT a.attname AS name FROM pg_index i
          JOIN pg_attribute a ON a.attrelid = i.indrelid AND a.attnum = ANY(i.indkey)
        WHERE i.indrelid = ($1)::regclass AND i.indisprimary`, [qualified(remote)]).catch(() => ({ rows: [] }));
      const key = keys.rows.length === 1 ? keys.rows[0].name : null;
      const fields = cols.rows.map((c) => ({ name: c.name, type: fieldType({ ...c, identity: c.identity && c.name === key }), required: Boolean(c.required) && !c.identity }));
      typeOf.set(qualified(remote), new Map(fields.map((f) => [f.name, f.type])));
      return { fields, key };
    },

    /** Its records, at most `limit`, in its key's order: [{ field: value }]. */
    async read(remote, { fields, key, limit = 100000 }) {
      const r = await client.query(`SELECT ${fields.map((f) => q(f.name)).join(', ')} FROM ${qualified(remote)}${key ? ` ORDER BY ${q(key)}` : ''} LIMIT ${Math.max(1, limit | 0)}`);
      return r.rows.map((row) => Object.fromEntries(fields.map((f) => [f.name, fromEngine(f.type, row[f.name])])));
    },

    /** A record added; answers its key. */
    async insert(remote, { key }, values) {
      const types = typeOf.get(qualified(remote)) || new Map();
      const names = Object.keys(values).filter((n) => values[n] !== undefined && types.get(n) !== 'autonumber');
      const sql = names.length
        ? `INSERT INTO ${qualified(remote)} (${names.map(q).join(', ')}) VALUES (${names.map((_, i) => `$${i + 1}`).join(', ')})`
        : `INSERT INTO ${qualified(remote)} DEFAULT VALUES`;
      const r = await client.query(`${sql}${key ? ` RETURNING ${q(key)}` : ''}`, names.map((n) => toEngine(types.get(n), values[n])));
      return key ? r.rows[0]?.[key] ?? null : null;
    },

    /** A record changed, found by its key. */
    async update(remote, { key }, keyValue, values) {
      if (!key) throw new Error('A table with no primary key in the engine can be read, not changed.');
      const types = typeOf.get(qualified(remote)) || new Map();
      const names = Object.keys(values).filter((n) => n !== key);
      if (!names.length) return 0;
      const r = await client.query(`UPDATE ${qualified(remote)} SET ${names.map((n, i) => `${q(n)} = $${i + 1}`).join(', ')} WHERE ${q(key)} = $${names.length + 1}`,
        [...names.map((n) => toEngine(types.get(n), values[n])), keyValue]);
      return r.rowCount;
    },

    /** Records deleted, found by their keys. */
    async remove(remote, { key }, keyValues) {
      if (!key) throw new Error('A table with no primary key in the engine can be read, not changed.');
      if (!keyValues.length) return 0;
      const r = await client.query(`DELETE FROM ${qualified(remote)} WHERE ${q(key)} = ANY($1)`, [keyValues]);
      return r.rowCount;
    },

    async close() {
      try { await client.end(); } catch { /* already gone */ }
    },
  };
}
