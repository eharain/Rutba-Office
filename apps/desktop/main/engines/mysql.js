// MySQL and MariaDB, for Data Files' tables linked to a database engine
// (D-2026-10-22-13): the engine holds the records; this reads its tables'
// designs and records, and writes each change to it at once.

import mysql from 'mysql2/promise';

const q = (name) => `\`${String(name).replace(/`/g, '``')}\``;
const qualified = ({ schema, name }) => (schema ? `${q(schema)}.${q(name)}` : q(name));

/** A MySQL column as a Data Files field type. */
export function fieldType(col) {
  const t = String(col.type || '').toLowerCase();
  const full = String(col.columnType || '').toLowerCase();
  if (col.identity) return 'autonumber';
  if (full === 'tinyint(1)' || t === 'bit' || t === 'boolean') return 'yesno';
  if (['tinyint', 'smallint', 'mediumint', 'int', 'integer', 'bigint', 'year'].includes(t)) return 'integer';
  if (['decimal', 'numeric', 'float', 'double', 'real'].includes(t)) return 'number';
  if (t === 'date') return 'date';
  if (t === 'datetime' || t === 'timestamp') return 'datetime';
  if (['text', 'mediumtext', 'longtext', 'json'].includes(t)) return 'longtext';
  return 'text';
}

/** A value read from MySQL as Data Files keeps it. */
export function fromEngine(type, v) {
  if (v === null || v === undefined) return null;
  if (Buffer.isBuffer(v)) return type === 'yesno' ? (v[0] ? 1 : 0) : v.toString('utf8');
  if ((type === 'date' || type === 'datetime') && typeof v === 'string') {
    const m = /^(\d{4}-\d{2}-\d{2})(?:[ T](\d{2}:\d{2}(?::\d{2})?))?/.exec(v);
    if (m) return type === 'date' || !m[2] ? m[1] : `${m[1]}T${m[2].length === 5 ? `${m[2]}:00` : m[2]}`;
  }
  if (type === 'yesno') return v ? 1 : 0;
  if (['integer', 'number', 'currency', 'autonumber'].includes(type) && typeof v === 'string') {
    const n = Number(v);
    return Number.isFinite(n) ? n : v;
  }
  if (typeof v === 'object') return JSON.stringify(v);
  return v;
}

/** A value as MySQL takes it for a column of `type`. */
function toEngine(type, v) {
  if (v === null || v === undefined || v === '') return null;
  if (type === 'yesno') return v === true || v === 1 || v === '1' || /^(yes|true)$/i.test(String(v)) ? 1 : 0;
  if (type === 'datetime' && typeof v === 'string') return v.replace('T', ' ');
  return v;
}

/** Connected to one MySQL or MariaDB database: { tables, describe, read, insert, update, remove, close }. */
export async function connect({ host, port = 3306, database, user, ssl = false }, password) {
  let conn;
  try {
    conn = await mysql.createConnection({ host, port: Number(port) || 3306, database, user, password, ssl: ssl ? { rejectUnauthorized: false } : undefined, dateStrings: true, supportBigNumbers: true, bigNumberStrings: true, connectTimeout: 8000 });
  } catch (err) {
    const m = String(err.message || err);
    if (/Access denied/i.test(m)) throw new Error('The engine did not take that account and password.');
    if (/ECONNREFUSED|ENOTFOUND|ETIMEDOUT|timeout/i.test(m)) throw new Error(`The engine at ${host}:${port} could not be reached.`);
    if (/Unknown database/i.test(m)) throw new Error(`There is no database called ${database} there.`);
    throw err;
  }
  conn.on('error', () => {});
  const typeOf = new Map();

  return {
    engine: 'mysql',

    async tables() {
      const [rows] = await conn.query(`SELECT table_schema AS \`schema\`, table_name AS name, table_type AS kind FROM information_schema.tables WHERE table_schema = DATABASE() ORDER BY table_name`);
      return rows.map((x) => ({ schema: x.schema, name: x.name, kind: /VIEW/i.test(x.kind) ? 'view' : 'table' }));
    },

    async describe(remote) {
      const [cols] = await conn.query(`SELECT column_name AS name, data_type AS type, column_type AS columnType, is_nullable = 'NO' AS required,
          extra LIKE '%auto_increment%' AS identity, column_key = 'PRI' AS pk
        FROM information_schema.columns WHERE table_schema = ? AND table_name = ? ORDER BY ordinal_position`, [remote.schema, remote.name]);
      if (!cols.length) throw new Error(`There is no table called ${remote.name} in the engine.`);
      const pks = cols.filter((c) => Number(c.pk));
      const key = pks.length === 1 ? pks[0].name : null;
      const fields = cols.map((c) => ({ name: c.name, type: fieldType({ ...c, identity: Number(c.identity) && c.name === key }), required: Boolean(Number(c.required)) && !Number(c.identity) }));
      typeOf.set(qualified(remote), new Map(fields.map((f) => [f.name, f.type])));
      return { fields, key };
    },

    async read(remote, { fields, key, limit = 100000, since = null }) {
      const where = since?.field && since.value != null ? ` WHERE ${q(since.field)} >= ?` : '';
      const value = where && typeof since.value === 'string' ? since.value.replace('T', ' ') : since?.value;
      const [rows] = await conn.query(`SELECT ${fields.map((f) => q(f.name)).join(', ')} FROM ${qualified(remote)}${where}${key ? ` ORDER BY ${q(key)}` : ''} LIMIT ${Math.max(1, limit | 0)}`, where ? [value] : []);
      return rows.map((row) => Object.fromEntries(fields.map((f) => [f.name, fromEngine(f.type, row[f.name])])));
    },

    async keys(remote, { key, type = 'integer' }) {
      const [rows] = await conn.query(`SELECT ${q(key)} AS k FROM ${qualified(remote)}`);
      return rows.map((row) => fromEngine(type, row.k));
    },

    async insert(remote, { key }, values) {
      const types = typeOf.get(qualified(remote)) || new Map();
      const names = Object.keys(values).filter((n) => values[n] !== undefined && types.get(n) !== 'autonumber');
      const sql = names.length
        ? `INSERT INTO ${qualified(remote)} (${names.map(q).join(', ')}) VALUES (${names.map(() => '?').join(', ')})`
        : `INSERT INTO ${qualified(remote)} () VALUES ()`;
      const [r] = await conn.query(sql, names.map((n) => toEngine(types.get(n), values[n])));
      if (!key) return null;
      return types.get(key) === 'autonumber' ? r.insertId : values[key] ?? null;
    },

    async update(remote, { key }, keyValue, values) {
      if (!key) throw new Error('A table with no primary key in the engine can be read, not changed.');
      const types = typeOf.get(qualified(remote)) || new Map();
      const names = Object.keys(values).filter((n) => n !== key);
      if (!names.length) return 0;
      const [r] = await conn.query(`UPDATE ${qualified(remote)} SET ${names.map((n) => `${q(n)} = ?`).join(', ')} WHERE ${q(key)} = ?`,
        [...names.map((n) => toEngine(types.get(n), values[n])), keyValue]);
      return r.affectedRows;
    },

    async remove(remote, { key }, keyValues) {
      if (!key) throw new Error('A table with no primary key in the engine can be read, not changed.');
      if (!keyValues.length) return 0;
      const [r] = await conn.query(`DELETE FROM ${qualified(remote)} WHERE ${q(key)} IN (?)`, [keyValues]);
      return r.affectedRows;
    },

    async close() {
      try { await conn.end(); } catch { /* already gone */ }
    },
  };
}
