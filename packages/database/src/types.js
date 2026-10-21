// The kinds of field a person chooses in the database app, each a real
// SQLite column with a CHECK for what SQLite does not enforce itself, so the
// file keeps its promises whatever program writes to it next.

/** The field types, in the order the designer offers them. */
export const FIELD_TYPES = ['autonumber', 'text', 'longtext', 'number', 'integer', 'currency', 'date', 'datetime', 'yesno'];

/** An identifier, quoted for SQL. */
export const q = (name) => `"${String(name).replace(/"/g, '""')}"`;

/** A literal for a DEFAULT clause. */
export function literal(value) {
  if (value == null) return 'NULL';
  if (typeof value === 'boolean') return value ? '1' : '0';
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error('A default must be a finite number.');
    return String(value);
  }
  return `'${String(value).replace(/'/g, "''")}'`;
}

const DATE = "'[0-9][0-9][0-9][0-9]-[0-1][0-9]-[0-3][0-9]'";
const DATETIME = "'[0-9][0-9][0-9][0-9]-[0-1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9]*'";

/**
 * A field's column definition. `field`: { name, type, size?, required?,
 * unique?, default? }. An AutoNumber is the table's own key.
 */
export function columnSql(field) {
  const c = q(field.name);
  const type = FIELD_TYPES.includes(field.type) ? field.type : 'text';
  const nullable = (inner) => `CHECK (${c} IS NULL OR (${inner}))`;
  let sql;
  switch (type) {
    case 'autonumber': return `${c} INTEGER PRIMARY KEY AUTOINCREMENT`;
    case 'text': {
      const size = Math.max(1, Math.min(4000, Number(field.size) || 255));
      sql = `${c} TEXT ${nullable(`typeof(${c}) = 'text' AND length(${c}) <= ${size}`)}`;
      break;
    }
    case 'longtext': sql = `${c} TEXT ${nullable(`typeof(${c}) = 'text'`)}`; break;
    case 'number': sql = `${c} REAL ${nullable(`typeof(${c}) IN ('real', 'integer')`)}`; break;
    case 'integer': sql = `${c} INTEGER ${nullable(`typeof(${c}) = 'integer'`)}`; break;
    case 'currency': sql = `${c} REAL ${nullable(`typeof(${c}) IN ('real', 'integer') AND round(${c}, 4) = ${c}`)}`; break;
    case 'date': sql = `${c} TEXT ${nullable(`${c} GLOB ${DATE}`)}`; break;
    case 'datetime': sql = `${c} TEXT ${nullable(`${c} GLOB ${DATETIME}`)}`; break;
    case 'yesno': return `${c} INTEGER NOT NULL DEFAULT ${field.default ? 1 : 0} CHECK (${c} IN (0, 1))`;
    default: sql = `${c} TEXT`;
  }
  if (field.required) sql += ' NOT NULL';
  if (field.unique) sql += ' UNIQUE';
  if (field.default !== undefined && field.default !== null) sql += ` DEFAULT ${literal(field.default)}`;
  return sql;
}

/**
 * A value as a field of this type keeps it: text trimmed to nothing is null,
 * a number from its text, a date as YYYY-MM-DD, a yes/no as 1 or 0. Throws
 * with a sentence a person can act on when it cannot be.
 */
export function toStored(field, value) {
  if (value === undefined || value === null || value === '') return field.type === 'yesno' ? 0 : null;
  switch (field.type) {
    case 'number': case 'currency': {
      const n = typeof value === 'number' ? value : Number(String(value).replace(/[,\s]/g, ''));
      if (!Number.isFinite(n)) throw new Error(`${field.name} takes a number.`);
      return field.type === 'currency' ? Math.round(n * 10000) / 10000 : n;
    }
    case 'integer': case 'autonumber': {
      const n = typeof value === 'number' ? value : Number(String(value).replace(/[,\s]/g, ''));
      if (!Number.isInteger(n)) throw new Error(`${field.name} takes a whole number.`);
      return n;
    }
    case 'yesno': return value === true || value === 1 || /^(1|yes|true|y)$/i.test(String(value)) ? 1 : 0;
    case 'date': {
      const d = value instanceof Date ? value : null;
      const s = d ? d.toISOString().slice(0, 10) : String(value).trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error(`${field.name} takes a date, as 2026-10-21.`);
      return s;
    }
    case 'datetime': {
      const s = value instanceof Date ? value.toISOString().slice(0, 19) : String(value).trim().replace(' ', 'T');
      if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(s)) throw new Error(`${field.name} takes a date and a time, as 2026-10-21 14:30:00.`);
      return s.slice(0, 19);
    }
    default: return String(value);
  }
}
