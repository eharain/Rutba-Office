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

// A date as the UK writes one, 12/04/2025 (or 12.04.2025, 12-04-2025), with a time after it or not.
export const WRITTEN_DATE = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
// A number as a file writes it: 1,250.50, -3, 1e3; with a pound, dollar or euro sign before it, an amount of money.
const NUMBER = /^-?[\d,]*\.?\d+(e[+-]?\d+)?$/i;
export const MONEY_SIGN = /^(-?)\s*[£$€]\s*/;

/** A number written in a file, its money sign and thousands commas taken off; NaN when it is not one. */
export function numberIn(v) {
  if (typeof v === 'number') return v;
  if (typeof v !== 'string') return NaN;
  const s = v.trim().replace(MONEY_SIGN, '$1');
  return NUMBER.test(s) ? Number(s.replace(/,/g, '')) : NaN;
}

/** A written date as the file keeps one, YYYY-MM-DD (and its time), read in `order` ('dmy', day first, or 'mdy'); null when there is no such day. */
export function fromWritten(text, order = 'dmy') {
  const m = WRITTEN_DATE.exec(String(text).trim());
  if (!m) return null;
  const [day, month] = order === 'mdy' ? [Number(m[2]), Number(m[1])] : [Number(m[1]), Number(m[2])];
  const year = Number(m[3]);
  const d = new Date(Date.UTC(year, month - 1, day));
  if (d.getUTCFullYear() !== year || d.getUTCMonth() !== month - 1 || d.getUTCDate() !== day) return null;
  const date = d.toISOString().slice(0, 10);
  if (m[4] === undefined) return date;
  const hh = Number(m[4]);
  const mm = Number(m[5]);
  const ss = Number(m[6] || 0);
  if (hh > 23 || mm > 59 || ss > 59) return null;
  return `${date}T${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}:${String(ss).padStart(2, '0')}`;
}

/** Why a value does not fit a field of a type, in words that follow "a Joined that". */
export const UNFIT = { text: (f) => `is longer than ${Number(f.size) || 255} characters`, number: () => 'is not a number', currency: () => 'is not an amount', integer: () => 'is not a whole number', date: () => 'is not a date', datetime: () => 'is not a date and time', yesno: () => 'is not Yes or No' };

/**
 * A value kept in a field, as a field of another type (or size) keeps it,
 * when a field's type is changed: a number from its text (1,250.50, £3.10),
 * a date from 2026-10-21 or 21/10/2026, a whole number only when it is one.
 * `undefined` when it does not fit, so nothing is changed into something else
 * without a word.
 */
export function fitTo(field, value) {
  if (value === undefined || value === null || value === '') return field.type === 'yesno' ? 0 : null;
  switch (field.type) {
    case 'text': {
      const s = String(value);
      return s.length > (Number(field.size) || 255) ? undefined : s;
    }
    case 'longtext': return String(value);
    case 'number': case 'currency': {
      const n = numberIn(value);
      if (!Number.isFinite(n)) return undefined;
      return field.type === 'currency' ? Math.round(n * 10000) / 10000 : n;
    }
    case 'integer': case 'autonumber': {
      const n = numberIn(value);
      return Number.isInteger(n) ? n : undefined;
    }
    case 'yesno': {
      if (value === 1 || value === 0) return value;
      const s = String(value).trim();
      if (/^(1|yes|true|y)$/i.test(s)) return 1;
      if (/^(0|no|false|n)$/i.test(s)) return 0;
      return undefined;
    }
    case 'date': {
      const s = String(value).trim();
      if (/^\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2})?)?$/.test(s)) return s.slice(0, 10);
      return fromWritten(s)?.slice(0, 10) ?? undefined;
    }
    case 'datetime': {
      const s = String(value).trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return `${s}T00:00:00`;
      const m = /^(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2})(:\d{2})?/.exec(s);
      if (m) return `${m[1]}T${m[2]}${m[3] || ':00'}`;
      const w = fromWritten(s);
      return w ? (w.length === 10 ? `${w}T00:00:00` : w) : undefined;
    }
    default: return value;
  }
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
      let s = value instanceof Date ? value.toISOString().slice(0, 19) : String(value).trim().replace(' ', 'T');
      // A time to the minute (a form's date and time box gives 2026-10-22T17:24) is kept to the second, :00.
      if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(s)) s += ':00';
      if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(s)) throw new Error(`${field.name} takes a date and a time, as 2026-10-21 14:30:00.`);
      return s.slice(0, 19);
    }
    default: return String(value);
  }
}
