// Searching and filtering a grid of records, the same for a table's records
// and a query's results (owner, 2026-10-22, D-2026-10-22-07: "there should be
// a way to further filter and text search tables and running queries exactly
// as the tables would be"; "dates and numbers to be searchable by range and
// text by matching"). Everything a person types becomes a parameter, never
// part of the SQL, and a column is named only when it is one of the grid's.
//
// A filter, by the column's type:
//   { column, kind: 'text', op: 'contains' | 'starts' | 'is', value }
//   { column, kind: 'range', from, to }      dates and numbers, either end open
//   { column, kind: 'yesno', value: true | false }
//   { column, kind: 'value', value }         filter by selection: only this value
//   { column, kind: 'notValue', value }      all but this value
//   { column, kind: 'empty' } | { column, kind: 'notEmpty' }
// and `search`, words looked for in every column.

import { q } from './types.js';

const NUMERIC = new Set(['number', 'integer', 'currency', 'autonumber']);
const DATES = new Set(['date', 'datetime']);

/** A LIKE pattern for `text` anywhere (or at the start), its own % and _ taken as themselves. */
const like = (text, start = false) => `${start ? '' : '%'}${String(text).replace(/[\\%_]/g, (c) => `\\${c}`)}%`;

/** How a filter's kind is offered for a column of this type. */
export const filterKindOf = (type) => (DATES.has(type) || NUMERIC.has(type) ? 'range' : type === 'yesno' ? 'yesno' : 'text');

/**
 * The WHERE a grid's search and filters stand for: { sql, params }, sql
 * empty when there is nothing to filter by. `columns` [{ name, type }] are
 * the grid's; a filter on any other column is refused.
 */
export function filterSql(columns, { search = '', filters = [] } = {}) {
  const byName = new Map(columns.map((c) => [c.name, c]));
  const parts = [];
  const params = [];
  const words = String(search || '').trim();
  if (words && columns.length) {
    // Every word somewhere in the record, each in any column.
    for (const word of words.split(/\s+/).slice(0, 8)) {
      parts.push(`(${columns.map((c) => `CAST(${q(c.name)} AS TEXT) LIKE ? ESCAPE '\\'`).join(' OR ')})`);
      for (let i = 0; i < columns.length; i++) params.push(like(word));
    }
  }
  for (const f of filters || []) {
    const c = byName.get(f?.column);
    if (!c) throw new Error(`There is no column called ${f?.column} to filter by.`);
    const col = q(c.name);
    const dates = DATES.has(c.type);
    switch (f.kind) {
      case 'text': {
        const v = String(f.value ?? '');
        if (!v) break;
        if (f.op === 'is') { parts.push(`CAST(${col} AS TEXT) = ? COLLATE NOCASE`); params.push(v); }
        else { parts.push(`CAST(${col} AS TEXT) LIKE ? ESCAPE '\\'`); params.push(like(v, f.op === 'starts')); }
        break;
      }
      case 'range': {
        // A date compared by its day (a date and time's first ten characters), a number as a number.
        const side = (v, op) => {
          if (v === null || v === undefined || v === '') return;
          if (dates) { parts.push(`substr(CAST(${col} AS TEXT), 1, 10) ${op} ?`); params.push(String(v).slice(0, 10)); return; }
          const n = Number(v);
          if (!Number.isFinite(n)) throw new Error(`${v} is not a number.`);
          parts.push(`${col} ${op} ?`);
          params.push(n);
        };
        side(f.from, '>=');
        side(f.to, '<=');
        break;
      }
      case 'yesno':
        parts.push(`${col} = ?`);
        params.push(f.value ? 1 : 0);
        break;
      case 'value':
        parts.push(`${col} IS ?`);
        params.push(f.value ?? null);
        break;
      case 'notValue':
        parts.push(`${col} IS NOT ?`);
        params.push(f.value ?? null);
        break;
      case 'empty':
        parts.push(`(${col} IS NULL OR CAST(${col} AS TEXT) = '')`);
        break;
      case 'notEmpty':
        parts.push(`(${col} IS NOT NULL AND CAST(${col} AS TEXT) <> '')`);
        break;
      default:
        throw new Error(`${f.kind} is not a way to filter.`);
    }
  }
  return { sql: parts.join(' AND '), params };
}

/** A column's type read from its values, for a calculated column or a total: number, date, date and time, or words. */
export function sniffType(values) {
  const seen = values.filter((v) => v !== null && v !== undefined && v !== '');
  if (!seen.length) return 'text';
  if (seen.every((v) => typeof v === 'number' || typeof v === 'bigint')) return 'number';
  if (seen.every((v) => /^\d{4}-\d{2}-\d{2}$/.test(String(v)))) return 'date';
  if (seen.every((v) => /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}/.test(String(v)))) return 'datetime';
  return 'text';
}
