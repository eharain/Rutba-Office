// Imports into the database app, each as a new table: rows with a header
// (from a CSV or a sheet), and the old files' sheets (dBASE, Access, Lotus
// and the rest, read by office-formats' legacy readers). Each column's type
// is worked out from what is in it, and the import says what it made.

import { Base } from './base.js';

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
const ISO_DATETIME = /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2})?/;
const YES_NO = /^(yes|no|true|false|y|n)$/i;

/** A serial date (days since 1899-12-30) as YYYY-MM-DD, with the time when it has one. */
function fromSerial(n, withTime) {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n * 86400000));
  return withTime ? d.toISOString().slice(0, 19) : d.toISOString().slice(0, 10);
}

/**
 * The type a column's values call for: the narrowest every value fits.
 * `dated` is 'date' or 'datetime' when the column was formatted as one.
 */
export function inferType(values, dated = null) {
  const present = values.filter((v) => v !== null && v !== undefined && v !== '');
  if (!present.length) return { type: 'text', size: 255 };
  if (dated && present.every((v) => typeof v === 'number')) return { type: dated };
  if (present.every((v) => typeof v === 'boolean' || (typeof v === 'string' && YES_NO.test(v.trim())))) return { type: 'yesno' };
  const nums = present.map((v) => (typeof v === 'number' ? v : typeof v === 'string' && /^-?[\d,]*\.?\d+(e[+-]?\d+)?$/i.test(v.trim()) ? Number(v.replace(/,/g, '')) : NaN));
  if (nums.every(Number.isFinite)) return { type: nums.every(Number.isInteger) && nums.every((n) => Math.abs(n) <= Number.MAX_SAFE_INTEGER) ? 'integer' : 'number' };
  if (present.every((v) => typeof v === 'string' && ISO_DATE.test(v.trim()))) return { type: 'date' };
  if (present.every((v) => typeof v === 'string' && ISO_DATETIME.test(v.trim()))) return { type: 'datetime' };
  const longest = Math.max(...present.map((v) => String(v).length));
  return longest > 255 ? { type: 'longtext' } : { type: 'text', size: 255 };
}

/** Header names made into field names: each there, each its own. */
function fieldNames(header, count) {
  const used = new Set();
  return Array.from({ length: count }, (_, i) => {
    // A heading's line breaks and other control characters become spaces.
    let base = String(header[i] ?? '').replace(/[\u0000-\u001f]+/g, ' ').trim().replace(/"/g, "'").replace(/^(rutba_|sqlite_)/i, '').slice(0, 60) || `Field ${i + 1}`;
    if (/^id$/i.test(base)) base = `${base} (imported)`;
    let name = base;
    for (let n = 2; used.has(name.toLowerCase()); n++) name = `${base} ${n}`;
    used.add(name.toLowerCase());
    return name;
  });
}

/**
 * A table from rows: `rows[0]` the header, the rest records. `dated` maps a
 * column index to 'date' or 'datetime' where its numbers are serial dates.
 * Answers { table, fields, count, skipped }.
 */
export function importRows(base, name, rows, { dated = {} } = {}) {
  const [header = [], ...records] = rows;
  const width = Math.max(header.length, ...records.map((r) => r?.length || 0), 0);
  if (!width) throw new Error('There is nothing to import.');
  const names = fieldNames(header, width);
  const fields = names.map((n, c) => ({ name: n, ...inferType(records.map((r) => r?.[c] ?? null), dated[c] || null) }));
  let table = String(name || 'Imported').replace(/[\u0000-\u001f]+/g, ' ').trim().replace(/"/g, "'").replace(/^(rutba_|sqlite_)/i, '').slice(0, 60) || 'Imported';
  const taken = new Set(base.tables().map((t) => t.name.toLowerCase()));
  for (let n = 2; taken.has(table.toLowerCase()); n++) table = `${String(name).slice(0, 56)} ${n}`;
  base.createTable({ name: table, fields });
  let count = 0;
  let skipped = 0;
  base.db.exec('BEGIN');
  try {
    for (const r of records) {
      if (!r || r.every((v) => v === null || v === undefined || v === '')) continue;
      const values = {};
      fields.forEach((f, c) => {
        let v = r[c] ?? null;
        if ((f.type === 'date' || f.type === 'datetime') && typeof v === 'number') v = fromSerial(v, f.type === 'datetime');
        values[f.name] = v;
      });
      try {
        base.insert(table, values);
        count++;
      } catch {
        skipped++;
      }
    }
    base.db.exec('COMMIT');
  } catch (err) {
    base.db.exec('ROLLBACK');
    throw err;
  }
  return { table, fields, count, skipped };
}

/**
 * The old files' sheets (office-formats legacy readers' model) as tables, a
 * table a sheet, its first row the field names; a column whose cells are
 * formatted as dates becomes a date field.
 */
export function importSheets(base, model) {
  const made = [];
  for (const sheet of model.sheets) {
    if (!sheet.rows.length) continue;
    const dated = {};
    for (const [ref, style] of Object.entries(sheet.styles || {})) {
      const m = /^([A-Z]+)(\d+)$/.exec(ref);
      if (!m || !style?.numFmt || Number(m[2]) < 2) continue;
      const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      if (/y/i.test(style.numFmt) && /d/i.test(style.numFmt)) dated[col] = /h/i.test(style.numFmt) ? 'datetime' : 'date';
    }
    made.push(importRows(base, sheet.name, sheet.rows, { dated }));
  }
  return made;
}

export { Base };
