// Imports into the database app, each as a new table: rows with a header
// (from a CSV, or a sheet of a workbook), and the old files' sheets (dBASE,
// Access, Lotus and the rest, read by office-formats' legacy readers). Each
// column's type is worked out from what is in it, and the import says what
// it made: how it read a column of dates written 12/04/2025, and whether it
// kept the file's own ID numbers as the table's.

import { Base } from './base.js';
import { WRITTEN_DATE, MONEY_SIGN, numberIn, fromWritten } from './types.js';

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
const ISO_DATETIME = /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2})?/;
const YES_NO = /^(yes|no|true|false|y|n)$/i;

/** A serial date (days since 1899-12-30) as YYYY-MM-DD, with the time when it has one. */
function fromSerial(n, withTime) {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n * 86400000));
  return withTime ? d.toISOString().slice(0, 19) : d.toISOString().slice(0, 10);
}

/**
 * Written dates' order: day first unless a value says otherwise. A day above
 * 12 in the first place settles it day first, in the second month first;
 * both in one column, and it is not a column of dates. Null when a value is
 * not a date that can be.
 */
function writtenOrder(strings) {
  let dayFirst = false;
  let monthFirst = false;
  for (const s of strings) {
    const m = WRITTEN_DATE.exec(s.trim());
    if (!m) return null;
    const [a, b] = [Number(m[1]), Number(m[2])];
    if (!a || !b || a > 31 || b > 31 || (a > 12 && b > 12)) return null;
    if (a > 12) dayFirst = true;
    if (b > 12) monthFirst = true;
  }
  if (dayFirst && monthFirst) return null;
  return monthFirst ? 'mdy' : 'dmy';
}

/**
 * The type a column's values call for: the narrowest every value fits.
 * `dated` is 'date' or 'datetime' when the column was formatted as one;
 * `money` when it was formatted as an amount (a sheet's £#,##0.00). A
 * column of written dates (12/04/2025) answers the order it was read in.
 */
export function inferType(values, dated = null, money = false) {
  const present = values.filter((v) => v !== null && v !== undefined && v !== '');
  if (!present.length) return { type: 'text', size: 255 };
  if (dated && present.every((v) => typeof v === 'number')) return { type: dated };
  if (present.every((v) => typeof v === 'boolean' || (typeof v === 'string' && YES_NO.test(v.trim())))) return { type: 'yesno' };
  const nums = present.map(numberIn);
  if (nums.every(Number.isFinite)) {
    // Money: a sign before an amount, a sheet's money format, or every amount written with two places (1.10, 3.00).
    const strings = present.filter((v) => typeof v === 'string');
    const signed = strings.some((v) => MONEY_SIGN.test(v.trim()));
    const twoPlaces = strings.length === present.length && strings.every((v) => /\.\d{2}$/.test(v.trim()));
    if (signed || twoPlaces || (money && !nums.every(Number.isInteger))) return { type: 'currency' };
    return { type: nums.every(Number.isInteger) && nums.every((n) => Math.abs(n) <= Number.MAX_SAFE_INTEGER) ? 'integer' : 'number' };
  }
  if (present.every((v) => typeof v === 'string' && ISO_DATE.test(v.trim()))) return { type: 'date' };
  if (present.every((v) => typeof v === 'string' && ISO_DATETIME.test(v.trim()))) return { type: 'datetime' };
  if (present.every((v) => typeof v === 'string')) {
    const order = writtenOrder(present);
    if (order && present.every((v) => fromWritten(v, order))) return { type: present.some((v) => WRITTEN_DATE.exec(v.trim())[4] !== undefined) ? 'datetime' : 'date', order };
  }
  const longest = Math.max(...present.map((v) => String(v).length));
  return longest > 255 ? { type: 'longtext' } : { type: 'text', size: 255 };
}

/** Header names made into field names: each there, each its own. `kept` is the ID column used as the key, whose name stays. */
function fieldNames(header, count, kept = -1) {
  const used = new Set();
  return Array.from({ length: count }, (_, i) => {
    // A heading's line breaks and other control characters become spaces.
    let base = String(header[i] ?? '').replace(/[\u0000-\u001f]+/g, ' ').trim().replace(/"/g, "'").replace(/^(rutba_|sqlite_)/i, '').slice(0, 60) || `Field ${i + 1}`;
    if (/^id$/i.test(base) && i !== kept) base = `${base} (imported)`;
    let name = base;
    for (let n = 2; used.has(name.toLowerCase()); n++) name = `${base} ${n}`;
    used.add(name.toLowerCase());
    return name;
  });
}

/**
 * The column named ID whose numbers can be the table's own key: whole
 * numbers above nought, one in every record, no two the same. -1 when there
 * is none, and the table numbers its records itself.
 */
function keyColumn(header, records) {
  const ids = header.map((h, i) => (/^id$/i.test(String(h ?? '').trim()) ? i : -1)).filter((i) => i >= 0);
  if (ids.length !== 1) return -1;
  const c = ids[0];
  const seen = new Set();
  for (const r of records) {
    const n = numberIn(r?.[c]);
    if (!Number.isInteger(n) || n < 1 || n > Number.MAX_SAFE_INTEGER || seen.has(n)) return -1;
    seen.add(n);
  }
  return seen.size ? c : -1;
}

/**
 * A table from rows: `rows[0]` the header, the rest records. `dated` maps a
 * column index to 'date' or 'datetime' where its numbers are serial dates,
 * `money` to true where a sheet formats it as an amount. Answers { table,
 * fields, count, skipped, key, dates }: `key` the file's ID column kept as
 * the table's AutoNumber (or null), `dates` [{ field, order }] for each
 * column of written dates, `order` 'dmy' (day first) or 'mdy'.
 */
export function importRows(base, name, rows, { dated = {}, money = {} } = {}) {
  const [header = [], ...all] = rows;
  // A record with nothing in it is not one.
  const records = all.filter((r) => r && !r.every((v) => v === null || v === undefined || v === ''));
  const width = Math.max(header.length, ...records.map((r) => r?.length || 0), 0);
  if (!width) throw new Error('There is nothing to import.');
  const key = keyColumn(header, records);
  const names = fieldNames(header, width, key);
  const inferred = names.map((n, c) => (c === key ? { type: 'autonumber' } : inferType(records.map((r) => r?.[c] ?? null), dated[c] || null, Boolean(money[c]))));
  const fields = names.map((n, c) => {
    const { order, ...type } = inferred[c];
    return { name: n, ...type };
  });
  let table = String(name || 'Imported').replace(/[\u0000-\u001f]+/g, ' ').trim().replace(/"/g, "'").replace(/^(rutba_|sqlite_)/i, '').slice(0, 60) || 'Imported';
  const taken = new Set(base.tables().map((t) => t.name.toLowerCase()));
  for (let n = 2; taken.has(table.toLowerCase()); n++) table = `${String(name).slice(0, 56)} ${n}`;
  base.createTable({ name: table, fields });
  let count = 0;
  let skipped = 0;
  base.db.exec('BEGIN');
  try {
    for (const r of records) {
      const values = {};
      fields.forEach((f, c) => {
        let v = r[c] ?? null;
        if ((f.type === 'date' || f.type === 'datetime') && typeof v === 'number') v = fromSerial(v, f.type === 'datetime');
        else if ((f.type === 'date' || f.type === 'datetime') && inferred[c].order && typeof v === 'string' && v.trim()) v = fromWritten(v, inferred[c].order);
        else if (['currency', 'number', 'integer', 'autonumber'].includes(f.type) && typeof v === 'string' && v.trim()) v = numberIn(v);
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
  const dates = fields.flatMap((f, c) => (inferred[c].order ? [{ field: f.name, order: inferred[c].order }] : []));
  return { table, fields, count, skipped, key: key >= 0 ? fields[key].name : null, dates };
}

/**
 * Sheets as tables, a table a sheet, its first row the field names: the old
 * files' (office-formats legacy readers' model) and a workbook's (.xlsx,
 * .ods). A column whose cells are formatted as dates becomes a date field;
 * one formatted as an amount, with a money sign or two places, Currency.
 */
export function importSheets(base, model) {
  const made = [];
  for (const sheet of model.sheets) {
    if (!sheet.rows.length) continue;
    const dated = {};
    const money = {};
    for (const [ref, style] of Object.entries(sheet.styles || {})) {
      const m = /^([A-Z]+)(\d+)$/.exec(ref);
      if (!m || !style?.numFmt || Number(m[2]) < 2) continue;
      const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      // A format's colours and quoted words are not its parts: "£"#,##0.00 is money, [Red]0.00 two places.
      const fmt = String(style.numFmt).replace(/\[[^\]]*\]/g, (s) => (/[£$€]/.test(s) ? '£' : '')).replace(/"([^"]*)"/g, (s, w) => (/[£$€]/.test(w) ? '£' : ''));
      if (/y/i.test(fmt) && /d/i.test(fmt)) dated[col] = /h/i.test(fmt) ? 'datetime' : 'date';
      else if (/[£$€]|0\.00(?!0)/.test(fmt)) money[col] = true;
    }
    made.push(importRows(base, sheet.name, sheet.rows, { dated, money }));
  }
  return made;
}

/**
 * Another program's database (A2.20, `from` opened with Base.foreign) made
 * into tables of `base`, a Data Files database: each of its tables with the
 * fields Data Files showed it with (their types, sizes, required and
 * unique), its INTEGER PRIMARY KEY kept as the AutoNumber with its numbers,
 * and its foreign keys made relationships where the field they point at is
 * the AutoNumber or unique. A field holding pictures or other binary values
 * has nothing to go into and is left out. Answers { tables, left,
 * relationships }: each table made { table, from, count, skipped }, the
 * fields left out ("Table.Field"), and how many relationships were made.
 */
export function copyForeign(from, base) {
  const made = [];
  const left = [];
  const names = new Map();
  // A name of the other program's that a person could not give (a quotation mark, too long) made one they could.
  const fit = (name, taken, fallback) => {
    const b = String(name).replace(/[\u0000-\u001f]+/g, ' ').trim().replace(/"/g, "'").replace(/^(rutba_|sqlite_)/i, '').slice(0, 60) || fallback;
    let n = b;
    for (let i = 2; taken.has(n.toLowerCase()); i++) n = `${b.slice(0, 56)} ${i}`;
    taken.add(n.toLowerCase());
    return n;
  };
  const tableNames = new Set();
  for (const t of from.tables()) {
    const all = from.fields(t.name);
    const binary = new Set(all.filter((f) => from.foreign?.fixed?.has(`${t.name}\u0000${f.name}`)).map((f) => f.name));
    for (const b of binary) left.push(`${t.name}.${b}`);
    const keep = all.filter((f) => !binary.has(f.name));
    if (!keep.length) continue;
    const key = from.foreign?.keys?.get(t.name) || null;
    const taken = new Set();
    const fieldNames = new Map(keep.map((f, i) => [f.name, fit(f.name, taken, `Field ${i + 1}`)]));
    const fields = keep.map((f) => {
      const { readOnly, ...plain } = f;
      return { ...plain, name: fieldNames.get(f.name), type: f.name === key ? 'autonumber' : f.type, ...(f.name === key ? { required: false, unique: false } : {}) };
    });
    // One AutoNumber at most: the key's, or one the table is given.
    const numbered = fields.filter((f) => f.type === 'autonumber');
    for (const f of numbered.slice(1)) f.type = 'integer';
    const table = base.createTable({ name: fit(t.name, tableNames, 'Table'), fields });
    let count = 0;
    let skipped = 0;
    base.db.exec('BEGIN');
    try {
      for (let offset = 0; ; offset += 5000) {
        const page = from.rows(t.name, { offset, limit: 5000 });
        for (const r of page) {
          try {
            base.insert(table, Object.fromEntries(keep.map((f) => [fieldNames.get(f.name), r[f.name] ?? null])));
            count++;
          } catch {
            skipped++;
          }
        }
        if (page.length < 5000) break;
      }
      base.db.exec('COMMIT');
    } catch (err) {
      base.db.exec('ROLLBACK');
      throw err;
    }
    names.set(t.name, { table, fields: fieldNames });
    made.push({ table, from: t.name, count, skipped });
  }
  let relationships = 0;
  for (const rel of from.relationships()) {
    const a = names.get(rel.from.table);
    const b = names.get(rel.to.table);
    const fromField = a?.fields.get(rel.from.field);
    const toField = b?.fields.get(rel.to.field);
    if (!fromField || !toField) continue;
    try {
      base.addRelationship({ from: { table: a.table, field: fromField }, to: { table: b.table, field: toField }, cascadeUpdate: rel.cascadeUpdate, cascadeDelete: rel.cascadeDelete });
      relationships++;
    } catch { /* the field it points at is neither the AutoNumber nor unique here, or records point at nothing: the tables are kept without it */ }
  }
  return { tables: made, left, relationships };
}

export { Base, fromWritten };
