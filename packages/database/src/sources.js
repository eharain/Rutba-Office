// Other sources a table is read from (D-2026-10-22-13, step 5): the rows of a
// CSV file or a spreadsheet's sheet, and the records a web address answers
// in JSON, made into a table's fields and records. Each column's type is
// worked out as Import works it out (import.js), and a value that does not
// fit it is kept as it came rather than lost.

import { inferType } from './import.js';
import { toStored } from './types.js';

/** Names from a header or a JSON record's keys made into field names: each there, each its own, none the file keeps for itself. */
export function sourceFieldNames(names) {
  const used = new Set();
  return names.map((raw, i) => {
    let base = String(raw ?? '').replace(/[\u0000-\u001f]+/g, ' ').trim().replace(/"/g, "'").replace(/^(rutba_|sqlite_)/i, '').slice(0, 60) || `Field ${i + 1}`;
    let name = base;
    for (let n = 2; used.has(name.toLowerCase()); n++) name = `${base} ${n}`;
    used.add(name.toLowerCase());
    return name;
  });
}

/** A serial date (days since 1899-12-30) as YYYY-MM-DD, with the time when it has one. */
function fromSerial(n, withTime) {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n * 86400000));
  return withTime ? d.toISOString().slice(0, 19) : d.toISOString().slice(0, 10);
}

/** A value as its field keeps it; one that does not fit is kept as the source gave it. */
function stored(field, v) {
  if (v === null || v === undefined || v === '') return null;
  if ((field.type === 'date' || field.type === 'datetime') && typeof v === 'number') return fromSerial(v, field.type === 'datetime');
  if (typeof v === 'boolean' && field.type !== 'yesno') v = v ? 'Yes' : 'No';
  try {
    return toStored(field, typeof v === 'string' ? v.trim() : v);
  } catch {
    return typeof v === 'object' ? JSON.stringify(v) : v;
  }
}

/**
 * A sheet's rows (or a CSV's) as a table: the row numbered `header` (1, the
 * first, unless the person says) names the fields, the rows below it are the
 * records, empty ones left out. `dated` maps a column to 'date' or
 * 'datetime' where its numbers are serial dates. Answers { fields, rows }.
 */
export function tableFromRows(allRows, { header = 1, dated = {} } = {}) {
  const at = Math.max(0, (Number(header) || 1) - 1);
  const head = allRows[at] || [];
  const body = allRows.slice(at + 1).filter((r) => r && r.some((v) => v !== null && v !== undefined && String(v).trim() !== ''));
  let width = Math.max(head.length, ...body.map((r) => r.length), 0);
  // A column with no heading and nothing under it is not a field.
  while (width > 0 && !String(head[width - 1] ?? '').trim() && body.every((r) => r[width - 1] == null || String(r[width - 1]).trim() === '')) width--;
  if (!width) throw new Error('There are no columns to read.');
  const names = sourceFieldNames(Array.from({ length: width }, (_, c) => head[c]));
  const fields = names.map((name, c) => ({ name, ...inferType(body.map((r) => r[c] ?? null), dated[c] || null) }));
  // A date column that is words as well (a totals row's "Total" under the
  // weeks) is a text field, and its dates are kept as dates, not as the
  // sheet's day numbers (46027).
  const asDate = (f, c, v) => (dated[c] && typeof v === 'number' && f.type !== 'date' && f.type !== 'datetime' ? fromSerial(v, dated[c] === 'datetime') : v);
  const rows = body.map((r) => Object.fromEntries(fields.map((f, c) => [f.name, stored(f, asDate(f, c, r[c]))])));
  return { fields, rows };
}

/** The value at a dotted path in an answer (`data.items`); the answer itself for none. */
export function valueAt(answer, path) {
  const parts = String(path || '').split('.').map((s) => s.trim()).filter(Boolean);
  let v = answer;
  for (const p of parts) {
    if (v == null || typeof v !== 'object') return undefined;
    v = Array.isArray(v) && /^\d+$/.test(p) ? v[Number(p)] : v[p];
  }
  return v;
}

/** A JSON record flattened one level: { address: { city } } gives address.city; deeper values and lists are kept as their JSON. */
export function flatten(record) {
  if (record === null || typeof record !== 'object' || Array.isArray(record)) return { value: record };
  const out = {};
  for (const [k, v] of Object.entries(record)) {
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      for (const [k2, v2] of Object.entries(v)) out[`${k}.${k2}`] = v2 && typeof v2 === 'object' ? JSON.stringify(v2) : v2;
    } else {
      out[k] = Array.isArray(v) ? JSON.stringify(v) : v;
    }
  }
  return out;
}

/**
 * Records a web address answered, as a table: fields found from the first
 * `sample` records in the order they first appear, typed by their values.
 * Answers { fields, rows }.
 */
export function tableFromRecords(records, { sample = 200 } = {}) {
  const flat = records.map(flatten);
  const keys = [];
  const seen = new Set();
  for (const r of flat.slice(0, sample)) for (const k of Object.keys(r)) if (!seen.has(k)) { seen.add(k); keys.push(k); }
  if (!keys.length) throw new Error('The records have no fields to read.');
  const names = sourceFieldNames(keys);
  const fields = keys.map((k, i) => ({ name: names[i], ...inferType(flat.slice(0, sample).map((r) => (r[k] === undefined ? null : typeof r[k] === 'boolean' ? (r[k] ? 'yes' : 'no') : r[k]))) }));
  const rows = flat.map((r) => Object.fromEntries(fields.map((f, i) => [f.name, stored(f, r[keys[i]])])));
  return { fields, rows };
}
