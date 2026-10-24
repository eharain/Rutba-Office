// Data Files' reports: the records of a table or a saved query laid out to
// print, grouped by a field with totals for each group and for the whole.
// The report wizard makes a good one from a table or a query in one step;
// the report's data is worked out here and drawn by the window as a page,
// printed or saved as a PDF exactly as it is shown.
//
// A layout: { name, title, source: { table } | { query },
//   columns: [{ field, label, align: 'left' | 'right' }],
//   group: null | field, sort: null | field, desc: false,
//   totals: [{ field, fn: 'sum' | 'avg' | 'count' | 'min' | 'max' }] }
//
// A column a query shows may say its type ({ field, type }), for a total of
// an amount or a whole number the query works out.

import { labelFor } from './form-design.js';

const NUMERIC = new Set(['number', 'integer', 'currency']);

/**
 * A report for a table (its fields from `fields`) or a query (its columns
 * from `columns`, names or { name, type }), as the wizard makes it: each
 * amount (a Number or Currency) summed, the totals then chosen in the window.
 */
export function wizardReport({ table = null, query = null, fields = [], columns = [], relationships = [], name = null }) {
  const source = table ? { table } : { query };
  const looks = table
    ? fields.filter((f) => f.type !== 'autonumber').map((f) => ({ field: f.name, label: f.caption || labelFor(f.name), type: f.type }))
    : columns.map((c) => (typeof c === 'string' ? { field: c, label: c, type: null } : { field: c.name, label: c.name, type: c.type || null }));
  if (!looks.length) throw new Error('There is nothing to put in a report.');
  const pointing = table ? relationships.filter((r) => r.from.table === table).map((r) => r.from.field) : [];
  // Grouped by the first field that points at another table's record, as one would by customer.
  const group = looks.find((c) => pointing.includes(c.field))?.field || null;
  const cols = looks.filter((c) => c.field !== group).slice(0, 8).map((c) => ({ field: c.field, label: c.label, align: NUMERIC.has(c.type) && !pointing.includes(c.field) ? 'right' : 'left', ...(query && c.type ? { type: c.type } : {}) }));
  const totals = looks.filter((c) => (c.type === 'currency' || c.type === 'number') && c.field !== group).map((c) => ({ field: c.field, fn: 'sum' }));
  return { name: name || (table || query), title: labelFor(table || query), source, columns: cols, group, sort: null, desc: false, totals };
}

/** The totals a report's chooser offers: none, or one of these, for each column. */
export const TOTAL_FNS = ['sum', 'avg', 'count'];

/**
 * The columns of a report that may carry a total, for its Totals chooser:
 * each number shown (Number, Whole number or Currency, a query's columns
 * by their type), not the field it is grouped by nor one that points at
 * another table's record (shown as words). `fields` are reportData's.
 */
export function totalFields(layout, fields = []) {
  const typeOf = new Map(fields.map((f) => [f.name, f.type]));
  for (const c of layout.columns || []) if (c.type && !typeOf.get(c.field)) typeOf.set(c.field, c.type);
  return (layout.columns || []).filter((c) => c.field !== layout.group && NUMERIC.has(typeOf.get(c.field))).map((c) => ({ field: c.field, label: c.label || c.field, type: typeOf.get(c.field) }));
}

/**
 * A layout with `field`'s total set to `fn` (sum, avg, count, min or max),
 * or taken off with none (null or ''): one total for each column, in the
 * columns' order, the others as they were.
 */
export function withTotal(layout, field, fn) {
  if (fn && !['sum', 'avg', 'count', 'min', 'max'].includes(fn)) throw new Error(`A report's total is Sum, Average or Count, not ${fn}.`);
  const others = (layout.totals || []).filter((x) => x.field !== field);
  const totals = fn ? [...others, { field, fn }] : others;
  const at = new Map((layout.columns || []).map((c, i) => [c.field, i]));
  totals.sort((a, b) => (at.get(a.field) ?? 1e9) - (at.get(b.field) ?? 1e9));
  return { ...layout, totals };
}

/** A total of `values` by `fn`; null where there is nothing to count up. */
export function total(fn, values) {
  if (fn === 'count') return values.filter((v) => v !== null && v !== undefined && v !== '').length;
  // Empty values are left out, not counted as nought: an average of 10 and nothing is 10.
  const nums = values.filter((v) => v !== null && v !== undefined && v !== '').map(Number).filter(Number.isFinite);
  if (!nums.length) return null;
  // Fifteen significant digits, all a double holds for certain: 0.1 + 0.2 is 0.3 here.
  const clean = (v) => (Number.isInteger(v) ? v : Number(v.toPrecision(15)));
  if (fn === 'sum') return clean(nums.reduce((a, b) => a + b, 0));
  if (fn === 'avg') return clean(nums.reduce((a, b) => a + b, 0) / nums.length);
  if (fn === 'min') return Math.min(...nums);
  if (fn === 'max') return Math.max(...nums);
  return null;
}

const compare = (a, b) => (a === b ? 0 : a === null || a === undefined ? 1 : b === null || b === undefined ? -1 : typeof a === 'number' && typeof b === 'number' ? a - b : String(a).localeCompare(String(b), undefined, { numeric: true }));

/**
 * The report's data from its records: { groups: [{ key, label, rows, totals }],
 * totals, count }. `shownAs(field, value)` turns a pointing field's value into
 * the words it stands for (a customer's name for their ID).
 */
export function reportData(layout, records, shownAs = (f, v) => v) {
  const totalsOf = (rows) => Object.fromEntries((layout.totals || []).map((t) => [`${t.fn}:${t.field}`, total(t.fn, rows.map((r) => r[t.field]))]));
  const shownRow = (r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, shownAs(k, v)]));
  const sorted = [...records].sort((a, b) => {
    if (layout.group) {
      const g = compare(shownAs(layout.group, a[layout.group]), shownAs(layout.group, b[layout.group]));
      if (g) return g;
    }
    if (!layout.sort) return 0;
    // An empty value stays at the foot, largest first or smallest.
    const x = a[layout.sort];
    const y = b[layout.sort];
    const empty = (v) => v === null || v === undefined || v === '';
    if (empty(x) || empty(y)) return empty(x) - empty(y);
    return (layout.desc ? -1 : 1) * compare(x, y);
  });
  const groups = [];
  if (layout.group) {
    for (const r of sorted) {
      const key = r[layout.group];
      let g = groups[groups.length - 1];
      if (!g || g.key !== key) { g = { key, label: shownAs(layout.group, key), rows: [] }; groups.push(g); }
      g.rows.push(r);
    }
  } else {
    groups.push({ key: null, label: null, rows: sorted });
  }
  return {
    groups: groups.map((g) => ({ key: g.key, label: g.label, rows: g.rows.map(shownRow), totals: totalsOf(g.rows) })),
    totals: totalsOf(sorted),
    count: sorted.length,
  };
}
