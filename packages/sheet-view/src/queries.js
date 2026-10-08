/**
 * Data → Get & Transform: queries, as Power Query keeps them — a source (a
 * table or a range of this workbook, or a CSV file) and a list of applied
 * steps, each one a change to the table the step before it left: columns
 * removed, kept, renamed or given a type; rows filtered, sorted, cut to the
 * top so many, their duplicates or blanks taken out; values replaced, split
 * at a delimiter, trimmed or put in capitals; rows grouped with a total,
 * count, average, least or most; an index added.
 *
 * Running a query reads its source afresh and applies the steps in order,
 * so Refresh brings in what the source holds now. A step is data, not code:
 * nothing a query does runs a formula or a script.
 *
 * A table here is { columns: [name], rows: [[value]] }, values numbers,
 * strings, booleans, dates as serials, null for blank.
 */

export const STEP_KINDS = [
  'removeColumns', 'keepColumns', 'renameColumn', 'changeType', 'filterRows', 'sort', 'removeDuplicates',
  'removeBlankRows', 'keepTopRows', 'replaceValues', 'splitColumn', 'groupBy', 'addIndex', 'transformText', 'promoteHeaders',
  'appendQuery', 'mergeQueries',
];

/** Merge Queries' kinds of join, as Power Query names them. */
export const JOIN_KINDS = [['left', 'Left Outer (all from the first, matching from the second)'], ['inner', 'Inner (only matching rows)'], ['leftAnti', 'Left Anti (rows only in the first)'], ['full', 'Full Outer (all rows from both)']];

const blank = (v) => v === null || v === undefined || v === '';
const indexOf = (t, name) => {
  const i = t.columns.indexOf(name);
  if (i < 0) throw new Error(`There is no column "${name}"`);
  return i;
};
const asNumber = (v) => {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = Number(String(v ?? '').replace(/[,\s£$€]/g, ''));
  return String(v ?? '').trim() !== '' && Number.isFinite(n) ? n : null;
};
const asText = (v) => (blank(v) ? '' : String(v));
const compare = (a, b) => {
  if (blank(a) && blank(b)) return 0;
  if (blank(a)) return 1;
  if (blank(b)) return -1;
  const x = asNumber(a);
  const y = asNumber(b);
  if (x !== null && y !== null && typeof a !== 'string' && typeof b !== 'string') return x - y;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  return asText(a).localeCompare(asText(b), undefined, { numeric: true, sensitivity: 'base' });
};

/** Make column names unique and non-empty, as Power Query names them: Column1, Name.1 … */
function uniqueNames(names) {
  const seen = new Map();
  return names.map((n, i) => {
    let name = blank(n) ? `Column${i + 1}` : String(n).trim() || `Column${i + 1}`;
    if (seen.has(name)) { const k = seen.get(name) + 1; seen.set(name, k); name = `${name}.${k}`; } else seen.set(name, 0);
    return name;
  });
}

/**
 * One step applied to a table: a new table. Append and Merge read another
 * table with `ctx.table(source)` — a table or range of the workbook, a file,
 * or another query.
 */
export function applyStep(t, step, ctx = {}) {
  const s = step || {};
  switch (s.kind) {
    // Append Queries: the other table's rows after these, its columns matched by name, any new ones added.
    case 'appendQuery': {
      if (typeof ctx.table !== 'function') throw new Error('Append needs the other table');
      // Headings made unique first: two columns of one name would both be
      // written into the first, and the second's values lost.
      const raw = ctx.table(s.with);
      const other = { ...raw, columns: uniqueNames(raw.columns) };
      const columns = [...t.columns, ...other.columns.filter((c) => !t.columns.includes(c))];
      const at = other.columns.map((c) => columns.indexOf(c));
      const rows = t.rows.map((r) => columns.map((_, i) => (i < r.length ? r[i] : null)));
      for (const r of other.rows) {
        const row = columns.map(() => null);
        at.forEach((to, from) => { row[to] = r[from] === undefined ? null : r[from]; });
        rows.push(row);
      }
      return { columns, rows };
    }
    // Merge Queries: each row joined to the other table's rows whose key matches, their other columns brought in.
    case 'mergeQueries': {
      if (typeof ctx.table !== 'function') throw new Error('Merge needs the other table');
      const other = ctx.table(s.with);
      const mine = indexOf(t, s.on);
      const theirs = other.columns.indexOf(s.withOn);
      if (theirs < 0) throw new Error(`The other table has no column "${s.withOn}"`);
      const how = s.how || 'left';
      const key = (v) => (blank(v) ? null : asText(v).trim().toLowerCase());
      const index = new Map();
      other.rows.forEach((r, i) => { const k = key(r[theirs]); if (k === null) return; if (!index.has(k)) index.set(k, []); index.get(k).push(i); });
      if (how === 'leftAnti') return { columns: t.columns.slice(), rows: t.rows.filter((r) => !index.has(key(r[mine]))) };
      // The other table's columns but its key, named apart from these where they clash.
      const brought = other.columns.map((c, i) => [c, i]).filter(([, i]) => i !== theirs);
      const prefix = s.prefix || 'Merged';
      const names = brought.map(([c]) => (t.columns.includes(c) ? `${prefix}.${c}` : c));
      const columns = uniqueNames([...t.columns, ...names]);
      const rows = [];
      const used = new Set();
      for (const r of t.rows) {
        const hits = index.get(key(r[mine])) || [];
        if (!hits.length) { if (how !== 'inner') rows.push([...r, ...brought.map(() => null)]); continue; }
        for (const h of hits) { used.add(h); rows.push([...r, ...brought.map(([, i]) => (other.rows[h][i] === undefined ? null : other.rows[h][i]))]); }
      }
      // Full outer: the other table's rows nothing matched, its key in this key's column.
      if (how === 'full') {
        other.rows.forEach((r, h) => {
          if (used.has(h)) return;
          const row = t.columns.map(() => null);
          row[mine] = r[theirs] === undefined ? null : r[theirs];
          rows.push([...row, ...brought.map(([, i]) => (r[i] === undefined ? null : r[i]))]);
        });
      }
      return { columns, rows };
    }
    case 'removeColumns': {
      const drop = new Set((s.columns || []).map((c) => indexOf(t, c)));
      return { columns: t.columns.filter((_, i) => !drop.has(i)), rows: t.rows.map((r) => r.filter((_, i) => !drop.has(i))) };
    }
    case 'keepColumns': {
      const keep = (s.columns || []).map((c) => indexOf(t, c));
      return { columns: keep.map((i) => t.columns[i]), rows: t.rows.map((r) => keep.map((i) => r[i] ?? null)) };
    }
    case 'renameColumn': {
      const i = indexOf(t, s.from);
      const to = String(s.to || '').trim();
      if (!to) throw new Error('A column needs a name');
      if (t.columns.some((c, j) => j !== i && c === to)) throw new Error(`There is already a column "${to}"`);
      return { columns: t.columns.map((c, j) => (j === i ? to : c)), rows: t.rows };
    }
    case 'changeType': {
      const i = indexOf(t, s.column);
      const to = (v) => {
        if (blank(v)) return null;
        switch (s.type) {
          case 'text': return asText(v);
          case 'number': return asNumber(v);
          case 'integer': { const n = asNumber(v); return n === null ? null : Math.round(n); }
          case 'boolean': return typeof v === 'boolean' ? v : /^(true|yes|1)$/i.test(String(v).trim()) ? true : /^(false|no|0)$/i.test(String(v).trim()) ? false : null;
          case 'date': {
            if (typeof v === 'number') return Math.floor(v);
            const d = new Date(String(v));
            return Number.isNaN(d.getTime()) ? null : Math.floor((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
          }
          default: throw new Error(`"${s.type}" is not a type a column can be given`);
        }
      };
      return { columns: t.columns, rows: t.rows.map((r) => r.map((v, j) => (j === i ? to(v) : v))), types: { ...(t.types || {}), [s.column]: s.type } };
    }
    case 'filterRows': {
      const i = indexOf(t, s.column);
      const want = s.value;
      const lower = (v) => asText(v).toLowerCase();
      const test = {
        equals: (v) => compare(v, want) === 0,
        notEquals: (v) => compare(v, want) !== 0,
        contains: (v) => lower(v).includes(lower(want)),
        notContains: (v) => !lower(v).includes(lower(want)),
        beginsWith: (v) => lower(v).startsWith(lower(want)),
        endsWith: (v) => lower(v).endsWith(lower(want)),
        greater: (v) => !blank(v) && compare(v, want) > 0,
        greaterOrEqual: (v) => !blank(v) && compare(v, want) >= 0,
        less: (v) => !blank(v) && compare(v, want) < 0,
        lessOrEqual: (v) => !blank(v) && compare(v, want) <= 0,
        blank: (v) => blank(v),
        notBlank: (v) => !blank(v),
      }[s.op];
      if (!test) throw new Error(`"${s.op}" is not a filter`);
      return { ...t, rows: t.rows.filter((r) => test(r[i])) };
    }
    case 'sort': {
      const keys = (s.by || [{ column: s.column, descending: s.descending }]).map((k) => ({ i: indexOf(t, k.column), sign: k.descending ? -1 : 1 }));
      return { ...t, rows: t.rows.map((r, n) => [r, n]).sort(([a, n], [b, m]) => { for (const k of keys) { const c = compare(a[k.i], b[k.i]); if (c) return blank(a[k.i]) || blank(b[k.i]) ? c : c * k.sign; } return n - m; }).map(([r]) => r) };
    }
    case 'removeDuplicates': {
      const cols = (s.columns?.length ? s.columns : t.columns).map((c) => indexOf(t, c));
      const seen = new Set();
      return { ...t, rows: t.rows.filter((r) => { const k = JSON.stringify(cols.map((i) => asText(r[i]).toLowerCase())); if (seen.has(k)) return false; seen.add(k); return true; }) };
    }
    case 'removeBlankRows':
      return { ...t, rows: t.rows.filter((r) => r.some((v) => !blank(v))) };
    case 'keepTopRows': {
      const n = Math.max(0, Math.floor(Number(s.count) || 0));
      return { ...t, rows: t.rows.slice(0, n) };
    }
    case 'replaceValues': {
      const i = indexOf(t, s.column);
      const find = asText(s.find);
      return { ...t, rows: t.rows.map((r) => r.map((v, j) => {
        if (j !== i) return v;
        if (s.whole !== false && compare(v, s.find) === 0) return blank(s.replace) ? null : s.replace;
        if (s.whole === false && find && typeof v === 'string') return v.split(find).join(asText(s.replace));
        return v;
      })) };
    }
    case 'splitColumn': {
      const i = indexOf(t, s.column);
      const d = s.delimiter === 'tab' ? '\t' : s.delimiter === 'space' ? ' ' : s.delimiter === 'comma' ? ',' : s.delimiter === 'semicolon' ? ';' : String(s.delimiter ?? ',');
      if (!d) throw new Error('A column is split at something');
      const parts = t.rows.map((r) => asText(r[i]).split(d).map((p) => p.trim()));
      const width = Math.max(1, ...parts.map((p) => p.length));
      const names = uniqueNames([...t.columns.slice(0, i), ...Array.from({ length: width }, (_, k) => `${t.columns[i]}.${k + 1}`), ...t.columns.slice(i + 1)]);
      // A part that is a number becomes one, but not one whose leading zeros
      // would go: "A-007" splits into A and 007, a code, not 7.
      const part = (p) => (p === undefined || p === '' ? null : /^[-+]?0\d/.test(p) ? p : asNumber(p) ?? p);
      return { columns: names, rows: t.rows.map((r, n) => [...r.slice(0, i), ...Array.from({ length: width }, (_, k) => part(parts[n][k])), ...r.slice(i + 1)]) };
    }
    case 'groupBy': {
      const keys = (s.columns || [s.column]).filter(Boolean).map((c) => indexOf(t, c));
      const aggs = (s.aggregations || []).map((a) => ({ ...a, i: a.fn === 'count' && !a.column ? -1 : indexOf(t, a.column) }));
      if (!aggs.length) aggs.push({ fn: 'count', name: 'Count', i: -1 });
      const groups = new Map();
      for (const r of t.rows) {
        const k = JSON.stringify(keys.map((i) => asText(r[i]).toLowerCase()));
        if (!groups.has(k)) groups.set(k, { key: keys.map((i) => r[i]), rows: [] });
        groups.get(k).rows.push(r);
      }
      const fold = (a, rows) => {
        if (a.fn === 'count') return a.i < 0 ? rows.length : rows.filter((r) => !blank(r[a.i])).length;
        const nums = rows.map((r) => asNumber(r[a.i])).filter((n) => n !== null);
        if (a.fn === 'sum') return nums.reduce((x, y) => x + y, 0);
        if (a.fn === 'average') return nums.length ? nums.reduce((x, y) => x + y, 0) / nums.length : null;
        if (a.fn === 'min') return nums.length ? Math.min(...nums) : null;
        if (a.fn === 'max') return nums.length ? Math.max(...nums) : null;
        throw new Error(`"${a.fn}" is not a way rows are added up`);
      };
      return {
        columns: uniqueNames([...keys.map((i) => t.columns[i]), ...aggs.map((a) => a.name || `${a.fn[0].toUpperCase()}${a.fn.slice(1)}${a.i >= 0 ? ` of ${t.columns[a.i]}` : ''}`)]),
        rows: [...groups.values()].map((g) => [...g.key, ...aggs.map((a) => fold(a, g.rows))]),
      };
    }
    case 'addIndex': {
      const start = Number.isFinite(Number(s.start)) ? Number(s.start) : 1;
      const name = uniqueNames([...t.columns, s.name || 'Index']).pop();
      return { columns: [...t.columns, name], rows: t.rows.map((r, n) => [...r, start + n]) };
    }
    case 'transformText': {
      const i = indexOf(t, s.column);
      const fn = { trim: (x) => x.trim().replace(/\s+/g, ' '), upper: (x) => x.toUpperCase(), lower: (x) => x.toLowerCase(), proper: (x) => x.toLowerCase().replace(/(^|[\s'-])(\p{L})/gu, (m, a, b) => a + b.toUpperCase()) }[s.how];
      if (!fn) throw new Error(`"${s.how}" is not a change to text`);
      return { ...t, rows: t.rows.map((r) => r.map((v, j) => (j === i && typeof v === 'string' ? fn(v) : v))) };
    }
    case 'promoteHeaders': {
      if (!t.rows.length) return t;
      return { columns: uniqueNames(t.rows[0].map((v) => (blank(v) ? '' : String(v)))), rows: t.rows.slice(1) };
    }
    default:
      throw new Error(`"${s.kind}" is not a step a query takes`);
  }
}

/** A query run on its source table: the table after each step, or after `upTo` of them. */
export function runSteps(source, steps = [], { upTo = null, table = null } = {}) {
  let t = { columns: uniqueNames(source.columns), rows: source.rows.map((r) => source.columns.map((_, i) => (r[i] === undefined ? null : r[i]))) };
  const last = upTo === null ? steps.length : Math.min(upTo, steps.length);
  for (let i = 0; i < last; i++) t = applyStep(t, steps[i], { table });
  return t;
}

/** Another table a step reads, in words. */
const sourceWords = (src) => (!src ? 'a table' : src.kind === 'query' ? `query ${src.name || src.id}` : src.kind === 'table' ? `table ${src.table}` : src.kind === 'range' ? `${src.sheet}!${src.ref}` : src.kind === 'csv' ? String(src.path || '').split(/[\\/]/).pop() : src.kind);

/** What a step says it does, for the Applied Steps list. */
export function describeStep(s) {
  const cols = (list) => (list || []).map((c) => `"${c}"`).join(', ');
  switch (s.kind) {
    case 'removeColumns': return `Removed ${cols(s.columns)}`;
    case 'keepColumns': return `Kept ${cols(s.columns)}`;
    case 'renameColumn': return `Renamed "${s.from}" to "${s.to}"`;
    case 'changeType': return `"${s.column}" made ${s.type === 'integer' ? 'whole numbers' : s.type === 'number' ? 'numbers' : s.type === 'boolean' ? 'true or false' : s.type === 'date' ? 'dates' : 'text'}`;
    case 'filterRows': return `Rows where "${s.column}" ${({ equals: 'is', notEquals: 'is not', contains: 'contains', notContains: 'does not contain', beginsWith: 'begins with', endsWith: 'ends with', greater: 'is more than', greaterOrEqual: 'is at least', less: 'is less than', lessOrEqual: 'is at most', blank: 'is blank', notBlank: 'is not blank' })[s.op] || s.op}${s.op === 'blank' || s.op === 'notBlank' ? '' : ` ${JSON.stringify(s.value)}`}`;
    case 'sort': return `Sorted by "${(s.by?.[0] || s).column}"${(s.by?.[0] || s).descending ? ', Z to A' : ''}`;
    case 'removeDuplicates': return s.columns?.length ? `Duplicates of ${cols(s.columns)} removed` : 'Duplicate rows removed';
    case 'removeBlankRows': return 'Blank rows removed';
    case 'keepTopRows': return `The first ${s.count} rows kept`;
    case 'replaceValues': return `In "${s.column}", ${JSON.stringify(s.find)} replaced with ${JSON.stringify(s.replace)}`;
    case 'splitColumn': return `"${s.column}" split at ${s.delimiter === 'tab' || s.delimiter === 'space' || s.delimiter === 'comma' || s.delimiter === 'semicolon' ? `each ${s.delimiter}` : JSON.stringify(s.delimiter)}`;
    case 'groupBy': return `Grouped by ${cols(s.columns || [s.column])}`;
    case 'addIndex': return `Index column "${s.name || 'Index'}" added`;
    case 'transformText': return `"${s.column}" ${({ trim: 'trimmed', upper: 'in capitals', lower: 'in small letters', proper: 'in title case' })[s.how]}`;
    case 'promoteHeaders': return 'First row used as headers';
    case 'appendQuery': return `Appended ${sourceWords(s.with)}`;
    case 'mergeQueries': return `Merged with ${sourceWords(s.with)} on "${s.on}" = "${s.withOn}"${s.how && s.how !== 'left' ? ` (${({ inner: 'inner', leftAnti: 'left anti', full: 'full outer' })[s.how] || s.how})` : ''}`;
    default: return s.kind;
  }
}

/**
 * Rows of a CSV or TSV text, quotes and doubled quotes honoured, each value a
 * number where it reads as one — but a quoted one only where nothing is lost:
 * "1" is 1, and "007", a code the quotes were there to keep, stays 007.
 */
export function parseDelimited(text, delimiter = null) {
  const src = String(text || '').replace(/^﻿/, '');
  const first = src.split(/\r?\n/, 1)[0] || '';
  const d = delimiter || (first.split('\t').length > first.split(',').length ? '\t' : first.split(';').length > first.split(',').length ? ';' : ',');
  const rows = [];
  let row = [];
  let cell = '';
  let quoted = false;
  let wasQuoted = false;
  // A quoted cell is marked, for `value` below.
  const end = () => { row.push(wasQuoted ? { text: cell } : cell); cell = ''; wasQuoted = false; };
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"' && src[i + 1] === '"') { cell += '"'; i += 1; } else if (ch === '"') quoted = false; else cell += ch;
      continue;
    }
    if (ch === '"' && cell === '') { quoted = true; wasQuoted = true; continue; }
    if (ch === d) { end(); continue; }
    if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && src[i + 1] === '\n') i += 1;
      end(); rows.push(row); row = [];
      continue;
    }
    cell += ch;
  }
  if (cell !== '' || row.length || wasQuoted) { end(); rows.push(row); }
  const value = (v) => {
    // Quoted, it is a number only if nothing is lost: "1" is 1, "007" stays 007.
    if (typeof v === 'object') {
      if (v.text.trim() === '') return null;
      if (/^[-+]?0\d/.test(v.text.trim())) return v.text;
      v = v.text;
    }
    const t = v.trim();
    if (t === '') return null;
    const n = Number(t);
    return /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(t) && Number.isFinite(n) ? n : v;
  };
  return rows.map((r) => r.map(value));
}
