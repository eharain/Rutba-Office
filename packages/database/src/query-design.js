// A query as a person builds it in Data Files' designer, and the SQL it
// stands for. The design names the tables put in, the fields shown, each
// field's criteria in the words Access uses (> 10, Like "A*", Between 1 And
// 5, Is Null, In ("a", "b"), a date as #2026-10-21#), the sort, and the
// totals; the tables are joined by the relationships the file keeps between
// them. Values a person types become parameters, never SQL text, and names
// are quoted, so a query written here cannot be made to do anything but
// read.
//
// A design: {
//   tables: ['Customers', 'Orders'],
//   columns: [{ table, field, show = true, sort: null | 'asc' | 'desc',
//               total: null | 'group' | 'sum' | 'avg' | 'count' | 'min' | 'max' | 'where',
//               criteria: ['> 10', ...] (one per Or row), as: null | 'Heading',
//               decimals: null | 0..6 (the answer rounded, an average to 2 places) }
//             | { expr: '[Quantity] * [Unit Price]', as: 'Line Total', ...the same }],
// }
// | { sql: 'SELECT ...' }, a query written in the SQL view and kept as it is written.
//
// A column with `expr` is calculated, as Access writes one (expression.js):
// its [Field] names are the query's tables' fields.

import { expressionSql } from './expression.js';

const q = (name) => `"${String(name).replace(/"/g, '""')}"`;
const col = (c) => `${q(c.table)}.${q(c.field)}`;

/**
 * "Line Total: [Quantity] * [Unit Price]" as Access's designer takes it in a
 * Field cell: { as, expr }. A cell with no name before a colon is the
 * expression alone. A colon inside brackets or quotes is part of the text.
 */
export function splitNamedExpression(cell) {
  const s = String(cell || '').trim();
  const m = /^([^[\]"'#:]+?)\s*:\s*(.+)$/s.exec(s);
  return m ? { as: m[1].trim(), expr: m[2].trim() } : { as: null, expr: s };
}

/** A column's keys the query designer sets, and the ones of its own it holds while the column is on screen. */
const DESIGNER_SETS = ['table', 'field', 'expr', 'show', 'sort', 'total', 'decimals', 'criteria'];
const DESIGNER_HOLDS = ['exprText'];

/**
 * The design the query designer saves: `columns` as it holds them (a field
 * chosen, or an expression typed as "Name: expression" in `exprText`), the
 * empty ones left out. What a column or the design has that the designer
 * does not set (a field column's heading, or a key a newer version wrote)
 * is kept as it was (A2.20), `kept` the design's own such keys.
 */
export function designerDesign(tables, columns, totals, kept = {}) {
  const used = (c) => Boolean((c.table && c.field) || (c.exprText !== null && c.exprText !== undefined && String(c.exprText).trim()));
  const top = Object.fromEntries(Object.entries(kept || {}).filter(([k]) => k !== 'tables' && k !== 'columns' && k !== 'sql'));
  return {
    ...top,
    tables,
    columns: columns.filter(used).map((c) => {
      const expression = c.exprText !== null && c.exprText !== undefined;
      const other = Object.fromEntries(Object.entries(c).filter(([k]) => !DESIGNER_SETS.includes(k) && !DESIGNER_HOLDS.includes(k) && !(expression && k === 'as')));
      return {
        ...other,
        ...(expression ? splitNamedExpression(c.exprText) : { table: c.table, field: c.field }),
        show: c.show,
        sort: c.sort || null,
        total: totals ? c.total || 'group' : null,
        // Rounded where the person said so: an average spend to 2 places.
        ...(c.decimals !== '' && c.decimals !== null && c.decimals !== undefined ? { decimals: Number(c.decimals) } : {}),
        criteria: (c.criteria || []).filter((x, i, all) => String(x).trim() || all.slice(i + 1).some((y) => String(y).trim())),
      };
    }),
  };
}

/** The aggregate each total names, in SQL. */
const AGGREGATES = { sum: 'SUM', avg: 'AVG', count: 'COUNT', min: 'MIN', max: 'MAX' };
/** A total's heading, as the designer's Total list names it: "Average of Total", "Sum of Quantity". */
const TOTAL_WORDS = { sum: 'Sum', avg: 'Average', count: 'Count', min: 'Min', max: 'Max' };
/** A heading as 2.17 wrote an average's ("Avg of Total"), as it is written now, for a report that still names it so. */
export const newHeading = (heading) => String(heading).replace(/^Avg of /, 'Average of ');

/** A typed value as SQL takes it: a quoted string, a #date#, a number, true/false. */
function valueOf(token) {
  const t = token.trim();
  if (/^"(.*)"$/s.test(t)) return t.slice(1, -1).replace(/""/g, '"');
  if (/^'(.*)'$/s.test(t)) return t.slice(1, -1).replace(/''/g, "'");
  const date = /^#(\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?)#$/.exec(t);
  if (date) return date[1].replace(' ', 'T');
  if (/^(true|yes)$/i.test(t)) return 1;
  if (/^(false|no)$/i.test(t)) return 0;
  if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t);
  return t;
}

/** The values of an In list: "a", 'b', 3. */
function listOf(text) {
  const out = [];
  const re = /\s*("(?:[^"]|"")*"|'(?:[^']|'')*'|[^,]+)\s*(?:,|$)/g;
  let m;
  while ((m = re.exec(text)) && m[0]) out.push(valueOf(m[1]));
  return out;
}

/** An Access wildcard (* and ?) as a LIKE pattern (% and _), the pattern's own % and _ kept literal. */
const likePattern = (s) => String(s).replace(/[\\%_]/g, (c) => `\\${c}`).replace(/\*/g, '%').replace(/\?/g, '_');

/** Whether a criterion's value names a field ([Reorder Level], [Products].[Price] * 2) rather than being one: brackets outside quotes. */
const namesField = (token) => {
  const t = token.trim();
  return !/^["'#]/.test(t) && /\[[^\]]+\]/.test(t);
};

/**
 * One cell of criteria for one field as SQL and its parameters, or throws
 * what is wrong with it. `c` is the column, `text` what the person typed.
 * A value that names a field (<= [Reorder Level]) is compared with that
 * field, worked out as an expression is: `resolve(table, field)` answers its
 * column, as designToSql's does.
 */
export function criterionSql(column, text, { resolve = null, today = null } = {}) {
  const t = String(text).trim();
  if (!t) return null;
  let m;
  // A field named where a value goes: that field's value, not the words "[Reorder Level]".
  const side = (token) => {
    if (!namesField(token)) return { sql: '?', params: [valueOf(token)] };
    if (!resolve) throw new Error(`"${t}" names a field. Compare fields in the query designer, where the query's tables are known.`);
    return { sql: `(${expressionSql(token, resolve, { today })})`, params: [] };
  };
  if ((m = /^not\s+(.+)$/i.exec(t)) && !/^not\s+(like|in|between)\b/i.test(t) && !/^is\b/i.test(t)) {
    const inner = criterionSql(column, m[1], { resolve, today });
    return { sql: `NOT (${inner.sql})`, params: inner.params };
  }
  if (/^is\s+null$/i.test(t)) return { sql: `${column} IS NULL`, params: [] };
  if (/^is\s+not\s+null$/i.test(t)) return { sql: `${column} IS NOT NULL`, params: [] };
  if ((m = /^(not\s+)?like\s+(.+)$/i.exec(t))) return { sql: `${column} ${m[1] ? 'NOT ' : ''}LIKE ? ESCAPE '\\'`, params: [likePattern(valueOf(m[2]))] };
  if ((m = /^(not\s+)?in\s*\((.*)\)$/is.exec(t))) {
    const values = listOf(m[2]);
    if (!values.length) throw new Error(`"${t}" lists nothing to look for.`);
    return { sql: `${column} ${m[1] ? 'NOT ' : ''}IN (${values.map(() => '?').join(', ')})`, params: values };
  }
  if ((m = /^(not\s+)?between\s+(.+?)\s+and\s+(.+)$/i.exec(t))) {
    const [lo, hi] = [side(m[2]), side(m[3])];
    return { sql: `${column} ${m[1] ? 'NOT ' : ''}BETWEEN ${lo.sql} AND ${hi.sql}`, params: [...lo.params, ...hi.params] };
  }
  if ((m = /^(<>|<=|>=|=|<|>)\s*(.+)$/.exec(t))) {
    const v = side(m[2]);
    return { sql: `${column} ${m[1]} ${v.sql}`, params: v.params };
  }
  // A value alone means equal to it; with a wildcard in it, like it. A field named alone, equal to that field.
  if (namesField(t)) { const f = side(t); return { sql: `${column} = ${f.sql}`, params: f.params }; }
  const v = valueOf(t);
  if (typeof v === 'string' && /[*?]/.test(v)) return { sql: `${column} LIKE ? ESCAPE '\\'`, params: [likePattern(v)] };
  return { sql: `${column} = ?`, params: [v] };
}

/** A query's own SQL, checked as the SQL view runs it: SELECT (or WITH ... SELECT), one statement. */
export function sqlText(sql) {
  const text = String(sql ?? '').trim().replace(/;\s*$/, '');
  if (!text) throw new Error('Write the query first.');
  if (!/^(select|with)\b/i.test(text)) throw new Error('A query here starts with SELECT.');
  if (/;/.test(text.replace(/'(?:[^']|'')*'|"(?:[^"]|"")*"/g, ''))) throw new Error('One statement at a time.');
  return text;
}

/**
 * The joins between the design's tables, from the relationships: each table
 * after the first joined to one already in by a relationship between them.
 * A table no relationship reaches is refused, rather than crossed with the
 * rest (every row of one against every row of the other).
 */
export function joinsFor(tables, relationships) {
  if (!tables.length) throw new Error('Put a table in the query first.');
  const inQuery = [tables[0]];
  const joins = [];
  const left = tables.slice(1);
  while (left.length) {
    const i = left.findIndex((t) => relationships.some((r) => (r.from.table === t && inQuery.includes(r.to.table)) || (r.to.table === t && inQuery.includes(r.from.table))));
    if (i < 0) throw new Error(`${left[0]} is not related to the other tables in the query. Relate them in Relationships first.`);
    const t = left.splice(i, 1)[0];
    const r = relationships.find((x) => (x.from.table === t && inQuery.includes(x.to.table)) || (x.to.table === t && inQuery.includes(x.from.table)));
    joins.push(`JOIN ${q(t)} ON ${q(r.from.table)}.${q(r.from.field)} = ${q(r.to.table)}.${q(r.to.field)}`);
    inQuery.push(t);
  }
  return joins;
}

/**
 * The SQL a design stands for: { sql, params, columns } with `columns` the
 * headings of what it shows. `relationships` are the file's.
 */
export function designToSql(design, relationships = [], { fieldsOf = null, today = null } = {}) {
  // A query written in SQL and kept from the SQL view: its own text, a SELECT and nothing else.
  if (typeof design?.sql === 'string') return { sql: sqlText(design.sql), params: [], columns: [], sources: [], written: true };
  const tables = [...new Set(design.tables || [])];
  const columns = (design.columns || []).filter((c) => c && ((c.table && c.field) || (typeof c.expr === 'string' && c.expr.trim())));
  for (const c of columns) if (!c.expr && !tables.includes(c.table)) throw new Error(`${c.field} is from ${c.table}, which is not in the query.`);
  // A [Field] in an expression: the one table of the query that has it.
  const resolve = (table, field) => {
    if (table) {
      if (!tables.includes(table)) throw new Error(`[${table}].[${field}]: ${table} is not in the query.`);
      if (fieldsOf && !fieldsOf(table).includes(field)) throw new Error(`${table} has no field called ${field}.`);
      return `${q(table)}.${q(field)}`;
    }
    if (!fieldsOf) {
      if (tables.length === 1) return `${q(tables[0])}.${q(field)}`;
      throw new Error(`Say which table [${field}] is from: [Table].[${field}].`);
    }
    const having = tables.filter((t) => fieldsOf(t).includes(field));
    if (!having.length) throw new Error(`No table in the query has a field called ${field}.`);
    if (having.length > 1) throw new Error(`${having.join(' and ')} both have ${field}: write [${having[0]}].[${field}] to say which.`);
    return `${q(having[0])}.${q(field)}`;
  };
  let n = 0;
  const named = new Map(columns.filter((c) => c.expr).map((c) => [c, c.as || `Expr${++n}`]));
  const colSql = (c) => (c.expr ? `(${expressionSql(c.expr, resolve, { today })})` : col(c));
  const nameOf = (c) => (c.expr ? named.get(c) : c.field);
  const totals = columns.some((c) => c.total);
  const shown = columns.filter((c) => c.show !== false && c.total !== 'where');
  if (!shown.length) throw new Error('Choose a field to show.');
  const params = [];
  const select = shown.map((c) => {
    const agg = totals ? AGGREGATES[c.total] : null;
    const raw = agg ? `${agg}(${colSql(c)})` : colSql(c);
    // Rounded where the design says so: an average spend of 15.08, not 15.0754455445545.
    const places = Number.isInteger(c.decimals) ? Math.max(0, Math.min(6, c.decimals)) : null;
    const expr = places === null ? raw : `ROUND(${raw}, ${places})`;
    const heading = (c.expr && !agg ? named.get(c) : c.as) || (agg ? `${TOTAL_WORDS[c.total]} of ${nameOf(c)}` : nameOf(c));
    return { sql: `${expr} AS ${q(heading)}`, heading };
  });
  // Criteria: a row of cells joined by AND, the rows by OR. With totals, a
  // criterion on a summed field is a HAVING; on a grouped or Where field, a WHERE.
  const rows = Math.max(0, ...columns.map((c) => (c.criteria || []).length));
  const where = [];
  const having = [];
  const havingParams = [];
  for (let r = 0; r < rows; r++) {
    const w = [];
    const h = [];
    const wp = [];
    const hp = [];
    for (const c of columns) {
      const one = criterionSql(totals && AGGREGATES[c.total] ? `${AGGREGATES[c.total]}(${colSql(c)})` : colSql(c), (c.criteria || [])[r] || '', { resolve, today });
      if (!one) continue;
      if (totals && AGGREGATES[c.total]) { h.push(one.sql); hp.push(...one.params); } else { w.push(one.sql); wp.push(...one.params); }
    }
    if (w.length) { where.push(`(${w.join(' AND ')})`); params.push(...wp); }
    if (h.length) { having.push(`(${h.join(' AND ')})`); havingParams.push(...hp); }
  }
  const group = totals ? columns.filter((c) => c.total === 'group').map(colSql) : [];
  const order = columns.filter((c) => c.sort === 'asc' || c.sort === 'desc').map((c) => {
    const s = shown.find((x) => x === c);
    const target = s && totals && AGGREGATES[c.total] ? q(select[shown.indexOf(s)].heading) : colSql(c);
    return `${target} ${c.sort === 'desc' ? 'DESC' : 'ASC'}`;
  });
  const sql = [
    `SELECT ${select.map((s) => s.sql).join(', ')}`,
    `FROM ${q(tables[0])}`,
    ...joinsFor(tables, relationships),
    where.length ? `WHERE ${where.join(' OR ')}` : '',
    group.length ? `GROUP BY ${group.join(', ')}` : '',
    having.length ? `HAVING ${having.join(' OR ')}` : '',
    order.length ? `ORDER BY ${order.join(', ')}` : '',
  ].filter(Boolean).join('\n');
  // Each column's source, for its type: the field it shows, or null for a calculation or a total.
  const sources = shown.map((c) => (!c.expr && !(totals && AGGREGATES[c.total]) ? { table: c.table, field: c.field } : null));
  return { sql, params: [...params, ...havingParams], columns: select.map((s) => s.heading), sources };
}
