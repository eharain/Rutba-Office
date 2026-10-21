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
//               criteria: ['> 10', ...] (one per Or row), as: null | 'Heading' }],
// }

const q = (name) => `"${String(name).replace(/"/g, '""')}"`;
const col = (c) => `${q(c.table)}.${q(c.field)}`;

/** The aggregate each total names, in SQL. */
const AGGREGATES = { sum: 'SUM', avg: 'AVG', count: 'COUNT', min: 'MIN', max: 'MAX' };

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

/**
 * One cell of criteria for one field as SQL and its parameters, or throws
 * what is wrong with it. `c` is the column, `text` what the person typed.
 */
export function criterionSql(column, text) {
  const t = String(text).trim();
  if (!t) return null;
  let m;
  if ((m = /^not\s+(.+)$/i.exec(t)) && !/^not\s+(like|in|between)\b/i.test(t) && !/^is\b/i.test(t)) {
    const inner = criterionSql(column, m[1]);
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
  if ((m = /^(not\s+)?between\s+(.+?)\s+and\s+(.+)$/i.exec(t))) return { sql: `${column} ${m[1] ? 'NOT ' : ''}BETWEEN ? AND ?`, params: [valueOf(m[2]), valueOf(m[3])] };
  if ((m = /^(<>|<=|>=|=|<|>)\s*(.+)$/.exec(t))) return { sql: `${column} ${m[1] === '<>' ? '<>' : m[1]} ?`, params: [valueOf(m[2])] };
  // A value alone means equal to it; with a wildcard in it, like it.
  const v = valueOf(t);
  if (typeof v === 'string' && /[*?]/.test(v)) return { sql: `${column} LIKE ? ESCAPE '\\'`, params: [likePattern(v)] };
  return { sql: `${column} = ?`, params: [v] };
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
export function designToSql(design, relationships = []) {
  const tables = [...new Set(design.tables || [])];
  const columns = (design.columns || []).filter((c) => c && c.table && c.field);
  for (const c of columns) if (!tables.includes(c.table)) throw new Error(`${c.field} is from ${c.table}, which is not in the query.`);
  const totals = columns.some((c) => c.total);
  const shown = columns.filter((c) => c.show !== false && c.total !== 'where');
  if (!shown.length) throw new Error('Choose a field to show.');
  const params = [];
  const select = shown.map((c) => {
    const agg = totals ? AGGREGATES[c.total] : null;
    const expr = agg ? `${agg}(${col(c)})` : col(c);
    const heading = c.as || (agg ? `${c.total[0].toUpperCase()}${c.total.slice(1)} of ${c.field}` : c.field);
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
      const one = criterionSql(totals && AGGREGATES[c.total] ? `${AGGREGATES[c.total]}(${col(c)})` : col(c), (c.criteria || [])[r] || '');
      if (!one) continue;
      if (totals && AGGREGATES[c.total]) { h.push(one.sql); hp.push(...one.params); } else { w.push(one.sql); wp.push(...one.params); }
    }
    if (w.length) { where.push(`(${w.join(' AND ')})`); params.push(...wp); }
    if (h.length) { having.push(`(${h.join(' AND ')})`); havingParams.push(...hp); }
  }
  const group = totals ? columns.filter((c) => c.total === 'group').map(col) : [];
  const order = columns.filter((c) => c.sort === 'asc' || c.sort === 'desc').map((c) => {
    const s = shown.find((x) => x === c);
    const target = s && totals && AGGREGATES[c.total] ? q(select[shown.indexOf(s)].heading) : col(c);
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
  return { sql, params: [...params, ...havingParams], columns: select.map((s) => s.heading) };
}
