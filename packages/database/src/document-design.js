// Data Files' print templates ("documents"): an invoice, a sale receipt, a
// rental invoice, a payslip, a statement or an agreement, one for each record
// of a table, laid out in blocks a person arranges:
//
//   heading     a line of large type: "Invoice"
//   fields      the record's fields, by label, in one to three columns; a name
//               like Customer.Address follows the field that points at another
//               table's record and takes that record's field
//   lines       the related records (an invoice's lines) as a table, with
//               calculated columns written as a query's are
//   totals      rows worked out from the lines' sums, the record's fields, the
//               letterhead's and the rows above: Subtotal, VAT, Total
//   text        words with {Field} blanks filled from the record
//   signatures  lines to sign on, with who signs under each
//   spacer      room between blocks
//
// The letterhead is the first record of a table of the business's own
// details, put at the top of every page. Paper is A4, Letter, A5, or a till
// roll 80 or 58 mm wide. The expressions are the query designer's
// (expression.js), worked out by SQLite with every name given as a literal
// value, so a template can compute but never reach anything else.

import { expressionSql } from './expression.js';
import { shownField } from './form-design.js';

export const PAPERS = ['a4', 'letter', 'a5', 'receipt-80', 'receipt-58'];
export const BLOCKS = ['heading', 'fields', 'lines', 'totals', 'text', 'signatures', 'spacer'];

/** A value as SQL takes it, for an expression worked out with known values. */
function literal(v) {
  if (v === null || v === undefined || v === '') return 'NULL';
  if (typeof v === 'boolean') return v ? '1' : '0';
  if (typeof v === 'number' && Number.isFinite(v)) return String(v);
  return `'${String(v).replace(/'/g, "''")}'`;
}

/**
 * Work out an expression over named values: `scope` maps a name (Field, or
 * Table.Field) to its value. A name the scope does not have is refused.
 */
export function evaluate(base, src, scope) {
  const sql = expressionSql(src, (table, field) => {
    const key = table ? `${table}.${field}` : field;
    if (Object.hasOwn(scope, key)) return literal(scope[key]);
    if (!table && Object.hasOwn(scope, field)) return literal(scope[field]);
    throw new Error(`[${table ? `${table}].[` : ''}${field}] is not a field of this document.`);
  });
  return base.query(`SELECT ${sql} AS v`).rows[0]?.v ?? null;
}

/** The field a table's records are numbered by (its AutoNumber, ID by default). */
const keyOf = (base, table) => (base.fields(table).find((f) => f.type === 'autonumber') || { name: 'ID' }).name;

/**
 * Words a related record is known by: a person's first and last name, else
 * its first text field. The same as a form's lists and a report's groups.
 */
function wordsFor(base, table, record) {
  if (!record) return null;
  const show = shownField(base.fields(table));
  return Array.isArray(show) ? show.map((f) => record[f]).filter((v) => v !== null && v !== undefined && v !== '').join(' ') : record[show];
}

/** What reads a document's values: records by key, relationships followed, cached for a run of many documents. */
function reader(base) {
  const rels = base.relationships();
  const cache = new Map();
  const typesCache = new Map();
  const record = (table, key) => {
    const id = `${table}\u0001${key}`;
    if (!cache.has(id)) cache.set(id, base.rows(table, { where: { [keyOf(base, table)]: key }, limit: 1 })[0] || null);
    return cache.get(id);
  };
  const typesOf = (table) => {
    if (!typesCache.has(table)) typesCache.set(table, Object.fromEntries(base.fields(table).map((f) => [f.name, f.type])));
    return typesCache.get(table);
  };
  const pointsTo = (table, field) => rels.find((r) => r.from.table === table && r.from.field === field) || null;
  /**
   * A name's value and type on a record of `table`: Field, or a path through
   * pointing fields (Customer.Address, Order.Customer.Name). A pointing field
   * named alone is shown by the words of the record it points at.
   */
  const value = (table, row, name) => {
    const steps = String(name).split('.');
    let t = table;
    let r = row;
    for (let i = 0; i < steps.length; i++) {
      const f = steps[i];
      if (!r) return { value: null, type: null };
      if (!typesOf(t)[f] && !Object.hasOwn(r, f)) throw new Error(`${t} has no field called ${f}.`);
      const rel = pointsTo(t, f);
      const last = i === steps.length - 1;
      if (rel && !last) { r = r[f] === null || r[f] === undefined ? null : record(rel.to.table, r[f]); t = rel.to.table; continue; }
      if (rel && last) return { value: wordsFor(base, rel.to.table, r[f] === null || r[f] === undefined ? null : record(rel.to.table, r[f])), type: 'text' };
      return { value: r[f] ?? null, type: typesOf(t)[f] || null };
    }
    return { value: null, type: null };
  };
  return { record, value, pointsTo, typesOf };
}

/** The field of `child` that points at `parent`, when the template does not say. */
function linkOf(base, child, parent) {
  const r = base.relationships().find((x) => x.from.table === child && x.to.table === parent);
  if (!r) throw new Error(`${child} has no field pointing at ${parent}: relate them, or say which field links them.`);
  return r.from.field;
}

/** {Field} and {Customer.Name} in words, filled from the record: [{ text } | { value, type }]. */
function fillText(text, get) {
  const parts = [];
  const re = /\{([^{}]+)\}/g;
  let at = 0;
  let m;
  while ((m = re.exec(text))) {
    if (m.index > at) parts.push({ text: text.slice(at, m.index) });
    const v = get(m[1].trim());
    parts.push({ value: v.value, type: v.type });
    at = m.index + m[0].length;
  }
  if (at < text.length) parts.push({ text: text.slice(at) });
  return parts;
}

/**
 * The documents for `layout`, one for each record of its table (or each of
 * `keys`): { layout, letterhead, documents: [{ key, blocks }] }, each block
 * with its values worked out, ready to draw.
 */
export function documentData(base, layout, { keys = null, limit = 1000 } = {}) {
  const table = layout.table;
  if (!table || !base.tables().some((t) => t.name === table)) throw new Error(`The template's table, ${table}, is not in this database.`);
  const read = reader(base);
  const keyField = keyOf(base, table);

  // The letterhead: the first record of the business's own details.
  let letterhead = null;
  const head = {};
  const headTypes = {};
  const lt = layout.letterhead?.table || null;
  if (lt) {
    const first = base.rows(lt, { limit: 1 })[0] || null;
    if (first) {
      for (const f of base.fields(lt)) { head[f.name] = first[f.name]; head[`${lt}.${f.name}`] = first[f.name]; headTypes[f.name] = f.type; }
      const names = layout.letterhead.fields?.length ? layout.letterhead.fields : base.fields(lt).filter((f) => f.type === 'text').map((f) => f.name);
      letterhead = { fields: names.map((n) => ({ name: n, ...read.value(lt, first, n) })) };
    }
  }

  const rows = keys ? keys.map((k) => read.record(table, k)).filter(Boolean) : base.rows(table, { limit });
  const documents = rows.map((row) => {
    // A name: the record's field (or a path through its pointing fields), else the letterhead's.
    // Either may say whose it is first: {Invoices.Date}, {Business Details.Address}.
    const get = (name) => {
      if (name.startsWith(`${table}.`)) return read.value(table, row, name.slice(table.length + 1));
      if (lt && name.startsWith(`${lt}.`)) { const f = name.slice(lt.length + 1); return { value: head[f] ?? null, type: headTypes[f] ?? null }; }
      if (Object.hasOwn(row, name) || name.includes('.')) return read.value(table, row, name);
      if (Object.hasOwn(head, name)) return { value: head[name], type: headTypes[name] ?? null };
      return read.value(table, row, name);
    };
    // What a total may use: the record's fields, the letterhead's, and the lines' sums.
    const scope = { ...head, ...row };
    const blocks = (layout.blocks || []).map((b) => {
      switch (b.type) {
        case 'heading':
          return { type: 'heading', parts: fillText(String(b.text || ''), get) };
        case 'fields':
          return { type: 'fields', columns: Math.min(3, Math.max(1, b.columns || 2)), items: (b.items || []).map((it) => ({ label: it.label || it.field.split('.').pop(), ...get(it.field) })) };
        case 'lines': {
          const child = b.table;
          const link = b.link || linkOf(base, child, table);
          const lines = base.rows(child, { where: { [link]: row[keyField] }, limit: 10000, order: b.sort || null });
          const types = read.typesOf(child);
          const columns = (b.columns || []).map((c) => ({ name: c.as || c.field, label: c.label || c.as || c.field, type: c.type || (c.field ? types[c.field] : null) || null, align: null }));
          for (const c of columns) c.align = ['currency', 'number', 'integer'].includes(c.type) ? 'right' : 'left';
          const out = lines.map((l) => {
            const values = {};
            for (const [i, c] of (b.columns || []).entries()) {
              if (c.expr) values[columns[i].name] = evaluate(base, c.expr, { ...head, ...l, ...values });
              else {
                const v = read.value(child, l, c.field);
                values[columns[i].name] = v.value;
              }
            }
            return values;
          });
          // The lines' sums, for the totals below: a column summed under its own name.
          for (const c of columns) {
            if (!['currency', 'number', 'integer'].includes(c.type) && !out.some((v) => typeof v[c.name] === 'number')) continue;
            scope[c.name] = Math.round(out.reduce((n, v) => n + (Number(v[c.name]) || 0), 0) * 1e6) / 1e6;
          }
          return { type: 'lines', columns, rows: out };
        }
        case 'totals': {
          const items = (b.rows || []).map((t) => {
            const v = evaluate(base, t.expr, scope);
            scope[t.label] = v;
            return { label: t.label, value: v, type: t.type || 'currency', strong: Boolean(t.strong) };
          });
          return { type: 'totals', rows: items };
        }
        case 'text':
          return { type: 'text', parts: fillText(String(b.text || ''), get), small: Boolean(b.small) };
        case 'signatures':
          return { type: 'signatures', labels: b.labels?.length ? b.labels : ['Signed'] };
        case 'spacer':
          return { type: 'spacer', size: b.size || 16 };
        default:
          throw new Error(`A template has no block called ${b.type}.`);
      }
    });
    return { key: row[keyField], blocks };
  });
  return { layout, letterhead, documents };
}

/**
 * A template for `table` as the wizard makes one: its fields at the top, the
 * first table whose records point at it as lines (with an Amount where they
 * have a quantity and a price), and a total. `fieldsOf`, `relationships` and
 * `tables` describe the database.
 */
export function wizardDocument({ table, fieldsOf, relationships, tables = [], name = null }) {
  const fields = fieldsOf(table).filter((f) => f.type !== 'autonumber');
  const children = relationships.filter((r) => r.to.table === table).map((r) => ({ table: r.from.table, link: r.from.field }));
  const business = tables.find((t) => /^(business|company)( details| info(rmation)?)?$/i.test(t)) || null;
  // Headed by what one record is (Invoices: Invoice), whatever the template is called.
  const blocks = [{ type: 'heading', text: table.replace(/s$/, '') }];
  blocks.push({ type: 'fields', columns: 2, items: [{ field: (fieldsOf(table).find((f) => f.type === 'autonumber') || { name: 'ID' }).name, label: 'Number' }, ...fields.slice(0, 7).map((f) => ({ field: f.name }))] });
  const child = children[0];
  if (child) {
    const cf = fieldsOf(child.table).filter((f) => f.type !== 'autonumber' && f.name !== child.link);
    const qty = cf.find((f) => /^(quantity|qty|hours|units)$/i.test(f.name));
    const price = cf.find((f) => f.type === 'currency');
    const columns = cf.slice(0, 5).map((f) => ({ field: f.name }));
    const totalCol = qty && price ? 'Amount' : price ? price.name : null;
    if (qty && price) columns.push({ expr: `Round([${qty.name}] * [${price.name}], 2)`, as: 'Amount', type: 'currency' });
    blocks.push({ type: 'lines', table: child.table, link: child.link, columns });
    if (totalCol) blocks.push({ type: 'totals', rows: [{ label: 'Total', expr: `[${totalCol}]`, strong: true }] });
  }
  return { kind: 'document', name: name || `${table} document`, title: name || table, table, paper: 'a4', ...(business ? { letterhead: { table: business } } : {}), blocks };
}
