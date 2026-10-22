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
// Any block (and any totals row) can say when it is shown: `if`, an
// expression over the record, the letterhead and what came before. A template
// can say which records it is for (`where`), its lines which of them they list
// (`where`) and how they are found (`on`, any field of the record, not only
// its number), a column keep a running balance (`running`), and the lines be
// grouped under headings with subtotals (`group`). A value can say how it
// shows (`type`, `format`, a Format pattern as the query designer writes one);
// one that is empty shows as nothing.
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
  }, { today: base.today });
  return base.query(`SELECT ${sql} AS v`).rows[0]?.v ?? null;
}

/** Whether an expression's answer holds: anything but empty, zero, false or "0". */
const holds = (v) => !(v === null || v === undefined || v === '' || v === 0 || v === '0' || v === false);

/** A value as a Format pattern writes it ("0.00", "#,##0", "dd/mm/yyyy"), or as it is when it has none. */
function formatted(base, v, format) {
  if (!format || v === null || v === undefined || v === '') return v;
  return evaluate(base, `Format([v], "${String(format).replace(/"/g, '""')}")`, { v });
}

/** A value whose type nobody said, read from the value: a date, a date and time, or as it is. */
const typeOfValue = (v) => (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v) ? 'date' : typeof v === 'string' && /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}/.test(v) ? 'datetime' : null);

/** The field a table's records are numbered by (its AutoNumber, ID by default). */
const keyOf = (base, table) => (base.fields(table).find((f) => f.type === 'autonumber') || { name: 'ID' }).name;

/** A field that holds a name: First Name, Last Name, Surname, Name, Title, Company Name. */
const NAMED = /^(name|title|surname)$|[ _]?name$/i;

/** A person's first and last name, or a field called Name or Title: what a record is called by, if it has one. */
function nameOf(fields) {
  const named = (re) => fields.find((f) => f.type === 'text' && re.test(f.name))?.name;
  const first = named(/^(first|given)[ _]?name$/i);
  const last = named(/^(last|family)[ _]?name$|^surname$/i);
  return first && last ? [first, last] : named(/^(name|title)$/i) || null;
}

/**
 * Words a related record is called by on paper: the names among the fields
 * that name its table's records (Camille Sharma, not the list's EMP-0143
 * Camille Sharma); with no name among them, its own name if it has one; else
 * the words its lists show (a sale by its receipt number).
 */
function wordsFor(base, table, record) {
  if (!record) return null;
  const types = Object.fromEntries(base.fields(table).map((f) => [f.name, f.type]));
  const show = [].concat(base.shownOf ? base.shownOf(table) : shownField(base.fields(table))).filter((n) => n in types);
  let names = show.filter((n) => types[n] === 'text' && NAMED.test(n));
  // Named by a code (a client by its account, CL-0004), a record that has a person's name or a Name is called by it.
  if (!names.length) names = [].concat(nameOf(base.fields(table)) || []);
  if (!names.length && base.wordsOf && record._rowid !== undefined) return base.wordsOf(table, record._rowid);
  return (names.length ? names : show).map((f) => record[f]).filter((v) => v !== null && v !== undefined && v !== '').join(' ');
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

  // What a condition reads of a record: its fields, by name or by Table.Field, and the letterhead's.
  const scopeOf = (row) => ({ ...head, ...row, ...Object.fromEntries(Object.entries(row).map(([k, v]) => [`${table}.${k}`, v])) });
  // The records it is for: the ones asked for, or those its `where` holds for (a delivery note for shipped orders only).
  let rows;
  if (keys) rows = keys.map((k) => read.record(table, k)).filter(Boolean);
  else if (!layout.where) rows = base.rows(table, { limit });
  else {
    // The condition first, then the limit: this month's payslips are found however many came before them.
    rows = [];
    for (let offset = 0; rows.length < limit; offset += 1000) {
      const page = base.rows(table, { offset, limit: 1000 });
      for (const row of page) if (rows.length < limit && holds(evaluate(base, layout.where, scopeOf(row)))) rows.push(row);
      if (page.length < 1000) break;
    }
  }
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
    const scope = scopeOf(row);
    const blocks = (layout.blocks || []).map((b) => {
      // Shown only when its condition holds: cash tendered and change for a cash sale only.
      if (b.if && !holds(evaluate(base, b.if, scope))) return null;
      switch (b.type) {
        case 'heading':
          return { type: 'heading', parts: fillText(String(b.text || ''), get) };
        case 'fields':
          return {
            type: 'fields',
            columns: Math.min(3, Math.max(1, b.columns || 2)),
            items: (b.items || []).filter((it) => !it.if || holds(evaluate(base, it.if, scope))).map((it) => {
              const v = get(it.field);
              const type = it.type || v.type || typeOfValue(v.value);
              return it.format ? { label: it.label || it.field.split('.').pop(), value: formatted(base, v.value, it.format), type: 'text' } : { label: it.label || it.field.split('.').pop(), value: v.value, type };
            }),
          };
        case 'lines': {
          const child = b.table;
          // Found by the field that points at this record, or by `on`: a field of theirs to a field of this record's.
          const by = b.on && typeof b.on === 'object'
            // The record's own value (a pointing field's number, not its words), or one reached through a path.
            ? Object.fromEntries(Object.entries(b.on).map(([theirs, mine]) => [theirs, Object.hasOwn(row, mine) ? row[mine] : get(mine).value]))
            : { [b.link || linkOf(base, child, table)]: row[keyField] };
          let lines = base.rows(child, { where: by, limit: 10000, order: b.sort || null });
          // Only those its `where` holds for: a statement's unpaid invoices.
          if (b.where) lines = lines.filter((l) => holds(evaluate(base, b.where, { ...scope, ...l })));
          const types = read.typesOf(child);
          const columns = (b.columns || []).map((c) => ({ name: c.as || c.field || c.running, label: c.label || c.as || c.field || c.running, type: c.type || (c.field ? types[c.field] : null) || (c.running ? 'currency' : null), align: null, decimals: c.decimals ?? null }));
          // A running balance starts from its `start` (an opening balance), else nothing.
          const running = Object.fromEntries((b.columns || []).filter((c) => c.running).map((c) => [c.as || c.running, Number(c.start ? evaluate(base, c.start, scope) : 0) || 0]));
          for (const c of columns) c.align = ['currency', 'number', 'integer'].includes(c.type) ? 'right' : 'left';
          const out = lines.map((l) => {
            const values = {};
            for (const [i, c] of (b.columns || []).entries()) {
              if (c.running) {
                const name = columns[i].name;
                running[name] = Math.round((running[name] + (Number(values[c.running]) || 0)) * 1e6) / 1e6;
                values[name] = running[name];
              } else if (c.expr) values[columns[i].name] = evaluate(base, c.expr, { ...scope, ...l, ...values });
              else {
                const v = read.value(child, l, c.field);
                values[columns[i].name] = v.value;
                // A field pointing at another table shows that record's words: the column is words, not the number it holds.
                if (!c.type && v.type === 'text') columns[i].type = 'text';
              }
            }
            return values;
          });
          // A calculated column nobody typed: a date or a date and time read from its values.
          for (const c of columns) if (!c.type) c.type = out.map((v) => typeOfValue(v[c.name])).find(Boolean) || null;
          for (const c of columns) c.align = ['currency', 'number', 'integer'].includes(c.type) ? 'right' : 'left';
          // The lines' sums, for the totals below: a column summed under its own name. A running balance is not summed: its last is its total.
          const isNumber = (c) => ['currency', 'number', 'integer'].includes(c.type) || out.some((v) => typeof v[c.name] === 'number');
          for (const c of columns) {
            if (!isNumber(c)) continue;
            scope[c.name] = c.name in running ? (out.length ? out[out.length - 1][c.name] : running[c.name]) : Math.round(out.reduce((n, v) => n + (Number(v[c.name]) || 0), 0) * 1e6) / 1e6;
          }
          scope[`${b.as || child} count`] = out.length;
          // Each column's own Format pattern, applied once the sums are taken.
          const shownRows = out.map((v) => {
            const r = { ...v };
            for (const [i, c] of (b.columns || []).entries()) if (c.format) r[columns[i].name] = formatted(base, v[columns[i].name], c.format);
            return r;
          });
          for (const [i, c] of (b.columns || []).entries()) if (c.format) { columns[i].type = 'text'; }
          // Grouped: a heading for each value of the column, in the order they come, and a subtotal under each.
          if (b.group && columns.some((c) => c.name === b.group)) {
            const order = [];
            const groups = new Map();
            out.forEach((v, k) => {
              const key = v[b.group] ?? '';
              if (!groups.has(key)) { groups.set(key, []); order.push(key); }
              groups.get(key).push(k);
            });
            const rows = [];
            for (const key of order) {
              const members = groups.get(key);
              rows.push({ _group: shownRows[members[0]][b.group] ?? '', _count: members.length });
              for (const k of members) rows.push(shownRows[k]);
              if (b.subtotals !== false) {
                const sub = {};
                // `subtotals` may name the columns worth adding up (an amount, not a unit price); else every number.
                const named = Array.isArray(b.subtotals) ? new Set(b.subtotals) : null;
                for (const c of columns) if (isNumber(c) && !(c.name in running) && c.name !== b.group && (!named || named.has(c.name))) sub[c.name] = Math.round(members.reduce((n, k) => n + (Number(out[k][c.name]) || 0), 0) * 1e6) / 1e6;
                rows.push({ _subtotal: sub, _count: members.length });
              }
            }
            return { type: 'lines', columns, rows, group: b.group };
          }
          return { type: 'lines', columns, rows: shownRows };
        }
        case 'totals': {
          // A row with an `if` shows only when it holds; one that does not still counts as nothing for the rows after it.
          const items = (b.rows || []).flatMap((t) => {
            if (t.if && !holds(evaluate(base, t.if, scope))) { scope[t.label] = null; return []; }
            const v = evaluate(base, t.expr, scope);
            scope[t.label] = v;
            return [t.format ? { label: t.label, value: formatted(base, v, t.format), type: 'text', strong: Boolean(t.strong), decimals: null } : { label: t.label, value: v, type: t.type || 'currency', strong: Boolean(t.strong), decimals: t.decimals ?? null }];
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
    }).filter(Boolean);
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
