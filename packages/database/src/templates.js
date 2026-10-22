// Data Files' simple ready-made databases, built here: their designs are
// kept in the templates' own repository, Rutba-Office-Templates (taken as
// packages/templates, @rutba/templates/data-files). buildFromSpec makes any
// design (a template's, a sample business's) through the same calls a
// person's own design goes through, so every table, relationship and query
// is a numbered step like any other and can be changed after.

import { TEMPLATE_IDS, templateTables, templateSpec } from '@rutba/templates/data-files/active';
import { recordOrigin } from './origin.js';

export { TEMPLATE_IDS, templateTables };

/**
 * Make template `id` in `base`, an empty database: its tables, the
 * relationships its pointing fields stand for, its queries, a form for each
 * table and its reports. Answers the names made.
 */
export function buildTemplate(base, id) {
  const made = buildFromSpec(base, templateSpec(id));
  // Which template it came from, so its page can be linked and a newer design offered (origin.js).
  recordOrigin(base, { kind: 'template', id });
  return made;
}

/**
 * Make a database from a spec: { tables: [{ name, fields: [{ name, type,
 * points?: table, ...}] }], queries: { name: design }, reports: [{ name,
 * table | query, columns: [field | { field, label?, type? }], group?, sort?,
 * desc?, totals? }] }. Relationships are made in the order the pointing
 * fields come, which is the order a query joins two tables by.
 */
export function buildFromSpec(base, t) {
  if (base.tables().length) throw new Error('A template is made in an empty database.');
  for (const x of t.tables) {
    base.createTable({ name: x.name, fields: x.fields.map(({ points, ...f }) => f) });
  }
  for (const x of t.tables) {
    for (const f of x.fields.filter((g) => g.points)) {
      base.addRelationship({ from: { table: x.name, field: f.name }, to: { table: f.points, field: 'ID' } });
    }
  }
  // What names each table's records, where the spec says (a sale by its Receipt No, not its payment).
  for (const x of t.tables) if (x.show) base.setShown(x.name, x.show);
  for (const [name, design] of Object.entries(t.queries || {})) base.saveQuery(name, design);
  for (const x of t.tables) base.makeForm(x.name);
  for (const r of t.reports || []) {
    const made = base.makeReport(r.query ? { query: r.query, name: r.name } : { table: r.table, name: r.name });
    const label = (field) => made.columns.find((c) => c.field === field)?.label || field;
    // A field that points at another table's record shows its words, so it stands on the left.
    const spec = r.table ? t.tables.find((x) => x.name === r.table) : null;
    const pointing = new Set(spec ? spec.fields.filter((f) => f.points).map((f) => f.name) : []);
    const types = r.table ? Object.fromEntries(base.fields(r.table).map((f) => [f.name, pointing.has(f.name) ? 'text' : f.type])) : {};
    base.saveReport({
      ...made,
      title: r.title || r.name,
      columns: r.columns.map((c) => {
        const col = typeof c === 'string' ? { field: c } : c;
        const type = col.type || types[col.field] || null;
        return { field: col.field, label: col.label || label(col.field), align: ['currency', 'number', 'integer'].includes(type) ? 'right' : 'left', ...(type ? { type } : {}), ...(col.decimals !== undefined ? { decimals: col.decimals } : {}) };
      }),
      group: r.group || null,
      sort: r.sort || null,
      desc: Boolean(r.desc),
      totals: r.totals || [],
    });
  }
  // Print templates: an invoice, a receipt, a payslip (document-design.js).
  for (const d of t.documents || []) base.saveDocument(d);
  return {
    tables: t.tables.map((x) => x.name),
    queries: Object.keys(t.queries || {}),
    reports: (t.reports || []).map((r) => r.name),
    documents: (t.documents || []).map((d) => d.name),
  };
}
