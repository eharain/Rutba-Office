// A newer design for a database made from a Data Files template or sample
// (owner, 2026-10-22, D-2026-10-22-04: "we will have to create some way to
// update the templates themself as well"; "i meant the data file templates").
//
// The set in use is built afresh in memory and set against what the database
// kept when it was made (origin.js): a table or field the newer design has and
// the database lacks is added; a query, form, report or print template the
// person has not changed since is replaced by the newer one; one they changed
// is left as it is, the newer one put beside it as "<name> (updated)"; one
// that is new is added. Records are never touched, and nothing is removed.

import { Base } from './base.js';
import { buildFromSpec } from './templates.js';
import { DESIGN_KINDS, designPrint, printOf, originOf, originPrint, activeSet } from './origin.js';
import { templateSpec, sampleModule, designOfModule, TEMPLATE_IDS, SAMPLE_IDS } from '@rutba/templates/data-files/active';

const SAVE = {
  queries: (b, name, v) => b.saveQuery(name, v),
  forms: (b, name, v) => b.saveForm({ ...v, name }),
  reports: (b, name, v) => b.saveReport({ ...v, name }),
  documents: (b, name, v) => b.saveDocument({ ...v, name }),
};

/** The design the set in use has for this origin, or null when it has none now. */
function specFor(origin) {
  try {
    if (origin.kind === 'template' && TEMPLATE_IDS.includes(origin.id)) return templateSpec(origin.id);
    if (origin.kind === 'sample' && SAMPLE_IDS.includes(origin.id)) return designOfModule(sampleModule(origin.id));
  } catch { /* gone from the set */ }
  return null;
}

/**
 * What a newer design would change in `base`: { available, origin, tables,
 * fields, added, replaced, beside } (each a list), available false when the
 * database came from nowhere, kept no fingerprints, or is up to date.
 */
export function planUpgrade(base) {
  const origin = originOf(base);
  const was = originPrint(base);
  const none = { available: false, origin, tables: [], fields: [], added: [], replaced: [], beside: [] };
  if (!origin || !was) return none;
  const spec = specFor(origin);
  if (!spec) return none;
  const fresh = Base.open(':memory:');
  try {
    buildFromSpec(fresh, spec);
    const now = designPrint(base);
    const plan = { ...none, spec, fresh: designPrint(fresh), items: {} };
    const have = new Set(base.tables().map((t) => t.name));
    for (const t of spec.tables) {
      if (!have.has(t.name)) { plan.tables.push(t.name); continue; }
      const fields = new Set(base.fields(t.name).map((f) => f.name));
      for (const f of t.fields) if (!fields.has(f.name)) plan.fields.push({ table: t.name, field: f.name });
    }
    for (const [kind, read] of Object.entries(DESIGN_KINDS)) {
      plan.items[kind] = Object.fromEntries(read(fresh));
      for (const [name, value] of Object.entries(plan.items[kind])) {
        const print = printOf(value);
        const kept = was[kind]?.[name];
        if (!(name in now[kind])) { if (print !== kept) plan.added.push({ kind, name }); continue; }
        // The template's own item has not changed: nothing to do.
        if (print === kept) continue;
        if (now[kind][name] === kept) plan.replaced.push({ kind, name });
        else plan.beside.push({ kind, name, as: `${name} (updated)` });
      }
    }
    plan.available = Boolean(plan.tables.length || plan.fields.length || plan.added.length || plan.replaced.length || plan.beside.length);
    return plan;
  } finally {
    fresh.close();
  }
}

/** Bring `base` up to the newer design as planUpgrade says; answers what was done. */
export function applyUpgrade(base) {
  const plan = planUpgrade(base);
  if (!plan.available) return { done: false, ...strip(plan) };
  const { spec } = plan;
  const newTables = new Set(plan.tables);
  // Tables first, their relationships once every table is there, then the design that reads them.
  for (const t of spec.tables.filter((x) => newTables.has(x.name))) base.createTable({ name: t.name, fields: t.fields.map(({ points, ...f }) => f) });
  for (const { table, field } of plan.fields) {
    const f = spec.tables.find((t) => t.name === table).fields.find((x) => x.name === field);
    const { points, ...plain } = f;
    base.addField(table, plain);
  }
  const rels = base.relationships();
  for (const t of spec.tables) {
    for (const f of t.fields.filter((g) => g.points)) {
      if (!newTables.has(t.name) && !plan.fields.some((x) => x.table === t.name && x.field === f.name)) continue;
      if (rels.some((r) => r.from.table === t.name && r.from.field === f.name)) continue;
      base.addRelationship({ from: { table: t.name, field: f.name }, to: { table: f.points, field: 'ID' } });
    }
  }
  for (const kind of Object.keys(DESIGN_KINDS)) {
    for (const x of plan.added.filter((i) => i.kind === kind)) SAVE[kind](base, x.name, plan.items[kind][x.name]);
    for (const x of plan.replaced.filter((i) => i.kind === kind)) SAVE[kind](base, x.name, plan.items[kind][x.name]);
    for (const x of plan.beside.filter((i) => i.kind === kind)) SAVE[kind](base, x.as, plan.items[kind][x.name]);
  }
  // What it was made from is now the newer design: the same change is not offered again.
  base.setMeta('origin', JSON.stringify({ ...plan.origin, set: activeSet(), updated: new Date().toISOString() }));
  base.setMeta('origin.design', JSON.stringify(plan.fresh));
  return { done: true, ...strip(plan) };
}

const strip = ({ available, origin, tables, fields, added, replaced, beside }) => ({ available, origin, tables, fields, added, replaced, beside });
