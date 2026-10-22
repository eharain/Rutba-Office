// Building Data Files' sample businesses. Their content (each sample's
// design, invented records and documentation) is kept in the templates'
// own repository, Rutba-Office-Templates, which Office takes as
// packages/templates (@rutba/templates/data-files); here each is made
// through the same calls a person's own design goes through (templates.js
// buildFromSpec), then, when asked for, filled with its records counted
// back from the day it is made.

import { buildFromSpec } from './templates.js';
import { q, toStored } from './types.js';
import { sampleModule, designOfModule, guideOf, GUIDE_TABLE, random, calendar } from '@rutba/templates/data-files/active';

/**
 * Make sample `id` in `base`, an empty database: its design, and its
 * invented records unless `withData` is false. `today` and `seed` fix what
 * the records are, for a test. Answers the names made and the records added.
 */
export function buildSample(base, id, options = {}) {
  return buildSampleFrom(base, sampleModule(id), options);
}

/** Make the sample module `s` ({ spec, fill, about }) in `base`, as buildSample does. */
export function buildSampleFrom(base, s, { withData = true, today = new Date(), seed = 20261021 } = {}) {
  const made = buildFromSpec(base, designOfModule(s));
  // Thousands of records in one transaction: one write to the disk, not one each.
  base.db.exec('BEGIN');
  try {
    const w = writer(base);
    // A sample's Start Here guide is there whether it is made with its records or empty.
    const guide = guideOf(s);
    for (const row of guide) w.insert(GUIDE_TABLE, row);
    const records = withData ? s.fill(w, { r: random(seed), cal: calendar(today) }) : {};
    base.db.exec('COMMIT');
    return { ...made, records: guide.length && withData ? { [GUIDE_TABLE]: guide.length, ...records } : records };
  } catch (err) {
    base.db.exec('ROLLBACK');
    throw err;
  }
}

/**
 * The database as a sample's fill writes to it: insert and update as the
 * Base has them, each table's fields read once and each statement prepared
 * once, so ten thousand records take a second rather than ten.
 */
function writer(base) {
  const fields = new Map();
  const statements = new Map();
  const fieldsOf = (table) => {
    if (!fields.has(table)) {
      const list = base.fields(table);
      if (!list.length) throw new Error(`There is no table called ${table}.`);
      fields.set(table, new Map(list.map((f) => [f.name, f])));
    }
    return fields.get(table);
  };
  const prepared = (sql) => {
    if (!statements.has(sql)) statements.set(sql, base.db.prepare(sql));
    return statements.get(sql);
  };
  const w = Object.create(base);
  w.insert = (table, values) => {
    const known = fieldsOf(table);
    for (const k of Object.keys(values)) if (!known.has(k)) throw new Error(`${table} has no field called ${k}.`);
    // A field not given takes its default from the table itself, as Base.insert does.
    const set = [...known.values()].filter((f) => Object.hasOwn(values, f.name) && values[f.name] !== undefined);
    const sql = set.length
      ? `INSERT INTO ${q(table)} (${set.map((f) => q(f.name)).join(', ')}) VALUES (${set.map(() => '?').join(', ')})`
      : `INSERT INTO ${q(table)} DEFAULT VALUES`;
    return Number(prepared(sql).run(...set.map((f) => toStored(f, values[f.name]))).lastInsertRowid);
  };
  w.update = (table, rowid, values) => {
    const known = fieldsOf(table);
    const set = Object.keys(values).map((k) => known.get(k) || (() => { throw new Error(`${table} has no field called ${k}.`); })());
    if (!set.length) return 0;
    return Number(prepared(`UPDATE ${q(table)} SET ${set.map((f) => `${q(f.name)} = ?`).join(', ')} WHERE rowid = ?`).run(...set.map((f) => toStored(f, values[f.name])), rowid).changes);
  };
  return w;
}
