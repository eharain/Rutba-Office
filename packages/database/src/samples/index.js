// Data Files' sample databases: a whole working database for each kind of
// business we make Office for, to learn from with its records or to start
// from empty. Each is made through the same calls a person's own design goes
// through (templates.js buildFromSpec), then, when asked for, filled with
// invented records counted back from the day it is made.

import { buildFromSpec } from '../templates.js';
import { q, toStored } from '../types.js';
import { calendar, random } from './util.js';
import * as trading from './trading.js';
import * as shop from './shop.js';
import * as clinic from './clinic.js';
import * as salon from './salon.js';
import * as repairs from './repairs.js';
import * as agency from './agency.js';
import * as cafe from './cafe.js';
import * as rentals from './rentals.js';
import * as school from './school.js';
import * as charity from './charity.js';
import * as workshop from './workshop.js';

const SAMPLES = { trading, shop, clinic, salon, repairs, agency, cafe, rentals, school, charity, workshop };

/** The samples, in the order a window offers them. */
export const SAMPLE_IDS = Object.keys(SAMPLES);

/** A sample's documentation: what the business is, each table, field, query and report in words. */
export function sampleAbout(id) {
  const s = SAMPLES[id];
  if (!s) throw new Error(`There is no sample database called ${id}.`);
  // A summary written as one string holds its paragraphs between blank lines.
  const summary = Array.isArray(s.about.summary) ? s.about.summary : String(s.about.summary).split(/\n\s*\n/).map((p) => p.replace(/\s+/g, ' ').trim()).filter(Boolean);
  return { ...s.about, summary };
}

/** A sample's design, as buildFromSpec takes it. */
export function sampleSpec(id) {
  const s = SAMPLES[id];
  if (!s) throw new Error(`There is no sample database called ${id}.`);
  return s.spec;
}

/**
 * Make sample `id` in `base`, an empty database: its design, and its
 * invented records unless `withData` is false. `today` and `seed` fix what
 * the records are, for a test. Answers the names made and the records added.
 */
export function buildSample(base, id, options = {}) {
  const s = SAMPLES[id];
  if (!s) throw new Error(`There is no sample database called ${id}.`);
  return buildSampleFrom(base, s, options);
}

/** Make the sample module `s` ({ spec, fill }) in `base`, as buildSample does. */
export function buildSampleFrom(base, s, { withData = true, today = new Date(), seed = 20261021 } = {}) {
  const made = buildFromSpec(base, s.spec);
  if (!withData) return { ...made, records: {} };
  // Thousands of records in one transaction: one write to the disk, not one each.
  base.db.exec('BEGIN');
  try {
    const records = s.fill(writer(base), { r: random(seed), cal: calendar(today) });
    base.db.exec('COMMIT');
    return { ...made, records };
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
    // A Yes/No field left out takes its default: Active is Yes unless a record says otherwise.
    const set = [...known.values()].filter((f) => Object.hasOwn(values, f.name) || f.type === 'yesno');
    for (const k of Object.keys(values)) if (!known.has(k)) throw new Error(`${table} has no field called ${k}.`);
    const value = (f) => (Object.hasOwn(values, f.name) ? values[f.name] : f.default ?? false);
    const sql = `INSERT INTO ${q(table)} (${set.map((f) => q(f.name)).join(', ')}) VALUES (${set.map(() => '?').join(', ')})`;
    return Number(prepared(sql).run(...set.map((f) => toStored(f, value(f)))).lastInsertRowid);
  };
  w.update = (table, rowid, values) => {
    const known = fieldsOf(table);
    const set = Object.keys(values).map((k) => known.get(k) || (() => { throw new Error(`${table} has no field called ${k}.`); })());
    if (!set.length) return 0;
    return Number(prepared(`UPDATE ${q(table)} SET ${set.map((f) => `${q(f.name)} = ?`).join(', ')} WHERE rowid = ?`).run(...set.map((f) => toStored(f, values[f.name])), rowid).changes);
  };
  return w;
}
