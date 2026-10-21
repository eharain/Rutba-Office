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

/**
 * The guide inside each sample, as Northwind's start-up form is: a Start
 * Here table with the tutorial's tasks a step at a time, each step naming
 * the table, form, query, report or print template to open, and a report
 * to print it. It is there whether the sample is made with its records or
 * empty, made from the sample's `about.tutorial`.
 */
const GUIDE = { name: 'Start Here', fields: [
  { name: 'Order', type: 'integer' }, { name: 'Task', type: 'text', required: true }, { name: 'Step', type: 'integer' }, { name: 'What To Do', type: 'longtext' },
] };
const GUIDE_REPORT = { name: 'Start Here', title: 'Start here: the everyday tasks, step by step', table: 'Start Here', group: 'Task', sort: 'Order', columns: [{ field: 'Step', type: 'integer' }, 'What To Do'] };
const GUIDE_ABOUT = { about: 'The guide to this database: the everyday tasks a step at a time, each step saying which table, form, query, report or print template to open. Open its report, Start Here, to read or print it. Delete the table when you no longer need it; nothing else uses it.', fields: {
  Order: 'The order the steps are read in.', Task: 'The everyday task the step belongs to.', Step: 'The step\'s number within its task.', 'What To Do': 'What to open and what to do there.',
} };

/** A sample's own name as a web address: Northbank Trading is northbank-trading. */
const slugOf = (name) => String(name).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

/** A summary written as one string holds its paragraphs between blank lines. */
const paragraphs = (summary) => (Array.isArray(summary) ? summary : String(summary).split(/\n\s*\n/).map((p) => p.replace(/\s+/g, ' ').trim()).filter(Boolean));

const abouts = new WeakMap();
/** A sample module's documentation, the guide's table and report described with its own. */
function aboutOf(s) {
  if (abouts.has(s)) return abouts.get(s);
  const a = s.about;
  const guided = Boolean(a.tutorial?.length);
  const about = {
    ...a,
    summary: paragraphs(a.summary),
    tables: guided ? { 'Start Here': GUIDE_ABOUT, ...a.tables } : a.tables,
    reports: guided ? { ...a.reports, 'Start Here': 'The guide to this database printed: each everyday task and its steps, from the Start Here table.' } : a.reports,
  };
  abouts.set(s, about);
  return about;
}

const designs = new WeakMap();
/** A sample module's design: its own, with the guide first when it has a tutorial. */
function designOf(s) {
  if (designs.has(s)) return designs.get(s);
  const guided = Boolean(s.about?.tutorial?.length);
  const design = guided ? { ...s.spec, tables: [GUIDE, ...s.spec.tables], reports: [...(s.spec.reports || []), GUIDE_REPORT] } : s.spec;
  designs.set(s, design);
  return design;
}

/** The guide's steps for a sample module, as Start Here records. */
function guideRows(s) {
  const a = s.about;
  const rows = [];
  let order = 0;
  const step = (task, n, text) => rows.push({ Order: ++order, Task: task, Step: n, 'What To Do': text });
  const first = '0. About this database';
  step(first, 1, `${a.name}: ${a.line} Everything in it is invented: the business, the people, the addresses and the figures. The dates count back from the day it was made.`);
  step(first, 2, 'Every table has a form of the same name: open it from Forms to see one record at a time and to add one. Queries answer the questions the business asks; reports and print templates are what it prints.');
  step(first, 3, `The full guide, with every table, field, query and report explained, is at https://office.rutba.io/samples/${slugOf(a.name)}`);
  step(first, 4, 'Keep a copy before you change things: File, Save As makes one. Delete this Start Here table when you no longer need it.');
  a.tutorial.forEach((t, i) => t.steps.forEach((text, j) => step(`${i + 1}. ${t.title}`, j + 1, text)));
  return rows;
}

/** A sample's documentation: what the business is, each table, field, query and report in words. */
export function sampleAbout(id) {
  const s = SAMPLES[id];
  if (!s) throw new Error(`There is no sample database called ${id}.`);
  return aboutOf(s);
}

/** A sample's design, as buildFromSpec takes it: tables, queries, reports and print templates. */
export function sampleSpec(id) {
  const s = SAMPLES[id];
  if (!s) throw new Error(`There is no sample database called ${id}.`);
  return designOf(s);
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
  const made = buildFromSpec(base, designOf(s));
  // Thousands of records in one transaction: one write to the disk, not one each.
  base.db.exec('BEGIN');
  try {
    const w = writer(base);
    // The guide is there empty or not: it is how a person finds their way in.
    const guide = s.about?.tutorial?.length ? guideRows(s) : [];
    for (const row of guide) w.insert(GUIDE.name, row);
    const records = withData ? s.fill(w, { r: random(seed), cal: calendar(today) }) : {};
    base.db.exec('COMMIT');
    return { ...made, records: guide.length && withData ? { [GUIDE.name]: guide.length, ...records } : records };
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
