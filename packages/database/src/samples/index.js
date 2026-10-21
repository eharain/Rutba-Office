// Data Files' sample databases: a whole working database for each kind of
// business we make Office for, to learn from with its records or to start
// from empty. Each is made through the same calls a person's own design goes
// through (templates.js buildFromSpec), then, when asked for, filled with
// invented records counted back from the day it is made.

import { buildFromSpec } from '../templates.js';
import { calendar, random } from './util.js';
import * as trading from './trading.js';
import * as shop from './shop.js';

const SAMPLES = { trading, shop };

/** The samples, in the order a window offers them. */
export const SAMPLE_IDS = Object.keys(SAMPLES);

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
export function buildSample(base, id, { withData = true, today = new Date(), seed = 20261021 } = {}) {
  const s = SAMPLES[id];
  if (!s) throw new Error(`There is no sample database called ${id}.`);
  const made = buildFromSpec(base, s.spec);
  if (!withData) return { ...made, records: {} };
  // Thousands of records in one transaction: one write to the disk, not one each.
  base.db.exec('BEGIN');
  try {
    const records = s.fill(base, { r: random(seed), cal: calendar(today) });
    base.db.exec('COMMIT');
    return { ...made, records };
  } catch (err) {
    base.db.exec('ROLLBACK');
    throw err;
  }
}
