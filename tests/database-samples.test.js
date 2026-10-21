// Data Files' sample databases (packages/templates/data-files, built by packages/database/src/samples.js): each one
// builds with its invented records and without them; with them, every table
// has records, every query answers something, every report has records and
// totals that come to a number, nothing is below nought where it should not
// be, and every e-mail address and phone number is one kept for fiction
// (check.js); the same day and seed make the same file; and every table,
// field, query and report is described for the documentation.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Base, SAMPLE_IDS, buildSample, sampleSpec, sampleAbout, checkSample } from '../packages/database/src/index.js';

/** A sample made in memory, handed to `fn`, then closed. */
function made(id, options, fn) {
  const b = Base.open(':memory:');
  try {
    const out = buildSample(b, id, options);
    return fn(b, out);
  } finally {
    b.close();
  }
}

test('there are eleven samples, one for each kind of business we make Office for', () => {
  assert.deepEqual(SAMPLE_IDS, ['trading', 'shop', 'clinic', 'salon', 'repairs', 'agency', 'cafe', 'rentals', 'school', 'charity', 'workshop']);
});

for (const id of SAMPLE_IDS) {
  test(`${id}: with its records, every table, query and report has something and checkSample finds nothing`, () => {
    made(id, {}, (b, out) => {
      const spec = sampleSpec(id);
      assert.deepEqual(out.tables, spec.tables.map((t) => t.name));
      assert.ok(Object.values(out.records).reduce((a, n) => a + n, 0) > 500, 'enough records to learn from');
      assert.deepEqual(checkSample(b, spec), []);
      assert.equal(b.forms().length, spec.tables.length, 'a form for each table');
      assert.ok(b.queries().length >= 8, 'the questions the business asks');
      assert.ok(b.reports().length >= 4, 'reports to print');
    });
  });

  test(`${id}: another seed makes other records that are just as sound`, () => {
    made(id, { seed: 7 }, (b) => assert.deepEqual(checkSample(b, sampleSpec(id)), []));
  });

  test(`${id}: empty, it has the whole design and no records, and every query runs`, () => {
    made(id, { withData: false }, (b, out) => {
      assert.deepEqual(out.records, {});
      for (const t of b.tables()) assert.equal(b.count(t.name), 0, t.name);
      for (const q of b.queries()) assert.equal(b.runQuery({ name: q.name }).count, 0, q.name);
      for (const r of b.reports()) assert.equal(b.reportData({ name: r.name }).count, 0, r.name);
      assert.ok(b.relationships().length >= 2);
    });
  });

  test(`${id}: the same day and seed make the same records`, () => {
    const today = new Date(2026, 9, 21, 12);
    const snapshot = () => made(id, { today, seed: 3 }, (b) => b.tables().map((t) => JSON.stringify(b.rows(t.name, { limit: 50 }))).join('\n'));
    assert.equal(snapshot(), snapshot());
  });

  test(`${id}: every table, field, query and report is described for the documentation`, () => {
    const spec = sampleSpec(id);
    const about = sampleAbout(id);
    for (const k of ['name', 'kind', 'line', 'empty']) assert.ok(typeof about[k] === 'string' && about[k].length > 2, k);
    assert.ok(about.summary.length >= 1 && about.adapt.length >= 3 && about.suits.length >= 2);
    for (const t of spec.tables) {
      assert.ok(about.tables[t.name]?.about, `the table ${t.name}`);
      for (const f of t.fields) assert.ok(about.tables[t.name].fields?.[f.name], `${t.name}.${f.name}`);
    }
    assert.deepEqual(Object.keys(about.tables).sort(), spec.tables.map((t) => t.name).sort(), 'nothing described that is not there');
    assert.deepEqual(Object.keys(about.queries).sort(), Object.keys(spec.queries).sort());
    assert.deepEqual(Object.keys(about.reports).sort(), spec.reports.map((r) => r.name).sort());
    // The words go on the website as they are: plain ASCII punctuation, no long dashes.
    const words = JSON.stringify(about);
    assert.doesNotMatch(words, /[\u2013\u2014\u2018\u2019\u201c\u201d\u00a0\u200b]/);
  });
}

test('a sample is made only in an empty database, and an unknown one is refused', () => {
  const b = Base.open(':memory:');
  try {
    assert.throws(() => buildSample(b, 'nursery'), /no sample database called nursery/);
    buildSample(b, 'shop', { withData: false });
    assert.throws(() => buildSample(b, 'shop'), /empty database/);
  } finally {
    b.close();
  }
});
