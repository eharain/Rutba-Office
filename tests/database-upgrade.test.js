// A database made from a Data Files template or sample keeps where it came
// from (D-2026-10-22-03), and is offered a newer design when the set in use
// has one (D-2026-10-22-04): a field or item the newer design adds is added,
// an item the person left alone is replaced, one they changed keeps their
// version with the newer beside it, and the records are left as they are.
// The "older" design here is the same template with parts taken back out of
// the database and out of what it remembers, as a database made from an
// earlier set would be.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Base, buildTemplate, buildSample, originOf, planUpgrade, applyUpgrade, designPrint } from '../packages/database/src/index.js';
import { printOf } from '../packages/database/src/origin.js';

/** Pretend the database was made from a design without `undo`'s parts, and remembers it so. */
function older(base, undo) {
  const print = designPrint(base);
  undo(base, print);
  base.setMeta('origin.design', JSON.stringify(print));
}

test('a database made from a template or a sample remembers which, from which set, and is up to date', () => {
  const t = Base.open(':memory:');
  const s = Base.open(':memory:');
  try {
    buildTemplate(t, 'contacts');
    buildSample(s, 'shop', { withData: false });
    assert.deepEqual({ kind: originOf(t).kind, id: originOf(t).id }, { kind: 'template', id: 'contacts' });
    assert.match(originOf(t).set, /^\d+\.\d+\.\d+$/);
    assert.equal(originOf(s).kind, 'sample');
    assert.equal(planUpgrade(t).available, false);
    assert.equal(planUpgrade(s).available, false);
    // One of the person's own: nothing to offer.
    const own = Base.open(':memory:');
    own.createTable({ name: 'Things', fields: [{ name: 'Name', type: 'text' }] });
    assert.equal(originOf(own), null);
    assert.equal(planUpgrade(own).available, false);
    own.close();
  } finally { t.close(); s.close(); }
});

test('a newer design adds the fields and items it brings, and replaces those the person left alone, records and all kept', () => {
  const b = Base.open(':memory:');
  try {
    buildTemplate(b, 'contacts');
    const query = b.queries()[0];
    const report = b.reports()[0];
    b.insert('Contacts', { 'First Name': 'Amira', 'Last Name': 'Khan', Notes: 'met at the fair' });
    older(b, (base, print) => {
      // Made before the template had Notes or its report, and with an older query.
      base.removeField('Contacts', 'Notes');
      print.tables.Contacts = print.tables.Contacts.filter((f) => f !== 'Notes');
      base.removeReport(report.name);
      delete print.reports[report.name];
      const old = { ...query.design, columns: query.design.columns.slice(0, 1) };
      base.saveQuery(query.name, old);
      print.queries[query.name] = printOf(old);
    });
    const plan = planUpgrade(b);
    assert.equal(plan.available, true);
    assert.deepEqual(plan.fields, [{ table: 'Contacts', field: 'Notes' }]);
    assert.deepEqual(plan.added, [{ kind: 'reports', name: report.name }]);
    assert.deepEqual(plan.replaced, [{ kind: 'queries', name: query.name }]);
    assert.deepEqual(plan.beside, []);
    const done = applyUpgrade(b);
    assert.equal(done.done, true);
    assert.ok(b.fields('Contacts').some((f) => f.name === 'Notes'));
    assert.ok(b.reports().some((r) => r.name === report.name));
    assert.deepEqual(b.queries().find((x) => x.name === query.name).design, query.design);
    // The record is there, untouched; the field taken away and added back starts empty.
    assert.equal(b.count('Contacts'), 1);
    assert.equal(b.rows('Contacts')[0]['Last Name'], 'Khan');
    // And the same change is not offered twice.
    assert.equal(planUpgrade(b).available, false);
  } finally { b.close(); }
});

test('an item the person changed keeps their version, the newer one put beside it', () => {
  const b = Base.open(':memory:');
  try {
    buildTemplate(b, 'inventory');
    const query = b.queries()[0];
    older(b, (base, print) => {
      // The template's older query, which the person then changed for their own.
      print.queries[query.name] = printOf({ ...query.design, columns: query.design.columns.slice(0, 1) });
      base.saveQuery(query.name, { ...query.design, columns: query.design.columns.slice(0, 2) });
    });
    const mine = b.queries().find((x) => x.name === query.name).design;
    const plan = planUpgrade(b);
    assert.deepEqual(plan.beside, [{ kind: 'queries', name: query.name, as: `${query.name} (updated)` }]);
    applyUpgrade(b);
    assert.deepEqual(b.queries().find((x) => x.name === query.name).design, mine);
    assert.deepEqual(b.queries().find((x) => x.name === `${query.name} (updated)`).design, query.design);
    assert.equal(planUpgrade(b).available, false);
  } finally { b.close(); }
});
