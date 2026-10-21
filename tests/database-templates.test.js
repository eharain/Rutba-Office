// Data Files' ready-made databases (packages/database/src/templates.js): each
// one made in an empty database with its tables, the relationships its
// pointing fields stand for, queries that run, a form for each table and
// reports that draw; a template refused in a database that has tables; and
// a person shown by first and last name in lists and reports.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Base, TEMPLATE_IDS, templateTables, buildTemplate, shownField } from '../packages/database/src/index.js';

test('every template is made whole: tables, relationships, queries that run, a form a table, reports that draw', () => {
  assert.deepEqual(TEMPLATE_IDS, ['contacts', 'inventory', 'membership', 'crm', 'library']);
  for (const id of TEMPLATE_IDS) {
    const b = Base.open(':memory:');
    try {
      const made = buildTemplate(b, id);
      assert.deepEqual(b.tables().map((t) => t.name), made.tables, id);
      assert.deepEqual(templateTables(id).map((t) => t.table), made.tables, id);
      assert.equal(b.forms().length, made.tables.length, `${id}: a form a table`);
      for (const q of made.queries) {
        const out = b.runQuery({ name: q });
        assert.equal(new Set(out.columns).size, out.columns.length, `${id}: ${q}'s columns are named apart`);
      }
      for (const r of made.reports) assert.equal(b.reportData({ name: r }).count, 0, `${id}: ${r}`);
      const pointing = templateTables(id).length - 1;
      assert.ok(b.relationships().length >= Math.min(1, pointing), `${id}: related`);
    } finally {
      b.close();
    }
  }
});

test("the inventory's query finds what is running low, by product and supplier", () => {
  const b = Base.open(':memory:');
  try {
    buildTemplate(b, 'inventory');
    b.insert('Suppliers', { Name: 'Acme' });
    b.insert('Products', { Name: 'Bolts', SKU: 'B1', Supplier: 1, 'Units in Stock': 3 });
    b.insert('Products', { Name: 'Nuts', SKU: 'N1', Supplier: 1, 'Units in Stock': 50 });
    b.insert('Products', { Name: 'Rivets', SKU: 'R1', Supplier: 1, 'Units in Stock': 1, Discontinued: true });
    assert.deepEqual(b.runQuery({ name: 'Low stock' }).rows.map((r) => [r.Product, r.Supplier]), [['Bolts', 'Acme']]);
    assert.throws(() => b.createTable({ name: 'Products' }), /already/);
    assert.throws(() => buildTemplate(b, 'contacts'), /empty database/);
    assert.throws(() => buildTemplate(Base.open(':memory:'), 'nothing'), /no template/);
  } finally {
    b.close();
  }
});

test('a person is shown by first and last name, in a lookup list and in a report', () => {
  assert.deepEqual(shownField([{ name: 'ID', type: 'autonumber' }, { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text' }]), ['First Name', 'Last Name']);
  assert.equal(shownField([{ name: 'ID', type: 'autonumber' }, { name: 'Title', type: 'text' }]), 'Title');
  const b = Base.open(':memory:');
  try {
    buildTemplate(b, 'library');
    b.insert('Books', { Title: 'Middlemarch' });
    b.insert('Borrowers', { 'First Name': 'Amira', 'Last Name': 'Haddad' });
    b.insert('Borrowers', { 'Last Name': 'Bell' });
    b.insert('Loans', { Book: 1, Borrower: 2, 'Lent On': '2026-10-01' });
    b.insert('Loans', { Book: 1, Borrower: 1, 'Lent On': '2026-10-02' });
    const lookup = b.forms().find((f) => f.name === 'Loans').layout.controls.find((c) => c.field === 'Borrower').lookup;
    assert.deepEqual(b.lookup(lookup).map((o) => o.text), ['Bell', 'Amira Haddad'], 'listed by last name');
    const d = b.reportData({ name: 'Loans by borrower' });
    assert.deepEqual(d.groups.map((g) => g.label), ['Amira Haddad', 'Bell']);
    assert.equal(d.groups[0].rows[0].Book, 'Middlemarch');
    assert.equal(d.totals['count:Book'], 2);
  } finally {
    b.close();
  }
});
