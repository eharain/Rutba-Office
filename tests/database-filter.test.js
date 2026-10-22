// Searching and filtering Data Files' records (D-2026-10-22-07): a table's
// records and a query's results alike, searched for words across every
// column, dates and numbers by a range, words by matching, yes/no, and by
// selection; a query's results paged, sorted and typed by the fields they
// show. Nothing typed becomes SQL.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Base, buildFromSpec } from '../packages/database/src/index.js';
import { filterSql, sniffType, filterKindOf } from '../packages/database/src/filter.js';

const ref = (t) => ({ type: 'integer', points: t });

function shop() {
  const b = Base.open(':memory:');
  buildFromSpec(b, {
    tables: [
      { name: 'Customers', fields: [{ name: 'Name', type: 'text' }, { name: 'Town', type: 'text' }, { name: 'Member', type: 'yesno' }] },
      { name: 'Orders', fields: [{ name: 'Customer', ...ref('Customers') }, { name: 'Ordered', type: 'date' }, { name: 'Amount', type: 'currency' }, { name: 'Note', type: 'text' }] },
    ],
    queries: {
      'Big orders': { tables: ['Orders', 'Customers'], columns: [{ table: 'Customers', field: 'Name' }, { table: 'Orders', field: 'Ordered' }, { table: 'Orders', field: 'Amount' }, { expr: '[Amount] * 2', as: 'Double' }] },
      'By customer': { tables: ['Orders', 'Customers'], columns: [{ table: 'Customers', field: 'Name', total: 'group' }, { table: 'Orders', field: 'Amount', total: 'sum' }] },
    },
  });
  b.insert('Customers', { Name: 'Amira Khan', Town: 'Leeds', Member: true });
  b.insert('Customers', { Name: 'Ben Ortiz', Town: 'York', Member: false });
  b.insert('Customers', { Name: 'Grace 50% Okafor', Town: 'Leeds', Member: true });
  b.insert('Orders', { Customer: 1, Ordered: '2026-10-01', Amount: 120, Note: 'paid' });
  b.insert('Orders', { Customer: 1, Ordered: '2026-10-15', Amount: 30, Note: '' });
  b.insert('Orders', { Customer: 2, Ordered: '2026-09-28', Amount: 75.5, Note: 'late' });
  b.insert('Orders', { Customer: 3, Ordered: '2026-10-20', Amount: 400, Note: null });
  return b;
}
const names = (rows) => rows.map((r) => r.Name);

test('a table: words searched across its columns, each word anywhere in the record', () => {
  const b = shop();
  try {
    assert.deepEqual(names(b.rows('Customers', { search: 'leeds' })), ['Amira Khan', 'Grace 50% Okafor']);
    assert.deepEqual(names(b.rows('Customers', { search: 'leeds amira' })), ['Amira Khan']);
    assert.equal(b.count('Customers', {}, { search: 'york' }), 1);
    // % and _ are themselves, not patterns.
    assert.deepEqual(names(b.rows('Customers', { search: '50%' })), ['Grace 50% Okafor']);
    assert.deepEqual(names(b.rows('Customers', { search: '_' })), []);
  } finally { b.close(); }
});

test('a table: words by matching, dates and numbers by a range, yes/no, by selection and empty', () => {
  const b = shop();
  try {
    assert.deepEqual(names(b.rows('Customers', { filters: [{ column: 'Name', kind: 'text', op: 'starts', value: 'b' }] })), ['Ben Ortiz']);
    assert.deepEqual(names(b.rows('Customers', { filters: [{ column: 'Town', kind: 'text', op: 'is', value: 'LEEDS' }] })), ['Amira Khan', 'Grace 50% Okafor']);
    assert.deepEqual(names(b.rows('Customers', { filters: [{ column: 'Member', kind: 'yesno', value: false }] })), ['Ben Ortiz']);
    const amounts = (filters) => b.rows('Orders', { filters }).map((r) => r.Amount);
    assert.deepEqual(amounts([{ column: 'Ordered', kind: 'range', from: '2026-10-01', to: '2026-10-15' }]), [120, 30]);
    assert.deepEqual(amounts([{ column: 'Amount', kind: 'range', from: 50 }]), [120, 75.5, 400]);
    assert.deepEqual(amounts([{ column: 'Amount', kind: 'range', from: 50, to: 200 }]), [120, 75.5]);
    assert.deepEqual(amounts([{ column: 'Note', kind: 'value', value: 'late' }]), [75.5]);
    assert.deepEqual(amounts([{ column: 'Note', kind: 'notValue', value: 'late' }]), [120, 30, 400]);
    assert.deepEqual(amounts([{ column: 'Note', kind: 'empty' }]), [30, 400]);
    assert.equal(b.count('Orders', {}, { filters: [{ column: 'Note', kind: 'notEmpty' }] }), 2);
    // Two filters at once: both hold.
    assert.deepEqual(amounts([{ column: 'Amount', kind: 'range', from: 50 }, { column: 'Ordered', kind: 'range', to: '2026-10-05' }]), [120, 75.5]);
  } finally { b.close(); }
});

test('a query\'s results: a page at a time, sorted, searched and filtered in the database, each column typed', () => {
  const b = shop();
  try {
    const all = b.queryRows({ name: 'Big orders' });
    assert.equal(all.count, 4);
    assert.deepEqual(all.columns, [{ name: 'Name', type: 'text' }, { name: 'Ordered', type: 'date' }, { name: 'Amount', type: 'currency' }, { name: 'Double', type: 'number' }]);
    const page = b.queryRows({ name: 'Big orders', order: '-Amount', limit: 2 });
    assert.deepEqual(page.rows.map((r) => r.Amount), [400, 120]);
    assert.equal(page.count, 4);
    assert.deepEqual(b.queryRows({ name: 'Big orders', order: '-Amount', limit: 2, offset: 2 }).rows.map((r) => r.Amount), [75.5, 30]);
    assert.deepEqual(b.queryRows({ name: 'Big orders', search: 'amira' }).rows.map((r) => r.Amount), [120, 30]);
    const ranged = b.queryRows({ name: 'Big orders', filters: [{ column: 'Ordered', kind: 'range', from: '2026-10-10' }, { column: 'Double', kind: 'range', from: 100 }] });
    assert.deepEqual(ranged.rows.map((r) => r.Name), ['Grace 50% Okafor']);
    assert.equal(ranged.count, 1);
    // A total's column is typed by what it holds.
    const totals = b.queryRows({ name: 'By customer' });
    assert.equal(totals.columns[1].type, 'number');
    assert.deepEqual(b.queryRows({ name: 'By customer', filters: [{ column: totals.columns[1].name, kind: 'range', from: 100 }] }).rows.map((r) => r.Name), ['Amira Khan', 'Grace 50% Okafor']);
  } finally { b.close(); }
});

test('nothing typed is SQL, and only the grid\'s own columns can be filtered', () => {
  const cols = [{ name: 'Name', type: 'text' }, { name: 'Amount', type: 'number' }];
  const f = filterSql(cols, { search: "x' OR 1=1 --", filters: [{ column: 'Name', kind: 'text', op: 'contains', value: '"; DROP TABLE x' }] });
  assert.ok(!/DROP|1=1/.test(f.sql), f.sql);
  assert.ok(f.params.some((p) => /DROP/.test(p)));
  assert.throws(() => filterSql(cols, { filters: [{ column: 'Secret', kind: 'value', value: 1 }] }), /no column called Secret/);
  assert.throws(() => filterSql(cols, { filters: [{ column: 'Amount', kind: 'range', from: 'many' }] }), /not a number/);
  assert.deepEqual(filterSql(cols, {}), { sql: '', params: [] });
});

test('each type is offered the filter it takes, and a calculation is typed by its values', () => {
  assert.equal(filterKindOf('date'), 'range');
  assert.equal(filterKindOf('currency'), 'range');
  assert.equal(filterKindOf('yesno'), 'yesno');
  assert.equal(filterKindOf('text'), 'text');
  assert.equal(sniffType([1, 2.5, null]), 'number');
  assert.equal(sniffType(['2026-10-01', '2026-10-02']), 'date');
  assert.equal(sniffType(['2026-10-01T09:30:00']), 'datetime');
  assert.equal(sniffType(['Leeds', 3]), 'text');
  assert.equal(sniffType([]), 'text');
});
