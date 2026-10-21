// Data Files' calculated columns (packages/database/src/expression.js): an
// expression written as Access writes one, read into SQL from a fixed set
// of pieces; a [Field] found in the one table of the query that has it; the
// functions a business asks for (Round, Nz, IIf, Format, DateDiff, Year...);
// totals, criteria and a sort on a calculated column; and anything else
// refused in words, never run.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Base, expressionSql, splitNamedExpression } from '../packages/database/src/index.js';

function shop() {
  const b = Base.open(':memory:');
  b.createTable({ name: 'Customers', fields: [{ name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text' }, { name: 'Country', type: 'text' }] });
  b.createTable({ name: 'Orders', fields: [{ name: 'Customer', type: 'integer' }, { name: 'Order Date', type: 'date' }, { name: 'Paid', type: 'yesno' }] });
  b.createTable({ name: 'Lines', fields: [{ name: 'Order', type: 'integer' }, { name: 'Quantity', type: 'integer' }, { name: 'Unit Price', type: 'currency' }, { name: 'Discount', type: 'number' }] });
  b.addRelationship({ from: { table: 'Orders', field: 'Customer' }, to: { table: 'Customers', field: 'ID' } });
  b.addRelationship({ from: { table: 'Lines', field: 'Order' }, to: { table: 'Orders', field: 'ID' } });
  b.insert('Customers', { 'First Name': 'Amira', 'Last Name': 'Haddad', Country: 'Jordan' });
  b.insert('Customers', { 'First Name': 'Ben', 'Last Name': 'Cole', Country: 'UK' });
  b.insert('Orders', { Customer: 1, 'Order Date': '2026-01-05', Paid: true });
  b.insert('Orders', { Customer: 2, 'Order Date': '2026-02-11', Paid: false });
  b.insert('Orders', { Customer: 1, 'Order Date': '2026-02-20', Paid: false });
  b.insert('Lines', { Order: 1, Quantity: 3, 'Unit Price': 10, Discount: 0.1 });
  b.insert('Lines', { Order: 1, Quantity: 1, 'Unit Price': 5 });
  b.insert('Lines', { Order: 2, Quantity: 2, 'Unit Price': 7.5 });
  b.insert('Lines', { Order: 3, Quantity: 4, 'Unit Price': 2.5 });
  return b;
}

test("a named expression in a Field cell, as Access's designer takes it", () => {
  assert.deepEqual(splitNamedExpression('Line Total: [Quantity] * [Unit Price]'), { as: 'Line Total', expr: '[Quantity] * [Unit Price]' });
  assert.deepEqual(splitNamedExpression('[Quantity] * 2'), { as: null, expr: '[Quantity] * 2' });
  assert.deepEqual(splitNamedExpression('IIf([Paid], "Paid: yes", "No")'), { as: null, expr: 'IIf([Paid], "Paid: yes", "No")' }, 'a colon in quotes or brackets is the text');
});

test('line totals with discounts, and sales by month: a total on a calculated column, grouped by another', () => {
  const b = shop();
  try {
    const lines = b.runQuery({ design: { tables: ['Lines'], columns: [{ table: 'Lines', field: 'Order' }, { expr: 'Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)', as: 'Line Total' }] } });
    assert.deepEqual(lines.columns, ['Order', 'Line Total']);
    assert.deepEqual(lines.rows.map((r) => r['Line Total']), [27, 5, 15, 10]);
    const months = b.runQuery({ design: { tables: ['Orders', 'Lines'], columns: [
      { expr: 'Format([Order Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { expr: '[Quantity] * [Unit Price] * (1 - Nz([Discount]))', as: 'Sales', total: 'sum' },
    ] } });
    assert.deepEqual(months.rows.map((r) => [r.Month, Math.round(r.Sales * 100) / 100]), [['2026-01', 32], ['2026-02', 25]]);
  } finally {
    b.close();
  }
});

test('text joined with &, IIf, criteria and a sort on a calculated column', () => {
  const b = shop();
  try {
    const r = b.runQuery({ design: { tables: ['Customers', 'Orders'], columns: [
      { expr: '[First Name] & " " & [Last Name]', as: 'Name', sort: 'asc' },
      { expr: 'IIf([Paid], "Paid", "Due")', as: 'State', criteria: ['"Due"'] },
      { expr: 'DateDiff("d", [Order Date], #2026-03-01#)', as: 'Days' },
    ] } });
    assert.deepEqual(r.rows, [{ Name: 'Amira Haddad', State: 'Due', Days: 9 }, { Name: 'Ben Cole', State: 'Due', Days: 18 }]);
    const unnamed = b.runQuery({ design: { tables: ['Lines'], columns: [{ expr: '[Quantity] / 2' }, { expr: 'Year(#2026-10-21#) Mod 100' }] } });
    assert.deepEqual(unnamed.columns, ['Expr1', 'Expr2']);
    assert.equal(unnamed.rows[0].Expr1, 1.5, 'division in decimals, as Access divides');
    assert.equal(unnamed.rows[0].Expr2, 26);
  } finally {
    b.close();
  }
});

test('a [Field] two tables have must say which; [Table].[Field] does', () => {
  const b = shop();
  try {
    assert.throws(() => b.runQuery({ design: { tables: ['Orders', 'Lines'], columns: [{ expr: '[ID] + 1' }] } }), /Orders and Lines both have ID: write \[Orders\]\.\[ID\]/);
    assert.equal(b.runQuery({ design: { tables: ['Orders', 'Lines'], columns: [{ expr: '[Orders].[ID] * 10' }] } }).rows[0].Expr1, 10);
    assert.throws(() => b.runQuery({ design: { tables: ['Orders'], columns: [{ expr: '[Nope]' }] } }), /No table in the query has a field called Nope/);
    assert.throws(() => b.runQuery({ design: { tables: ['Orders'], columns: [{ expr: '[Customers].[Country]' }] } }), /Customers is not in the query/);
  } finally {
    b.close();
  }
});

test('anything that is not an expression is refused in words, never run', () => {
  const resolve = (t, f) => `"${f}"`;
  for (const [src, why] of [
    ['1; DROP TABLE Orders', /";" .* is not part of an expression/],
    ['[Quantity] *', /ends too soon/],
    ['Sleep(5)', /Sleep is not a function Data Files knows/],
    ['Round([a], 2, 3)', /Round takes 1 or 2 values/],
    ['"open', /quotation mark .* is not closed/],
    ['#21/10/2026#', /Write a date as #2026-10-21#/],
    ['(1 + 2', /a \) is missing/],
    ['[a] [b]', /\[b\] was not expected there/],
    ['Format([d], "dddd")', /Format takes/],
  ]) assert.throws(() => expressionSql(src, resolve), why, src);
  // Text is written as SQL text, its quotes doubled: it can never close the literal early.
  assert.equal(expressionSql(`"it's" & "x"`, resolve), `(COALESCE('it''s', '') || COALESCE('x', ''))`);
  const b = shop();
  try {
    assert.match(b.sqlOf({ tables: ['Customers'], columns: [{ expr: '"Is it?"', as: 'Q' }, { table: 'Customers', field: 'Country', criteria: ['UK'] }] }), /\('Is it\?'\) AS "Q".*"Country" = 'UK'/s, 'a ? inside text is not taken for a parameter');
  } finally {
    b.close();
  }
});
