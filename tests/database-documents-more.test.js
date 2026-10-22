// Print templates that say when and what (the samples' list, 2026-10-22): a
// template for some records only (`where`), lines kept by a condition
// (`where`) and found by any field (`on`), blocks and totals rows shown only
// when they hold (`if`), a running balance, lines grouped with subtotals,
// values formatted and typed by the template, and an empty value as nothing.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Base, buildFromSpec } from '../packages/database/src/index.js';

const ref = (t) => ({ type: 'integer', points: t });

function shop() {
  const b = Base.open(':memory:');
  buildFromSpec(b, {
    tables: [
      { name: 'Customers', fields: [{ name: 'Name', type: 'text' }, { name: 'Opening', type: 'currency' }] },
      { name: 'Orders', fields: [{ name: 'Customer', ...ref('Customers') }, { name: 'Status', type: 'text' }, { name: 'Payment', type: 'text' }, { name: 'Tendered', type: 'currency' }, { name: 'Placed', type: 'datetime' }, { name: 'Gift', type: 'yesno' }] },
      { name: 'Order Lines', fields: [{ name: 'Order', ...ref('Orders') }, { name: 'Item', type: 'text' }, { name: 'Kind', type: 'text' }, { name: 'Amount', type: 'currency' }] },
      { name: 'Invoices', fields: [{ name: 'Customer', ...ref('Customers') }, { name: 'Number', type: 'text' }, { name: 'Amount', type: 'currency' }, { name: 'Paid', type: 'yesno' }] },
    ],
  });
  b.insert('Customers', { Name: 'Amira Khan', Opening: 100 });
  b.insert('Customers', { Name: 'Ben Ortiz', Opening: 0 });
  b.insert('Orders', { Customer: 1, Status: 'Shipped', Payment: 'Cash', Tendered: 50, Placed: '2026-10-21T09:30:00', Gift: true });
  b.insert('Orders', { Customer: 2, Status: 'Open', Payment: 'Card', Placed: '2026-10-21T10:00:00', Gift: false });
  b.insert('Orders', { Customer: 2, Status: 'Shipped', Payment: 'Card', Placed: '2026-10-22T11:15:00', Gift: false });
  for (const [o, item, kind, amount] of [[1, 'Bread', 'Bakery', 2.5], [1, 'Milk', 'Dairy', 1.2], [1, 'Buns', 'Bakery', 3], [3, 'Cheese', 'Dairy', 6]]) b.insert('Order Lines', { Order: o, Item: item, Kind: kind, Amount: amount });
  for (const [c, n, a, p] of [[1, 'INV-1', 40, true], [1, 'INV-2', 25, false], [1, 'INV-3', 10, false], [2, 'INV-4', 5, false]]) b.insert('Invoices', { Customer: c, Number: n, Amount: a, Paid: p });
  return b;
}
const block = (doc, type) => doc.blocks.find((x) => x.type === type);

test('a template for some records only, its lines kept by a condition and found by any field', () => {
  const b = shop();
  try {
    const notes = b.documentData({ layout: { name: 'Delivery note', table: 'Orders', where: '[Status] = "Shipped"', blocks: [{ type: 'heading', text: 'Delivery note' }] } });
    assert.deepEqual(notes.documents.map((d) => d.key), [1, 3]);
    // Asked for by its key, a record is shown whatever the condition says.
    assert.deepEqual(b.documentData({ layout: { name: 'x', table: 'Orders', where: '[Status] = "Shipped"', blocks: [] }, keys: [2] }).documents.map((d) => d.key), [2]);
    const statement = b.documentData({ layout: { name: 'Statement', table: 'Customers', blocks: [
      { type: 'lines', table: 'Invoices', where: 'Not [Paid]', columns: [{ field: 'Number' }, { field: 'Amount' }] },
      { type: 'totals', rows: [{ label: 'Owing', expr: '[Amount]', strong: true }] },
    ] } }).documents[0];
    assert.deepEqual(block(statement, 'lines').rows.map((r) => r.Number), ['INV-2', 'INV-3']);
    assert.equal(block(statement, 'totals').rows[0].value, 35);
    // Found by `on`: the customer's other orders, by the order's own customer.
    const others = b.documentData({ layout: { name: 'x', table: 'Orders', blocks: [{ type: 'lines', table: 'Orders', on: { Customer: 'Customer' }, columns: [{ field: 'Status' }] }] }, keys: [3] }).documents[0];
    assert.deepEqual(block(others, 'lines').rows.map((r) => r.Status), ['Open', 'Shipped']);
  } finally { b.close(); }
});

test('the template condition is applied before the limit: the matching records among the last of many are found', () => {
  const b = shop();
  try {
    for (let i = 0; i < 2500; i++) b.insert('Orders', { Customer: 1, Status: 'Open' });
    b.insert('Orders', { Customer: 2, Status: 'Late' });
    b.insert('Orders', { Customer: 2, Status: 'Late' });
    const late = b.documentData({ layout: { name: 'x', table: 'Orders', where: '[Status] = "Late"', blocks: [] }, limit: 1 });
    assert.deepEqual(late.documents.map((d) => d.key), [2504]);
    assert.deepEqual(b.documentData({ layout: { name: 'x', table: 'Orders', where: '[Status] = "Late"', blocks: [] }, limit: 5 }).documents.map((d) => d.key), [2504, 2505]);
  } finally { b.close(); }
});

test('blocks and totals rows shown only when they hold: cash tendered and change for cash only', () => {
  const b = shop();
  try {
    const layout = { name: 'Receipt', table: 'Orders', blocks: [
      { type: 'lines', table: 'Order Lines', columns: [{ field: 'Item' }, { field: 'Amount' }] },
      { type: 'totals', rows: [{ label: 'Total', expr: '[Amount]', strong: true }, { label: 'Tendered', expr: '[Tendered]', if: '[Payment] = "Cash"' }, { label: 'Change', expr: '[Tendered] - [Total]', if: '[Payment] = "Cash"' }] },
      { type: 'text', text: 'Thank you for paying by card.', if: '[Payment] = "Card"' },
    ] };
    const [cash, , card] = b.documentData({ layout }).documents;
    assert.deepEqual(block(cash, 'totals').rows.map((r) => [r.label, r.value]), [['Total', 6.7], ['Tendered', 50], ['Change', 43.3]]);
    assert.equal(block(cash, 'text'), undefined);
    assert.deepEqual(block(card, 'totals').rows.map((r) => r.label), ['Total']);
    assert.ok(block(card, 'text'));
  } finally { b.close(); }
});

test('a running balance from an opening balance, and lines grouped with subtotals', () => {
  const b = shop();
  try {
    const statement = b.documentData({ layout: { name: 'Statement', table: 'Customers', blocks: [
      { type: 'lines', table: 'Invoices', columns: [{ field: 'Number' }, { field: 'Amount' }, { running: 'Amount', as: 'Balance', start: '[Opening]' }] },
      { type: 'totals', rows: [{ label: 'Closing balance', expr: '[Balance]' }] },
    ] } }).documents[0];
    assert.deepEqual(block(statement, 'lines').rows.map((r) => r.Balance), [140, 165, 175]);
    assert.equal(block(statement, 'totals').rows[0].value, 175);
    const grouped = b.documentData({ layout: { name: 'Picking list', table: 'Orders', blocks: [
      { type: 'lines', table: 'Order Lines', group: 'Kind', columns: [{ field: 'Kind' }, { field: 'Item' }, { field: 'Amount' }] },
    ] }, keys: [1] }).documents[0];
    const rows = block(grouped, 'lines').rows;
    assert.deepEqual(rows.map((r) => (r._group !== undefined ? `# ${r._group}` : r._subtotal ? `= ${r._subtotal.Amount}` : r.Item)), ['# Bakery', 'Bread', 'Buns', '= 5.5', '# Dairy', 'Milk', '= 1.2']);
    // Subtotals of the columns named, only.
    const named = b.documentData({ layout: { name: 'x', table: 'Orders', blocks: [{ type: 'lines', table: 'Order Lines', group: 'Kind', subtotals: ['Amount'], columns: [{ field: 'Kind' }, { expr: '1', as: 'One', type: 'integer' }, { field: 'Amount' }] }] }, keys: [1] }).documents[0];
    assert.deepEqual(block(named, 'lines').rows.find((r) => r._subtotal)._subtotal, { Amount: 5.5 });
  } finally { b.close(); }
});

test('values typed and formatted by the template, a date and time read from its value, empty as nothing', () => {
  const b = shop();
  try {
    const doc = b.documentData({ layout: { name: 'x', table: 'Orders', blocks: [
      { type: 'fields', items: [{ field: 'Gift' }, { field: 'Placed' }, { field: 'Placed', label: 'Day', format: 'dd/mm/yyyy' }, { field: 'Tendered', format: '#,##0.00' }, { field: 'Tendered', label: 'Change due', if: '[Payment] = "Cash"' }] },
      { type: 'lines', table: 'Order Lines', columns: [{ field: 'Item' }, { expr: '[Placed]', as: 'When' }, { field: 'Amount', format: '0.00' }] },
    ] }, keys: [1, 2] }).documents;
    const [first, second] = doc;
    const items = Object.fromEntries(block(first, 'fields').items.map((i) => [i.label, i]));
    assert.equal(items.Gift.type, 'yesno');
    assert.equal(items.Placed.type, 'datetime');
    assert.equal(items.Day.value, '21/10/2026');
    assert.equal(items.Tendered.value, '50.00');
    assert.ok(items['Change due']);
    // Not cash: the item that says so is left out; an empty value stays empty, not "null".
    assert.ok(!block(second, 'fields').items.some((i) => i.label === 'Change due'));
    assert.equal(Object.fromEntries(block(second, 'fields').items.map((i) => [i.label, i.value])).Tendered, null);
    const lines = block(first, 'lines');
    assert.equal(lines.columns.find((c) => c.name === 'When').type, 'datetime');
    assert.equal(lines.rows[0].Amount, '2.50');
  } finally { b.close(); }
});
