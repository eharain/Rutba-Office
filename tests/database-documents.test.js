// Data Files' print templates (packages/database/src/document-design.js): an
// invoice for each record with the business's letterhead, fields followed
// through relationships, lines with calculated columns, totals built on the
// lines and on each other, words with blanks filled in, and signatures; a
// till receipt; a rent agreement made of words; the wizard's template; kept
// in the file and made from a spec; and anything wrong refused in words.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { Base, buildFromSpec, evaluate, DESIGN_VERSION } from '../packages/database/src/index.js';

const ref = (t) => ({ type: 'integer', points: t });
const INVOICE = {
  name: 'Invoice', table: 'Invoices', paper: 'a4',
  letterhead: { table: 'Business Details', fields: ['Name', 'Address', 'Tax Number'] },
  blocks: [
    { type: 'heading', text: 'Invoice' },
    { type: 'fields', columns: 2, items: [{ field: 'Number', label: 'Invoice' }, { field: 'Invoice Date' }, { field: 'Customer', label: 'Bill to' }, { field: 'Customer.Address', label: 'Address' }] },
    { type: 'lines', table: 'Invoice Lines', columns: [{ field: 'Product' }, { field: 'Quantity' }, { field: 'Unit Price' }, { expr: 'Round([Quantity] * [Unit Price], 2)', as: 'Amount', type: 'currency' }] },
    { type: 'totals', rows: [{ label: 'Subtotal', expr: '[Amount]' }, { label: 'VAT', expr: 'Round([Subtotal] * [Tax Rate] / 100, 2)' }, { label: 'Total', expr: '[Subtotal] + [VAT]', strong: true }] },
    { type: 'text', text: 'Please pay within {Customer.Payment Days} days. Thank you.' },
    { type: 'signatures', labels: ['Received by'] },
  ],
};

function shop(file = ':memory:') {
  const b = Base.open(file);
  buildFromSpec(b, {
    tables: [
      { name: 'Business Details', fields: [{ name: 'Name', type: 'text' }, { name: 'Address', type: 'text' }, { name: 'Tax Number', type: 'text' }, { name: 'Tax Rate', type: 'number' }] },
      { name: 'Customers', fields: [{ name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text' }, { name: 'Address', type: 'text' }, { name: 'Payment Days', type: 'integer' }] },
      { name: 'Products', fields: [{ name: 'Name', type: 'text' }] },
      { name: 'Invoices', fields: [{ name: 'Number', type: 'text' }, { name: 'Customer', ...ref('Customers') }, { name: 'Invoice Date', type: 'date' }] },
      { name: 'Invoice Lines', fields: [{ name: 'Invoice', ...ref('Invoices') }, { name: 'Product', ...ref('Products') }, { name: 'Quantity', type: 'integer' }, { name: 'Unit Price', type: 'currency' }] },
    ],
    documents: [INVOICE],
  });
  b.insert('Business Details', { Name: 'Larkspur Repairs', Address: '4 Mill Lane, Leeds', 'Tax Number': 'GB 000 0000 00', 'Tax Rate': 20 });
  b.insert('Customers', { 'First Name': 'Amira', 'Last Name': 'Haddad', Address: '12 Orchard Way', 'Payment Days': 14 });
  b.insert('Customers', { 'First Name': 'Ben', 'Last Name': 'Cole', Address: '3 Canal Walk', 'Payment Days': 30 });
  b.insert('Products', { Name: 'Screen repair' });
  b.insert('Products', { Name: 'Phone case' });
  b.insert('Invoices', { Number: 'INV-00001', Customer: 1, 'Invoice Date': '2026-10-21' });
  b.insert('Invoices', { Number: 'INV-00002', Customer: 2, 'Invoice Date': '2026-10-22' });
  b.insert('Invoice Lines', { Invoice: 1, Product: 1, Quantity: 1, 'Unit Price': 89 });
  b.insert('Invoice Lines', { Invoice: 1, Product: 2, Quantity: 2, 'Unit Price': 7.5 });
  b.insert('Invoice Lines', { Invoice: 2, Product: 2, Quantity: 3, 'Unit Price': 7.5 });
  return b;
}
const block = (doc, type) => doc.blocks.find((x) => x.type === type);

test('an invoice for each record: the letterhead, fields through relationships, lines, totals on totals, words filled in', () => {
  const b = shop();
  try {
    const d = b.documentData({ name: 'Invoice' });
    assert.equal(d.documents.length, 2);
    assert.deepEqual(d.letterhead.fields.map((f) => f.value), ['Larkspur Repairs', '4 Mill Lane, Leeds', 'GB 000 0000 00']);
    const [first, second] = d.documents;
    assert.deepEqual(block(first, 'fields').items.map((i) => [i.label, i.value]), [['Invoice', 'INV-00001'], ['Invoice Date', '2026-10-21'], ['Bill to', 'Amira Haddad'], ['Address', '12 Orchard Way']]);
    assert.deepEqual(block(first, 'lines').rows, [{ Product: 'Screen repair', Quantity: 1, 'Unit Price': 89, Amount: 89 }, { Product: 'Phone case', Quantity: 2, 'Unit Price': 7.5, Amount: 15 }]);
    assert.deepEqual(block(first, 'totals').rows.map((r) => [r.label, r.value]), [['Subtotal', 104], ['VAT', 20.8], ['Total', 124.8]]);
    assert.equal(block(first, 'text').parts.map((p) => p.text ?? p.value).join(''), 'Please pay within 14 days. Thank you.');
    assert.deepEqual(block(second, 'totals').rows.map((r) => r.value), [22.5, 4.5, 27]);
    assert.deepEqual(block(first, 'signatures').labels, ['Received by']);
    // One record alone: the invoice a person is looking at.
    assert.deepEqual(b.documentData({ name: 'Invoice', keys: [2] }).documents.map((x) => x.key), [2]);
  } finally {
    b.close();
  }
});

test('a till receipt on a narrow roll, and a rent agreement made of words with blanks and two signatures', () => {
  const b = shop();
  try {
    b.saveDocument({ name: 'Receipt', table: 'Invoices', paper: 'receipt-80', blocks: [
      { type: 'heading', text: 'Receipt {Number}' },
      { type: 'lines', table: 'Invoice Lines', columns: [{ field: 'Product' }, { expr: '[Quantity] * [Unit Price]', as: 'Price', type: 'currency' }] },
      { type: 'totals', rows: [{ label: 'Total', expr: '[Price]', strong: true }] },
    ] });
    const r = b.documentData({ name: 'Receipt', keys: [1] }).documents[0];
    assert.equal(r.blocks[0].parts.map((p) => p.text ?? p.value).join(''), 'Receipt INV-00001');
    assert.equal(block(r, 'totals').rows[0].value, 104);
    b.saveDocument({ name: 'Agreement', table: 'Customers', paper: 'letter', blocks: [
      { type: 'heading', text: 'Agreement' },
      { type: 'text', text: 'This agreement is between {Name} of {Address} and {First Name} {Last Name} of {Customers.Address}.' },
      { type: 'spacer', size: 24 },
      { type: 'signatures', labels: ['Landlord', 'Tenant'] },
    ], letterhead: { table: 'Business Details' } });
    const a = b.documentData({ name: 'Agreement', keys: [2] }).documents[0];
    // {Address} alone is the record's own; {Name}, which the record has not, is the letterhead's.
    assert.equal(block(a, 'text').parts.map((p) => p.text ?? p.value).join(''), 'This agreement is between Larkspur Repairs of 3 Canal Walk and Ben Cole of 3 Canal Walk.');
    b.saveDocument({ name: 'Agreement 2', table: 'Customers', letterhead: { table: 'Business Details' }, blocks: [{ type: 'text', text: '{Business Details.Name} of {Business Details.Address}, and {Customers.First Name}.' }] });
    assert.equal(b.documentData({ name: 'Agreement 2', keys: [1] }).documents[0].blocks[0].parts.map((p) => p.text ?? p.value).join(''), 'Larkspur Repairs of 4 Mill Lane, Leeds, and Amira.');
    assert.deepEqual(block(a, 'signatures').labels, ['Landlord', 'Tenant']);
  } finally {
    b.close();
  }
});

test('the wizard makes an invoice-like template from a table with lines, and templates are kept and read back', () => {
  const file = path.join(os.tmpdir(), `docs-${process.pid}.rbase`);
  try {
    const b = shop(file);
    const w = b.makeDocument({ table: 'Invoices' });
    assert.equal(w.name, 'Invoice 2');
    assert.equal(w.blocks[0].text, 'Invoice');
    assert.deepEqual(w.letterhead, { table: 'Business Details' });
    const lines = w.blocks.find((x) => x.type === 'lines');
    assert.equal(lines.table, 'Invoice Lines');
    assert.equal(lines.columns.at(-1).expr, 'Round([Quantity] * [Unit Price], 2)');
    assert.equal(block(b.documentData({ name: 'Invoice 2', keys: [1] }).documents[0], 'totals').rows[0].value, 104);
    assert.ok(b.removeDocument('Invoice 2'));
    b.close();
    const again = Base.open(file);
    assert.deepEqual(again.documents().map((d) => d.name), ['Invoice']);
    again.close();
  } finally {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  }
});

test('a template that cannot be printed is refused in words', () => {
  const b = shop();
  try {
    assert.throws(() => b.saveDocument({ name: 'X', table: 'Nowhere', blocks: [] }), /table, Nowhere, is not in this database/);
    assert.throws(() => b.saveDocument({ name: 'X', table: 'Invoices', paper: 'a3', blocks: [] }), /prints on a4, letter, a5/);
    assert.throws(() => b.saveDocument({ name: 'X', table: 'Invoices', blocks: [{ type: 'video' }] }), /no block called video/);
    b.saveDocument({ name: 'Bad', table: 'Invoices', blocks: [{ type: 'fields', items: [{ field: 'Customer.Shoe Size' }] }] });
    assert.throws(() => b.documentData({ name: 'Bad' }), /Customers has no field called Shoe Size/);
    assert.throws(() => evaluate(b, '[Nope] + 1', { A: 1 }), /\[Nope\] is not a field of this document/);
    assert.throws(() => b.documentData({ name: 'Missing' }), /no print template called Missing/);
  } finally {
    b.close();
  }
});

test('a file made before print templates opens with their table, its version brought up', () => {
  const file = path.join(os.tmpdir(), `v3-${process.pid}.rbase`);
  try {
    const old = new DatabaseSync(file);
    old.exec("CREATE TABLE rutba_meta (key TEXT PRIMARY KEY, value TEXT); INSERT INTO rutba_meta VALUES ('version', '3');");
    old.close();
    const b = Base.open(file);
    assert.equal(DESIGN_VERSION, 5);
    assert.equal(b.meta('version'), '5');
    assert.deepEqual(b.documents(), []);
    b.close();
  } finally {
    for (const s of ['', '-wal', '-shm']) fs.rmSync(file + s, { force: true });
  }
});

test('a lines column naming a field that points at another table shows that record\'s words, as words', () => {
  const b = shop();
  try {
    const lines = block(b.documentData({ name: 'Invoice' }).documents[0], 'lines');
    const product = lines.columns.find((c) => c.name === 'Product');
    assert.equal(product.type, 'text');
    assert.equal(product.align, 'left');
    assert.deepEqual(lines.rows.map((r) => r.Product), ['Screen repair', 'Phone case']);
    // Quantity holds whole numbers and stays a number.
    assert.equal(lines.columns.find((c) => c.name === 'Quantity').type, 'integer');
  } finally {
    b.close();
  }
});
