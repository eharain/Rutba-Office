#!/usr/bin/env node
// The sample files of the Data Files format (docs/DATA-FILES-FORMAT.md,
// section 12), made with Rutba Office's own database code so that they are
// exactly what it writes:
//
//   node tools/make-rbase-samples.mjs [folder]
//
// into docs/data-files-format/samples/ unless a folder is named. Each is
// left in rollback journal mode, one file with nothing beside it. The
// passwords of protected.rbase are examples, published with it.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DatabaseSync } from 'node:sqlite';
import { Base } from '../packages/database/src/index.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const SAMPLES_DIR = path.join(ROOT, 'docs', 'data-files-format', 'samples');

/** The example sign-ins of protected.rbase, as section 12 gives them. */
export const EXAMPLE_PEOPLE = [
  { name: 'Owner', role: 'Admin', password: 'example-password' },
  { name: 'Clerk', role: 'Staff', password: 'example-clerk' },
];

/** Two related tables with every field type, and a few invented records. */
function tablesAndRecords(base) {
  base.createTable({
    name: 'Customers',
    fields: [
      { name: 'Name', type: 'text', size: 80, required: true },
      { name: 'Email', type: 'text', unique: true },
      { name: 'Notes', type: 'longtext' },
      { name: 'Joined', type: 'date' },
      { name: 'Active', type: 'yesno', default: true },
      { name: 'Country', type: 'text', size: 2, default: 'GB' },
    ],
  });
  base.createTable({
    name: 'Orders',
    fields: [
      { name: 'Customer', type: 'integer', required: true },
      { name: 'Placed', type: 'datetime' },
      { name: 'Quantity', type: 'integer' },
      { name: 'Unit Price', type: 'currency' },
      { name: 'Weight', type: 'number', caption: 'Weight (kg)' },
      { name: 'Paid', type: 'yesno' },
    ],
  });
  base.addRelationship({ from: { table: 'Orders', field: 'Customer' }, to: { table: 'Customers', field: 'ID' } });
  const customers = [
    { Name: 'Example Stationers', Email: 'orders@example.com', Notes: 'Delivers on Mondays.', Joined: '2025-03-14', Country: 'GB' },
    { Name: 'Sample Bakery', Email: 'hello@example.org', Joined: '2026-01-05', Active: false, Country: 'PK' },
    { Name: 'Test Garden Centre', Email: 'shop@example.net', Joined: '2026-07-21' },
  ];
  const ids = customers.map((c) => base.insert('Customers', c));
  const orders = [
    [0, '2026-10-01T09:30:00', 12, 1.25, 0.4, true],
    [0, '2026-10-08T14:05:00', 3, 18.5, 2.2, false],
    [1, '2026-10-12T08:00:00', 40, 0.35, 6, true],
    [2, '2026-10-20T16:45:00', 1, 249.99, 31.5, false],
    [2, '2026-10-22T11:15:00', 6, 4.2, 1.8, true],
  ];
  for (const [c, placed, qty, price, weight, paid] of orders) {
    base.insert('Orders', { Customer: ids[c], Placed: placed, Quantity: qty, 'Unit Price': price, Weight: weight, Paid: paid });
  }
  base.setShown('Customers', 'Name');
}

/** A saved query with criteria, a calculation and totals; a form; a report; a print template. */
function designs(base) {
  base.saveQuery('Order Lines', {
    tables: ['Orders', 'Customers'],
    columns: [
      { table: 'Customers', field: 'Name', sort: 'asc' },
      { table: 'Orders', field: 'Placed', sort: 'desc' },
      { table: 'Orders', field: 'Quantity', criteria: ['> 0'] },
      { expr: 'Round([Quantity] * [Unit Price], 2)', as: 'Amount' },
      { table: 'Customers', field: 'Active', show: false, criteria: ['Yes'] },
    ],
  });
  base.saveQuery('Spend by Customer', {
    tables: ['Orders', 'Customers'],
    columns: [
      { table: 'Customers', field: 'Name', total: 'group', sort: 'asc' },
      { expr: '[Quantity] * [Unit Price]', as: 'Spend', total: 'sum', decimals: 2 },
      { table: 'Orders', field: 'ID', total: 'count', as: 'Orders' },
    ],
  });
  base.makeForm('Customers');
  base.makeForm('Orders');
  const report = base.makeReport({ table: 'Orders', name: 'Orders by Customer' });
  base.saveReport({ ...report, sort: 'Placed', desc: true, totals: [{ field: 'Quantity', fn: 'sum' }, { field: 'Unit Price', fn: 'avg' }] });
  base.makeReport({ query: 'Spend by Customer' });
  base.saveDocument({
    kind: 'document',
    name: 'Statement',
    title: 'Statement',
    table: 'Customers',
    paper: 'a4',
    blocks: [
      { type: 'heading', text: 'Statement for {Name}' },
      { type: 'fields', columns: 2, items: [{ field: 'ID', label: 'Account' }, { field: 'Email' }, { field: 'Joined', format: 'd mmmm yyyy' }] },
      { type: 'lines', table: 'Orders', link: 'Customer', sort: 'Placed', columns: [{ field: 'Placed', label: 'Date', format: 'dd/mm/yyyy' }, { field: 'Quantity' }, { field: 'Unit Price' }, { expr: 'Round([Quantity] * [Unit Price], 2)', as: 'Amount', type: 'currency' }] },
      { type: 'totals', rows: [{ label: 'Total', expr: '[Amount]', strong: true }] },
      { type: 'text', text: 'Thank you, {Name}.', small: true },
      { type: 'signatures', labels: ['Checked by'] },
    ],
  });
}

/** Protected, with the example people, a role of its own, and a trail. */
function protection(base) {
  const [owner, clerk] = EXAMPLE_PEOPLE;
  base.protect({ name: owner.name, password: owner.password });
  base.saveRole('Orders only', { tables: { default: 'hidden', each: { Orders: { access: 'add', own: true, fields: { 'Unit Price': 'read' } }, Customers: 'read' } }, hidden: { reports: ['Orders by Customer'] } });
  base.addPerson({ name: clerk.name, role: clerk.role, password: clerk.password });
  const id = base.insert('Customers', { Name: 'Added While Protected', Email: 'new@example.com' });
  base.update('Customers', id, { Notes: 'This change is in the trail.' });
}

/** A database made by `make` at `file`, then left as one self-contained file in rollback journal mode. */
function build(file, make) {
  for (const s of ['', '-wal', '-shm', '-journal']) fs.rmSync(file + s, { force: true });
  const base = Base.open(file);
  try { make(base); } finally { base.close(); }
  const db = new DatabaseSync(file);
  try {
    db.exec('PRAGMA journal_mode = DELETE');
    db.exec('VACUUM');
  } finally {
    db.close();
  }
  return file;
}

/** Make the three samples in `dir`. Answers their paths. */
export function makeSamples(dir = SAMPLES_DIR) {
  fs.mkdirSync(dir, { recursive: true });
  return [
    build(path.join(dir, 'plain-tables.rbase'), (b) => tablesAndRecords(b)),
    build(path.join(dir, 'query-form-report.rbase'), (b) => { tablesAndRecords(b); designs(b); }),
    build(path.join(dir, 'protected.rbase'), (b) => { tablesAndRecords(b); designs(b); protection(b); }),
  ];
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  for (const file of makeSamples(process.argv[2] ? path.resolve(process.argv[2]) : SAMPLES_DIR)) {
    process.stdout.write(`${path.relative(process.cwd(), file)}  ${fs.statSync(file).size} bytes\n`);
  }
}
