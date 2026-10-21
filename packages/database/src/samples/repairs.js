// The repair shop: Ironbridge Repairs, an invented shop that mends phones,
// tablets, laptops and small appliances, with three technicians at the bench
// and a counter where work is booked in and collected. Customers and the
// devices they bring, the businesses that have an account, the staff, every
// job from booking in to collection, the parts on the shelf and those ordered
// in, the parts and labour each job used, quotes, invoices, warranties and
// what was paid.
//
// Its records come from fifteen months of work, run a day at a time: jobs
// are booked in, diagnosed, quoted and approved or turned down, wait for a
// part when the shelf has none, are repaired and then collected (most within
// days, a few left for weeks). Phones come in more in the summer holidays,
// laptops when term starts, kitchen appliances before the year ends. A few
// repairs come back under warranty. The questions a repair shop asks at the
// counter and at the bench are its saved queries; its reports print them.
//
// It prints what a repair shop hands over the counter: a booking-in receipt
// for every job, a written quote after diagnosis, an invoice with the parts
// and labour and the VAT in them, a receipt when the device is collected, a
// 90-day warranty certificate, and a statement for each business account.
// Prices are shown to the public with VAT included, so each line keeps its
// rate and the VAT is worked out from inside the price, line by line.

import { BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, companies, money, person, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const CUSTOMER = '[Customers].[First Name] & " " & [Customers].[Last Name]';
const TECHNICIAN = '[Staff].[First Name] & " " & [Staff].[Last Name]';
const DEVICE = '[Make] & " " & [Model]';
// The job board's order, as the shop works through it, rather than A to Z.
const STAGE = 'IIf([Status] = "Booked in", "1 Booked in", IIf([Status] = "Diagnosing", "2 Diagnosing", IIf([Status] = "Awaiting approval", "3 Awaiting approval", '
  + 'IIf([Status] = "Awaiting parts", "4 Awaiting parts", IIf([Status] = "Repairing", "5 Repairing", IIf([Status] = "Ready", "6 Ready", "7 " & [Status]))))))';
// A line's amount, VAT included, and the VAT inside it: each line rounded, then added up.
const GROSS = 'Round([Quantity] * [Unit Price], 2)';
const VAT_IN = `Round(${GROSS} - ${GROSS} / (1 + [Tax Rate]), 2)`;
const OVERDUE = 'IIf(DateDiff("d", [Due Date], Date()) > 0, DateDiff("d", [Due Date], Date()), 0)';
// What every document closes with: the business's terms and its footer.
const FOOT = { type: 'text', text: '{Business.Document Footer}', small: true };

export const spec = {
  // Where a value below nought is right: a job already past its promised
  // date, and a part that was fitted free under warranty more than it was sold.
  negative: { 'Job board': ['Days Left'], 'Parts margin': ['Margin'] },
  tables: [
    BUSINESS_TABLE,
    { name: 'Accounts', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Contact', type: 'text' }, { name: 'Phone', type: 'text' },
      { name: 'E-mail', type: 'text' }, { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' },
      { name: 'Payment Days', type: 'integer', default: 30 }, { name: 'Credit Limit', type: 'currency' }, { name: 'Opened', type: 'date' },
    ] },
    { name: 'Customers', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Account', ...ref('Accounts') },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Address', type: 'text' }, { name: 'City', type: 'text' },
      { name: 'Postcode', type: 'text' }, { name: 'Contact By', type: 'text', default: 'Text message' }, { name: 'Joined', type: 'date' },
    ] },
    { name: 'Staff', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Role', type: 'text' },
      { name: 'Skills', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Started', type: 'date' },
      { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Suppliers', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Contact', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' },
      { name: 'Supplies', type: 'text' }, { name: 'Lead Days', type: 'integer', default: 3 }, { name: 'Our Account', type: 'text' },
    ] },
    { name: 'Parts', fields: [
      { name: 'Code', type: 'text', unique: true }, { name: 'Name', type: 'text', required: true }, { name: 'Fits', type: 'text' },
      { name: 'Supplier', ...ref('Suppliers') }, { name: 'Cost', type: 'currency' }, { name: 'Price', type: 'currency' },
      { name: 'In Stock', type: 'integer', default: 0 }, { name: 'Reorder Level', type: 'integer', default: 2 },
      { name: 'Reorder Quantity', type: 'integer', default: 4 }, { name: 'Stocked', type: 'yesno', default: true }, { name: 'Bin', type: 'text' },
    ] },
    { name: 'Devices', fields: [
      { name: 'Customer', ...ref('Customers'), required: true }, { name: 'Kind', type: 'text' }, { name: 'Make', type: 'text' },
      { name: 'Model', type: 'text' }, { name: 'Serial', type: 'text' }, { name: 'Colour', type: 'text' },
    ] },
    // Number is the first text field, so a document that names a job alone prints its number.
    { name: 'Jobs', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Device', ...ref('Devices'), required: true }, { name: 'Technician', ...ref('Staff') },
      { name: 'Booked In', type: 'date' }, { name: 'Fault', type: 'text' }, { name: 'Left With It', type: 'text' }, { name: 'Promised', type: 'date' },
      { name: 'Status', type: 'text', default: 'Booked in' }, { name: 'Quote', type: 'currency' }, { name: 'Quoted On', type: 'date' },
      { name: 'Quote Approved', type: 'yesno' }, { name: 'Deposit', type: 'currency' }, { name: 'Outcome', type: 'text' },
      { name: 'Charge', type: 'currency' }, { name: 'Finished', type: 'date' }, { name: 'Collected', type: 'date' }, { name: 'Paid', type: 'yesno' },
      { name: 'Warranty Until', type: 'date' }, { name: 'Warranty Of', ...ref('Jobs') }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Part Orders', fields: [
      { name: 'Supplier', ...ref('Suppliers'), required: true }, { name: 'Part', ...ref('Parts'), required: true }, { name: 'Job', ...ref('Jobs') },
      { name: 'Quantity', type: 'integer', default: 1 }, { name: 'Unit Cost', type: 'currency' }, { name: 'Ordered On', type: 'date' },
      { name: 'Expected', type: 'date' }, { name: 'Received On', type: 'date' }, { name: 'Status', type: 'text', default: 'On order' },
    ] },
    // The parts and labour on each job, one line each; once invoiced, a line also points at its invoice.
    { name: 'Job Lines', fields: [
      { name: 'Job', ...ref('Jobs'), required: true }, { name: 'Invoice', ...ref('Invoices') }, { name: 'Kind', type: 'text', default: 'Part' },
      { name: 'Part', ...ref('Parts') }, { name: 'Description', type: 'text' }, { name: 'Quantity', type: 'number', default: 1 },
      { name: 'Unit Price', type: 'currency' }, { name: 'Tax Rate', type: 'number', default: 0.2 }, { name: 'Done On', type: 'date' },
    ] },
    { name: 'Quotes', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Job', ...ref('Jobs'), required: true }, { name: 'Quoted On', type: 'date' },
      { name: 'Valid Until', type: 'date' }, { name: 'Net', type: 'currency' }, { name: 'Tax', type: 'currency' }, { name: 'Total', type: 'currency' },
      { name: 'Answer', type: 'text' }, { name: 'Answered On', type: 'date' },
    ] },
    { name: 'Quote Lines', fields: [
      { name: 'Quote', ...ref('Quotes'), required: true }, { name: 'Kind', type: 'text', default: 'Part' }, { name: 'Description', type: 'text' },
      { name: 'Quantity', type: 'number', default: 1 }, { name: 'Unit Price', type: 'currency' }, { name: 'Tax Rate', type: 'number', default: 0.2 },
    ] },
    { name: 'Invoices', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Job', ...ref('Jobs'), required: true }, { name: 'Account', ...ref('Accounts') },
      { name: 'Invoice Date', type: 'date' }, { name: 'Due Date', type: 'date' }, { name: 'Net', type: 'currency' }, { name: 'Tax', type: 'currency' },
      { name: 'Total', type: 'currency' }, { name: 'Deposit', type: 'currency', default: 0 }, { name: 'Balance', type: 'currency' }, { name: 'Paid On', type: 'date' },
    ] },
    { name: 'Warranties', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Job', ...ref('Jobs'), required: true }, { name: 'Issued', type: 'date' },
      { name: 'Covers', type: 'text' },
    ] },
    { name: 'Payments', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Job', ...ref('Jobs'), required: true }, { name: 'Paid On', type: 'date' },
      { name: 'Amount', type: 'currency' }, { name: 'Method', type: 'text', default: 'Card' }, { name: 'Kind', type: 'text', default: 'On collection' },
    ] },
  ],
  queries: {
    'Job board': { tables: ['Jobs', 'Devices', 'Customers', 'Staff'], columns: [
      { expr: STAGE, as: 'Stage', sort: 'asc' }, { table: 'Jobs', field: 'Number', as: 'Job' }, { table: 'Jobs', field: 'Booked In' },
      { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' }, { expr: DEVICE, as: 'Device' }, { table: 'Jobs', field: 'Fault' },
      { expr: TECHNICIAN, as: 'Technician' }, { table: 'Jobs', field: 'Promised', sort: 'asc' },
      { expr: 'IIf(Nz([Finished], "") = "", DateDiff("d", Date(), [Promised]), Null)', as: 'Days Left' },
      { table: 'Jobs', field: 'Collected', criteria: ['Is Null'], show: false },
    ] },
    'Jobs by status': { tables: ['Jobs'], columns: [
      { expr: STAGE, as: 'Stage', total: 'group', sort: 'asc' }, { table: 'Jobs', field: 'ID', as: 'Jobs', total: 'count' },
      { table: 'Jobs', field: 'Booked In', as: 'Oldest Booked In', total: 'min' },
      { table: 'Jobs', field: 'Collected', total: 'where', criteria: ['Is Null'] },
    ] },
    'Awaiting parts': { tables: ['Part Orders', 'Parts', 'Suppliers', 'Jobs', 'Devices', 'Customers'], columns: [
      { table: 'Jobs', field: 'Number', as: 'Job' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' }, { expr: DEVICE, as: 'Device' },
      { table: 'Parts', field: 'Name', as: 'Part' }, { table: 'Suppliers', field: 'Name', as: 'Supplier' }, { table: 'Part Orders', field: 'Ordered On' },
      { table: 'Part Orders', field: 'Expected', sort: 'asc' }, { table: 'Jobs', field: 'Promised' },
      { table: 'Part Orders', field: 'Received On', criteria: ['Is Null'], show: false }, { table: 'Jobs', field: 'Status', criteria: ['Awaiting parts'], show: false },
    ] },
    'Past promised date': { tables: ['Jobs', 'Devices', 'Customers', 'Staff'], columns: [
      { table: 'Jobs', field: 'Number', as: 'Job' }, { table: 'Jobs', field: 'Status' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' },
      { expr: DEVICE, as: 'Device' }, { expr: TECHNICIAN, as: 'Technician' }, { table: 'Jobs', field: 'Promised' },
      { expr: 'DateDiff("d", [Promised], Date())', as: 'Days Late', sort: 'desc', criteria: ['> 0'] },
      { table: 'Jobs', field: 'Finished', criteria: ['Is Null'], show: false },
    ] },
    'Ready for collection': { tables: ['Jobs', 'Devices', 'Customers'], columns: [
      { table: 'Jobs', field: 'Number', as: 'Job' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' },
      { table: 'Customers', field: 'Contact By' }, { expr: DEVICE, as: 'Device' }, { table: 'Jobs', field: 'Outcome' }, { table: 'Jobs', field: 'Finished' },
      { expr: 'DateDiff("d", [Finished], Date())', as: 'Days Waiting', sort: 'desc' }, { expr: 'Round(Nz([Charge]) - Nz([Deposit]), 2)', as: 'To Pay' },
      { table: 'Jobs', field: 'Finished', criteria: ['Is Not Null'], show: false }, { table: 'Jobs', field: 'Collected', criteria: ['Is Null'], show: false },
    ] },
    'Uncollected 14 days or more': { tables: ['Jobs', 'Devices', 'Customers'], columns: [
      { table: 'Jobs', field: 'Number', as: 'Job' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' }, { table: 'Customers', field: 'E-mail' },
      { expr: DEVICE, as: 'Device' }, { table: 'Jobs', field: 'Finished' },
      { expr: 'DateDiff("d", [Finished], Date())', as: 'Days Waiting', sort: 'desc', criteria: ['>= 14'] }, { expr: 'Round(Nz([Charge]) - Nz([Deposit]), 2)', as: 'To Pay' },
      { table: 'Jobs', field: 'Notes' }, { table: 'Jobs', field: 'Collected', criteria: ['Is Null'], show: false },
    ] },
    'Quotes awaiting approval': { tables: ['Jobs', 'Devices', 'Customers'], columns: [
      { table: 'Jobs', field: 'Number', as: 'Job' }, { table: 'Jobs', field: 'Quoted On', sort: 'asc' },
      { expr: 'DateDiff("d", [Quoted On], Date())', as: 'Days Waiting' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' },
      { expr: DEVICE, as: 'Device' }, { table: 'Jobs', field: 'Fault' }, { table: 'Jobs', field: 'Quote' },
      { table: 'Jobs', field: 'Status', criteria: ['Awaiting approval'], show: false },
    ] },
    'Parts to reorder': { tables: ['Parts', 'Suppliers'], columns: [
      { table: 'Suppliers', field: 'Name', as: 'Supplier', sort: 'asc' }, { table: 'Parts', field: 'Code', sort: 'asc' }, { table: 'Parts', field: 'Name', as: 'Part' },
      { table: 'Parts', field: 'In Stock' }, { table: 'Parts', field: 'Reorder Level' }, { table: 'Parts', field: 'Reorder Quantity' },
      { table: 'Parts', field: 'Cost' }, { expr: 'Round([Reorder Quantity] * [Cost], 2)', as: 'Order Value' },
      { expr: '[In Stock] - [Reorder Level]', as: 'Spare', criteria: ['<= 0'], show: false }, { table: 'Parts', field: 'Stocked', criteria: ['Yes'], show: false },
    ] },
    'Parts on order': { tables: ['Part Orders', 'Parts', 'Suppliers'], columns: [
      { table: 'Suppliers', field: 'Name', as: 'Supplier', sort: 'asc' }, { table: 'Parts', field: 'Name', as: 'Part' }, { table: 'Part Orders', field: 'Quantity' },
      { table: 'Part Orders', field: 'Unit Cost' }, { table: 'Part Orders', field: 'Ordered On' }, { table: 'Part Orders', field: 'Expected', sort: 'asc' },
      { table: 'Part Orders', field: 'Job', as: 'For Job' }, { table: 'Part Orders', field: 'Received On', criteria: ['Is Null'], show: false },
    ] },
    'Takings by month': { tables: ['Payments'], columns: [
      { expr: 'Format([Paid On], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Payments', field: 'ID', as: 'Payments', total: 'count' },
      { table: 'Payments', field: 'Amount', as: 'Takings', total: 'sum' },
    ] },
    'Takings last 30 days': { tables: ['Payments'], columns: [
      { table: 'Payments', field: 'Method', total: 'group', sort: 'asc' }, { table: 'Payments', field: 'ID', as: 'Payments', total: 'count' },
      { table: 'Payments', field: 'Amount', as: 'Takings', total: 'sum' },
      { expr: 'DateDiff("d", [Paid On], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 30'] },
    ] },
    'Technician workload': { tables: ['Jobs', 'Staff'], columns: [
      { expr: TECHNICIAN, as: 'Technician', total: 'group', sort: 'asc' }, { table: 'Jobs', field: 'ID', as: 'Open Jobs', total: 'count' },
      { expr: 'IIf([Status] = "Awaiting parts", 1, 0)', as: 'Awaiting Parts', total: 'sum' },
      { expr: 'IIf([Status] = "Awaiting approval", 1, 0)', as: 'Awaiting Approval', total: 'sum' },
      { table: 'Jobs', field: 'Booked In', as: 'Oldest Booked In', total: 'min' }, { table: 'Jobs', field: 'Finished', total: 'where', criteria: ['Is Null'] },
    ] },
    'Technician turnaround': { tables: ['Jobs', 'Staff'], columns: [
      { expr: TECHNICIAN, as: 'Technician', total: 'group', sort: 'asc' }, { table: 'Jobs', field: 'ID', as: 'Jobs Repaired', total: 'count' },
      { expr: 'DateDiff("d", [Booked In], [Finished])', as: 'Average Days', total: 'avg' }, { table: 'Jobs', field: 'Charge', as: 'Charged', total: 'sum' },
      { table: 'Jobs', field: 'Outcome', total: 'where', criteria: ['Repaired'] },
      { expr: 'DateDiff("d", [Finished], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Work by device kind': { tables: ['Jobs', 'Devices'], columns: [
      { table: 'Devices', field: 'Kind', total: 'group' }, { table: 'Jobs', field: 'ID', as: 'Jobs', total: 'count', sort: 'desc' },
      { table: 'Jobs', field: 'Charge', as: 'Charged', total: 'sum' }, { expr: 'DateDiff("d", [Booked In], [Finished])', as: 'Average Days', total: 'avg' },
    ] },
    'Warranty returns': { tables: ['Jobs', 'Devices', 'Customers', 'Staff'], columns: [
      { table: 'Jobs', field: 'Number', as: 'Job' }, { table: 'Jobs', field: 'Booked In', sort: 'desc' },
      { expr: CUSTOMER, as: 'Customer' }, { expr: DEVICE, as: 'Device' }, { table: 'Jobs', field: 'Fault' }, { expr: TECHNICIAN, as: 'Technician' },
      { table: 'Jobs', field: 'Status' }, { table: 'Jobs', field: 'Warranty Of', criteria: ['Is Not Null'], show: false },
    ] },
    'Job sheets': { tables: ['Jobs', 'Devices', 'Customers', 'Staff'], columns: [
      { expr: TECHNICIAN, as: 'Technician', sort: 'asc' }, { table: 'Jobs', field: 'Number', as: 'Job', sort: 'asc' }, { table: 'Jobs', field: 'Booked In' },
      { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' }, { expr: DEVICE, as: 'Device' }, { table: 'Devices', field: 'Serial' },
      { table: 'Jobs', field: 'Fault' }, { table: 'Jobs', field: 'Left With It' }, { table: 'Jobs', field: 'Status' }, { table: 'Jobs', field: 'Quote' },
      { table: 'Jobs', field: 'Promised' }, { expr: 'DateDiff("d", [Booked In], Date())', as: 'Days Ago', criteria: ['Between 0 And 14'], show: false },
    ] },
    // Charged is what the customer paid, VAT included; the margin is on the price without its VAT.
    'Parts margin': { tables: ['Job Lines', 'Parts'], columns: [
      { table: 'Parts', field: 'Fits', total: 'group', sort: 'asc' }, { table: 'Parts', field: 'Name', as: 'Part', total: 'group', sort: 'asc' },
      { table: 'Job Lines', field: 'Quantity', as: 'Fitted', total: 'sum' }, { expr: GROSS, as: 'Charged', total: 'sum' },
      { expr: 'Round([Quantity] * ([Unit Price] / (1 + [Tax Rate]) - [Cost]), 2)', as: 'Margin', total: 'sum' },
      { table: 'Job Lines', field: 'Done On', total: 'where', criteria: ['Is Not Null'] },
    ] },
    'Accounts to chase': { tables: ['Invoices', 'Accounts', 'Jobs'], columns: [
      { table: 'Accounts', field: 'Name', as: 'Account', sort: 'asc' }, { table: 'Accounts', field: 'Contact' }, { table: 'Accounts', field: 'E-mail' },
      { table: 'Invoices', field: 'Number', as: 'Invoice' }, { table: 'Jobs', field: 'Number', as: 'Job' }, { table: 'Invoices', field: 'Invoice Date' },
      { table: 'Invoices', field: 'Due Date', sort: 'asc' }, { table: 'Invoices', field: 'Balance' }, { expr: OVERDUE, as: 'Days Overdue' },
      { table: 'Invoices', field: 'Paid On', criteria: ['Is Null'], show: false },
    ] },
    'Under warranty': { tables: ['Warranties', 'Jobs', 'Devices', 'Customers'], columns: [
      { table: 'Warranties', field: 'Number', as: 'Certificate' }, { table: 'Jobs', field: 'Number', as: 'Job' }, { expr: CUSTOMER, as: 'Customer' },
      { table: 'Customers', field: 'Phone' }, { expr: DEVICE, as: 'Device' }, { table: 'Devices', field: 'Serial' }, { table: 'Warranties', field: 'Covers' },
      { table: 'Jobs', field: 'Warranty Until', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Warranty Until])', as: 'Days Left', criteria: ['>= 0'] },
    ] },
    'VAT by month': { tables: ['Invoices'], columns: [
      { expr: 'Format([Invoice Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Invoices', field: 'ID', as: 'Invoices', total: 'count' },
      { table: 'Invoices', field: 'Net', as: 'Net Sales', total: 'sum' }, { table: 'Invoices', field: 'Tax', as: 'VAT', total: 'sum' },
      { table: 'Invoices', field: 'Total', as: 'Invoiced', total: 'sum' },
    ] },
  },
  reports: [
    { name: 'Job board', query: 'Job board', group: 'Stage', sort: 'Promised', columns: ['Job', { field: 'Booked In', type: 'date' }, 'Customer', 'Device', 'Fault', 'Technician', { field: 'Promised', type: 'date' }, { field: 'Days Left', type: 'integer' }], totals: [{ field: 'Job', fn: 'count' }] },
    { name: 'Ready for collection', title: 'Ready for collection: to phone', query: 'Ready for collection', sort: 'Days Waiting', desc: true, columns: ['Job', 'Customer', 'Phone', 'Contact By', 'Device', 'Outcome', { field: 'Finished', type: 'date' }, { field: 'Days Waiting', type: 'integer' }, { field: 'To Pay', type: 'currency' }], totals: [{ field: 'Job', fn: 'count' }, { field: 'To Pay', fn: 'sum' }] },
    { name: 'Parts to order', query: 'Parts to reorder', group: 'Supplier', columns: ['Code', 'Part', { field: 'In Stock', type: 'integer' }, { field: 'Reorder Level', type: 'integer' }, { field: 'Reorder Quantity', type: 'integer' }, { field: 'Cost', type: 'currency' }, { field: 'Order Value', type: 'currency' }], totals: [{ field: 'Order Value', fn: 'sum' }] },
    { name: 'Takings by month', query: 'Takings by month', columns: ['Month', { field: 'Payments', type: 'integer' }, { field: 'Takings', type: 'currency' }], totals: [{ field: 'Payments', fn: 'sum' }, { field: 'Takings', fn: 'sum' }] },
    { name: 'Job sheets', title: 'Job sheets: the last two weeks', query: 'Job sheets', group: 'Technician', sort: 'Job', columns: ['Job', { field: 'Booked In', type: 'date' }, 'Customer', 'Phone', 'Device', 'Serial', 'Fault', 'Left With It', 'Status', { field: 'Quote', type: 'currency' }, { field: 'Promised', type: 'date' }], totals: [{ field: 'Job', fn: 'count' }] },
    { name: 'Technician turnaround', query: 'Technician turnaround', columns: ['Technician', { field: 'Jobs Repaired', type: 'integer' }, { field: 'Average Days', type: 'number' }, { field: 'Charged', type: 'currency' }], totals: [{ field: 'Jobs Repaired', fn: 'sum' }, { field: 'Charged', fn: 'sum' }] },
    { name: 'Account invoices due', query: 'Accounts to chase', group: 'Account', sort: 'Due Date', columns: ['Invoice', 'Job', { field: 'Invoice Date', type: 'date' }, { field: 'Due Date', type: 'date' }, { field: 'Days Overdue', type: 'integer' }, { field: 'Balance', type: 'currency' }], totals: [{ field: 'Balance', fn: 'sum' }] },
    { name: 'VAT by month', title: 'VAT by month, from the invoices', query: 'VAT by month', columns: ['Month', { field: 'Invoices', type: 'integer' }, { field: 'Net Sales', type: 'currency' }, { field: 'VAT', type: 'currency' }, { field: 'Invoiced', type: 'currency' }], totals: [{ field: 'Net Sales', fn: 'sum' }, { field: 'VAT', fn: 'sum' }, { field: 'Invoiced', fn: 'sum' }] },
  ],
  // The print templates: one document for each record of the table named.
  documents: [
    { name: 'Booking-in receipt', table: 'Jobs', paper: 'a5',
      letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail'] },
      blocks: [
        { type: 'heading', text: 'Booking-in receipt {Number}' },
        { type: 'fields', columns: 2, items: [
          { field: 'Number', label: 'Job number' }, { field: 'Booked In', label: 'Booked in' },
          { field: 'Device.Customer', label: 'Customer' }, { field: 'Device.Customer.Phone', label: 'Phone' },
          { field: 'Promised', label: 'Ready by' }, { field: 'Technician', label: 'Technician' },
        ] },
        { type: 'fields', columns: 2, items: [
          { field: 'Device.Kind', label: 'Device' }, { field: 'Device.Make', label: 'Make' }, { field: 'Device.Model', label: 'Model' },
          { field: 'Device.Serial', label: 'Serial number' }, { field: 'Device.Colour', label: 'Colour' }, { field: 'Left With It', label: 'Left with it' },
        ] },
        { type: 'fields', columns: 1, items: [{ field: 'Fault', label: 'Fault as described' }, { field: 'Quote', label: 'Quoted price, VAT included' }] },
        { type: 'text', text: 'Our terms, in short:', small: true },
        { type: 'text', text: '1. If no price is quoted above, we will check the device and tell you the price before we do any work. If we cannot repair it, there is nothing to pay.', small: true },
        { type: 'text', text: '2. Please keep a copy of your photos and files. We take care with your data, but a repair can lose it, and we cannot be held to blame for that.', small: true },
        { type: 'text', text: '3. When a part has to be ordered in for you, we take its price as a deposit. It comes off your bill, and is kept only if you cancel after the part has arrived.', small: true },
        { type: 'text', text: '4. The ready-by date is our best estimate. We will tell you if it moves, and tell you by your chosen way when the device is ready.', small: true },
        { type: 'text', text: '5. A device not collected within 90 days of our final reminder may be recycled or sold to cover what is owed. Please bring this receipt, or quote {Number}, when you collect.', small: true },
        { type: 'spacer', size: 16 },
        { type: 'signatures', labels: ['Customer', 'Booked in by'] },
        FOOT,
      ] },
    { name: 'Quote', table: 'Quotes', paper: 'a4',
      letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Tax Number'] },
      blocks: [
        { type: 'heading', text: 'Quote {Number}' },
        { type: 'fields', columns: 2, items: [
          { field: 'Number', label: 'Quote' }, { field: 'Quoted On', label: 'Date' }, { field: 'Valid Until', label: 'Valid until' }, { field: 'Job', label: 'Job number' },
          { field: 'Job.Device.Customer', label: 'For' }, { field: 'Job.Device.Customer.Phone', label: 'Phone' },
          { field: 'Job.Device.Make', label: 'Make' }, { field: 'Job.Device.Model', label: 'Model' }, { field: 'Job.Device.Serial', label: 'Serial number' },
        ] },
        { type: 'fields', columns: 1, items: [{ field: 'Job.Fault', label: 'Fault reported' }] },
        { type: 'lines', table: 'Quote Lines', columns: [
          { field: 'Kind', label: 'Item' }, { field: 'Description' }, { field: 'Quantity', label: 'Qty or hours', type: 'number' }, { field: 'Unit Price', label: 'Price', type: 'currency' },
          { expr: GROSS, as: 'Amount', type: 'currency' }, { expr: VAT_IN, as: 'VAT', label: 'VAT included', type: 'currency' },
        ] },
        { type: 'totals', rows: [
          { label: 'Net of VAT', expr: 'Round([Amount] - [VAT], 2)' }, { label: 'VAT', expr: '[VAT]' }, { label: 'Total', expr: '[Amount]', strong: true },
        ] },
        { type: 'text', text: 'This is the price for the repair described above, VAT included, and it holds until {Valid Until}. If we find another fault once the work starts, we will ask you before doing anything more.' },
        { type: 'text', text: 'To go ahead, reply to this quote or call us on {Business.Phone}, quoting {Number}. If you would rather not, there is a charge of half an hour at our bench rate for the diagnosis, and you can collect the device as it is.' },
        FOOT,
      ] },
    { name: 'Invoice', table: 'Invoices', paper: 'a4',
      letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Tax Number'] },
      blocks: [
        { type: 'heading', text: 'VAT invoice {Number}' },
        { type: 'fields', columns: 2, items: [
          { field: 'Number', label: 'Invoice' }, { field: 'Invoice Date', label: 'Date and tax point' }, { field: 'Due Date', label: 'Due' }, { field: 'Job', label: 'Job number' },
          { field: 'Job.Device.Customer', label: 'Customer' }, { field: 'Account.Name', label: 'Account' },
          { field: 'Job.Device.Customer.Address', label: 'Address' }, { field: 'Job.Device.Customer.Postcode', label: 'Postcode' },
          { field: 'Job.Device.Make', label: 'Make' }, { field: 'Job.Device.Model', label: 'Model' }, { field: 'Job.Device.Serial', label: 'Serial number' },
        ] },
        { type: 'lines', table: 'Job Lines', link: 'Invoice', columns: [
          { field: 'Kind', label: 'Item' }, { field: 'Description' }, { field: 'Quantity', label: 'Qty or hours', type: 'number' }, { field: 'Unit Price', label: 'Price', type: 'currency' },
          { expr: GROSS, as: 'Amount', type: 'currency' }, { field: 'Tax Rate', label: 'VAT rate', type: 'number' }, { expr: VAT_IN, as: 'VAT', label: 'VAT included', type: 'currency' },
        ] },
        { type: 'totals', rows: [
          { label: 'Net of VAT', expr: 'Round([Amount] - [VAT], 2)' }, { label: 'VAT', expr: '[VAT]' }, { label: 'Total', expr: '[Amount]', strong: true },
          { label: 'Deposit paid', expr: 'Nz([Deposit])' }, { label: 'Balance due', expr: 'Round([Total] - [Deposit paid], 2)', strong: true },
        ] },
        { type: 'text', text: '{Business.Payment Terms} By bank transfer to {Business.Bank Details}, quoting {Number}.', small: true },
        FOOT,
      ] },
    { name: 'Collection receipt', table: 'Jobs', paper: 'receipt-80',
      letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Phone', 'Tax Number'] },
      blocks: [
        { type: 'heading', text: 'Collection receipt' },
        { type: 'fields', columns: 1, items: [
          { field: 'Number', label: 'Job' }, { field: 'Collected', label: 'Collected' }, { field: 'Device.Customer', label: 'Customer' },
          { field: 'Device.Make', label: 'Make' }, { field: 'Device.Model', label: 'Model' }, { field: 'Device.Serial', label: 'Serial' }, { field: 'Outcome', label: 'Outcome' }, { field: 'Warranty Until', label: 'Guaranteed until' },
        ] },
        { type: 'lines', table: 'Job Lines', link: 'Job', columns: [
          { field: 'Description' }, { expr: GROSS, as: 'Amount', type: 'currency' }, { expr: VAT_IN, as: 'VAT', type: 'currency' },
        ] },
        { type: 'totals', rows: [
          { label: 'Total', expr: '[Amount]', strong: true }, { label: 'VAT included', expr: '[VAT]' }, { label: 'Deposit paid', expr: 'Nz([Deposit])' },
          { label: 'Paid today', expr: 'IIf([Paid], Round([Total] - [Deposit paid], 2), 0)' },
          { label: 'Still to pay', expr: 'IIf([Paid], 0, Round([Total] - [Deposit paid], 2))', strong: true },
        ] },
        { type: 'text', text: 'Handed back with: {Left With It}.', small: true },
        { type: 'signatures', labels: ['Received in working order'] },
        FOOT,
      ] },
    { name: 'Warranty certificate', table: 'Warranties', paper: 'a5',
      letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail'] },
      blocks: [
        { type: 'heading', text: 'Repair warranty {Number}' },
        { type: 'fields', columns: 2, items: [
          { field: 'Number', label: 'Certificate' }, { field: 'Job', label: 'Job number' }, { field: 'Issued', label: 'Repaired on' }, { field: 'Job.Warranty Until', label: 'Covered until' },
          { field: 'Job.Device.Customer', label: 'Customer' }, { field: 'Job.Device.Kind', label: 'Device' }, { field: 'Job.Device.Make', label: 'Make' },
          { field: 'Job.Device.Model', label: 'Model' }, { field: 'Job.Device.Serial', label: 'Serial number' },
        ] },
        { type: 'fields', columns: 1, items: [{ field: 'Covers', label: 'What it covers' }] },
        { type: 'text', text: '1. For 90 days from the repair, until {Job.Warranty Until}, we will put right free of charge, parts and labour included, any fault in the part we fitted or in our work, and the same fault if it comes back.' },
        { type: 'text', text: '2. It does not cover new damage after the repair (a drop, a crack, liquid or heat), parts of the device we did not work on, or a device that someone else has opened since. It does not cover lost data.' },
        { type: 'text', text: '3. To claim, bring the device to the shop with this certificate, or quote {Number}, before the warranty ends. We look at warranty work first, usually within two working days.' },
        { type: 'text', text: '4. This warranty is given on top of your rights as a customer and takes nothing away from them.' },
        { type: 'signatures', labels: ['For the shop'] },
        FOOT,
      ] },
    { name: 'Account statement', table: 'Accounts', paper: 'a4',
      letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Tax Number'] },
      blocks: [
        { type: 'heading', text: 'Statement of account' },
        { type: 'fields', columns: 2, items: [
          { field: 'Name', label: 'Account' }, { field: 'Contact', label: 'For the attention of' }, { field: 'Address' }, { field: 'City' },
          { field: 'Postcode' }, { field: 'Payment Days', label: 'Days to pay' }, { field: 'Credit Limit', label: 'Credit limit' },
        ] },
        { type: 'lines', table: 'Invoices', link: 'Account', sort: 'Invoice Date', columns: [
          { field: 'Number', label: 'Invoice' }, { field: 'Job', label: 'Job' }, { field: 'Invoice Date', label: 'Date' }, { field: 'Due Date', label: 'Due' },
          { field: 'Total', label: 'Amount', type: 'currency' }, { field: 'Paid On', label: 'Paid on' },
          { expr: 'IIf(Nz([Paid On], "") = "", [Balance], 0)', as: 'Owing', type: 'currency' },
          { expr: 'IIf(Nz([Paid On], "") = "" And [Due Date] < Date(), [Balance], 0)', as: 'Overdue', type: 'currency' },
        ] },
        { type: 'totals', rows: [{ label: 'Of which overdue', expr: '[Overdue]' }, { label: 'Balance owing', expr: '[Owing]', strong: true }] },
        { type: 'text', text: 'Invoices are due {Payment Days} days from their date. Please pay by bank transfer to {Business.Bank Details}, quoting the invoice numbers. If you have paid in the last few days, thank you, and please ignore what is shown as owing.', small: true },
        FOOT,
      ] },
  ],
};

export const about = {
  name: 'Ironbridge Repairs',
  kind: 'Repair shop',
  line: 'Jobs, devices, quotes, parts, invoices, warranties and payments for a shop that mends phones, tablets, laptops and small appliances.',
  summary: 'Ironbridge Repairs is an invented repair shop on a high street in Bristol. Three technicians work at the bench and one assistant runs the counter, where customers book in a phone with a cracked screen or a coffee machine that has stopped pumping, and come back later to collect it. A handful of local businesses have an account: their staff bring work in, and the shop invoices the business and sends it a statement.\n\n'
    + 'Every job is one record with its own number, JOB-2026-0042. It starts as Booked in, is diagnosed, gets a quote, waits for the customer to approve it, waits for a part if the shelf has none, is repaired and is marked Ready, then Collected. A job can also end as Unrepairable, or Ready with the quote turned down and a small diagnosis charge. The parts and labour each job used are its Job Lines, so the charge can be explained line by line; every payment, deposits included, is recorded against its job.\n\n'
    + 'Prices are shown to customers with VAT included, as a high-street shop shows them. Each line keeps its VAT rate, and the quote, the invoice and the receipt work out the VAT inside each line and add it up, so the VAT on the invoice and in the VAT by month query agree to the penny.\n\n'
    + 'The records cover about fifteen months up to the day the file is made. Phones come in more in the summer holidays and laptops when the school year starts; kitchen appliances pick up before the year ends. Some jobs are waiting on parts today, some customers have not picked up their devices after weeks, a few repairs have come back under warranty, an account or two is behind with its invoices, and the parts shelf has items at or below their reorder level, one of them held up because the supplier has it on back order.',
  suits: [
    'A phone and tablet repair kiosk',
    'A computer support firm that takes machines in',
    'A bike shop workshop, with bikes as the devices',
    'A watch and jewellery repairer',
    'An appliance engineer who also visits homes',
    'A musical instrument or audio repair bench',
  ],
  tables: {
    Business: BUSINESS_ABOUT,
    'Accounts': { about: 'The businesses that have an account with the shop: their staff bring work in, the shop invoices the business, and the business pays by bank transfer within its payment days.', fields: {
      'Name': 'The business\'s name, printed on its invoices and statement. Each name is used once.',
      'Contact': 'The person who looks after the account there, printed on the statement.',
      'Phone': 'The business\'s main number.',
      'E-mail': 'Where invoices and statements are sent.',
      'Address': 'The street address, printed on the statement.',
      'City': 'The town or city.',
      'Postcode': 'The postcode.',
      'Payment Days': 'How many days after the invoice date payment is due: 30 means thirty days.',
      'Credit Limit': 'The most the business may owe at once before the counter asks for payment first.',
      'Opened': 'The day the account was opened.',
    } },
    'Customers': { about: 'One record for each person who brings work in. Someone who brings work for a business on account is linked to that account; everyone else pays at the counter.', fields: {
      'First Name': 'The customer\'s first name.',
      'Last Name': 'The family name. Needed for every customer.',
      'Account': 'The business account the person brings work in for. Empty for members of the public.',
      'Phone': 'The number the shop rings or texts when the device is ready.',
      'E-mail': 'Where receipts and reminders go, when the customer gave one.',
      'Address': 'The street address, used for invoices and for home visits.',
      'City': 'The town or city.',
      'Postcode': 'The postcode.',
      'Contact By': 'How the customer wants to hear that the job is ready: Text message, Phone call or E-mail.',
      'Joined': 'The day the customer first brought something in.',
    } },
    'Staff': { about: 'The people at the shop: the technicians who take jobs and the counter assistant.', fields: {
      'First Name': 'First name.',
      'Last Name': 'Family name.',
      'Role': 'Senior Technician, Technician or Counter Assistant.',
      'Skills': 'What each person is best at, so the counter books jobs to the right bench.',
      'Phone': 'Work mobile.',
      'E-mail': 'Work e-mail address.',
      'Started': 'The day they joined the shop.',
      'Active': 'No once someone has left, so they stop appearing in lists of who can take a job.',
    } },
    'Suppliers': { about: 'The firms parts are bought from.', fields: {
      'Name': 'The supplier\'s name.',
      'Contact': 'The person the shop deals with there.',
      'Phone': 'Their trade counter number.',
      'E-mail': 'Where orders are sent.',
      'Supplies': 'What the shop buys from them.',
      'Lead Days': 'How many working days an order usually takes to arrive, used to set the expected date.',
      'Our Account': 'The shop\'s account reference with the supplier (an invented reference here).',
    } },
    'Parts': { about: 'Every part the shop fits, whether kept on the shelf or ordered in for each job.', fields: {
      'Code': 'The shop\'s own code for the part, printed on the bin label.',
      'Name': 'What the part is, in plain words. It is copied onto the job line, quote and invoice.',
      'Fits': 'The kind of device it is for: Phone, Tablet, Laptop or an appliance.',
      'Supplier': 'Who it is bought from.',
      'Cost': 'What one costs the shop, without VAT.',
      'Price': 'What one is charged at on a job, VAT included.',
      'In Stock': 'How many are on the shelf now, including any already set aside for a job.',
      'Reorder Level': 'When stock falls to this number the part is ordered again.',
      'Reorder Quantity': 'How many to order at a time.',
      'Stocked': 'Yes for parts kept on the shelf. No for parts ordered in for each job, such as main boards and model-specific keyboards.',
      'Bin': 'Where it is kept on the parts wall.',
    } },
    'Devices': { about: 'Each device a customer has brought in. One customer can have several, and a device that comes back keeps its record, so its history is in one place.', fields: {
      'Customer': 'Who owns it, or who brought it in for their business.',
      'Kind': 'Phone, Tablet, Laptop, Coffee machine, Vacuum cleaner, Food mixer or Sewing machine.',
      'Make': 'The maker\'s name (invented makes in this sample).',
      'Model': 'The model name or number.',
      'Serial': 'The serial number, checked when the device is handed back.',
      'Colour': 'Its colour, which helps find it on the shelf.',
    } },
    'Jobs': { about: 'One record for each repair, from booking in to collection. The Booking-in receipt and the Collection receipt print from it.', fields: {
      'Number': 'The job number on the ticket, given in the order jobs are booked in, a new run each year: JOB-2026-0042.',
      'Device': 'The device being repaired, and through it the customer.',
      'Technician': 'Who the job is booked to.',
      'Booked In': 'The day the device was handed in.',
      'Fault': 'The fault as the customer described it at the counter. A warranty return names the first job.',
      'Left With It': 'Anything left with the device, such as a charger, case or water tank.',
      'Promised': 'The day the customer was told it would be ready.',
      'Status': 'Booked in, Diagnosing, Awaiting approval, Awaiting parts, Repairing, Ready, Collected or Unrepairable.',
      'Quote': 'The price given for the repair, VAT included. Fixed-price jobs such as screens are quoted at the counter; others after diagnosis, with a written quote in Quotes.',
      'Quoted On': 'The day the quote was given.',
      'Quote Approved': 'Yes once the customer has agreed to the quote.',
      'Deposit': 'Money taken up front when a part has to be ordered in.',
      'Outcome': 'How the job ended: Repaired, Quote declined or Beyond economic repair.',
      'Charge': 'What the job finally costs the customer, VAT included: the total of its lines. Nothing for a device that cannot be repaired ("no fix, no fee") or a warranty return.',
      'Finished': 'The day the work was done and the customer could be told. The invoice is dated that day.',
      'Collected': 'The day the customer took the device away.',
      'Paid': 'Yes once everything charged has been paid.',
      'Warranty Until': 'The last day of the 90-day warranty on the repair, printed on its certificate.',
      'Warranty Of': 'For a warranty return, the first job it came back from.',
      'Notes': 'Anything else the counter or the bench needs to know, such as reminders sent.',
    } },
    'Part Orders': { about: 'Parts ordered from suppliers: stock to refill the shelf, or one part for one job.', fields: {
      'Supplier': 'Who the order went to.',
      'Part': 'The part ordered.',
      'Job': 'The job the part is for, when it was ordered for one (its record number). Empty for shelf stock.',
      'Quantity': 'How many were ordered.',
      'Unit Cost': 'The price paid for each, without VAT.',
      'Ordered On': 'The day the order was placed.',
      'Expected': 'The day the supplier said it would arrive.',
      'Received On': 'The day it actually arrived. Empty while it is still on its way.',
      'Status': 'On order or Received.',
    } },
    'Job Lines': { about: 'The parts and labour on each job, one line each, at the price charged with VAT included. They are the lines of the job\'s invoice and its collection receipt.', fields: {
      'Job': 'The job the line belongs to.',
      'Invoice': 'The invoice the line is on, once the job is invoiced. Empty before then, and on jobs that cost nothing.',
      'Kind': 'Part or Labour.',
      'Part': 'The part fitted or waiting to be fitted. Empty on a labour line.',
      'Description': 'What the line is: the part\'s name, or the work done, as the invoice prints it.',
      'Quantity': 'How many parts, or for labour the hours: 0.75 is three quarters of an hour.',
      'Unit Price': 'The price of one part or one hour, VAT included (nothing on a warranty return or a check that found no fix).',
      'Tax Rate': 'The VAT rate inside the price, as a fraction: 0.2 is twenty percent. Taken from the Business record when the line is made.',
      'Done On': 'The day the part was fitted or the work done. Empty while the job is waiting for it.',
    } },
    'Quotes': { about: 'The written quotes given after a device has been diagnosed, with the price, VAT included, and whether the customer said yes. The Quote print template prints from it.', fields: {
      'Number': 'The quote\'s number, in date order, a new run each year: QUO-2026-0042.',
      'Job': 'The job it is for.',
      'Quoted On': 'The day it was given.',
      'Valid Until': 'The last day the price holds: thirty days from the quote.',
      'Net': 'The total without VAT.',
      'Tax': 'The VAT inside the total, worked out line by line.',
      'Total': 'The price the customer is asked to pay, VAT included.',
      'Answer': 'Accepted or Declined; empty while the customer has not answered.',
      'Answered On': 'The day the customer answered.',
    } },
    'Quote Lines': { about: 'The parts and labour on each quote.', fields: {
      'Quote': 'The quote the line belongs to.',
      'Kind': 'Part or Labour.',
      'Description': 'The part, or the work to be done.',
      'Quantity': 'How many parts, or for labour the hours.',
      'Unit Price': 'The price of one part or one hour, VAT included.',
      'Tax Rate': 'The VAT rate inside the price, as a fraction.',
    } },
    'Invoices': { about: 'An invoice for every job that costs something, raised the day the work is finished. It lists the job\'s lines, the VAT inside them, any deposit already paid and the balance. Members of the public pay it on collection; businesses on account pay it within their payment days.', fields: {
      'Number': 'The invoice\'s number, given in date order and once only, a new run each year: INV-2026-0042.',
      'Job': 'The job invoiced.',
      'Account': 'The business account billed, for work brought in on account. Empty for members of the public.',
      'Invoice Date': 'The day it was raised, the day the work was finished. It is the tax point for VAT.',
      'Due Date': 'When payment is due: the invoice date for the public (pay on collection), or that date and the account\'s payment days.',
      'Net': 'The total without VAT.',
      'Tax': 'The VAT inside the total: each line\'s VAT rounded to the penny, then added up, as the invoice prints it.',
      'Total': 'What the invoice comes to, VAT included. The same as the job\'s Charge.',
      'Deposit': 'The deposit already paid against the job, taken off the total.',
      'Balance': 'What was left to pay when the invoice was raised: the total less the deposit.',
      'Paid On': 'The day the balance was paid; empty while it is owed.',
    } },
    'Warranties': { about: 'A 90-day warranty certificate for every job repaired (not for a warranty return, which is covered by the first one). The Warranty certificate print template prints from it.', fields: {
      'Number': 'The certificate\'s number, in date order, a new run each year: WTY-2026-0042.',
      'Job': 'The job whose repair it covers. The last day of cover is the job\'s Warranty Until.',
      'Issued': 'The day the repair was finished and the certificate given.',
      'Covers': 'What it covers, in words: the part fitted and the work done.',
    } },
    'Payments': { about: 'Every payment taken, each against its job, each with its own receipt number.', fields: {
      'Number': 'The receipt number, in date order, a new run each year: RCP-2026-0042.',
      'Job': 'The job paid for.',
      'Paid On': 'The day the money came in.',
      'Amount': 'How much.',
      'Method': 'Card, Cash or Bank transfer.',
      'Kind': 'Deposit, On collection, or On account for a business paying its invoice.',
    } },
  },
  queries: {
    'Job board': 'Every job not yet collected, in the order the shop works through them, with who it belongs to, the technician, the promised date and how many days are left (below nought means it is late). Open it first thing in the morning.',
    'Jobs by status': 'How many jobs sit at each stage now, and the oldest booking at each. A quick count for the morning meeting.',
    'Awaiting parts': 'Jobs held up for a part, with the part, the supplier and when it is expected. Use it when a customer rings to ask, or to chase a supplier.',
    'Past promised date': 'Unfinished jobs whose promised date has gone by, latest first, with the customer\'s number so the counter can ring them with news.',
    'Ready for collection': 'Every finished device still on the shelf, longest waiting first, with what is left to pay and how the customer wants to be told. The calling list.',
    'Uncollected 14 days or more': 'Devices ready for two weeks or longer and not picked up, with any reminders already sent. Use it to send a final reminder before the shop\'s storage terms apply.',
    'Quotes awaiting approval': 'Quotes given and not yet answered, oldest first, with the price and the phone number to follow up.',
    'Parts to reorder': 'Shelf parts at or below their reorder level, by supplier, with what a full reorder would cost.',
    'Parts on order': 'Everything ordered and not yet arrived, with the expected date and the record number of the job it is for when there is one.',
    'Takings by month': 'Money taken each month, deposits and account payments included, and how many payments made it up.',
    'Takings last 30 days': 'The last thirty days of takings split by payment method, for checking against the card machine and the bank statement.',
    'Technician workload': 'How many unfinished jobs each technician has, how many are stuck on parts or quotes, and their oldest job.',
    'Technician turnaround': 'For repairs finished in the last year: how many each technician did, the average days from booking in to finished, and what they were charged at.',
    'Work by device kind': 'Jobs and money by kind of device over the whole period, with the average days each kind takes.',
    'Warranty returns': 'Jobs that came back under warranty, newest first; the fault names the first job, so repeat faults on a part or a technician\'s work show up.',
    'Job sheets': 'Every job booked in over the last fourteen days with its full detail: serial, fault, what was left with it, quote and promised date.',
    'Parts margin': 'Each part fitted over the period: how many, what they were charged at with VAT, and the margin over cost on the price without VAT. Warranty fittings count against the margin.',
    'Accounts to chase': 'Every invoice a business account has not paid yet, by account and due date, with how many days overdue it is (nought when it is not due yet).',
    'Under warranty': 'Repairs still inside their 90 days, with the certificate number and what it covers, those ending soonest first. Look a customer up here when a device comes back.',
    'VAT by month': 'For each month, the invoices raised, their net value, the VAT inside them and the total: the figures for the VAT return.',
  },
  reports: {
    'Job board': 'The job board printed for the bench, one section per stage, each job with its customer, device, fault, technician and promised date.',
    'Ready for collection': 'A calling list for the counter: each device waiting to be collected, the customer\'s number, how they like to be contacted and what they owe, longest waiting first.',
    'Parts to order': 'The parts to reorder, one section per supplier, with the cost of each order. Take it to the phone or the supplier\'s website.',
    'Takings by month': 'Monthly takings with a grand total, for the accountant or the owner\'s month-end look.',
    'Job sheets': 'The last two weeks of jobs, one section per technician, with every detail taken at the counter. Print it to check work in progress against the shelf.',
    'Technician turnaround': 'Each technician\'s finished repairs over the last year, their average turnaround and the value of their work.',
    'Account invoices due': 'Unpaid account invoices, a section for each business, with days overdue and the total owed: the list for chasing payment.',
    'VAT by month': 'Net sales, VAT and invoiced totals for each month with the grand totals, to fill in the VAT return or hand to the accountant.',
  },
  documents: {
    'Booking-in receipt': 'Printed at the counter when a device is booked in, on A5, one copy for the customer: the job number, the device and its serial, the fault as described, what was left with it, the price if one was quoted, the ready-by date, the shop\'s terms in short and two lines to sign. It prints from Jobs, so print the job on screen.',
    'Quote': 'The written quote sent after diagnosis: the parts and labour, the VAT inside them, the total and how long the price holds, with how to say yes. One for each record in Quotes.',
    'Invoice': 'The VAT invoice raised when the work is finished: the parts and labour lines with the VAT in each, the net, VAT and total, the deposit already paid and the balance due. One for each record in Invoices.',
    'Collection receipt': 'The till receipt handed over with the device: the work done and its price, VAT included, the deposit, what was paid today and a line to sign that the device came back in working order. It prints from Jobs, so print the job on screen as it is collected; a job not yet collected prints with what is still to pay.',
    'Warranty certificate': 'The 90-day warranty given with every repair: the device, the dates, what it covers and what it does not, and how to claim. One for each record in Warranties. Have the wording checked against the consumer law of your own country.',
    'Account statement': 'Sent to each business account at the month end: every invoice raised to the account with the day it was paid, what is still owing and how much of that is overdue. One for each record in Accounts.',
  },
  tutorial: [
    { title: 'Book a device in and print the receipt', steps: [
      'Open the Customers form and look the person up by name or phone. If they are new, press New and type their name, phone and how they want to be told (Contact By). If they bring work for a business with an account, choose it in Account.',
      'Open the Devices form, press New, choose the customer in Customer, then fill in Kind, Make, Model, Serial and Colour. A device that has been in before already has a record: use that one.',
      'Open the Jobs form and press New. Type the next job number in Number (JOB-2026- and the next number in the run), choose the device in Device and the technician in Technician, and fill in Booked In, Fault, Left With It and Promised. For a fixed-price repair such as a screen, type the price in Quote and tick Quote Approved.',
      'With the job on screen, open the print template Booking-in receipt from Print templates, and press Print. The customer signs one copy and keeps it.',
    ] },
    { title: 'Quote after diagnosis, and hear back', steps: [
      'When the bench has found the fault, open the Quotes form and press New. Type the next number (QUO-2026-...), choose the job in Job, and set Quoted On to today and Valid Until to thirty days on.',
      'In Quote Lines, add a Part line with the part\'s name and price and a Labour line with the hours and the bench rate, VAT included in both. Put the total in Total and in the job\'s Quote, and set the job\'s Status to Awaiting approval.',
      'Open the print template Quote, choose the quote, and press Save as PDF to e-mail it, or Print.',
      'Each morning open the query Quotes awaiting approval and ring anyone who has not answered. When they answer, fill in Answer and Answered On on the quote, and on the job tick Quote Approved or set the outcome to Quote declined.',
    ] },
    { title: 'Order a part for a job and take a deposit', steps: [
      'If the part is not on the shelf (In Stock on the Parts form), open the Part Orders form, press New, choose the supplier, the part and the job, and fill in Quantity, Unit Cost, Ordered On and Expected. Set the job\'s Status to Awaiting parts.',
      'For a part ordered in specially, take its price as a deposit: type it in the job\'s Deposit, then open the Payments form, press New, choose the job, and record the amount with Kind Deposit and the next receipt number.',
      'When the part arrives, type Received On on the part order and set its Status to Received. The query Awaiting parts shows what is still on its way.',
    ] },
    { title: 'Finish a repair, invoice it and give the warranty', steps: [
      'On the Jobs form, add the job\'s lines in Job Lines: the part fitted (Kind Part, the part, its name in Description, Quantity 1 and its price) and the labour (Kind Labour, what was done, the hours in Quantity and the rate in Unit Price). Fill in Done On on each.',
      'Set the job\'s Status to Ready and Outcome to Repaired, type today in Finished, the total of the lines in Charge, and in Warranty Until the day 90 days on.',
      'Open the Invoices form and press New. Type the next number (INV-2026-...), choose the job, choose the account if it is business work, and fill in Invoice Date, Due Date, Net, Tax, Total, Deposit and Balance. Then on each of the job\'s lines choose the new invoice in Invoice.',
      'Open the Warranties form, press New, type the next number (WTY-2026-...), choose the job and say in Covers what was fitted and done.',
      'Open the print templates Invoice and Warranty certificate, choose the new records, and press Save as PDF to e-mail them, or Print to put them in the bag with the device.',
      'Ring or text the customer: the query Ready for collection is the list, with how each wants to be told.',
    ] },
    { title: 'Hand a device back and take payment', steps: [
      'Find the job on the Job board query or the Jobs form by its number from the customer\'s receipt, and check the serial on the device against Devices.',
      'Open the Payments form, press New, choose the job, type the amount left to pay, the Method and Kind On collection, with the next receipt number. On the invoice fill in Paid On; on the job type Collected, tick Paid and set Status to Collected.',
      'With the job on screen, open the print template Collection receipt and press Print. The customer signs that the device came back in working order.',
    ] },
    { title: 'Send the business accounts their statements', steps: [
      'At the month end, open the report Account invoices due to see what each business owes, and how late.',
      'Open the print template Account statement from Print templates. It makes one statement for each record in Accounts; choose the accounts that owe something, and press Save as PDF to e-mail each one to the address on its Accounts record.',
      'When a business pays, open the Payments form, record each invoice paid with Kind On account and Method Bank transfer, fill in Paid On on the invoice, and tick Paid on the job.',
    ] },
    { title: 'Reorder parts and do the VAT', steps: [
      'Every Monday, open the report Parts to order: a page for each supplier with what has fallen to its reorder level and what a full order costs. Place the orders and record them in Part Orders.',
      'At the end of each VAT quarter, open the report VAT by month and add up the three months\' Net Sales and VAT for the return.',
      'The query Parts margin shows what each part has earned over its cost, with the VAT taken out of the price.',
    ] },
  ],
  notes: [
    'Prices are typed with VAT included, as the customer sees them: a phone screen at 79.00 is 65.83 and 13.17 VAT. The VAT is worked out inside each line and rounded to the penny, then added up, so an invoice\'s Tax is the sum of its lines and not 20 percent of its total worked out once.',
    'The numbers on jobs, quotes, invoices, warranties and receipts (JOB-, QUO-, INV-, WTY-, RCP-) are typed, not made by Data Files. Keep each run in date order and start again at 0001 each January. The record number (ID) is separate and is what Part Orders shows in For Job.',
    'An invoice\'s lines are the job\'s own lines, Job Lines, pointed at the invoice once it is raised. Do not add lines to a job after its invoice is made; raise a second invoice for extra work.',
    'A print template prints every record of its table. The Booking-in receipt and the Collection receipt print from Jobs and the Account statement from Accounts, so choose the record you want (or print from the form with it on screen) rather than printing them all.',
    'The Account statement lists every invoice the business has ever had, the paid ones with the day they were paid; the Owing column counts only those still unpaid.',
    'A job that is beyond economic repair, a warranty return and a declined quote still have lines, at nought or at the diagnosis charge, so the bench time is recorded even when nothing is charged.',
    'Customers who bring work for a business point at its Accounts record. The invoice is made out to the account and its balance is paid by transfer, not at the counter, so Accounts to chase is the list to watch.',
  ],
  faq: [
    { q: 'How do I print an older invoice or quote?', a: 'Open the print template Invoice (or Quote) and choose the record you want by its number; every invoice ever raised is there. To find the number, look the job up on the Jobs form or in the query Job board.' },
    { q: 'A customer has lost their receipt. How do I find their device?', a: 'Open the Customers form and find them by name or phone, then look at their devices and jobs. The query Ready for collection lists every finished device still on the shelf with its customer.' },
    { q: 'Is this device still under warranty?', a: 'Open the query Under warranty and look for the customer or the serial. If the job is there, its Days Left is how long the cover has to run. Book the return in as a new job with Warranty Of set to the first job and nothing to pay.' },
    { q: 'Why does the invoice say less VAT than 20 percent of the total?', a: 'The prices include VAT, so the VAT is one sixth of each line, not a fifth: 20 percent of the price without VAT. On a total of 112.75 the VAT is 18.79.' },
    { q: 'How do I change the bench rate or the VAT rate?', a: 'The rate per hour is whatever you type in Unit Price on each labour line, so a new rate applies from the next job. The VAT rate is in the Business record (Tax Rate) and on each line; old lines keep the rate they were charged at.' },
    { q: 'How do I see what an account owes?', a: 'Open the query Accounts to chase, or print the Account statement for that business. Both show only invoices without a Paid On date.' },
    { q: 'Can I start with my own shop instead of Ironbridge?', a: 'Yes. Make the database empty from the sample, fill in the Business record with your own name, address, VAT number and bank details, and every document prints with them.' },
  ],
  empty: 'Start with Business, your own name, address, VAT number and payment details, which every document prints. Then Staff, then Suppliers, then Parts (each part pointing at its supplier, with Stocked set to No for parts you only order in), and Accounts for any businesses you invoice. When a customer comes to the counter, add them to Customers, add their device to Devices, then open a Job for that device with its number, the fault, the technician and the promised date. As the job moves on, change its Status, fill in the quote (and a record in Quotes when it is written), and add its parts and labour in Job Lines. When it is finished, raise its Invoice and its Warranty, record each payment in Payments as it is taken, and put Finished and Collected on the job when they happen.',
  adapt: [
    'For a bike workshop, use Devices for bikes (Kind as Road, Mountain, Hybrid, E-bike; Serial as the frame number) and add a Booked For date for the day the customer wants it back.',
    'For an engineer who visits homes, add Visit Date and Visit Address to Jobs and a call-out line in Job Lines, and make a query of visits in the next 7 days ordered by postcode.',
    'Add a Courier Reference to Part Orders, and a query of orders past their Expected date, to chase late suppliers.',
    'Add a status such as Sent to maker for repairs sent away, and add it to the Stage calculation in the Job board query so it sorts in the right place.',
    'Add a Disposal Date to Jobs for devices never collected, and a query of devices uncollected for 90 days or more, if your terms say you may then recycle or sell them.',
    'If you are not registered for VAT, set Tax Rate to 0 in the Business record and on new lines; the documents then show no VAT, and the heading of the Invoice template can be changed from VAT invoice to Invoice.',
    'Add Passcode Given (Yes/No) and Data Backed Up (Yes/No) to Jobs, so the bench knows what it can test before starting, and add them to the Booking-in receipt.',
  ],
};

// The bench rate per hour, and the charge for checking a device when its quote is turned down.
const RATE = 45;
const DIAGNOSIS = 0.5;
const WARRANTY_DAYS = 90;
// The VAT rate inside every price, as a fraction (the Business record keeps the same).
const VAT = 0.2;

// Suppliers: [name, what they supply, lead days].
const SUPPLIERS = [
  ['Copperfield Spares', 'Phone and tablet screens, batteries and cameras', 2],
  ['Saffron Components', 'Laptop screens, drives, batteries and boards', 3],
  ['Granite Appliance Spares', 'Parts for coffee machines, vacuums, mixers and sewing machines', 4],
  ['Juniper Parts Direct', 'Ports, jacks, leads and small parts', 2],
];

// Parts: [code, name, fits, supplier, cost, price, reorder level, reorder quantity, bin].
// A reorder level of 0 is a part ordered in for each job and not kept on the shelf.
const PARTS = [
  ['PH-SCR61', 'Phone screen, 6.1 in', 'Phone', 0, 38, 79, 4, 6, 'A1'],
  ['PH-SCR67', 'Phone screen, 6.7 in', 'Phone', 0, 52, 109, 3, 5, 'A1'],
  ['PH-BAT1', 'Phone battery, standard', 'Phone', 0, 9, 39, 4, 8, 'A2'],
  ['PH-BAT2', 'Phone battery, large', 'Phone', 0, 12, 45, 3, 6, 'A2'],
  ['PH-PORT', 'Charging port, USB-C', 'Phone', 3, 4.5, 29, 3, 8, 'A3'],
  ['PH-CAM', 'Rear camera module', 'Phone', 0, 22, 59, 0, 0, 'A3'],
  ['PH-BACK', 'Back glass panel', 'Phone', 0, 8, 35, 0, 0, 'A4'],
  ['PH-EAR', 'Earpiece speaker', 'Phone', 3, 3.5, 25, 1, 4, 'A4'],
  ['TB-GLS', 'Tablet glass digitiser, 10 in', 'Tablet', 0, 28, 69, 2, 3, 'A5'],
  ['TB-BAT', 'Tablet battery', 'Tablet', 0, 15, 49, 1, 3, 'A5'],
  ['TB-CHG', 'Tablet charging board', 'Tablet', 0, 9, 35, 0, 0, 'A6'],
  ['LT-SCR14', 'Laptop screen panel, 14 in', 'Laptop', 1, 55, 119, 1, 3, 'B1'],
  ['LT-SCR15', 'Laptop screen panel, 15.6 in', 'Laptop', 1, 60, 129, 1, 2, 'B1'],
  ['LT-KEY', 'Laptop keyboard, UK layout', 'Laptop', 1, 18, 55, 0, 0, 'B2'],
  ['LT-BAT', 'Laptop battery, 4-cell', 'Laptop', 1, 32, 79, 1, 3, 'B2'],
  ['LT-SSD5', 'SSD, 500 GB', 'Laptop', 1, 28, 65, 2, 5, 'B3'],
  ['LT-SSD1', 'SSD, 1 TB', 'Laptop', 1, 45, 95, 1, 3, 'B3'],
  ['LT-RAM', 'Memory, 8 GB module', 'Laptop', 1, 16, 45, 1, 3, 'B3'],
  ['LT-PSU', 'Power adaptor, 65 W', 'Laptop', 1, 12, 35, 2, 4, 'B4'],
  ['LT-FAN', 'Cooling fan, laptop', 'Laptop', 1, 8, 29, 1, 3, 'B4'],
  ['LT-JACK', 'DC power jack', 'Laptop', 3, 3, 25, 2, 5, 'B5'],
  ['LT-HNG', 'Laptop hinge set', 'Laptop', 1, 14, 39, 0, 0, 'B5'],
  ['LT-MB', 'Laptop main board', 'Laptop', 1, 180, 299, 0, 0, 'B6'],
  ['AP-PUMP', 'Pump, coffee machine', 'Coffee machine', 2, 14, 39, 1, 3, 'C1'],
  ['AP-HEAT', 'Thermoblock heater', 'Coffee machine', 2, 35, 79, 0, 0, 'C1'],
  ['AP-SEAL', 'Seal and gasket kit', 'Coffee machine', 2, 4, 15, 2, 6, 'C2'],
  ['AP-VFLT', 'Vacuum filter set', 'Vacuum cleaner', 2, 6, 18, 2, 5, 'C3'],
  ['AP-BRSH', 'Brush roll', 'Vacuum cleaner', 2, 9, 25, 1, 3, 'C3'],
  ['AP-VMOT', 'Vacuum motor', 'Vacuum cleaner', 2, 38, 85, 0, 0, 'C4'],
  ['AP-GEAR', 'Mixer drive gear set', 'Food mixer', 2, 7, 25, 1, 3, 'C5'],
  ['AP-MCB', 'Mixer speed control board', 'Food mixer', 2, 26, 59, 0, 0, 'C5'],
  ['AP-LEAD', 'Mains lead, fused', 'Food mixer', 3, 3, 12, 2, 5, 'C6'],
  ['AP-BELT', 'Sewing machine drive belt', 'Sewing machine', 2, 4, 15, 1, 4, 'C6'],
  ['AP-BOB', 'Bobbin case', 'Sewing machine', 2, 8, 22, 1, 3, 'C6'],
];

// What comes in: [make, model, size]. The makes are invented; no real brand is meant.
const MODELS = {
  Phone: [['Veltro', 'V12', 61], ['Veltro', 'V14 Pro', 67], ['Kestrin', 'K5', 61], ['Kestrin', 'K7 Max', 67], ['Ostrel', 'Lark 3', 61], ['Tamberlin', 'T20', 67]],
  Tablet: [['Veltro', 'Slate 10', 10], ['Kestrin', 'Tab 8', 8], ['Pellow', 'Pad 11', 11]],
  Laptop: [['Halvard', 'Book 14', 14], ['Halvard', 'Book 16', 15], ['Tarnwick', 'Air 13', 14], ['Pellow', 'Notebook 15', 15], ['Kestrin', 'Work 14', 14]],
  'Coffee machine': [['Quillon', 'Barista 2', 0], ['Quillon', 'Crema 5', 0]],
  'Vacuum cleaner': [['Corriby', 'Cyclone 5', 0], ['Corriby', 'Stick 22', 0]],
  'Food mixer': [['Quillon', 'Stand Mixer 500', 0], ['Ambrel', 'Mix Pro', 0]],
  'Sewing machine': [['Fennick', 'Stitch 40', 0], ['Fennick', 'Home 12', 0]],
};
const COLOURS = { Phone: ['Black', 'Silver', 'Blue', 'White', 'Green', 'Rose'], Tablet: ['Grey', 'Silver', 'Black'], Laptop: ['Grey', 'Silver', 'Black', 'Blue'] };
const APPLIANCE_COLOURS = ['White', 'Black', 'Red', 'Cream', 'Stainless'];
const LEFT_WITH = {
  Phone: [['Nothing', 6], ['Case', 3], ['Case and charger', 1]],
  Tablet: [['Nothing', 4], ['Charger', 2], ['Keyboard cover', 1]],
  Laptop: [['Charger', 6], ['Nothing', 2], ['Charger and bag', 2]],
  'Coffee machine': [['Water tank', 3], ['Nothing', 2], ['Water tank and milk jug', 1]],
  'Vacuum cleaner': [['Nothing', 3], ['Hose and tools', 2]],
  'Food mixer': [['Bowl', 2], ['Bowl and beater', 2], ['Nothing', 1]],
  'Sewing machine': [['Foot pedal', 3], ['Foot pedal and cover', 2]],
};

// Faults by kind: [weight, fault, part, hours, work, quoted after diagnosis, chance beyond repair, chance quote approved].
// A part of 'PH-SCR', 'PH-BAT', 'LT-SCR' or 'LT-SSD' is chosen for the device's size.
const FAULTS = {
  Phone: [
    [20, 'Cracked screen, touch still works', 'PH-SCR', 0.75, 'Replace screen', false, 0, 1],
    [14, 'Cracked screen, no display', 'PH-SCR', 0.75, 'Replace screen', false, 0, 1],
    [12, 'Battery drains quickly', 'PH-BAT', 0.5, 'Replace battery', false, 0, 1],
    [8, 'Turns off at 20 per cent', 'PH-BAT', 0.5, 'Replace battery', false, 0, 1],
    [9, 'Will not charge', 'PH-PORT', 0.75, 'Replace charging port', false, 0, 1],
    [5, 'Charges only at an angle', null, 0.25, 'Clean charging port and test', false, 0, 1],
    [4, 'Rear camera blurry', 'PH-CAM', 0.75, 'Replace rear camera', false, 0, 1],
    [5, 'Back glass cracked', 'PH-BACK', 1.25, 'Replace back glass', false, 0, 1],
    [4, 'Cannot hear callers', 'PH-EAR', 0.5, 'Replace earpiece speaker', false, 0, 1],
    [6, 'Dropped in water, will not turn on', 'PH-BAT', 1.5, 'Clean and dry board, new battery, test', true, 0.35, 0.8],
    [3, 'Move photos and contacts to a new phone', null, 0.5, 'Transfer data and set up', false, 0, 1],
  ],
  Tablet: [
    [45, 'Cracked glass', 'TB-GLS', 1.25, 'Replace glass digitiser', false, 0, 1],
    [20, 'Battery swollen, case lifting', 'TB-BAT', 1, 'Replace battery', false, 0, 1],
    [20, 'Will not charge', 'TB-CHG', 1, 'Replace charging board', true, 0.08, 0.85],
    [15, 'Stuck on the start-up screen', null, 1, 'Restore software and test', true, 0.1, 0.9],
  ],
  Laptop: [
    [18, 'Cracked screen', 'LT-SCR', 1, 'Replace screen panel', false, 0, 1],
    [10, 'Keys not working after a spill', 'LT-KEY', 1.5, 'Clean and replace keyboard', true, 0.12, 0.8],
    [14, 'Very slow, wants an SSD', 'LT-SSD', 1.5, 'Fit SSD and copy data across', false, 0, 1],
    [10, 'Battery will not hold a charge', 'LT-BAT', 0.5, 'Replace battery', false, 0, 1],
    [9, 'Overheats and shuts down', 'LT-FAN', 1, 'Replace fan and clean heatsink', true, 0.03, 0.9],
    [8, 'Will not power on', 'LT-JACK', 1.5, 'Replace DC power jack', true, 0.15, 0.85],
    [4, 'Will not power on', 'LT-MB', 1.5, 'Replace main board', true, 0.1, 0.55],
    [9, 'Virus warnings and pop-ups', null, 2, 'Remove malware and update system', false, 0, 1],
    [6, 'Hinge broken, lid loose', 'LT-HNG', 1.5, 'Replace hinge set', true, 0.05, 0.85],
    [4, 'Needs more memory', 'LT-RAM', 0.5, 'Fit memory and test', false, 0, 1],
    [3, 'Charger lost, needs a new one', 'LT-PSU', 0.25, 'Supply and test adaptor', false, 0, 1],
    [4, 'Recover files from a dead drive', null, 3, 'Recover data to a new drive', true, 0.25, 0.85],
  ],
  'Coffee machine': [
    [40, 'No water coming through', 'AP-PUMP', 1, 'Replace pump and descale', true, 0.05, 0.85],
    [35, 'Leaking from the base', 'AP-SEAL', 0.75, 'Replace seals and test', true, 0.02, 0.9],
    [25, 'Not heating', 'AP-HEAT', 1, 'Replace thermoblock', true, 0.1, 0.75],
  ],
  'Vacuum cleaner': [
    [40, 'Loses suction', 'AP-VFLT', 0.5, 'Clear blockage, new filters', false, 0, 1],
    [30, 'Brush not spinning', 'AP-BRSH', 0.75, 'Replace brush roll', true, 0.03, 0.9],
    [30, 'Burning smell, then cuts out', 'AP-VMOT', 1, 'Replace motor', true, 0.15, 0.55],
  ],
  'Food mixer': [
    [50, 'Grinding noise, bowl not turning', 'AP-GEAR', 1, 'Replace drive gears', true, 0.05, 0.85],
    [30, 'Stuck on one speed', 'AP-MCB', 1, 'Replace speed control board', true, 0.08, 0.7],
    [20, 'Mains lead damaged', 'AP-LEAD', 0.5, 'Replace mains lead and test', false, 0, 1],
  ],
  'Sewing machine': [
    [45, 'Skipping stitches', null, 1, 'Service and reset timing', false, 0, 1],
    [35, 'Motor runs, needle does not move', 'AP-BELT', 0.75, 'Replace drive belt', true, 0.03, 0.9],
    [20, 'Thread keeps jamming', 'AP-BOB', 0.75, 'Replace bobbin case and service', true, 0.03, 0.9],
  ],
};
// What a customer says when a repair comes back, by the part that was fitted.
const RETURNS = {
  'PH-SCR61': 'touch not working in one corner', 'PH-SCR67': 'touch not working in one corner', 'PH-BAT1': 'new battery draining fast',
  'PH-BAT2': 'new battery draining fast', 'PH-PORT': 'charging cuts in and out', 'LT-JACK': 'power cuts out when the lead moves',
  'AP-PUMP': 'pump noisy again', 'AP-SEAL': 'leaking again', 'TB-GLS': 'screen lifting at one edge',
};

// The kinds of device booked in each month (0 January): phones up in the
// summer holidays and after new year, laptops when term starts, appliances
// before the year ends.
function kindWeights(month) {
  const w = { Phone: 50, Tablet: 10, Laptop: 24, 'Coffee machine': 6, 'Vacuum cleaner': 5, 'Food mixer': 3, 'Sewing machine': 2 };
  if (month === 6 || month === 7) w.Phone *= 1.3;
  if (month === 0) { w.Phone *= 1.2; w.Tablet *= 1.3; }
  if (month === 8) w.Laptop *= 1.7;
  if (month === 10 || month === 11) { w['Coffee machine'] *= 1.6; w['Food mixer'] *= 1.8; w['Sewing machine'] *= 1.4; }
  return Object.entries(w);
}
const SEASON = [1.15, 0.9, 0.95, 1, 1, 1.05, 1.15, 1.15, 1.1, 0.95, 1, 1.1];

/** Fill an Ironbridge Repairs database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 450 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };
  // Records that belong to a job are kept here and written once, with the
  // job, at the end: one insert each rather than an insert and updates. A
  // value that is a job, quote, invoice or account becomes its record number.
  const later = { 'Part Orders': [], Quotes: [], 'Quote Lines': [], Invoices: [], 'Job Lines': [], Warranties: [], Payments: [] };
  const keep = (table, values) => { later[table].push(values); return values; };
  const ids = new WeakMap();
  // The shop opens Monday to Saturday; anything due on a Sunday happens on the Monday.
  const opened = (d) => (cal.weekday(d) === 0 ? d + 1 : d);
  const slug = (s) => s.toLowerCase().replace(/[^a-z]+/g, '');
  // Document numbers, in date order, a new run each year: JOB-2026-0042.
  const runs = {};
  const numberFor = (prefix, d) => { const y = cal.at(d).getFullYear(); const k = `${prefix}${y}`; runs[k] = (runs[k] || 0) + 1; return `${prefix}-${y}-${String(runs[k]).padStart(4, '0')}`; };
  // A line's amount and the VAT inside it, rounded as the print templates round them (by SQLite, not JavaScript).
  const roundings = new Map();
  const vatSql = base.db.prepare('SELECT ROUND(? * ?, 2) AS gross, ROUND(ROUND(? * ?, 2) - ROUND(? * ?, 2) / (1 + ?), 2) AS vat');
  const amounts = (qty, price, rate) => {
    const k = `${qty}|${price}|${rate}`;
    if (!roundings.has(k)) { const v = vatSql.get(qty, price, qty, price, qty, price, rate); roundings.set(k, { gross: v.gross, vat: v.vat }); }
    return roundings.get(k);
  };
  /** Net, VAT and total of some lines, each line rounded first, as a document adds them. */
  const totalsOf = (lines) => {
    let total = 0; let tax = 0;
    for (const l of lines) { const a = amounts(l.Quantity, l['Unit Price'], l['Tax Rate']); total += a.gross; tax += a.vat; }
    total = money(total); tax = money(tax);
    return { Net: money(total - tax), Tax: tax, Total: total };
  };

  add('Business', businessRecord({
    name: 'Ironbridge Repairs', domain: 'ironbridge.example', phone: '07700 900300',
    address: { Address: '42 Canal Walk', City: 'Bristol', Postcode: 'BS1 4QR', Country: 'United Kingdom' },
    terms: 'Payment on collection, by card or cash. Business accounts: payment within 30 days of the invoice date.',
    taxRate: VAT,
    footer: 'Every repair carries a 90-day warranty on the parts fitted and the work. Open Monday to Saturday, 9 am to 5.30 pm. Thank you for bringing it to us.',
  }));

  // Staff: three at the bench and one at the counter.
  const staff = [
    ['Senior Technician', 'Laptops, desktops and data recovery'],
    ['Technician', 'Phones and tablets, board-level soldering'],
    ['Technician', 'Small appliances and sewing machines, phones when busy'],
    ['Counter Assistant', 'Booking in, quotes and collections'],
  ].map(([role, skills], i) => {
    const p = person(r, { domain: 'ironbridge.example' });
    return add('Staff', { ...p, Role: role, Skills: skills, Phone: `07700 900${String(310 + i).padStart(3, '0')}`, Started: cal.day(-r.int(500, 3200)), Active: true });
  });
  const [laptops, phones, appliances] = staff;
  const technicianFor = (kind) => {
    if (kind === 'Phone' || kind === 'Tablet') return r.chance(0.8) ? phones : appliances;
    if (kind === 'Laptop') return r.chance(0.9) ? laptops : phones;
    return appliances;
  };

  const suppliers = SUPPLIERS.map(([name, supplies, lead]) => {
    const c = person(r);
    return { lead, id: add('Suppliers', {
      Name: name, Contact: `${c['First Name']} ${c['Last Name']}`, Phone: phoneFor(r, 'United Kingdom'), 'E-mail': `trade@${slug(name)}.example`,
      Supplies: supplies, 'Lead Days': lead, 'Our Account': `IBR-${r.int(10000, 99999)}`,
    }) };
  });

  const parts = new Map();
  for (const [code, name, fits, s, cost, price, level, qty, bin] of PARTS) {
    const p = { code, name, cost, price, level, qty, stocked: level > 0, supplier: suppliers[s], stock: 0, reserved: 0, onOrder: false };
    p.id = add('Parts', { Code: code, Name: name, Fits: fits, Supplier: p.supplier.id, Cost: cost, Price: price, 'In Stock': 0, 'Reorder Level': level, 'Reorder Quantity': qty, Stocked: p.stocked, Bin: bin });
    parts.set(code, p);
  }

  // Part orders: placed, then received some working days later (now and then later than promised).
  const pendingOrders = [];
  const orderPart = (p, qty, d, job = null) => {
    // A part not kept on the shelf is a special order and takes the supplier longer.
    const expected = cal.workday(d + p.supplier.lead + (p.stocked ? 0 : 2));
    // The supplier has the larger phone screens on back order: shelf stock
    // ordered in the last three weeks has not come, so that bin runs low
    // (single screens for a job still come from their trade counter).
    const backOrdered = !job && p.code === 'PH-SCR67' && d >= -21;
    const arrive = backOrdered ? Infinity : cal.workday(expected + r.weighted([[0, 7], [1, 2], [2, 1], [4, 0.4]]));
    const row = keep('Part Orders', { Supplier: p.supplier.id, Part: p.id, Job: job, Quantity: qty, 'Unit Cost': p.cost, 'Ordered On': cal.day(d), Expected: cal.day(expected), Status: 'On order' });
    pendingOrders.push({ row, p, qty, job, arrive });
    if (!job) p.onOrder = true;
  };
  // Opening stock, in before the first job.
  for (const p of parts.values()) {
    if (!p.stocked) continue;
    const n = p.qty + p.level;
    keep('Part Orders', { Supplier: p.supplier.id, Part: p.id, Job: null, Quantity: n, 'Unit Cost': p.cost, 'Ordered On': cal.day(-days - 12), Expected: cal.day(cal.workday(-days - 8)), 'Received On': cal.day(cal.workday(-days - 8)), Status: 'Received' });
    p.stock = n;
  }

  // Customers: members of the public as they walk in, and the staff of a few businesses on account.
  const area = () => ({ Address: `${r.int(1, 220)} ${r.pick(['Mill Lane', 'Station Road', 'Orchard Way', 'Church Street', 'Canal Walk', 'Quarry Road', 'Cedar Avenue', 'Linden Close', 'Harbour Street'])}`, City: 'Bristol', Postcode: `BS${r.int(1, 16)} ${r.int(1, 9)}${r.pick(['AB', 'DE', 'FG', 'HJ', 'NP', 'QR', 'TU'])}` });
  const walkins = [];
  const newCustomer = (d) => {
    const p = person(r, { domain: r.pick(['mail.example', 'post.example', 'example.com']) });
    const email = r.chance(0.7);
    const c = { devices: [], account: null };
    c.id = add('Customers', {
      'First Name': p['First Name'], 'Last Name': p['Last Name'], Phone: phoneFor(r, 'United Kingdom'), 'E-mail': email ? p['E-mail'] : null,
      ...(r.chance(0.5) ? area() : {}), 'Contact By': r.weighted([['Text message', 6], ['Phone call', 3], ...(email ? [['E-mail', 2]] : [])]), Joined: cal.day(d),
    });
    walkins.push(c);
    return c;
  };
  const accounts = companies(r, 5, ['Lettings', 'Surveyors', 'Accountants', 'Dental', 'Architects', 'Solicitors']).map((name) => {
    const where = area();
    const domain = `${slug(name)}.example`;
    const since = -days - r.int(30, 700);
    const lead = person(r, { domain });
    const a = { payDays: 30, people: [] };
    a.id = add('Accounts', {
      Name: name, Contact: `${lead['First Name']} ${lead['Last Name']}`, Phone: phoneFor(r, 'United Kingdom'), 'E-mail': `accounts@${domain}`,
      ...where, 'Payment Days': a.payDays, 'Credit Limit': r.pick([500, 1000, 1500]), Opened: cal.day(since),
    });
    // The account's contact and one or two colleagues bring work in.
    for (let k = 0; k < r.int(2, 3); k++) {
      const p = k === 0 ? lead : person(r, { domain });
      const c = { devices: [], account: a };
      c.id = add('Customers', {
        'First Name': p['First Name'], 'Last Name': p['Last Name'], Account: a.id, Phone: phoneFor(r, 'United Kingdom'), 'E-mail': p['E-mail'],
        ...where, 'Contact By': 'E-mail', Joined: cal.day(since + r.int(0, 60)),
      });
      a.people.push(c);
    }
    return a;
  });

  const newDevice = (c, kind) => {
    const [make, model, size] = r.pick(MODELS[kind]);
    const letters = () => String.fromCharCode(65 + r.int(0, 25)) + String.fromCharCode(65 + r.int(0, 25));
    const dev = { kind, size, customer: c };
    dev.id = add('Devices', { Customer: c.id, Kind: kind, Make: make, Model: model, Serial: `${make.slice(0, 2).toUpperCase()}${r.int(10, 99)}${letters()}${r.int(100000, 999999)}`, Colour: r.pick(COLOURS[kind] || APPLIANCE_COLOURS) });
    c.devices.push(dev);
    return dev;
  };
  const partFor = (code, dev) => {
    if (code === 'PH-SCR') return dev.size === 67 ? 'PH-SCR67' : 'PH-SCR61';
    if (code === 'PH-BAT') return dev.size === 67 ? 'PH-BAT2' : 'PH-BAT1';
    if (code === 'LT-SCR') return dev.size === 15 ? 'LT-SCR15' : 'LT-SCR14';
    if (code === 'LT-SSD') return r.chance(0.6) ? 'LT-SSD5' : 'LT-SSD1';
    return code;
  };
  const PROMISE = { Phone: 1, Tablet: 2, Laptop: 3 };
  const repairDelay = (kind) => (kind === 'Phone' ? r.weighted([[0, 55], [1, 35], [2, 10]])
    : kind === 'Tablet' ? r.weighted([[0, 2], [1, 5], [2, 3]])
      : kind === 'Laptop' ? r.weighted([[0, 15], [1, 45], [2, 30], [3, 10]]) : r.weighted([[0, 1], [1, 4], [2, 3.5], [3, 1.5]]));

  const jobs = [];
  const present = [
    { at: -9, kind: 'Laptop', fault: 'Hinge broken, lid loose', slow: true },
    { at: -2, kind: 'Phone', fault: 'Back glass cracked' },
    { at: -1, kind: 'Phone', fault: 'Back glass cracked' },
  ];
  let active = [];
  const returns = [];
  const transfers = [];

  /** Book a job in on day `d`: the device, its fault (one of FAULTS, or chosen by weight) and how it will go. */
  const book = (d, dev, warrantyOf = null, chosen = null) => {
    const c = dev.customer;
    let fault; let part; let hours; let work; let quoted; let beyond; let approve;
    if (warrantyOf) {
      part = warrantyOf.part;
      fault = `Warranty return of ${warrantyOf.number}: ${RETURNS[part?.code] || 'the same fault has come back'}`;
      hours = 0.5; work = 'Warranty repair'; quoted = false; beyond = 0; approve = 1;
    } else {
      const f = chosen || r.weighted(FAULTS[dev.kind].map((x) => [x, x[0]]));
      [, fault, part, hours, work, quoted, beyond, approve] = f;
      part = part ? parts.get(partFor(part, dev)) : null;
    }
    const price = warrantyOf ? 0 : money((part ? part.price : 0) + hours * RATE);
    const promise = opened(d + (PROMISE[dev.kind] || 5) + (quoted ? 1 : 0));
    const job = {
      number: numberFor('JOB', d), dev, c, d, part, hours, work, quoted, beyond, approve, price, warrantyOf, tech: warrantyOf ? warrantyOf.tech : technicianFor(dev.kind),
      status: 'Booked in', quote: null, quotedOn: null, approved: false, deposit: null, outcome: null, charge: null,
      finished: null, collected: null, paid: false, warrantyUntil: warrantyOf ? warrantyOf.warrantyUntil : null, notes: [], lines: [],
      step: 'start', at: d + (r.chance(0.65) ? 0 : 1),
    };
    if (!quoted) { job.quote = price; job.quotedOn = d; job.approved = true; }
    if (warrantyOf) job.notes.push(`Covered by ${warrantyOf.warranty.Number}.`);
    if (r.chance(0.08)) job.notes.push(r.pick(['Customer has a backup of the data.', 'Passcode not given: display tested only.', 'Collect after 5 pm.', 'Second visit for this fault.', 'Customer wants the old part back.']));
    job.row = { Number: job.number, Device: dev.id, Technician: job.tech, 'Booked In': cal.day(d), Fault: fault, 'Left With It': r.weighted(LEFT_WITH[dev.kind]), Promised: cal.day(promise) };
    jobs.push(job);
    active.push(job);
    return job;
  };

  /** A line on the job, a part or labour, at its price with VAT included. */
  const line = (job, kind, part, description, quantity, price, done) => {
    const row = keep('Job Lines', { Job: job, Invoice: null, Kind: kind, Part: part ? part.id : null, Description: description, Quantity: quantity, 'Unit Price': price, 'Tax Rate': VAT, 'Done On': done == null ? null : cal.day(done) });
    job.lines.push(row);
    return row;
  };
  const labour = (job, d, description, hours, rate) => line(job, 'Labour', null, description, hours, rate, d);
  const collectIn = (job, d) => {
    const wait = r.weighted([[0, 30], [1, 18], [2, 14], [3, 10], [5, 8], [8, 7], [12, 4], [r.int(14, 45), 5], [Infinity, 0.6]]);
    job.step = 'collect';
    // A device never fetched is kept a few months, then the customer comes for it after the final reminder.
    job.at = wait !== Infinity ? opened(d + wait) : d > -100 ? Infinity : opened(d + r.int(60, 95));
  };
  /** The work is done: the job is invoiced the same day when it costs anything. */
  const finish = (job, d, status, outcome, charge) => {
    job.status = status; job.outcome = outcome; job.charge = charge; job.finished = d;
    if (charge > 0) {
      const t = totalsOf(job.lines);
      const deposit = job.deposit || 0;
      const a = job.c.account;
      job.invoice = keep('Invoices', {
        Number: numberFor('INV', d), Job: job, Account: a, 'Invoice Date': cal.day(d), 'Due Date': cal.day(a ? d + a.payDays : d),
        ...t, Deposit: deposit, Balance: money(t.Total - deposit), 'Paid On': null,
      });
      for (const l of job.lines) l.Invoice = job.invoice;
    }
    collectIn(job, d);
  };

  /** The customer has said yes (or the price was fixed): set the part aside, or order it. */
  const goAhead = (job, d) => {
    job.approved = true;
    if (job.quoteRow) { job.quoteRow.Answer = 'Accepted'; job.quoteRow['Answered On'] = cal.day(d); }
    if (!job.part) { job.status = 'Repairing'; job.step = 'repair'; job.at = d + repairDelay(job.dev.kind); return; }
    const p = job.part;
    job.jobPart = line(job, 'Part', p, p.name, 1, job.warrantyOf ? 0 : p.price, null);
    if (p.stock - p.reserved >= 1) {
      p.reserved += 1;
      job.status = 'Repairing'; job.step = 'repair'; job.at = d + repairDelay(job.dev.kind);
      return;
    }
    orderPart(p, 1, d, job);
    job.status = 'Awaiting parts'; job.step = 'wait'; job.at = Infinity;
    // A part ordered in specially is paid for up front, unless the customer is on account.
    if (!p.stocked && !job.c.account && !job.warrantyOf) {
      job.deposit = p.price;
      keep('Payments', { Number: numberFor('RCP', d), Job: job, 'Paid On': cal.day(d), Amount: p.price, Method: r.chance(0.8) ? 'Card' : 'Cash', Kind: 'Deposit' });
    }
  };

  const steps = {
    start(job, d) {
      if (!job.quoted) return goAhead(job, d);
      job.status = 'Diagnosing'; job.step = 'diagnose'; job.at = d + r.weighted([[0, 3], [1, 4], [2, 2]]);
    },
    diagnose(job, d) {
      job.diagnosed = d;
      if (r.chance(job.beyond)) {
        // No fix, no fee: the check is recorded but not charged.
        labour(job, d, 'Diagnosis, no fix no fee', DIAGNOSIS, 0);
        return finish(job, d, 'Unrepairable', 'Beyond economic repair', 0);
      }
      job.quote = job.price; job.quotedOn = d; job.status = 'Awaiting approval';
      // The written quote: the part and the labour, VAT included, good for thirty days.
      const quote = keep('Quotes', { Number: numberFor('QUO', d), Job: job, 'Quoted On': cal.day(d), 'Valid Until': cal.day(d + 30), Net: 0, Tax: 0, Total: 0, Answer: null, 'Answered On': null });
      const lines = [];
      if (job.part) lines.push(keep('Quote Lines', { Quote: quote, Kind: 'Part', Description: job.part.name, Quantity: 1, 'Unit Price': job.part.price, 'Tax Rate': VAT }));
      lines.push(keep('Quote Lines', { Quote: quote, Kind: 'Labour', Description: job.work, Quantity: job.hours, 'Unit Price': RATE, 'Tax Rate': VAT }));
      Object.assign(quote, totalsOf(lines));
      job.quoteRow = quote;
      job.step = 'decide'; job.at = job.slow ? d + r.int(12, 16) : d + r.weighted([[0, 10], [1, 25], [2, 22], [3, 15], [r.int(4, 6), 15], [r.int(7, 14), 13]]);
    },
    decide(job, d) {
      if (r.chance(job.approve)) return goAhead(job, d);
      job.quoteRow.Answer = 'Declined'; job.quoteRow['Answered On'] = cal.day(d);
      labour(job, job.diagnosed, 'Diagnosis and quote', DIAGNOSIS, RATE);
      return finish(job, d, 'Ready', 'Quote declined', money(DIAGNOSIS * RATE));
    },
    repair(job, d) {
      const p = job.part;
      if (p) { p.stock -= 1; p.reserved -= 1; job.jobPart['Done On'] = cal.day(d); }
      const rate = job.warrantyOf ? 0 : RATE;
      if (job.quoted) { labour(job, job.diagnosed, 'Diagnosis', 0.25, rate); labour(job, d, job.work, job.hours - 0.25, rate); }
      else labour(job, d, job.work, job.hours, rate);
      if (!job.warrantyOf) {
        job.warrantyUntil = d + WARRANTY_DAYS;
        job.warranty = keep('Warranties', { Number: numberFor('WTY', d), Job: job, Issued: cal.day(d), Covers: p ? `${p.name}, fitted, and the work: ${job.work.toLowerCase()}.` : `The work: ${job.work.toLowerCase()}.` });
      }
      finish(job, d, 'Ready', 'Repaired', job.warrantyOf ? 0 : job.price);
    },
    collect(job, d) {
      job.collected = d;
      if (job.status !== 'Unrepairable') job.status = 'Collected';
      const due = money((job.charge || 0) - (job.deposit || 0));
      if (due <= 0) {
        job.paid = true;
        if (job.invoice) job.invoice['Paid On'] = cal.day(d);
      } else if (job.c.account) transfers.push({ job, at: cal.workday(d + r.int(10, 45)), amount: due });
      else {
        keep('Payments', { Number: numberFor('RCP', d), Job: job, 'Paid On': cal.day(d), Amount: due, Method: r.weighted([['Card', 72], ['Cash', 24], ['Bank transfer', 4]]), Kind: 'On collection' });
        job.paid = true;
        job.invoice['Paid On'] = cal.day(d);
      }
      // Now and then a repair comes back inside its warranty.
      if (job.outcome === 'Repaired' && !job.warrantyOf && r.chance(0.035)) returns.push({ at: Math.min(d + r.int(4, 80), job.warrantyUntil - 1), job });
      job.step = 'done'; job.at = Infinity;
    },
  };

  for (let d = -days; d <= 0; d++) {
    const wd = cal.weekday(d);
    if (wd === 0) continue;

    // Parts arriving today: shelf stock, or the part a job has been waiting for.
    for (let i = pendingOrders.length - 1; i >= 0; i--) {
      const o = pendingOrders[i];
      if (o.arrive > d) continue;
      pendingOrders.splice(i, 1);
      o.row['Received On'] = cal.day(d); o.row.Status = 'Received';
      o.p.stock += o.qty;
      if (o.job) { o.p.reserved += o.qty; o.job.status = 'Repairing'; o.job.step = 'repair'; o.job.at = d + r.weighted([[0, 1], [1, 2]]); }
      else o.p.onOrder = false;
    }
    // Business accounts paying their invoices by transfer.
    for (let i = transfers.length - 1; i >= 0; i--) {
      const t = transfers[i];
      if (t.at > d) continue;
      transfers.splice(i, 1);
      keep('Payments', { Number: numberFor('RCP', d), Job: t.job, 'Paid On': cal.day(d), Amount: t.amount, Method: 'Bank transfer', Kind: 'On account' });
      t.job.paid = true;
      t.job.invoice['Paid On'] = cal.day(d);
    }
    // Warranty returns brought back today.
    for (let i = returns.length - 1; i >= 0; i--) {
      if (returns[i].at > d) continue;
      const { job } = returns.splice(i, 1)[0];
      book(d, job.dev, job);
    }

    // This week's jobs, whatever day the file is made: a customer slow to
    // answer a quote booked nine days ago, and back glass booked in over
    // the last two days, waiting on a special-order part.
    for (const x of present) {
      if (x.done || d < x.at) continue;
      x.done = true;
      const c = newCustomer(d);
      const job = book(d, newDevice(c, x.kind), null, FAULTS[x.kind].find((f) => f[1] === x.fault));
      job.beyond = 0;
      job.approve = 1;
      if (x.slow) job.slow = true;
    }

    // New work at the counter: fewer on a Saturday, more in the busy months.
    const month = cal.at(d).getMonth();
    const n = Math.round((wd === 6 ? 4.5 : 6) * SEASON[month] * (0.6 + r.next() * 0.8));
    const kinds = kindWeights(month);
    for (let k = 0; k < n; k++) {
      let c;
      let dev;
      if (r.chance(0.06)) {
        c = r.pick(r.pick(accounts).people);
        dev = c.devices.length && r.chance(0.4) ? r.pick(c.devices) : newDevice(c, r.chance(0.8) ? 'Laptop' : 'Phone');
      } else {
        c = walkins.length > 40 && r.chance(0.25) ? r.pick(walkins) : newCustomer(d);
        dev = c.devices.length && r.chance(0.5) ? r.pick(c.devices) : newDevice(c, r.weighted(kinds));
      }
      // A device already in the shop is not booked in again.
      if (active.some((j) => j.dev === dev)) dev = newDevice(c, r.weighted(kinds));
      book(d, dev);
    }

    // The bench: move each job on as far as today allows.
    for (const job of active) while (job.at <= d) steps[job.step](job, d);
    active = active.filter((j) => j.step !== 'done');

    // Each Monday: reorder shelf parts at or below their level.
    if (wd === 1) {
      for (const p of parts.values()) if (p.stocked && !p.onOrder && p.stock - p.reserved <= p.level) orderPart(p, p.qty, d);
    }
  }

  // Each job as it stands today.
  for (const job of jobs) {
    if (job.finished != null && job.collected == null && -job.finished >= 14) job.notes.push(`Reminder sent by ${job.c.account ? 'e-mail' : 'text'} on ${cal.day(job.finished + 7)}.`);
    if (job.finished != null && job.collected == null && -job.finished >= 35) job.notes.push(`Final reminder sent on ${cal.day(job.finished + 28)}.`);
    job.id = add('Jobs', {
      ...job.row, 'Warranty Of': job.warrantyOf ? job.warrantyOf.id : null, Status: job.status, Quote: job.quote, 'Quoted On': job.quotedOn == null ? null : cal.day(job.quotedOn), 'Quote Approved': job.approved,
      Deposit: job.deposit, Outcome: job.outcome, Charge: job.charge, Finished: job.finished == null ? null : cal.day(job.finished),
      Collected: job.collected == null ? null : cal.day(job.collected), Paid: job.paid,
      'Warranty Until': job.warrantyUntil == null ? null : cal.day(job.warrantyUntil), Notes: job.notes.length ? job.notes.join(' ') : null,
    });
  }
  // Then what hangs off the jobs, quotes and invoices before the lines that point at them.
  for (const [table, rows] of Object.entries(later)) {
    for (const row of rows) {
      const values = {};
      for (const [k, v] of Object.entries(row)) values[k] = v !== null && typeof v === 'object' ? (v.id ?? ids.get(v)) : v;
      ids.set(row, add(table, values));
    }
  }
  for (const p of parts.values()) base.update('Parts', p.id, { 'In Stock': p.stock });
  return counts;
}
