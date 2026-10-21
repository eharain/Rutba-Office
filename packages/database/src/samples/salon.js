// The salon: Marigold Hair and Beauty, an invented high-street salon with
// four stylists, a beauty therapist and a junior, open Tuesday to Saturday
// with a late night on Thursday. Its clients, the team and what each earns
// in commission, the menu of services priced by the stylist's level, every
// booking and the services done in it, retail and back-bar products with
// their suppliers and deliveries, what is sold over the counter, and gift
// vouchers sold and redeemed.
//
// Its records are thirteen months of the appointment book, run a day at a
// time, and six weeks of bookings ahead: clients come back on their own
// rhythm (a gents cut every few weeks, a root tint every six), some drift
// away, one stylist leaves and most of her clients move to the stylist who
// replaces her, December is full and January quiet. Colour and tint need a
// patch test in the six months before, a few clients do not turn up, and
// stock goes down as it is sold and used at the basin until the weekly order
// brings it back. The questions a salon owner asks at the desk and at the
// month's end are its saved queries; its reports print them.
//
// Prices are shown to clients with VAT in them, as a salon's menu is, so a
// receipt shows the VAT within each line. The salon prints what a salon
// gives out: a receipt at the desk for each visit, gift vouchers, a
// commission statement for each stylist each month, and a letter
// confirming each appointment.

import { BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, companies, FIRST, LAST, money } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const CLIENT = '[Clients].[First Name] & " " & [Clients].[Last Name]';
const STYLIST = '[Staff].[First Name] & " " & [Staff].[Last Name]';
const WHEN = 'Left([Start], 10) & " " & Mid([Start], 12, 5)';
// A date as the desk writes it: 24/10/2026.
const UK_DATE = (f) => `Format([${f}], "dd") & "/" & Format([${f}], "mm") & "/" & Format([${f}], "yyyy")`;
// A receipt line's amount, and the VAT within it: prices are shown with VAT in.
const AMOUNT = 'Round([Quantity] * [Unit Price], 2)';
const VAT_IN = `Round(${AMOUNT} - ${AMOUNT} / (1 + [Tax Rate]), 2)`;

export const spec = {
  tables: [
    BUSINESS_TABLE,
    { name: 'Suppliers', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Contact', type: 'text' }, { name: 'Phone', type: 'text' },
      { name: 'E-mail', type: 'text' }, { name: 'Account Number', type: 'text' }, { name: 'Lead Days', type: 'integer', default: 3 },
    ] },
    { name: 'Staff', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Role', type: 'text' },
      { name: 'Level', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Works', type: 'text' },
      { name: 'Started', type: 'date' }, { name: 'Left', type: 'date' },
      { name: 'Service Commission', type: 'number', default: 0 }, { name: 'Retail Commission', type: 'number', default: 0 },
      { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Services', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Category', type: 'text' }, { name: 'Minutes', type: 'integer' },
      { name: 'Junior Price', type: 'currency' }, { name: 'Stylist Price', type: 'currency' }, { name: 'Senior Price', type: 'currency' },
      { name: 'Patch Test', type: 'yesno' }, { name: 'Tax Rate', type: 'number', default: 0.2 }, { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Clients', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Phone', type: 'text' },
      { name: 'E-mail', type: 'text' }, { name: 'Heard Of Us', type: 'text' }, { name: 'Client Since', type: 'date' },
      { name: 'Patch Test Date', type: 'date' }, { name: 'Allergies', type: 'text' }, { name: 'Colour Notes', type: 'text' },
      { name: 'Marketing Consent', type: 'yesno' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Bookings', fields: [
      { name: 'Client', ...ref('Clients'), required: true }, { name: 'Staff', ...ref('Staff'), required: true },
      { name: 'Start', type: 'datetime' }, { name: 'Minutes', type: 'integer' }, { name: 'Status', type: 'text', default: 'Booked' },
      { name: 'Booked On', type: 'date' }, { name: 'Booked Via', type: 'text' }, { name: 'Deposit', type: 'currency', default: 0 },
      { name: 'Deposit Refunded', type: 'yesno' }, { name: 'Cancelled On', type: 'date' }, { name: 'Paid By', type: 'text' },
      { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Booking Services', fields: [
      { name: 'Booking', ...ref('Bookings'), required: true }, { name: 'Service', ...ref('Services'), required: true },
      { name: 'Price', type: 'currency' },
    ] },
    { name: 'Products', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Kind', type: 'text', default: 'Retail' }, { name: 'Category', type: 'text' },
      { name: 'Supplier', ...ref('Suppliers') }, { name: 'Size', type: 'text' }, { name: 'Cost Price', type: 'currency' },
      { name: 'Sale Price', type: 'currency' }, { name: 'Tax Rate', type: 'number', default: 0.2 }, { name: 'In Stock', type: 'integer', default: 0 },
      { name: 'Reorder Level', type: 'integer', default: 4 }, { name: 'Reorder Quantity', type: 'integer', default: 12 },
      { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Product Sales', fields: [
      { name: 'Date', type: 'date' }, { name: 'Staff', ...ref('Staff') }, { name: 'Client', ...ref('Clients') },
      { name: 'Product', ...ref('Products'), required: true }, { name: 'Quantity', type: 'integer', default: 1 },
      { name: 'Unit Price', type: 'currency' }, { name: 'Receipt', ...ref('Receipts') },
    ] },
    { name: 'Stock Deliveries', fields: [
      { name: 'Supplier', ...ref('Suppliers'), required: true }, { name: 'Product', ...ref('Products'), required: true },
      { name: 'Received', type: 'date' }, { name: 'Quantity', type: 'integer' }, { name: 'Unit Cost', type: 'currency' },
      { name: 'Invoice Number', type: 'text' },
    ] },
    { name: 'Gift Vouchers', fields: [
      { name: 'Code', type: 'text', required: true, unique: true }, { name: 'Sold On', type: 'date' }, { name: 'Bought By', type: 'text' },
      { name: 'For', type: 'text' }, { name: 'Value', type: 'currency' }, { name: 'Expires', type: 'date' },
      { name: 'Redeemed On', type: 'date' }, { name: 'Booking', ...ref('Bookings') },
    ] },
    // The till's record of each visit paid for: its own copy of what was charged, kept as it was
    // printed, so a receipt reprinted next year says what the client was given on the day.
    { name: 'Receipts', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Booking', ...ref('Bookings') }, { name: 'Taken', type: 'datetime' },
      { name: 'Net', type: 'currency' }, { name: 'VAT', type: 'currency' }, { name: 'Total', type: 'currency' },
      { name: 'Deposit Paid', type: 'currency', default: 0 }, { name: 'Voucher', type: 'currency', default: 0 }, { name: 'Voucher Code', type: 'text' },
      { name: 'Paid Today', type: 'currency' }, { name: 'Paid By', type: 'text' },
    ] },
    { name: 'Receipt Lines', fields: [
      { name: 'Receipt', ...ref('Receipts'), required: true }, { name: 'Kind', type: 'text', default: 'Service' }, { name: 'Item', type: 'text' },
      { name: 'Quantity', type: 'integer', default: 1 }, { name: 'Unit Price', type: 'currency' }, { name: 'Tax Rate', type: 'number', default: 0.2 },
    ] },
    // A statement for each person on commission each month, its lines the takings by kind of work.
    { name: 'Commission Statements', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Staff', ...ref('Staff'), required: true }, { name: 'Month', type: 'text' },
      { name: 'Prepared On', type: 'date' }, { name: 'Commission Due', type: 'currency' }, { name: 'Paid On', type: 'date' },
    ] },
    { name: 'Commission Lines', fields: [
      { name: 'Statement', ...ref('Commission Statements'), required: true }, { name: 'Line', type: 'integer' }, { name: 'Work', type: 'text' },
      { name: 'Done', type: 'integer' }, { name: 'Takings', type: 'currency' }, { name: 'Rate', type: 'number' }, { name: 'Commission', type: 'currency' },
    ] },
  ],
  queries: {
    'Bookings next 7 days': { tables: ['Bookings', 'Clients', 'Staff'], columns: [
      { expr: WHEN, as: 'When' }, { table: 'Bookings', field: 'Start', sort: 'asc', show: false },
      { expr: STYLIST, as: 'Stylist' }, { expr: CLIENT, as: 'Client' }, { table: 'Clients', field: 'Phone' },
      { table: 'Bookings', field: 'Minutes' }, { table: 'Bookings', field: 'Deposit' }, { table: 'Bookings', field: 'Booked Via' },
      { expr: 'DateDiff("d", Date(), [Start])', as: 'Days Away', criteria: ['Between 0 And 7'], show: false },
      { table: 'Bookings', field: 'Status', criteria: ['Booked'], show: false },
    ] },
    'Takings by stylist last month': { tables: ['Bookings', 'Staff', 'Booking Services'], columns: [
      { expr: STYLIST, as: 'Stylist', total: 'group' }, { table: 'Staff', field: 'ID', total: 'group', show: false },
      { table: 'Staff', field: 'Role', total: 'group' }, { table: 'Booking Services', field: 'ID', as: 'Services Done', total: 'count' },
      { table: 'Booking Services', field: 'Price', as: 'Takings', total: 'sum', sort: 'desc' },
      { table: 'Bookings', field: 'Status', total: 'where', criteria: ['Completed'] },
      { expr: 'DateDiff("m", [Start], Date())', as: 'Months Ago', total: 'where', criteria: ['1'] },
    ] },
    'Service commission last month': { tables: ['Booking Services', 'Bookings', 'Staff', 'Clients', 'Services'], columns: [
      { expr: STYLIST, as: 'Stylist', sort: 'asc' }, { expr: 'Format([Start], "yyyy-mm-dd")', as: 'Day', sort: 'asc' },
      { table: 'Bookings', field: 'Start', sort: 'asc', show: false }, { expr: CLIENT, as: 'Client' },
      { table: 'Services', field: 'Name', as: 'Service' }, { table: 'Booking Services', field: 'Price' },
      { table: 'Staff', field: 'Service Commission', as: 'Rate', criteria: ['> 0'] },
      { expr: 'Round([Price] * [Service Commission], 2)', as: 'Commission' },
      { table: 'Bookings', field: 'Status', criteria: ['Completed'], show: false },
      { expr: 'DateDiff("m", [Start], Date())', as: 'Months Ago', criteria: ['1'], show: false },
    ] },
    'Retail commission last month': { tables: ['Product Sales', 'Staff', 'Products'], columns: [
      { expr: STYLIST, as: 'Stylist', total: 'group', sort: 'asc' }, { table: 'Staff', field: 'ID', total: 'group', show: false },
      { table: 'Product Sales', field: 'Quantity', as: 'Items Sold', total: 'sum' },
      { expr: 'Round([Quantity] * [Unit Price], 2)', as: 'Retail Sales', total: 'sum' },
      { table: 'Staff', field: 'Retail Commission', as: 'Rate', total: 'group', criteria: ['> 0'] },
      { expr: 'Round([Quantity] * [Unit Price] * [Retail Commission], 2)', as: 'Commission', total: 'sum' },
      { expr: 'DateDiff("m", [Date], Date())', as: 'Months Ago', total: 'where', criteria: ['1'] },
    ] },
    'Service takings by month': { tables: ['Bookings', 'Booking Services'], columns: [
      { expr: 'Format([Start], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { table: 'Booking Services', field: 'ID', as: 'Services Done', total: 'count' },
      { table: 'Booking Services', field: 'Price', as: 'Takings', total: 'sum' },
      { table: 'Bookings', field: 'Status', total: 'where', criteria: ['Completed'] },
    ] },
    'Retail sales by month': { tables: ['Product Sales', 'Products'], columns: [
      { expr: 'Format([Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { table: 'Product Sales', field: 'Quantity', as: 'Items Sold', total: 'sum' },
      { expr: 'Round([Quantity] * [Unit Price], 2)', as: 'Sales', total: 'sum' },
      { expr: 'Round([Quantity] * ([Unit Price] - [Cost Price]), 2)', as: 'Profit', total: 'sum' },
    ] },
    'Clients to win back': { tables: ['Clients', 'Bookings'], columns: [
      { table: 'Clients', field: 'ID', total: 'group', show: false }, { expr: CLIENT, as: 'Client', total: 'group' },
      { table: 'Clients', field: 'Phone', total: 'group' }, { table: 'Clients', field: 'E-mail', total: 'group' },
      { expr: 'Format([Start], "yyyy-mm-dd")', as: 'Last Visit', total: 'max' },
      { table: 'Bookings', field: 'ID', as: 'Visits', total: 'count' },
      { expr: 'DateDiff("d", [Start], Date())', as: 'Days Since', total: 'min', criteria: ['Between 91 And 365'], sort: 'asc' },
      { table: 'Bookings', field: 'Status', total: 'where', criteria: ['In ("Completed", "Booked")'] },
      { table: 'Clients', field: 'Marketing Consent', total: 'where', criteria: ['Yes'] },
    ] },
    'Most booked services': { tables: ['Booking Services', 'Services', 'Bookings'], columns: [
      { table: 'Services', field: 'Name', as: 'Service', total: 'group' }, { table: 'Services', field: 'Category', total: 'group' },
      { table: 'Booking Services', field: 'ID', as: 'Times Done', total: 'count', sort: 'desc' },
      { table: 'Booking Services', field: 'Price', as: 'Takings', total: 'sum' },
      { table: 'Bookings', field: 'Status', total: 'where', criteria: ['Completed'] },
      { expr: 'DateDiff("d", [Start], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Products to order': { tables: ['Products', 'Suppliers'], columns: [
      { table: 'Suppliers', field: 'Name', as: 'Supplier', sort: 'asc' }, { table: 'Products', field: 'Name', as: 'Product', sort: 'asc' },
      { table: 'Products', field: 'Kind' }, { table: 'Products', field: 'Size' }, { table: 'Products', field: 'In Stock' },
      { table: 'Products', field: 'Reorder Level' }, { table: 'Products', field: 'Reorder Quantity' },
      { expr: 'Round([Reorder Quantity] * [Cost Price], 2)', as: 'Order Cost' },
      { expr: '[In Stock] - [Reorder Level]', as: 'Above Reorder Level', criteria: ['<= 0'], show: false },
      { table: 'Products', field: 'Active', criteria: ['Yes'], show: false },
    ] },
    'No-shows by client': { tables: ['Bookings', 'Clients'], columns: [
      { table: 'Clients', field: 'ID', total: 'group', show: false }, { expr: CLIENT, as: 'Client', total: 'group' },
      { table: 'Clients', field: 'Phone', total: 'group' },
      { expr: 'IIf([Status] = "No-show", 1, 0)', as: 'No-shows', total: 'sum', criteria: ['>= 1'], sort: 'desc' },
      { expr: 'IIf([Status] = "Cancelled" And DateDiff("d", [Cancelled On], [Start]) < 2, 1, 0)', as: 'Late Cancellations', total: 'sum' },
      { expr: 'IIf([Deposit Refunded], 0, Nz([Deposit]))', as: 'Deposits Kept', total: 'sum' },
      { table: 'Bookings', field: 'Status', total: 'where', criteria: ['In ("No-show", "Cancelled")'] },
    ] },
    'Patch tests due': { tables: ['Bookings', 'Clients', 'Booking Services', 'Services'], columns: [
      { table: 'Bookings', field: 'Start', total: 'group', sort: 'asc', show: false }, { expr: WHEN, as: 'When', total: 'group' },
      { expr: CLIENT, as: 'Client', total: 'group' }, { table: 'Clients', field: 'Phone', total: 'group' },
      { table: 'Clients', field: 'Patch Test Date', as: 'Last Patch Test', total: 'group' },
      { table: 'Services', field: 'Name', as: 'Service', total: 'min' },
      { table: 'Services', field: 'Patch Test', total: 'where', criteria: ['Yes'] },
      { table: 'Bookings', field: 'Status', total: 'where', criteria: ['Booked'] },
      { expr: 'DateDiff("d", Date(), [Start])', as: 'Days Away', total: 'where', criteria: ['Between 0 And 14'] },
      { expr: 'DateDiff("d", Nz([Patch Test Date], #2000-01-01#), [Start])', as: 'Days Since Test', total: 'where', criteria: ['> 180'] },
    ] },
    'Vouchers outstanding': { tables: ['Gift Vouchers'], columns: [
      { table: 'Gift Vouchers', field: 'Code' }, { table: 'Gift Vouchers', field: 'Sold On' }, { table: 'Gift Vouchers', field: 'Bought By' },
      { table: 'Gift Vouchers', field: 'For' }, { table: 'Gift Vouchers', field: 'Value' }, { table: 'Gift Vouchers', field: 'Expires', sort: 'asc' },
      { expr: 'DateDiff("d", Date(), [Expires])', as: 'Days Left', criteria: ['>= 0'] },
      { table: 'Gift Vouchers', field: 'Redeemed On', criteria: ['Is Null'], show: false },
    ] },
    'New clients by source': { tables: ['Clients'], columns: [
      { table: 'Clients', field: 'Heard Of Us', total: 'group' }, { table: 'Clients', field: 'ID', as: 'New Clients', total: 'count', sort: 'desc' },
      { expr: 'DateDiff("d", [Client Since], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Retail best sellers': { tables: ['Product Sales', 'Products'], columns: [
      { table: 'Products', field: 'Name', as: 'Product', total: 'group' }, { table: 'Products', field: 'Category', total: 'group' },
      { table: 'Product Sales', field: 'Quantity', as: 'Sold', total: 'sum' },
      { expr: 'Round([Quantity] * [Unit Price], 2)', as: 'Sales', total: 'sum' },
      { expr: 'Round([Quantity] * ([Unit Price] - [Cost Price]), 2)', as: 'Profit', total: 'sum', sort: 'desc' },
      { expr: 'DateDiff("d", [Date], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Service menu': { tables: ['Services'], columns: [
      { table: 'Services', field: 'Category', sort: 'asc' }, { table: 'Services', field: 'Name', as: 'Service' },
      { table: 'Services', field: 'Minutes' }, { table: 'Services', field: 'Junior Price' }, { table: 'Services', field: 'Stylist Price' },
      { table: 'Services', field: 'Senior Price' }, { table: 'Services', field: 'Active', criteria: ['Yes'], show: false },
    ] },
    'VAT by month': { tables: ['Receipts'], columns: [
      { expr: 'Format([Taken], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { table: 'Receipts', field: 'ID', as: 'Receipts', total: 'count' },
      { table: 'Receipts', field: 'Total', as: 'Takings', total: 'sum' },
      { table: 'Receipts', field: 'Net', as: 'Net', total: 'sum' },
      { table: 'Receipts', field: 'VAT', as: 'VAT', total: 'sum' },
    ] },
    'Receipts last 7 days': { tables: ['Receipts', 'Bookings', 'Clients', 'Staff'], columns: [
      { table: 'Receipts', field: 'Number', sort: 'asc' }, { expr: WHEN.replace(/Start/g, 'Taken'), as: 'Taken At' },
      { expr: CLIENT, as: 'Client' }, { expr: STYLIST, as: 'Stylist' }, { table: 'Receipts', field: 'Total' },
      { table: 'Receipts', field: 'Paid Today' }, { table: 'Receipts', field: 'Paid By' },
      { expr: 'DateDiff("d", [Taken], Date())', as: 'Days Ago', criteria: ['Between 0 And 7'], show: false },
    ] },
  },
  documents: [
    // The till receipt handed over at the desk, on the 80 mm roll. Prices have VAT in them; the VAT is shown within.
    { name: 'Receipt', table: 'Receipts', paper: 'receipt-80', letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'Website', 'Tax Number'] }, blocks: [
      { type: 'heading', text: 'Receipt {Number}' },
      { type: 'fields', columns: 1, items: [
        { field: 'Taken', label: 'Date' }, { field: 'Booking.Client', label: 'Client' }, { field: 'Booking.Staff', label: 'Stylist' },
      ] },
      { type: 'lines', table: 'Receipt Lines', columns: [
        { field: 'Item' }, { field: 'Quantity', label: 'Qty', type: 'integer' }, { field: 'Unit Price', label: 'Price', type: 'currency' },
        { expr: AMOUNT, as: 'Amount', type: 'currency' }, { expr: 'Nz([Tax Rate]) * 100', as: 'VAT %', type: 'number' },
        { expr: VAT_IN, as: 'VAT In', label: 'VAT', type: 'currency' },
      ] },
      { type: 'totals', rows: [
        { label: 'Total', expr: '[Amount]', strong: true }, { label: 'Of which VAT', expr: '[VAT In]' },
        { label: 'Less deposit paid', expr: '[Deposit Paid]' }, { label: 'Less gift voucher', expr: '[Voucher]' },
        { label: 'Paid today', expr: '[Total] - [Deposit Paid] - [Voucher]', strong: true },
      ] },
      { type: 'text', text: 'Paid by {Paid By}. VAT number {Business.Tax Number}.', small: true },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    // The gift voucher, printed on A5 card and put in an envelope.
    { name: 'Gift voucher', table: 'Gift Vouchers', paper: 'a5', letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'Website'] }, blocks: [
      { type: 'heading', text: 'Gift voucher' },
      { type: 'text', text: 'For {For}, from {Bought By}.' },
      { type: 'fields', columns: 2, items: [
        { field: 'Value', label: 'Worth' }, { field: 'Code', label: 'Voucher number' },
        { field: 'Sold On', label: 'Issued' }, { field: 'Expires', label: 'Use by' },
      ] },
      { type: 'text', text: 'To spend on any service or product at {Business.Name}. Call {Business.Phone} or book at {Business.Website}, and give the voucher number when you book.' },
      { type: 'spacer', size: 12 },
      { type: 'text', text: 'Terms. Valid until the use-by date shown, twelve months from the day it was bought. To be used in full at one visit; it cannot be exchanged for cash and no change is given. Please bring the voucher with you. A lost voucher can be replaced only if we can find its number in our records.', small: true },
    ] },
    // A statement for each person on commission for each month, given out with their pay.
    { name: 'Commission statement', table: 'Commission Statements', paper: 'a4', letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone'] }, blocks: [
      { type: 'heading', text: 'Commission statement {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Staff', label: 'For' }, { field: 'Month', label: 'Month' },
        { field: 'Staff.Role', label: 'Role' }, { field: 'Prepared On', label: 'Prepared' },
        { field: 'Staff.Level', label: 'Level' }, { field: 'Paid On', label: 'Paid with wages on' },
      ] },
      { type: 'lines', table: 'Commission Lines', sort: 'Line', columns: [
        { field: 'Work' }, { field: 'Done', label: 'Done', type: 'integer' }, { field: 'Takings', type: 'currency' },
        { expr: 'Nz([Rate]) * 100', as: 'Rate %', type: 'number' }, { field: 'Commission', type: 'currency' },
      ] },
      { type: 'totals', rows: [
        { label: 'Takings', expr: '[Takings]' }, { label: 'Commission due', expr: '[Commission]', strong: true },
      ] },
      { type: 'text', text: 'Commission is worked out on what clients paid, VAT included, for completed appointments and retail you sold in the month. It is paid with your wages and taxed with them. If anything here looks wrong, please speak to the owner within 14 days.', small: true },
    ] },
    // The letter (or e-mail) confirming an appointment: every booking prints one; choose the ones to send.
    { name: 'Appointment confirmation', table: 'Bookings', paper: 'a4', letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Website'] }, blocks: [
      { type: 'heading', text: 'Your appointment' },
      { type: 'text', text: 'Dear {Client.First Name},' },
      { type: 'text', text: 'Thank you for booking with us. This is to confirm your appointment, booked on {Booked On}.' },
      { type: 'fields', columns: 2, items: [
        { field: 'Start', label: 'Date and time' }, { field: 'Staff.First Name', label: 'With' },
        { field: 'Minutes', label: 'Time set aside (minutes)' }, { field: 'Deposit', label: 'Deposit paid' },
      ] },
      { type: 'lines', table: 'Booking Services', columns: [{ field: 'Service.Name', label: 'Service' }, { field: 'Price', type: 'currency' }] },
      { type: 'totals', rows: [{ label: 'Price for the visit', expr: '[Price]' }, { label: 'Less deposit paid', expr: 'Nz([Deposit])' }, { label: 'To pay on the day', expr: '[Price] - Nz([Deposit])', strong: true }] },
      { type: 'text', text: 'Prices include VAT. If you would like a different service on the day, the price may change; your stylist will tell you before starting.' },
      { type: 'text', text: 'Colour, tints and lash lifts need a skin test in the six months before. If you have not had one with us in that time, please call in for a quick test at least 48 hours before your appointment. It takes five minutes.' },
      { type: 'text', text: 'Our cancellation policy. If you need to move or cancel, please tell us at least 48 hours before, by phone on {Business.Phone} or by e-mail to {Business.E-mail}, and any deposit is refunded or moved to your new booking. With less notice, or if you do not come, we keep the deposit and may ask for one when you next book. If you are running late, please call: we will do what we can in the time left, or offer you another appointment.' },
      { type: 'text', text: 'We look forward to seeing you.' },
      { type: 'text', text: 'With best wishes, the team at {Business.Name}' },
    ] },
  ],
  reports: [
    { name: 'The column', title: 'The column: the next 7 days by stylist', query: 'Bookings next 7 days', group: 'Stylist', sort: 'When', columns: ['When', 'Client', 'Phone', { field: 'Minutes', type: 'integer' }, { field: 'Deposit', type: 'currency' }], totals: [{ field: 'Minutes', fn: 'sum' }] },
    { name: 'Commission statement', title: 'Commission statement: last month', query: 'Service commission last month', group: 'Stylist', columns: [{ field: 'Day', type: 'date' }, 'Client', 'Service', { field: 'Price', type: 'currency' }, { field: 'Commission', type: 'currency' }], totals: [{ field: 'Price', fn: 'sum' }, { field: 'Commission', fn: 'sum' }] },
    { name: 'Retail commission', title: 'Retail commission: last month', query: 'Retail commission last month', columns: ['Stylist', { field: 'Items Sold', type: 'integer' }, { field: 'Retail Sales', type: 'currency' }, { field: 'Rate', type: 'number' }, { field: 'Commission', type: 'currency' }], totals: [{ field: 'Retail Sales', fn: 'sum' }, { field: 'Commission', fn: 'sum' }] },
    { name: 'Takings by month', query: 'Service takings by month', columns: ['Month', { field: 'Services Done', type: 'integer' }, { field: 'Takings', type: 'currency' }], totals: [{ field: 'Services Done', fn: 'sum' }, { field: 'Takings', fn: 'sum' }] },
    { name: 'Products to order', query: 'Products to order', group: 'Supplier', columns: ['Product', 'Kind', 'Size', { field: 'In Stock', type: 'integer' }, { field: 'Reorder Level', type: 'integer' }, { field: 'Reorder Quantity', type: 'integer' }, { field: 'Order Cost', type: 'currency' }], totals: [{ field: 'Order Cost', fn: 'sum' }] },
    { name: 'Clients to win back', query: 'Clients to win back', columns: ['Client', 'Phone', 'E-mail', { field: 'Last Visit', type: 'date' }, { field: 'Visits', type: 'integer' }, { field: 'Days Since', type: 'integer' }], totals: [{ field: 'Client', fn: 'count' }] },
    { name: 'Price list', query: 'Service menu', group: 'Category', columns: ['Service', { field: 'Minutes', type: 'integer' }, { field: 'Junior Price', type: 'currency' }, { field: 'Stylist Price', type: 'currency' }, { field: 'Senior Price', type: 'currency' }] },
    { name: 'VAT by month', title: 'VAT on takings by month', query: 'VAT by month', columns: ['Month', { field: 'Receipts', type: 'integer' }, { field: 'Takings', type: 'currency' }, { field: 'Net', type: 'currency' }, { field: 'VAT', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }, { field: 'Net', fn: 'sum' }, { field: 'VAT', fn: 'sum' }] },
    { name: 'Till roll', title: 'Receipts: the last 7 days', query: 'Receipts last 7 days', columns: ['Number', 'Taken At', 'Client', 'Stylist', { field: 'Total', type: 'currency' }, { field: 'Paid Today', type: 'currency' }, 'Paid By'], totals: [{ field: 'Total', fn: 'sum' }, { field: 'Paid Today', fn: 'sum' }] },
  ],
};

export const about = {
  name: 'Marigold Hair and Beauty',
  kind: 'Salon',
  line: 'Clients, the appointment book, services, commission, retail stock and gift vouchers for a salon.',
  summary: [
    'For a salon, barber or beauty room that takes bookings, pays its team commission and sells a little retail over the desk. Marigold Hair and Beauty is an invented high-street salon with four stylists, a beauty therapist and a junior. It opens Tuesday to Saturday, with a late night until eight on Thursday.',
    'The records cover the last thirteen months of the appointment book and the next six weeks. Clients come back on their own rhythm, a gents cut every few weeks and a root tint every six or so, and some stop coming. One stylist left about five months ago and most of her clients moved to the stylist who replaced her. December is the busiest month and January the quietest. A few clients cancel late or do not turn up, and the salon keeps their deposit.',
    'Every service done is kept with the price charged, which depends on whether a junior, a stylist or a senior stylist did it, and new clients have 20 per cent off their first visit. Colour work and lash lifts need a patch test in the six months before. Retail and back-bar products are counted in and out as they are sold or used, and a weekly order goes in on Tuesday for whatever has run low. Two lines have been on back order with the supplier for the last few weeks. Gift vouchers are sold, mostly in December and before Mothering Sunday, and redeemed against bookings.',
    'The salon is registered for VAT, and its prices, like any salon\'s menu, have VAT in them. Each visit paid for at the desk has a numbered receipt that lists the services and any products bought at the same time, with the VAT within each line, the deposit already paid and any gift voucher used. It prints the gift voucher itself, a commission statement for each stylist each month, and a letter confirming an appointment.',
  ].join('\n\n'),
  suits: ['A barber shop', 'A nail bar', 'A beauty or brow and lash studio', 'A massage or spa room', 'A dog groomer', 'A mobile hairdresser'],
  tables: {
    Business: BUSINESS_ABOUT,
    'Suppliers': { about: 'The wholesalers the salon buys its products from.', fields: {
      'Name': 'The supplier\'s trading name.', 'Contact': 'The person the salon deals with there.', 'Phone': 'Their order line.',
      'E-mail': 'Where orders and queries go.', 'Account Number': 'The salon\'s account number with the supplier, quoted when ordering.',
      'Lead Days': 'How many days an order usually takes to arrive.',
    } },
    'Staff': { about: 'One record per member of the team, past and present, with the commission each earns.', fields: {
      'First Name': 'The name clients know them by.', 'Last Name': 'Their family name.', 'Role': 'Their job in the salon, such as Senior Stylist or Beauty Therapist.',
      'Level': 'Junior, Stylist, Senior or Therapist: which price column of the service menu their clients pay.',
      'Phone': 'Their mobile, for rota changes.', 'E-mail': 'Their salon e-mail address.', 'Works': 'The days they normally work.',
      'Started': 'The day they joined.', 'Left': 'The day they left, empty while they work here.',
      'Service Commission': 'The share of their service takings paid as commission, 0.12 for 12 per cent. The owner takes none.',
      'Retail Commission': 'The share of the retail they sell paid as commission.', 'Active': 'Yes while they work here; No keeps their history after they leave.',
    } },
    'Services': { about: 'The service menu: every treatment the salon offers, how long it takes and what it costs at each level.', fields: {
      'Name': 'The service as it appears on the menu.', 'Category': 'The heading it sits under: Cut and Finish, Colour, Treatments, Nails, Brows and Lashes, Skin and Body.',
      'Minutes': 'The time to book for it, processing included.', 'Junior Price': 'The price with the junior. Empty where the junior does not do it.',
      'Stylist Price': 'The price with a stylist, and the price for beauty services.', 'Senior Price': 'The price with a senior stylist or the owner.',
      'Patch Test': 'Yes where a skin test is needed in the six months before, such as colour and lash lifts.',
      'Tax Rate': 'The VAT within the price, as a fraction: 0.2 is twenty per cent. Copied onto the receipt line when the visit is paid for.',
      'Active': 'No takes a service off the menu without losing its history.',
    } },
    'Clients': { about: 'One record per client, with what the salon needs to know before they sit down.', fields: {
      'First Name': 'Their first name.', 'Last Name': 'Their family name.', 'Phone': 'Their mobile, for reminders and running-late calls.',
      'E-mail': 'Their e-mail address, for booking confirmations.', 'Heard Of Us': 'How they found the salon, to see which kind of advertising brings people in.',
      'Client Since': 'The day they first booked.', 'Patch Test Date': 'The day of their last skin test for colour or tint.',
      'Allergies': 'Anything to avoid, such as latex gloves or a fragrance.', 'Colour Notes': 'The colour they have and how it is mixed, so anyone can repeat it.',
      'Marketing Consent': 'Yes if they agreed to hear about offers. Only these clients go on the win-back list.', 'Notes': 'Anything else worth knowing.',
    } },
    'Bookings': { about: 'The appointment book: one record per appointment, past and to come.', fields: {
      'Client': 'Who the appointment is for.', 'Staff': 'Who it is booked with.', 'Start': 'The day and time it starts.',
      'Minutes': 'How long it is booked for, the services added together.', 'Status': 'Booked, Completed, Cancelled or No-show.',
      'Booked On': 'The day the booking was made.', 'Booked Via': 'Phone, Online or In salon (rebooked at the desk).',
      'Deposit': 'Taken when booking new clients and long appointments.', 'Deposit Refunded': 'Yes when a cancellation came with two days\' notice or more and the deposit went back.',
      'Cancelled On': 'The day it was cancelled, for a cancelled booking.', 'Paid By': 'Card, Cash or Gift voucher, once the appointment is done.',
      'Notes': 'Anything to know for this visit.',
    } },
    'Booking Services': { about: 'The services in each booking, with the price charged for each.', fields: {
      'Booking': 'The appointment it belongs to.', 'Service': 'The service from the menu.',
      'Price': 'What was charged: the price for the stylist\'s level, less the new-client offer on a first visit.',
    } },
    'Products': { about: 'Everything the salon stocks: retail products sold to clients and back-bar products used in services.', fields: {
      'Name': 'The product, described plainly.', 'Kind': 'Retail (sold) or Back bar (used at the basin and in treatments).',
      'Category': 'Shampoo, Colour, Nails and so on.', 'Supplier': 'Who it is ordered from.', 'Size': 'The size of one unit.',
      'Cost Price': 'What one unit costs the salon.', 'Sale Price': 'The retail price, VAT included; empty for back-bar products.',
      'Tax Rate': 'The VAT within the sale price, as a fraction: 0.2 is twenty per cent.',
      'In Stock': 'Units on the shelf now.', 'Reorder Level': 'At or below this many, it goes on the next order.',
      'Reorder Quantity': 'How many to order at a time.', 'Active': 'No for a product no longer stocked.',
    } },
    'Product Sales': { about: 'Retail sold at the desk, one record per product per sale.', fields: {
      'Date': 'The day of the sale.', 'Staff': 'Who sold it, for retail commission.', 'Client': 'Who bought it.',
      'Product': 'What was sold.', 'Quantity': 'How many.', 'Unit Price': 'The price of one, as charged.',
      'Receipt': 'The receipt it was rung up on, with the services of the same visit.',
    } },
    'Stock Deliveries': { about: 'What came in from suppliers, one record per product per delivery.', fields: {
      'Supplier': 'Who delivered it.', 'Product': 'What came in.', 'Received': 'The day it arrived.',
      'Quantity': 'How many units.', 'Unit Cost': 'What each cost on the invoice.', 'Invoice Number': 'The supplier\'s invoice, to match against the bill.',
    } },
    'Gift Vouchers': { about: 'Vouchers sold, and when and where each was used.', fields: {
      'Code': 'The number printed on the voucher.', 'Sold On': 'The day it was bought.', 'Bought By': 'Who paid for it.',
      'For': 'Who it is for.', 'Value': 'What it is worth.', 'Expires': 'A year after it was sold.',
      'Redeemed On': 'The day it was used, empty until then.', 'Booking': 'The appointment it paid for.',
    } },
    'Receipts': { about: 'The till\'s record of each visit paid for at the desk, one per completed appointment. It keeps what was charged as it was printed, so a receipt printed again later says what the client was given on the day.', fields: {
      'Number': 'The receipt number, given in the order receipts are taken, a new run each year: RCP-2026-00042.',
      'Booking': 'The appointment paid for. The client and stylist printed on the receipt come from it.', 'Taken': 'The day and time the payment was taken.',
      'Net': 'The total less the VAT within it.', 'VAT': 'The VAT within the total: each line\'s VAT rounded to the penny, then added up.',
      'Total': 'Everything on the receipt, VAT included: the services and any products.',
      'Deposit Paid': 'The deposit taken when the appointment was booked, which comes off the bill.',
      'Voucher': 'The value of a gift voucher used to pay, nought when none was.', 'Voucher Code': 'The number of the gift voucher used.',
      'Paid Today': 'What the client paid at the desk: the total less the deposit and any voucher.',
      'Paid By': 'How it was paid: Card, Cash, Gift voucher, or a voucher and card or cash for the rest.',
    } },
    'Receipt Lines': { about: 'What is on each receipt: a line for each service done and each product bought at the visit.', fields: {
      'Receipt': 'The receipt the line is on.', 'Kind': 'Service or Retail.',
      'Item': 'The service or product as printed, with the size of a product and a note of the first-visit offer.',
      'Quantity': 'How many: one for a service.', 'Unit Price': 'The price of one, VAT included, as charged.',
      'Tax Rate': 'The VAT within the price, as a fraction, from the service or product.',
    } },
    'Commission Statements': { about: 'A statement for each person on commission for each month, made on the first of the next month and paid with the wages on the 7th.', fields: {
      'Number': 'The statement number, a new run each year: COM-2026-0042.', 'Staff': 'Who it is for.', 'Month': 'The month it covers, as 2026-09.',
      'Prepared On': 'The day it was made.', 'Commission Due': 'The commission on all its lines added up.',
      'Paid On': 'The day it was paid with the wages; empty until then.',
    } },
    'Commission Lines': { about: 'The lines of each commission statement: the services done by kind, and the retail sold.', fields: {
      'Statement': 'The statement the line is on.', 'Line': 'Its place on the statement.', 'Work': 'Which kind of work: a heading of the service menu, such as Colour, or Retail products sold.',
      'Done': 'How many services were done, or products sold.', 'Takings': 'What clients paid for them, VAT included.',
      'Rate': 'The commission rate for that kind of work, as a fraction.', 'Commission': 'The commission earned: each service\'s or sale\'s rounded to the penny, then added up.',
    } },
  },
  queries: {
    'Bookings next 7 days': 'Every booked appointment from today to a week ahead, by time, with the client\'s phone number. Use it to send reminders and plan the week.',
    'Takings by stylist last month': 'What each member of the team took in services last calendar month, and how many services they did.',
    'Service commission last month': 'Each service done last month by someone on commission, with the price and the commission it earns. The commission statement prints it.',
    'Retail commission last month': 'Retail each person sold last month and the commission due on it.',
    'Service takings by month': 'Service takings for each month, completed appointments only. Shows the busy and quiet months.',
    'Retail sales by month': 'Retail sold each month and the profit on it after cost.',
    'Clients to win back': 'Clients who agreed to hear from the salon, last seen between three months and a year ago, with nothing booked. The ones to send an offer to.',
    'Most booked services': 'Which services were done most over the last year, and what each took.',
    'Products to order': 'Retail and back-bar products at or below their reorder level, by supplier, with what the order will cost.',
    'No-shows by client': 'Clients who have missed an appointment, with their late cancellations and the deposits kept. Use it before taking another booking without a deposit.',
    'Patch tests due': 'Colour and lash appointments in the next two weeks where the client has had no patch test in the six months before. Call them in for one at least two days ahead.',
    'Vouchers outstanding': 'Gift vouchers sold and not yet used, soonest to expire first. What the salon still owes in services.',
    'New clients by source': 'New clients over the last year by how they heard of the salon.',
    'Retail best sellers': 'Retail products over the last year, by the profit they made.',
    'Service menu': 'The services on offer, by category, with the price at each level.',
    'VAT by month': 'For each month, how many receipts were given, the takings, and the net and VAT within them: the figures for the VAT return, three months at a time.',
    'Receipts last 7 days': 'Every receipt of the last week, with the client, the stylist, the total and what was paid at the desk and how. Use it to find a receipt to print again.',
  },
  reports: {
    'The column': 'The appointment book for the next seven days, a section for each stylist, with times and phone numbers. Print it for the desk.',
    'Commission statement': 'Last month\'s services for each person on commission, with a total of takings and commission each. For the owner and for payroll.',
    'Retail commission': 'Retail commission due for last month, one line a person.',
    'Takings by month': 'Service takings month by month with a grand total.',
    'Products to order': 'The order list for the week, a section for each supplier, with the cost of each line.',
    'Clients to win back': 'A call or mailing list of lapsed clients who agreed to marketing.',
    'Price list': 'The service menu, printed by category, with the junior, stylist and senior prices.',
    'VAT by month': 'Takings, net and VAT month by month, with the totals, for the bookkeeper.',
    'Till roll': 'The last seven days of receipts, with the total taken and the total paid at the desk, to check against the card machine and the cash.',
  },
  documents: {
    'Receipt': 'The till receipt for a visit, on an 80 mm roll: the salon\'s name, address and VAT number, the client and stylist, each service and product with the VAT within it, the total, the deposit and any voucher taken off, and what was paid today. Numbered RCP-2026-00042.',
    'Gift voucher': 'The voucher a buyer takes away, on A5 card: who it is for and from, what it is worth, its number and use-by date, and the terms.',
    'Commission statement': 'Given to each stylist and therapist on commission with their pay: the month\'s services by kind and the retail they sold, the takings, the rate and the commission due. Numbered COM-2026-0042.',
    'Appointment confirmation': 'A letter or e-mail confirming an appointment: the day and time, who with, the services and their prices, the deposit, and the skin test and cancellation policy. Every booking has one; print or send the one on screen, or choose the week\'s new bookings. Have the cancellation words checked against the consumer rules of your own country.',
  },
  tutorial: [
    { title: 'Book an appointment and confirm it', steps: [
      'Open the Clients form. Look the client up; for someone new, press New and fill in their name, mobile, e-mail and Heard Of Us, and tick Marketing Consent only if they agree.',
      'Open the Bookings form and press New. Choose the client in Client and the stylist in Staff, type the day and time in Start, and set Booked On to today and Booked Via to Phone, Online or In salon.',
      'In Booking Services at the foot of the form, add each service. Type the Price from the stylist\'s column of the menu (the Price list report has it), taking 20 per cent off for a first visit. Add the minutes of the services together in Minutes.',
      'If you take a deposit, type it in Deposit. Then open the print template Appointment confirmation from Print templates, choose the booking, and press Print or Save as PDF to send it.',
    ] },
    { title: 'Check tomorrow\'s column and the skin tests', steps: [
      'Run the query Bookings next 7 days, or print the report The column, which puts each stylist\'s week on its own page.',
      'Run Patch tests due. Call each client on it to come in for a skin test at least 48 hours before their colour or lash appointment, and when they have had it, type the day in Patch Test Date on the Clients form.',
    ] },
    { title: 'Take payment at the end of a visit', steps: [
      'On the Bookings form, set Status to Completed and Paid By to Card, Cash or Gift voucher.',
      'Open the Receipts form and press New. Type the next number in Number (the Receipts last 7 days query shows the latest), choose the booking, and type the time in Taken.',
      'In Receipt Lines below, add a line for each service and each product sold, with Kind, Item, Quantity, Unit Price and Tax Rate 0.2. For each product, also add a record to Product Sales with the same receipt, so the stock and the retail commission are counted.',
      'Type the Total, the VAT (each line\'s price less the price divided by 1.2), Net, any Deposit Paid and Voucher, and Paid Today. Open the print template Receipt, choose the receipt and press Print.',
    ] },
    { title: 'Sell a gift voucher and take it as payment', steps: [
      'Open the Gift Vouchers form and press New. Type the next Code, today in Sold On, who is paying in Bought By, who it is for in For, the Value, and a date a year ahead in Expires.',
      'Open the print template Gift voucher, choose the voucher, and print it on card.',
      'When it is used, type the day in Redeemed On and choose the booking in Booking. On that visit\'s receipt, type the voucher\'s value in Voucher and its code in Voucher Code.',
      'Run Vouchers outstanding to see what the salon still owes in services, soonest to expire first.',
    ] },
    { title: 'Order stock on Tuesday', steps: [
      'Print the report Products to order. It lists every product at or below its reorder level, a page for each supplier, with the cost of the order.',
      'Send each supplier its page. When the goods arrive, add a record to Stock Deliveries for each product, with the supplier\'s invoice number, and add the quantity to In Stock on the Products form.',
    ] },
    { title: 'Pay the month\'s commission', steps: [
      'On the first of the month, print the reports Commission statement and Retail commission to see last month\'s services and retail by person.',
      'Open the Commission Statements form and press New for each person on commission. Type the Number, choose them in Staff, and type the month (2026-09) in Month and today in Prepared On.',
      'In Commission Lines below, add a line for each kind of work with what was done, the takings, the rate and the commission, and put the total in Commission Due.',
      'Open the print template Commission statement, choose the statement, and print it to go with their pay. Type the pay day in Paid On.',
    ] },
    { title: 'Do the VAT return', steps: [
      'Print the report VAT by month. Add up the three months of the VAT quarter: the VAT column is the output tax on takings, and the Net column the sales before VAT.',
      'Gift vouchers sold are not in it: VAT is due when a voucher is spent, and that is in the receipt it was spent on.',
    ] },
  ],
  notes: [
    'Prices are VAT inclusive. A receipt shows the VAT within each line, worked out as the price less the price divided by one plus the rate, rounded to the penny; the receipt\'s VAT is those added up. Change the 0.2 in Services and Products if your rate is different, or make it 0 if you are not registered for VAT.',
    'The services of a visit are kept twice on purpose: in Booking Services, as booked, and in Receipt Lines, as charged and printed. A receipt should not change after it is handed over, so the till keeps its own copy.',
    'Every booking, past or to come, can print an Appointment confirmation. Choose the booking you want before printing; printing them all gives thousands of letters.',
    'Commission is worked out on takings with VAT in them, as many salons pay it. To pay it on takings before VAT, divide the price by 1.2 in the commission queries.',
    'The owner takes no commission, so she has no statements. Staff who have left keep their records with Active set to No.',
    'Stock is not counted by Data Files on its own: In Stock goes down as products are sold or used at the basin and up as deliveries arrive. In the sample this was worked out while the records were made; in your own salon, change In Stock when you add a sale or a delivery, or count the shelf each week.',
  ],
  faq: [
    { q: 'How do I print a receipt again for a client?', a: 'Run Receipts last 7 days, or open the Receipts form and find the number. Then open the print template Receipt, choose that receipt and print it. It prints exactly as it did on the day.' },
    { q: 'Where does the VAT number on the receipt come from?', a: 'From the Business table, with the salon\'s name, address and phone. Change it there once and every receipt, voucher, statement and letter takes the new details.' },
    { q: 'A client cancelled with less than 48 hours\' notice. What do I do?', a: 'Set the booking\'s Status to Cancelled and type the day in Cancelled On. Leave Deposit Refunded unticked, so the deposit is kept. No-shows by client counts late cancellations and no-shows together.' },
    { q: 'How do I pay part of a bill with a gift voucher?', a: 'In the sample a voucher is used in full on one visit. Type its value in Voucher on the receipt and the rest in Paid Today, with Paid By as Gift voucher and card. To let a balance stay on a voucher, add a Balance field to Gift Vouchers.' },
    { q: 'Why does the commission statement for a stylist not match the takings report?', a: 'Takings by stylist last month counts services only; the statement adds the retail they sold, at the retail rate. Both use completed appointments only.' },
    { q: 'How do I start with my own salon instead of Marigold?', a: 'Open the sample empty and fill in Business first. Then add your suppliers, staff, service menu and products, in that order, before any client or booking.' },
  ],
  empty: 'Start with Business, your own name, address, VAT number and the words at the foot of a receipt. Then Suppliers, then Staff with their level and commission rates, then the Services menu with a price for each level. Add your Products next, with their supplier and reorder level. Clients can be added as they book. Then make Bookings and add the services for each in Booking Services, changing the price where you charge something different. Record Product Sales and Stock Deliveries as they happen, and Gift Vouchers when you sell them. Each visit paid for gets a record in Receipts with its Receipt Lines, and at the end of each month a record in Commission Statements, with its lines, for each person on commission.',
  adapt: [
    'A barber can drop the beauty services, keep Junior and Senior prices and add a Walk-in choice to Booked Via.',
    'A dog groomer can rename Clients to Owners and add a Pets table pointing at it (name, breed, size, temperament), then make each booking point at a pet.',
    'To pay commission only above a target, add a Target field to Staff and a calculated column in Takings by stylist last month that takes it off the takings.',
    'Add a Reminder Sent yes/no to Bookings and tick it when the reminder goes, so nobody sends two.',
    'For a nail bar or brow studio, add a Station or Room field to Bookings to show which chair or room each appointment uses.',
    'Change the patch test rule from six months to another period by editing 180 in the Patch tests due query.',
    'Add a Tips field to Bookings and a query that totals tips by stylist each month.',
    'To sell products over the counter without a booking, add a Client field to Receipts and leave Booking empty; the Receipt template can then print {Client} instead of {Booking.Client}.',
    'Give receipts for gift vouchers too: add a Receipts record with one line, Kind Voucher and Tax Rate 0, since VAT on a voucher is paid when it is spent.',
  ],
};

/** The service menu: [category, [name, minutes, junior, stylist, senior, patch test]...]. Junior null where the junior does not do it. */
const MENU = [
  ['Cut and Finish', [
    ['Cut and blow dry', 60, 34, 46, 56], ['Restyle cut and blow dry', 75, null, 58, 70], ['Blow dry', 40, 20, 27, 32],
    ['Gents cut', 30, 16, 22, 26], ['Beard trim', 15, 8, 10, 12], ["Child's cut (under 12)", 30, 12, 16, 18],
    ['Fringe trim', 15, 6, 6, 6], ['Hair up for an occasion', 60, null, 50, 60],
  ]],
  ['Colour', [
    ['Root tint', 90, 44, 54, 62, true], ['Full head tint', 120, null, 72, 82, true], ['Half head highlights', 120, null, 78, 90, true],
    ['Full head highlights', 150, null, 102, 118, true], ['Balayage', 180, null, 125, 145, true], ['Toner', 30, 16, 22, 26, true],
  ]],
  ['Treatments', [['Conditioning treatment', 15, 12, 12, 12], ['Smoothing treatment', 150, null, 170, 195]]],
  ['Nails', [['Gel manicure', 45, null, 28, 28], ['Gel pedicure', 60, null, 34, 34], ['Gel removal', 15, null, 8, 8]]],
  ['Brows and Lashes', [['Eyebrow wax', 15, null, 14, 14], ['Brow tint', 15, null, 10, 10, true], ['Lash tint', 20, null, 15, 15, true], ['Lash lift', 60, null, 45, 45, true]]],
  ['Skin and Body', [['Facial, 60 minutes', 60, null, 55, 55], ['Back massage, 30 minutes', 30, null, 32, 32], ['Half leg wax', 30, null, 22, 22]]],
];

/** What is stocked: [kind, category, supplier index, [name, size, cost, sale price or null, level, order quantity, weekly demand]...]. Generic goods; no brand. */
const STOCK = [
  ['Retail', 'Shampoo', 1, [
    ['Colour care shampoo', '250 ml', 7.2, 16, 4, 12, 1.4], ['Moisture shampoo', '250 ml', 6.8, 15, 4, 12, 1.1], ['Volume shampoo', '250 ml', 6.6, 15, 3, 10, 0.8],
    ['Purple shampoo for blondes', '250 ml', 7.4, 16.5, 4, 12, 1.0], ['Scalp care shampoo', '250 ml', 7.9, 17, 2, 6, 0.4],
  ]],
  ['Retail', 'Conditioner', 1, [['Colour care conditioner', '250 ml', 7.6, 16.5, 4, 12, 1.2], ['Moisture conditioner', '250 ml', 7.1, 15.5, 3, 10, 0.9], ['Repair hair mask', '200 ml', 9.4, 21, 3, 8, 0.7]]],
  ['Retail', 'Styling', 1, [
    ['Heat protection spray', '150 ml', 6.4, 15, 3, 10, 0.9], ['Texture spray', '200 ml', 7.0, 17, 3, 8, 0.6], ['Firm hold hairspray', '300 ml', 5.6, 13, 3, 10, 0.8],
    ['Dry shampoo', '200 ml', 4.8, 11.5, 3, 10, 0.9], ['Hair oil', '100 ml', 10.2, 24, 3, 8, 0.7], ['Matt paste for men', '75 ml', 6.1, 14, 3, 8, 0.6],
  ]],
  ['Retail', 'Tools', 3, [['Paddle brush', 'each', 6.5, 16, 2, 6, 0.3], ['Round brush, medium', 'each', 7.2, 18, 2, 4, 0.2], ['Wide-tooth comb', 'each', 1.6, 5, 3, 10, 0.3]]],
  ['Retail', 'Nails and Skin', 2, [['Cuticle oil', '15 ml', 2.9, 9, 3, 10, 0.4], ['Hand cream', '75 ml', 3.3, 10, 3, 10, 0.4], ['Face moisturiser', '50 ml', 9.8, 26, 2, 6, 0.3]]],
  ['Retail', 'Gifts', 1, [['Shampoo and conditioner gift set', 'box', 13.5, 32, 2, 6, 0.3]]],
  ['Back bar', 'Colour', 0, [
    ['Permanent colour', '60 ml tube', 5.9, null, 30, 60, 0], ['Developer, 6 per cent', '1 litre', 4.6, null, 4, 10, 0], ['Lightening powder', '500 g tub', 14.5, null, 2, 4, 0],
    ['Toner', '60 ml tube', 6.3, null, 10, 24, 0], ['Foil roll', '100 m', 10.8, null, 2, 6, 0], ['Colour gloves', 'box of 100', 5.9, null, 2, 6, 0],
  ]],
  ['Back bar', 'Basin', 1, [['Backwash shampoo', '5 litre', 17.5, null, 1, 3, 0], ['Backwash conditioner', '5 litre', 18.9, null, 1, 3, 0], ['Treatment mask, salon size', '1 litre', 16.2, null, 1, 3, 0]]],
  ['Back bar', 'Beauty', 2, [
    ['Gel polish colour', '15 ml', 7.4, null, 6, 12, 0], ['Gel base and top coat', '15 ml pair', 9.2, null, 2, 6, 0], ['Gel remover', '1 litre', 6.8, null, 1, 3, 0],
    ['Hot wax', '800 g tin', 8.9, null, 2, 4, 0], ['Brow and lash tint', '15 ml', 4.9, null, 2, 6, 0], ['Lash lift kit', 'kit of 10', 22, null, 1, 2, 0],
    ['Facial cleanser, salon size', '500 ml', 12.4, null, 1, 2, 0], ['Massage oil', '1 litre', 11.6, null, 1, 2, 0],
  ]],
];

/** Back-bar used by each service, in units of the product: [product, units per service]. Basin products go with every hair service. */
const USES = {
  'Root tint': [['Permanent colour', 1.2], ['Developer, 6 per cent', 0.1], ['Colour gloves', 0.02]],
  'Full head tint': [['Permanent colour', 2], ['Developer, 6 per cent', 0.18], ['Colour gloves', 0.02]],
  'Half head highlights': [['Lightening powder', 0.12], ['Developer, 6 per cent', 0.12], ['Foil roll', 0.08]],
  'Full head highlights': [['Lightening powder', 0.22], ['Developer, 6 per cent', 0.2], ['Foil roll', 0.15]],
  'Balayage': [['Lightening powder', 0.25], ['Developer, 6 per cent', 0.2]],
  'Toner': [['Toner', 1], ['Developer, 6 per cent', 0.05]],
  'Conditioning treatment': [['Treatment mask, salon size', 0.04]],
  'Gel manicure': [['Gel polish colour', 0.08], ['Gel base and top coat', 0.07]],
  'Gel pedicure': [['Gel polish colour', 0.08], ['Gel base and top coat', 0.07]],
  'Gel removal': [['Gel remover', 0.04]],
  'Eyebrow wax': [['Hot wax', 0.04]], 'Half leg wax': [['Hot wax', 0.12]],
  'Brow tint': [['Brow and lash tint', 0.06]], 'Lash tint': [['Brow and lash tint', 0.06]], 'Lash lift': [['Lash lift kit', 0.1]],
  'Facial, 60 minutes': [['Facial cleanser, salon size', 0.06]], 'Back massage, 30 minutes': [['Massage oil', 0.04]],
};
const BASIN = [['Backwash shampoo', 0.012], ['Backwash conditioner', 0.01]];

/** The kinds of client: what they come in for, how often (days), and who books them. */
const PROFILES = {
  gent: { cadence: [21, 42], hair: true }, child: { cadence: [42, 84], hair: true }, cut: { cadence: [42, 91], hair: true },
  blowdry: { cadence: [7, 14], hair: true }, colour: { cadence: [35, 63], hair: true, patch: true },
  nails: { cadence: [21, 28] }, brows: { cadence: [28, 42], patch: true }, lashes: { cadence: [42, 63], patch: true },
  facial: { cadence: [28, 63] }, wax: { cadence: [28, 49] }, massage: { cadence: [21, 56] },
};
const MIX = {
  Senior: [['colour', 45], ['cut', 33], ['blowdry', 6], ['gent', 12], ['child', 4]],
  Stylist: [['colour', 36], ['cut', 33], ['gent', 18], ['child', 8], ['blowdry', 5]],
  Junior: [['gent', 40], ['child', 18], ['cut', 22], ['blowdry', 20]],
  Therapist: [['nails', 34], ['brows', 30], ['lashes', 14], ['facial', 11], ['wax', 8], ['massage', 3]],
};

/** Opening hours by weekday (0 Sunday), in minutes from midnight: Thursday is the late night. */
const OPEN = { 2: [540, 1050], 3: [540, 1050], 4: [600, 1200], 5: [540, 1080], 6: [510, 990] };
const DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

// Names by how the salon would book them: a gents cut goes to a man's name.
const MEN = new Set(['Ben', 'Emeka', 'Giorgio', 'Ivan', 'Kofi', 'Mateo', 'Oskar', 'Rhys', 'Tariq', 'Viktor', 'Yusuf', 'Bilal', 'Dmitri', 'Farhan', 'Hamza', 'Javier', 'Lukas', 'Noah', 'Paulo', 'Samir', 'Tomasz', 'Wale', 'Zain', 'Callum', 'Eamon', 'Hiroshi']);
const EITHER = new Set(['Chen', 'Wen', 'Olu']);
const MALE = FIRST.filter((n) => MEN.has(n));
const FEMALE = FIRST.filter((n) => !MEN.has(n) && !EITHER.has(n));

const pad = (n, w) => String(n).padStart(w, '0');
const hhmm = (m) => [Math.floor(m / 60), m % 60];

/** Fill a Marigold Hair and Beauty database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 395, ahead = 42 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };

  // Phone numbers from the fiction ranges, each given out once: mobiles first, then landlines.
  const mobiles = r.some([...Array(1000).keys()], 1000).map((i) => `07700 900${pad(i, 3)}`);
  const landlines = r.some([...Array(1000).keys()], 1000).map((i) => `020 7946 0${pad(i, 3)}`);
  const phone = (land = false) => (land ? landlines.pop() : mobiles.pop() || landlines.pop()) || '';
  const mails = new Set();
  const mailFor = (first, last, domain) => {
    let m = `${first}.${last}`.toLowerCase().replace(/[^a-z.]/g, '');
    for (let i = 2; mails.has(`${m}@${domain}`); i++) m = `${first}.${last}${i}`.toLowerCase().replace(/[^a-z.0-9]/g, '');
    mails.add(`${m}@${domain}`);
    return `${m}@${domain}`;
  };
  const nameOf = (male) => ({ first: r.pick(male === undefined ? FIRST : male ? MALE : FEMALE), last: r.pick(LAST) });

  // The salon's own details, printed at the head of its receipts, vouchers, statements and letters.
  add('Business', businessRecord({
    name: 'Marigold Hair and Beauty', domain: 'marigoldhair.example', phone: '020 7946 0418',
    address: { Address: '27 Market Square', City: 'London', Postcode: 'N16 7QR', Country: 'United Kingdom' },
    terms: 'Payment at the end of each visit, by card or cash.',
    footer: 'Open Tuesday to Saturday, until 8pm on Thursday. Please give us 48 hours notice to move or cancel an appointment. Unopened products can be returned within 14 days with this receipt. Thank you for coming in.',
  }));

  // Suppliers: colour, hair care, beauty, sundries.
  const suppliers = companies(r, 4, ['Salon Supplies', 'Beauty Wholesale', 'Hair Trade', 'Professional']).map((name, i) => {
    const n = nameOf();
    const lead = [2, 3, 3, 4][i];
    const id = add('Suppliers', {
      Name: name, Contact: `${n.first} ${n.last}`, Phone: phone(true), 'E-mail': `orders@${name.toLowerCase().replace(/[^a-z]+/g, '')}.example`,
      'Account Number': `MAR-${pad(r.int(100, 9999), 4)}`, 'Lead Days': lead,
    });
    return { id, lead };
  });

  // The team. One stylist left five months ago; another took her chair.
  const team = [
    { role: 'Owner and Senior Stylist', level: 'Senior', days: [2, 3, 4, 6], svc: 0, retail: 0, busy: 0.3, since: -3200 },
    { role: 'Senior Stylist', level: 'Senior', days: [2, 4, 5, 6], svc: 0.15, retail: 0.1, busy: 0.34, since: -1900 },
    { role: 'Stylist', level: 'Stylist', days: [3, 4, 5, 6], svc: 0.12, retail: 0.1, busy: 0.34, since: -1100 },
    { role: 'Stylist', level: 'Stylist', days: [2, 3, 5, 6], svc: 0.12, retail: 0.1, busy: 0.34, since: -950, to: -150 },
    { role: 'Stylist', level: 'Stylist', days: [2, 3, 5, 6], svc: 0.12, retail: 0.1, busy: 0.3, since: -140 },
    { role: 'Beauty Therapist', level: 'Therapist', days: [2, 3, 4, 5, 6], svc: 0.12, retail: 0.1, busy: 0.3, since: -1500 },
    { role: 'Junior Stylist', level: 'Junior', days: [2, 3, 4, 5, 6], svc: 0.08, retail: 0.05, busy: 0.24, since: -400 },
  ];
  const firstNames = new Set();
  for (const s of team) {
    let n;
    do n = nameOf(); while (firstNames.has(n.first));
    firstNames.add(n.first);
    s.name = n.first;
    s.from = Math.max(s.since, -days - 1);
    s.to = s.to ?? Infinity;
    s.clients = [];
    s.hair = s.level !== 'Therapist';
    // Two or three weeks' holiday a year, never in December.
    s.off = new Set();
    for (let k = 0; k < 4; k++) {
      const start = r.int(-days, ahead);
      if (cal.at(start).getMonth() === 11) continue;
      for (let i = 0; i < 7; i++) s.off.add(start + i);
    }
    s.id = add('Staff', {
      'First Name': n.first, 'Last Name': n.last, Role: s.role, Level: s.level, Phone: phone(), 'E-mail': `${n.first.toLowerCase()}@marigoldhair.example`,
      Works: s.days.map((d) => DAY_NAMES[d]).join(', '), Started: cal.day(s.since), Left: Number.isFinite(s.to) ? cal.day(s.to) : null,
      'Service Commission': s.svc, 'Retail Commission': s.retail, Active: !Number.isFinite(s.to),
    });
  }
  const leaver = team[3];
  const newcomer = team[4];

  // The menu.
  const services = new Map();
  for (const [cat, items] of MENU) {
    for (const [name, minutes, junior, stylist, senior, patch = false] of items) {
      const id = add('Services', { Name: name, Category: cat, Minutes: minutes, 'Junior Price': junior, 'Stylist Price': stylist, 'Senior Price': senior, 'Patch Test': patch, 'Tax Rate': 0.2 });
      services.set(name, { id, name, category: cat, minutes, prices: { Junior: junior, Stylist: stylist, Senior: senior, Therapist: stylist }, patch, hair: !['Nails', 'Brows and Lashes', 'Skin and Body'].includes(cat) });
    }
  }
  const S = (name) => services.get(name);

  // The stock.
  const products = [];
  const byName = new Map();
  for (const [kind, cat, si, items] of STOCK) {
    for (const [name, size, cost, price, level, qty, demand] of items) {
      const p = { name, size, kind, cost, price, level, qty, demand, supplier: suppliers[si], stock: 0, use: 0, pending: false };
      p.id = add('Products', { Name: name, Kind: kind, Category: cat, Supplier: p.supplier.id, Size: size, 'Cost Price': cost, 'Sale Price': price, 'Tax Rate': 0.2, 'In Stock': 0, 'Reorder Level': level, 'Reorder Quantity': qty });
      products.push(p);
      byName.set(name, p);
    }
  }
  const retail = products.filter((p) => p.kind === 'Retail');
  // Two lines the supplier has had on back order for the last few weeks, so they run down.
  const backOrdered = new Set([byName.get('Hair oil'), byName.get('Lightening powder')]);
  const giftSet = byName.get('Shampoo and conditioner gift set');

  let invoice = r.int(20000, 30000);
  const deliver = (lines, d) => {
    const no = `INV-${invoice += r.int(3, 40)}`;
    for (const { p, n } of lines) {
      add('Stock Deliveries', { Supplier: p.supplier.id, Product: p.id, Received: cal.day(d), Quantity: n, 'Unit Cost': p.cost, 'Invoice Number': no });
      p.stock += n;
      p.pending = false;
    }
  };
  // Opening stock, in before the first day of the records.
  for (const s of suppliers) deliver(products.filter((p) => p.supplier === s).map((p) => ({ p, n: p.level + p.qty })), -days - 3);

  // Clients.
  const sources = [['Recommended by a friend', 4], ['Social media', 3], ['Search online', 3], ['Walked past', 2], ['Gift voucher', 1], ['Local magazine', 0.6]];
  const NOTES = ['Prefers a quiet appointment', 'Sensitive scalp, go gently at the basin', 'Likes a strong coffee', 'Brings her own photos to show', 'Needs the chair by the window (bad back)', 'Usually runs ten minutes late'];
  const ALLERGIES = ['Latex', 'Nickel', 'Strong fragrance', 'Nut oils'];
  const allClients = [];
  const makeClient = (profile, s, since, initial) => {
    const prof = PROFILES[profile];
    const male = profile === 'gent' ? true : ['colour', 'blowdry', 'brows', 'lashes', 'nails'].includes(profile) ? r.chance(0.04) : r.chance(profile === 'child' ? 0.5 : 0.25);
    const n = nameOf(male);
    const c = {
      profile, staff: s, cadence: prof.cadence, next: 0, lapsed: false, visits: 0, fresh: !initial,
      unreliable: r.chance(0.05), patch: null,
    };
    if (prof.patch && initial && r.chance(0.9)) c.patch = -days - r.int(5, 175);
    const values = {
      'First Name': n.first, 'Last Name': n.last, Phone: phone(), 'E-mail': r.chance(0.85) ? mailFor(n.first, n.last, r.pick(['example.com', 'example.org', 'example.net'])) : null,
      'Heard Of Us': r.weighted(sources), 'Client Since': cal.day(since), 'Patch Test Date': c.patch === null ? null : cal.day(c.patch),
      Allergies: r.chance(0.04) ? r.pick(ALLERGIES) : null,
      'Colour Notes': profile === 'colour' ? `${r.int(4, 9)}/${r.pick(['0', '1', '3', '4', '43', '71', '66'])} with ${r.pick([20, 20, 30])} vol, ${r.pick([30, 35, 40])} minutes` : null,
      'Marketing Consent': r.chance(0.62), Notes: r.chance(0.05) ? r.pick(NOTES) : null,
    };
    c.id = add('Clients', values);
    s.clients.push(c);
    allClients.push(c);
    return c;
  };
  // The book as it stood when the records begin: each member of the team with their regulars.
  const regulars = { Senior: 175, Stylist: 165, Junior: 110, Therapist: 190 };
  for (const s of team) {
    if (s === newcomer) continue;
    for (let i = 0; i < regulars[s.level]; i++) {
      const c = makeClient(r.weighted(MIX[s.level]), s, -days - r.int(20, 1800), true);
      c.next = -days + r.int(0, c.cadence[1]);
    }
  }

  /** What a client has done this time, from their profile and who they see. */
  const servicesFor = (c, s, month) => {
    const level = s.level;
    const list = [];
    switch (c.profile) {
      case 'gent': list.push('Gents cut'); if (r.chance(0.2)) list.push('Beard trim'); break;
      case 'child': list.push("Child's cut (under 12)"); break;
      case 'blowdry': list.push('Blow dry'); break;
      case 'cut':
        if (level !== 'Junior' && r.chance(month === 11 ? 0.12 : 0.03)) list.push('Hair up for an occasion');
        else list.push(level !== 'Junior' && r.chance(0.07) ? 'Restyle cut and blow dry' : 'Cut and blow dry');
        if (r.chance(0.1)) list.push('Conditioning treatment');
        if (level !== 'Junior' && r.chance(0.008)) list.push('Smoothing treatment');
        break;
      case 'colour': {
        const colour = level === 'Junior' ? 'Root tint' : r.weighted([['Root tint', 50], ['Half head highlights', 20], ['Full head highlights', 13], ['Balayage', 9], ['Full head tint', 8]]);
        list.push(colour);
        if (/highlights|Balayage/.test(colour) && r.chance(0.5)) list.push('Toner');
        list.push(r.chance(0.6) ? 'Cut and blow dry' : 'Blow dry');
        break;
      }
      case 'nails': list.push(r.chance(0.8) ? 'Gel manicure' : 'Gel pedicure'); if (r.chance(0.7)) list.push('Gel removal'); break;
      case 'brows': list.push('Eyebrow wax'); if (r.chance(0.6)) list.push('Brow tint'); break;
      case 'lashes': list.push('Lash lift'); if (r.chance(0.5)) list.push('Lash tint'); break;
      case 'facial': list.push('Facial, 60 minutes'); break;
      case 'wax': list.push('Half leg wax'); if (r.chance(0.3)) list.push('Eyebrow wax'); break;
      default: list.push('Back massage, 30 minutes');
    }
    return list.map(S);
  };

  const vouchers = [];
  const voucherQueue = []; // vouchers whose owner means to use them: { v, at }
  let voucherNo = 1000 + r.int(0, 400);
  const BOOKING_NOTES = ['Running ten minutes late, called ahead', 'Wedding on Saturday', 'Wants to go a little shorter this time', 'Bringing a photo of the colour she wants', 'First time with us, allow a consultation'];

  // Receipts numbered by the year they are taken in, RCP-2026-00042, and what each person
  // on commission earns in each month by kind of work, for the month's statement.
  const receiptCount = {};
  const earned = new Map();
  const earn = (s, month, work, done, takings, rate) => {
    const key = `${s.id}|${month}`;
    if (!earned.has(key)) earned.set(key, { s, month, work: new Map() });
    const w = earned.get(key).work;
    if (!w.has(work)) w.set(work, { done: 0, takings: 0, commission: 0, rate });
    const x = w.get(work);
    x.done += done;
    x.takings += takings;
    // Each service's or sale's commission rounded to the penny, then added up, as the commission queries do.
    x.commission += money(takings * rate);
  };
  const receipt = (v, d, date) => {
    const y = date.slice(0, 4);
    receiptCount[y] = (receiptCount[y] || 0) + 1;
    const lines = [...v.lines, ...v.retail.map(({ p, q }) => ({ Kind: 'Retail', Item: `${p.name}, ${p.size}`, Quantity: q, 'Unit Price': p.price, 'Tax Rate': 0.2 }))];
    // The VAT within each line rounded to the penny, then added up, as the receipt prints it.
    let total = 0;
    let vat = 0;
    for (const l of lines) {
      const amount = money(l.Quantity * l['Unit Price']);
      total += amount;
      vat += money(amount - amount / (1 + l['Tax Rate']));
    }
    total = money(total);
    vat = money(vat);
    const voucher = v.voucher ? v.voucher.Value : 0;
    const paidToday = money(total - v.deposit - voucher);
    const paidBy = v.voucher ? (paidToday > 0 ? `Gift voucher and ${r.weighted([['card', 8], ['cash', 1.4]])}` : 'Gift voucher') : v.paidBy;
    const [hh, mm] = hhmm(v.at);
    const id = add('Receipts', {
      Number: `RCP-${y}-${pad(receiptCount[y], 5)}`, Booking: v.booking, Taken: cal.time(d, hh, mm), Net: money(total - vat), VAT: vat, Total: total,
      'Deposit Paid': v.deposit, Voucher: voucher, 'Voucher Code': v.voucher ? v.voucher.Code : null, 'Paid Today': paidToday, 'Paid By': paidBy,
    });
    for (const { category, ...l } of lines) add('Receipt Lines', { Receipt: id, ...l });
    for (const { p, q } of v.retail) add('Product Sales', { Date: date, Staff: v.s.id, Client: v.c.id, Product: p.id, Quantity: q, 'Unit Price': p.price, Receipt: id });
    const month = date.slice(0, 7);
    if (v.s.svc > 0) for (const l of v.lines) earn(v.s, month, l.category, 1, l['Unit Price'], v.s.svc);
    if (v.s.retail > 0) for (const { p, q } of v.retail) earn(v.s, month, 'Retail', q, money(q * p.price), v.s.retail);
  };

  for (let d = -days; d <= ahead; d++) {
    const date = cal.day(d);
    const at = cal.at(d);
    const wd = at.getDay();
    const month = at.getMonth();
    const hours = OPEN[wd];
    const closed = !hours || (month === 11 && (at.getDate() === 25 || at.getDate() === 26)) || (month === 0 && at.getDate() === 1);

    // The day the stylist leaves: most of her clients move to the new stylist, some follow her elsewhere.
    if (d === leaver.to) {
      for (const c of leaver.clients) {
        if (r.chance(0.25)) { c.lapsed = true; continue; }
        const to = r.chance(0.8) ? newcomer : r.pick([team[0], team[1], team[2]]);
        c.staff = to;
        c.next = Math.max(c.next, newcomer.from + r.int(0, 21));
        to.clients.push(c);
      }
      leaver.clients = [];
    }
    // Once a week: anyone more than three months overdue has gone elsewhere.
    if (wd === 2) {
      for (const c of allClients) if (!c.lapsed && d - c.next > 95) c.lapsed = true;
      for (const s of team) s.clients = s.clients.filter((c) => !c.lapsed);
    }

    if (closed) continue;

    // Stock, in the past only: deliveries ordered last week arrive; on Tuesday the order goes in.
    if (d <= 0) {
      const due = products.filter((p) => p.pending && p.arrives <= d);
      for (const s of suppliers) {
        const lines = due.filter((p) => p.supplier === s).map((p) => ({ p, n: p.qty }));
        if (lines.length) deliver(lines, d);
      }
      if (wd === 2) {
        for (const p of products) {
          if (p.pending || p.stock > p.level) continue;
          if (backOrdered.has(p) && d > -24) continue;
          let arrive = d + p.supplier.lead;
          while (!OPEN[cal.weekday(arrive)]) arrive++;
          p.pending = true;
          p.arrives = arrive;
        }
      }
    }

    // Gift vouchers over the desk: most in December, then before Mothering Sunday and Valentine's Day.
    if (d <= 0) {
      const rate = (month === 11 ? 2.2 : month === 2 ? 0.9 : month === 1 ? 0.6 : 0.3);
      const n = Math.floor(rate) + (r.chance(rate % 1) ? 1 : 0);
      for (let k = 0; k < n; k++) {
        const buyer = nameOf();
        const forWhom = nameOf(r.chance(0.15));
        const v = { Code: `GV-${voucherNo += 1}`, 'Sold On': date, 'Bought By': `${buyer.first} ${buyer.last}`, For: `${forWhom.first} ${forWhom.last}`, Value: r.weighted([[25, 3], [50, 4], [75, 2], [100, 2]]), Expires: cal.day(d + 365), expires: d + 365, 'Redeemed On': null, Booking: null };
        vouchers.push(v);
        if (r.chance(0.82)) voucherQueue.push({ v, at: d + r.int(10, 330) });
      }
    }

    // The appointment book: each person working today, from opening to closing.
    const season = month === 11 ? 1.35 : month === 0 ? 0.75 : month === 1 ? 0.9 : month === 7 ? 0.88 : 1;
    // A day still to come is only partly booked yet: most of next week, little of six weeks out.
    const bookedSoFar = d > 0 ? Math.exp(-d / 16) : 1;
    const visits = [];
    for (const s of team) {
      if (!s.days.includes(wd) || d < s.from || d >= s.to || s.off.has(d)) continue;
      const busy = Math.min(0.85, s.busy * season * bookedSoFar);
      let t = hours[0];
      // The regulars due by today, worked through as the column fills.
      const dueNow = s.clients.filter((c) => !c.lapsed && c.next <= d);
      while (t + 15 <= hours[1]) {
        if (!r.chance(busy)) { t += 15; continue; }
        // A regular who is due, now and then a new client, else a gap in the column.
        const pick = dueNow.length && !r.chance(0.04) ? r.int(0, dueNow.length - 1) : -1;
        let c = pick >= 0 ? dueNow[pick] : null;
        let profile = c ? c.profile : r.chance(dueNow.length ? 1 : 0.35) ? r.weighted(MIX[s.level]) : null;
        if (!profile) { t += 15; continue; }
        const probe = c || { profile };
        const done = servicesFor(probe, s, month);
        const minutes = done.reduce((n, x) => n + x.minutes, 0);
        if (t + minutes > hours[1]) { t += 15; continue; }
        if (pick >= 0) dueNow.splice(pick, 1);

        // When it was booked: regulars often rebook at the desk weeks ahead; a booking still to come was made in the last few weeks.
        const lead = d > 0 ? d + r.int(0, 20) : r.weighted([[0, 1], [r.int(1, 7), 4], [r.int(8, 21), 3.5], [r.int(22, 49), 2]]);
        // Never before a regular's last visit: that is when they rebooked at the soonest.
        let bookedOn = Math.min(0, Math.max(d - lead, c && c.last !== undefined ? c.last : -Infinity));
        const isNew = !c;
        const needsPatch = done.some((x) => x.patch);
        if (isNew) {
          if (needsPatch) bookedOn = Math.min(bookedOn, d - 3);
          c = makeClient(profile, s, bookedOn, false);
        }
        // Patch test: done two to seven days before when the last is over six months old.
        // One still to do for a booking more than two days away is what Patch tests due finds.
        if (needsPatch && (c.patch === null || d - c.patch > 180) && d <= 2) {
          c.patch = Math.min(0, d - r.int(2, Math.min(7, Math.max(2, d - bookedOn))));
          c.retested = true;
        }

        const firstVisit = c.fresh;
        const prices = done.map((x) => money((x.prices[s.level] ?? x.prices.Stylist) * (firstVisit ? 0.8 : 1)));
        const bill = money(prices.reduce((n, x) => n + x, 0));
        let status = 'Booked';
        let cancelledOn = null;
        if (d < 0) status = r.weighted([['Completed', 92.5], ['Cancelled', 5], ['No-show', c.unreliable ? 12 : 1.6]]);
        else if (r.chance(0.025)) status = 'Cancelled';
        if (status === 'Cancelled') cancelledOn = r.int(bookedOn, Math.min(d, 0));
        // A deposit for new clients and long appointments, and for anyone who has missed one before.
        // Never more than the bill: a child's cut or a fringe trim takes none.
        const deposit = bill >= 80 ? 20 : (isNew || c.missed) && bill >= 15 ? 10 : 0;
        const late = status === 'No-show' || (status === 'Cancelled' && d - cancelledOn < 2);
        let paidBy = null;
        let voucher = null;
        if (status === 'Completed') {
          // A voucher is spent in full, so only on a bill it does not exceed once the deposit is off.
          const i = voucherQueue.findIndex((x) => x.at <= d && x.v.expires >= d && x.v.Value <= bill - deposit);
          if (i >= 0) { voucher = voucherQueue.splice(i, 1)[0].v; paidBy = 'Gift voucher'; } else paidBy = r.weighted([['Card', 8], ['Cash', 1.4]]);
        }
        const [hh, mm] = hhmm(t);
        const booking = add('Bookings', {
          Client: c.id, Staff: s.id, Start: cal.time(d, hh, mm), Minutes: minutes, Status: status, 'Booked On': cal.day(bookedOn),
          'Booked Via': !isNew && d - bookedOn > 20 && r.chance(0.6) ? 'In salon' : r.weighted([['Online', 5], ['Phone', 4]]),
          Deposit: deposit, 'Deposit Refunded': status === 'Cancelled' && deposit > 0 && !late, 'Cancelled On': cancelledOn === null ? null : cal.day(cancelledOn),
          'Paid By': paidBy, Notes: r.chance(isNew ? 0.2 : 0.015) ? (isNew ? BOOKING_NOTES[4] : r.pick(BOOKING_NOTES.slice(0, 4))) : null,
        });
        done.forEach((x, i) => add('Booking Services', { Booking: booking, Service: x.id, Price: prices[i] }));
        if (voucher) { voucher['Redeemed On'] = date; voucher.Booking = booking; }
        // A visit paid for at the desk: its receipt is written when the day is done, in time order.
        const visit = status === 'Completed' ? { booking, at: t + minutes, s, c, deposit, voucher, paidBy, retail: [],
          lines: done.map((x, i) => ({ Kind: 'Service', Item: firstVisit ? `${x.name} (first visit, 20 per cent off)` : x.name, Quantity: 1, 'Unit Price': prices[i], 'Tax Rate': 0.2, category: x.category })) } : null;
        if (visit) visits.push(visit);

        c.last = d;
        if (status === 'No-show') c.missed = true;
        if (status === 'Completed') {
          c.fresh = false;
          c.visits++;
          // What was used at the basin and the trolley.
          for (const x of done) {
            for (const [name, units] of [...(USES[x.name] || []), ...(x.hair ? BASIN : [])]) {
              const p = byName.get(name);
              // Whole units come off the shelf; with the shelf empty the stylist makes do and nothing is counted.
              p.use += units;
              if (p.use >= 1) { const whole = Math.floor(p.use); p.stock -= Math.min(whole, p.stock); p.use -= whole; }
            }
          }
          // Retail over the desk at the end of the appointment.
          if (r.chance((s.level === 'Junior' ? 0.07 : s.level === 'Therapist' ? 0.12 : 0.2) * (month === 11 ? 1.5 : 1))) {
            const pool = retail.filter((p) => p.stock > 0 && (p !== giftSet || month >= 10));
            const items = r.chance(0.2) ? 2 : 1;
            for (const p of pool.length ? [...new Set(Array.from({ length: items }, () => r.weighted(pool.map((x) => [x, x.demand + (x === giftSet && month === 11 ? 3 : 0)]))))] : []) {
              const q = Math.min(p.stock, r.chance(0.1) ? 2 : 1);
              visit.retail.push({ p, q });
              p.stock -= q;
            }
          }
          if (r.chance(c.unreliable ? 0.08 : 0.035)) c.lapsed = true;
        }
        c.next = status === 'Cancelled' ? d + r.int(4, 14) : status === 'No-show' ? d + r.int(10, 35) : d + Math.round(r.int(c.cadence[0], c.cadence[1]) * (0.9 + r.next() * 0.3));
        t += minutes;
      }
    }
    // The day's receipts, numbered in the order they were taken at the desk.
    visits.sort((a, b) => a.at - b.at || a.booking - b.booking);
    for (const v of visits) receipt(v, d, date);
  }

  // Vouchers in the order they were sold.
  for (const v of vouchers) {
    const { expires, ...values } = v;
    add('Gift Vouchers', values);
  }
  // Commission statements: made on the first of the month for the month before, for each person on
  // commission who worked in it, and paid with the wages on the 7th. The first, part month is left out.
  const firsts = [];
  for (let d = -days; d <= 0; d++) if (cal.at(d).getDate() === 1) firsts.push(d);
  const WORK = [...MENU.map(([cat]) => cat), 'Retail'];
  const statementCount = {};
  for (let i = 0; i + 1 < firsts.length; i++) {
    const month = cal.day(firsts[i]).slice(0, 7);
    const made = firsts[i + 1];
    for (const s of team) {
      const e = earned.get(`${s.id}|${month}`);
      if (!e) continue;
      const y = cal.at(made).getFullYear();
      statementCount[y] = (statementCount[y] || 0) + 1;
      const lines = WORK.filter((w) => e.work.has(w)).map((w) => ({ w, ...e.work.get(w) }));
      const due = money(lines.reduce((n, l) => n + money(l.commission), 0));
      const id = add('Commission Statements', {
        Number: `COM-${y}-${pad(statementCount[y], 4)}`, Staff: s.id, Month: month, 'Prepared On': cal.day(made),
        'Commission Due': due, 'Paid On': made + 6 <= 0 ? cal.day(made + 6) : null,
      });
      lines.forEach((l, k) => add('Commission Lines', {
        Statement: id, Line: k + 1, Work: l.w === 'Retail' ? 'Retail products sold' : l.w,
        Done: l.done, Takings: money(l.takings), Rate: l.rate, Commission: money(l.commission),
      }));
    }
  }

  // Each client's latest patch test, and the stock as it stands.
  for (const c of allClients) if (c.retested) base.update('Clients', c.id, { 'Patch Test Date': cal.day(c.patch) });
  for (const p of products) base.update('Products', p.id, { 'In Stock': p.stock });
  return counts;
}
