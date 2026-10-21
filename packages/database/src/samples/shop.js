// The shop: Larkspur Corner Store, an invented neighbourhood shop. What it
// sells in categories, who supplies it, its till (each sale and its lines,
// paid by cash, card or a mobile wallet), the loyalty customers, the orders
// it sends its suppliers and the deliveries in, stock, the staff on the till,
// and what the shop costs to run.
//
// Its records are a year of trading a day at a time: busier at
// lunchtime and on Saturdays, an order to each supplier the evening before
// its delivery day for what has run low, rent and wages each month. Prices
// on the shelf include VAT: most food carries none, while drinks, snacks,
// household goods and cards carry the standard rate, and each receipt shows
// the VAT within its total. Its queries are the ones a shopkeeper asks at
// closing time and at the month's end: today's takings, what sells, what
// earns, what to order, the VAT for the return.

import { address, BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, companies, money, person, phoneFor, PLACES } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const LINE = 'Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)';
// The VAT inside a price that includes it, as the till works it out for each line.
const LINE_VAT = `Round(${LINE} - ${LINE} / (1 + [Sale Lines].[Tax Rate]), 2)`;
const DAYS_AGO = 'DateDiff("d", Format([When], "yyyy-mm-dd"), Date())';
// The shop's own details at the head of what it prints.
const LETTERHEAD = { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'Tax Number'] };

export const spec = {
  tables: [
    BUSINESS_TABLE,
    { name: 'Categories', fields: [{ name: 'Name', type: 'text', required: true, unique: true }] },
    { name: 'Suppliers', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Contact', type: 'text' }, { name: 'Address', type: 'text' },
      { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' }, { name: 'Phone', type: 'text' },
      { name: 'E-mail', type: 'text' }, { name: 'Account Number', type: 'text' }, { name: 'Delivers On', type: 'text' },
    ] },
    { name: 'Products', fields: [
      { name: 'Barcode', type: 'text', unique: true }, { name: 'Name', type: 'text', required: true }, { name: 'Category', ...ref('Categories') },
      { name: 'Supplier', ...ref('Suppliers') }, { name: 'VAT Code', type: 'text', default: 'Z' }, { name: 'Tax Rate', type: 'number', default: 0 },
      { name: 'Cost Price', type: 'currency' }, { name: 'Sale Price', type: 'currency' },
      { name: 'In Stock', type: 'integer', default: 0 }, { name: 'Reorder Level', type: 'integer', default: 6 }, { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Staff', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Role', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'Started', type: 'date' },
    ] },
    { name: 'Customers', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Phone', type: 'text' },
      { name: 'E-mail', type: 'text' }, { name: 'Joined', type: 'date' }, { name: 'Points', type: 'integer', default: 0 },
    ] },
    { name: 'Sales', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'When', type: 'datetime' }, { name: 'Staff', ...ref('Staff') }, { name: 'Customer', ...ref('Customers') },
      { name: 'Payment', type: 'text', default: 'Card' }, { name: 'Total', type: 'currency' }, { name: 'VAT', type: 'currency' },
      { name: 'Standard Rated', type: 'currency' }, { name: 'Saved', type: 'currency', default: 0 },
      { name: 'Tendered', type: 'currency' }, { name: 'Change', type: 'currency' }, { name: 'Points Earned', type: 'integer' },
      { name: 'Refunded', type: 'yesno' },
    ] },
    { name: 'Sale Lines', fields: [
      { name: 'Sale', ...ref('Sales'), required: true }, { name: 'Product', ...ref('Products'), required: true },
      { name: 'Quantity', type: 'integer', default: 1 }, { name: 'Unit Price', type: 'currency' }, { name: 'Discount', type: 'number', default: 0 },
      { name: 'Tax Rate', type: 'number', default: 0 },
    ] },
    { name: 'Supplier Orders', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Supplier', ...ref('Suppliers'), required: true }, { name: 'Staff', ...ref('Staff') },
      { name: 'Ordered On', type: 'date' }, { name: 'Delivery Due', type: 'date' }, { name: 'Status', type: 'text', default: 'Sent' },
    ] },
    { name: 'Supplier Order Lines', fields: [
      { name: 'Order', ...ref('Supplier Orders'), required: true }, { name: 'Product', ...ref('Products'), required: true },
      { name: 'Quantity', type: 'integer' }, { name: 'Unit Cost', type: 'currency' },
    ] },
    { name: 'Deliveries', fields: [
      { name: 'Supplier', ...ref('Suppliers'), required: true }, { name: 'Order', ...ref('Supplier Orders') }, { name: 'Received', type: 'date' },
      { name: 'Invoice Number', type: 'text' }, { name: 'Paid', type: 'yesno' },
    ] },
    { name: 'Delivery Lines', fields: [
      { name: 'Delivery', ...ref('Deliveries'), required: true }, { name: 'Product', ...ref('Products'), required: true },
      { name: 'Quantity', type: 'integer' }, { name: 'Unit Cost', type: 'currency' },
    ] },
    { name: 'Expenses', fields: [
      { name: 'Date', type: 'date' }, { name: 'Kind', type: 'text' }, { name: 'Description', type: 'text' }, { name: 'Amount', type: 'currency' },
    ] },
  ],
  queries: {
    "Today's takings": { tables: ['Sales', 'Sale Lines'], columns: [
      { table: 'Sales', field: 'Payment', total: 'group', sort: 'asc' }, { expr: LINE, as: 'Takings', total: 'sum' },
      { table: 'Sale Lines', field: 'Quantity', as: 'Items', total: 'sum' },
      { expr: DAYS_AGO, as: 'Days Ago', total: 'where', criteria: ['0'] },
      { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Daily takings': { tables: ['Sales', 'Sale Lines'], columns: [
      { expr: 'Format([When], "yyyy-mm-dd")', as: 'Day', total: 'group', sort: 'desc' }, { expr: LINE, as: 'Takings', total: 'sum' },
      { table: 'Sale Lines', field: 'Quantity', as: 'Items', total: 'sum' },
      { expr: DAYS_AGO, as: 'Days Ago', total: 'where', criteria: ['Between 0 And 30'] },
      { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Monthly takings': { tables: ['Sales', 'Sale Lines'], columns: [
      { expr: 'Format([When], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { expr: LINE, as: 'Takings', total: 'sum' },
      { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Busy hours': { tables: ['Sales'], columns: [
      { expr: 'Mid([When], 12, 2)', as: 'Hour', total: 'group', sort: 'asc' }, { table: 'Sales', field: 'ID', as: 'Sales', total: 'count' },
      { expr: DAYS_AGO, as: 'Days Ago', total: 'where', criteria: ['Between 0 And 90'] },
    ] },
    'Best sellers': { tables: ['Sale Lines', 'Products', 'Sales'], columns: [
      { table: 'Products', field: 'Name', as: 'Product', total: 'group' }, { table: 'Sale Lines', field: 'Quantity', as: 'Sold', total: 'sum', sort: 'desc' },
      { expr: LINE, as: 'Takings', total: 'sum' }, { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
      { expr: DAYS_AGO, as: 'Days Ago', total: 'where', criteria: ['Between 0 And 30'] },
    ] },
    // Profit is worked out on the price without its VAT, which goes to the tax office, not the shop.
    'Profit by category': { tables: ['Sale Lines', 'Products', 'Categories', 'Sales'], columns: [
      { table: 'Categories', field: 'Name', as: 'Category', total: 'group' }, { expr: LINE, as: 'Takings', total: 'sum' },
      { expr: LINE_VAT, as: 'VAT', total: 'sum' },
      { expr: `Round(${LINE} - ${LINE_VAT} - [Quantity] * [Cost Price], 2)`, as: 'Profit', total: 'sum', sort: 'desc' },
      { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Takings by staff': { tables: ['Sales', 'Staff', 'Sale Lines'], columns: [
      { expr: '[First Name] & " " & [Last Name]', as: 'Staff Member', total: 'group' }, { table: 'Staff', field: 'Role', total: 'group' },
      { expr: LINE, as: 'Takings', total: 'sum', sort: 'desc' },
      { expr: DAYS_AGO, as: 'Days Ago', total: 'where', criteria: ['Between 0 And 30'] },
      { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Low stock': { tables: ['Products', 'Suppliers'], columns: [
      { table: 'Suppliers', field: 'Name', as: 'Supplier', sort: 'asc' }, { table: 'Suppliers', field: 'Delivers On' },
      { table: 'Products', field: 'Name', as: 'Product', sort: 'asc' },
      { table: 'Products', field: 'In Stock' }, { table: 'Products', field: 'Reorder Level' },
      { expr: '[In Stock] - [Reorder Level]', as: 'Spare', criteria: ['<= 0'], show: false }, { table: 'Products', field: 'Active', criteria: ['Yes'], show: false },
    ] },
    'Stock value': { tables: ['Products', 'Categories'], columns: [
      { table: 'Categories', field: 'Name', as: 'Category', total: 'group', sort: 'asc' }, { table: 'Products', field: 'In Stock', as: 'Items', total: 'sum' },
      { expr: 'Round([In Stock] * [Cost Price], 2)', as: 'At Cost', total: 'sum' }, { expr: 'Round([In Stock] * [Sale Price], 2)', as: 'At Sale Price', total: 'sum' },
    ] },
    'Costs by month': { tables: ['Expenses'], columns: [
      { expr: 'Format([Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Expenses', field: 'Kind', total: 'group' },
      { table: 'Expenses', field: 'Amount', as: 'Amount', total: 'sum' },
    ] },
    'Best customers': { tables: ['Customers', 'Sales', 'Sale Lines'], columns: [
      { expr: '[First Name] & " " & [Last Name]', as: 'Customer', total: 'group' }, { expr: LINE, as: 'Spent', total: 'sum', sort: 'desc' },
      { table: 'Customers', field: 'Points', as: 'Points', total: 'max' }, { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Unpaid deliveries': { tables: ['Deliveries', 'Suppliers', 'Delivery Lines'], columns: [
      { table: 'Deliveries', field: 'ID', as: 'Delivery', total: 'group' }, { table: 'Suppliers', field: 'Name', as: 'Supplier', total: 'group' },
      { table: 'Deliveries', field: 'Received', total: 'group', sort: 'asc' }, { table: 'Deliveries', field: 'Invoice Number', total: 'group' },
      { expr: 'Round([Quantity] * [Unit Cost], 2)', as: 'Owed', total: 'sum' }, { table: 'Deliveries', field: 'Paid', total: 'where', criteria: ['No'] },
    ] },
    'VAT on sales by month': { tables: ['Sales'], columns: [
      { expr: 'Format([When], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Sales', field: 'Total', as: 'Takings', total: 'sum' },
      { table: 'Sales', field: 'Standard Rated', as: 'At 20%', total: 'sum' }, { table: 'Sales', field: 'VAT', as: 'VAT', total: 'sum' },
      { expr: 'Round([Total] - [Standard Rated], 2)', as: 'At 0% or Exempt', total: 'sum' },
      { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    // The totals a receipt prints, worked out from its lines: what to type into a sale entered by hand.
    'Receipt totals': { tables: ['Sales', 'Sale Lines'], columns: [
      { table: 'Sales', field: 'Number', as: 'Receipt', total: 'group', sort: 'desc' }, { table: 'Sales', field: 'When', total: 'group' },
      { expr: LINE, as: 'Total', total: 'sum' }, { expr: LINE_VAT, as: 'VAT', total: 'sum' },
      { expr: `IIf([Sale Lines].[Tax Rate] > 0, ${LINE}, 0)`, as: 'Standard Rated', total: 'sum' },
      { expr: 'Round([Quantity] * [Unit Price] * Nz([Discount]), 2)', as: 'Saved', total: 'sum' },
      { expr: DAYS_AGO, as: 'Days Ago', total: 'where', criteria: ['Between 0 And 1'] },
    ] },
    'Supplier orders, last 7 days': { tables: ['Supplier Orders', 'Suppliers'], columns: [
      { table: 'Supplier Orders', field: 'Number', sort: 'desc' }, { table: 'Suppliers', field: 'Name', as: 'Supplier' },
      { table: 'Supplier Orders', field: 'Ordered On' }, { table: 'Supplier Orders', field: 'Delivery Due' }, { table: 'Supplier Orders', field: 'Status' },
      { expr: 'DateDiff("d", [Ordered On], Date())', as: 'Days Ago', criteria: ['Between 0 And 6'], show: false },
    ] },
  },
  reports: [
    { name: 'Takings by month', query: 'Monthly takings', columns: ['Month', { field: 'Takings', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }] },
    { name: 'Last 30 days', title: 'Takings, the last 30 days', query: 'Daily takings', columns: [{ field: 'Day', type: 'date' }, { field: 'Items', type: 'integer' }, { field: 'Takings', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }, { field: 'Takings', fn: 'avg' }] },
    { name: 'Profit by category', query: 'Profit by category', columns: ['Category', { field: 'Takings', type: 'currency' }, { field: 'VAT', type: 'currency' }, { field: 'Profit', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }, { field: 'VAT', fn: 'sum' }, { field: 'Profit', fn: 'sum' }] },
    { name: 'Shopping list', title: 'To order', query: 'Low stock', group: 'Supplier', columns: ['Product', { field: 'In Stock', type: 'integer' }, { field: 'Reorder Level', type: 'integer' }] },
    { name: 'Stock take sheet', table: 'Products', group: 'Category', sort: 'Name', columns: ['Name', 'Barcode', 'In Stock'], totals: [{ field: 'In Stock', fn: 'sum' }] },
    { name: 'Running costs', query: 'Costs by month', group: 'Month', columns: ['Kind', { field: 'Amount', type: 'currency' }], totals: [{ field: 'Amount', fn: 'sum' }] },
    { name: 'Deliveries to pay', query: 'Unpaid deliveries', group: 'Supplier', columns: ['Delivery', { field: 'Received', type: 'date' }, 'Invoice Number', { field: 'Owed', type: 'currency' }], totals: [{ field: 'Owed', fn: 'sum' }] },
    { name: 'VAT on sales', title: 'VAT on sales, by month', query: 'VAT on sales by month', columns: ['Month', { field: 'Takings', type: 'currency' }, { field: 'At 20%', type: 'currency' }, { field: 'VAT', type: 'currency' }, { field: 'At 0% or Exempt', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }, { field: 'VAT', fn: 'sum' }] },
  ],
  documents: [
    { name: 'Sale receipt', table: 'Sales', paper: 'receipt-80', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Receipt {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'When', label: 'Date and time' }, { field: 'Staff.First Name', label: 'Served by' },
        { field: 'Payment', label: 'Paid by' }, { field: 'Customer', label: 'Loyalty card' },
      ] },
      { type: 'lines', table: 'Sale Lines', columns: [
        { field: 'Product.Name', label: 'Item' }, { field: 'Quantity', label: 'Qty', type: 'integer' },
        { field: 'Unit Price', label: 'Price', type: 'currency' }, { expr: LINE, as: 'Amount', type: 'currency' },
        { field: 'Product.VAT Code', label: 'VAT' },
      ] },
      { type: 'totals', rows: [
        { label: 'Total', expr: '[Total]', strong: true },
        { label: 'You saved', expr: '[Saved]' },
        { label: 'Cash tendered', expr: '[Tendered]' }, { label: 'Change', expr: '[Change]' },
        { label: 'S: sales at 20% VAT', expr: '[Standard Rated]' }, { label: 'VAT at 20% included', expr: '[VAT]' },
        { label: 'Z, E, O: sales at 0% or exempt', expr: 'Round([Total] - [Standard Rated], 2)' },
        { label: 'Loyalty points earned', expr: '[Points Earned]', type: 'integer' },
      ] },
      { type: 'text', text: 'VAT codes: S standard rate 20%, Z zero rate, E exempt, O outside the scope of VAT. VAT number {Business.Tax Number}.', small: true },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    { name: 'Supplier order', table: 'Supplier Orders', paper: 'a4', letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Tax Number'] }, blocks: [
      { type: 'heading', text: 'Order {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Supplier.Name', label: 'To' }, { field: 'Number', label: 'Order number' },
        { field: 'Supplier.Contact', label: 'For the attention of' }, { field: 'Ordered On', label: 'Date' },
        { field: 'Supplier.Address', label: 'Address' }, { field: 'Delivery Due', label: 'Deliver on' },
        { field: 'Supplier.City', label: 'City' }, { field: 'Supplier.Account Number', label: 'Our account' },
        { field: 'Supplier.Postcode', label: 'Postcode' }, { field: 'Staff', label: 'Ordered by' },
      ] },
      { type: 'lines', table: 'Supplier Order Lines', link: 'Order', columns: [
        { field: 'Product.Barcode', label: 'Barcode' }, { field: 'Product.Name', label: 'Product' },
        { field: 'Quantity', type: 'integer' }, { field: 'Unit Cost', type: 'currency' },
        { expr: 'Round([Quantity] * [Unit Cost], 2)', as: 'Line Total', label: 'Total', type: 'currency' },
      ] },
      { type: 'totals', rows: [{ label: 'Items', expr: '[Quantity]', type: 'integer' }, { label: 'Total before VAT', expr: '[Line Total]', strong: true }] },
      { type: 'text', text: 'Please deliver on {Delivery Due} to {Business.Name}, {Business.Address}, {Business.City} {Business.Postcode}, and quote {Number} on your invoice. If a line is out of stock, please tell us on the day rather than send a substitute.', small: true },
      { type: 'text', text: 'Chilled and fresh lines are checked on arrival and will be refused with less than three days left on their date.', small: true },
    ] },
  ],
};

export const about = {
  name: 'Larkspur Corner Store',
  kind: 'Shop',
  line: 'Products, the till and its receipts, VAT, stock, supplier orders, deliveries, loyalty customers and running costs for a neighbourhood shop.',
  summary: [
    'Larkspur Corner Store is an invented neighbourhood shop open from seven in the morning to nine at night, every day of the week. It sells bread, milk, fruit and vegetables, drinks, snacks, groceries, household goods and newspapers, from about forty lines bought from six suppliers.',
    'The sample holds a year of trading made up a day at a time. The till rings up a sale every few minutes, more at lunchtime and after school and more on Saturdays, and every sale has a numbered receipt. The evening before each supplier\'s delivery day the owner sends it an order for what has run low, so stock goes down between deliveries and back up again; now and then a supplier is short of a line. At each month end the rent, wages, power and card fees are paid, so the takings can be set against what the shop costs to run.',
    'Prices on the shelf include VAT, as a shop\'s do. Most food is zero-rated, while soft drinks, alcohol, sweets, crisps, household goods and greeting cards carry the standard rate of twenty percent, and stamps and phone top-ups carry none. Each product has its VAT code, each receipt shows the VAT within its total, and a query adds up the VAT on the month\'s sales for the VAT return.',
    'Use it to see how a till, stock and suppliers fit together in one database, or start from it empty for a shop of your own.',
  ],
  suits: ['A convenience store or minimarket', 'A farm shop', 'A newsagent', 'A gift shop', 'A pharmacy counter for the goods it sells off the shelf', 'A market stall'],
  tables: {
    Business: BUSINESS_ABOUT,
    Categories: { about: 'The sections of the shop, used to group products on reports.', fields: { Name: 'The section, such as Bakery or Drinks. Each name is used once.' } },
    Suppliers: { about: 'The wholesalers and producers the shop buys from.', fields: {
      Name: 'The supplier\'s name.', Contact: 'The person you deal with there.', Address: 'The supplier\'s street address, printed on your orders.',
      City: 'The town or city.', Postcode: 'The postcode.', Phone: 'Their phone number.', 'E-mail': 'Where orders and queries go.',
      'Account Number': 'The shop\'s account number with the supplier, quoted on every order.',
      'Delivers On': 'The days the supplier comes, so you know when an order will arrive.',
    } },
    Products: { about: 'Everything on the shelves, one record per line you sell.', fields: {
      Barcode: 'The number on the pack the till scans. Each barcode is used once.', Name: 'The product as it shows on the till and the receipt.',
      Category: 'The section it belongs to.', Supplier: 'Who you buy it from.',
      'VAT Code': 'The letter the receipt prints beside the line: S for the standard rate, Z for zero-rated food and papers, E for exempt (stamps), O for outside the scope of VAT (phone top-ups).',
      'Tax Rate': 'The VAT rate as a fraction: 0.2 for the standard rate, 0 for everything else. The till copies it to each sale line.',
      'Cost Price': 'What one costs you to buy, before VAT.',
      'Sale Price': 'What one sells for, with VAT included, as it is on the shelf label.', 'In Stock': 'How many are on the shelf and in the back now.',
      'Reorder Level': 'When stock falls to this number or below, it goes on the shopping list.', Active: 'Untick a product you no longer sell; it stays in the old sales.',
    } },
    Staff: { about: 'The people who work the till.', fields: {
      'First Name': 'Their first name, printed on the receipts they ring up.', 'Last Name': 'Their family name.', Role: 'Owner, shop assistant or weekend assistant.',
      Phone: 'A number to reach them.', Started: 'The day they started.',
    } },
    Customers: { about: 'Regulars who signed up for the loyalty card.', fields: {
      'First Name': 'Their first name.', 'Last Name': 'Their family name.', Phone: 'Their mobile number.', 'E-mail': 'For offers, if they asked for them.',
      Joined: 'The day they signed up.', Points: 'Loyalty points earned: one for each whole amount spent.',
    } },
    Sales: { about: 'Each time the till is rung up: one record per receipt, with the totals the receipt prints.', fields: {
      Number: 'The receipt number, given in time order and once only: RCP-2026-01234, a new run each year.',
      When: 'The date and time of the sale.', Staff: 'Who was on the till.', Customer: 'The loyalty customer, when they showed their card.',
      Payment: 'Card, Cash or Mobile.', Total: 'What the customer paid, VAT included.',
      VAT: 'The VAT inside the total: each line\'s VAT rounded to the penny, then added up.',
      'Standard Rated': 'The part of the total, VAT included, for goods at the standard rate. The rest is zero-rated or exempt.',
      Saved: 'What the customer saved on discounted lines.', Tendered: 'For a cash sale, the cash handed over; empty for card and mobile.',
      Change: 'For a cash sale, the change given back.', 'Points Earned': 'The loyalty points this sale earned; empty when no card was shown.',
      Refunded: 'Ticked when the whole sale was refunded; refunded sales are left out of the takings.',
    } },
    'Sale Lines': { about: 'The items on each receipt.', fields: {
      Sale: 'The receipt the line belongs to.', Product: 'What was sold.', Quantity: 'How many.',
      'Unit Price': 'The price of one at the time, VAT included, kept so a later price change does not alter old sales.', Discount: 'Any discount, as a fraction: 0.1 is ten percent off.',
      'Tax Rate': 'The VAT rate of the line, copied from the product when it was sold.',
    } },
    'Supplier Orders': { about: 'The orders the shop sends its suppliers, one record each, made from the shopping list the evening before the supplier\'s delivery day.', fields: {
      Number: 'The order\'s number, given in date order and once only: ORD-2026-0042, a new run each year.', Supplier: 'Who the order goes to.',
      Staff: 'Who sent it.', 'Ordered On': 'The day it was sent.', 'Delivery Due': 'The day the goods should come.',
      Status: 'Sent (not delivered yet), Delivered, Part delivered (some lines did not come) or Not delivered.',
    } },
    'Supplier Order Lines': { about: 'The products on each supplier order.', fields: {
      Order: 'The order the line belongs to.', Product: 'What was ordered.', Quantity: 'How many.', 'Unit Cost': 'What one costs from this supplier, before VAT.',
    } },
    Deliveries: { about: 'Each delivery from a supplier, with its invoice.', fields: {
      Supplier: 'Who delivered.', Order: 'The order it answers.', Received: 'The day it came.', 'Invoice Number': 'The number on the supplier\'s invoice.', Paid: 'Ticked once the invoice is paid.',
    } },
    'Delivery Lines': { about: 'What came in each delivery.', fields: {
      Delivery: 'The delivery the line belongs to.', Product: 'What came.', Quantity: 'How many.', 'Unit Cost': 'What one cost on this invoice.',
    } },
    Expenses: { about: 'What the shop costs to run, other than the goods it sells.', fields: {
      Date: 'The day it was paid.', Kind: 'Rent, Wages, Utilities, Card fees, Insurance, Phone and internet or Repairs.', Description: 'What it was for.', Amount: 'How much.',
    } },
  },
  queries: {
    "Today's takings": 'What the till has taken today so far, by how people paid, and how many items went out. Run it at closing time to check against the cash drawer and the card terminal.',
    'Daily takings': 'Takings and items sold on each of the last 30 days, the newest first, to spot a slow day.',
    'Monthly takings': 'Takings for each month of the year, refunds left out.',
    'Busy hours': 'How many sales the till rang up in each hour of the day over the last 90 days, to plan who works when.',
    'Best sellers': 'The products that sold most in the last 30 days, with what they took.',
    'Profit by category': 'For each section of the shop, what it took, the VAT within that, and what it earned after the VAT and the cost of the goods, the best earner first.',
    'Takings by staff': 'What each person rang up over the last 30 days.',
    'Low stock': 'Products at or below their reorder level, by supplier, with the days each supplier delivers. This is the shopping list each supplier order is made from.',
    'Stock value': 'What the stock on the shelves is worth by section, at cost and at sale price.',
    'Costs by month': 'What the shop paid out each month, by kind of cost.',
    'Best customers': 'Loyalty customers by what they have spent, with their points.',
    'Unpaid deliveries': 'Supplier invoices not yet paid, oldest first, with what each one comes to.',
    'VAT on sales by month': 'For each month, the takings, the part at the standard rate, the VAT within it and the part at zero or exempt: the output VAT for the VAT return, refunds left out.',
    'Receipt totals': 'Today\'s and yesterday\'s receipts with their total, VAT, standard-rated part and saving worked out from their lines. When you enter a sale by hand, read its figures here and type them into the sale.',
    'Supplier orders, last 7 days': 'The orders sent to suppliers in the last week, the newest first, with when each is due and whether it came.',
  },
  reports: {
    'Takings by month': 'One line a month with the takings and a total for the year.',
    'Last 30 days': 'Takings and items for each of the last 30 days, with the total and the average day.',
    'Profit by category': 'Takings, VAT and profit by section, with totals: the page to look at when deciding what to stock more of.',
    'Shopping list': 'What to order, on a page for each supplier, with what is left and the reorder level.',
    'Stock take sheet': 'Every product by section with its barcode and what the database says is in stock, to count against.',
    'Running costs': 'Each month\'s costs by kind, with a total for the month.',
    'Deliveries to pay': 'Supplier invoices still to pay, by supplier, with a total for each.',
    'VAT on sales': 'The VAT on sales month by month, for the accountant or the VAT return.',
  },
  documents: {
    'Sale receipt': 'The till receipt on an 80 mm roll: the shop\'s name, address and VAT number, the receipt number, date, time and who served, each item with its price and VAT code, the total, the VAT included by rate, the cash tendered and change for cash sales, loyalty points earned, and the returns policy from the Business record. Print it from the Sales form for the sale on screen.',
    'Supplier order': 'The order sent to a supplier: its number, the shop\'s account with them, the delivery day, each product with its barcode, quantity and cost, and the total before VAT. Print it from the Supplier Orders form for the order on screen, or e-mail it as a PDF.',
  },
  tutorial: [
    { title: 'Ring up a sale and print its receipt', steps: [
      'Open the Sales form and press New. Type the next receipt number in Number (look at the last one in the Sales table: RCP-2026-01234 is followed by RCP-2026-01235), and the date and time in When.',
      'Choose who is on the till in Staff and, if the customer shows a loyalty card, choose them in Customer. Choose Card, Cash or Mobile in Payment.',
      'In the Sale Lines below, choose each Product, type the Quantity and the Unit Price from the shelf label, and copy the product\'s Tax Rate (0.2 or 0).',
      'Run the query Receipt totals and find the sale. Type its Total, VAT, Standard Rated and Saved into the sale. For cash, type what was handed over in Tendered and the difference in Change.',
      'Open the print template Sale receipt from Print templates, choose the sale, and press Print, or press Print on the Sales form to print the sale on screen.',
    ] },
    { title: 'Cash up at closing time', steps: [
      'Run the query Today\'s takings. It shows the takings by Card, Cash and Mobile, with refunds left out.',
      'Count the cash drawer and set it against the Cash line, less the float you started with. Set the Card and Mobile lines against the card terminal\'s end-of-day slip.',
      'If a sale was refunded, open it in the Sales form and tick Refunded, then run Today\'s takings again.',
    ] },
    { title: 'Order from the shopping list', steps: [
      'The evening before a supplier comes, open the report Shopping list. It lists, on a page for each supplier, what is at or below its reorder level.',
      'Open the Supplier Orders form and press New. Type the next number in Number (ORD-2026-0042 is followed by ORD-2026-0043), choose the Supplier, and type today in Ordered On and the delivery day in Delivery Due. Status stays Sent.',
      'In the Supplier Order Lines below, choose each product from the shopping list and type how many to order and the Unit Cost.',
      'Open the print template Supplier order, choose the order, and press Save as PDF to e-mail it, or Print to hand it to the driver.',
    ] },
    { title: 'Take in a delivery', steps: [
      'When the goods come, open the Deliveries form and press New. Choose the Supplier and the Order it answers, type the day in Received and the number on the supplier\'s invoice in Invoice Number.',
      'In the Delivery Lines below, enter each product that came, how many, and the Unit Cost on the invoice.',
      'Add what came to each product\'s In Stock in the Products form, and set the order\'s Status in Supplier Orders to Delivered, or Part delivered if something was short.',
      'When you pay the invoice, tick Paid on the delivery. The report Deliveries to pay lists what is still owed.',
    ] },
    { title: 'Do the month end', steps: [
      'On the last day of the month, open the Expenses form and enter the rent, wages, power, card fees and anything else paid, one record each.',
      'Open the report Takings by month and the report Running costs to set the month\'s takings against its costs.',
      'Open the report Profit by category to see which sections earned most after VAT and the cost of the goods.',
      'For the VAT return, open the report VAT on sales and take the VAT for the months of the quarter. The VAT you paid on what you bought is on your suppliers\' invoices.',
    ] },
    { title: 'Count the stock', steps: [
      'Print the report Stock take sheet. It lists every product by section with its barcode and what the database thinks is in stock.',
      'Count the shelves and the back room and write the real number beside each line.',
      'Open the Products form and type the counted numbers into In Stock where they differ.',
    ] },
  ],
  notes: [
    'Prices in Products and Sale Lines include VAT, as the shelf labels do. Cost Price and Unit Cost on orders and deliveries are before VAT, as suppliers invoice them. Profit by category takes the VAT off the takings before it takes off the cost.',
    'A sale\'s Total, VAT, Standard Rated and Saved are what the till worked out and printed when the sale was made, each line\'s VAT rounded to the penny and then added up. They are stored, not worked out again, so the receipt reprints exactly as it was; the query Receipt totals works them out from the lines when you enter a sale by hand.',
    'Only one VAT rate other than nought is used: twenty percent. The receipt\'s VAT summary and the VAT queries count everything with a rate above nought as standard rated. If you sell goods at a reduced rate, see the ways to adapt it.',
    'In Stock goes down as the till sells and up as deliveries come in, in the sample\'s records. Entering a sale or a delivery by hand does not change it by itself: change In Stock in the Products form, or count it at a stock take.',
    'The print templates print every record of their table unless you choose one or a few: print a receipt from the Sales form for the sale on screen, and an order from the Supplier Orders form.',
    'Refunded sales stay in the Sales table with their receipt numbers, so no number is missing; they are left out of the takings, the VAT and the best sellers.',
  ],
  faq: [
    { q: 'How do I reprint a customer\'s receipt?', a: 'Find the sale in the Sales form, by its receipt number or by the date and time, and press Print with the print template Sale receipt. It prints from the stored totals, so it reads as the first one did.' },
    { q: 'Where do I see how much VAT I owe?', a: 'Open the report VAT on sales: the VAT column is the VAT within your takings each month. Add up the months of your VAT quarter and take off the VAT on your suppliers\' invoices, which this database does not keep.' },
    { q: 'Why is there no VAT on bread or milk?', a: 'Most food is zero-rated, so its Tax Rate is 0 and its VAT Code is Z. Soft drinks, alcohol, sweets, crisps and household goods carry the standard rate. Check the rules for your own country and set each product\'s VAT Code and Tax Rate to match.' },
    { q: 'How does the shop know what to order?', a: 'Each product has a Reorder Level. The query Low stock and the report Shopping list show everything at or below it, by supplier, and the supplier order is made from that list.' },
    { q: 'How are loyalty points worked out?', a: 'One point for each whole pound spent on a sale, stored in the sale\'s Points Earned and added to the customer\'s Points. A refunded sale earns none.' },
    { q: 'A supplier did not bring everything. How do I show that?', a: 'Enter only what came in the delivery\'s lines, and set the order\'s Status to Part delivered. The lines left out are still on the order, so you can see what to chase.' },
    { q: 'Can I start with my own products and none of the sample\'s sales?', a: 'Yes. Open the sample empty, fill in Business first, then Categories, Suppliers and Products, then Staff. Sales, orders and deliveries come in as you trade.' },
  ],
  empty: 'Start with Business, the shop\'s own name, address and VAT number and the returns policy in Document Footer, which every receipt prints. Then Categories and Suppliers, then your Products with their VAT code, tax rate, prices and reorder levels, then your Staff. Sales, supplier orders and deliveries come in as you trade. Put your running costs in Expenses as you pay them.',
  adapt: [
    'Sell by weight: add a Unit field to Products (each, kg, 100 g) and make Quantity in Sale Lines a Number rather than a whole number.',
    'Track sell-by dates: add a Best Before field to Delivery Lines and a query for lines within three days of it.',
    'Sell at a reduced VAT rate as well: add a Reduced Rated field to Sales beside Standard Rated, and lines to the receipt\'s totals and the VAT query for it.',
    'Keep more than one till or shop: add a Tills table, point Sales at it, and group the takings queries by till.',
    'Reward points differently: change Points to count one for every two spent, or add a Points Redeemed field to Sales.',
    'Record the VAT you pay: add a VAT field to Deliveries for the VAT on each supplier invoice, and a query that sets it against VAT on sales by month.',
  ],
};

/** The shelves: [category, [name, cost, price, units a day]...]. Generic goods; no brand. Prices include VAT. */
const SHELVES = [
  ['Bakery', [['White sliced loaf', 0.62, 1.25, 9], ['Wholemeal loaf', 0.75, 1.45, 5], ['Croissants, pack of 4', 1.1, 2.2, 3], ['Flatbreads, pack of 6', 0.8, 1.6, 4]]],
  ['Dairy and Eggs', [['Milk, 2 litres', 0.98, 1.65, 14], ['Free-range eggs, 6', 1.05, 1.9, 6], ['Butter, 250 g', 1.3, 2.25, 4], ['Natural yoghurt, 500 g', 0.7, 1.35, 4], ['Cheddar, 400 g', 2.1, 3.6, 3]]],
  ['Fruit and Veg', [['Bananas, bunch', 0.55, 1.1, 8], ['Apples, bag of 6', 0.9, 1.75, 5], ['Onions, 1 kg', 0.45, 0.95, 5], ['Potatoes, 2.5 kg', 0.9, 1.85, 4], ['Tomatoes, 6', 0.7, 1.4, 4]]],
  ['Chilled and Lunch', [['Sandwich, freshly made', 1.6, 3.75, 9], ['Ready meal, chilled', 1.9, 3.5, 4], ['Salad pot', 1.2, 2.6, 3]]],
  ['Drinks', [['Bottled water, 500 ml', 0.18, 0.75, 16], ['Orange juice, 1 litre', 0.85, 1.6, 5], ['Cola, 330 ml can', 0.32, 0.95, 12], ['Instant coffee, 100 g', 1.9, 3.5, 2], ['Tea bags, 80', 1.2, 2.3, 3], ['Red wine, 75 cl', 4.2, 7.5, 3], ['Lager, 4 cans', 3.1, 5.25, 4]]],
  ['Snacks', [['Crisps, single bag', 0.25, 0.85, 14], ['Chocolate bar', 0.38, 0.99, 15], ['Biscuits, packet', 0.55, 1.2, 6], ['Mixed nuts, 200 g', 1.1, 2.25, 2]]],
  ['Groceries', [['Pasta, 500 g', 0.42, 0.95, 5], ['Basmati rice, 1 kg', 1.15, 2.2, 3], ['Chopped tomatoes, tin', 0.32, 0.7, 6], ['Baked beans, tin', 0.36, 0.8, 6], ['Cooking oil, 1 litre', 1.05, 1.95, 2]]],
  ['Household', [['Washing-up liquid', 0.6, 1.3, 2], ['Toilet roll, 9 pack', 2.4, 4.5, 3], ['Bin bags, roll of 20', 0.9, 1.9, 2], ['Laundry liquid, 1 litre', 2.1, 4.0, 1]]],
  ['Newsagent', [['Daily newspaper', 0.9, 1.2, 10], ['Greeting card', 0.7, 2.5, 2], ['Phone top-up voucher', 9.5, 10, 2], ['Postage stamps, book of 4', 2.6, 2.6, 1]]],
];

// VAT as a UK shop charges it: most food is zero-rated (Z); soft drinks, alcohol, sweets, crisps,
// roasted nuts, household goods and cards are standard-rated (S); stamps are exempt (E) and a
// phone top-up is outside the scope of VAT (O). A section's code, then the lines that differ.
const SECTION_VAT = { Drinks: 'S', Snacks: 'S', Household: 'S' };
const LINE_VAT_CODE = { 'Instant coffee, 100 g': 'Z', 'Tea bags, 80': 'Z', 'Biscuits, packet': 'Z', 'Greeting card': 'S', 'Phone top-up voucher': 'O', 'Postage stamps, book of 4': 'E' };
const RATE = { S: 0.2, Z: 0, E: 0, O: 0 };

/** Who brings what, and on which days (0 Sunday): [section names, days, the end of the supplier's name]. */
const ROUNDS = [
  [['Bakery', 'Chilled and Lunch'], [1, 3, 5], 'Bakery'],
  [['Dairy and Eggs'], [1, 2, 4, 6], 'Dairies'],
  [['Fruit and Veg'], [2, 5], 'Fresh Produce'],
  [['Drinks', 'Snacks'], [3], 'Drinks Wholesale'],
  [['Groceries', 'Household'], [4], 'Cash and Carry'],
  [['Newsagent'], [1, 2, 3, 4, 5, 6], 'News Distribution'],
];
const DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Sales an hour, open at 7: the morning papers, lunch, after school, a quiet evening. */
const HOURLY = { 7: 1.2, 8: 1, 9: 0.8, 10: 0.8, 11: 1, 12: 2, 13: 1.9, 14: 0.9, 15: 1.2, 16: 1.6, 17: 1.6, 18: 1.4, 19: 0.8, 20: 0.6 };
const PER_HOUR = 3.6;

/** Pence as SQLite's Round(x, 2) gives them, so a stored total is the one a query works out. */
// Quick for most amounts; near a half penny it asks toFixed, which rounds the exact value as SQLite does.
const cents = (n) => {
  const y = n * 100;
  const k = Math.round(y);
  return Math.abs(Math.abs(y - k) - 0.5) < 1e-6 ? Number(n.toFixed(2)) : k / 100;
};

/** What a customer hands over for `total` in cash: the exact money, or the next coin or note up. */
function tendered(r, total) {
  const up = [Math.ceil(total), Math.ceil(total / 5) * 5, Math.ceil(total / 10) * 10, total <= 20 ? 20 : Math.ceil(total / 20) * 20].filter((x) => x >= total);
  return r.chance(0.2) ? total : r.pick(up);
}

export function fill(base, { r, cal, days = 365 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };

  add('Business', businessRecord({
    name: 'Larkspur Corner Store', domain: 'larkspurstore.example', phone: '020 7946 0420', website: false,
    address: { Address: '27 Linden Close', City: 'Leeds', Postcode: 'LS6 2HJ', Country: 'United Kingdom' },
    terms: 'Paid at the till by card, cash or phone.',
    footer: 'Thank you for shopping at Larkspur. Open 7am to 9pm, every day. Unopened goods can be returned with this receipt within 14 days for a refund or an exchange; fresh and chilled food only if it is faulty. Your legal rights are not affected.',
  }));

  const staff = ['Owner', 'Shop Assistant', 'Shop Assistant', 'Weekend Assistant'].map((role, i) => {
    const p = person(r);
    return add('Staff', { 'First Name': p['First Name'], 'Last Name': p['Last Name'], Role: role, Phone: `07700 900${String(400 + i).padStart(3, '0')}`, Started: cal.day(-r.int(600, 2400)) });
  });

  const taken = new Set(['Larkspur']);
  const suppliers = ROUNDS.map(([sections, on, tail], i) => {
    let name;
    do name = companies(r, 1, [tail])[0]; while (taken.has(name.split(' ')[0]));
    taken.add(name.split(' ')[0]);
    const id = add('Suppliers', {
      Name: name, Contact: (({ 'First Name': f, 'Last Name': l }) => `${f} ${l}`)(person(r)),
      ...(({ Address, City, Postcode }) => ({ Address, City, Postcode }))(address(r, PLACES.slice(0, 1))),
      Phone: phoneFor(r, 'United Kingdom'), 'E-mail': `orders@${name.toLowerCase().replace(/[^a-z]+/g, '')}.example`,
      'Account Number': `LCS-${String(1040 + i * 37).padStart(5, '0')}`,
      'Delivers On': on.length === 6 ? 'Monday to Saturday' : on.map((d) => DAY_NAMES[d]).join(', '),
    });
    // The longest gap between two of its deliveries, in days: what the stock must last.
    const gap = Math.max(...on.map((d, k) => ((on[(k + 1) % on.length] - d + 7) % 7) || 7));
    return { id, sections, on, gap };
  });

  const products = [];
  let code = 5012345000000;
  for (const [cat, items] of SHELVES) {
    const catId = add('Categories', { Name: cat });
    const s = suppliers.find((x) => x.sections.includes(cat));
    for (const [name, cost, price, demand] of items) {
      // Enough to last to the next delivery and a day over; reorder at about half that.
      const par = Math.ceil(demand * (s.gap + 1.5));
      const level = Math.max(3, Math.ceil(demand * s.gap * 0.5));
      const vat = LINE_VAT_CODE[name] || SECTION_VAT[cat] || 'Z';
      const p = { name, cost, price, demand, par, level, stock: par, supplier: s, rate: RATE[vat] };
      p.id = add('Products', { Barcode: String(code += r.int(7, 97)), Name: name, Category: catId, Supplier: s.id, 'VAT Code': vat, 'Tax Rate': p.rate, 'Cost Price': cost, 'Sale Price': price, 'In Stock': p.stock, 'Reorder Level': level });
      products.push(p);
    }
  }

  const customers = [];
  for (let i = 0; i < 60; i++) {
    const p = person(r);
    customers.push({ id: add('Customers', { ...p, Phone: `07700 900${String(i * 11 + 3).padStart(3, '0')}`, Joined: cal.day(-r.int(30, days)) }), points: 0 });
  }

  // Numbers that run in date order, a new run each year: RCP-2026-01234, ORD-2026-0042.
  const runs = {};
  const numberFor = (prefix, d, width) => { const y = cal.at(d).getFullYear(); const k = `${prefix}${y}`; runs[k] = (runs[k] || 0) + 1; return `${prefix}-${y}-${String(runs[k]).padStart(width, '0')}`; };
  // An order for what has run low, sent the evening before the delivery day `d`.
  const order = (s, d, low, status) => {
    const id = add('Supplier Orders', { Number: numberFor('ORD', d - 1, 4), Supplier: s.id, Staff: staff[0], 'Ordered On': cal.day(d - 1), 'Delivery Due': cal.day(d), Status: status });
    for (const p of low) add('Supplier Order Lines', { Order: id, Product: p.id, Quantity: p.par - p.stock, 'Unit Cost': p.cost });
    return id;
  };

  // A year from the first of a month, so each month in the costs is a whole one.
  let start = -days;
  while (cal.at(start).getDate() !== 1) start++;
  let cardTakings = 0;
  // Two lines the suppliers have been short of for the last fortnight: there is always something to chase.
  const short = r.some(products.filter((p) => p.demand >= 3 && p.supplier.sections[0] !== 'Newsagent'), 2);

  for (let d = start; d <= 0; d++) {
    const wd = cal.weekday(d);
    // Deliveries: each supplier on its days, bringing what the order of the evening before asked for.
    for (const s of suppliers) {
      if (!s.on.includes(wd)) continue;
      const low = products.filter((p) => p.supplier === s && p.stock < p.par * 0.8);
      if (!low.length) continue;
      const sent = low.filter((p) => !(d > -14 && short.includes(p)));
      const ordered = order(s, d, low, sent.length === low.length ? 'Delivered' : sent.length ? 'Part delivered' : 'Not delivered');
      if (!sent.length) continue;
      const delivery = add('Deliveries', { Supplier: s.id, Order: ordered, Received: cal.day(d), 'Invoice Number': `INV-${r.int(10000, 99999)}`, Paid: d < -14 || r.chance(0.3) });
      for (const p of sent) {
        const n = p.par - p.stock;
        add('Delivery Lines', { Delivery: delivery, Product: p.id, Quantity: n, 'Unit Cost': p.cost });
        p.stock += n;
      }
    }
    // The till: open 7 to 21, busier at lunch and on Saturdays, quieter on Sundays and in August.
    const dayFactor = (wd === 6 ? 1.3 : wd === 0 ? 0.75 : 1) * (cal.at(d).getMonth() === 7 ? 0.85 : 1) * (cal.at(d).getMonth() === 11 ? 1.15 : 1);
    const closeAt = d === 0 ? Math.min(21, Math.max(10, cal.hour + 1)) : 21;
    for (let h = 7; h < closeAt; h++) {
      const n = Math.floor(r.next() * 2 * PER_HOUR * HOURLY[h] * dayFactor + 0.5);
      // The minutes in order, so the receipt numbers run in time order.
      const minutes = Array.from({ length: n }, () => r.int(0, 59)).sort((a, b) => a - b);
      for (const minute of minutes) {
        const items = [];
        const want = r.weighted([[1, 30], [2, 25], [3, 18], [4, 12], [5, 8], [6, 7]]);
        const shelf = products.filter((p) => p.stock > 0);
        for (let j = 0; j < want && shelf.length; j++) {
          const p = r.weighted(shelf.map((x) => [x, x.demand]));
          shelf.splice(shelf.indexOf(p), 1);
          const q = Math.min(p.stock, r.weighted([[1, 80], [2, 15], [3, 5]]));
          items.push({ p, q });
        }
        if (!items.length) continue;
        const loyal = r.chance(0.3) ? r.pick(customers) : null;
        const payment = r.weighted([['Card', 6], ['Cash', 3], ['Mobile', 2]]);
        const refunded = r.chance(0.003);
        // The receipt's figures as the till prints them: each line rounded, its VAT rounded, then added up.
        let total = 0;
        let vat = 0;
        let standard = 0;
        let saved = 0;
        for (const it of items) {
          // Ten percent off now and then: a short date or a multibuy.
          it.discount = r.chance(0.04) ? 0.1 : 0;
          const line = cents(it.q * it.p.price * (1 - it.discount));
          total += line;
          vat += cents(line - line / (1 + it.p.rate));
          if (it.p.rate > 0) standard += line;
          saved += cents(it.q * it.p.price * it.discount);
        }
        total = cents(total);
        const paid = payment === 'Cash' ? tendered(r, total) : null;
        const points = loyal ? (refunded ? 0 : Math.floor(total)) : null;
        const sale = add('Sales', {
          Number: numberFor('RCP', d, 5), When: cal.time(d, h, minute), Staff: wd === 0 || wd === 6 ? r.pick([staff[0], staff[3]]) : r.pick(staff.slice(0, 3)), Customer: loyal?.id ?? null,
          Payment: payment, Total: total, VAT: cents(vat), 'Standard Rated': cents(standard), Saved: cents(saved),
          Tendered: paid, Change: paid === null ? null : cents(paid - total), 'Points Earned': points, Refunded: refunded,
        });
        for (const { p, q, discount } of items) {
          add('Sale Lines', { Sale: sale, Product: p.id, Quantity: q, 'Unit Price': p.price, Discount: discount, 'Tax Rate': p.rate });
          if (!refunded) p.stock -= q;
        }
        if (loyal && !refunded) loyal.points += points;
        if (payment !== 'Cash' && !refunded) cardTakings += total;
      }
    }
    // Month's end: rent, wages, power, the card terminal's share, and now and then a repair.
    if (cal.at(d + 1).getDate() === 1) {
      const date = cal.day(d);
      const winter = [10, 11, 0, 1].includes(cal.at(d).getMonth());
      add('Expenses', { Date: date, Kind: 'Rent', Description: 'Shop rent', Amount: 950 });
      add('Expenses', { Date: date, Kind: 'Wages', Description: 'Part-time and weekend wages', Amount: money(1720 + r.next() * 180) });
      add('Expenses', { Date: date, Kind: 'Utilities', Description: 'Electricity and water', Amount: money((winter ? 390 : 300) + r.next() * 60) });
      add('Expenses', { Date: date, Kind: 'Card fees', Description: 'Card terminal fees', Amount: money(cardTakings * 0.011 + 15) });
      add('Expenses', { Date: date, Kind: 'Insurance', Description: 'Shop insurance', Amount: 42 });
      add('Expenses', { Date: date, Kind: 'Phone and internet', Description: 'Phone line and broadband', Amount: 38 });
      if (r.chance(0.25)) add('Expenses', { Date: date, Kind: 'Repairs', Description: r.pick(['Fridge service', 'Door lock', 'Shelving', 'Lighting']), Amount: money(60 + r.next() * 240) });
      cardTakings = 0;
    }
  }

  // Tonight's orders, sent for the suppliers who come tomorrow: not delivered yet.
  for (const s of suppliers) {
    if (!s.on.includes(cal.weekday(1))) continue;
    const low = products.filter((p) => p.supplier === s && p.stock < p.par * 0.8);
    if (low.length) order(s, 1, low, 'Sent');
  }

  for (const p of products) base.update('Products', p.id, { 'In Stock': p.stock });
  for (const c of customers) base.update('Customers', c.id, { Points: c.points });
  return counts;
}
