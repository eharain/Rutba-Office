// The car dealer: Copperbeech Motors, an invented independent used car
// dealer in south-east London with about forty cars on the forecourt. Cars
// bought at auction, from the trade and taken in part-exchange, what
// preparing each one cost (valet, service, repairs, MOT), its asking price
// and what it sold for, the customers and their test drives, sales with
// their deposits, finance arranged through two invented lenders and the
// commission each pays, warranties sold, and the part-exchange valuations
// given.
//
// Its records are fifteen months of trading, run a day at a time: twice a
// week the buyer tops the stock up at auction, each car is prepared for a
// few days before it goes on sale, prices come down when a car has stood
// for a month, customers test drive and some buy, paying a deposit and the
// rest at handover, often with finance and a part-exchange that comes into
// stock the day the new car leaves. March and September, when the new
// registrations come out, are busy; December is quiet.
//
// Every car is invented: its make and model are made-up names, and its
// registration begins ZZ, which is never issued, so none is a real car's.
// The documents are the ones a dealer prints: the sales invoice, the order
// and sales agreement, a part-exchange valuation letter, the warranty
// certificate and the card in the car's window.

import { BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, money, person, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
// The business's own details at the head of every document.
const LETTERHEAD = { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Website', 'Tax Number'] };
const CUSTOMER = '[Customers].[First Name] & " " & [Customers].[Last Name]';
const SALESPERSON = '[Staff].[First Name] & " " & [Staff].[Last Name]';
const VEHICLE = '[Year] & " " & [Make] & " " & [Model]';
// VAT on a used car is paid on the dealer's margin: a sixth of what it sold for above what it cost.
const MARGIN_VAT = 'IIf([Sales].[Price] > [Purchase Price], Round(([Sales].[Price] - [Purchase Price]) / 6, 2), 0)';
const MARGIN = `Round([Sales].[Price] - [Purchase Price] - Nz([Buying Fee]) - Nz([Prep Cost]) - ${MARGIN_VAT}, 2)`;
// The VAT within a tax-inclusive amount.
const VAT_IN = 'Round([Amount] - [Amount] / (1 + Nz([Tax Rate])), 2)';

export const spec = {
  // Where a value below nought is right: a car that sold at a loss, a reserved car's handover already passed.
  negative: { 'Margin per vehicle': ['Margin'], 'Sales by month': ['Gross Margin'], 'Sales by salesperson': ['Gross Margin'], 'Stock over 60 days': ['Margin at Asking'] },
  tables: [
    BUSINESS_TABLE,
    { name: 'Staff', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Role', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Started', type: 'date' }, { name: 'Sells', type: 'yesno' },
    ] },
    { name: 'Suppliers', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Kind', type: 'text' }, { name: 'Address', type: 'text' }, { name: 'City', type: 'text' },
      { name: 'Postcode', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Account No', type: 'text' },
    ] },
    { name: 'Customers', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Address', type: 'text' },
      { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' },
      { name: 'Heard Of Us', type: 'text' }, { name: 'Added', type: 'date' }, { name: 'Marketing Consent', type: 'yesno' },
    ] },
    // Registration is the first text field, so a car is known by its registration in every list.
    { name: 'Vehicles', fields: [
      { name: 'Registration', type: 'text', required: true, unique: true }, { name: 'Stock No', type: 'text', unique: true },
      { name: 'Make', type: 'text' }, { name: 'Model', type: 'text' }, { name: 'Body', type: 'text' }, { name: 'Year', type: 'integer' },
      { name: 'First Registered', type: 'date' }, { name: 'Colour', type: 'text' }, { name: 'Fuel', type: 'text' }, { name: 'Transmission', type: 'text' },
      { name: 'Doors', type: 'integer' }, { name: 'Mileage', type: 'integer' }, { name: 'Owners', type: 'integer' }, { name: 'Service History', type: 'text' },
      { name: 'MOT Expiry', type: 'date' }, { name: 'Keys', type: 'integer', default: 2 }, { name: 'Source', type: 'text' }, { name: 'Supplier', ...ref('Suppliers') },
      { name: 'Bought On', type: 'date' }, { name: 'Purchase Price', type: 'currency' }, { name: 'Buying Fee', type: 'currency', default: 0 },
      { name: 'Prep Cost', type: 'currency', default: 0 }, { name: 'First Asking Price', type: 'currency' }, { name: 'Asking Price', type: 'currency' },
      { name: 'VAT Qualifying', type: 'yesno' }, { name: 'Status', type: 'text', default: 'In preparation' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Vehicle Features', fields: [{ name: 'Vehicle', ...ref('Vehicles'), required: true }, { name: 'Feature', type: 'text', required: true }] },
    { name: 'Preparation', fields: [
      { name: 'Vehicle', ...ref('Vehicles'), required: true }, { name: 'Done On', type: 'date' }, { name: 'Kind', type: 'text' },
      { name: 'Description', type: 'text' }, { name: 'Done By', type: 'text' }, { name: 'Cost', type: 'currency' },
    ] },
    { name: 'Test Drives', fields: [
      { name: 'Vehicle', ...ref('Vehicles'), required: true }, { name: 'Customer', ...ref('Customers'), required: true }, { name: 'Staff', ...ref('Staff') },
      { name: 'When', type: 'datetime' }, { name: 'Outcome', type: 'text', default: 'Booked' }, { name: 'Notes', type: 'text' },
    ] },
    { name: 'Sales', fields: [
      { name: 'Order No', type: 'text', unique: true }, { name: 'Invoice No', type: 'text', unique: true },
      { name: 'Vehicle', ...ref('Vehicles'), required: true }, { name: 'Customer', ...ref('Customers'), required: true }, { name: 'Salesperson', ...ref('Staff') },
      { name: 'Agreed On', type: 'date' }, { name: 'Handover', type: 'date' }, { name: 'Price', type: 'currency' }, { name: 'Total', type: 'currency' },
      { name: 'VAT', type: 'currency' }, { name: 'VAT Scheme', type: 'text' }, { name: 'Part-Ex Allowance', type: 'currency', default: 0 },
      { name: 'Deposit', type: 'currency', default: 0 }, { name: 'Finance Amount', type: 'currency', default: 0 }, { name: 'Balance', type: 'currency' },
      { name: 'Status', type: 'text', default: 'Reserved' }, { name: 'Notes', type: 'text' },
    ] },
    { name: 'Sale Lines', fields: [
      { name: 'Sale', ...ref('Sales'), required: true }, { name: 'Line', type: 'integer' }, { name: 'Kind', type: 'text' },
      { name: 'Description', type: 'text' }, { name: 'Amount', type: 'currency' }, { name: 'Tax Rate', type: 'number', default: 0 },
    ] },
    { name: 'Lenders', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Address', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' },
      { name: 'Commission Rate', type: 'number' }, { name: 'Typical APR', type: 'number' },
    ] },
    { name: 'Finance Agreements', fields: [
      { name: 'Sale', ...ref('Sales'), required: true }, { name: 'Lender', ...ref('Lenders'), required: true }, { name: 'Reference', type: 'text' },
      { name: 'Kind', type: 'text' }, { name: 'Amount Financed', type: 'currency' }, { name: 'Term Months', type: 'integer' }, { name: 'APR', type: 'number' },
      { name: 'Monthly Payment', type: 'currency' }, { name: 'Final Payment', type: 'currency', default: 0 }, { name: 'Commission', type: 'currency' },
      { name: 'Status', type: 'text' }, { name: 'Paid Out On', type: 'date' },
    ] },
    { name: 'Warranty Plans', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Months', type: 'integer' }, { name: 'Claim Limit', type: 'currency' },
      { name: 'Price', type: 'currency' }, { name: 'Cost', type: 'currency' }, { name: 'Administered By', type: 'text' }, { name: 'Cover', type: 'longtext' },
    ] },
    { name: 'Warranties', fields: [
      { name: 'Certificate No', type: 'text', unique: true }, { name: 'Sale', ...ref('Sales'), required: true }, { name: 'Plan', ...ref('Warranty Plans'), required: true },
      { name: 'Starts', type: 'date' }, { name: 'Ends', type: 'date' }, { name: 'Mileage at Start', type: 'integer' },
      { name: 'Price', type: 'currency' }, { name: 'Cost', type: 'currency' },
    ] },
    { name: 'Part Exchanges', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Customer', ...ref('Customers'), required: true }, { name: 'Valued By', ...ref('Staff') },
      { name: 'Valued On', type: 'date' }, { name: 'Registration', type: 'text' }, { name: 'Make', type: 'text' }, { name: 'Model', type: 'text' },
      { name: 'Year', type: 'integer' }, { name: 'Mileage', type: 'integer' }, { name: 'Colour', type: 'text' }, { name: 'Fuel', type: 'text' },
      { name: 'Condition', type: 'text' }, { name: 'Notes', type: 'text' }, { name: 'Guide Value', type: 'currency' }, { name: 'Offered', type: 'currency' },
      { name: 'Valid Until', type: 'date' }, { name: 'Outcome', type: 'text', default: 'Open' },
      { name: 'Sale', ...ref('Sales') }, { name: 'Taken In', ...ref('Vehicles') },
    ] },
  ],
  queries: {
    'Stock by days in stock': { tables: ['Vehicles'], columns: [
      { table: 'Vehicles', field: 'Stock No' }, { table: 'Vehicles', field: 'Registration' }, { expr: VEHICLE, as: 'Vehicle' },
      { table: 'Vehicles', field: 'Mileage' }, { table: 'Vehicles', field: 'Fuel' }, { table: 'Vehicles', field: 'Asking Price' },
      { table: 'Vehicles', field: 'Status', criteria: ['Not In ("Sold")'] },
      { expr: 'DateDiff("d", [Bought On], Date())', as: 'Days in Stock', sort: 'desc' },
    ] },
    'Stock over 60 days': { tables: ['Vehicles'], columns: [
      { table: 'Vehicles', field: 'Stock No' }, { table: 'Vehicles', field: 'Registration' }, { expr: VEHICLE, as: 'Vehicle' },
      { table: 'Vehicles', field: 'Bought On' }, { expr: 'DateDiff("d", [Bought On], Date())', as: 'Days in Stock', sort: 'desc', criteria: ['> 60'] },
      { expr: 'Round([Purchase Price] + Nz([Buying Fee]) + Nz([Prep Cost]), 2)', as: 'Cost to Date' },
      { table: 'Vehicles', field: 'First Asking Price' }, { table: 'Vehicles', field: 'Asking Price' },
      { expr: 'Round([Asking Price] - [Purchase Price] - Nz([Buying Fee]) - Nz([Prep Cost]), 2)', as: 'Margin at Asking' },
      { table: 'Vehicles', field: 'Status', criteria: ['In ("For sale", "In preparation")'] },
    ] },
    'Margin per vehicle': { tables: ['Sales', 'Vehicles', 'Staff'], columns: [
      { table: 'Sales', field: 'Handover', sort: 'desc' }, { table: 'Vehicles', field: 'Stock No' }, { table: 'Vehicles', field: 'Registration' },
      { expr: VEHICLE, as: 'Vehicle' }, { expr: SALESPERSON, as: 'Salesperson' }, { table: 'Sales', field: 'Price', as: 'Sold For' },
      { table: 'Vehicles', field: 'Purchase Price', as: 'Bought For' }, { table: 'Vehicles', field: 'Buying Fee' }, { table: 'Vehicles', field: 'Prep Cost' },
      { expr: MARGIN_VAT, as: 'VAT on Margin' }, { expr: MARGIN, as: 'Margin' }, { expr: 'DateDiff("d", [Bought On], [Handover])', as: 'Days in Stock' },
      { table: 'Sales', field: 'Status', criteria: ['Completed'], show: false },
    ] },
    'Sales by month': { tables: ['Sales', 'Vehicles'], columns: [
      { expr: 'Format([Handover], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Sales', field: 'ID', as: 'Cars Sold', total: 'count' },
      { table: 'Sales', field: 'Price', as: 'Vehicle Sales', total: 'sum' }, { table: 'Sales', field: 'Total', as: 'Takings', total: 'sum' },
      { expr: MARGIN, as: 'Gross Margin', total: 'sum' }, { table: 'Sales', field: 'Status', total: 'where', criteria: ['Completed'] },
    ] },
    'Sales by salesperson': { tables: ['Sales', 'Staff', 'Vehicles'], columns: [
      { expr: SALESPERSON, as: 'Salesperson', total: 'group' }, { table: 'Sales', field: 'ID', as: 'Cars Sold', total: 'count' },
      { table: 'Sales', field: 'Total', as: 'Takings', total: 'sum' }, { expr: MARGIN, as: 'Gross Margin', total: 'sum', sort: 'desc' },
      { table: 'Sales', field: 'Status', total: 'where', criteria: ['Completed'] },
      { expr: 'DateDiff("d", [Handover], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Deposits on unsold vehicles': { tables: ['Sales', 'Vehicles', 'Customers', 'Staff'], columns: [
      { table: 'Sales', field: 'Order No' }, { table: 'Sales', field: 'Agreed On' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' },
      { table: 'Vehicles', field: 'Registration' }, { expr: VEHICLE, as: 'Vehicle' }, { table: 'Sales', field: 'Deposit' },
      { table: 'Sales', field: 'Finance Amount' }, { table: 'Sales', field: 'Balance', as: 'Balance at Handover' },
      { table: 'Sales', field: 'Handover', sort: 'asc' }, { expr: SALESPERSON, as: 'Salesperson' },
      { table: 'Sales', field: 'Status', criteria: ['Reserved'], show: false },
    ] },
    'Test drives that did not buy': { tables: ['Test Drives', 'Customers', 'Vehicles', 'Staff'], columns: [
      { table: 'Test Drives', field: 'When', sort: 'desc' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' },
      { table: 'Customers', field: 'E-mail' }, { table: 'Vehicles', field: 'Registration' }, { expr: VEHICLE, as: 'Vehicle' },
      { expr: SALESPERSON, as: 'Salesperson' }, { table: 'Test Drives', field: 'Outcome', criteria: ['In ("Thinking it over", "Not for them", "Bought elsewhere")'] },
      { table: 'Test Drives', field: 'Notes' }, { expr: 'DateDiff("d", [When], Date())', as: 'Days Ago', criteria: ['Between 0 And 30'] },
    ] },
    'Test drives booked': { tables: ['Test Drives', 'Customers', 'Vehicles', 'Staff'], columns: [
      { table: 'Test Drives', field: 'When', sort: 'asc' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' },
      { table: 'Vehicles', field: 'Registration' }, { expr: VEHICLE, as: 'Vehicle' }, { table: 'Vehicles', field: 'Status', as: 'Vehicle Status' },
      { expr: SALESPERSON, as: 'With' }, { table: 'Test Drives', field: 'Outcome', criteria: ['Booked'], show: false },
      { expr: 'DateDiff("d", Date(), [When])', as: 'Days Ahead', criteria: ['Between 0 And 7'], show: false },
    ] },
    'Warranties expiring': { tables: ['Warranties', 'Warranty Plans', 'Sales', 'Customers', 'Vehicles'], columns: [
      { table: 'Warranties', field: 'Ends', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Ends])', as: 'Days Left', criteria: ['Between 0 And 60'] },
      { table: 'Warranties', field: 'Certificate No' }, { table: 'Warranty Plans', field: 'Name', as: 'Plan' }, { expr: CUSTOMER, as: 'Customer' },
      { table: 'Customers', field: 'Phone' }, { table: 'Vehicles', field: 'Registration' }, { expr: VEHICLE, as: 'Vehicle' },
      { table: 'Warranties', field: 'Mileage at Start' },
    ] },
    'Finance commission by month': { tables: ['Finance Agreements', 'Lenders', 'Sales'], columns: [
      { expr: 'Format([Paid Out On], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Lenders', field: 'Name', as: 'Lender', total: 'group', sort: 'asc' },
      { table: 'Finance Agreements', field: 'ID', as: 'Agreements', total: 'count' }, { table: 'Finance Agreements', field: 'Amount Financed', as: 'Amount Financed', total: 'sum' },
      { table: 'Finance Agreements', field: 'Commission', as: 'Commission', total: 'sum' },
      { table: 'Finance Agreements', field: 'Status', total: 'where', criteria: ['Paid out'] },
    ] },
    'Warranty sales by plan': { tables: ['Warranties', 'Warranty Plans'], columns: [
      { table: 'Warranty Plans', field: 'Name', as: 'Plan', total: 'group', sort: 'asc' }, { table: 'Warranties', field: 'ID', as: 'Sold', total: 'count' },
      { table: 'Warranties', field: 'Price', as: 'Takings', total: 'sum' }, { table: 'Warranties', field: 'Cost', as: 'Cost', total: 'sum' },
      { expr: 'Round([Warranties].[Price] - [Warranties].[Cost], 2)', as: 'Profit', total: 'sum' },
      { expr: 'DateDiff("d", [Starts], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Recent part-exchange valuations': { tables: ['Part Exchanges', 'Customers', 'Staff'], columns: [
      { table: 'Part Exchanges', field: 'Number' }, { table: 'Part Exchanges', field: 'Valued On', sort: 'desc' }, { expr: CUSTOMER, as: 'Customer' },
      { table: 'Part Exchanges', field: 'Registration' }, { expr: '[Part Exchanges].[Year] & " " & [Part Exchanges].[Make] & " " & [Part Exchanges].[Model]', as: 'Vehicle' },
      { table: 'Part Exchanges', field: 'Mileage' }, { table: 'Part Exchanges', field: 'Guide Value' }, { table: 'Part Exchanges', field: 'Offered' },
      { table: 'Part Exchanges', field: 'Valid Until' }, { table: 'Part Exchanges', field: 'Outcome' }, { expr: SALESPERSON, as: 'Valued By' },
      { expr: 'DateDiff("d", [Valued On], Date())', as: 'Days Ago', criteria: ['Between 0 And 30'], show: false },
    ] },
    'Purchases by source': { tables: ['Vehicles'], columns: [
      { table: 'Vehicles', field: 'Source', total: 'group', sort: 'asc' }, { table: 'Vehicles', field: 'ID', as: 'Bought', total: 'count' },
      { table: 'Vehicles', field: 'Purchase Price', as: 'Spent', total: 'sum' }, { table: 'Vehicles', field: 'Buying Fee', as: 'Fees', total: 'sum' },
      { table: 'Vehicles', field: 'Prep Cost', as: 'Preparation', total: 'sum' },
      { expr: 'DateDiff("d", [Bought On], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Preparation by vehicle': { tables: ['Preparation', 'Vehicles'], columns: [
      { table: 'Vehicles', field: 'Stock No', total: 'group', sort: 'asc' }, { table: 'Vehicles', field: 'Registration', total: 'group' },
      { expr: VEHICLE, as: 'Vehicle', total: 'group' }, { table: 'Vehicles', field: 'Status', total: 'group' },
      { table: 'Preparation', field: 'ID', as: 'Jobs', total: 'count' }, { table: 'Preparation', field: 'Cost', as: 'Cost', total: 'sum' },
      { table: 'Vehicles', field: 'Prep Cost', as: 'Prep Cost on Vehicle', total: 'max' },
      { table: 'Vehicles', field: 'Status', as: 'Still In Stock', total: 'where', criteria: ['Not In ("Sold")'] },
    ] },
    'Preparation spend by kind': { tables: ['Preparation'], columns: [
      { table: 'Preparation', field: 'Kind', total: 'group', sort: 'asc' }, { table: 'Preparation', field: 'ID', as: 'Jobs', total: 'count' },
      { table: 'Preparation', field: 'Cost', as: 'Cost', total: 'sum' },
      { expr: 'DateDiff("d", [Done On], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
  },
  reports: [
    { name: 'Stock list', query: 'Stock by days in stock', group: 'Status', columns: ['Stock No', 'Registration', 'Vehicle', { field: 'Mileage', type: 'integer' }, 'Fuel', { field: 'Asking Price', type: 'currency' }, { field: 'Days in Stock', type: 'integer' }], totals: [{ field: 'Asking Price', fn: 'sum' }, { field: 'Stock No', fn: 'count' }] },
    { name: 'Aged stock', query: 'Stock over 60 days', columns: ['Stock No', 'Registration', 'Vehicle', { field: 'Days in Stock', type: 'integer' }, { field: 'Cost to Date', type: 'currency' }, { field: 'Asking Price', type: 'currency' }, { field: 'Margin at Asking', type: 'currency' }], totals: [{ field: 'Cost to Date', fn: 'sum' }, { field: 'Margin at Asking', fn: 'sum' }] },
    { name: 'Vehicle margins', query: 'Margin per vehicle', columns: [{ field: 'Handover', type: 'date' }, 'Registration', 'Vehicle', 'Salesperson', { field: 'Sold For', type: 'currency' }, { field: 'Bought For', type: 'currency' }, { field: 'Prep Cost', type: 'currency' }, { field: 'VAT on Margin', type: 'currency' }, { field: 'Margin', type: 'currency' }], totals: [{ field: 'Margin', fn: 'sum' }, { field: 'Margin', fn: 'avg' }] },
    { name: 'Monthly sales', query: 'Sales by month', columns: ['Month', { field: 'Cars Sold', type: 'integer' }, { field: 'Vehicle Sales', type: 'currency' }, { field: 'Takings', type: 'currency' }, { field: 'Gross Margin', type: 'currency' }], totals: [{ field: 'Cars Sold', fn: 'sum' }, { field: 'Takings', fn: 'sum' }, { field: 'Gross Margin', fn: 'sum' }] },
    { name: 'Salesperson league', query: 'Sales by salesperson', columns: ['Salesperson', { field: 'Cars Sold', type: 'integer' }, { field: 'Takings', type: 'currency' }, { field: 'Gross Margin', type: 'currency' }], totals: [{ field: 'Cars Sold', fn: 'sum' }, { field: 'Gross Margin', fn: 'sum' }] },
    { name: 'Deposits held', query: 'Deposits on unsold vehicles', columns: ['Order No', 'Customer', 'Registration', 'Vehicle', { field: 'Handover', type: 'date' }, { field: 'Deposit', type: 'currency' }, { field: 'Balance at Handover', type: 'currency' }], totals: [{ field: 'Deposit', fn: 'sum' }] },
    { name: 'Test drive follow-ups', query: 'Test drives that did not buy', group: 'Salesperson', columns: ['When', 'Customer', 'Phone', 'Vehicle', 'Outcome'], totals: [{ field: 'Customer', fn: 'count' }] },
    { name: 'Warranty renewals', query: 'Warranties expiring', columns: [{ field: 'Ends', type: 'date' }, 'Certificate No', 'Plan', 'Customer', 'Phone', 'Registration', 'Vehicle'], totals: [{ field: 'Certificate No', fn: 'count' }] },
    { name: 'Finance commission', query: 'Finance commission by month', group: 'Month', columns: ['Lender', { field: 'Agreements', type: 'integer' }, { field: 'Amount Financed', type: 'currency' }, { field: 'Commission', type: 'currency' }], totals: [{ field: 'Commission', fn: 'sum' }, { field: 'Amount Financed', fn: 'sum' }] },
  ],
  documents: [
    { name: 'Vehicle sales invoice', table: 'Sales', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Sales invoice {Invoice No}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Customer', label: 'Sold to' }, { field: 'Invoice No', label: 'Invoice number' },
        { field: 'Customer.Address', label: 'Address' }, { field: 'Handover', label: 'Date of sale' },
        { field: 'Customer.City', label: 'Town' }, { field: 'Order No', label: 'Order' },
        { field: 'Customer.Postcode', label: 'Postcode' }, { field: 'Salesperson', label: 'Sold by' },
      ] },
      { type: 'fields', columns: 3, items: [
        { field: 'Vehicle.Registration', label: 'Registration' }, { field: 'Vehicle.Make', label: 'Make' }, { field: 'Vehicle.Model', label: 'Model' },
        { field: 'Vehicle.First Registered', label: 'First registered' }, { field: 'Vehicle.Mileage', label: 'Recorded mileage' }, { field: 'Vehicle.Colour', label: 'Colour' },
        { field: 'Vehicle.Fuel', label: 'Fuel' }, { field: 'Vehicle.Stock No', label: 'Stock number' }, { field: 'VAT Scheme', label: 'VAT' },
      ] },
      { type: 'lines', table: 'Sale Lines', sort: 'Line', columns: [
        { field: 'Description' }, { expr: 'Nz([Tax Rate]) * 100', as: 'VAT %', type: 'number' },
        { expr: VAT_IN, as: 'VAT Included', label: 'VAT included', type: 'currency' }, { field: 'Amount', type: 'currency' },
      ] },
      { type: 'totals', rows: [
        { label: 'Total price', expr: '[Amount]' }, { label: 'Of which VAT', expr: '[VAT]' },
        { label: 'Less part-exchange allowance', expr: '[Part-Ex Allowance]' }, { label: 'Less deposit paid', expr: '[Deposit]' },
        { label: 'Paid by finance', expr: '[Finance Amount]' }, { label: 'Balance paid at handover', expr: '[Total price] - [Less part-exchange allowance] - [Less deposit paid] - [Paid by finance]', strong: true },
      ] },
      { type: 'text', text: 'VAT: {VAT Scheme}. A used car bought from a private seller or under the margin scheme is sold under the margin scheme: the VAT is within the price, paid by us on our margin, and is not shown or reclaimable. On a VAT-qualifying car the VAT within the price is shown above and a VAT-registered buyer may reclaim it. Warranties are exempt from VAT.', small: true },
      { type: 'text', text: '{Business.Payment Terms} Bank: {Business.Bank Details}.', small: true },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    { name: 'Sales agreement', table: 'Sales', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Vehicle order and sales agreement {Order No}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Customer', label: 'Buyer' }, { field: 'Order No', label: 'Order number' },
        { field: 'Customer.Address', label: 'Address' }, { field: 'Agreed On', label: 'Date of order' },
        { field: 'Customer.City', label: 'Town' }, { field: 'Handover', label: 'Handover on or about' },
        { field: 'Customer.Postcode', label: 'Postcode' }, { field: 'Customer.Phone', label: 'Phone' },
      ] },
      { type: 'fields', columns: 3, items: [
        { field: 'Vehicle.Registration', label: 'Registration' }, { field: 'Vehicle.Make', label: 'Make' }, { field: 'Vehicle.Model', label: 'Model' },
        { field: 'Vehicle.Year', label: 'Year' }, { field: 'Vehicle.Mileage', label: 'Recorded mileage' }, { field: 'Vehicle.Colour', label: 'Colour' },
        { field: 'Vehicle.MOT Expiry', label: 'MOT until' }, { field: 'Vehicle.Keys', label: 'Keys' }, { field: 'Vehicle.Service History', label: 'Service history' },
      ] },
      { type: 'lines', table: 'Sale Lines', sort: 'Line', columns: [{ field: 'Description', label: 'What is bought' }, { field: 'Amount', type: 'currency' }] },
      { type: 'totals', rows: [
        { label: 'Total price', expr: '[Amount]', strong: true }, { label: 'Part-exchange allowance', expr: '[Part-Ex Allowance]' },
        { label: 'Deposit paid today', expr: '[Deposit]' }, { label: 'To be paid by finance', expr: '[Finance Amount]' },
        { label: 'Balance due at handover', expr: '[Total price] - [Part-exchange allowance] - [Deposit paid today] - [To be paid by finance]', strong: true },
      ] },
      { type: 'text', text: 'This agreement is made on {Agreed On} between {Business.Name} of {Business.Address}, {Business.City} {Business.Postcode} ("the Seller") and {Customer} of {Customer.Address}, {Customer.City} {Customer.Postcode} ("the Buyer").' },
      { type: 'text', text: '1. The vehicle. The Seller agrees to sell and the Buyer agrees to buy the vehicle described above, registration {Vehicle.Registration}, with the extras listed, for the total price above. The mileage shown is the reading on the day of this order; the Seller is not aware of anything suggesting it is wrong.' },
      { type: 'text', text: '2. Deposit. The Buyer has paid a deposit of {Deposit} today, and the Seller holds the vehicle for the Buyer until the handover date. The deposit is part of the price.' },
      { type: 'text', text: '3. Payment. The balance due at handover is paid by bank transfer or debit card before the vehicle is driven away. The vehicle remains the Seller\'s property until the whole price has been paid.' },
      { type: 'text', text: '4. Part-exchange. Where a vehicle is taken in part-exchange, the allowance above is given on the understanding that it is the Buyer\'s to sell, that it is handed over in the condition and with no more than 500 miles more than when it was valued, with its keys and registration document, and that any finance owing on it has been declared. The Seller settles declared finance with the lender and takes the amount settled from the allowance.' },
      { type: 'text', text: '5. Finance. Where part of the price is paid by finance, the finance agreement is between the Buyer and the lender, not the Seller. This sale depends on the lender accepting the Buyer; if it does not, either party may cancel this agreement and the deposit is returned in full.' },
      { type: 'text', text: '6. Cancelling. If the Buyer cancels for any other reason before handover, the Seller may keep from the deposit the reasonable costs it has already met for this sale, such as work done to the vehicle at the Buyer\'s request, and returns the rest within 14 days.' },
      { type: 'text', text: '7. Handover. The vehicle will be ready on or about {Handover}, valeted, with its keys, service records and registration document, and an MOT certificate running until {Vehicle.MOT Expiry}.' },
      { type: 'text', text: '8. Warranty. The warranty shown above starts on the day of handover. What it covers is set out on its certificate, given at handover.' },
      { type: 'text', text: '9. The Buyer\'s rights. Nothing in this agreement takes away any right the Buyer has in law when buying goods as a consumer, including the right to reject a vehicle that is faulty.' },
      { type: 'text', text: '10. The whole agreement. This is the whole of what has been agreed. Any change must be written down and signed by both parties.' },
      { type: 'spacer', size: 24 },
      { type: 'signatures', labels: ['For the Seller', 'The Buyer'] },
    ] },
    { name: 'Part-exchange valuation', table: 'Part Exchanges', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'fields', columns: 2, items: [
        { field: 'Customer', label: 'To' }, { field: 'Number', label: 'Our reference' },
        { field: 'Customer.Address', label: 'Address' }, { field: 'Valued On', label: 'Date' },
        { field: 'Customer.City', label: 'Town' }, { field: 'Customer.Postcode', label: 'Postcode' },
      ] },
      { type: 'heading', text: 'Your part-exchange valuation' },
      { type: 'text', text: 'Dear {Customer.First Name},' },
      { type: 'text', text: 'Thank you for letting us look at your {Year} {Make} {Model}, registration {Registration}, when you visited us on {Valued On}. We have checked it over and driven it, and taken account of its {Mileage} recorded miles and its {Condition} condition. What we noted: {Notes}' },
      { type: 'text', text: 'We can offer you {Offered} for it as a part-exchange against any car from our stock.' },
      { type: 'text', text: 'This offer stands until {Valid Until}, as long as the car is in the same condition when you bring it in, with no more than 500 miles more on it, both keys if it came with two, its registration document and its service records. If there is finance owing on the car, please tell us how much: we settle it with your lender and take what we pay from the allowance.' },
      { type: 'text', text: 'If you have any questions, or would like to talk about a car from our stock, call us on {Business.Phone} or reply to {Business.E-mail}.' },
      { type: 'text', text: 'Yours sincerely,' },
      { type: 'spacer', size: 24 },
      { type: 'text', text: '{Valued By}, {Business.Name}' },
    ] },
    { name: 'Warranty certificate', table: 'Warranties', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Warranty certificate {Certificate No}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Sale.Customer', label: 'Warranty holder' }, { field: 'Certificate No', label: 'Certificate' },
        { field: 'Sale.Customer.Address', label: 'Address' }, { field: 'Plan.Name', label: 'Plan' },
        { field: 'Sale.Customer.Postcode', label: 'Postcode' }, { field: 'Starts', label: 'Cover starts' },
        { field: 'Sale.Vehicle.Registration', label: 'Registration' }, { field: 'Ends', label: 'Cover ends' },
        { field: 'Sale.Vehicle.Make', label: 'Make' }, { field: 'Mileage at Start', label: 'Mileage at start' },
        { field: 'Sale.Vehicle.Model', label: 'Model' }, { field: 'Plan.Claim Limit', label: 'Most paid on one claim' },
      ] },
      { type: 'text', text: 'What is covered. {Plan.Cover}' },
      { type: 'text', text: 'How to claim. Before any repair is started, call {Business.Name} on {Business.Phone} with this certificate\'s number. We will tell you where to take the car and agree the work with the garage. Repairs started without our agreement cannot be paid for.' },
      { type: 'text', text: 'Conditions. The car must be serviced as its maker recommends, at any garage, and the receipts kept. Wear and tear, damage from accidents, misuse or neglect, and parts that wear in normal use (tyres, brake pads and discs, clutches, wiper blades, bulbs) are not covered. The warranty stays with the car if it is sold, if you tell us. This warranty is in addition to your rights in law.' },
      { type: 'text', text: 'Administered by {Plan.Administered By}.', small: true },
      { type: 'spacer', size: 16 },
      { type: 'signatures', labels: ['For the Seller'] },
    ] },
    { name: 'Window card', table: 'Vehicles', paper: 'a5', letterhead: { table: 'Business', fields: ['Name', 'Phone', 'Website'] }, blocks: [
      { type: 'heading', text: '{Year} {Make} {Model}' },
      { type: 'fields', columns: 3, items: [
        { field: 'Mileage', label: 'Miles' }, { field: 'Fuel' }, { field: 'Transmission', label: 'Gearbox' },
        { field: 'Colour' }, { field: 'Body' }, { field: 'Doors' },
        { field: 'Owners', label: 'Owners' }, { field: 'Service History', label: 'Service history' }, { field: 'MOT Expiry', label: 'MOT until' },
      ] },
      { type: 'lines', table: 'Vehicle Features', columns: [{ field: 'Feature', label: 'Key features' }] },
      { type: 'totals', rows: [{ label: 'Our price', expr: '[Asking Price]', strong: true }] },
      { type: 'text', text: 'Part-exchange welcome. Finance arranged. At least three months\' warranty on every car. {Registration}, stock number {Stock No}.', small: true },
          ] },
  ],
};

export const about = {
  name: 'Copperbeech Motors',
  kind: 'Used car dealer',
  line: 'Stock, preparation costs, test drives, sales, finance, warranties and part-exchanges for a used car dealer.',
  summary: [
    'Copperbeech Motors is an invented independent used car dealer in south-east London, with about forty cars on the forecourt, a sales manager who owns the business, three sales executives, a workshop technician, a valeter and an administrator. It buys most of its cars at two auctions twice a week, some from the trade, and takes about a third of its customers\' old cars in part-exchange.',
    'The sample holds fifteen months of trading, made up a day at a time. Each car is prepared before it goes on sale: always a valet, often a service, repairs where the inspection finds them, and an MOT when the old one is close to running out. Its asking price is set from what it cost, and comes down when it has stood for a month. Customers test drive; some buy, pay a deposit, and pay the rest at handover a few days later, nearly half with finance from one of two lenders, who pay the dealer a commission. Every car leaves with a warranty: three months included, or a longer plan bought. The part-exchange comes into stock the day the new car leaves.',
    'Every car and make in it is invented. The makes and models are made-up names, and every registration begins ZZ, which is never issued, so none is a real car\'s.',
  ],
  suits: ['A used car or van dealer', 'A motorcycle dealer', 'A car supermarket or a dealer\'s used car site', 'A specialist or classic car dealer', 'A caravan or motorhome dealer', 'A plant or machinery dealer selling used stock'],
  tables: {
    Business: BUSINESS_ABOUT,
    Staff: { about: 'The people who work at the dealer: sales, workshop, valeting and the office.', fields: {
      'First Name': 'Their first name.', 'Last Name': 'Their family name.', Role: 'Their job, such as Sales Executive or Workshop Technician.',
      Phone: 'Their work phone.', 'E-mail': 'Their work e-mail address.', Started: 'The day they joined.',
      Sells: 'Ticked for those who sell cars and take test drives, so they appear in the sales queries.',
    } },
    Suppliers: { about: 'Where cars are bought: the auctions and the trade sellers. Part-exchanges come from customers and have no supplier.', fields: {
      Name: 'The auction or trade seller\'s name.', Kind: 'Auction or Trade.', Address: 'The street address.', City: 'The town or city.',
      Postcode: 'The postcode.', Phone: 'Their phone number.', 'E-mail': 'Where invoices and queries go.', 'Account No': 'Your account or buyer number with them.',
    } },
    Customers: { about: 'Everyone who has test driven, bought or had a car valued, one record each.', fields: {
      'First Name': 'Their first name.', 'Last Name': 'Their family name.', Address: 'The street address, printed on the invoice and agreement.',
      City: 'The town.', Postcode: 'The postcode.', Phone: 'The number to call them on.', 'E-mail': 'Their e-mail address.',
      'Heard Of Us': 'How they found you: Website, Walk-in, Recommended or Returning customer.', Added: 'The day they were first entered.',
      'Marketing Consent': 'Ticked when they have agreed to hear about offers and warranty renewals.',
    } },
    Vehicles: { about: 'Every car the dealer has bought, one record each, in stock or sold.', fields: {
      Registration: 'The number plate. Each is used once. The sample\'s begin ZZ, which is never issued.',
      'Stock No': 'Your own stock number, given in the order cars are bought: STK-1042.', Make: 'The maker. The sample\'s are invented.',
      Model: 'The model, engine and trim, as it is advertised.', Body: 'Hatchback, Estate, Saloon, SUV or MPV.', Year: 'The year it was first registered.',
      'First Registered': 'The day it was first registered, from the registration document.', Colour: 'Its colour.', Fuel: 'Petrol, Diesel, Hybrid or Electric.',
      Transmission: 'Manual or Automatic.', Doors: 'How many doors.', Mileage: 'The mileage when it came in; for a sold car, when it left.',
      Owners: 'How many keepers it has had before you.', 'Service History': 'Full, Part or None.', 'MOT Expiry': 'When its MOT certificate runs out.',
      Keys: 'How many keys came with it.', Source: 'Auction, Trade or Part-exchange.', Supplier: 'The auction or trade seller it was bought from; empty for a part-exchange.',
      'Bought On': 'The day it was bought or taken in. Days in stock are counted from here.', 'Purchase Price': 'What was paid for it, or the allowance given for a part-exchange.',
      'Buying Fee': 'The auction\'s buyer\'s fee, nought for other cars.', 'Prep Cost': 'What preparing it has cost: its jobs in Preparation added up. Add each job\'s cost here when you enter it; the query Preparation by vehicle checks the two agree.',
      'First Asking Price': 'The price it first went on sale at.', 'Asking Price': 'The price on its window card now, lower than the first when it has been reduced.',
      'VAT Qualifying': 'Ticked for a car bought with VAT shown (usually from a business or a lease company), sold with VAT shown. Others are sold under the margin scheme.',
      Status: 'In preparation, For sale, Reserved (a deposit taken) or Sold.', Notes: 'Anything worth remembering about it.',
    } },
    'Vehicle Features': { about: 'The features listed on a car\'s window card and advert, one record each.', fields: { Vehicle: 'The car.', Feature: 'One feature, such as Rear parking sensors.' } },
    Preparation: { about: 'Every job done to get a car ready for sale, with what it cost.', fields: {
      Vehicle: 'The car the work was done on.', 'Done On': 'The day it was done.', Kind: 'Valet, Service, MOT, Repair, Tyres or Bodywork.',
      Description: 'What was done.', 'Done By': 'In house, or the garage that did it.', Cost: 'What it cost, before any VAT you reclaim.',
    } },
    'Test Drives': { about: 'Every test drive, taken or booked, and how it ended.', fields: {
      Vehicle: 'The car driven.', Customer: 'Who drove it.', Staff: 'Who went out with them.', When: 'The day and time.',
      Outcome: 'Booked (not taken yet), Bought, Thinking it over, Not for them or Bought elsewhere.', Notes: 'What the customer said, for the follow-up call.',
    } },
    Sales: { about: 'Every car sold or reserved: the order, and the invoice when it is handed over.', fields: {
      'Order No': 'The order\'s number, given when the deal is agreed: ORD-2026-0042.', 'Invoice No': 'The invoice\'s number, given at handover in date order: INV-2026-0042. Empty while the car is reserved.',
      Vehicle: 'The car sold.', Customer: 'Who bought it.', Salesperson: 'Who sold it.', 'Agreed On': 'The day the deal was agreed and the deposit taken.',
      Handover: 'The day the car is collected, or is to be.', Price: 'The agreed price of the car itself, after any discount.',
      Total: 'Everything on the sale: the car, the warranty and any extras, VAT included. The sale\'s lines add up to it.',
      VAT: 'The VAT within the total that the invoice shows: on the extras, and on the car itself only when it is VAT qualifying.',
      'VAT Scheme': 'Margin scheme or VAT qualifying, printed on the invoice.', 'Part-Ex Allowance': 'What is allowed for the customer\'s car taken in part-exchange.',
      Deposit: 'The deposit paid when the deal was agreed.', 'Finance Amount': 'What the lender pays, when part is paid by finance.',
      Balance: 'What the customer pays at handover: the total less the part-exchange, the deposit and the finance.',
      Status: 'Reserved (agreed, not handed over), Completed or Cancelled.', Notes: 'Anything to remember, such as why a sale was cancelled and when the deposit went back.',
    } },
    'Sale Lines': { about: 'What is on each sale: the car, the warranty and any extras, each with its price.', fields: {
      Sale: 'The sale the line belongs to.', Line: 'The order the lines print in: the car first.', Kind: 'Vehicle, Warranty or Extra.',
      Description: 'The words printed on the invoice.', Amount: 'Its price, VAT included.',
      'Tax Rate': 'The VAT within it, as a fraction: 0.2 for extras and a VAT-qualifying car, 0 for a margin-scheme car and for a warranty, which is exempt.',
    } },
    Lenders: { about: 'The finance companies the dealer introduces customers to. The sample\'s are invented.', fields: {
      Name: 'The lender\'s name.', Address: 'Their address.', Phone: 'The dealer support line.', 'E-mail': 'Where proposals go.',
      'Commission Rate': 'What they pay the dealer, as a fraction of the amount financed: 0.04 is four percent.', 'Typical APR': 'The rate they usually lend at, in percent.',
    } },
    'Finance Agreements': { about: 'Each sale paid partly by finance: the lender, the terms and the commission earned.', fields: {
      Sale: 'The sale it pays for.', Lender: 'The lender.', Reference: 'The lender\'s reference for the proposal.', Kind: 'Hire purchase, or Personal contract purchase with a final payment.',
      'Amount Financed': 'What the lender pays the dealer.', 'Term Months': 'How many monthly payments.', APR: 'The rate, in percent.',
      'Monthly Payment': 'What the customer pays each month.', 'Final Payment': 'The optional last payment of a personal contract purchase; nought for hire purchase.',
      Commission: 'What the lender pays the dealer for the introduction.', Status: 'Approved (car not handed over yet), Paid out or Cancelled.',
      'Paid Out On': 'The day the lender paid the dealer.',
    } },
    'Warranty Plans': { about: 'The warranties offered: three months included with every car, and longer plans to buy.', fields: {
      Name: 'The plan\'s name. Each is used once.', Months: 'How long it lasts.', 'Claim Limit': 'The most paid on any one claim.',
      Price: 'What the customer pays; nought for the one included.', Cost: 'What the dealer pays the administrator for it.',
      'Administered By': 'Who handles claims. The sample\'s administrator is invented.', Cover: 'What it covers, in words, printed on the certificate.',
    } },
    Warranties: { about: 'The warranty each car left with, and when it runs out.', fields: {
      'Certificate No': 'The certificate\'s number, given at handover: WAR-2026-0042.', Sale: 'The sale it came with.', Plan: 'The plan.',
      Starts: 'The day it starts, the handover.', Ends: 'The day it runs out.', 'Mileage at Start': 'The mileage at handover.',
      Price: 'What the customer paid for it.', Cost: 'What it cost the dealer.',
    } },
    'Part Exchanges': { about: 'Every valuation of a customer\'s car, taken in or not.', fields: {
      Number: 'The valuation\'s number, printed on the letter: PX-2026-0042.', Customer: 'Whose car it is.', 'Valued By': 'Who valued it.',
      'Valued On': 'The day it was looked at.', Registration: 'Its number plate.', Make: 'Its maker.', Model: 'Its model.', Year: 'The year it was first registered.',
      Mileage: 'The mileage when valued.', Colour: 'Its colour.', Fuel: 'Its fuel.', Condition: 'Good, Fair or Poor, as it was found.',
      Notes: 'What was noted on the inspection, printed in the letter.', 'Guide Value': 'What the trade price guide gives for it.',
      Offered: 'The allowance offered.', 'Valid Until': 'The last day the offer stands.', Outcome: 'Open, Accepted, Declined, Expired or Withdrawn.',
      Sale: 'The sale it was part-exchanged against, when it was.', 'Taken In': 'Its record in Vehicles once it has come into stock.',
    } },
  },
  queries: {
    'Stock by days in stock': 'Every car not yet sold, the longest in stock first, with its asking price and status. The forecourt list.',
    'Stock over 60 days': 'Cars on sale or in preparation for more than sixty days, with what each has cost so far and the margin left at its asking price: the cars to reduce, advertise harder or send back to auction.',
    'Margin per vehicle': 'Each car sold, the newest first: what it sold for, what it cost, the buying fee and preparation, the VAT on the margin and what is left, and how long it took to sell.',
    'Sales by month': 'Cars handed over each month, with the takings and the gross margin on the cars.',
    'Sales by salesperson': 'Cars sold, takings and gross margin for each salesperson over the last twelve months, the best first.',
    'Deposits on unsold vehicles': 'Reserved cars: the deposit held, the finance and the balance still to come, by handover date. Use it each morning to get the cars ready.',
    'Test drives that did not buy': 'Test drives in the last thirty days that did not end in a sale, the newest first, with the customer\'s phone and what they said: the follow-up calls.',
    'Test drives booked': 'Test drives booked for today and the next seven days, in time order.',
    'Warranties expiring': 'Warranties that run out in the next sixty days, the soonest first, with the customer\'s phone: who to offer a renewal or a service.',
    'Finance commission by month': 'Finance agreements paid out each month by lender, with the amount financed and the commission earned.',
    'Warranty sales by plan': 'Warranties sold in the last twelve months by plan, with what they brought in, what they cost and the profit.',
    'Recent part-exchange valuations': 'Valuations in the last thirty days, the newest first, with what was offered and how each ended.',
    'Purchases by source': 'Cars bought in the last twelve months from auctions, the trade and part-exchange: how many, what they cost, the fees and the preparation.',
    'Preparation by vehicle': 'For each car still in stock, the preparation jobs and their cost beside the Prep Cost on its record, to check the two agree.',
    'Preparation spend by kind': 'What was spent on valets, services, MOTs, repairs, tyres and bodywork in the last twelve months.',
  },
  reports: {
    'Stock list': 'The forecourt list by status, with mileage, asking price and days in stock, and the stock\'s value at asking prices.',
    'Aged stock': 'The cars over sixty days, with what they have cost and the margin left, for the weekly stock meeting.',
    'Vehicle margins': 'Every car sold with its costs and margin, and the total and average margin.',
    'Monthly sales': 'One line a month: cars sold, takings and gross margin, with the totals.',
    'Salesperson league': 'The sales team over the last twelve months, by gross margin.',
    'Deposits held': 'Every deposit held on a reserved car, with the handover date and the total held.',
    'Test drive follow-ups': 'The calls to make, on a page for each salesperson.',
    'Warranty renewals': 'Warranties running out in the next sixty days, for the renewal letters and calls.',
    'Finance commission': 'Commission earned from each lender, month by month.',
  },
  documents: {
    'Vehicle sales invoice': 'The invoice given at handover: the car with its registration, mileage and first registration date, the warranty and extras, the total, the part-exchange allowance, the deposit and finance taken off, and the balance paid. It says whether the car is sold under the margin scheme or is VAT qualifying. A reserved car prints without an invoice number until it is handed over.',
    'Sales agreement': 'The order the customer signs when the deal is agreed and the deposit paid: the parties, the car, what is bought, the deposit, finance and balance, and plainly worded terms on payment, part-exchange, finance, cancelling, handover and warranty. Have it checked against the law of your own country before you use it.',
    'Part-exchange valuation': 'The letter sent after valuing a customer\'s car: the car, what was noted, the allowance offered, and how long the offer stands.',
    'Warranty certificate': 'Given at handover with every car: the holder, the car, the plan, when it starts and ends, the claim limit, what is covered and how to claim.',
    'Window card': 'The card in the car\'s window on the forecourt: the year, make and model, its mileage and details, its key features and the price. Print it again after a price change.',
  },
  tutorial: [
    { title: 'Buy a car and put it on sale', steps: [
      'Open the Vehicles form and press New. Type the Registration, the next Stock No, the Make, Model, Year, Mileage and the other details from the registration document.',
      'Choose Auction or Trade in Source and the seller in Supplier. Type the Bought On date, the Purchase Price and, from an auction, the Buying Fee. Leave Status as In preparation.',
      'Add each preparation job in the Preparation form as it is done (choose the car in Vehicle, then the Kind, Description, Done By and Cost), and add its cost to the car\'s Prep Cost.',
      'Add the car\'s features in the Vehicle Features form, one record each.',
      'When it is ready, type its First Asking Price and Asking Price and set Status to For sale. Open the print template Window card, choose the car and press Print.',
    ] },
    { title: 'Book a test drive and follow it up', steps: [
      'Find the customer in the Customers form, or press New and type their name, phone and e-mail.',
      'Open the Test Drives form and press New. Choose the car in Vehicle, the customer in Customer and yourself in Staff, and type the day and time in When. Outcome starts as Booked.',
      'After the drive, set Outcome to Bought, Thinking it over, Not for them or Bought elsewhere, and type what they said in Notes.',
      'Each morning, run the query Test drives that did not buy, or print the report Test drive follow-ups, and call the people on it.',
    ] },
    { title: 'Value a part-exchange', steps: [
      'Open the Part Exchanges form and press New. Type the next Number, choose the customer and yourself in Valued By, and type the car\'s details, its Condition and what you noted.',
      'Type the Guide Value from your price guide, the allowance you offer in Offered, and Valid Until (usually a week). Outcome starts as Open.',
      'Open the print template Part-exchange valuation, choose the valuation and press Print or Save as PDF to send the letter.',
    ] },
    { title: 'Take a deposit and agree a sale', steps: [
      'Open the Sales form and press New. Type the next Order No, choose the car in Vehicle, the customer and the Salesperson, and type Agreed On and the Handover date.',
      'In the sale\'s lines below, add the car (Kind Vehicle, the description and the Amount), the warranty and any extras, with each line\'s Tax Rate. Type the Price, the Total of the lines and the VAT they hold.',
      'Type the Part-Ex Allowance, the Deposit taken and any Finance Amount, and the Balance left to pay. Leave Status as Reserved and set the car\'s Status to Reserved.',
      'Where finance is used, open the Finance Agreements form and press New: choose the sale and the lender and type the terms and the commission. Set its Status to Approved.',
      'Open the print template Sales agreement, choose the sale, print two copies and sign them with the customer.',
    ] },
    { title: 'Hand a car over', steps: [
      'Run the query Deposits on unsold vehicles to see which cars go out today and what each customer still owes.',
      'On the sale in the Sales form, type the next Invoice No, check the Handover date and set Status to Completed. Set the car\'s Status to Sold.',
      'Open the Warranties form and press New. Type the next Certificate No, choose the sale and the plan, and type Starts, Ends and the Mileage at Start.',
      'Open the print templates Vehicle sales invoice and Warranty certificate in turn, choose the sale or the warranty, and press Print.',
      'If a car came in part-exchange, add it to the Vehicles form with Source Part-exchange and the allowance as its Purchase Price, then set the valuation\'s Outcome to Accepted and choose the new car in Taken In.',
      'When the lender pays, set the finance agreement\'s Status to Paid out and type Paid Out On.',
    ] },
    { title: 'Review the stock each week', steps: [
      'Print the report Aged stock. For each car, decide whether to reduce it, advertise it or send it back to auction.',
      'To reduce a car, type its new Asking Price in the Vehicles form (keep First Asking Price as it was) and print its Window card again.',
      'Run the query Preparation by vehicle to check that each car\'s Prep Cost agrees with its jobs.',
    ] },
    { title: 'See how the month went', steps: [
      'Print the report Monthly sales for cars sold, takings and gross margin by month.',
      'Print the report Salesperson league for the team, and Finance commission for what the lenders paid.',
      'Run the query Warranty sales by plan for the warranty profit, and print Warranty renewals to start the renewal calls.',
    ] },
  ],
  notes: [
    'The margin in the queries is the car\'s own: what it sold for, less what it cost, the buying fee, its preparation and the VAT on the margin. Warranty profit and finance commission are counted in their own queries, and overheads (wages, rent, advertising) are not in this database.',
    'VAT on used cars: most are sold under the margin scheme, where VAT is a sixth of the margin, paid by the dealer and not shown on the invoice. A VAT-qualifying car shows its VAT on the invoice, but because the dealer reclaims the VAT it paid when buying, what it owes comes to the same sixth of the margin, so the margin query treats both alike. Check the rules for your own country with your accountant.',
    'Prep Cost on a car is kept by hand: add each job\'s cost when you enter it in Preparation. The query Preparation by vehicle puts the two side by side.',
    'Every sale is printed by the print templates, whatever its Status: a reserved sale\'s invoice has no number yet, and a cancelled sale still has its agreement. Choose the record to print, or print the one on screen.',
    'The Window card prints every car on file, sold ones too. Choose the cars on sale (the query Stock by days in stock lists them) and print those.',
    'A part-exchange comes into stock as a new record in Vehicles, with its allowance as the purchase price and no supplier. Its valuation points at the sale it was taken against and at the new stock record.',
    'Numbers (orders, invoices, certificates, valuations, stock numbers) are typed in date order, a new run each year: look at the last one in the form and add one.',
  ],
  faq: [
    { q: 'Why is there no VAT on most of my invoices?', a: 'Most used cars are sold under the margin scheme: VAT is paid on your margin, not on the price, and the invoice must not show it. The invoice says so under the lines. A VAT-qualifying car, ticked in Vehicles, shows the VAT within its price.' },
    { q: 'How do I see which cars have been in stock too long?', a: 'Run the query Stock over 60 days, or print the report Aged stock. Change the criteria under Days in Stock (> 60) in the query designer for another number of days.' },
    { q: 'How do I record a sale that falls through?', a: 'Set the sale\'s Status to Cancelled, write in Notes when and how the deposit was returned, set the car back to For sale, and set any finance agreement and part-exchange valuation to Cancelled and Withdrawn.' },
    { q: 'Where does the finance commission go?', a: 'Each finance agreement holds its commission. The query Finance commission by month and the report Finance commission add it up by month and lender, counting agreements that have been paid out.' },
    { q: 'How do I print one invoice?', a: 'Open the Sales form on the sale and press Print, choosing Vehicle sales invoice; or open the print template from Print templates and choose the sale.' },
    { q: 'Can I change the warranty plans?', a: 'Yes. Edit the Warranty Plans records: the name, months, claim limit, price, cost and the Cover words that print on the certificate. Old warranties keep the price they were sold at.' },
    { q: 'How do I find who to call about a warranty running out?', a: 'Run the query Warranties expiring, or print the report Warranty renewals. Only call customers who have ticked Marketing Consent about offers.' },
  ],
  empty: 'Start with Business, your own name, address and bank details, which every document prints. Then Staff (tick Sells for the sales team), Suppliers, Lenders and Warranty Plans. Then add the cars you hold in Vehicles, with their features and preparation. Customers, test drives, valuations, sales, finance agreements and warranties come in as you trade.',
  adapt: [
    'Sell vans or motorcycles too: add their bodies to Body, and a Payload or Engine Size field to Vehicles.',
    'Keep a price history: add a Price Changes table (vehicle, date, old price, new price) and a query of the cars reduced this month.',
    'Track where enquiries come from: add a Leads table (customer, source, vehicle asked about, date) and a query of leads by source and how many bought.',
    'Send cars back to auction: add a Disposals table (vehicle, date, auction, price) and a Status of Sent to auction, so slow stock leaves with its loss recorded.',
    'Record warranty claims: add a Claims table pointing at Warranties, with the fault, the garage and what was paid, and a query of claims against each plan.',
    'Hold overheads: add an Expenses table (date, kind, amount) and a monthly query beside Sales by month to see the profit after costs.',
    'Pay commission to the sales team: add a Commission Rate to Staff and a calculated column in Sales by salesperson.',
  ],
};

// The invented makes and models: [make, model, body, new price, fuels [[fuel, weight]...], engines, weight].
const RANGE = [
  ['Velmar', 'Aspa', 'Hatchback', 15500, [['Petrol', 1]], ['1.0', '1.2'], 5],
  ['Velmar', 'Corla', 'Hatchback', 22500, [['Petrol', 5], ['Diesel', 3], ['Hybrid', 2]], ['1.5', '1.6'], 6],
  ['Velmar', 'Tavio', 'SUV', 29500, [['Petrol', 4], ['Diesel', 3], ['Hybrid', 3]], ['1.5', '2.0'], 4],
  ['Torvik', 'Fenna', 'Estate', 26500, [['Petrol', 3], ['Diesel', 5]], ['1.6', '2.0'], 3],
  ['Torvik', 'Tarn', 'SUV', 33500, [['Diesel', 5], ['Hybrid', 3]], ['2.0'], 3],
  ['Torvik', 'Sund', 'Saloon', 30500, [['Petrol', 3], ['Diesel', 4]], ['1.8', '2.0'], 2],
  ['Kordal', 'Brin', 'Hatchback', 13500, [['Petrol', 1]], ['1.0'], 4],
  ['Kordal', 'Vixa', 'Hatchback', 18500, [['Petrol', 4], ['Diesel', 2]], ['1.2', '1.4'], 5],
  ['Kordal', 'Ostra', 'MPV', 25500, [['Petrol', 2], ['Diesel', 3]], ['1.6', '2.0'], 2],
  ['Selvane', 'Orra', 'SUV', 27500, [['Petrol', 3], ['Hybrid', 4]], ['1.4', '1.6'], 4],
  ['Selvane', 'Ivel', 'Hatchback', 31000, [['Electric', 1]], [], 2],
  ['Brevan', 'Kesta', 'Estate', 24500, [['Petrol', 2], ['Diesel', 4]], ['1.5', '2.0'], 2],
  ['Brevan', 'Marn', 'Saloon', 27000, [['Petrol', 3], ['Hybrid', 2]], ['1.6', '2.0'], 2],
  ['Duvaro', 'Pesk', 'SUV', 21000, [['Petrol', 4], ['Hybrid', 2]], ['1.0', '1.3'], 4],
  ['Duvaro', 'Lenta', 'Hatchback', 17000, [['Petrol', 1]], ['1.2'], 3],
  ['Orrin', 'Kaldo', 'SUV', 41000, [['Diesel', 4], ['Hybrid', 3], ['Electric', 1]], ['2.0', '2.2'], 2],
  ['Orrin', 'Halm', 'Saloon', 36000, [['Petrol', 2], ['Diesel', 3]], ['2.0'], 1],
  ['Ameno', 'Pira', 'Hatchback', 24000, [['Hybrid', 1]], ['1.8'], 3],
  ['Ameno', 'Soli', 'SUV', 38500, [['Electric', 1]], [], 2],
];
// Trims, dearer up the list, each adding features to those below it.
const TRIMS = [
  { name: 'Comfort', uplift: 1, weight: 4, features: ['Air conditioning', 'Bluetooth', 'DAB radio', 'Electric windows', 'Cruise control'] },
  { name: 'Design', uplift: 1.08, weight: 5, features: ['Alloy wheels', 'Rear parking sensors', 'Touchscreen with phone mirroring', 'Climate control'] },
  { name: 'Sport', uplift: 1.13, weight: 2, features: ['Sports seats', 'Tinted rear glass', 'LED headlights'] },
  { name: 'Premium', uplift: 1.2, weight: 2, features: ['Heated front seats', 'Reversing camera', 'Sat nav', 'Part leather seats', 'Keyless entry'] },
];
const COLOURS = [['Silver', 5], ['Grey', 6], ['Black', 5], ['White', 5], ['Blue', 3], ['Red', 2], ['Green', 1], ['Bronze', 1]];
const PLATE = 'ABCDEFGHJKLMNOPRSTUVWXY';
// Where the customers live: towns around the dealer and their postcode districts.
const AREAS = [['London', ['SE9', 'SE12', 'SE13', 'SE6', 'SE3', 'SE18']], ['Bromley', ['BR1', 'BR2']], ['Sidcup', ['DA14', 'DA15']], ['Bexleyheath', ['DA6', 'DA7']], ['Dartford', ['DA1', 'DA2']], ['Croydon', ['CR0']]];
const STREETS = ['Harbour Street', 'Mill Lane', 'Station Road', 'Orchard Way', 'Church Street', 'Beech Grove', 'Riverside Drive', 'Cedar Avenue', 'Linden Close', 'Hawthorn Road', 'Kingfisher Way', 'Ashdown Crescent', 'Meadow Rise', 'Elm Park Road'];
const REPAIRS = [
  ['Repair', 'Front brake pads and discs', 170, 320, 3], ['Tyres', 'Two front tyres', 130, 220, 4], ['Bodywork', 'Smart repair to a bumper scuff', 90, 180, 3],
  ['Bodywork', 'Alloy wheel refurbishment, two wheels', 70, 160, 2], ['Repair', 'Windscreen replaced', 180, 340, 1], ['Repair', 'New battery', 90, 160, 2],
  ['Repair', 'Wiper blades and bulbs', 20, 45, 3], ['Repair', 'Timing belt and water pump', 320, 520, 1], ['Repair', 'Air conditioning re-gas', 55, 90, 2],
  ['Repair', 'Clutch replaced', 450, 760, 1], ['Bodywork', 'Dent removal, driver\'s door', 60, 140, 2], ['Tyres', 'Four tyres', 260, 420, 1],
];
const GARAGES = ['Linden Close Garage', 'Canal Walk Bodyshop', 'Station Road Tyres', 'Hawthorn Auto Glass'];
const PX_NOTES = {
  Good: ['Clean inside and out, tyres good, service book stamped.', 'Light marks on the rear bumper, otherwise tidy.', 'One small stone chip on the bonnet, drives well.'],
  Fair: ['Scuffed alloys and a dent in the rear door; brakes will be due soon.', 'Interior worn on the driver\'s seat, two tyres near the limit.', 'Service history incomplete; scratches on both bumpers.'],
  Poor: ['Warning light on the dash, clutch slipping, needs a service.', 'Heavy wear inside, bodywork damage on the near side, MOT due soon.'],
};
const NOT_BOUGHT = {
  'Thinking it over': ['Liked it, wants to talk it over at home.', 'Waiting to hear about the finance figures.', 'Wants to bring a partner back to see it.', 'Comparing it with one elsewhere.'],
  'Not for them': ['Too small for the family.', 'Did not like the drive.', 'Wanted an automatic after all.', 'Over their budget.'],
  'Bought elsewhere': ['Bought a newer one from another dealer.', 'Found one privately.', 'Went for a different make.'],
};

/** Fill a Copperbeech Motors database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 450 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };
  const iso = (d) => cal.day(d);
  // The dealer closes on Christmas Day, Boxing Day and New Year's Day.
  const closed = (d) => { const t = cal.at(d); const m = t.getMonth(); const day = t.getDate(); return (m === 11 && (day === 25 || day === 26)) || (m === 0 && day === 1); };
  const nextOpen = (d) => { while (closed(d)) d++; return d; };
  const yearOf = (d) => cal.at(d).getFullYear();
  const numbers = {};
  const number = (prefix, d) => { const y = yearOf(d); const k = `${prefix}${y}`; numbers[k] = (numbers[k] || 0) + 1; return `${prefix}-${y}-${String(numbers[k]).padStart(4, '0')}`; };
  const round25 = (n) => Math.max(250, Math.round(n / 25) * 25);
  const askFor = (n) => Math.ceil(n / 100) * 100 - 5;

  add('Business', businessRecord({
    name: 'Copperbeech Motors', domain: 'copperbeech.example', phone: '020 7946 0420',
    address: { Address: '41 Quarry Road', City: 'London', Postcode: 'SE12 8QR', Country: 'United Kingdom' },
    terms: 'The balance is due at handover, by bank transfer or debit card, before the car is driven away.',
    footer: 'Open Monday to Saturday 9am to 6pm, Sunday 10am to 4pm. Every car is sold with at least three months\' warranty. Thank you for buying from us.',
  }));

  // Staff: the owner sells too; one sales executive joined part-way through.
  const ROLES = [['Owner and Sales Manager', 3400, true, 0.6], ['Sales Executive', 1900, true, 1], ['Sales Executive', 950, true, 1], ['Sales Executive', 210, true, 1], ['Workshop Technician', 2600, false], ['Valeter', 720, false], ['Administrator', 1400, false]];
  const staff = ROLES.map(([role, since, sells, weight], i) => {
    const p = person(r, { domain: 'copperbeech.example' });
    const id = add('Staff', { 'First Name': p['First Name'], 'Last Name': p['Last Name'], Role: role, Phone: `020 7946 0${String(421 + i).padStart(3, '0')}`, 'E-mail': p['E-mail'], Started: iso(-since), Sells: sells });
    return { id, role, since: -since, sells, weight: weight || 0 };
  });
  const sellers = staff.filter((s) => s.sells);
  const sellerOn = (d) => r.weighted(sellers.filter((s) => s.since <= d).map((s) => [s, s.weight])).id;

  // Where cars come from.
  const supplier = (name, kind, city, post) => ({ id: add('Suppliers', { Name: name, Kind: kind, Address: `${r.int(1, 90)} ${r.pick(STREETS)}`, City: city, Postcode: post, Phone: phoneFor(r, 'United Kingdom'), 'E-mail': `buyers@${name.toLowerCase().replace(/[^a-z]+/g, '')}.example`, 'Account No': `BUY-${r.int(10000, 99999)}` }), kind });
  const auctions = [supplier('Riverbend Vehicle Auctions', 'Auction', 'Dartford', 'DA1 5QR'), supplier('Northgate Car Auctions', 'Auction', 'London', 'E6 7TU')];
  const trade = [supplier('Granite Fleet Disposals', 'Trade', 'Croydon', 'CR0 4HJ'), supplier('Kestrel Lease Returns', 'Trade', 'London', 'SE18 6NP'), supplier('Marlow Motor Traders', 'Trade', 'Bromley', 'BR1 3DE')];

  const lenders = [
    { name: 'Northlight Motor Finance', rate: 0.045, apr: 10.9, prefix: 'NMF' },
    { name: 'Ashgrove Credit', rate: 0.035, apr: 12.9, prefix: 'ASG' },
  ].map((l) => ({ ...l, id: add('Lenders', { Name: l.name, Address: `${r.int(1, 200)} ${r.pick(STREETS)}, Leeds`, Phone: phoneFor(r, 'United Kingdom'), 'E-mail': `dealers@${l.name.toLowerCase().replace(/[^a-z]+/g, '')}.example`, 'Commission Rate': l.rate, 'Typical APR': l.apr }) }));

  const ADMIN = 'Steadfast Motor Cover, for Copperbeech Motors';
  const COVER_BASIC = 'The internal parts of the engine and gearbox, the clutch plate excepted, and the drive shafts, for faults that appear in normal use. Parts and labour are paid up to the claim limit.';
  const COVER_FULL = 'The engine and gearbox internal parts, the turbocharger, the cooling system, the fuel system, the steering and suspension, the brakes (but not pads and discs), the electrical parts and the starter motor, the air conditioning, and for a hybrid or electric car the drive motor and the battery management. Parts and labour are paid up to the claim limit, with no excess to pay.';
  const plans = [
    { name: 'Dealer warranty, 3 months', months: 3, limit: 750, price: 0, cost: 0, cover: COVER_BASIC, weight: 6 },
    { name: 'Essential, 12 months', months: 12, limit: 1000, price: 349, cost: 145, cover: COVER_FULL, weight: 2.4 },
    { name: 'Extended, 24 months', months: 24, limit: 1500, price: 549, cost: 230, cover: COVER_FULL, weight: 1.2 },
    { name: 'Extended, 36 months', months: 36, limit: 2000, price: 749, cost: 320, cover: COVER_FULL, weight: 0.6 },
  ].map((p) => ({ ...p, id: add('Warranty Plans', { Name: p.name, Months: p.months, 'Claim Limit': p.limit, Price: p.price, Cost: p.cost, 'Administered By': ADMIN, Cover: p.cover }) }));

  // Customers, made as they first come in.
  const customers = [];
  const newCustomer = (d, heard) => {
    const p = person(r, { domain: r.pick(['example.com', 'example.org', 'example.net']) });
    const [town, districts] = r.weighted(AREAS.map((a, i) => [a, i === 0 ? 4 : 1]));
    const id = add('Customers', {
      'First Name': p['First Name'], 'Last Name': p['Last Name'], Address: `${r.int(1, 180)} ${r.pick(STREETS)}`, City: town,
      Postcode: `${r.pick(districts)} ${r.int(1, 9)}${r.pick(['AB', 'DE', 'FG', 'HJ', 'NP', 'QR', 'TU', 'WX'])}`, Phone: phoneFor(r, 'United Kingdom'),
      'E-mail': p['E-mail'], 'Heard Of Us': heard || r.weighted([['Website', 5], ['Walk-in', 3], ['Recommended', 2]]), Added: iso(d), 'Marketing Consent': r.chance(0.55),
    });
    const c = { id, first: p['First Name'] };
    customers.push(c);
    return c;
  };

  // A car as it is found: invented make and model, an age, the mileage that goes with it, and what it is worth to the trade.
  const plates = new Set();
  const plateFor = (reg) => {
    const y = reg.getFullYear() % 100;
    const m = reg.getMonth() + 1;
    const age = m >= 3 && m <= 8 ? y : m >= 9 ? y + 50 : y - 1 + 50;
    for (;;) {
      const p = `ZZ${String(age).padStart(2, '0')} ${PLATE[r.int(0, PLATE.length - 1)]}${PLATE[r.int(0, PLATE.length - 1)]}${PLATE[r.int(0, PLATE.length - 1)]}`;
      if (!plates.has(p)) { plates.add(p); return p; }
    }
  };
  const carSpec = (d, lo, hi) => {
    const [make, model, body, list, fuels, engines] = r.weighted(RANGE.map((m) => [m, m[6]]));
    const trim = r.weighted(TRIMS.map((t) => [t, t.weight]));
    const fuel = r.weighted(fuels);
    const months = r.int(lo * 12, hi * 12);
    const first = cal.at(d - Math.round(months * 30.44) - r.int(0, 20));
    const years = months / 12;
    const mileage = Math.max(1800, Math.round(years * r.int(5200, 11800)) + r.int(0, 999));
    const engine = fuel === 'Electric' ? `${r.pick(['50', '64', '77'])} kWh` : r.pick(engines);
    const value = Math.max(900, list * trim.uplift * 0.74 * Math.pow(0.85, years) - (mileage - years * 8000) * 0.035);
    const ti = TRIMS.indexOf(trim);
    const features = TRIMS.slice(0, ti + 1).flatMap((t) => t.features);
    const pickFeatures = [...r.some(features, Math.min(features.length, r.int(5, 8)))];
    if (r.chance(0.12)) pickFeatures.push('Tow bar');
    return {
      make, model: `${model} ${engine} ${trim.name}`, body, fuel, year: first.getFullYear(), first: cal.iso(first), reg: plateFor(first),
      colour: r.weighted(COLOURS), transmission: fuel === 'Electric' || fuel === 'Hybrid' || r.chance(0.3) ? 'Automatic' : 'Manual',
      doors: body === 'Saloon' ? 4 : body === 'Hatchback' ? r.pick([3, 5, 5, 5]) : 5, mileage, owners: Math.max(1, Math.min(6, Math.round(years / 3.2) + r.int(0, 1))),
      history: r.weighted([['Full', 6], ['Part', 3], ['None', 1]]), keys: r.chance(0.85) ? 2 : 1, value: Math.round(value), features: pickFeatures,
      // When its MOT runs out: three years from new, then yearly.
      mot: years < 3 ? (() => { const t = cal.at(0); t.setTime(first.getTime()); t.setFullYear(first.getFullYear() + 3); return Math.round((t - cal.at(0)) / 86400000); })() : d + r.int(20, 340),
      appeal: r.weighted([[1, 7], [0.6, 2], [0.22, 1]]), years,
    };
  };

  // The cars, in stock or sold.
  const cars = [];
  let stockNo = 1000;
  const prepJob = (car, d, kind, description, by, cost) => {
    if (d > 0) return;
    add('Preparation', { Vehicle: car.id, 'Done On': iso(d), Kind: kind, Description: description, 'Done By': by, Cost: money(cost) });
    car.prep = money(car.prep + cost);
  };
  /** A car comes in on day `d`: its record, its features, and its preparation over the next few days. */
  const takeIn = (d, s, source, sup, price, fee) => {
    const vatq = source !== 'Part-exchange' && s.years <= 4 && r.chance(0.35);
    const car = { ...s, source, boughtAt: d, price, fee, prep: 0, status: 'In preparation', vatq, nextCut: 35 };
    car.stockNo = `STK-${++stockNo}`;
    car.id = add('Vehicles', {
      Registration: s.reg, 'Stock No': car.stockNo, Make: s.make, Model: s.model, Body: s.body, Year: s.year, 'First Registered': s.first, Colour: s.colour,
      Fuel: s.fuel, Transmission: s.transmission, Doors: s.doors, Mileage: s.mileage, Owners: s.owners, 'Service History': s.history,
      'MOT Expiry': iso(s.mot), Keys: s.keys, Source: source, Supplier: sup, 'Bought On': iso(d), 'Purchase Price': price, 'Buying Fee': fee,
      'VAT Qualifying': vatq, Status: 'In preparation',
    });
    for (const f of s.features) add('Vehicle Features', { Vehicle: car.id, Feature: f });
    // The inspection the day after: a service when the history is thin, repairs it finds, an MOT when the old one is close to its end.
    let at = d + 1;
    if (s.history !== 'Full' || r.chance(0.35)) prepJob(car, at, 'Service', r.chance(0.5) ? 'Full service' : 'Interim service, oil and filters', 'In house', r.int(140, 320));
    const repairs = r.chance(0.5) ? r.int(1, s.years > 6 ? 3 : 2) : 0;
    for (const [kind, what, lo, hi] of r.some(REPAIRS.flatMap((x) => Array(x[4]).fill(x)), repairs).filter((x, i, a) => a.indexOf(x) === i)) {
      at += r.int(0, 1);
      prepJob(car, at, kind, what, kind === 'Bodywork' || what.startsWith('Windscreen') ? r.pick(GARAGES) : 'In house', r.int(lo, hi));
    }
    if (car.mot - d < 150) { at += 1; prepJob(car, at, 'MOT', 'MOT test, passed', 'Station Road MOT Centre', 54.85); if (at <= 0) car.mot = at + 365; }
    at += r.int(0, 2);
    prepJob(car, at, 'Valet', r.chance(0.7) ? 'Full valet' : 'Full valet and machine polish', 'In house', r.int(35, 70));
    car.readyAt = at + 1;
    // The price: what it cost and its preparation, and enough over for the margin.
    car.first = askFor((price + fee + 380) * 1.1 + 1000);
    car.asking = car.first;
    cars.push(car);
    return car;
  };
  const buy = (d, from) => {
    const sup = from || (r.chance(0.8) ? r.pick(auctions) : r.pick(trade));
    const s = carSpec(d, 2, 10);
    const auction = sup.kind === 'Auction';
    const price = round25(s.value * (auction ? 0.93 + r.next() * 0.08 : 0.97 + r.next() * 0.06));
    return takeIn(d, s, sup.kind, sup.id, price, auction ? Math.min(650, Math.round(150 + price * 0.03)) : 0);
  };
  const inStock = () => cars.filter((c) => c.status !== 'Sold');

  // Fifteen months from the first of a month, so each month's figures are a whole month's.
  let start = -days;
  while (cal.at(start).getDate() !== 1) start++;
  // The forecourt as it stood then: cars bought in the weeks before.
  for (let k = 0; k < 38; k++) buy(start - r.int(2, 80));

  const sales = [];
  const valuations = [];
  /** A valuation of the customer's own car on day `d`. */
  const value = (d, c, sellerId, forSale) => {
    const s = carSpec(d, 4, 13);
    const condition = r.weighted([['Good', 5], ['Fair', 4], ['Poor', 1]]);
    const guide = round25(s.value * (condition === 'Poor' ? 0.8 : condition === 'Fair' ? 0.92 : 1));
    let offered = round25(guide * (0.86 + r.next() * 0.09));
    if (forSale) offered = Math.min(offered, round25(forSale * 0.6));
    const v = { s, d, c, offered, outcome: 'Open', validUntil: d + 7 };
    v.id = add('Part Exchanges', {
      Number: number('PX', d), Customer: c.id, 'Valued By': sellerId, 'Valued On': iso(d), Registration: s.reg, Make: s.make, Model: s.model, Year: s.year,
      Mileage: s.mileage, Colour: s.colour, Fuel: s.fuel, Condition: condition, Notes: r.pick(PX_NOTES[condition]), 'Guide Value': guide, Offered: offered,
      'Valid Until': iso(d + 7), Outcome: 'Open',
    });
    valuations.push(v);
    return v;
  };
  /** A deal agreed on day `d`: the order, its lines, the deposit, finance and the handover to come. */
  const sell = (d, car, c, sellerId, px) => {
    const price = car.asking - r.weighted([[0, 4], [100, 2], [200, 2], [300, 1], [500, 1]]);
    const plan = r.weighted(plans.map((p) => [p, p.weight]));
    const lines = [{ kind: 'Vehicle', text: `${car.year} ${car.make} ${car.model}, ${car.colour}, ${car.reg}, ${car.mileage.toLocaleString('en-GB')} miles`, amount: price, rate: car.vatq ? 0.2 : 0 }];
    lines.push({ kind: 'Warranty', text: plan.price ? `${plan.name} warranty` : `${plan.name}, included`, amount: plan.price, rate: 0 });
    if (r.chance(0.15)) lines.push({ kind: 'Extra', text: 'Fitted mats, set of four', amount: 35, rate: 0.2 });
    if (r.chance(0.12)) lines.push({ kind: 'Extra', text: 'Paint and interior protection', amount: 199, rate: 0.2 });
    if (r.chance(0.08)) lines.push({ kind: 'Extra', text: 'Delivery to your home', amount: 95, rate: 0.2 });
    const total = money(lines.reduce((n, l) => n + l.amount, 0));
    const vat = money(lines.reduce((n, l) => n + money(l.amount - l.amount / (1 + l.rate)), 0));
    const allowance = px ? px.offered : 0;
    const deposit = Math.min(r.pick([100, 200, 250, 500, 500, 1000]), Math.floor(price * 0.1 / 50) * 50);
    let finance = 0;
    let lender = null;
    if (r.chance(0.45)) {
      const amount = total - allowance - deposit - r.pick([0, 0, 500, 1000, 2000]);
      if (amount >= 1500) { finance = money(amount); lender = r.weighted([[lenders[0], 3], [lenders[1], 2]]); }
    }
    const handoverAt = nextOpen(d + r.int(2, 8));
    const sale = { car, c, sellerId, px, plan, d, handoverAt, price, total, vat, allowance, deposit, finance, lender, status: 'Reserved' };
    sale.cancelAt = handoverAt - d >= 3 && r.chance(0.04) ? d + r.int(1, handoverAt - d - 1) : null;
    sale.id = add('Sales', {
      'Order No': number('ORD', d), Vehicle: car.id, Customer: c.id, Salesperson: sellerId, 'Agreed On': iso(d), Handover: iso(handoverAt),
      Price: price, Total: total, VAT: vat, 'VAT Scheme': car.vatq ? 'VAT qualifying' : 'Margin scheme', 'Part-Ex Allowance': allowance, Deposit: deposit,
      'Finance Amount': finance, Balance: money(total - allowance - deposit - finance), Status: 'Reserved',
    });
    lines.forEach((l, i) => add('Sale Lines', { Sale: sale.id, Line: i + 1, Kind: l.kind, Description: l.text, Amount: l.amount, 'Tax Rate': l.rate }));
    if (lender) {
      const term = r.pick([24, 36, 48, 48, 60]);
      const pcp = r.chance(0.35);
      const balloon = pcp ? Math.min(Math.round(price * (0.42 - term / 300)), Math.round(finance * 0.5)) : 0;
      const i = lender.apr / 1200;
      const monthly = money((finance - balloon / Math.pow(1 + i, term)) * i / (1 - Math.pow(1 + i, -term)));
      sale.finance_id = add('Finance Agreements', {
        Sale: sale.id, Lender: lender.id, Reference: `${lender.prefix}-${r.int(100000, 999999)}`, Kind: pcp ? 'Personal contract purchase' : 'Hire purchase',
        'Amount Financed': finance, 'Term Months': term, APR: lender.apr, 'Monthly Payment': monthly, 'Final Payment': balloon,
        Commission: money(finance * lender.rate), Status: 'Approved',
      });
    }
    if (px) { px.outcome = 'Accepted'; base.update('Part Exchanges', px.id, { Outcome: 'Accepted', Sale: sale.id }); }
    car.status = 'Reserved';
    sales.push(sale);
    return sale;
  };

  const SEASON = [0.8, 0.95, 1.3, 1.1, 1.05, 1, 0.95, 0.95, 1.2, 1, 0.9, 0.65];
  const drives = [];
  for (let d = start; d <= 0; d++) {
    if (closed(d)) continue;
    const wd = cal.weekday(d);
    // Cars ready for sale, and price cuts on those that have stood.
    for (const car of cars) {
      if (car.status === 'In preparation' && car.readyAt <= d) car.status = 'For sale';
      if (car.status === 'For sale' && d - car.boughtAt >= car.nextCut) {
        car.asking = Math.max(askFor(car.price + car.fee + car.prep + 250), askFor(car.asking * 0.96));
        car.appeal *= 1.2;
        car.nextCut += 25;
      }
    }
    // Deals that end today: handed over, or cancelled before handover.
    for (const s of sales) {
      if (s.status !== 'Reserved') continue;
      if (s.cancelAt === d) {
        s.status = 'Cancelled';
        s.car.status = 'For sale';
        base.update('Sales', s.id, { Status: 'Cancelled', Notes: `${r.pick(['Finance declined', 'Customer changed their mind', 'Customer found the money short'])}. Deposit of ${s.deposit} returned on ${iso(d)}.` });
        if (s.finance_id) base.update('Finance Agreements', s.finance_id, { Status: 'Cancelled' });
        if (s.px) { s.px.outcome = 'Withdrawn'; base.update('Part Exchanges', s.px.id, { Outcome: 'Withdrawn' }); }
      } else if (s.handoverAt === d) {
        s.status = 'Completed';
        s.car.status = 'Sold';
        if (s.finance_id) {
          const paid = d + r.int(1, 3);
          base.update('Finance Agreements', s.finance_id, paid <= 0 ? { Status: 'Paid out', 'Paid Out On': iso(paid) } : { Status: 'Approved' });
        }
        // The part-exchange comes into stock as the new car leaves.
        if (s.px) { const pxCar = takeIn(d, s.px.s, 'Part-exchange', null, s.px.offered, 0); base.update('Part Exchanges', s.px.id, { 'Taken In': pxCar.id }); }
      }
    }
    // Buying: at auction on Tuesdays and Thursdays to bring the stock back to forty, and now and then from the trade.
    const stock = inStock().length;
    if ((wd === 2 || wd === 4) && stock < 40) for (let k = 0; k < Math.min(6, 40 - stock + r.int(0, 1)); k++) buy(d, r.pick(auctions));
    if ((wd === 1 || wd === 3 || wd === 5) && stock < 42 && r.chance(0.35)) buy(d, r.pick(trade));

    // Test drives: more on Saturdays, fewer on Sundays, more in the plate-change months.
    const forSale = () => cars.filter((c) => c.status === 'For sale');
    const n = Math.round(3.3 * SEASON[cal.at(d).getMonth()] * (wd === 6 ? 1.6 : wd === 0 ? 0.7 : 1) * (0.5 + r.next()));
    const lastHour = d === 0 ? Math.min(17, cal.hour) : 17;
    for (let k = 0; k < n; k++) {
      const list = forSale();
      if (!list.length) break;
      const car = r.weighted(list.map((x) => [x, x.appeal]));
      const c = customers.length > 50 && r.chance(0.06) ? r.pick(customers) : newCustomer(d);
      const seller = sellerOn(d);
      const hour = wd === 0 ? r.int(10, 15) : r.int(9, 17);
      const when = cal.time(d, hour, r.pick([0, 15, 30, 45]));
      const done = hour <= lastHour;
      const buys = done && r.chance(0.27 * Math.min(1.2, car.appeal + 0.2));
      const outcome = !done ? 'Booked' : buys ? 'Bought' : r.weighted([['Thinking it over', 4], ['Not for them', 3], ['Bought elsewhere', 2]]);
      drives.push({ Vehicle: car.id, Customer: c.id, Staff: seller, When: when, Outcome: outcome, Notes: done && !buys ? r.pick(NOT_BOUGHT[outcome]) : null });
      if (buys) sell(d, car, c, seller, r.chance(0.35) ? value(d, c, seller, car.asking) : null);
      else if (done && r.chance(0.12)) value(d, c, seller, null);
    }
    // A few bought without a test drive, from the website.
    if (r.chance(0.12)) { const list = forSale(); if (list.length) { const car = r.pick(list); const c = newCustomer(d, 'Website'); sell(d, car, c, sellerOn(d), null); } }
    // Every day ends with at least one car reserved, as a dealer this size has.
    if (d === 0 && !sales.some((s) => s.status === 'Reserved')) {
      const list = forSale();
      const car = list.length ? list[0] : null;
      if (car) { const c = newCustomer(d); const seller = sellerOn(d); drives.push({ Vehicle: car.id, Customer: c.id, Staff: seller, When: cal.time(d, 9, 30), Outcome: 'Bought', Notes: null }); sell(d, car, c, seller, null); }
    }
  }
  // Test drives booked for the coming week.
  for (let d = 1; d <= 7; d++) {
    if (closed(d)) continue;
    const list = cars.filter((c) => c.status === 'For sale');
    for (let k = 0; k < r.int(1, 3) && list.length; k++) {
      const c = newCustomer(0);
      drives.push({ Vehicle: r.pick(list).id, Customer: c.id, Staff: sellerOn(0), When: cal.time(d, r.int(10, 16), r.pick([0, 30])), Outcome: 'Booked', Notes: null });
    }
  }
  drives.sort((a, b) => (a.When < b.When ? -1 : a.When > b.When ? 1 : 0));
  for (const t of drives) add('Test Drives', t);

  // Valuations not taken up: declined, or the offer has run out.
  for (const v of valuations) {
    if (v.outcome !== 'Open' || v.validUntil >= 0) continue;
    base.update('Part Exchanges', v.id, { Outcome: r.chance(0.5) ? 'Declined' : 'Expired' });
  }

  // Invoices and warranties for the cars handed over, numbered in date order.
  const done = sales.filter((s) => s.status === 'Completed').sort((a, b) => a.handoverAt - b.handoverAt || a.id - b.id);
  for (const s of done) {
    base.update('Sales', s.id, { 'Invoice No': number('INV', s.handoverAt), Status: 'Completed' });
    const startAt = cal.at(s.handoverAt);
    const end = new Date(startAt.getFullYear(), startAt.getMonth() + s.plan.months, startAt.getDate() - 1);
    add('Warranties', { 'Certificate No': number('WAR', s.handoverAt), Sale: s.id, Plan: s.plan.id, Starts: iso(s.handoverAt), Ends: cal.iso(end), 'Mileage at Start': s.car.mileage + r.int(5, 40), Price: s.plan.price, Cost: s.plan.cost });
  }

  // Each car's record as it stands today.
  for (const car of cars) base.update('Vehicles', car.id, { Status: car.status, 'Asking Price': car.asking, 'First Asking Price': car.first, 'Prep Cost': car.prep, 'MOT Expiry': iso(car.mot) });
  return counts;
}
