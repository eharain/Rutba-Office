// The cafe: Little Fig Cafe, an invented all-day cafe with about thirty
// seats and a takeaway counter. Its menu in categories, each dish's recipe
// (so every dish has a food cost), the ingredients and who supplies them,
// deliveries in and the bills for them, the till's orders and their lines,
// the staff and their rota, what was thrown away, and the cash-up sheet
// filled in at each day's close. It prints a till receipt for each bill
// (VAT within the menu prices), a purchase order for each order to a
// supplier, and a weekly payslip for each person from the rota.
//
// Its records come from fourteen months of trading, run a day at a time:
// busier at weekends, in summer and on sunny days, breakfast until half past
// eleven and lunch after, cakes baked on site each morning. Every dish sold
// uses up its ingredients, deliveries come on each supplier's days for what
// has run down to its reorder level, and the fridge throws a little away.
// The till's orders and their lines are kept for the last eight weeks; the cash-up
// sheets, the rota, the deliveries and the waste go back the whole fourteen
// months. The rota runs two weeks ahead.

import { BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, money, person, phoneFor, random } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const LINE = 'Round([Quantity] * [Unit Price], 2)';
const RECIPE_COST = 'Round([Quantity] * [Cost per Unit], 2)';
const DAYS_AGO = 'DateDiff("d", Format([When], "yyyy-mm-dd"), Date())';
// Days counted from Monday 1 January 2024, so a date's place in its week
// (0 is Monday) and the week it falls in can be worked out from the date alone.
const DAY_NO = 'DateDiff("d", #2024-01-01#, [When]) Mod 7';
const DAY_NAME = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
  .reduceRight((rest, name, i) => `IIf(${DAY_NO} = ${i}, "${name}", ${rest})`, '"Sunday"');
const WEEK_NO = '(DateDiff("d", #2024-01-01#, [Date]) - DateDiff("d", #2024-01-01#, [Date]) Mod 7) / 7';
const TAKINGS = '[Card Takings] + [Cash Takings]';
// The VAT within a line's tax-inclusive price, as the till works it out.
const VAT_IN = 'Round([Price] - [Price] / (1 + [Tax Rate]), 2)';
// The day of the week a shift starts on, for the payslip.
const SHIFT_DAY_NO = 'DateDiff("d", #2024-01-01#, [Start]) Mod 7';
const SHIFT_DAY = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
  .reduceRight((rest, name, i) => `IIf(${SHIFT_DAY_NO} = ${i}, "${name}", ${rest})`, '"Sunday"');

export const spec = {
  // Where a value below nought is right: cash short in the till, a bill already past its day.
  negative: { 'Cash over or short': ['Over or Short'], 'Bills to pay': ['Days to Pay'] },
  tables: [
    BUSINESS_TABLE,
    { name: 'Menu Categories', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Served', type: 'text' }, { name: 'Sort Order', type: 'integer' },
    ] },
    { name: 'Menu Items', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Category', ...ref('Menu Categories'), required: true },
      { name: 'Price', type: 'currency' }, { name: 'Takeaway', type: 'yesno' }, { name: 'Baked Here', type: 'yesno' },
      { name: 'Diet', type: 'text' }, { name: 'Takeaway VAT Rate', type: 'number', default: 0.2 }, { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Suppliers', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Supplies', type: 'text' }, { name: 'Address', type: 'text' },
      { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' }, { name: 'Contact', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Delivers On', type: 'text' },
      { name: 'Payment Days', type: 'integer', default: 30 },
    ] },
    { name: 'Ingredients', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Supplier', ...ref('Suppliers'), required: true },
      { name: 'Unit', type: 'text' }, { name: 'Cost per Unit', type: 'currency' }, { name: 'Bought As', type: 'text' },
      { name: 'Storage', type: 'text' }, { name: 'In Stock', type: 'number', default: 0 }, { name: 'On Order', type: 'number', default: 0 },
      { name: 'Reorder Level', type: 'number', default: 0 }, { name: 'Reorder Quantity', type: 'number', default: 1 },
    ] },
    { name: 'Recipe Lines', fields: [
      { name: 'Menu Item', ...ref('Menu Items'), required: true }, { name: 'Line', type: 'integer', default: 1 },
      { name: 'Ingredient', ...ref('Ingredients'), required: true }, { name: 'Quantity', type: 'number' },
    ] },
    { name: 'Staff', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Role', type: 'text' },
      { name: 'Hourly Rate', type: 'currency' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Started', type: 'date' },
    ] },
    // A payslip for each person for each week, made from the week's shifts once it has been worked.
    { name: 'Payslips', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Staff', ...ref('Staff'), required: true }, { name: 'Week Beginning', type: 'date' },
      { name: 'Week Ending', type: 'date' }, { name: 'Pay Date', type: 'date' }, { name: 'Hours', type: 'number' },
      { name: 'Hourly Rate', type: 'currency' }, { name: 'Gross Pay', type: 'currency' },
    ] },
    { name: 'Shifts', fields: [
      { name: 'Staff', ...ref('Staff'), required: true }, { name: 'Start', type: 'datetime' }, { name: 'End', type: 'datetime' },
      { name: 'Break Minutes', type: 'integer', default: 0 }, { name: 'Paid Hours', type: 'number' },
      { name: 'Station', type: 'text' }, { name: 'Status', type: 'text', default: 'Rota' }, { name: 'Payslip', ...ref('Payslips') },
    ] },
    { name: 'Orders', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'When', type: 'datetime' }, { name: 'Service', type: 'text', default: 'Dine in' }, { name: 'Table No', type: 'integer' },
      { name: 'Covers', type: 'integer', default: 0 }, { name: 'Staff', ...ref('Staff'), required: true },
      { name: 'Payment', type: 'text', default: 'Card' }, { name: 'Net', type: 'currency' }, { name: 'Tax', type: 'currency' },
      { name: 'Total', type: 'currency' }, { name: 'Voided', type: 'yesno' },
    ] },
    { name: 'Order Lines', fields: [
      { name: 'Order', ...ref('Orders'), required: true }, { name: 'Menu Item', ...ref('Menu Items'), required: true },
      { name: 'Quantity', type: 'integer', default: 1 }, { name: 'Unit Price', type: 'currency' }, { name: 'Tax Rate', type: 'number', default: 0.2 },
    ] },
    { name: 'Purchases', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Supplier', ...ref('Suppliers'), required: true }, { name: 'Ordered On', type: 'date' }, { name: 'Due On', type: 'date' },
      { name: 'Received On', type: 'date' }, { name: 'Invoice Number', type: 'text' }, { name: 'Paid On', type: 'date' },
    ] },
    { name: 'Purchase Lines', fields: [
      { name: 'Purchase', ...ref('Purchases'), required: true }, { name: 'Ingredient', ...ref('Ingredients'), required: true },
      { name: 'Quantity', type: 'number' }, { name: 'Unit Cost', type: 'currency' },
    ] },
    { name: 'Waste', fields: [
      { name: 'Ingredient', ...ref('Ingredients'), required: true }, { name: 'Date', type: 'date' }, { name: 'Quantity', type: 'number' },
      { name: 'Reason', type: 'text' }, { name: 'Recorded By', ...ref('Staff') },
    ] },
    { name: 'Cash Ups', fields: [
      { name: 'Date', type: 'date', required: true, unique: true }, { name: 'Closed By', ...ref('Staff') }, { name: 'Weather', type: 'text' },
      { name: 'Orders', type: 'integer' }, { name: 'Covers', type: 'integer' }, { name: 'Card Takings', type: 'currency' },
      { name: 'Cash Takings', type: 'currency' }, { name: 'VAT', type: 'currency' }, { name: 'Cash Counted', type: 'currency' }, { name: 'Wage Cost', type: 'currency' },
      { name: 'Notes', type: 'longtext' },
    ] },
  ],
  queries: {
    'Menu costing': { tables: ['Menu Items', 'Menu Categories', 'Recipe Lines', 'Ingredients'], columns: [
      { table: 'Menu Categories', field: 'Name', as: 'Category', total: 'group' }, { table: 'Menu Categories', field: 'Sort Order', total: 'group', sort: 'asc', show: false },
      { table: 'Menu Items', field: 'Name', as: 'Menu Item', total: 'group', sort: 'asc' }, { table: 'Menu Items', field: 'Price', total: 'group' },
      { expr: RECIPE_COST, as: 'Food Cost', total: 'sum' },
      { expr: 'Round(100 * [Quantity] * [Cost per Unit] / [Price], 2)', as: 'Food Cost %', total: 'sum' },
      // The price is counted once, on the recipe's first line, so the sum over the lines is the price less the food cost.
      { expr: `Round(IIf([Line] = 1, [Price], 0) - ${RECIPE_COST}, 2)`, as: 'Gross Profit', total: 'sum' },
      { expr: 'Round(IIf([Line] = 1, 100, 0) - 100 * [Quantity] * [Cost per Unit] / [Price], 2)', as: 'Gross Margin %', total: 'sum' },
      { table: 'Menu Items', field: 'Active', total: 'where', criteria: ['Yes'] },
    ] },
    'Recipe card': { tables: ['Recipe Lines', 'Menu Items', 'Ingredients'], columns: [
      { table: 'Menu Items', field: 'Name', as: 'Menu Item', sort: 'asc' }, { table: 'Recipe Lines', field: 'Line', sort: 'asc' },
      { table: 'Ingredients', field: 'Name', as: 'Ingredient' }, { table: 'Recipe Lines', field: 'Quantity' }, { table: 'Ingredients', field: 'Unit' },
      { table: 'Ingredients', field: 'Cost per Unit' }, { expr: RECIPE_COST, as: 'Line Cost' },
      { table: 'Menu Items', field: 'Active', criteria: ['Yes'], show: false },
    ] },
    'Menu': { tables: ['Menu Items', 'Menu Categories'], columns: [
      { table: 'Menu Categories', field: 'Sort Order', sort: 'asc', show: false }, { table: 'Menu Categories', field: 'Name', as: 'Category' },
      { table: 'Menu Categories', field: 'Served' }, { table: 'Menu Items', field: 'Name', as: 'Menu Item', sort: 'asc' },
      { table: 'Menu Items', field: 'Diet' }, { table: 'Menu Items', field: 'Takeaway' }, { table: 'Menu Items', field: 'Price' },
      { table: 'Menu Items', field: 'Active', criteria: ['Yes'], show: false },
    ] },
    'Best sellers, last 30 days': { tables: ['Order Lines', 'Orders', 'Menu Items', 'Menu Categories'], columns: [
      { table: 'Menu Items', field: 'Name', as: 'Menu Item', total: 'group' }, { table: 'Menu Categories', field: 'Name', as: 'Category', total: 'group' },
      { table: 'Order Lines', field: 'Quantity', as: 'Sold', total: 'sum', sort: 'desc' }, { expr: LINE, as: 'Takings', total: 'sum' },
      { expr: DAYS_AGO, as: 'Days Ago', total: 'where', criteria: ['Between 0 And 30'] },
      { table: 'Orders', field: 'Voided', total: 'where', criteria: ['No'] },
    ] },
    'Sales by hour': { tables: ['Orders'], columns: [
      { expr: 'Mid([When], 12, 2)', as: 'Hour', total: 'group', sort: 'asc' }, { table: 'Orders', field: 'ID', as: 'Orders', total: 'count' },
      { table: 'Orders', field: 'Total', as: 'Takings', total: 'sum' }, { table: 'Orders', field: 'Total', as: 'Average Spend', total: 'avg' },
      { expr: DAYS_AGO, as: 'Days Ago', total: 'where', criteria: ['Between 1 And 56'] },
      { table: 'Orders', field: 'Voided', total: 'where', criteria: ['No'] },
    ] },
    'Sales by day of the week': { tables: ['Orders'], columns: [
      { expr: DAY_NO, as: 'Day No', total: 'group', sort: 'asc', show: false }, { expr: DAY_NAME, as: 'Day', total: 'group' },
      { table: 'Orders', field: 'ID', as: 'Orders', total: 'count' }, { table: 'Orders', field: 'Covers', as: 'Covers', total: 'sum' },
      { table: 'Orders', field: 'Total', as: 'Takings', total: 'sum' }, { table: 'Orders', field: 'Total', as: 'Average Spend', total: 'avg' },
      { expr: DAYS_AGO, as: 'Days Ago', total: 'where', criteria: ['Between 1 And 56'] },
      { table: 'Orders', field: 'Voided', total: 'where', criteria: ['No'] },
    ] },
    'Average spend by service': { tables: ['Orders'], columns: [
      { table: 'Orders', field: 'Service', total: 'group', sort: 'asc' }, { table: 'Orders', field: 'ID', as: 'Orders', total: 'count' },
      { table: 'Orders', field: 'Covers', as: 'Covers', total: 'sum' }, { table: 'Orders', field: 'Total', as: 'Takings', total: 'sum' },
      { table: 'Orders', field: 'Total', as: 'Average Spend', total: 'avg' },
      { expr: DAYS_AGO, as: 'Days Ago', total: 'where', criteria: ['Between 0 And 30'] },
      { table: 'Orders', field: 'Voided', total: 'where', criteria: ['No'] },
    ] },
    'Takings and wages by week': { tables: ['Cash Ups'], columns: [
      { expr: WEEK_NO, as: 'Week No', total: 'group', show: false }, { table: 'Cash Ups', field: 'Date', as: 'Week Beginning', total: 'min', sort: 'desc' },
      { table: 'Cash Ups', field: 'Date', as: 'Days Open', total: 'count' }, { table: 'Cash Ups', field: 'Orders', as: 'Orders', total: 'sum' },
      { table: 'Cash Ups', field: 'Covers', as: 'Covers', total: 'sum' }, { expr: TAKINGS, as: 'Takings', total: 'sum' },
      { table: 'Cash Ups', field: 'Wage Cost', as: 'Wages', total: 'sum' },
      { expr: `Round(${TAKINGS} - [Wage Cost], 2)`, as: 'Left After Wages', total: 'sum' },
      { expr: `Round(100 * [Wage Cost] / (${TAKINGS}), 1)`, as: 'Average Day Wage %', total: 'avg' },
    ] },
    'Takings by month': { tables: ['Cash Ups'], columns: [
      { expr: 'Format([Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Cash Ups', field: 'Date', as: 'Days Open', total: 'count' },
      { table: 'Cash Ups', field: 'Orders', as: 'Orders', total: 'sum' }, { expr: TAKINGS, as: 'Takings', total: 'sum' },
      { table: 'Cash Ups', field: 'VAT', as: 'VAT', total: 'sum' }, { expr: `Round(${TAKINGS} - [VAT], 2)`, as: 'Takings less VAT', total: 'sum' },
      { table: 'Cash Ups', field: 'Wage Cost', as: 'Wages', total: 'sum' },
      { expr: `Round(${TAKINGS} - [Wage Cost], 2)`, as: 'Left After Wages', total: 'sum' },
    ] },
    'Takings by weather': { tables: ['Cash Ups'], columns: [
      { table: 'Cash Ups', field: 'Weather', total: 'group' }, { table: 'Cash Ups', field: 'Date', as: 'Days', total: 'count' },
      { expr: TAKINGS, as: 'Average Takings', total: 'avg', sort: 'desc' }, { table: 'Cash Ups', field: 'Covers', as: 'Average Covers', total: 'avg' },
      { expr: 'Format([Date], "mm")', as: 'Month', total: 'where', criteria: ['In ("05", "06", "07", "08", "09")'] },
    ] },
    'Cash over or short': { tables: ['Cash Ups', 'Staff'], columns: [
      { table: 'Cash Ups', field: 'Date', sort: 'desc' }, { expr: '[First Name] & " " & [Last Name]', as: 'Closed By' },
      { table: 'Cash Ups', field: 'Cash Takings' }, { table: 'Cash Ups', field: 'Cash Counted' },
      { expr: 'Round([Cash Counted] - [Cash Takings], 2)', as: 'Over or Short', criteria: ['<> 0'] }, { table: 'Cash Ups', field: 'Notes' },
    ] },
    'Hours and wages by staff, last 28 days': { tables: ['Shifts', 'Staff'], columns: [
      { expr: '[First Name] & " " & [Last Name]', as: 'Name', total: 'group' }, { table: 'Staff', field: 'Role', total: 'group' },
      { table: 'Staff', field: 'Hourly Rate', total: 'group' }, { table: 'Shifts', field: 'ID', as: 'Shifts', total: 'count' },
      { table: 'Shifts', field: 'Paid Hours', as: 'Hours', total: 'sum' }, { expr: 'Round([Paid Hours] * [Hourly Rate], 2)', as: 'Wages', total: 'sum', sort: 'desc' },
      { expr: 'DateDiff("d", Format([Start], "yyyy-mm-dd"), Date())', as: 'Days Ago', total: 'where', criteria: ['Between 1 And 28'] },
      { table: 'Shifts', field: 'Status', total: 'where', criteria: ['In ("Worked", "Holiday")'] },
    ] },
    'Rota for the next 7 days': { tables: ['Shifts', 'Staff'], columns: [
      { expr: 'Format([Start], "yyyy-mm-dd")', as: 'Day', sort: 'asc' }, { expr: 'Mid([Start], 12, 5)', as: 'From', sort: 'asc' },
      { expr: 'Mid([End], 12, 5)', as: 'To' }, { expr: '[First Name] & " " & [Last Name]', as: 'Name' },
      { table: 'Shifts', field: 'Station' }, { table: 'Shifts', field: 'Paid Hours' }, { table: 'Shifts', field: 'Status' },
      { expr: 'DateDiff("d", Date(), Format([Start], "yyyy-mm-dd"))', as: 'Days Ahead', criteria: ['Between 0 And 6'], show: false },
    ] },
    'Ingredients to order': { tables: ['Ingredients', 'Suppliers'], columns: [
      { table: 'Suppliers', field: 'Name', as: 'Supplier', sort: 'asc' }, { table: 'Suppliers', field: 'Delivers On' },
      { table: 'Ingredients', field: 'Name', as: 'Ingredient', sort: 'asc' }, { table: 'Ingredients', field: 'Unit' },
      { table: 'Ingredients', field: 'In Stock' }, { table: 'Ingredients', field: 'On Order' }, { table: 'Ingredients', field: 'Reorder Level' },
      { table: 'Ingredients', field: 'Reorder Quantity', as: 'Order' }, { table: 'Ingredients', field: 'Bought As' },
      { expr: 'Round([Reorder Quantity] * [Cost per Unit], 2)', as: 'Estimated Cost' },
      { expr: '[In Stock] + Nz([On Order]) - [Reorder Level]', as: 'Above Reorder Level', criteria: ['<= 0'], show: false },
    ] },
    'Deliveries due': { tables: ['Purchases', 'Suppliers', 'Purchase Lines'], columns: [
      { table: 'Purchases', field: 'ID', as: 'Purchase', total: 'group' }, { table: 'Suppliers', field: 'Name', as: 'Supplier', total: 'group' },
      { table: 'Purchases', field: 'Ordered On', total: 'group' }, { table: 'Purchases', field: 'Due On', total: 'group', sort: 'asc' },
      { table: 'Purchase Lines', field: 'ID', as: 'Lines', total: 'count' },
      { expr: 'Round([Quantity] * [Unit Cost], 2)', as: 'Value', total: 'sum' },
      { table: 'Purchases', field: 'Received On', total: 'where', criteria: ['Is Null'] },
    ] },
    'Bills to pay': { tables: ['Purchases', 'Suppliers', 'Purchase Lines'], columns: [
      { table: 'Purchases', field: 'ID', as: 'Purchase', total: 'group' }, { table: 'Suppliers', field: 'Name', as: 'Supplier', total: 'group' },
      { table: 'Purchases', field: 'Invoice Number', total: 'group' }, { table: 'Purchases', field: 'Received On', total: 'group', criteria: ['Is Not Null'] },
      { expr: '[Payment Days] - DateDiff("d", [Received On], Date())', as: 'Days to Pay', total: 'group', sort: 'asc' },
      { expr: 'Round([Quantity] * [Unit Cost], 2)', as: 'Amount', total: 'sum' },
      { table: 'Purchases', field: 'Paid On', total: 'where', criteria: ['Is Null'] },
    ] },
    'Spend by supplier by month': { tables: ['Purchases', 'Suppliers', 'Purchase Lines'], columns: [
      { expr: 'Format([Received On], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Suppliers', field: 'Name', as: 'Supplier', total: 'group', sort: 'asc' },
      { table: 'Purchases', field: 'ID', as: 'Deliveries', total: 'count' }, { expr: 'Round([Quantity] * [Unit Cost], 2)', as: 'Spent', total: 'sum' },
      { table: 'Purchases', field: 'Received On', total: 'where', criteria: ['Is Not Null'] },
    ] },
    'Waste by ingredient, last 90 days': { tables: ['Waste', 'Ingredients'], columns: [
      { table: 'Ingredients', field: 'Name', as: 'Ingredient', total: 'group' }, { table: 'Ingredients', field: 'Unit', total: 'group' },
      { table: 'Waste', field: 'ID', as: 'Times', total: 'count' }, { table: 'Waste', field: 'Quantity', as: 'Wasted', total: 'sum' },
      { expr: RECIPE_COST, as: 'Cost', total: 'sum', sort: 'desc' },
      { expr: 'DateDiff("d", [Date], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 90'] },
    ] },
    'Waste in the last 30 days': { tables: ['Waste', 'Ingredients', 'Staff'], columns: [
      { table: 'Waste', field: 'Date', sort: 'desc' }, { table: 'Waste', field: 'Reason', sort: 'asc' }, { table: 'Ingredients', field: 'Name', as: 'Ingredient' },
      { table: 'Waste', field: 'Quantity' }, { table: 'Ingredients', field: 'Unit' }, { expr: RECIPE_COST, as: 'Cost' },
      { expr: '[First Name] & " " & [Last Name]', as: 'Recorded By' },
      { expr: 'DateDiff("d", [Date], Date())', as: 'Days Ago', criteria: ['Between 0 And 30'], show: false },
    ] },
    'Stock value by storage': { tables: ['Ingredients'], columns: [
      { table: 'Ingredients', field: 'Storage', total: 'group', sort: 'asc' }, { table: 'Ingredients', field: 'ID', as: 'Ingredients', total: 'count' },
      { expr: 'Round([In Stock] * [Cost per Unit], 2)', as: 'Value at Cost', total: 'sum' },
    ] },
  },
  // The print templates: one document for each record of the table named, its lines from the records pointing at it.
  documents: [
    { name: 'Receipt', table: 'Orders', paper: 'receipt-80',
      letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'Tax Number'] },
      blocks: [
        { type: 'heading', text: 'Receipt {Number}' },
        { type: 'fields', columns: 2, items: [{ field: 'When', label: 'Date and time' }, { field: 'Service' }, { field: 'Table No', label: 'Table' }, { field: 'Payment', label: 'Paid by' }] },
        { type: 'lines', table: 'Order Lines', columns: [
          { field: 'Quantity', label: 'Qty', type: 'integer' }, { field: 'Menu Item.Name', label: 'Item' }, { field: 'Unit Price', label: 'Each', type: 'currency' },
          { expr: LINE, as: 'Price', type: 'currency' }, { expr: 'Round([Tax Rate] * 100, 0)', as: 'VAT %', type: 'integer' },
          { expr: VAT_IN, as: 'VAT', type: 'currency' },
        ] },
        { type: 'totals', rows: [
          { label: 'Total', expr: '[Total]', strong: true }, { label: 'Net of VAT', expr: '[Net]' }, { label: 'VAT included', expr: '[Tax]' },
        ] },
        { type: 'text', text: 'You were served by {Staff.First Name}. Please keep this receipt as your proof of purchase.' },
        { type: 'text', text: 'VAT number {Business.Tax Number}. {Business.Document Footer}', small: true },
      ] },
    { name: 'Purchase order', table: 'Purchases', paper: 'a4',
      letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail'] },
      blocks: [
        { type: 'heading', text: 'Purchase order {Number}' },
        { type: 'fields', columns: 2, items: [
          { field: 'Supplier.Name', label: 'To' }, { field: 'Number', label: 'Order number' }, { field: 'Supplier.Address', label: 'Address' },
          { field: 'Ordered On', label: 'Order date' }, { field: 'Supplier.City', label: 'City' }, { field: 'Due On', label: 'Deliver on' },
          { field: 'Supplier.Postcode', label: 'Postcode' }, { field: 'Supplier.Contact', label: 'For the attention of' },
        ] },
        { type: 'lines', table: 'Purchase Lines', columns: [
          { field: 'Ingredient.Name', label: 'Ingredient' }, { field: 'Ingredient.Bought As', label: 'Comes as' }, { field: 'Quantity', type: 'number' },
          { field: 'Ingredient.Unit', label: 'Unit' }, { field: 'Unit Cost', label: 'Price per unit', type: 'currency' },
          { expr: 'Round([Quantity] * [Unit Cost], 2)', as: 'Amount', type: 'currency' },
        ] },
        { type: 'totals', rows: [{ label: 'Order total', expr: '[Amount]', strong: true }] },
        { type: 'text', text: 'Please deliver to {Business.Name}, {Business.Address}, {Business.City} {Business.Postcode} on {Due On}, before 7 in the morning if you can. Quote {Number} on your delivery note and your invoice. Prices are those we last paid; tell us before delivery if any has changed.' },
        { type: 'text', text: 'Questions about this order: {Business.Phone} or {Business.E-mail}.', small: true },
        { type: 'signatures', labels: ['Ordered by', 'Received by'] },
      ] },
    { name: 'Payslip', table: 'Payslips', paper: 'a4',
      letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Company Number'] },
      blocks: [
        { type: 'heading', text: 'Payslip {Number}' },
        { type: 'fields', columns: 2, items: [
          { field: 'Staff', label: 'Employee' }, { field: 'Week Beginning', label: 'Week beginning' }, { field: 'Staff.Role', label: 'Role' },
          { field: 'Week Ending', label: 'Week ending' }, { field: 'Hourly Rate', label: 'Hourly rate' }, { field: 'Pay Date', label: 'Paid on' },
        ] },
        { type: 'lines', table: 'Shifts', sort: 'Start', columns: [
          { expr: 'Format([Start], "yyyy-mm-dd")', as: 'Date', type: 'date' }, { expr: SHIFT_DAY, as: 'Day' }, { expr: 'Mid([Start], 12, 5)', as: 'From' }, { expr: 'Mid([End], 12, 5)', as: 'To' },
          { field: 'Break Minutes', label: 'Break (min)', type: 'integer' }, { field: 'Station' }, { field: 'Status' },
          { field: 'Paid Hours', label: 'Hours', type: 'number' }, { field: 'Payslip.Hourly Rate', as: 'Rate', type: 'currency' },
          { expr: 'Round([Paid Hours] * [Rate], 2)', as: 'Pay', type: 'currency' },
        ] },
        { type: 'totals', rows: [
          { label: 'Hours paid', expr: '[Hours]', type: 'number' }, { label: 'Gross pay', expr: '[Gross Pay]', strong: true },
        ] },
        { type: 'text', text: 'Gross pay for the hours on the rota, holiday included at the usual rate and sick days unpaid here. Tax, National Insurance, pension and any other deductions are worked out by the payroll and shown on the payroll\'s own statement of net pay.', small: true },
      ] },
  ],
  reports: [
    { name: 'Menu costing sheet', query: 'Menu costing', group: 'Category', columns: ['Menu Item', { field: 'Price', type: 'currency' }, { field: 'Food Cost', type: 'currency' }, { field: 'Food Cost %', type: 'number' }, { field: 'Gross Profit', type: 'currency' }, { field: 'Gross Margin %', type: 'number' }], totals: [{ field: 'Food Cost %', fn: 'avg' }, { field: 'Gross Profit', fn: 'avg' }] },
    { name: 'Recipe cards', query: 'Recipe card', group: 'Menu Item', columns: [{ field: 'Line', type: 'integer' }, 'Ingredient', { field: 'Quantity', type: 'number' }, 'Unit', { field: 'Line Cost', type: 'currency' }], totals: [{ field: 'Line Cost', fn: 'sum' }] },
    { name: 'Price list', title: 'Menu', query: 'Menu', group: 'Category', columns: ['Menu Item', 'Diet', { field: 'Price', type: 'currency' }] },
    { name: 'Weekly takings and wages', query: 'Takings and wages by week', columns: [{ field: 'Week Beginning', type: 'date' }, { field: 'Days Open', type: 'integer' }, { field: 'Orders', type: 'integer' }, { field: 'Covers', type: 'integer' }, { field: 'Takings', type: 'currency' }, { field: 'Wages', type: 'currency' }, { field: 'Left After Wages', type: 'currency' }, { field: 'Average Day Wage %', type: 'number' }], totals: [{ field: 'Takings', fn: 'sum' }, { field: 'Wages', fn: 'sum' }, { field: 'Left After Wages', fn: 'sum' }] },
    { name: 'Monthly takings', query: 'Takings by month', columns: ['Month', { field: 'Days Open', type: 'integer' }, { field: 'Orders', type: 'integer' }, { field: 'Takings', type: 'currency' }, { field: 'VAT', type: 'currency' }, { field: 'Takings less VAT', type: 'currency' }, { field: 'Wages', type: 'currency' }, { field: 'Left After Wages', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }, { field: 'VAT', fn: 'sum' }, { field: 'Takings less VAT', fn: 'sum' }, { field: 'Wages', fn: 'sum' }, { field: 'Left After Wages', fn: 'sum' }] },
    { name: 'Rota', title: 'Rota for the next 7 days', query: 'Rota for the next 7 days', group: 'Day', columns: ['From', 'To', 'Name', 'Station', { field: 'Paid Hours', type: 'number' }, 'Status'], totals: [{ field: 'Paid Hours', fn: 'sum' }] },
    { name: 'Order list for suppliers', query: 'Ingredients to order', group: 'Supplier', columns: ['Ingredient', 'Unit', { field: 'In Stock', type: 'number' }, { field: 'On Order', type: 'number' }, { field: 'Reorder Level', type: 'number' }, { field: 'Order', type: 'number' }, 'Bought As', { field: 'Estimated Cost', type: 'currency' }], totals: [{ field: 'Estimated Cost', fn: 'sum' }] },
    { name: 'Waste report', query: 'Waste in the last 30 days', group: 'Reason', columns: [{ field: 'Date', type: 'date' }, 'Ingredient', { field: 'Quantity', type: 'number' }, 'Unit', { field: 'Cost', type: 'currency' }, 'Recorded By'], totals: [{ field: 'Cost', fn: 'sum' }] },
    { name: 'Bills to pay', query: 'Bills to pay', group: 'Supplier', columns: ['Purchase', 'Invoice Number', { field: 'Received On', type: 'date' }, { field: 'Days to Pay', type: 'integer' }, { field: 'Amount', type: 'currency' }], totals: [{ field: 'Amount', fn: 'sum' }] },
  ],
};

export const about = {
  name: 'Little Fig Cafe',
  kind: 'Cafe or restaurant',
  line: 'Menu, recipes and food costs, till orders, rota, deliveries, waste and the nightly cash-up for a small cafe.',
  summary: [
    `For the owner or manager of a small cafe who wants to see what each dish costs to make, what sells and when, whether the wages are in hand against the takings, what to order from each supplier and what is being thrown away.`,
    `Little Fig Cafe is an invented all-day cafe with about thirty seats and a takeaway counter. It opens every day but Christmas Day, from half past seven (half past eight at weekends) until half past four. Breakfast is served until half past eleven and lunch until half past three; coffee, cold drinks and cakes all day. Its head baker makes the cakes on site each morning. Eight people work there, and six invented suppliers deliver coffee, dairy, produce, meat and fish, bread and dry goods, each on its own days.`,
    `The records cover the last fourteen months: a cash-up sheet for every day the cafe opened, the rota as it was worked, every delivery and its bill, and what was thrown away. The till's orders are here line by line for the last eight weeks, about six thousand of them; older days are in their cash-up sheets. Trade is busier at weekends, in July and August and on sunny days, and quieter in January. The rota runs two weeks ahead. The records stop at the start of today: this morning's deliveries are not booked in yet and this morning's orders to suppliers are not placed, so the order list shows what to ask for.`,
    `Every dish has a recipe in its ingredients' own units, so its food cost comes from the ingredients' current prices. Prices are pounds, as a cafe in the UK would charge them, but nothing in the design depends on the currency.`,
    `The menu prices include VAT, as a cafe's prices to the public do. The sample keeps the VAT simple and says so: everything eaten in, every drink and any hot food taken away carries the standard rate of 20%, and cold food taken away (the cakes and bakes, the cold bagel and rolls) carries none. The real rules have more cases than that, such as hot food kept warm, cold drinks and snacks, so set each dish's Takeaway VAT Rate to what applies to you. Each line of a bill keeps the rate it was sold at, each bill keeps its VAT, and each day's cash-up keeps the day's VAT for the return.`,
    `It prints three documents: a till receipt for every bill, numbered RCP-2026-00042; a purchase order for every order to a supplier, numbered PO-2026-0042, made each morning from the order list; and a weekly payslip for each person, numbered PAY-2026-0042, from the shifts on the rota. The payslip shows hours, rate and gross pay only: tax, National Insurance, pension and other deductions are left for the payroll to work out.`,
  ].join('\n\n'),
  suits: [
    'A coffee shop or kiosk',
    'A small restaurant or bistro',
    'A bakery with a few tables',
    'A food truck or market stall',
    'A sandwich bar or takeaway',
    'A tea room or farm shop cafe',
  ],
  tables: {
    Business: BUSINESS_ABOUT,
    'Menu Categories': { about: 'The sections of the menu, such as breakfast or cakes. Enter these first.', fields: {
      'Name': 'The section as the menu heads it.',
      'Served': 'When the section is served, such as All day or Until 11:30, printed on the menu.',
      'Sort Order': 'The place of the section on the printed menu and the costing sheet: 1 comes first.',
    } },
    'Menu Items': { about: 'One record per dish or drink on the menu, past or present.', fields: {
      'Name': 'The dish as the menu and the till name it. No two dishes share a name.',
      'Category': 'The menu section it belongs to.',
      'Price': 'What a customer pays for one, as the menu shows it.',
      'Takeaway': 'Yes when it can be taken away as well as eaten in.',
      'Baked Here': 'Yes for what the cafe bakes itself, such as cakes and pancakes.',
      'Diet': 'Short marks for the menu: V for vegetarian, VG for vegan, GF for gluten free.',
      'Takeaway VAT Rate': 'The VAT rate when it is taken away, as a fraction: 0.2 for drinks and hot food, 0 for cold food such as cakes. Eaten in, everything carries the standard rate.',
      'Active': 'No for a dish taken off the menu. It stays for the orders that sold it, and drops out of the menu and the costing.',
    } },
    'Suppliers': { about: 'The firms the cafe buys from, one record each.', fields: {
      'Name': 'The supplier\'s trading name.',
      'Supplies': 'What the cafe buys from them, in a few words.',
      'Address': 'Their street address, printed on the purchase order.',
      'City': 'Their town or city.',
      'Postcode': 'Their postcode.',
      'Contact': 'The person to speak to when ordering.',
      'Phone': 'The number to call an order through.',
      'E-mail': 'Where to send an order in writing.',
      'Delivers On': 'The days of the week they deliver. An order goes in the day before.',
      'Payment Days': 'How many days after a delivery its bill is due.',
    } },
    'Ingredients': { about: 'Everything that goes into a dish or is sold as it comes, with its price, where it is kept and how much is in.', fields: {
      'Name': 'The ingredient as the kitchen knows it.',
      'Supplier': 'Who it is ordered from.',
      'Unit': 'The unit its cost, stock and recipe quantities are counted in: kg, litre or each.',
      'Cost per Unit': 'What one unit costs now, from the latest bill. Changing it changes every dish\'s food cost.',
      'Bought As': 'How it comes, such as a 1 kg bag or a tray of 30, to help when ordering.',
      'Storage': 'Where it is kept: fridge, dry store or the counter. Used for the stock count and the stock value.',
      'In Stock': 'How much is in, in its unit. Set it from a stock count; Data Files does not take sales off it by itself.',
      'On Order': 'How much is ordered and not yet delivered. Clear it when the delivery comes in.',
      'Reorder Level': 'When stock and what is on order come to this or less, it goes on the order list.',
      'Reorder Quantity': 'The usual amount to order, in its unit.',
    } },
    'Recipe Lines': { about: 'What goes into one portion of each dish: one line per ingredient. These give every dish its food cost.', fields: {
      'Menu Item': 'The dish this line belongs to.',
      'Line': 'The line\'s place in the recipe, from 1. Every dish needs a line 1: the costing counts the price once, on that line.',
      'Ingredient': 'The ingredient used.',
      'Quantity': 'How much of it one portion uses, in the ingredient\'s unit: 0.018 kg of beans for a shot of espresso.',
    } },
    'Staff': { about: 'One record per person on the payroll.', fields: {
      'First Name': 'Their first name.',
      'Last Name': 'Their family name.',
      'Role': 'Their job, such as Head Baker or Barista.',
      'Hourly Rate': 'What they are paid an hour before tax. The wage cost of each shift is worked out from it.',
      'Phone': 'A number to reach them about shifts.',
      'E-mail': 'Where the rota is sent.',
      'Started': 'The day they started.',
    } },
    'Payslips': { about: 'One payslip per person per week, Monday to Sunday, made once the week has been worked. Its lines are the week\'s shifts.', fields: {
      'Number': 'The payslip\'s number, given in date order with a new run each year: PAY-2026-0042.',
      'Staff': 'Who it pays.',
      'Week Beginning': 'The Monday the week starts.',
      'Week Ending': 'The Sunday it ends.',
      'Pay Date': 'The day the money goes out: the Friday after the week ends.',
      'Hours': 'The hours paid that week, holiday included, added up from the shifts.',
      'Hourly Rate': 'The rate paid that week, kept so a later pay rise leaves old payslips as they were.',
      'Gross Pay': 'Pay before any deduction: each shift\'s hours times the rate, rounded to the penny, added up. Tax, National Insurance and pension are left for the payroll.',
    } },
    'Shifts': { about: 'The rota: one record per person per shift, planned ahead and kept once worked.', fields: {
      'Staff': 'Who works the shift.',
      'Start': 'The day and time the shift starts.',
      'End': 'The day and time it ends, later than planned when a busy day kept someone on.',
      'Break Minutes': 'The unpaid break in the shift.',
      'Paid Hours': 'The hours paid: from start to end less the break, or nought for a day off sick.',
      'Station': 'Where they work that shift: Barista, Kitchen, Floor, Bakery or Manager.',
      'Status': 'Rota for a shift still to come, Worked, Holiday (paid, not worked) or Sick.',
      'Payslip': 'The payslip the shift is paid on; empty until its week is over and paid.',
    } },
    'Orders': { about: 'One record per bill rung up on the till: a table\'s meal or a takeaway. Each prints as a receipt.', fields: {
      'Number': 'The receipt number, given in order through the day with a new run each year: RCP-2026-00042. A voided bill keeps its number.',
      'When': 'The day and time it was rung up.',
      'Service': 'Dine in or Takeaway.',
      'Table No': 'The table it was served at; empty for a takeaway.',
      'Covers': 'How many people sat down to it; nought for a takeaway.',
      'Staff': 'Who took the order.',
      'Payment': 'Card or Cash.',
      'Net': 'The bill less the VAT within it.',
      'Tax': 'The VAT within the bill: each line\'s VAT worked out from its price and rate, rounded to the penny, then added up.',
      'Total': 'What the bill came to, VAT included: the sum of its lines.',
      'Voided': 'Yes for a bill cancelled before it was paid. It stays on record and is left out of the takings.',
    } },
    'Order Lines': { about: 'What was on each bill: one line per dish, with how many.', fields: {
      'Order': 'The bill this line is on.',
      'Menu Item': 'The dish or drink sold.',
      'Quantity': 'How many.',
      'Unit Price': 'The price of one when it was sold, VAT included, kept so a later price change leaves old bills as they were.',
      'Tax Rate': 'The VAT rate within the price, as a fraction: the standard 0.2 eaten in, the dish\'s Takeaway VAT Rate taken away.',
    } },
    'Purchases': { about: 'One record per order to a supplier, from the day it was placed to the day its bill was paid. Each prints as a purchase order.', fields: {
      'Number': 'The purchase order\'s number, given in date order with a new run each year: PO-2026-0042. The supplier quotes it on the delivery note.',
      'Supplier': 'Who it was ordered from.',
      'Ordered On': 'The day the order went in.',
      'Due On': 'The day the delivery is expected.',
      'Received On': 'The day it arrived; empty while it is still to come.',
      'Invoice Number': 'The supplier\'s number on the bill that came with the delivery.',
      'Paid On': 'The day the bill was paid; empty while it is owed.',
    } },
    'Purchase Lines': { about: 'What was ordered on each purchase: one line per ingredient.', fields: {
      'Purchase': 'The purchase this line is on.',
      'Ingredient': 'The ingredient ordered.',
      'Quantity': 'How much, in the ingredient\'s unit.',
      'Unit Cost': 'What one unit cost on this bill. Prices rose a little over the year, so older bills are cheaper.',
    } },
    'Waste': { about: 'Everything thrown away, as the kitchen notes it on the waste sheet.', fields: {
      'Ingredient': 'What was thrown away.',
      'Date': 'The day.',
      'Quantity': 'How much, in the ingredient\'s unit.',
      'Reason': 'Why: out of date, overripe, stale, spilt, dropped, a kitchen error and so on.',
      'Recorded By': 'Who wrote it down.',
    } },
    'Cash Ups': { about: 'The sheet filled in at close each day the cafe opens: the till\'s reading, the cash counted, and the day\'s wages from the rota.', fields: {
      'Date': 'The day. One sheet per day.',
      'Closed By': 'Who counted up and locked up.',
      'Weather': 'Sunny, Cloudy, Rain or Cold: worth noting, as it moves trade.',
      'Orders': 'How many bills the till rang up that day, voided ones left out.',
      'Covers': 'How many people sat down.',
      'Card Takings': 'Taken by card, from the till\'s end of day reading.',
      'Cash Takings': 'Taken in cash, from the same reading.',
      'VAT': 'The VAT within the day\'s takings, from the same reading, for the VAT return.',
      'Cash Counted': 'The cash actually in the drawer, less the float. A difference from Cash Takings is the till over or short.',
      'Wage Cost': 'What the day\'s shifts cost in wages, from the rota: hours worked or on holiday times each person\'s rate.',
      'Notes': 'Anything worth remembering about the day: something ran out, the card machine went down, the till was out.',
    } },
  },
  queries: {
    'Menu costing': 'Each dish on the menu with its price, food cost, food cost as a share of the price, and the gross profit and margin left over. Use it when setting prices or when a supplier puts a price up: most cafes aim for a food cost under about 30% of the price.',
    'Recipe card': 'Every line of every recipe with what it costs, for checking a recipe or showing a new cook how a dish is made up.',
    'Menu': 'The dishes on the menu now, by section, with their diet marks and prices.',
    'Best sellers, last 30 days': 'How many of each dish sold in the last 30 days and what they took, most sold first. Use it to plan the bake and the prep, and to spot dishes worth dropping.',
    'Sales by hour': 'Bills, takings and the average bill for each hour of the day over the last eight weeks: when the rush comes and when the cafe is quiet, for setting shift times.',
    'Sales by day of the week': 'Bills, covers, takings and the average bill for each day of the week over the last eight weeks, for planning who works when.',
    'Average spend by service': 'For the last 30 days, how many bills were eaten in and how many taken away, what each brought in and the average bill for each.',
    'Takings and wages by week': 'Each week\'s takings from the cash-up sheets against the wages on the rota, what was left after wages, and the average day\'s wages as a share of its takings. The week starts on Monday; the newest week comes first.',
    'Takings by month': 'Takings, the VAT within them and the takings less VAT, and wages, for each month, oldest first: to see the seasons, compare with the same month last year, and fill in the VAT return.',
    'Takings by weather': 'The average day\'s takings and covers in the warmer months (May to September) by the weather noted on the cash-up sheet.',
    'Cash over or short': 'The days the cash counted at close did not match the till\'s cash reading, and by how much, with who closed up.',
    'Hours and wages by staff, last 28 days': 'Each person\'s shifts, hours and wages over the last four weeks, with holiday paid in and sick days left out. Use it for the payroll and to keep hours fair.',
    'Rota for the next 7 days': 'Who is on, when and where, today and for the next six days, earliest first.',
    'Ingredients to order': 'Ingredients at or below their reorder level with nothing on order, by supplier, with the usual amount to order and what it will cost. Work through it each morning before phoning the orders in.',
    'Deliveries due': 'Orders placed with suppliers and not yet delivered, the earliest due first, with how many lines and their value.',
    'Bills to pay': 'Delivered orders whose bills are not yet paid, with the days left until each is due; below nought means it is overdue.',
    'Spend by supplier by month': 'What was spent with each supplier each month, from the delivered orders.',
    'Waste by ingredient, last 90 days': 'For each ingredient thrown away in the last 90 days, how often, how much and what it cost, the dearest first. Use it to order less of what goes off.',
    'Waste in the last 30 days': 'Each entry on the waste sheet in the last 30 days with its cost and who recorded it.',
    'Stock value by storage': 'What the stock in each store is worth at today\'s cost, for the month end stock count and the accounts.',
  },
  reports: {
    'Menu costing sheet': 'Every dish on the menu by section with its price, food cost, food cost percentage, gross profit and margin, and the average for each section. For the owner when pricing the menu.',
    'Recipe cards': 'A card for each dish listing its ingredients, the quantity of each and its cost, with the dish\'s total food cost. For the kitchen wall and for training.',
    'Price list': 'The menu by section with prices and diet marks, ready to check against the printed menu.',
    'Weekly takings and wages': 'One line a week: days open, bills, covers, takings, wages, what was left after wages and the wage share, with the totals. For the owner\'s weekly look at the business.',
    'Monthly takings': 'Takings, VAT, takings less VAT and wages month by month with the totals, for the accountant, the VAT return or the bank.',
    'Rota': 'The coming week\'s shifts, a section for each day with its total hours, to pin up in the back.',
    'Order list for suppliers': 'What to order today, a section for each supplier with its delivery days, the quantities and the estimated cost. For whoever phones the orders in.',
    'Waste report': 'The last 30 days of waste grouped by reason, with the cost of each and in all. For the weekly kitchen meeting.',
    'Bills to pay': 'Unpaid supplier bills grouped by supplier, with their invoice numbers, the days left to pay and the total owed to each.',
  },
  documents: {
    Receipt: 'The till receipt for a bill, on an 80 mm roll: the cafe\'s name, address and VAT number, the receipt number, time, eat in or take away and the table, each item with its price and VAT rate, then the total, the net and the VAT within it, and how it was paid. Numbered RCP-2026-00042. Every bill on record prints, voided ones too, so choose the bill to print.',
    'Purchase order': 'The order sent to a supplier each morning from the order list: what to send, how much, at the last price paid, and the day to deliver, with the cafe\'s address to deliver to. Numbered PO-2026-0042, which the supplier quotes on the delivery note and the bill. Every purchase prints, delivered ones too, so choose the ones placed today.',
    Payslip: 'The weekly payslip for one person: the week, the hourly rate, each shift worked or on holiday with its hours and pay, and the gross pay. Numbered PAY-2026-0042. It shows gross pay only; tax, National Insurance, pension and other deductions are for the payroll to work out, and its statement of net pay goes with this one.',
  },
  tutorial: [
    { title: 'Ring up a bill and print the receipt', steps: [
      'Open the Orders form and press New. Type the next receipt number in Number, one on from the last (RCP-2026-00042 after RCP-2026-00041), and the date and time in When.',
      'Set Service to Dine in or Takeaway. For a table, type the Table No and the Covers; for a takeaway leave Table No empty and Covers at nought. Choose who took it in Staff and how it was paid in Payment.',
      'In the Order Lines below, choose each dish in Menu Item, type the Quantity and its menu price in Unit Price. Set Tax Rate to 0.2 eaten in, or the dish\'s Takeaway VAT Rate (on the Menu Items form) when it is taken away.',
      'Type the bill\'s Total, the VAT within it in Tax and the rest in Net. A form does not add the lines up by itself, so check the Total against the lines.',
      'Open the print template Receipt from Print templates, choose the bill, and press Print. The receipt prints the stored Total, Net and VAT under the lines.',
    ] },
    { title: 'Order from the suppliers in the morning', steps: [
      'Run the query Ingredients to order, or print the report Order list for suppliers: what is at or below its reorder level with nothing on order, a section for each supplier with its delivery days and the usual amount.',
      'For each supplier on the list, open the Purchases form and press New. Type the next number in Number (PO-2026-0042), choose the Supplier, type today in Ordered On and the next delivery day in Due On.',
      'In the Purchase Lines below, choose each Ingredient and type the Quantity in its unit and the Unit Cost from the last bill.',
      'Open the Ingredients form and type the same amount in On Order for each ingredient ordered, so it leaves the order list.',
      'Open the print template Purchase order from Print templates, choose today\'s purchases, and print them or press Save as PDF and e-mail each to the address on its Suppliers record.',
    ] },
    { title: 'Book in a delivery and pay the bill', steps: [
      'When the delivery comes, open the Purchases form, find it by its Number (the query Deliveries due lists the ones expected), and type today in Received On and the supplier\'s bill number in Invoice Number.',
      'Check the goods against the printed purchase order and sign it. Open the Ingredients form and add what came to In Stock for each line, and set its On Order back to nought.',
      'Each week, run the query Bills to pay or print the report Bills to pay: the bills not yet paid, with the days left to pay; below nought is overdue.',
      'When you pay one, type the day in Paid On on the Purchases form, and it leaves the list.',
    ] },
    { title: 'Plan the rota and pay the week', steps: [
      'Open the Shifts form and press New for each shift in the coming weeks: choose the Staff, type Start and End, the Break Minutes, the Paid Hours (end less start less the break) and the Station, and leave Status at Rota. Print the report Rota to pin up.',
      'After each day, change each shift\'s Status to Worked, Holiday or Sick, and its End if someone stayed on. A sick day has Paid Hours of nought.',
      'When the week is over (Monday to Sunday), run the query Hours and wages by staff, last 28 days to see the hours. Open the Payslips form and press New for each person: the next Number (PAY-2026-0042), the Staff, the Week Beginning and Week Ending, the Pay Date, their Hourly Rate, and the week\'s Hours and Gross Pay.',
      'Open the Shifts form as a datasheet and set Payslip to the new payslip on each of that person\'s shifts in the week: they are the payslip\'s lines.',
      'Open the print template Payslip from Print templates, choose the week\'s payslips, and print them. Give them out with the payroll\'s statement of net pay, which has the tax and deductions.',
    ] },
    { title: 'Cash up at close', steps: [
      'Take the till\'s end of day reading. Open the Cash Ups form and press New: the Date, who closed in Closed By, the Weather, the Orders and Covers, Card Takings, Cash Takings and the VAT from the reading.',
      'Count the cash in the drawer less the float and type it in Cash Counted. Type the day\'s Wage Cost from the rota, and anything worth remembering in Notes.',
      'Run the query Cash over or short now and then: the days the cash counted did not match the reading.',
      'On Monday, print the report Weekly takings and wages to see last week\'s takings against the wages.',
    ] },
    { title: 'Cost the menu when a price goes up', steps: [
      'When a supplier\'s bill shows a new price, open the Ingredients form and change Cost per Unit. Every dish that uses it costs more at once.',
      'Print the report Menu costing sheet: each dish\'s price, food cost and margin by section. Most cafes aim to keep the food cost under about 30% of the price.',
      'To change a price, open the Menu Items form and change Price. Old bills keep the price they were sold at, so past takings do not change.',
      'For a new dish, press New on the Menu Items form, then add its Recipe Lines from Line 1, one ingredient each. Print the report Recipe cards for the kitchen.',
    ] },
    { title: 'Close the month', steps: [
      'Print the report Monthly takings: each month\'s takings, the VAT within them, the takings less VAT and the wages. Add the VAT of the three months of your VAT quarter for the return.',
      'Count the stock in each store, put the counts in In Stock on the Ingredients form, and run the query Stock value by storage for the accounts.',
      'Print the report Waste report for the kitchen meeting, and run Waste by ingredient, last 90 days to see what to order less of.',
      'Run Spend by supplier by month to see what each supplier was paid.',
    ] },
  ],
  notes: [
    'The till\'s bills, Orders and Order Lines, go back eight weeks; the Cash Ups go back the whole fourteen months and are where the weekly and monthly figures come from. If your till keeps its own records you can skip Orders and keep only the cash-ups.',
    'Nothing here takes sales off the stock by itself. In Stock is what the last count, the deliveries and the waste sheet say, so count the stock now and then and correct it.',
    'A bill\'s Total, Net and Tax, and a payslip\'s Hours and Gross Pay, are stored figures, typed when the record is made, not added up from the lines by the form. The sample\'s own are exactly the sums of their lines, rounded line by line as a till rounds them.',
    'The VAT is kept simple: 20% on everything eaten in, on drinks and on hot food taken away; nought on cold food taken away. Set each dish\'s Takeaway VAT Rate to what the rules say for you, and have your VAT checked by an accountant.',
    'A print template prints every record of its table unless you choose some: Receipt prints every bill (voided ones too), Purchase order every purchase, Payslip every payslip. Choose the ones you want before pressing Print.',
    'The records stop at the start of today: this morning\'s deliveries are not booked in and this morning\'s orders not yet placed, so the order list shows what to ask for today. The rota runs two weeks ahead, and the week now running has no payslips yet.',
    'The VAT number, company number and bank details on the Business record are placeholders of noughts and Xs. Put in your own before printing anything for real.',
  ],
  faq: [
    { q: 'How do I reprint a customer\'s receipt from last week?', a: 'Find the bill on the Orders form by its time in When or its Number, then open the print template Receipt, choose that bill and press Print. Bills older than eight weeks are not in the sample; in your own file every bill you keep can be reprinted.' },
    { q: 'Why does a carrot cake have VAT eaten in but none taken away?', a: 'In the UK, food eaten on the premises carries the standard rate, while cold food taken away is zero-rated. Each line of a bill keeps the rate it was sold at in Tax Rate, from the dish\'s Takeaway VAT Rate when it went out the door. The real rules have more cases; the sample keeps to these two.' },
    { q: 'Where are tax and National Insurance on the payslip?', a: 'They are not worked out here. The payslip shows the hours from the rota, the rate and the gross pay; the payroll works out tax, National Insurance and pension from the gross pay and gives its own statement of net pay. You can add a Payslip Deductions table pointing at Payslips if you want them here.' },
    { q: 'I placed an order but the ingredient is still on Ingredients to order. Why?', a: 'The list leaves out what has something On Order. Type the amount ordered in On Order on the Ingredients form, and set it back to nought when the delivery is booked in.' },
    { q: 'If I put a price up, do last month\'s takings change?', a: 'No. Each bill line keeps its own Unit Price, so old bills and their totals stay as they were. Menu costing uses today\'s price and today\'s ingredient costs.' },
    { q: 'How do I find the VAT for my quarter\'s return?', a: 'Print the report Monthly takings and add up the VAT column for the three months of the quarter. It comes from the VAT on each day\'s cash-up, which is the till\'s reading.' },
    { q: 'Can I start with my own menu and keep the design?', a: 'Yes. Make the sample empty, fill in Business with your own details, then follow the order under starting empty: suppliers and ingredients first, then the menu and its recipes, then your staff.' },
  ],
  empty: `Start with Business: your own name, address, VAT number and the footer your receipts carry. Then Suppliers, then Ingredients with each one's unit, cost per unit, supplier and reorder level. Then Menu Categories, then Menu Items, then the Recipe Lines for each dish, numbering each recipe's lines from 1; Menu costing works as soon as a dish has its recipe. Add your Staff and plan the Shifts for the coming week. From then on, day to day: a Purchase and its lines when you order from a supplier, its Received On when it arrives and its Paid On when the bill is paid; the Waste sheet as things are thrown away; and a Cash Up at each close. If your till keeps its own records, the Cash Ups alone give the weekly and monthly figures; enter Orders and Order Lines only if you want best sellers and the busy hours here too. Keep In Stock up to date from a stock count, as Data Files does not take sales off it by itself.`,
  adapt: [
    'Measure the food cost against what the cafe keeps rather than what the customer pays: in a copy of Menu costing, divide by [Price] / 1.2 in place of [Price], so the VAT within the price is left out.',
    'Add a Packaging Cost field to Menu Items for the cup, lid or box a takeaway goes out in, and add it to the food cost in a copy of Menu costing to see what a takeaway really costs.',
    'Add yes/no allergen fields to Ingredients (gluten, milk, eggs, nuts and the rest) and a query that lists the dishes using any ingredient marked Yes, for the allergen folder.',
    'A bakery can add a Bakes table (date, dish, made, left over at close) to see how much of each cake to make on each day of the week.',
    'A food truck can drop Table No and Covers and add a Pitch field to Cash Ups, then group the takings by pitch to see which places pay.',
    'A restaurant can add a Bookings table (date, time, covers, name, phone, table) and a query for the bookings in the next 7 days.',
    'Add a Running Costs table (date, kind, amount) for rent, power and card fees, and a query by month beside Takings by month to see the profit.',
    'Add a Tips field to Cash Ups and a column summing it in Takings and wages by week, so tips are kept apart from the takings.',
  ],
};

// The suppliers: [key, name, what they supply, delivery weekdays (0 Sunday), payment days, invoice prefix,
// and 'standing' where the order is a standing one, topped up every delivery as bread is].
const SUPPLIERS = [
  ['roaster', 'Granite Hill Roasters', 'Coffee, tea and hot chocolate', [2], 30, 'GHR'],
  ['dairy', 'Juniper Vale Dairy', 'Milk, cream, butter, eggs, cheese and juice', [1, 2, 3, 4, 5, 6], 14, 'JVD'],
  ['produce', 'Alder Farm Produce', 'Fruit, vegetables and salad', [1, 3, 5, 6], 14, 'AFP'],
  ['butcher', 'Riverbend Meat and Fish', 'Bacon, sausages, ham, chicken and smoked fish', [2, 5], 14, 'RMF'],
  ['bakery', 'Thistle Street Bakery', 'Bread, bagels and rolls', [0, 1, 2, 3, 4, 5, 6], 7, 'TSB', 'standing'],
  ['dry', 'Larkspur Dry Goods', 'Flour, sugar, chocolate, nuts, oils, tins and bottled drinks', [4], 30, 'LDG'],
];
const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const TRADE_ESTATES = ['Canal Walk Trading Estate', 'Quarry Road Industrial Park', 'Riverside Drive Works', 'Mill Lane Business Centre', 'Station Road Yard'];

// The ingredients: [key, name, unit, cost per unit, supplier, storage, bought as, pack size in units].
const INGREDIENTS = [
  ['espresso', 'Espresso beans', 'kg', 23.5, 'roaster', 'Dry store', '1 kg bag', 1],
  ['decaf', 'Decaf espresso beans', 'kg', 26, 'roaster', 'Dry store', '1 kg bag', 1],
  ['tea', 'Breakfast tea bags', 'each', 0.04, 'roaster', 'Dry store', 'box of 250', 250],
  ['mint', 'Peppermint tea bags', 'each', 0.09, 'roaster', 'Dry store', 'box of 100', 100],
  ['chocpowder', 'Hot chocolate powder', 'kg', 9.8, 'roaster', 'Dry store', '1 kg tin', 1],
  ['milk', 'Whole milk', 'litre', 1.15, 'dairy', 'Fridge', '2 litre bottle', 2],
  ['oat', 'Oat drink', 'litre', 2.1, 'dairy', 'Dry store', '1 litre carton', 1],
  ['butter', 'Butter, unsalted', 'kg', 8.2, 'dairy', 'Fridge', '250 g block', 0.25],
  ['cream', 'Double cream', 'litre', 4.6, 'dairy', 'Fridge', '1 litre tub', 1],
  ['eggs', 'Free-range eggs', 'each', 0.32, 'dairy', 'Fridge', 'tray of 30', 30],
  ['cheddar', 'Mature cheddar', 'kg', 11, 'dairy', 'Fridge', '2.5 kg block', 2.5],
  ['creamcheese', 'Cream cheese', 'kg', 6.4, 'dairy', 'Fridge', '1 kg tub', 1],
  ['yoghurt', 'Greek-style yoghurt', 'kg', 3.9, 'dairy', 'Fridge', '1 kg tub', 1],
  ['halloumi', 'Halloumi', 'kg', 13, 'dairy', 'Fridge', '1 kg pack', 1],
  ['oj', 'Orange juice, fresh', 'litre', 2.6, 'dairy', 'Fridge', '1 litre bottle', 1],
  ['avocado', 'Avocados', 'each', 1.05, 'produce', 'Counter', 'loose', 1],
  ['tomatoes', 'Vine tomatoes', 'kg', 2.8, 'produce', 'Fridge', 'by the kg', 1],
  ['mushrooms', 'Chestnut mushrooms', 'kg', 4.2, 'produce', 'Fridge', 'by the kg', 1],
  ['spinach', 'Baby spinach', 'kg', 6.5, 'produce', 'Fridge', '500 g bag', 0.5],
  ['salad', 'Mixed salad leaves', 'kg', 7.8, 'produce', 'Fridge', '500 g bag', 0.5],
  ['lemons', 'Lemons', 'each', 0.3, 'produce', 'Counter', 'net of 10', 10],
  ['bananas', 'Bananas', 'kg', 1.4, 'produce', 'Counter', 'by the kg', 1],
  ['berries', 'Mixed berries', 'kg', 9.5, 'produce', 'Fridge', '500 g punnet', 0.5],
  ['onions', 'Red onions', 'kg', 1.2, 'produce', 'Dry store', 'by the kg', 1],
  ['carrots', 'Carrots', 'kg', 0.9, 'produce', 'Fridge', 'by the kg', 1],
  ['potatoes', 'Baking potatoes', 'kg', 0.95, 'produce', 'Dry store', 'by the kg', 1],
  ['bacon', 'Back bacon', 'kg', 11.5, 'butcher', 'Fridge', '1 kg pack', 1],
  ['sausages', 'Pork sausages', 'each', 0.65, 'butcher', 'Fridge', 'pack of 20', 20],
  ['salmon', 'Smoked salmon', 'kg', 32, 'butcher', 'Fridge', '500 g pack', 0.5],
  ['ham', 'Roast ham, sliced', 'kg', 14, 'butcher', 'Fridge', '1 kg pack', 1],
  ['chicken', 'Chicken breast', 'kg', 10.5, 'butcher', 'Fridge', '2 kg pack', 2],
  ['sourdough', 'Sourdough loaf', 'each', 3.6, 'bakery', 'Counter', 'loaf of 16 slices', 1],
  ['bagels', 'Plain bagels', 'each', 0.45, 'bakery', 'Counter', 'bag of 6', 6],
  ['ciabatta', 'Ciabatta rolls', 'each', 0.55, 'bakery', 'Counter', 'bag of 6', 6],
  ['flour', 'Plain flour', 'kg', 0.95, 'dry', 'Dry store', '16 kg sack', 16],
  ['sugar', 'Caster sugar', 'kg', 1.2, 'dry', 'Dry store', '5 kg bag', 5],
  ['icing', 'Icing sugar', 'kg', 1.5, 'dry', 'Dry store', '3 kg bag', 3],
  ['oats', 'Porridge oats', 'kg', 1.6, 'dry', 'Dry store', '3 kg bag', 3],
  ['darkchoc', 'Dark chocolate buttons', 'kg', 9.6, 'dry', 'Dry store', '2.5 kg bag', 2.5],
  ['cocoa', 'Cocoa powder', 'kg', 7.9, 'dry', 'Dry store', '1 kg bag', 1],
  ['walnuts', 'Walnut pieces', 'kg', 11, 'dry', 'Dry store', '1 kg bag', 1],
  ['almonds', 'Ground almonds', 'kg', 9.6, 'dry', 'Dry store', '1 kg bag', 1],
  ['bakingpowder', 'Baking powder', 'kg', 4.5, 'dry', 'Dry store', '500 g tub', 0.5],
  ['vanilla', 'Vanilla extract', 'litre', 38, 'dry', 'Dry store', '250 ml bottle', 0.25],
  ['oliveoil', 'Olive oil', 'litre', 6.5, 'dry', 'Dry store', '5 litre tin', 5],
  ['sunoil', 'Sunflower oil', 'litre', 2.1, 'dry', 'Dry store', '5 litre tin', 5],
  ['maple', 'Maple syrup', 'litre', 14, 'dry', 'Dry store', '1 litre bottle', 1],
  ['bakedbeans', 'Beans in tomato sauce', 'kg', 1.6, 'dry', 'Dry store', '2.6 kg tin', 2.6],
  ['lentils', 'Red lentils', 'kg', 2.1, 'dry', 'Dry store', '5 kg bag', 5],
  ['stock', 'Vegetable stock powder', 'kg', 6, 'dry', 'Dry store', '1 kg tub', 1],
  ['granola', 'Granola', 'kg', 5.5, 'dry', 'Dry store', '3 kg bag', 3],
  ['honey', 'Honey', 'kg', 7, 'dry', 'Dry store', '1 kg jar', 1],
  ['jam', 'Strawberry jam', 'kg', 3.8, 'dry', 'Dry store', '3 kg tub', 3],
  ['water', 'Sparkling water, 330 ml bottle', 'each', 0.45, 'dry', 'Dry store', 'case of 24', 24],
  ['lemonade', 'Cloudy lemonade, 330 ml can', 'each', 0.62, 'dry', 'Dry store', 'case of 24', 24],
];

// The menu categories: [name, served, sort order].
const CATEGORIES = [
  ['Coffee and Tea', 'All day', 1], ['Cold Drinks', 'All day', 2], ['Breakfast', 'Until 11:30', 3],
  ['Lunch', 'From 11:30', 4], ['Cakes and Bakes', 'All day', 5],
];

// The menu: [name, category, pool, price, takeaway, baked here, diet, weight, { summer, winter }, recipe].
// The pool is when it is chosen (hot or cold drink, breakfast, lunch, cake); the
// weight is how often against the rest of its pool; summer and winter lift or
// lower it in those months. A recipe is [ingredient, quantity in its unit].
const MENU = [
  ['Espresso', 'Coffee and Tea', 'hot', 2.4, true, false, 'VG', 4, {}, [['espresso', 0.018]]],
  ['Americano', 'Coffee and Tea', 'hot', 3.1, true, false, 'VG', 12, {}, [['espresso', 0.018]]],
  ['Flat white', 'Coffee and Tea', 'hot', 3.5, true, false, 'V', 14, {}, [['espresso', 0.018], ['milk', 0.15]]],
  ['Latte', 'Coffee and Tea', 'hot', 3.7, true, false, 'V', 16, { summer: 0.8 }, [['espresso', 0.018], ['milk', 0.25]]],
  ['Cappuccino', 'Coffee and Tea', 'hot', 3.6, true, false, 'V', 12, {}, [['espresso', 0.018], ['milk', 0.2], ['cocoa', 0.001]]],
  ['Oat latte', 'Coffee and Tea', 'hot', 4.1, true, false, 'VG', 8, {}, [['espresso', 0.018], ['oat', 0.25]]],
  ['Decaf latte', 'Coffee and Tea', 'hot', 3.8, true, false, 'V', 3, {}, [['decaf', 0.018], ['milk', 0.25]]],
  ['Hot chocolate', 'Coffee and Tea', 'hot', 3.8, true, false, 'V', 5, { summer: 0.4, winter: 1.8 }, [['chocpowder', 0.03], ['milk', 0.25], ['cream', 0.02]]],
  ['Pot of tea', 'Coffee and Tea', 'hot', 2.9, true, false, 'V', 9, {}, [['tea', 2], ['milk', 0.05]]],
  ['Peppermint tea', 'Coffee and Tea', 'hot', 2.8, true, false, 'VG', 3, {}, [['mint', 1]]],
  ['Iced latte', 'Coffee and Tea', 'hot', 4, true, false, 'V', 2, { summer: 5, winter: 0.2 }, [['espresso', 0.018], ['milk', 0.2]]],
  ['Fresh orange juice', 'Cold Drinks', 'cold', 3.2, true, false, 'VG', 5, {}, [['oj', 0.25]]],
  ['Sparkling water', 'Cold Drinks', 'cold', 1.9, true, false, 'VG', 4, {}, [['water', 1]]],
  ['Cloudy lemonade', 'Cold Drinks', 'cold', 2.4, true, false, 'VG', 3, { summer: 2 }, [['lemonade', 1]]],
  ['Banana and berry smoothie', 'Cold Drinks', 'cold', 4.8, true, false, 'V', 3, { summer: 1.6 }, [['bananas', 0.12], ['berries', 0.08], ['yoghurt', 0.1], ['milk', 0.1], ['honey', 0.01]]],
  ['Full breakfast', 'Breakfast', 'breakfast', 12.5, false, false, '', 14, {}, [['bacon', 0.06], ['sausages', 2], ['eggs', 2], ['mushrooms', 0.06], ['tomatoes', 0.08], ['bakedbeans', 0.12], ['sourdough', 0.125], ['butter', 0.01]]],
  ['Veggie breakfast', 'Breakfast', 'breakfast', 11.5, false, false, 'V', 8, {}, [['halloumi', 0.06], ['eggs', 2], ['mushrooms', 0.08], ['spinach', 0.04], ['tomatoes', 0.08], ['bakedbeans', 0.12], ['sourdough', 0.125], ['butter', 0.01]]],
  ['Avocado on sourdough', 'Breakfast', 'breakfast', 9.5, false, false, 'V', 10, {}, [['avocado', 1], ['sourdough', 0.125], ['eggs', 2], ['lemons', 0.25], ['oliveoil', 0.01], ['spinach', 0.02]]],
  ['Bacon bap', 'Breakfast', 'breakfast', 6.2, true, false, '', 9, {}, [['bacon', 0.08], ['ciabatta', 1], ['butter', 0.01]]],
  ['Smoked salmon bagel', 'Breakfast', 'breakfast', 8.9, true, false, '', 6, {}, [['salmon', 0.06], ['bagels', 1], ['creamcheese', 0.04], ['lemons', 0.1]]],
  ['Porridge with berries', 'Breakfast', 'breakfast', 6.5, true, false, 'V', 5, { summer: 0.5, winter: 1.6 }, [['oats', 0.07], ['milk', 0.25], ['berries', 0.06], ['maple', 0.02]]],
  ['Granola bowl', 'Breakfast', 'breakfast', 6.8, false, false, 'V', 4, { summer: 1.4 }, [['granola', 0.08], ['yoghurt', 0.15], ['berries', 0.06], ['honey', 0.015]]],
  ['Toast and jam', 'Breakfast', 'breakfast', 3.9, false, false, 'VG', 5, {}, [['sourdough', 0.125], ['butter', 0.015], ['jam', 0.03]]],
  ['Pancakes with bacon and maple', 'Breakfast', 'breakfast', 9.8, false, true, '', 7, {}, [['flour', 0.08], ['eggs', 1], ['milk', 0.12], ['butter', 0.01], ['bakingpowder', 0.003], ['bacon', 0.06], ['maple', 0.03]]],
  ['Soup of the day with sourdough', 'Lunch', 'lunch', 7.5, true, false, 'VG', 9, { summer: 0.5, winter: 1.6 }, [['lentils', 0.06], ['carrots', 0.1], ['onions', 0.05], ['stock', 0.01], ['oliveoil', 0.01], ['sourdough', 0.125]]],
  ['Ham and cheese toastie', 'Lunch', 'lunch', 7.9, true, false, '', 11, {}, [['sourdough', 0.25], ['ham', 0.06], ['cheddar', 0.05], ['butter', 0.015]]],
  ['Chicken and avocado ciabatta', 'Lunch', 'lunch', 9.5, true, false, '', 9, {}, [['chicken', 0.1], ['ciabatta', 1], ['avocado', 0.5], ['salad', 0.02], ['oliveoil', 0.01]]],
  ['Halloumi salad', 'Lunch', 'lunch', 10.5, false, false, 'V', 7, { summer: 1.6, winter: 0.7 }, [['halloumi', 0.1], ['salad', 0.08], ['tomatoes', 0.08], ['onions', 0.02], ['oliveoil', 0.02], ['lemons', 0.25]]],
  ['Cheese omelette with salad', 'Lunch', 'lunch', 8.9, false, false, 'V', 5, {}, [['eggs', 3], ['cheddar', 0.04], ['butter', 0.01], ['salad', 0.04]]],
  ['Jacket potato with cheese and beans', 'Lunch', 'lunch', 8.5, false, false, 'V', 6, { summer: 0.7, winter: 1.4 }, [['potatoes', 0.35], ['cheddar', 0.05], ['bakedbeans', 0.15], ['butter', 0.01], ['salad', 0.02]]],
  ['Chocolate brownie', 'Cakes and Bakes', 'cake', 3.4, true, true, 'V', 10, {}, [['darkchoc', 0.025], ['butter', 0.02], ['sugar', 0.025], ['eggs', 0.4], ['flour', 0.01], ['cocoa', 0.004]]],
  ['Carrot cake', 'Cakes and Bakes', 'cake', 4.2, true, true, 'V', 8, {}, [['carrots', 0.04], ['flour', 0.035], ['sugar', 0.03], ['eggs', 0.4], ['sunoil', 0.025], ['walnuts', 0.01], ['creamcheese', 0.03], ['icing', 0.02]]],
  ['Lemon drizzle cake', 'Cakes and Bakes', 'cake', 3.6, true, true, 'V', 7, {}, [['flour', 0.04], ['butter', 0.03], ['sugar', 0.04], ['eggs', 0.5], ['lemons', 0.15], ['icing', 0.01]]],
  ['Banana bread', 'Cakes and Bakes', 'cake', 3.2, true, true, 'V', 7, {}, [['bananas', 0.06], ['flour', 0.04], ['sugar', 0.025], ['butter', 0.02], ['eggs', 0.3], ['bakingpowder', 0.002], ['walnuts', 0.008]]],
  ['Scone with jam and cream', 'Cakes and Bakes', 'cake', 3.8, false, true, 'V', 6, { summer: 1.2 }, [['flour', 0.06], ['butter', 0.015], ['sugar', 0.01], ['milk', 0.03], ['bakingpowder', 0.003], ['jam', 0.025], ['cream', 0.03]]],
  ['Almond and berry cake', 'Cakes and Bakes', 'cake', 4, true, true, 'V, GF', 4, {}, [['almonds', 0.05], ['eggs', 0.6], ['sugar', 0.035], ['berries', 0.02], ['butter', 0.01]]],
  ['Chocolate chip cookie', 'Cakes and Bakes', 'cake', 2.5, true, true, 'V', 8, {}, [['flour', 0.025], ['butter', 0.015], ['sugar', 0.02], ['darkchoc', 0.015], ['eggs', 0.15], ['vanilla', 0.001]]],
  // Off the menu since the spring: sold earlier, still in the records, no longer offered.
  ['Tomato soup with sourdough', 'Lunch', 'lunch', 7, true, false, 'V', 8, {}, [['tomatoes', 0.25], ['onions', 0.05], ['stock', 0.01], ['cream', 0.03], ['sourdough', 0.125]], 'retired'],
  ['Egg and cress roll', 'Lunch', 'lunch', 5.9, true, false, 'V', 5, {}, [['eggs', 2], ['ciabatta', 1], ['salad', 0.01], ['butter', 0.01]], 'retired'],
];
const RETIRED_BEFORE = -150; // the two old dishes came off the menu about five months ago

// VAT, kept simple: everything eaten in, every drink and hot food taken away
// carry the standard rate; cold food taken away (cakes, cold rolls and
// bagels) carries none. The menu's prices include it.
const COLD_TO_TAKE_AWAY = new Set(['Smoked salmon bagel', 'Chicken and avocado ciabatta', 'Egg and cress roll']);
const STANDARD_RATE = 0.2;
const takeawayRate = (it) => (it.pool === 'cake' || COLD_TO_TAKE_AWAY.has(it.name) ? 0 : STANDARD_RATE);

// The staff: [role, hourly rate, station, [weekdays, from, to, break minutes, 'summer' for Jun to Aug only]...].
// Hours are in decimals: 15.5 is half past three.
const STAFF = [
  ['Owner and Manager', 16.5, 'Manager', [[[1, 2, 3, 4, 5], 7, 15.5, 30]]],
  ['Head Baker', 15.2, 'Bakery', [[[2, 3, 4, 5, 6], 6, 12, 0]]],
  ['Cook', 14.2, 'Kitchen', [[[1, 2, 3, 4, 5], 7.5, 15.5, 30]]],
  ['Weekend Cook', 13.9, 'Kitchen', [[[6, 0], 8, 16, 30]]],
  ['Lead Barista', 13.8, 'Barista', [[[1, 2, 3, 4, 5], 7, 15, 30]]],
  ['Barista', 12.9, 'Barista', [[[3, 4, 5, 6, 0], 10, 17, 30]]],
  ['Barista', 12.75, 'Barista', [[[6, 0], 8, 16.5, 30], [[1, 2], 10, 17, 30]]],
  ['Server', 12.75, 'Floor', [[[5, 6, 0], 9, 16.5, 30], [[4], 9, 16.5, 30, 'summer']]],
];

// Orders expected in each hour of a weekday and of a weekend day, before the
// day of the week, the season and the weather lift or lower them.
const WEEKDAY_HOURS = { 7: 5, 8: 13, 9: 13, 10: 12, 11: 10, 12: 15, 13: 14, 14: 9, 15: 7, 16: 2 };
const WEEKEND_HOURS = { 8: 7, 9: 17, 10: 19, 11: 18, 12: 19, 13: 18, 14: 13, 15: 9, 16: 3 };
const DAY_LIFT = [1, 0.88, 0.92, 0.97, 1, 1.1, 1.08]; // Sunday first
const SEASON = [0.82, 0.85, 0.93, 1, 1.05, 1.12, 1.22, 1.25, 1.05, 0.98, 0.95, 1.08]; // January first
const isSummer = (month) => month >= 5 && month <= 7;
const isWinter = (month) => month === 0 || month === 1 || month >= 10;

/** The day's weather, as the cash-up sheet notes it, and how much it lifts trade. */
function weatherFor(r, month) {
  if (isSummer(month)) return r.weighted([['Sunny', 45], ['Cloudy', 35], ['Rain', 20]]);
  if (isWinter(month)) return r.weighted([['Sunny', 15], ['Cloudy', 30], ['Rain', 30], ['Cold', 25]]);
  return r.weighted([['Sunny', 30], ['Cloudy', 40], ['Rain', 30]]);
}
const weatherLift = (weather, weekend) => ({ Sunny: weekend ? 1.15 : 1.06, Cloudy: 1, Rain: 0.88, Cold: 0.94 })[weather];

/** The dishes to choose from in `month`, by pool, weighted; with the retired dishes when `old`. */
function poolsFor(items, month, old) {
  const lists = { hot: [], cold: [], breakfast: [], lunch: [], cake: [], takeBreakfast: [], takeLunch: [], takeCake: [] };
  for (const it of items) {
    if (it.retired && !old) continue;
    const w = it.weight * (isSummer(month) ? it.season.summer ?? 1 : isWinter(month) ? it.season.winter ?? 1 : 1);
    lists[it.pool].push([it, w]);
    if (it.takeaway && it.pool === 'breakfast') lists.takeBreakfast.push([it, w]);
    if (it.takeaway && it.pool === 'lunch') lists.takeLunch.push([it, w]);
    if (it.takeaway && it.pool === 'cake') lists.takeCake.push([it, w]);
  }
  // Each pool as running totals, so a choice is one draw and a short walk.
  const pools = {};
  for (const [name, list] of Object.entries(lists)) {
    let sum = 0;
    pools[name] = { items: list.map(([it]) => it), upTo: list.map(([, w]) => (sum += w)), total: sum };
  }
  return pools;
}

/** An item of `pool` chosen by weight. */
function choose(r, pool) {
  const x = r.next() * pool.total;
  let i = 0;
  while (i < pool.upTo.length - 1 && x >= pool.upTo[i]) i++;
  return pool.items[i];
}
const COVERS = [[1, 28], [2, 46], [3, 12], [4, 10], [5, 2], [6, 2]];

/** One order at `t` hours (8.5 is half past eight): dine in or takeaway, its covers, and the dishes asked for. */
function oneOrder(r, t, weekend, month, pools) {
  const breakfast = t < 11.5;
  const lunch = !breakfast && t < 15.5; // the kitchen stops at half past three; cakes and drinks after
  const coldShare = isSummer(month) ? 0.25 : 0.1;
  const drink = () => choose(r, r.chance(coldShare) ? pools.cold : pools.hot);
  const items = [];
  let takeaway = r.chance((t < 9 ? 0.62 : t < 11.5 ? 0.45 : t < 14.5 ? 0.4 : 0.5) - (weekend ? 0.1 : 0));
  let covers = 0;
  if (takeaway) {
    const people = r.chance(0.15) ? 2 : 1;
    for (let p = 0; p < people; p++) {
      if (r.chance(0.94)) items.push(drink());
      if (r.chance(0.3)) {
        const mains = breakfast ? pools.takeBreakfast : lunch ? pools.takeLunch : null;
        items.push(mains && r.chance(0.55) ? choose(r, mains) : choose(r, pools.takeCake));
      }
    }
    if (!items.length) items.push(drink());
  } else {
    covers = r.weighted(COVERS);
    for (let p = 0; p < covers; p++) {
      if (r.chance(0.92)) items.push(drink());
      const main = breakfast ? r.chance(0.6) : lunch ? r.chance(t >= 14.5 ? 0.4 : 0.72) : false;
      if (main) items.push(choose(r, breakfast ? pools.breakfast : pools.lunch));
      if (r.chance(main ? 0.12 : t >= 14.5 ? 0.55 : 0.3)) items.push(choose(r, pools.cake));
    }
    if (!items.length) items.push(drink());
  }
  return { takeaway, covers, items };
}

/** A day's orders in time order: [{ h, m, takeaway, covers, items }]. */
function dayOrders(r, { month, wd, weather, pools }) {
  const weekend = wd === 0 || wd === 6;
  const hours = weekend ? WEEKEND_HOURS : WEEKDAY_HOURS;
  const lift = DAY_LIFT[wd] * SEASON[month] * weatherLift(weather, weekend);
  const out = [];
  for (const [hs, expected] of Object.entries(hours)) {
    const h = Number(hs);
    const n = Math.round(expected * lift * (0.75 + r.next() * 0.5));
    // The doors open at half past seven (half past eight at weekends); last orders a quarter past four.
    const first = h === (weekend ? 8 : 7) ? 30 : 0;
    const last = h === 16 ? 15 : 59;
    const minutes = Array.from({ length: n }, () => r.int(first, last)).sort((a, b) => a - b);
    for (const m of minutes) {
      const o = oneOrder(r, h + m / 60, weekend, month, pools);
      o.h = h;
      o.m = m;
      out.push(o);
    }
  }
  return out;
}

/** `x` rounded up to a sensible step for its unit. */
function niceUp(x, unit) {
  if (unit === 'each') return Math.max(1, Math.ceil(x));
  const step = x < 1 ? 0.05 : x < 10 ? 0.5 : 1;
  return Math.round(Math.ceil(x / step) * step * 1000) / 1000;
}
const r3 = (n) => Math.round(n * 1000) / 1000;
const clock = (x) => [Math.floor(x), Math.round((x - Math.floor(x)) * 60)];

/** Fill a Little Fig Cafe database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 425, orderDays = 56, ahead = 14 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };

  add('Business', businessRecord({
    name: 'Little Fig Cafe', domain: 'littlefig.example', phone: '020 7946 0418',
    address: { Address: '27 Orchard Way', City: 'London', Postcode: 'N16 7QR', Country: 'United Kingdom' },
    terms: 'Paid in full at the till.',
    footer: 'Thank you for coming in. Open every day but Christmas Day, 7:30 to 4:30 (8:30 at weekends). Prices include VAT. Please tell us about any allergy before you order.',
  }));

  // Document numbers, given in date order with a new run each year: RCP-2026-00042.
  const runs = {};
  const numbered = (prefix, d, width) => {
    const y = cal.at(d).getFullYear();
    const k = `${prefix}${y}`;
    runs[k] = (runs[k] || 0) + 1;
    return `${prefix}-${y}-${String(runs[k]).padStart(width, '0')}`;
  };

  // Suppliers, all in or near London.
  const suppliers = new Map();
  for (const [key, name, supplies, deliver, terms, prefix, standing] of SUPPLIERS) {
    const domain = `${name.toLowerCase().replace(/[^a-z]+/g, '')}.example`;
    const c = person(r, { domain });
    const id = add('Suppliers', {
      Name: name, Supplies: supplies, Address: `Unit ${r.int(1, 30)}, ${r.pick(TRADE_ESTATES)}`, City: 'London',
      Postcode: `${r.pick(['E9', 'E15', 'N17', 'N18', 'NW10', 'SE16'])} ${r.int(1, 9)}${r.pick(['AB', 'DF', 'GH', 'JL', 'NP', 'RT', 'WX'])}`,
      Contact: `${c['First Name']} ${c['Last Name']}`, Phone: phoneFor(r, 'United Kingdom'),
      'E-mail': `orders@${domain}`, 'Delivers On': deliver.length === 7 ? 'Every day' : deliver.map((w) => WEEKDAYS[w]).join(', '), 'Payment Days': terms,
    });
    // The longest wait between two deliveries, which sets how much to keep in.
    const sorted = [...deliver].sort((a, b) => a - b);
    const gap = Math.max(...sorted.map((w, i) => ((sorted[(i + 1) % sorted.length] - w + 7) % 7) || 7));
    suppliers.set(key, { id, deliver, terms, prefix, gap, standing: standing === 'standing', invoice: r.int(10000, 40000) });
  }

  // Menu categories and items; the ingredients come after, once their levels are known.
  const categories = new Map(CATEGORIES.map(([name, served, order]) => [name, add('Menu Categories', { Name: name, Served: served, 'Sort Order': order })]));
  const items = MENU.map(([name, cat, pool, price, takeaway, baked, diet, weight, season, recipe, retired]) => ({
    name, cat, pool, price, takeaway, baked, diet, weight, season, recipe, retired: retired === 'retired',
  }));

  // How much of each ingredient a day uses, from a trial week in each month run on
  // the side: the busiest month's daily use is what the levels are set from, so
  // they hold through the summer for salads and through the winter for soup.
  const trial = random(r.int(1, 2 ** 30));
  const used = new Map(INGREDIENTS.map(([key]) => [key, 0]));
  for (let month = 0; month < 12; month++) {
    const pools = poolsFor(items, month, true);
    const week = new Map();
    for (let wd = 0; wd < 7; wd++) {
      for (const o of dayOrders(trial, { month, wd, weather: 'Cloudy', pools })) {
        for (const it of o.items) for (let i = 0; i < it.recipe.length; i++) week.set(it.recipe[i][0], (week.get(it.recipe[i][0]) || 0) + it.recipe[i][1] / 7);
      }
    }
    for (const [k, q] of week) used.set(k, Math.max(used.get(k), q));
  }

  const ingredients = new Map();
  for (const [key, name, unit, cost, sup, storage, boughtAs, pack] of INGREDIENTS) {
    const s = suppliers.get(sup);
    const daily = used.get(key) * 1.1; // and a little for waste
    const level = niceUp(daily * (s.gap + 1) * 1.25, unit);
    const qty = r3(Math.max(1, Math.ceil(daily * s.gap * 1.4 / pack)) * pack);
    const ing = { key, name, unit, cost, sup: s, pack, level, qty, stock: r3(level + qty), onOrder: 0 };
    ing.id = add('Ingredients', {
      Name: name, Supplier: s.id, Unit: unit, 'Cost per Unit': cost, 'Bought As': boughtAs, Storage: storage,
      'In Stock': ing.stock, 'Reorder Level': level, 'Reorder Quantity': qty,
    });
    ingredients.set(key, ing);
  }
  for (const it of items) {
    it.takeRate = takeawayRate(it);
    it.id = add('Menu Items', {
      Name: it.name, Category: categories.get(it.cat), Price: it.price, Takeaway: it.takeaway, 'Baked Here': it.baked, Diet: it.diet,
      'Takeaway VAT Rate': it.takeRate, Active: !it.retired,
    });
    it.recipe.forEach(([k, q], i) => add('Recipe Lines', { 'Menu Item': it.id, Line: i + 1, Ingredient: ingredients.get(k).id, Quantity: q }));
    it.uses = it.recipe.map(([k, q]) => ({ ing: ingredients.get(k), q }));
  }

  // Staff, each with a block or two of holiday in the year.
  const staff = STAFF.map(([role, rate, station, patterns], i) => {
    const p = person(r, { domain: 'littlefig.example' });
    const started = -r.int(i === 0 ? 1900 : 450, i === 0 ? 2600 : 1500);
    const id = add('Staff', { ...p, Role: role, 'Hourly Rate': rate, Phone: `07700 900${String(310 + i * 7).padStart(3, '0')}`, Started: cal.day(started) });
    const holidays = Array.from({ length: i === 0 ? 1 : 2 }, () => r.int(-days, ahead - 4));
    return { id, role, rate, station, patterns, holidays };
  });

  const closed = (d) => { const at = cal.at(d); return at.getMonth() === 11 && at.getDate() === 25; };
  // The day a supplier next delivers after day d, the cafe being open.
  const nextDelivery = (s, d) => { let e = d + 1; while (!s.deliver.includes(cal.weekday(e)) || closed(e)) e++; return e; };
  const lastOpenBefore = (e) => { let d = e - 1; while (closed(d)) d--; return d; };

  // A payslip for each person for each week (Monday to Sunday) once the week is over,
  // paid on the Friday after; its hours and gross pay are added up from the shifts.
  const payslips = new Map();
  const payslipFor = (s, d) => {
    const monday = d - ((cal.weekday(d) + 6) % 7);
    if (monday + 6 >= 0) return null; // the week is not over yet
    const k = `${s.id}|${monday}`;
    if (!payslips.has(k)) {
      const id = add('Payslips', {
        Number: numbered('PAY', monday + 11, 4), Staff: s.id, 'Week Beginning': cal.day(monday), 'Week Ending': cal.day(monday + 6),
        'Pay Date': cal.day(monday + 11), 'Hourly Rate': s.rate, Hours: 0, 'Gross Pay': 0,
      });
      payslips.set(k, { id, hours: 0, gross: 0 });
    }
    return payslips.get(k);
  };

  // A day's shifts: who is on, from when to when, and whether they worked it.
  const shiftsFor = (d) => {
    const wd = cal.weekday(d);
    const month = cal.at(d).getMonth();
    const out = [];
    for (const s of staff) {
      for (const [weekdays, from, to, brk, only] of s.patterns) {
        if (!weekdays.includes(wd) || (only === 'summer' && !isSummer(month))) continue;
        const holiday = s.holidays.some((h) => d >= h && d < h + 7);
        let end = to;
        let status = holiday ? 'Holiday' : d >= 0 ? 'Rota' : r.chance(0.015) ? 'Sick' : 'Worked';
        // A busy day keeps someone half an hour longer.
        if (status === 'Worked' && (wd === 0 || wd === 6 || isSummer(month)) && r.chance(0.12)) end += 0.5;
        const paid = status === 'Sick' ? 0 : r3(end - from - brk / 60);
        const [fh, fm] = clock(from);
        const [th, tm] = clock(end);
        const wage = status === 'Worked' || status === 'Holiday' ? money(paid * s.rate) : 0;
        const slip = payslipFor(s, d);
        if (slip) { slip.hours = r3(slip.hours + paid); slip.gross = money(slip.gross + wage); }
        add('Shifts', { Staff: s.id, Start: cal.time(d, fh, fm), End: cal.time(d, th, tm), 'Break Minutes': brk, 'Paid Hours': paid, Station: s.station, Status: status, Payslip: slip ? slip.id : null });
        out.push({ s, from, to: end, status, wage });
      }
    }
    return out;
  };

  const pending = []; // purchases not yet in: { id, at, lines: [{ ing, n }] }
  const bills = []; // purchases received: { id, at, terms, slow }
  const poolCache = new Map();

  // The records stop at the start of today: today's deliveries are not booked in
  // yet, today's supplier orders not yet placed and the till not yet open.
  for (let d = -days; d < 0; d++) {
    if (closed(d)) continue;
    const date = cal.day(d);
    const wd = cal.weekday(d);
    const month = cal.at(d).getMonth();

    // Deliveries come first thing.
    for (let i = pending.length - 1; i >= 0; i--) {
      const po = pending[i];
      if (po.at !== d) continue;
      base.update('Purchases', po.id, { 'Received On': date, 'Invoice Number': `${po.s.prefix}-${po.s.invoice++}` });
      for (const l of po.lines) { l.ing.stock = r3(l.ing.stock + l.n); l.ing.onOrder = 0; }
      bills.push({ id: po.id, at: d, terms: po.s.terms, slow: r.chance(0.07) });
      pending.splice(i, 1);
    }

    // The morning order for each supplier whose next delivery this is the last open day before:
    // what is at or below its reorder level, brought back up to the level and a usual order
    // (on a standing order, whatever has been used since the last delivery).
    for (const [, s] of suppliers) {
      const e = nextDelivery(s, d);
      if (lastOpenBefore(e) !== d) continue;
      const lines = [];
      for (const ing of ingredients.values()) {
        if (ing.sup !== s || ing.onOrder || ing.stock > (s.standing ? ing.level + ing.qty - ing.pack : ing.level)) continue;
        const n = r3(Math.max(1, Math.ceil((ing.level + ing.qty - ing.stock) / ing.pack)) * ing.pack);
        lines.push({ ing, n });
      }
      if (!lines.length) continue;
      const id = add('Purchases', { Number: numbered('PO', d, 4), Supplier: s.id, 'Ordered On': date, 'Due On': cal.day(e) });
      // Prices crept up through the year: older deliveries cost a little less.
      for (const l of lines) l.ing.onOrder = l.n;
      for (const l of lines) add('Purchase Lines', { Purchase: id, Ingredient: l.ing.id, Quantity: l.n, 'Unit Cost': Math.round(l.ing.cost * (1 + 0.05 * Math.min(0, e) / 365) * 10000) / 10000 });
      pending.push({ id, at: e, s, lines });
    }

    // The rota as worked.
    const shifts = shiftsFor(d);
    const on = shifts.filter((x) => x.status === 'Worked');
    const front = on.filter((x) => ['Barista', 'Floor', 'Manager'].includes(x.s.station));
    const servedBy = (t) => (front.find((x) => x.from <= t && t < x.to) || front[0] || on[0] || shifts[0]).s.id;

    // The till.
    const weather = weatherFor(r, month);
    const key = `${month}|${d < RETIRED_BEFORE}`;
    if (!poolCache.has(key)) poolCache.set(key, poolsFor(items, month, d < RETIRED_BEFORE));
    const cardDown = r.chance(0.01);
    const soldOut = new Set();
    let card = 0;
    let cash = 0;
    let orders = 0;
    let covers = 0;
    let vat = 0;
    for (const o of dayOrders(r, { month, wd, weather, pools: poolCache.get(key) })) {
      const voided = r.chance(0.004);
      const lines = new Map();
      for (const it of o.items) {
        // Nothing is sold that the kitchen has run out of.
        if (!voided) {
          const uses = it.uses;
          let short = null;
          for (let i = 0; i < uses.length; i++) if (uses[i].ing.stock < uses[i].q - 1e-9) { short = uses[i].ing; break; }
          if (short) {
            soldOut.add(short.name.toLowerCase());
            continue;
          }
          for (let i = 0; i < uses.length; i++) uses[i].ing.stock -= uses[i].q;
        }
        lines.set(it, (lines.get(it) || 0) + 1);
      }
      if (!lines.size) continue;
      // Each line's VAT is worked out within its price and rounded, then added up, as the till does.
      let total = 0;
      let tax = 0;
      for (const [it, q] of lines) {
        const gross = money(q * it.price);
        const rate = o.takeaway ? it.takeRate : STANDARD_RATE;
        total += gross;
        tax += money(gross - gross / (1 + rate));
      }
      total = money(total);
      tax = money(tax);
      // Every bill rung up takes the next receipt number, a voided one too.
      const number = numbered('RCP', d, 5);
      const payment = r.chance(cardDown ? 0.4 : 0.86) ? 'Card' : 'Cash';
      if (!voided) {
        if (payment === 'Card') card += total; else cash += total;
        orders++;
        covers += o.covers;
        vat += tax;
      }
      if (d < -orderDays) continue; // older till records are summed in the cash-up only
      const t = o.h + o.m / 60;
      const order = add('Orders', {
        Number: number, When: cal.time(d, o.h, o.m), Service: o.takeaway ? 'Takeaway' : 'Dine in', 'Table No': o.takeaway ? null : r.int(1, 10),
        Covers: o.covers, Staff: servedBy(t), Payment: payment, Net: money(total - tax), Tax: tax, Total: total, Voided: voided,
      });
      for (const [it, q] of lines) add('Order Lines', { Order: order, 'Menu Item': it.id, Quantity: q, 'Unit Price': it.price, 'Tax Rate': o.takeaway ? it.takeRate : STANDARD_RATE });
    }

    for (const ing of ingredients.values()) ing.stock = r3(ing.stock);

    // What was thrown away: produce past its best, a dropped tray, a dish made twice.
    const recorder = () => (on.find((x) => x.s.station === 'Kitchen') || on[0] || shifts[0]).s.id;
    for (const [k, reason, p, lo, hi, step] of WASTE) {
      if (!r.chance(p)) continue;
      const ing = ingredients.get(k);
      const n = r3(Math.min(ing.stock, (lo + Math.round(r.next() * (hi - lo) / step) * step)));
      if (n <= 0) continue;
      ing.stock = r3(ing.stock - n);
      add('Waste', { Ingredient: ing.id, Date: date, Quantity: n, Reason: reason, 'Recorded By': recorder() });
    }

    // The cash-up at close: the till's reading, the cash counted, the day's wages from the rota.
    const short = r.chance(0.12) ? r.pick([-10, -5, -2, -1, -0.5, -0.2, 0.3, 0.5, 1, 2, 5]) : 0;
    const notes = [];
    if (cardDown) notes.push('Card machine down for an hour at lunch.');
    if (soldOut.size) notes.push(`Ran out of ${[...soldOut].slice(0, 2).join(' and ')}.`);
    if (short < 0) notes.push(`Till short by ${(-short).toFixed(2)}.`);
    if (short > 0) notes.push(`Till over by ${short.toFixed(2)}.`);
    const closer = [...on].sort((a, b) => b.to - a.to)[0] || shifts[0];
    add('Cash Ups', {
      Date: date, 'Closed By': closer.s.id, Weather: weather, Orders: orders, Covers: covers, 'Card Takings': money(card), 'Cash Takings': money(cash), VAT: money(vat),
      'Cash Counted': money(Math.max(0, cash + short)), 'Wage Cost': money(shifts.reduce((n, x) => n + x.wage, 0)), Notes: notes.join(' ') || null,
    });
  }

  // The rota ahead: today and the next two weeks.
  for (let d = 0; d <= ahead; d++) if (!closed(d)) shiftsFor(d);
  for (const slip of payslips.values()) base.update('Payslips', slip.id, { Hours: slip.hours, 'Gross Pay': slip.gross });

  // Bills: paid around their due day, a few late, the newest not yet.
  for (const b of bills) {
    const paidAt = b.at + b.terms + (b.slow ? r.int(8, 25) : r.int(-4, 2));
    if (paidAt <= -1) base.update('Purchases', b.id, { 'Paid On': cal.day(paidAt) });
  }

  for (const ing of ingredients.values()) base.update('Ingredients', ing.id, { 'In Stock': ing.stock, 'On Order': ing.onOrder });
  return counts;
}

// Waste: [ingredient, reason, chance on a given day, least, most, step].
const WASTE = [
  ['milk', 'Out of date', 0.05, 1, 2, 1], ['milk', 'Spilt', 0.06, 0.5, 1, 0.5], ['oat', 'Spilt', 0.03, 0.5, 1, 0.5],
  ['salad', 'Out of date', 0.18, 0.1, 0.4, 0.05], ['spinach', 'Out of date', 0.12, 0.1, 0.3, 0.05], ['berries', 'Gone mouldy', 0.12, 0.1, 0.3, 0.05],
  ['avocado', 'Overripe', 0.2, 1, 3, 1], ['bananas', 'Overripe', 0.1, 0.3, 0.8, 0.1], ['sourdough', 'Stale', 0.15, 0.5, 1.5, 0.5],
  ['bagels', 'Out of date', 0.08, 1, 3, 1], ['cream', 'Out of date', 0.04, 0.5, 1, 0.5], ['eggs', 'Dropped', 0.07, 1, 4, 1],
  ['tomatoes', 'Out of date', 0.07, 0.2, 0.5, 0.1], ['mushrooms', 'Out of date', 0.05, 0.2, 0.5, 0.1], ['bacon', 'Kitchen error', 0.06, 0.06, 0.12, 0.06],
  ['eggs', 'Kitchen error', 0.05, 2, 3, 1], ['halloumi', 'Kitchen error', 0.03, 0.06, 0.1, 0.02], ['yoghurt', 'Out of date', 0.03, 0.5, 1, 0.5],
];
