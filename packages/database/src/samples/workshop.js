// The workshop: Granite and Grain Woodworks, an invented workshop of five
// makers building furniture and joinery from timber, sheet materials and
// fittings. It sells to gift and homeware shops, to interior designers
// (kitchen doors and drawer fronts made to measure, tables, worktops) and to
// the public. Its materials and suppliers, what it buys, its products with
// their bills of materials, work orders and the hours booked to them, its
// customers, their orders and the stock of finished goods.
//
// Its records come from fifteen months run a day at a time. Orders come in;
// a made-to-order line becomes a work order of its own, and a stock line is
// sent from the shelf, with a batch made when the shelf runs low. Planning a
// work order sets its materials aside (Allocated), starting it takes them
// from stock, finishing it puts the goods on the shelf, and delivery takes
// them off again. Twice a week the buyer orders what the next fortnight's
// work leaves at or below its reorder level, waiting for a supplier's
// minimum order unless work would be held up. Shops buy more boards before
// Christmas; designers' kitchens slow in August and December; the workshop
// shuts over Christmas and January catches up late. The file is made first
// thing in the morning, before the day's work.

import { address, companies, money, person, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const LINE = 'Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)';
// The labour rate the margin queries cost a product's minutes at: about the
// makers' average hourly cost with holiday and insurance on top.
const RATE = 21;
const LABOUR = `[Labour Minutes] / 60 * ${RATE}`;

export const spec = {
  // Where a value below nought is right: stock going out, work or an order
  // already late, free stock spoken for, an invoice not yet due.
  negative: {
    'Material Movements': ['Quantity'], 'Stock Movements': ['Quantity'], 'Open work orders': ['Days Left'],
    'Orders to deliver': ['Days Left'], 'Unpaid invoices': ['Days Overdue'], 'Materials to reorder': ['Free Stock'],
  },
  tables: [
    { name: 'Suppliers', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Supplies', type: 'text' }, { name: 'Contact', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'City', type: 'text' },
      { name: 'Lead Days', type: 'integer', default: 5 }, { name: 'Minimum Order', type: 'currency', default: 0 },
    ] },
    { name: 'Materials', fields: [
      { name: 'Code', type: 'text', unique: true }, { name: 'Name', type: 'text', required: true }, { name: 'Kind', type: 'text' },
      { name: 'Unit', type: 'text' }, { name: 'Unit Cost', type: 'currency' }, { name: 'Supplier', ...ref('Suppliers') },
      { name: 'In Stock', type: 'number', default: 0 }, { name: 'On Order', type: 'number', default: 0 }, { name: 'Allocated', type: 'number', default: 0 },
      { name: 'Reorder Level', type: 'number', default: 0 }, { name: 'Reorder Quantity', type: 'number', default: 0 },
    ] },
    { name: 'Material Purchases', fields: [
      { name: 'Supplier', ...ref('Suppliers'), required: true }, { name: 'Ordered On', type: 'date' }, { name: 'Expected', type: 'date' },
      { name: 'Received On', type: 'date' }, { name: 'Status', type: 'text', default: 'Ordered' },
    ] },
    { name: 'Purchase Lines', fields: [
      { name: 'Purchase', ...ref('Material Purchases'), required: true }, { name: 'Material', ...ref('Materials'), required: true },
      { name: 'Quantity', type: 'number' }, { name: 'Unit Cost', type: 'currency' },
    ] },
    { name: 'Material Movements', fields: [
      { name: 'Material', ...ref('Materials'), required: true }, { name: 'When', type: 'date' }, { name: 'Kind', type: 'text' },
      { name: 'Quantity', type: 'number' }, { name: 'Reference', type: 'text' },
    ] },
    { name: 'Products', fields: [
      { name: 'Code', type: 'text', unique: true }, { name: 'Name', type: 'text', required: true }, { name: 'Category', type: 'text' },
      { name: 'Price', type: 'currency' }, { name: 'Material Cost', type: 'currency' }, { name: 'Labour Minutes', type: 'integer' },
      { name: 'Made to Order', type: 'yesno' }, { name: 'In Stock', type: 'integer', default: 0 }, { name: 'Stock Level', type: 'integer', default: 0 },
      { name: 'Batch Size', type: 'integer', default: 1 }, { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Bill of Materials', fields: [
      { name: 'Product', ...ref('Products'), required: true }, { name: 'Material', ...ref('Materials'), required: true }, { name: 'Quantity', type: 'number' },
    ] },
    { name: 'Makers', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Role', type: 'text' },
      { name: 'Hourly Cost', type: 'currency' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Started', type: 'date' },
    ] },
    { name: 'Customers', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Kind', type: 'text', default: 'Public' }, { name: 'Contact', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Address', type: 'text' }, { name: 'City', type: 'text' },
      { name: 'Postcode', type: 'text' }, { name: 'Payment Days', type: 'integer', default: 0 }, { name: 'Trade Discount', type: 'number', default: 0 },
    ] },
    { name: 'Sales Orders', fields: [
      { name: 'Customer', ...ref('Customers'), required: true }, { name: 'Order Date', type: 'date' }, { name: 'Delivery Due', type: 'date' },
      { name: 'Delivered On', type: 'date' }, { name: 'Status', type: 'text', default: 'Open' }, { name: 'Delivery Fee', type: 'currency', default: 0 },
      { name: 'Invoice Amount', type: 'currency' }, { name: 'Invoice Due', type: 'date' }, { name: 'Paid On', type: 'date' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Sales Lines', fields: [
      { name: 'Order', ...ref('Sales Orders'), required: true }, { name: 'Product', ...ref('Products'), required: true },
      { name: 'Quantity', type: 'integer', default: 1 }, { name: 'Unit Price', type: 'currency' }, { name: 'Discount', type: 'number', default: 0 },
    ] },
    // Time Sheets comes before Work Orders so that a query with both and
    // Makers joins a maker to the hours they booked, not to the work order's
    // lead maker (the first relationship declared is the one a query uses).
    { name: 'Time Sheets', fields: [
      { name: 'Maker', ...ref('Makers'), required: true }, { name: 'Work Order', ...ref('Work Orders'), required: true },
      { name: 'Worked On', type: 'date' }, { name: 'Hours', type: 'number' },
    ] },
    { name: 'Work Orders', fields: [
      { name: 'Product', ...ref('Products'), required: true }, { name: 'Quantity', type: 'integer', default: 1 }, { name: 'Maker', ...ref('Makers') },
      { name: 'Sales Order', ...ref('Sales Orders') }, { name: 'Planned Start', type: 'date' }, { name: 'Due Date', type: 'date' },
      { name: 'Started', type: 'date' }, { name: 'Finished', type: 'date' }, { name: 'Status', type: 'text', default: 'Planned' },
    ] },
    { name: 'Stock Movements', fields: [
      { name: 'Product', ...ref('Products'), required: true }, { name: 'When', type: 'date' }, { name: 'Kind', type: 'text' },
      { name: 'Quantity', type: 'integer' }, { name: 'Reference', type: 'text' },
    ] },
  ],
  queries: {
    'Costed bill of materials': { tables: ['Products', 'Bill of Materials', 'Materials'], columns: [
      { table: 'Products', field: 'Code', as: 'Product Code', sort: 'asc' }, { table: 'Products', field: 'Name', as: 'Product' },
      { table: 'Materials', field: 'Name', as: 'Material' }, { table: 'Bill of Materials', field: 'Quantity' }, { table: 'Materials', field: 'Unit' },
      { table: 'Materials', field: 'Unit Cost' }, { expr: 'Round([Bill of Materials].[Quantity] * [Unit Cost], 2)', as: 'Line Cost' },
      { table: 'Products', field: 'Active', criteria: ['Yes'], show: false },
    ] },
    'Product margins': { tables: ['Products'], columns: [
      { table: 'Products', field: 'Code' }, { table: 'Products', field: 'Name', as: 'Product' }, { table: 'Products', field: 'Category' },
      { table: 'Products', field: 'Price' }, { table: 'Products', field: 'Material Cost' }, { expr: `Round(${LABOUR}, 2)`, as: 'Labour Cost' },
      { expr: `Round([Price] - [Material Cost] - ${LABOUR}, 2)`, as: 'Margin' },
      { expr: `Round(100 * ([Price] - [Material Cost] - ${LABOUR}) / [Price], 1)`, as: 'Margin %', sort: 'desc' },
      { table: 'Products', field: 'Active', criteria: ['Yes'], show: false },
    ] },
    'Materials for planned work': { tables: ['Materials', 'Bill of Materials', 'Products', 'Work Orders'], columns: [
      { table: 'Materials', field: 'Code', total: 'group', sort: 'asc' }, { table: 'Materials', field: 'Name', as: 'Material', total: 'group' },
      { table: 'Materials', field: 'Unit', total: 'group' },
      { expr: 'Round([Work Orders].[Quantity] * [Bill of Materials].[Quantity], 2)', as: 'Needed', total: 'sum' },
      { table: 'Materials', field: 'Allocated', total: 'group' }, { table: 'Materials', field: 'In Stock', total: 'group' },
      { table: 'Work Orders', field: 'Status', total: 'where', criteria: ['Planned'] },
    ] },
    'Material shortages': { tables: ['Materials', 'Suppliers'], columns: [
      { table: 'Suppliers', field: 'Name', as: 'Supplier', sort: 'asc' }, { table: 'Materials', field: 'Code', sort: 'asc' },
      { table: 'Materials', field: 'Name', as: 'Material' }, { table: 'Materials', field: 'Unit' }, { table: 'Materials', field: 'Allocated' },
      { table: 'Materials', field: 'In Stock' }, { table: 'Materials', field: 'On Order' },
      { expr: 'Round([Allocated] - [In Stock] - [On Order], 2)', as: 'Short By', criteria: ['> 0'] },
    ] },
    'Materials to reorder': { tables: ['Materials', 'Suppliers'], columns: [
      { table: 'Suppliers', field: 'Name', as: 'Supplier', sort: 'asc' }, { table: 'Suppliers', field: 'Minimum Order' },
      { table: 'Materials', field: 'Code', sort: 'asc' }, { table: 'Materials', field: 'Name', as: 'Material' }, { table: 'Materials', field: 'Unit' },
      { expr: 'Round([In Stock] + [On Order] - [Allocated], 2)', as: 'Free Stock' }, { table: 'Materials', field: 'Reorder Level' },
      { table: 'Materials', field: 'Reorder Quantity' },
      { expr: '[In Stock] + [On Order] - [Allocated] - [Reorder Level]', as: 'Above Reorder Level', criteria: ['<= 0'], show: false },
    ] },
    'Purchases on order': { tables: ['Material Purchases', 'Suppliers', 'Purchase Lines'], columns: [
      { table: 'Material Purchases', field: 'ID', as: 'Purchase', total: 'group' }, { table: 'Suppliers', field: 'Name', as: 'Supplier', total: 'group' },
      { table: 'Material Purchases', field: 'Ordered On', total: 'group' }, { table: 'Material Purchases', field: 'Expected', total: 'group', sort: 'asc' },
      { expr: 'Round([Quantity] * [Unit Cost], 2)', as: 'Value', total: 'sum' },
      { table: 'Material Purchases', field: 'Status', total: 'where', criteria: ['Ordered'] },
    ] },
    'Open work orders': { tables: ['Work Orders', 'Products', 'Makers'], columns: [
      { table: 'Work Orders', field: 'ID', as: 'Work Order' }, { table: 'Products', field: 'Name', as: 'Product' }, { table: 'Work Orders', field: 'Quantity' },
      { expr: '[First Name] & " " & [Last Name]', as: 'Maker' }, { table: 'Work Orders', field: 'Sales Order', as: 'For Order' },
      { table: 'Work Orders', field: 'Planned Start' }, { table: 'Work Orders', field: 'Due Date', sort: 'asc' }, { table: 'Work Orders', field: 'Started' },
      { table: 'Work Orders', field: 'Status', criteria: ['In ("Planned", "In progress")'] }, { expr: 'DateDiff("d", Date(), [Due Date])', as: 'Days Left' },
    ] },
    'Work order costs': { tables: ['Time Sheets', 'Makers', 'Work Orders', 'Products'], columns: [
      { table: 'Work Orders', field: 'ID', as: 'Work Order', total: 'group' }, { table: 'Products', field: 'Name', as: 'Product', total: 'group' },
      { table: 'Work Orders', field: 'Quantity', total: 'group' }, { table: 'Work Orders', field: 'Finished', total: 'group', sort: 'desc' },
      { expr: 'Round([Work Orders].[Quantity] * [Labour Minutes] / 60, 2)', as: 'Standard Hours', total: 'group' },
      { table: 'Time Sheets', field: 'Hours', as: 'Hours Booked', total: 'sum' },
      { expr: 'Round([Hours] * [Hourly Cost], 2)', as: 'Labour Cost', total: 'sum' },
      { expr: 'Round([Work Orders].[Quantity] * [Material Cost], 2)', as: 'Material Cost', total: 'group' },
      { table: 'Work Orders', field: 'Status', total: 'where', criteria: ['Done'] },
      { expr: 'DateDiff("d", [Finished], Date())', as: 'Days Since', total: 'where', criteria: ['Between 0 And 90'] },
    ] },
    'Maker hours by month': { tables: ['Time Sheets', 'Makers'], columns: [
      { expr: '[First Name] & " " & [Last Name]', as: 'Maker', total: 'group', sort: 'asc' },
      { expr: 'Format([Worked On], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { table: 'Time Sheets', field: 'Hours', as: 'Hours', total: 'sum' }, { expr: 'Round([Hours] * [Hourly Cost], 2)', as: 'Cost', total: 'sum' },
    ] },
    'Maker hours, last 4 weeks': { tables: ['Time Sheets', 'Makers'], columns: [
      { expr: '[First Name] & " " & [Last Name]', as: 'Maker', total: 'group', sort: 'asc' }, { table: 'Makers', field: 'Role', total: 'group' },
      { table: 'Time Sheets', field: 'Hours', as: 'Hours', total: 'sum' }, { expr: 'Round([Hours] * [Hourly Cost], 2)', as: 'Cost', total: 'sum' },
      { expr: 'DateDiff("d", [Worked On], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 27'] },
    ] },
    'Orders to deliver': { tables: ['Sales Orders', 'Customers', 'Sales Lines'], columns: [
      { table: 'Sales Orders', field: 'ID', as: 'Order', total: 'group' }, { table: 'Customers', field: 'Name', as: 'Customer', total: 'group' },
      { table: 'Customers', field: 'Kind', total: 'group' }, { table: 'Sales Orders', field: 'Order Date', total: 'group' },
      { table: 'Sales Orders', field: 'Delivery Due', total: 'group', sort: 'asc' },
      { expr: 'DateDiff("d", Date(), [Delivery Due])', as: 'Days Left', total: 'group', criteria: ['<= 14'] },
      { expr: LINE, as: 'Goods', total: 'sum' },
      { table: 'Sales Orders', field: 'Status', total: 'where', criteria: ['Open'] },
    ] },
    'Late deliveries': { tables: ['Sales Orders', 'Customers'], columns: [
      { table: 'Sales Orders', field: 'ID', as: 'Order' }, { table: 'Customers', field: 'Name', as: 'Customer' }, { table: 'Customers', field: 'Kind' },
      { table: 'Sales Orders', field: 'Delivery Due' }, { table: 'Sales Orders', field: 'Delivered On', sort: 'desc' },
      { expr: 'DateDiff("d", [Delivery Due], [Delivered On])', as: 'Days Late', criteria: ['> 0'] },
      { expr: 'DateDiff("d", [Delivered On], Date())', as: 'Days Ago', criteria: ['Between 0 And 365'], show: false },
    ] },
    'Unpaid invoices': { tables: ['Sales Orders', 'Customers'], columns: [
      { table: 'Sales Orders', field: 'ID', as: 'Order' }, { table: 'Customers', field: 'Name', as: 'Customer' }, { table: 'Customers', field: 'Kind' },
      { table: 'Sales Orders', field: 'Delivered On' }, { table: 'Sales Orders', field: 'Invoice Due', sort: 'asc' }, { table: 'Sales Orders', field: 'Invoice Amount' },
      { expr: 'DateDiff("d", [Invoice Due], Date())', as: 'Days Overdue' },
      { table: 'Sales Orders', field: 'Status', criteria: ['Delivered'], show: false }, { table: 'Sales Orders', field: 'Paid On', criteria: ['Is Null'], show: false },
    ] },
    'Finished goods in stock': { tables: ['Products'], columns: [
      { table: 'Products', field: 'Category', sort: 'asc' }, { table: 'Products', field: 'Code', sort: 'asc' }, { table: 'Products', field: 'Name', as: 'Product' },
      { table: 'Products', field: 'In Stock', criteria: ['> 0'] }, { table: 'Products', field: 'Stock Level' },
      { expr: `Round([In Stock] * ([Material Cost] + ${LABOUR}), 2)`, as: 'Value at Cost' },
      { expr: 'Round([In Stock] * [Price], 2)', as: 'Value at Price' },
    ] },
    'Sales by month': { tables: ['Sales Orders', 'Sales Lines'], columns: [
      { expr: 'Format([Delivered On], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { table: 'Sales Lines', field: 'Quantity', as: 'Items', total: 'sum' }, { expr: LINE, as: 'Sales', total: 'sum' },
      { table: 'Sales Orders', field: 'Status', total: 'where', criteria: ['Delivered'] },
    ] },
    'Sales by product': { tables: ['Sales Lines', 'Products', 'Sales Orders'], columns: [
      { table: 'Products', field: 'Category', total: 'group' }, { table: 'Products', field: 'Name', as: 'Product', total: 'group' },
      { table: 'Sales Lines', field: 'Quantity', as: 'Items', total: 'sum' }, { expr: LINE, as: 'Sales', total: 'sum', sort: 'desc' },
      { table: 'Sales Orders', field: 'Status', total: 'where', criteria: ['Delivered'] },
    ] },
    'Sales by customer': { tables: ['Customers', 'Sales Orders', 'Sales Lines'], columns: [
      { table: 'Customers', field: 'Name', as: 'Customer', total: 'group' }, { table: 'Customers', field: 'Kind', total: 'group' },
      { table: 'Sales Orders', field: 'ID', as: 'Order Lines', total: 'count' }, { expr: LINE, as: 'Sales', total: 'sum', sort: 'desc' },
      { table: 'Sales Orders', field: 'Status', total: 'where', criteria: ['Delivered'] },
    ] },
  },
  reports: [
    { name: 'Costed bill of materials', query: 'Costed bill of materials', group: 'Product', columns: ['Material', { field: 'Quantity', type: 'number' }, 'Unit', { field: 'Unit Cost', type: 'currency' }, { field: 'Line Cost', type: 'currency' }], totals: [{ field: 'Line Cost', fn: 'sum' }] },
    { name: 'Product margins', query: 'Product margins', columns: ['Code', 'Product', { field: 'Price', type: 'currency' }, { field: 'Material Cost', type: 'currency' }, { field: 'Labour Cost', type: 'currency' }, { field: 'Margin', type: 'currency' }, { field: 'Margin %', type: 'number' }] },
    { name: 'Production plan', query: 'Open work orders', group: 'Maker', sort: 'Planned Start', columns: ['Work Order', 'Product', { field: 'Quantity', type: 'integer' }, 'Status', { field: 'Planned Start', type: 'date' }, { field: 'Due Date', type: 'date' }, { field: 'Days Left', type: 'integer' }], totals: [{ field: 'Work Order', fn: 'count' }] },
    { name: 'Shortages for buying', query: 'Material shortages', group: 'Supplier', columns: ['Code', 'Material', 'Unit', { field: 'Allocated', type: 'number' }, { field: 'In Stock', type: 'number' }, { field: 'On Order', type: 'number' }, { field: 'Short By', type: 'number' }] },
    { name: 'Reorder list', query: 'Materials to reorder', group: 'Supplier', columns: ['Code', 'Material', 'Unit', { field: 'Free Stock', type: 'number' }, { field: 'Reorder Level', type: 'number' }, { field: 'Reorder Quantity', type: 'number' }] },
    { name: 'Delivery schedule', query: 'Orders to deliver', columns: ['Order', 'Customer', 'Kind', { field: 'Order Date', type: 'date' }, { field: 'Delivery Due', type: 'date' }, { field: 'Days Left', type: 'integer' }, { field: 'Goods', type: 'currency' }], totals: [{ field: 'Goods', fn: 'sum' }] },
    { name: 'Monthly sales', query: 'Sales by month', columns: ['Month', { field: 'Items', type: 'integer' }, { field: 'Sales', type: 'currency' }], totals: [{ field: 'Items', fn: 'sum' }, { field: 'Sales', fn: 'sum' }] },
    { name: 'Invoices due', query: 'Unpaid invoices', group: 'Customer', columns: ['Order', { field: 'Delivered On', type: 'date' }, { field: 'Invoice Due', type: 'date' }, { field: 'Days Overdue', type: 'integer' }, { field: 'Invoice Amount', type: 'currency' }], totals: [{ field: 'Invoice Amount', fn: 'sum' }] },
    { name: 'Hours by maker', query: 'Maker hours by month', group: 'Maker', columns: ['Month', { field: 'Hours', type: 'number' }, { field: 'Cost', type: 'currency' }], totals: [{ field: 'Hours', fn: 'sum' }, { field: 'Cost', fn: 'sum' }] },
  ],
};

export const about = {
  name: 'Granite and Grain Woodworks',
  kind: 'Workshop or maker',
  line: 'Materials, bills of materials, work orders, makers\' hours, orders and deliveries for a small furniture workshop.',
  summary: [
    'Granite and Grain Woodworks is an invented workshop of five makers in a unit on an industrial estate. They build dining tables, chairs, stools, benches, shelves, bookcases and chopping boards from oak, ash, walnut and beech, and make kitchen doors, drawer fronts and worktops to measure. About ten gift and homeware shops buy boards, coasters and shelves at a trade discount, interior designers order kitchens and furniture for their clients, and members of the public buy from the workshop\'s showroom and website.',
    'Every product has a bill of materials: the timber, sheet, fittings, finish and packing that one of it takes. Small goods, chairs and shelves are made in batches to keep a stock on the shelf. Tables, bookcases, worktops and kitchen doors are made to order, so each order line of them gets a work order of its own. Planning a work order sets its materials aside, starting it takes them from stock, and the makers book their hours to it each day until it is finished.',
    'The records cover fifteen months up to the morning the file is made, with orders already taken for the weeks ahead. Shops order more boards before Christmas, kitchens slow in August, the workshop closes between Christmas and New Year, and timber prices went up part way through. Some deliveries are late, most of them in January after the break or when a maker was off sick, and a few trade customers pay late. The buyer orders twice a week for the work starting in the next fortnight, so a job further off can show as a shortage until it comes closer.',
    'Use it to see how a bill of materials, stock and production fit together, or start from it empty for a workshop of your own.',
  ],
  suits: ['A cabinet maker or joiner', 'A bakery\'s production side', 'A print shop', 'A sign maker', 'A soap or candle maker', 'A garment or textile maker', 'An electronics assembler', 'A metal fabricator'],
  tables: {
    Suppliers: { about: 'The merchants the workshop buys its materials from.', fields: {
      Name: 'The supplier\'s name.', Supplies: 'What you buy from them, such as Timber or Fittings and fixings.', Contact: 'The person you order from.',
      Phone: 'Their phone number.', 'E-mail': 'Where orders go.', City: 'Where they are.',
      'Lead Days': 'How many days an order usually takes to arrive, used to judge when to order.',
      'Minimum Order': 'The smallest order they deliver without a charge. The buyer waits until enough is needed from them, unless work would be held up or the order has waited a fortnight.',
    } },
    Materials: { about: 'Everything that goes into a product: timber, sheet materials, fittings, fixings, finishes and packing. One record per thing you buy and keep.', fields: {
      Code: 'A short code for the material, used on the bench and on purchase orders. Each code is used once.', Name: 'What it is, with its size.',
      Kind: 'Timber, Sheet, Fittings, Fixings, Finishes or Packing.', Unit: 'How it is counted: metre, sheet, each, pair, box, litre or roll.',
      'Unit Cost': 'What one unit costs now. Change it when the supplier\'s price changes, then work out each product\'s Material Cost again.',
      Supplier: 'Who you usually buy it from.', 'In Stock': 'How much is in the workshop now.',
      'On Order': 'How much has been ordered and not yet arrived.',
      'Allocated': 'How much is set aside for work orders that are planned but not started. It goes up when a work order is planned and down when work starts and the material is taken.',
      'Reorder Level': 'When the free stock (in stock, plus on order, less allocated) falls to this or below, the material goes on the reorder list.',
      'Reorder Quantity': 'How much you usually order at a time.',
    } },
    'Material Purchases': { about: 'Each order placed with a supplier.', fields: {
      Supplier: 'Who it was ordered from.', 'Ordered On': 'The day it was placed.', Expected: 'When the supplier said it would come.',
      'Received On': 'The day it arrived; empty until then.', Status: 'Ordered until it arrives, then Received.',
    } },
    'Purchase Lines': { about: 'What was ordered on each purchase.', fields: {
      Purchase: 'The purchase the line belongs to.', Material: 'What was ordered.', Quantity: 'How much, in the material\'s unit.',
      'Unit Cost': 'The price of one unit on this order, kept so a later price change does not alter it.',
    } },
    'Material Movements': { about: 'Every change to a material\'s stock, so the In Stock figure can be traced.', fields: {
      Material: 'The material.', When: 'The day it happened.',
      Kind: 'Received (a delivery came in), Used (taken for a work order) or Counted (the opening count, or a correction after a stock count).',
      Quantity: 'How much came in (above nought) or went out (below nought).', Reference: 'The purchase or work order it belongs to, such as PO 12 or WO 340.',
    } },
    Products: { about: 'What the workshop makes and sells, one record per product.', fields: {
      Code: 'A short product code. Each code is used once.', Name: 'The product as it shows on orders and invoices.',
      Category: 'Tables, Seating, Storage, Small goods or Kitchen.', Price: 'The price to the public before any trade discount.',
      'Material Cost': 'The cost of one product\'s materials from its bill of materials, worked out at today\'s unit costs. The Costed bill of materials query shows how it is made up.',
      'Labour Minutes': 'How long one takes to make, in minutes, by a maker working at a normal pace.',
      'Made to Order': 'Ticked for products made for each order (tables, bookcases, worktops, kitchen doors); left clear for products kept on the shelf.',
      'In Stock': 'Finished pieces on the shelf, including made-to-order pieces finished and waiting for delivery.',
      'Stock Level': 'How many to keep on the shelf beyond what is already ordered. A batch is planned when stock and work in hand fall short of it.',
      'Batch Size': 'How many are made at once.', Active: 'Untick a product you no longer make; it stays in the old orders.',
    } },
    'Bill of Materials': { about: 'What one of each product is made from: one record per material in it.', fields: {
      Product: 'The product.', Material: 'A material it takes.', Quantity: 'How much of the material one product uses, in the material\'s unit, waste included.',
    } },
    Makers: { about: 'The people in the workshop who make things and book their hours.', fields: {
      'First Name': 'Their first name.', 'Last Name': 'Their family name.', Role: 'Workshop lead, cabinet maker, chair maker or apprentice.',
      'Hourly Cost': 'What an hour of their time costs the workshop, with holiday pay and insurance.', Phone: 'A number to reach them.',
      'E-mail': 'Their work address.', Started: 'The day they joined.',
    } },
    Customers: { about: 'The shops, designers and people who buy from the workshop.', fields: {
      Name: 'The business, or the person\'s name for a member of the public.', Kind: 'Shop, Designer or Public.',
      Contact: 'The person you deal with at a business.', Phone: 'Their phone number.', 'E-mail': 'Where invoices go.',
      Address: 'The street address for deliveries.', City: 'The town or city.', Postcode: 'The postcode.',
      'Payment Days': 'Days allowed to pay after delivery: 30 for shops, 14 or 30 for designers, nought for the public, who pay on delivery.',
      'Trade Discount': 'The discount on the public price, as a fraction: 0.3 is thirty percent off.',
    } },
    'Sales Orders': { about: 'Each order a customer places.', fields: {
      Customer: 'Who ordered.', 'Order Date': 'The day the order was taken.', 'Delivery Due': 'The day you promised to deliver.',
      'Delivered On': 'The day it was delivered; empty until then.', Status: 'Open, Delivered or Cancelled.',
      'Delivery Fee': 'What was charged for delivery; nought when collected.', 'Invoice Amount': 'The invoice total, goods after discount plus delivery, set on delivery.',
      'Invoice Due': 'The day payment is due, the delivery day plus the customer\'s payment days.', 'Paid On': 'The day it was paid; empty while it is owed.',
      Notes: 'Anything to remember about the order.',
    } },
    'Sales Lines': { about: 'The products on each order.', fields: {
      Order: 'The order the line belongs to.', Product: 'What was ordered.', Quantity: 'How many.',
      'Unit Price': 'The price of one at the time, kept so a later price change does not alter old orders.',
      Discount: 'The discount given, as a fraction: 0.1 is ten percent off.',
    } },
    'Time Sheets': { about: 'The hours each maker booked to each work order, one record per maker, work order and day.', fields: {
      Maker: 'Who did the work.', 'Work Order': 'What they worked on.', 'Worked On': 'The day.', Hours: 'Hours worked on it that day, in quarter hours.',
    } },
    'Work Orders': { about: 'A job to make a quantity of one product, either for a customer\'s order or to fill the shelf.', fields: {
      Product: 'What is to be made.', Quantity: 'How many.', Maker: 'The maker the job is given to.',
      'Sales Order': 'The order it is made for; empty for a batch made for stock.', 'Planned Start': 'When work is planned to begin.',
      'Due Date': 'When it must be finished: two working days before the delivery for an order.', Started: 'The day work began and the materials were taken.',
      Finished: 'The day the last piece was finished.', Status: 'Planned, In progress, Done or Cancelled.',
    } },
    'Stock Movements': { about: 'Every change to the stock of finished goods.', fields: {
      Product: 'The product.', When: 'The day it happened.',
      Kind: 'Opening (the count the records start from), Made (a work order finished), Delivered (sent to a customer) or Adjusted (damaged or found at a count).',
      Quantity: 'How many came in (above nought) or went out (below nought).', Reference: 'The work order or sales order it belongs to.',
    } },
  },
  queries: {
    'Costed bill of materials': 'Every material in every product, with how much it takes, the unit cost and the line cost. Use it to check a price or to see what a price rise in one timber does to each product.',
    'Product margins': `Each product\'s price less its material cost and its labour, with the margin as a percentage of the price, the best first. Labour is costed at ${RATE} an hour; change the number in the query to your own rate.`,
    'Materials for planned work': 'What the planned work orders need of each material, from their bills of materials, beside what is allocated and in stock. The Needed column is how the Allocated figure is made up.',
    'Material shortages': 'Materials where the planned work needs more than is in stock and on order, with how much is short. Buy these before the work that needs them is due to start, or it will wait.',
    'Materials to reorder': 'Materials whose free stock (in stock, plus on order, less allocated to planned work) is at or below the reorder level, by supplier, with the minimum order for that supplier and how much you usually order. Free stock below nought is a shortage.',
    'Purchases on order': 'Orders placed with suppliers that have not arrived, with when each is expected and what it is worth.',
    'Open work orders': 'Work planned or in progress, the soonest due first, with the maker, the order it is for and the days left. Days left below nought means it is late.',
    'Work order costs': 'Work orders finished in the last 90 days, with the hours booked against the standard hours, the labour cost at each maker\'s rate and the material cost.',
    'Maker hours by month': 'The hours each maker booked to work orders in each month, and what they cost.',
    'Maker hours, last 4 weeks': 'The hours each maker booked over the last four weeks, to see who is busy and who has time.',
    'Orders to deliver': 'Open orders due in the next 14 days, and any already late, the soonest first, with what the goods come to.',
    'Late deliveries': 'Orders delivered after the promised day in the last year, with how many days late.',
    'Unpaid invoices': 'Orders delivered and not yet paid, the oldest due date first, with the days overdue. Below nought means it is not due yet.',
    'Finished goods in stock': 'Products on the shelf, with what they are worth at cost and at price.',
    'Sales by month': 'Goods delivered in each month, after discounts and without delivery fees.',
    'Sales by product': 'What each product has sold, the biggest seller first.',
    'Sales by customer': 'What each customer has bought, the biggest first, with how many order lines.',
  },
  reports: {
    'Costed bill of materials': 'A page for each product listing its materials and their cost, with the total material cost.',
    'Product margins': 'Every product with its price, costs and margin, for setting prices and deciding what to push.',
    'Production plan': 'The open work orders for each maker in the order they are planned to start, for the Monday meeting and the workshop wall.',
    'Shortages for buying': 'The materials short for planned work, by supplier, for the buyer to order first.',
    'Reorder list': 'Materials at or below their reorder level, by supplier, with how much to order.',
    'Delivery schedule': 'Orders to deliver in the next two weeks and those late, for planning the van.',
    'Monthly sales': 'Sales for each month, with a total.',
    'Invoices due': 'What each customer owes, with the due dates and days overdue, for chasing payment.',
    'Hours by maker': 'Each maker\'s booked hours and their cost month by month, with totals.',
  },
  empty: 'Start with Suppliers, then Materials with their units, costs and reorder levels, then Products, and then the Bill of Materials for each product. Fill in each product\'s Material Cost from the Costed bill of materials query. Add your Makers and Customers. After that, Sales Orders and their lines come in as you trade, Work Orders are raised for them, and the makers book Time Sheets as they work. Keep In Stock, On Order and Allocated up to date as deliveries arrive and work starts, or count once a month and correct them.',
  adapt: [
    'For a bakery or candle maker, change the units to kilogram and gram, make Quantity in Bill of Materials hold small amounts, and add a Batch Yield field to Products for how many one batch makes.',
    'Record finishes and options: add Finish and Size fields to Sales Lines for made-to-order pieces, and copy them to the work order.',
    'Track offcuts and waste: add a Waste Percent field to Bill of Materials and use it in the Costed bill of materials query.',
    'Add VAT or sales tax: add a Tax Rate field to Products and a calculated column in the sales queries for the tax in each line.',
    'Take deposits on commissions: add Deposit and Deposit Paid On fields to Sales Orders, and a query for orders waiting for one.',
    'Plan by machine as well as maker: add a Machines table and a Machine field to Work Orders, then group the production plan by it.',
    'Keep batch or lot numbers for traceability: add a Batch field to Material Movements and Stock Movements.',
  ],
};

/** Suppliers: [key, name, what they supply, lead days, minimum order]. Invented names. */
const SUPPLIERS = [
  ['timber', 'Hollybrook Timber Merchants', 'Timber', 6, 400],
  ['hardwood', 'Riverbend Hardwoods', 'Timber', 8, 250],
  ['sheet', 'Kestrel Board and Sheet', 'Sheet materials', 5, 150],
  ['fittings', 'Larkspur Fittings', 'Fittings and fixings', 5, 100],
  ['finishes', 'Thistle Coatings', 'Finishes and abrasives', 5, 120],
  ['packing', 'Cobalt Packaging', 'Packing', 5, 60],
];

/** Materials: [code, name, kind, unit, cost, supplier key, reorder level, reorder quantity]. */
const MATERIALS = [
  ['T-OAK27', 'Oak board, 27 x 150 mm', 'Timber', 'metre', 14.5, 'timber', 40, 100],
  ['T-OAK45', 'Oak board, 45 x 200 mm', 'Timber', 'metre', 32, 'timber', 25, 60],
  ['T-OAK70', 'Oak square, 70 x 70 mm', 'Timber', 'metre', 26, 'timber', 12, 30],
  ['T-ASH27', 'Ash board, 27 x 150 mm', 'Timber', 'metre', 9.8, 'timber', 20, 50],
  ['T-ASH45', 'Ash square, 45 x 45 mm', 'Timber', 'metre', 6.4, 'timber', 25, 60],
  ['T-TUL22', 'Tulipwood, 22 x 95 mm', 'Timber', 'metre', 4.2, 'timber', 40, 120],
  ['T-WAL27', 'Walnut board, 27 x 150 mm', 'Timber', 'metre', 28, 'hardwood', 20, 50],
  ['T-BEE32', 'Beech board, 32 x 200 mm', 'Timber', 'metre', 12.5, 'hardwood', 8, 20],
  ['S-MDF06', 'MDF sheet, 6 mm, 2440 x 1220', 'Sheet', 'sheet', 18.5, 'sheet', 2, 5],
  ['S-PLY18', 'Birch plywood, 18 mm, 2440 x 1220', 'Sheet', 'sheet', 72, 'sheet', 1, 2],
  ['S-VEN18', 'Oak veneered MDF, 18 mm, 2440 x 1220', 'Sheet', 'sheet', 95, 'sheet', 1, 2],
  ['F-HNG35', 'Soft-close hinge, 35 mm', 'Fittings', 'each', 3.2, 'fittings', 40, 100],
  ['F-RUN45', 'Drawer runners, full extension, 450 mm', 'Fittings', 'pair', 14.5, 'fittings', 2, 4],
  ['F-BTN', 'Table top fixing button', 'Fittings', 'each', 0.35, 'fittings', 15, 50],
  ['F-BRK15', 'Concealed shelf bracket, 150 mm', 'Fittings', 'each', 6.8, 'fittings', 10, 30],
  ['F-HOOK', 'Coat hook, black steel', 'Fittings', 'each', 2.1, 'fittings', 15, 40],
  ['F-KNOB', 'Cabinet knob, brushed steel', 'Fittings', 'each', 2.6, 'fittings', 10, 25],
  ['F-SCR', 'Wood screws, 4 x 40 mm, box of 200', 'Fixings', 'box', 7.5, 'fittings', 1, 2],
  ['F-DOW', 'Beech dowels, 8 mm, bag of 100', 'Fixings', 'bag', 4.2, 'fittings', 1, 3],
  ['X-GLUE', 'PVA wood glue', 'Finishes', 'litre', 5.8, 'finishes', 2, 5],
  ['X-HWO', 'Hardwax oil, clear', 'Finishes', 'litre', 24.8, 'finishes', 3, 6],
  ['X-BOIL', 'Food-safe board oil', 'Finishes', 'litre', 18, 'finishes', 2, 5],
  ['X-PRIM', 'Waterborne primer, white', 'Finishes', 'litre', 14, 'finishes', 2, 5],
  ['X-PNT', 'Waterborne eggshell paint', 'Finishes', 'litre', 22, 'finishes', 3, 6],
  ['X-ABR', 'Abrasive discs, 125 mm, mixed grits', 'Finishes', 'each', 0.45, 'finishes', 150, 400],
  ['P-CTN', 'Corrugated carton, small', 'Packing', 'each', 0.8, 'packing', 80, 200],
  ['P-WRAP', 'Bubble wrap, 750 mm x 50 m', 'Packing', 'roll', 21, 'packing', 2, 4],
];

/**
 * Products: [code, name, category, price, labour minutes, made to order,
 * stock level, batch size, easy (an apprentice can make it), bill of materials].
 */
const PRODUCTS = [
  ['TB-OAK18', 'Oak dining table, 1800 x 900', 'Tables', 1850, 1500, true, 0, 1, false,
    [['T-OAK45', 10], ['T-OAK70', 3], ['T-OAK27', 5], ['F-BTN', 10], ['F-DOW', 0.1], ['X-GLUE', 0.15], ['X-HWO', 0.4], ['X-ABR', 8], ['P-WRAP', 0.3]]],
  ['TB-OAK24', 'Oak dining table, 2400 x 1000', 'Tables', 2450, 1860, true, 0, 1, false,
    [['T-OAK45', 13.5], ['T-OAK70', 3.2], ['T-OAK27', 6.5], ['F-BTN', 12], ['F-DOW', 0.12], ['X-GLUE', 0.2], ['X-HWO', 0.55], ['X-ABR', 10], ['P-WRAP', 0.4]]],
  ['TB-WALC', 'Walnut coffee table, 1100 x 600', 'Tables', 980, 720, true, 0, 1, false,
    [['T-WAL27', 11], ['F-BTN', 8], ['X-GLUE', 0.08], ['X-HWO', 0.25], ['X-ABR', 6], ['P-WRAP', 0.2]]],
  ['TB-ASHS', 'Ash side table, 500 x 500', 'Tables', 340, 240, false, 3, 4, false,
    [['T-ASH27', 3], ['T-ASH45', 2.4], ['F-BTN', 4], ['X-GLUE', 0.04], ['X-HWO', 0.1], ['X-ABR', 3], ['P-WRAP', 0.1]]],
  ['SE-ASHC', 'Ash dining chair', 'Seating', 285, 300, false, 8, 8, false,
    [['T-ASH45', 4.2], ['T-ASH27', 1.6], ['F-DOW', 0.08], ['X-GLUE', 0.05], ['X-HWO', 0.12], ['X-ABR', 4], ['P-WRAP', 0.1]]],
  ['SE-BARS', 'Bar stool, oak seat, ash legs', 'Seating', 210, 210, false, 4, 4, false,
    [['T-OAK27', 1.2], ['T-ASH45', 3.2], ['F-DOW', 0.05], ['X-GLUE', 0.03], ['X-HWO', 0.08], ['X-ABR', 3], ['P-WRAP', 0.08]]],
  ['SE-STL3', 'Three-legged stool, ash', 'Seating', 95, 90, false, 6, 6, true,
    [['T-ASH27', 0.5], ['T-ASH45', 1.5], ['X-GLUE', 0.02], ['X-HWO', 0.04], ['X-ABR', 2], ['P-WRAP', 0.03]]],
  ['SE-BNCH', 'Oak bench, 1500 mm', 'Seating', 520, 360, false, 2, 2, false,
    [['T-OAK45', 3.2], ['T-OAK70', 1.6], ['F-DOW', 0.05], ['X-GLUE', 0.06], ['X-HWO', 0.15], ['X-ABR', 4], ['P-WRAP', 0.15]]],
  ['ST-SHF9', 'Floating shelf, oak, 900 mm', 'Storage', 95, 50, false, 12, 10, true,
    [['T-OAK45', 1], ['F-BRK15', 2], ['X-HWO', 0.03], ['X-ABR', 1], ['P-CTN', 1]]],
  ['ST-BOOK', 'Oak bookcase, 1800 x 900', 'Storage', 1150, 960, true, 0, 1, false,
    [['T-OAK27', 14], ['S-VEN18', 0.5], ['F-SCR', 0.1], ['F-DOW', 0.1], ['X-GLUE', 0.12], ['X-HWO', 0.35], ['X-ABR', 8], ['P-WRAP', 0.3]]],
  ['ST-PLYU', 'Wall shelf unit, birch plywood', 'Storage', 260, 150, false, 3, 4, true,
    [['S-PLY18', 0.6], ['F-SCR', 0.05], ['X-GLUE', 0.03], ['X-HWO', 0.1], ['X-ABR', 3], ['P-WRAP', 0.1]]],
  ['ST-BEDS', 'Bedside cabinet, oak', 'Storage', 480, 420, false, 2, 2, false,
    [['T-OAK27', 4.5], ['S-VEN18', 0.25], ['F-RUN45', 1], ['F-KNOB', 1], ['F-SCR', 0.05], ['X-GLUE', 0.05], ['X-HWO', 0.15], ['X-ABR', 4], ['P-WRAP', 0.15]]],
  ['SG-CBEG', 'Chopping board, oak end grain, 400 x 300', 'Small goods', 95, 70, false, 14, 12, true,
    [['T-OAK70', 0.9], ['X-GLUE', 0.03], ['X-BOIL', 0.04], ['X-ABR', 2], ['P-CTN', 1]]],
  ['SG-CBWA', 'Chopping board, walnut, 350 x 250', 'Small goods', 68, 40, false, 20, 16, true,
    [['T-WAL27', 0.75], ['X-BOIL', 0.03], ['X-ABR', 1], ['P-CTN', 1]]],
  ['SG-SRVA', 'Serving board, ash, with handle', 'Small goods', 34, 22, false, 25, 20, true,
    [['T-ASH27', 0.6], ['X-BOIL', 0.02], ['X-ABR', 0.5], ['P-CTN', 1]]],
  ['SG-COAS', 'Coasters, walnut, set of four', 'Small goods', 22, 12, false, 40, 40, true,
    [['T-WAL27', 0.15], ['X-BOIL', 0.01], ['X-ABR', 0.3], ['P-CTN', 1]]],
  ['SG-BRBE', 'Bread board, beech, 450 x 300', 'Small goods', 42, 30, false, 12, 12, true,
    [['T-BEE32', 0.9], ['X-BOIL', 0.03], ['X-ABR', 1], ['P-CTN', 1]]],
  ['SG-COAT', 'Coat rack, oak, five hooks', 'Small goods', 65, 45, false, 8, 8, true,
    [['T-OAK27', 1], ['F-HOOK', 5], ['F-SCR', 0.02], ['X-HWO', 0.02], ['X-ABR', 1], ['P-CTN', 1]]],
  ['SG-PLNT', 'Plant stand, ash, 600 mm', 'Small goods', 120, 80, false, 4, 4, true,
    [['T-ASH45', 2.5], ['T-ASH27', 0.6], ['X-GLUE', 0.02], ['X-HWO', 0.05], ['X-ABR', 2], ['P-WRAP', 0.05]]],
  ['KI-DSHP', 'Kitchen door, shaker, painted, 715 x 497', 'Kitchen', 145, 100, true, 0, 1, false,
    [['T-TUL22', 2.8], ['S-MDF06', 0.12], ['F-HNG35', 2], ['X-GLUE', 0.02], ['X-PRIM', 0.08], ['X-PNT', 0.12], ['X-ABR', 2]]],
  ['KI-FSHP', 'Drawer front, shaker, painted, 140 x 497', 'Kitchen', 75, 55, true, 0, 1, false,
    [['T-TUL22', 1.4], ['S-MDF06', 0.03], ['F-KNOB', 1], ['X-GLUE', 0.01], ['X-PRIM', 0.03], ['X-PNT', 0.05], ['X-ABR', 1]]],
  ['KI-DOAK', 'Kitchen door, solid oak, 715 x 497', 'Kitchen', 175, 90, true, 0, 1, false,
    [['T-OAK27', 4.2], ['F-HNG35', 2], ['F-DOW', 0.05], ['X-GLUE', 0.03], ['X-HWO', 0.06], ['X-ABR', 2]]],
  ['KI-FOAK', 'Drawer front, solid oak, 140 x 497', 'Kitchen', 85, 45, true, 0, 1, false,
    [['T-OAK27', 1], ['F-KNOB', 1], ['X-GLUE', 0.01], ['X-HWO', 0.02], ['X-ABR', 1]]],
  ['KI-WTOP', 'Oak worktop, 3000 x 620 x 40', 'Kitchen', 780, 300, true, 0, 1, false,
    [['T-OAK45', 10], ['F-DOW', 0.2], ['X-GLUE', 0.25], ['X-HWO', 0.4], ['X-ABR', 6], ['P-WRAP', 0.3]]],
];

// What each kind of customer orders: [product code, weight, fewest, most].
const SHOP_RANGE = [['SG-CBEG', 3, 4, 10], ['SG-CBWA', 4, 6, 12], ['SG-SRVA', 4, 6, 16], ['SG-COAS', 4, 10, 24], ['SG-BRBE', 3, 4, 10], ['SG-COAT', 2, 2, 6], ['ST-SHF9', 2, 2, 6], ['SE-STL3', 1, 2, 4], ['SG-PLNT', 1, 2, 4]];
const PUBLIC_STOCK = [['SE-ASHC', 3, 2, 6], ['SE-BARS', 3, 2, 4], ['SE-STL3', 2, 1, 2], ['SE-BNCH', 2, 1, 1], ['TB-ASHS', 2, 1, 2], ['ST-SHF9', 3, 1, 3], ['ST-PLYU', 2, 1, 1], ['ST-BEDS', 2, 1, 2], ['SG-CBEG', 3, 1, 1], ['SG-CBWA', 2, 1, 2], ['SG-BRBE', 2, 1, 1], ['SG-COAS', 2, 1, 2], ['SG-COAT', 2, 1, 1], ['SG-PLNT', 1, 1, 1]];
const PUBLIC_MADE = [['TB-OAK18', 4, 1, 1], ['TB-OAK24', 2, 1, 1], ['TB-WALC', 3, 1, 1], ['ST-BOOK', 2, 1, 2], ['KI-WTOP', 1, 1, 2]];
const DESIGNER_FURNITURE = [['TB-OAK18', 3, 1, 1], ['TB-OAK24', 2, 1, 1], ['TB-WALC', 2, 1, 1], ['ST-BOOK', 2, 1, 2], ['ST-BEDS', 2, 2, 2], ['SE-ASHC', 2, 6, 8], ['SE-BNCH', 1, 1, 2]];

const slug = (name) => name.toLowerCase().replace(/[^a-z]+/g, '');
const HOURS_A_DAY = 7.5;
const quarter = (h) => Math.round(h * 4) / 4;

/** Fill a Granite and Grain Woodworks database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 455 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };
  const update = (table, id, values) => base.update(table, id, values);
  const isWeekend = (d) => { const w = cal.weekday(d); return w === 0 || w === 6; };
  // Closed between Christmas Eve and New Year's Day.
  const closed = (d) => { const t = cal.at(d); return (t.getMonth() === 11 && t.getDate() >= 24) || (t.getMonth() === 0 && t.getDate() === 1); };
  const workday = (d) => !isWeekend(d) && !closed(d);
  // The first day the workshop is open on or after `d`, and the last on or before it.
  const nextOpen = (d) => { let o = d; while (!workday(o)) o++; return o; };
  const lastOpen = (d) => { let o = d; while (!workday(o)) o--; return o; };
  // Calendar days for a number of working days.
  const calendarDays = (workdays) => Math.ceil(workdays * 7 / 5);

  // Suppliers.
  const suppliers = new Map();
  for (const [key, name, supplies, lead, minimum] of SUPPLIERS) {
    const contact = person(r, { domain: `${slug(name)}.example` });
    const id = add('Suppliers', {
      Name: name, Supplies: supplies, Contact: `${contact['First Name']} ${contact['Last Name']}`, Phone: phoneFor(r, 'United Kingdom'),
      'E-mail': `orders@${slug(name)}.example`, City: r.pick(['Leeds', 'Bristol', 'Birmingham', 'Manchester', 'Glasgow', 'Sheffield']), 'Lead Days': lead, 'Minimum Order': minimum,
    });
    suppliers.set(key, { id, key, lead, minimum, waiting: null });
  }

  // Materials, with an opening stock counted the day before the records start.
  const materials = new Map();
  for (const [code, name, kind, unit, cost, sup, level, qty] of MATERIALS) {
    const m = { code, name, unit, cost, supplier: suppliers.get(sup), level, qty, stock: 0, onOrder: 0, allocated: 0, timber: kind === 'Timber' };
    m.id = add('Materials', { Code: code, Name: name, Kind: kind, Unit: unit, 'Unit Cost': cost, Supplier: m.supplier.id, 'Reorder Level': level, 'Reorder Quantity': qty });
    materials.set(code, m);
  }
  const moveMaterial = (m, d, kind, quantity, reference) => {
    m.stock = money(m.stock + quantity);
    add('Material Movements', { Material: m.id, When: cal.day(d), Kind: kind, Quantity: money(quantity), Reference: reference });
  };
  for (const m of materials.values()) moveMaterial(m, -days - 1, 'Counted', m.level + m.qty, 'Opening count');

  // Products and their bills of materials.
  const products = new Map();
  for (const [code, name, category, price, minutes, mto, level, batch, easy, bom] of PRODUCTS) {
    const p = { code, name, price, minutes, mto, level, batch, easy, active: true, stock: 0, wanted: 0, building: 0, bom: bom.map(([c, q]) => ({ m: materials.get(c), q })) };
    p.id = add('Products', { Code: code, Name: name, Category: category, Price: price, 'Labour Minutes': minutes, 'Made to Order': mto, 'Stock Level': level, 'Batch Size': batch, Active: true });
    for (const { m, q } of p.bom) add('Bill of Materials', { Product: p.id, Material: m.id, Quantity: q });
    products.set(code, p);
  }
  const moveStock = (p, d, kind, quantity, reference) => {
    p.stock += quantity;
    add('Stock Movements', { Product: p.id, When: cal.day(d), Kind: kind, Quantity: quantity, Reference: reference });
  };
  for (const p of products.values()) if (!p.mto) moveStock(p, -days - 1, 'Opening', p.level + Math.ceil(p.batch / 2), 'Opening count');
  // The plant stand stops being made part way through; what is left is sold off.
  const retiredOn = -Math.round(days * 0.45);

  // The makers: [role, hourly cost, pace (1 is the standard time), years in the workshop before the records start].
  const makers = [['Workshop Lead', 26, 0.9, 8], ['Cabinet Maker', 23, 1, 5], ['Cabinet Maker', 22, 1, 2], ['Chair Maker', 22, 1, 3], ['Apprentice', 13.5, 1.35, 0]].map(([role, rate, pace, years], i) => {
    const p = person(r, { domain: 'granitegrain.example' });
    const id = add('Makers', { ...p, Role: role, 'Hourly Cost': rate, Phone: `07700 900${String(310 + i * 7).padStart(3, '0')}`, Started: cal.day(-days - Math.round(years * 365) - r.int(30, 200)) });
    // Two weeks of holiday each, at different times of the year.
    const away = new Set();
    for (const start of [r.int(-days + 20, -days + 200), r.int(-200, -20)]) for (let k = 0; k < 7; k++) away.add(start + k);
    return { id, role, pace, apprentice: role === 'Apprentice', queue: [], away };
  });
  // Hours of work a maker has in hand: all of it, or only what is due to start within a week of `d`.
  const backlog = (m, d = null) => m.queue.reduce((n, w) => n + (d == null || w.start <= d + 7 ? w.remaining : 0), 0);
  const allBacklog = () => makers.reduce((n, m) => n + backlog(m), 0);

  // Plan a work order: give it to the least busy maker who can make it, set its materials aside.
  const plan = (p, qty, d, order = null, deliveryDue = null) => {
    const hours = qty * p.minutes / 60;
    // Work for an order is planned to finish a little before it is due; a batch for the shelf as soon as can be.
    const latest = deliveryDue != null ? deliveryDue - calendarDays(hours / HOURS_A_DAY) - 7 : d;
    const can = makers.filter((m) => p.easy || !m.apprentice);
    const maker = can.reduce((a, b) => (backlog(b, latest) < backlog(a, latest) ? b : a));
    const length = calendarDays(hours * maker.pace / HOURS_A_DAY);
    const start = nextOpen(Math.max(latest, d + 1 + calendarDays(backlog(maker, latest) / HOURS_A_DAY)));
    const due = deliveryDue != null ? lastOpen(deliveryDue - 2) : nextOpen(start + length + 3);
    const id = add('Work Orders', { Product: p.id, Quantity: qty, Maker: maker.id, 'Sales Order': order ? order.id : null, 'Planned Start': cal.day(start), 'Due Date': cal.day(due), Status: 'Planned' });
    const wo = { id, p, qty, maker, start, due, order, status: 'Planned', remaining: Math.max(0.25, quarter(hours * maker.pace * (0.85 + r.next() * 0.35))) };
    for (const { m, q } of p.bom) m.allocated = money(m.allocated + qty * q);
    if (!p.mto) p.building += qty;
    maker.queue.push(wo);
    return wo;
  };
  const release = (wo) => { for (const { m, q } of wo.p.bom) m.allocated = money(m.allocated - wo.qty * q); };
  const canStart = (wo) => wo.p.bom.every(({ m, q }) => m.stock >= money(wo.qty * q));

  // A day at the bench for each maker: carry on with the job in hand, or start
  // the next one due whose materials are in, and book the hours to it.
  const work = (d) => {
    for (const mk of makers) {
      // Now and then someone is off sick for a few days, and their work waits.
      if (r.chance(0.005)) for (let k = 0; k < r.int(3, 8); k++) mk.away.add(d + k);
      if (mk.away.has(d) || r.chance(0.02)) continue;
      let hours = HOURS_A_DAY;
      const booked = new Map();
      while (hours > 0) {
        let wo = mk.queue.find((w) => w.status === 'In progress');
        if (!wo) {
          wo = mk.queue.filter((w) => w.start <= d + 3).sort((a, b) => a.start - b.start || a.due - b.due).find(canStart);
          if (!wo) break;
          wo.status = 'In progress';
          release(wo);
          for (const { m, q } of wo.p.bom) moveMaterial(m, d, 'Used', -wo.qty * q, `WO ${wo.id}`);
          update('Work Orders', wo.id, { Started: cal.day(d), Status: 'In progress' });
        }
        const h = Math.min(hours, wo.remaining);
        hours -= h;
        wo.remaining -= h;
        booked.set(wo, (booked.get(wo) || 0) + h);
        if (wo.remaining <= 0) {
          wo.status = 'Done';
          mk.queue.splice(mk.queue.indexOf(wo), 1);
          if (!wo.p.mto) wo.p.building -= wo.qty;
          moveStock(wo.p, d, 'Made', wo.qty, `WO ${wo.id}`);
          update('Work Orders', wo.id, { Finished: cal.day(d), Status: 'Done' });
        }
      }
      for (const [wo, h] of booked) add('Time Sheets', { Maker: mk.id, 'Work Order': wo.id, 'Worked On': cal.day(d), Hours: h });
    }
  };

  // Customers: shops and designers to start with; members of the public as they come.
  // No two businesses share a first word, nor with the workshop or a supplier, so none is mistaken for another.
  const taken = new Set(['Granite', ...SUPPLIERS.map(([, name]) => name.split(' ')[0])]);
  const fresh = (names, k) => {
    const out = [];
    for (const n of names) {
      const w = n.split(' ')[0];
      if (out.length < k && !taken.has(w)) { taken.add(w); out.push(n); }
    }
    return out;
  };
  const shopNames = fresh(companies(r, 40, ['Homeware', 'Kitchen Shop', 'Living', 'Home Store', 'Gift Shop', 'Interiors']), 10);
  const designerNames = fresh(companies(r, 40, ['Design Studio', 'Interiors', 'Kitchens', 'Design']), 7);
  const business = (name, kind, terms, discount) => {
    const a = place(r, TRADE_TOWNS);
    const c = person(r);
    const id = add('Customers', {
      Name: name, Kind: kind, Contact: `${c['First Name']} ${c['Last Name']}`, Phone: phoneFor(r, 'United Kingdom'), 'E-mail': `accounts@${slug(name)}.example`,
      ...a, 'Payment Days': terms, 'Trade Discount': discount,
    });
    return { id, kind, terms, discount, slow: r.chance(0.25), weight: 0.5 + r.next() * 1.5 };
  };
  const shops = shopNames.map((n) => business(n, 'Shop', 30, 0.3));
  const designers = designerNames.map((n) => business(n, 'Designer', r.pick([14, 14, 30]), 0.1));
  const public_ = [];
  const member = () => {
    if (public_.length > 20 && r.chance(0.15)) return r.pick(public_);
    const p = person(r);
    const a = place(r, LOCAL_TOWNS);
    const c = { id: add('Customers', { Name: `${p['First Name']} ${p['Last Name']}`, Kind: 'Public', Phone: phoneFor(r, 'United Kingdom'), 'E-mail': p['E-mail'], ...a, 'Payment Days': 0, 'Trade Discount': 0 }), kind: 'Public', terms: 0, discount: 0, slow: false };
    public_.push(c);
    return c;
  };

  // Lines for an order from one of the ranges above.
  const choose = (range, k) => {
    const live = range.filter(([code]) => products.get(code).active);
    const out = [];
    for (let i = 0; i < k && live.length; i++) {
      const pick = r.weighted(live.map((x) => [x, x[1]]));
      live.splice(live.indexOf(pick), 1);
      out.push({ p: products.get(pick[0]), qty: r.int(pick[2], pick[3]) });
    }
    return out;
  };

  const open = []; // orders not yet delivered or cancelled
  const take = (d, customer, lines, dueIn, fee, note = null) => {
    const madeHours = lines.filter((l) => l.p.mto).reduce((n, l) => n + l.qty * l.p.minutes / 60, 0);
    // A made-to-order promise allows for the work already in the workshop.
    const quote = madeHours ? calendarDays((allBacklog() + madeHours) / 30) + 7 : 0;
    const due = nextOpen(d + Math.max(dueIn, quote));
    const o = { customer, lines, due, fee, wos: [], cancelAt: null };
    o.id = add('Sales Orders', { Customer: customer.id, 'Order Date': cal.day(d), 'Delivery Due': cal.day(due), Status: 'Open', 'Delivery Fee': fee, Notes: note });
    for (const l of lines) {
      add('Sales Lines', { Order: o.id, Product: l.p.id, Quantity: l.qty, 'Unit Price': l.p.price, Discount: customer.discount });
      if (l.p.mto) o.wos.push(plan(l.p, l.qty, d, o, due));
      else l.p.wanted += l.qty;
    }
    // A few members of the public change their minds within the first fortnight.
    if (customer.kind === 'Public' && o.wos.length && r.chance(0.06)) o.cancelAt = d + r.int(2, 12);
    open.push(o);
  };

  // A designer's kitchen: doors, drawer fronts and sometimes worktops, made to measure.
  const kitchen = (d, oak, worktops) => {
    const lines = [{ p: products.get(oak ? 'KI-DOAK' : 'KI-DSHP'), qty: r.int(10, 24) }, { p: products.get(oak ? 'KI-FOAK' : 'KI-FSHP'), qty: r.int(3, 9) }];
    if (worktops) lines.push({ p: products.get('KI-WTOP'), qty: r.int(1, 3) });
    take(d, r.pick(designers), lines, r.int(28, 42), 45, oak ? 'Solid oak kitchen, oiled.' : `Painted shaker kitchen, colour ${r.pick(['sage', 'off white', 'navy', 'charcoal', 'stone', 'putty'])}.`);
  };

  // The order book the records start with: orders taken in the weeks before.
  for (let k = 0; k < 6; k++) kitchen(-days - r.int(3, 20), r.chance(0.3), r.chance(0.3));
  for (let k = 0; k < 5; k++) take(-days - r.int(3, 20), member(), choose(PUBLIC_MADE, 1), r.int(35, 56), r.pick([0, 45, 65]));
  for (let k = 0; k < 3; k++) take(-days - r.int(1, 6), r.pick(shops), choose(SHOP_RANGE, r.int(2, 5)), r.int(10, 14), 15);

  const pending = []; // purchases on their way: { id, at, lines }
  const invoices = [];

  // The file is made first thing in the morning: today has its orders from
  // overnight and the work they need planned, but no hours at the bench yet,
  // no van out, no lorries in and no buying round.
  for (let d = -days; d <= 0; d++) {
    const month = cal.at(d).getMonth();
    const opens = workday(d);
    const atWork = opens && d < 0;

    // Deliveries from suppliers.
    if (atWork) {
      for (let i = pending.length - 1; i >= 0; i--) {
        const po = pending[i];
        if (po.at > d) continue;
        update('Material Purchases', po.id, { 'Received On': cal.day(d), Status: 'Received' });
        for (const l of po.lines) { l.m.onOrder = money(l.m.onOrder - l.n); moveMaterial(l.m, d, 'Received', l.n, `PO ${po.id}`); }
        pending.splice(i, 1);
      }
    }

    // A price rise from the timber merchants part way through.
    if (d === -Math.round(days * 0.4)) for (const m of materials.values()) if (m.timber) m.cost = money(m.cost * 1.06);
    if (d === retiredOn) products.get('SG-PLNT').active = false;

    // Cancellations, while the work has not started.
    for (const o of open) {
      if (o.cancelAt !== d || o.wos.some((w) => w.status !== 'Planned')) continue;
      for (const w of o.wos) { release(w); w.status = 'Cancelled'; w.maker.queue.splice(w.maker.queue.indexOf(w), 1); update('Work Orders', w.id, { Status: 'Cancelled' }); }
      for (const l of o.lines) if (!l.p.mto) l.p.wanted -= l.qty;
      o.cancelled = true;
      update('Sales Orders', o.id, { Status: 'Cancelled', Notes: 'Cancelled by the customer before work started.' });
    }

    if (atWork) {
      work(d);
      // The van goes out with what is ready and due within five days, the soonest due first.
      for (const o of open.filter((x) => !x.cancelled && d >= x.due - 5).sort((a, b) => a.due - b.due)) {
        if (o.wos.some((w) => w.status !== 'Done') || o.lines.some((l) => l.p.stock < l.qty)) continue;
        let goods = 0;
        for (const l of o.lines) {
          moveStock(l.p, d, 'Delivered', -l.qty, `SO ${o.id}`);
          if (!l.p.mto) l.p.wanted -= l.qty;
          goods += money(l.qty * l.p.price * (1 - o.customer.discount));
        }
        o.delivered = true;
        const amount = money(goods + o.fee);
        update('Sales Orders', o.id, { 'Delivered On': cal.day(d), Status: 'Delivered', 'Invoice Amount': amount, 'Invoice Due': cal.day(d + o.customer.terms) });
        invoices.push({ o, d });
      }
      for (let i = open.length - 1; i >= 0; i--) if (open[i].delivered || open[i].cancelled) open.splice(i, 1);
    }

    // New orders. Shops and designers on working days, the public any day.
    if (opens) {
      const shopSeason = [0.5, 0.8, 1, 1, 1, 0.9, 0.9, 0.9, 1, 1.3, 1.9, 1.1][month];
      for (let k = 0, n = count(r, 0.42 * shopSeason); k < n; k++) {
        const s = r.weighted(shops.map((x) => [x, x.weight]));
        const lines = choose(SHOP_RANGE, r.int(2, 5));
        const goods = lines.reduce((t, l) => t + l.qty * l.p.price * 0.7, 0);
        take(d, s, lines, r.int(10, 14), goods >= 500 ? 0 : 15);
      }
      const kitchenSeason = [1.2, 1.15, 1.15, 1.1, 1, 1, 0.9, 0.6, 1, 1.1, 1, 0.5][month];
      for (let k = 0, n = count(r, 0.2 * kitchenSeason); k < n; k++) kitchen(d, r.chance(0.3), r.chance(0.3));
      for (let k = 0, n = count(r, 0.08); k < n; k++) take(d, r.pick(designers), choose(DESIGNER_FURNITURE, r.int(1, 2)), r.int(28, 49), 45);
    }
    const publicSeason = [1.1, 0.9, 1.1, 1.1, 1.1, 1, 0.8, 0.8, 1, 1.1, 1.4, 1.3][month];
    for (let k = 0, n = count(r, 0.5 * publicSeason * (isWeekend(d) ? 0.6 : 1)); k < n; k++) {
      const c = member();
      if (r.chance(0.3)) {
        const lines = choose(PUBLIC_MADE, 1);
        // Half of those buying a dining table want chairs or a bench with it.
        if (/^TB-OAK/.test(lines[0].p.code) && r.chance(0.5)) {
          lines.push(r.chance(0.7) ? { p: products.get('SE-ASHC'), qty: r.pick([4, 6, 6, 8]) } : { p: products.get('SE-BNCH'), qty: 1 });
        }
        take(d, c, lines, r.int(35, 56), r.pick([0, 45, 65]));
      } else {
        take(d, c, choose(PUBLIC_STOCK, r.int(1, 2)), r.int(5, 10), r.pick([0, 0, 15, 25]));
      }
    }

    // Batches for the shelf: enough for what is ordered and the stock level.
    if (opens) {
      for (const p of products.values()) {
        if (p.mto) continue;
        const target = p.wanted + (p.active ? p.level : 0);
        const short = target - p.stock - p.building;
        if (short > 0) plan(p, Math.ceil(short / p.batch) * p.batch, d);
      }
    }

    // A stock count each quarter finds a little less timber than the books say, and a damaged piece.
    if (atWork && d % 91 === 0) {
      for (const m of r.some([...materials.values()].filter((x) => x.timber && x.stock > 10), 3)) moveMaterial(m, d, 'Counted', -money(r.int(5, 20) / 10), 'Stock count');
      const p = r.pick([...products.values()].filter((x) => !x.mto && x.stock > x.wanted + 1));
      if (p) moveStock(p, d, 'Adjusted', -1, 'Damaged in the showroom');
    }

    // The buyer orders on Mondays and Thursdays, for work due to start in
    // the next fortnight: timber for a job further off is bought nearer the
    // time, so Material shortages can show a shortage the buyer means to
    // cover later.
    const wd = cal.weekday(d);
    if (atWork && (wd === 1 || wd === 4)) {
      const soon = new Map();
      for (const mk of makers) for (const w of mk.queue) {
        if (w.status !== 'Planned' || w.start > d + 14) continue;
        for (const { m, q } of w.p.bom) soon.set(m, (soon.get(m) || 0) + w.qty * q);
      }
      const bySupplier = new Map();
      for (const m of materials.values()) {
        const free = m.stock + m.onOrder - (soon.get(m) || 0);
        if (free > m.level) continue;
        const n = Math.ceil(m.qty + Math.max(0, -free));
        if (!bySupplier.has(m.supplier)) bySupplier.set(m.supplier, []);
        bySupplier.get(m.supplier).push({ m, n, short: free < 0 });
      }
      for (const [s, lines] of bySupplier) {
        const value = lines.reduce((t, l) => t + l.n * l.m.cost, 0);
        if (s.waiting == null) s.waiting = d;
        // Wait for the minimum order unless work is held up or it has waited a fortnight.
        if (value < s.minimum && !lines.some((l) => l.short) && d - s.waiting < 14) continue;
        s.waiting = null;
        const at = nextOpen(d + s.lead + r.int(0, 3));
        const id = add('Material Purchases', { Supplier: s.id, 'Ordered On': cal.day(d), Expected: cal.day(nextOpen(d + s.lead)), Status: 'Ordered' });
        for (const l of lines) { add('Purchase Lines', { Purchase: id, Material: l.m.id, Quantity: l.n, 'Unit Cost': l.m.cost }); l.m.onOrder = money(l.m.onOrder + l.n); }
        pending.push({ id, at, lines });
      }
      for (const s of suppliers.values()) if (!bySupplier.has(s)) s.waiting = null;
    }
  }

  // Something is always on its way: when the last round ordered nothing, the
  // buyer topped up from the supplier whose materials were lowest.
  if (!pending.length) {
    let d = -1;
    while (!workday(d) || ![1, 4].includes(cal.weekday(d))) d--;
    const low = [...materials.values()].sort((a, b) => (a.stock + a.onOrder - a.allocated) / a.level - (b.stock + b.onOrder - b.allocated) / b.level);
    const s = low[0].supplier;
    const id = add('Material Purchases', { Supplier: s.id, 'Ordered On': cal.day(d), Expected: cal.day(nextOpen(d + s.lead)), Status: 'Ordered' });
    for (const m of low.filter((x) => x.supplier === s).slice(0, 3)) {
      add('Purchase Lines', { Purchase: id, Material: m.id, Quantity: m.qty, 'Unit Cost': m.cost });
      m.onOrder = money(m.onOrder + m.qty);
    }
  }

  // A large oak kitchen or two came in overnight, as happens most weeks,
  // with timber to buy for it at the next round.
  const short = () => [...materials.values()].some((m) => m.allocated - m.stock - m.onOrder > 0.005);
  for (let k = 0; k < 4 && !short(); k++) kitchen(0, true, true);

  // Payment: the public pays on delivery, most trade customers in their terms, a few late.
  for (const { o, d } of invoices) {
    const c = o.customer;
    let at;
    if (c.kind === 'Public') at = r.chance(0.85) ? d : d + r.int(1, 6);
    else if (c.slow && r.chance(0.6)) at = d + c.terms + r.int(8, 40);
    else at = d + r.int(Math.max(3, c.terms - 12), c.terms + 2);
    if (at <= 0) update('Sales Orders', o.id, { 'Paid On': cal.day(at) });
  }

  for (const m of materials.values()) update('Materials', m.id, { 'Unit Cost': m.cost, 'In Stock': m.stock, 'On Order': m.onOrder, Allocated: m.allocated });
  for (const p of products.values()) {
    const cost = money(p.bom.reduce((t, { m, q }) => t + money(q * m.cost), 0));
    update('Products', p.id, { 'In Stock': p.stock, 'Material Cost': cost, Active: p.active, ...(p.active ? {} : { 'Stock Level': 0 }) });
  }
  return counts;
}

/** How many of something happen today, when on average `rate` do. */
function count(r, rate) {
  const whole = Math.floor(rate);
  return whole + (r.chance(rate - whole) ? 1 : 0);
}

// Towns the trade customers are in, and the towns near the workshop the
// public buy from, each with the letters its postcodes start with.
const TRADE_TOWNS = [['Leeds', ['LS']], ['Manchester', ['M']], ['Bristol', ['BS']], ['London', ['N', 'SE', 'EC', 'W']], ['Birmingham', ['B']], ['Cardiff', ['CF']], ['Glasgow', ['G']]];
const LOCAL_TOWNS = [['Leeds', ['LS']], ['Harrogate', ['HG']], ['York', ['YO']], ['Wakefield', ['WF']], ['Bradford', ['BD']], ['Huddersfield', ['HD']]];

/** An invented street address in one of `towns`, with a postcode that fits the town. */
function place(r, towns) {
  const [city, letters] = r.pick(towns);
  const a = address(r, [{ country: 'United Kingdom', cities: [city], post: () => `${r.pick(letters)}${r.int(1, 19)} ${r.int(1, 9)}${r.pick(['AB', 'DE', 'FG', 'HJ', 'NP', 'QR', 'TU', 'WX'])}` }]);
  return { Address: a.Address, City: a.City, Postcode: a.Postcode };
}
