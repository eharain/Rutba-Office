// The trading company: Northbank Trading, an invented wholesaler of
// specialty foods and household goods selling to shops, cafes and hotels in
// ten countries. Customers, suppliers and carriers in one Companies table,
// their contacts, the staff, products in categories, orders and their lines,
// purchase orders and theirs, every movement of stock, and invoices.
//
// Its records come from eighteen months of trading, run a day at a time:
// customers order more as the year ends, stock goes down as orders ship, a
// weekly check raises purchase orders that arrive some days later, and
// invoices are paid on time, late or not yet. The questions a wholesaler
// asks are its saved queries; its reports print them.

import { address, BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, companies, money, person, phoneFor, PLACES } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const LINE = 'Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)';
// The business's own details at the head of every document.
const LETTERHEAD = { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Website', 'Tax Number'] };
const VAT = 'Round(Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2) * Nz([Tax Rate]), 2)';

export const spec = {
  // Where a value below nought is right: stock going out, an order already late.
  negative: { 'Stock Movements': ['Quantity'], 'Orders to ship': ['Days Left'] },
  tables: [
    BUSINESS_TABLE,
    { name: 'Companies', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Kind', type: 'text', default: 'Customer' },
      { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' }, { name: 'Country', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Payment Days', type: 'integer', default: 30 },
      { name: 'Credit Limit', type: 'currency' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Contacts', fields: [
      { name: 'Company', ...ref('Companies'), required: true }, { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
      { name: 'Job Title', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' },
    ] },
    { name: 'Employees', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Job Title', type: 'text' },
      { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'Hired', type: 'date' },
    ] },
    // Orders come before Products: a query joins two tables by the first relationship between them,
    // so an order's Companies record is its customer, not its products' supplier.
    { name: 'Orders', fields: [
      { name: 'Customer', ...ref('Companies'), required: true }, { name: 'Employee', ...ref('Employees') }, { name: 'Order Date', type: 'date' },
      { name: 'Required Date', type: 'date' }, { name: 'Shipped Date', type: 'date' }, { name: 'Shipper', ...ref('Companies') },
      { name: 'Shipping Fee', type: 'currency' }, { name: 'Status', type: 'text', default: 'New' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Order Lines', fields: [
      { name: 'Order', ...ref('Orders'), required: true }, { name: 'Product', ...ref('Products'), required: true },
      { name: 'Quantity', type: 'integer', default: 1 }, { name: 'Unit Price', type: 'currency' }, { name: 'Discount', type: 'number', default: 0 },
      { name: 'Tax Rate', type: 'number', default: 0 }, { name: 'Invoice', ...ref('Invoices') },
    ] },
    { name: 'Categories', fields: [{ name: 'Name', type: 'text', required: true, unique: true }, { name: 'Description', type: 'text' }, { name: 'Tax Rate', type: 'number', default: 0 }] },
    { name: 'Products', fields: [
      { name: 'Code', type: 'text', unique: true }, { name: 'Name', type: 'text', required: true }, { name: 'Category', ...ref('Categories') },
      { name: 'Supplier', ...ref('Companies') }, { name: 'Unit', type: 'text' }, { name: 'Cost Price', type: 'currency' }, { name: 'List Price', type: 'currency' },
      { name: 'Units in Stock', type: 'integer', default: 0 }, { name: 'Reorder Level', type: 'integer', default: 10 },
      { name: 'Reorder Quantity', type: 'integer', default: 40 }, { name: 'Discontinued', type: 'yesno' },
    ] },
    { name: 'Purchase Orders', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Supplier', ...ref('Companies'), required: true }, { name: 'Employee', ...ref('Employees') }, { name: 'Ordered On', type: 'date' },
      { name: 'Expected', type: 'date' }, { name: 'Received On', type: 'date' }, { name: 'Status', type: 'text', default: 'Submitted' },
    ] },
    { name: 'Purchase Lines', fields: [
      { name: 'Purchase Order', ...ref('Purchase Orders'), required: true }, { name: 'Product', ...ref('Products'), required: true },
      { name: 'Quantity', type: 'integer' }, { name: 'Unit Cost', type: 'currency' },
    ] },
    { name: 'Stock Movements', fields: [
      { name: 'Product', ...ref('Products'), required: true }, { name: 'When', type: 'date' }, { name: 'Kind', type: 'text' },
      { name: 'Quantity', type: 'integer' }, { name: 'Reference', type: 'text' },
    ] },
    { name: 'Invoices', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Order', ...ref('Orders'), required: true }, { name: 'Customer', ...ref('Companies') },
      { name: 'Invoice Date', type: 'date' }, { name: 'Due Date', type: 'date' },
      { name: 'Net', type: 'currency' }, { name: 'Tax', type: 'currency' }, { name: 'Amount', type: 'currency' }, { name: 'Paid On', type: 'date' },
    ] },
  ],
  queries: {
    'Order totals': { tables: ['Orders', 'Companies', 'Order Lines'], columns: [
      { table: 'Orders', field: 'ID', as: 'Order', total: 'group' }, { table: 'Orders', field: 'Order Date', total: 'group', sort: 'desc' },
      { table: 'Companies', field: 'Name', as: 'Customer', total: 'group' }, { table: 'Orders', field: 'Status', total: 'group' },
      { expr: LINE, as: 'Goods', total: 'sum' },
    ] },
    'Sales by month': { tables: ['Orders', 'Order Lines'], columns: [
      { expr: 'Format([Order Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { expr: LINE, as: 'Sales', total: 'sum' }, { table: 'Order Lines', field: 'Quantity', as: 'Units', total: 'sum' },
      { table: 'Orders', field: 'Status', total: 'where', criteria: ['<> "Cancelled"'] },
    ] },
    'Sales by category': { tables: ['Order Lines', 'Orders', 'Products', 'Categories'], columns: [
      { table: 'Categories', field: 'Name', as: 'Category', total: 'group' }, { expr: LINE, as: 'Sales', total: 'sum', sort: 'desc' },
      { table: 'Order Lines', field: 'Quantity', as: 'Units', total: 'sum' }, { table: 'Orders', field: 'Status', total: 'where', criteria: ['<> "Cancelled"'] },
    ] },
    'Sales by employee': { tables: ['Orders', 'Employees', 'Order Lines'], columns: [
      { expr: '[First Name] & " " & [Last Name]', as: 'Employee', total: 'group' }, { expr: LINE, as: 'Sales', total: 'sum', sort: 'desc' },
      { table: 'Orders', field: 'Status', total: 'where', criteria: ['<> "Cancelled"'] },
    ] },
    'Top customers': { tables: ['Companies', 'Orders', 'Order Lines'], columns: [
      { table: 'Companies', field: 'Name', as: 'Customer', total: 'group' }, { table: 'Companies', field: 'Country', total: 'group' },
      { expr: LINE, as: 'Sales', total: 'sum', sort: 'desc' }, { table: 'Orders', field: 'Status', total: 'where', criteria: ['<> "Cancelled"'] },
    ] },
    'Products to reorder': { tables: ['Products', 'Companies'], columns: [
      { table: 'Companies', field: 'Name', as: 'Supplier', sort: 'asc' }, { table: 'Products', field: 'Code' }, { table: 'Products', field: 'Name', as: 'Product' },
      { table: 'Products', field: 'Units in Stock' }, { table: 'Products', field: 'Reorder Level' }, { table: 'Products', field: 'Reorder Quantity' },
      { expr: '[Units in Stock] - [Reorder Level]', as: 'Above Reorder Level', criteria: ['<= 0'], show: false },
      { table: 'Products', field: 'Discontinued', criteria: ['No'], show: false },
    ] },
    'Orders to ship': { tables: ['Orders', 'Companies'], columns: [
      { table: 'Orders', field: 'ID', as: 'Order' }, { table: 'Orders', field: 'Order Date', sort: 'asc' }, { table: 'Companies', field: 'Name', as: 'Customer' },
      { table: 'Orders', field: 'Required Date' }, { expr: 'DateDiff("d", Date(), [Required Date])', as: 'Days Left' },
      { table: 'Orders', field: 'Shipped Date', criteria: ['Is Null'], show: false }, { table: 'Orders', field: 'Status', criteria: ['<> "Cancelled"'], show: false },
    ] },
    'Unpaid invoices': { tables: ['Invoices', 'Orders', 'Companies'], columns: [
      { table: 'Invoices', field: 'ID', as: 'Invoice' }, { table: 'Companies', field: 'Name', as: 'Customer' }, { table: 'Invoices', field: 'Invoice Date' },
      { table: 'Invoices', field: 'Due Date', sort: 'asc' }, { table: 'Invoices', field: 'Amount' },
      { expr: 'IIf(DateDiff("d", [Due Date], Date()) > 0, DateDiff("d", [Due Date], Date()), 0)', as: 'Days Overdue' },
      { table: 'Invoices', field: 'Paid On', criteria: ['Is Null'], show: false },
    ] },
    'Stock value': { tables: ['Products', 'Categories'], columns: [
      { table: 'Categories', field: 'Name', as: 'Category', total: 'group', sort: 'asc' }, { table: 'Products', field: 'Units in Stock', as: 'Units', total: 'sum' },
      { expr: 'Round([Units in Stock] * [Cost Price], 2)', as: 'Value at Cost', total: 'sum' },
    ] },
    'Product list': { tables: ['Products', 'Categories'], columns: [
      { table: 'Categories', field: 'Name', as: 'Category', sort: 'asc' }, { table: 'Products', field: 'Code', sort: 'asc' }, { table: 'Products', field: 'Name', as: 'Product' },
      { table: 'Products', field: 'Unit' }, { table: 'Products', field: 'List Price' }, { table: 'Products', field: 'Discontinued', criteria: ['No'], show: false },
    ] },
    'Purchases on order': { tables: ['Purchase Orders', 'Companies'], columns: [
      { table: 'Purchase Orders', field: 'ID', as: 'Purchase Order' }, { table: 'Companies', field: 'Name', as: 'Supplier' },
      { table: 'Purchase Orders', field: 'Ordered On' }, { table: 'Purchase Orders', field: 'Expected', sort: 'asc' },
      { table: 'Purchase Orders', field: 'Status', criteria: ['Submitted'] },
    ] },
  },
  // The print templates: each a document for one record, the business's details at the head.
  documents: [
    { name: 'Invoice', table: 'Invoices', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Invoice {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Order.Customer', label: 'Invoice to' }, { field: 'Number', label: 'Invoice number' },
        { field: 'Order.Customer.Address', label: 'Address' }, { field: 'Invoice Date', label: 'Date' },
        { field: 'Order.Customer.City', label: 'City' }, { field: 'Due Date', label: 'Payment due' },
        { field: 'Order.Customer.Postcode', label: 'Postcode' }, { field: 'Order.ID', label: 'Your order' },
        { field: 'Order.Customer.Country', label: 'Country' }, { field: 'Order.Shipped Date', label: 'Shipped' },
      ] },
      { type: 'lines', table: 'Order Lines', link: 'Invoice', columns: [
        { field: 'Product.Code', label: 'Code' }, { field: 'Product.Name', label: 'Product' }, { field: 'Product.Unit', label: 'Unit' },
        { field: 'Quantity', type: 'integer' }, { field: 'Unit Price', type: 'currency' },
        { expr: 'Nz([Discount]) * 100', as: 'Discount %', type: 'number' }, { expr: LINE, as: 'Line Net', label: 'Net', type: 'currency' },
        { expr: 'Nz([Tax Rate]) * 100', as: 'VAT %', type: 'number' }, { expr: VAT, as: 'Line VAT', label: 'VAT', type: 'currency' },
      ] },
      { type: 'totals', rows: [
        { label: 'Goods', expr: '[Line Net]' }, { label: 'Delivery', expr: '[Net] - [Line Net]' },
        { label: 'VAT', expr: '[Tax]' }, { label: 'Total due', expr: '[Amount]', strong: true },
      ] },
      { type: 'text', text: '{Business.Payment Terms} Please pay to {Business.Bank Details}, quoting {Number}.', small: true },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    { name: 'Delivery note', table: 'Orders', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Delivery note DN-{ID}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Customer', label: 'Deliver to' }, { field: 'ID', label: 'Order' },
        { field: 'Customer.Address', label: 'Address' }, { field: 'Order Date', label: 'Ordered' },
        { field: 'Customer.City', label: 'City' }, { field: 'Shipped Date', label: 'Shipped' },
        { field: 'Customer.Postcode', label: 'Postcode' }, { field: 'Shipper', label: 'Carrier' },
        { field: 'Customer.Country', label: 'Country' }, { field: 'Notes', label: 'Instructions' },
      ] },
      { type: 'lines', table: 'Order Lines', link: 'Order', columns: [
        { field: 'Product.Code', label: 'Code' }, { field: 'Product.Name', label: 'Product' }, { field: 'Product.Unit', label: 'Unit' }, { field: 'Quantity', type: 'integer' },
      ] },
      { type: 'totals', rows: [{ label: 'Units', expr: '[Quantity]', type: 'integer', strong: true }] },
      { type: 'text', text: 'Please check the goods against this note. Shortages or damage to be reported within 48 hours of delivery.', small: true },
      { type: 'signatures', labels: ['Received in good order by', 'Date'] },
    ] },
    { name: 'Purchase order', table: 'Purchase Orders', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Purchase order {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Supplier', label: 'To' }, { field: 'Number', label: 'Order number' },
        { field: 'Supplier.Address', label: 'Address' }, { field: 'Ordered On', label: 'Date' },
        { field: 'Supplier.City', label: 'City' }, { field: 'Expected', label: 'Deliver by' },
        { field: 'Supplier.Country', label: 'Country' }, { field: 'Employee', label: 'Ordered by' },
      ] },
      { type: 'lines', table: 'Purchase Lines', columns: [
        { field: 'Product.Code', label: 'Code' }, { field: 'Product.Name', label: 'Product' }, { field: 'Product.Unit', label: 'Unit' },
        { field: 'Quantity', type: 'integer' }, { field: 'Unit Cost', type: 'currency' }, { expr: 'Round([Quantity] * [Unit Cost], 2)', as: 'Line Total', label: 'Total', type: 'currency' },
      ] },
      { type: 'totals', rows: [{ label: 'Order total', expr: '[Line Total]', strong: true }] },
      { type: 'text', text: 'Please quote {Number} on your delivery note and invoice. Deliver to {Business.Address}, {Business.City} {Business.Postcode}.', small: true },
    ] },
    { name: 'Statement', table: 'Companies', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Statement of account' },
      { type: 'fields', columns: 2, items: [
        { field: 'Name', label: 'Account' }, { field: 'Payment Days', label: 'Terms (days)' },
        { field: 'Address', label: 'Address' }, { field: 'City', label: 'City' }, { field: 'Postcode', label: 'Postcode' }, { field: 'Country', label: 'Country' },
      ] },
      { type: 'lines', table: 'Invoices', link: 'Customer', sort: 'Invoice Date', columns: [
        { field: 'Number', label: 'Invoice' }, { field: 'Invoice Date', label: 'Date' }, { field: 'Due Date', label: 'Due' },
        { field: 'Amount', type: 'currency' }, { field: 'Paid On', label: 'Paid' },
        { expr: 'IIf(Nz([Paid On], "") = "", [Amount], 0)', as: 'Owing', type: 'currency' },
      ] },
      { type: 'totals', rows: [{ label: 'Invoiced', expr: '[Amount]' }, { label: 'Balance owing', expr: '[Owing]', strong: true }] },
      { type: 'text', text: 'Please pay the balance owing to {Business.Bank Details}. If you have paid since this statement was made, thank you.', small: true },
    ] },
  ],
  reports: [
    { name: 'Product catalogue', query: 'Product list', group: 'Category', columns: ['Code', 'Product', 'Unit', { field: 'List Price', type: 'currency' }] },
    { name: 'Monthly sales', query: 'Sales by month', columns: ['Month', { field: 'Units', type: 'integer' }, { field: 'Sales', type: 'currency' }], totals: [{ field: 'Sales', fn: 'sum' }, { field: 'Units', fn: 'sum' }] },
    { name: 'Orders by customer', query: 'Order totals', group: 'Customer', sort: 'Order Date', desc: true, columns: ['Order', { field: 'Order Date', type: 'date' }, 'Status', { field: 'Goods', type: 'currency' }], totals: [{ field: 'Goods', fn: 'sum' }] },
    { name: 'Invoices due', query: 'Unpaid invoices', group: 'Customer', columns: ['Invoice', { field: 'Invoice Date', type: 'date' }, { field: 'Due Date', type: 'date' }, { field: 'Days Overdue', type: 'integer' }, { field: 'Amount', type: 'currency' }], totals: [{ field: 'Amount', fn: 'sum' }] },
    { name: 'Stock to reorder', query: 'Products to reorder', group: 'Supplier', columns: ['Code', 'Product', { field: 'Units in Stock', type: 'integer' }, { field: 'Reorder Level', type: 'integer' }, { field: 'Reorder Quantity', type: 'integer' }] },
  ],
};

export const about = {
  name: 'Northbank Trading',
  kind: 'Wholesale and trading',
  line: 'Customers, suppliers, products, orders, purchase orders, stock and invoices for a wholesaler.',
  summary: [
    'Northbank Trading is an invented wholesaler of specialty foods and supplies for the trade: teas and coffee, spices, rice and grains, dried fruit and nuts, oils, sauces, baking goods and kitchen supplies. It sells to shops, delis, cafes, restaurants and hotels in ten countries, and buys from a dozen suppliers.',
    'The sample holds eighteen months of trading, made up a day at a time. Customers order more as the year ends. Each order is cut to what is on the shelf, so stock never goes below nought, and it goes down as orders ship. Twice a week the buyer raises purchase orders for whatever has fallen to its reorder level, and those arrive some days later. Every shipped order is invoiced; most invoices are paid on time, some late, and the newest are not yet due.',
    'Customers, suppliers and carriers share one Companies table, told apart by their Kind, so a company that both buys from you and sells to you is entered once.',
  ],
  suits: ['A distributor or importer', 'A cash and carry', 'A builders\' merchant or parts supplier', 'A wine merchant selling to the trade', 'An online shop that sells to businesses', 'A manufacturer\'s sales office'],
  tables: {
    Business: BUSINESS_ABOUT,
    Companies: { about: 'Every business you deal with: customers, suppliers and the carriers who deliver for you.', fields: {
      Name: 'The company\'s name.', Kind: 'Customer, Supplier or Shipper.', Address: 'The street address.', City: 'The town or city.',
      Postcode: 'The postcode or ZIP code.', Country: 'The country.', Phone: 'The main phone number.', 'E-mail': 'Where orders or invoices go.',
      'Payment Days': 'How many days after the invoice payment is due: 30 means thirty days.', 'Credit Limit': 'The most a customer may owe at once.',
      Notes: 'Anything worth remembering about the company.',
    } },
    Contacts: { about: 'The people at each company.', fields: {
      Company: 'The company they work for.', 'First Name': 'Their first name.', 'Last Name': 'Their family name.',
      'Job Title': 'Their job, such as Buyer or Account Manager.', 'E-mail': 'Their e-mail address.', Phone: 'Their phone number.',
    } },
    Employees: { about: 'Your own staff: sales, buying, the warehouse and accounts.', fields: {
      'First Name': 'Their first name.', 'Last Name': 'Their family name.', 'Job Title': 'Their job.', 'E-mail': 'Their work e-mail address.',
      Phone: 'Their work phone number.', Hired: 'The day they joined.',
    } },
    Categories: { about: 'The groups your products are sold in.', fields: { Name: 'The group\'s name. Each name is used once.', Description: 'What the group holds.', 'Tax Rate': 'The VAT rate its products are sold at, as a fraction: 0.2 is twenty percent, 0 for food that carries none.' } },
    Products: { about: 'What you sell, one record for each line in the catalogue.', fields: {
      Code: 'Your own product code. Each code is used once.', Name: 'The product as it appears on the catalogue and invoices.',
      Category: 'The group it is sold in.', Supplier: 'Who you buy it from.', Unit: 'How it is sold, such as a 10 kg sack or a box of 100.',
      'Cost Price': 'What one unit costs you.', 'List Price': 'What one unit sells for before any discount.',
      'Units in Stock': 'How many units are in the warehouse now.', 'Reorder Level': 'When stock falls to this level, the product is ordered again.',
      'Reorder Quantity': 'How many to order each time.', Discontinued: 'Ticked for a product you no longer sell; old orders keep it.',
    } },
    Orders: { about: 'Customers\' orders, one record each.', fields: {
      Customer: 'Who ordered.', Employee: 'The sales person who took it.', 'Order Date': 'The day it was placed.',
      'Required Date': 'When the customer needs it.', 'Shipped Date': 'The day it left the warehouse; empty until then.',
      Shipper: 'The carrier taking it.', 'Shipping Fee': 'What the customer pays for delivery.',
      Status: 'New (not shipped yet), Shipped, Closed (shipped and paid) or Cancelled.', Notes: 'Delivery instructions or anything else.',
    } },
    'Order Lines': { about: 'The products on each order.', fields: {
      Order: 'The order the line belongs to.', Product: 'What was ordered.', Quantity: 'How many units.',
      'Unit Price': 'The price of one unit on this order, kept so a later price change does not alter it.', Discount: 'The discount as a fraction: 0.05 is five percent.',
      'Tax Rate': 'The VAT rate of the line, taken from its product\'s group when the order was taken.', Invoice: 'The invoice the line went on, once the order shipped and was invoiced.',
    } },
    'Purchase Orders': { about: 'Orders you place with your suppliers.', fields: {
      Number: 'The purchase order\'s number, printed on it and quoted by the supplier: PO-2026-0042.', Supplier: 'Who you ordered from.', Employee: 'Who placed the order.', 'Ordered On': 'The day it was sent.',
      Expected: 'When the goods should arrive.', 'Received On': 'The day they arrived; empty until then.', Status: 'Submitted or Received.',
    } },
    'Purchase Lines': { about: 'The products on each purchase order.', fields: {
      'Purchase Order': 'The purchase order the line belongs to.', Product: 'What was ordered.', Quantity: 'How many units.', 'Unit Cost': 'What one unit costs on this order.',
    } },
    'Stock Movements': { about: 'Every change to stock: goods received, goods sold and corrections from a stock count. Added up, they give what is on the shelf.', fields: {
      Product: 'The product that moved.', When: 'The day it moved.', Kind: 'Received, Sold or Adjusted.',
      Quantity: 'How many: above nought coming in, below nought going out.', Reference: 'The order or purchase order it came from, or Stock count.',
    } },
    Invoices: { about: 'An invoice for each shipped order.', fields: {
      Number: 'The invoice\'s number, given in date order and once only: INV-2026-0042, a new run each year.', Order: 'The order invoiced.', Customer: 'Who the invoice is to: the order\'s customer, kept here so a statement can list a customer\'s invoices.', 'Invoice Date': 'The day it was raised, the day the order shipped.', 'Due Date': 'When payment is due, from the customer\'s payment days.',
      Net: 'The goods after discounts, and delivery, before VAT.', Tax: 'The VAT on the goods and the delivery.', Amount: 'What the invoice comes to: net and VAT together.', 'Paid On': 'The day it was paid; empty while it is owed.',
    } },
  },
  queries: {
    'Order totals': 'Every order with its customer, status and the value of its goods after discounts, the newest first.',
    'Sales by month': 'Sales and units for each month, cancelled orders left out.',
    'Sales by category': 'What each product group has sold, the biggest first.',
    'Sales by employee': 'What each sales person has sold.',
    'Top customers': 'Customers by what they have bought, with their country.',
    'Products to reorder': 'Products at or below their reorder level, by supplier, with how many to order. Discontinued lines are left out.',
    'Orders to ship': 'Orders not shipped yet, oldest first, with the days left before the customer needs them.',
    'Unpaid invoices': 'Invoices not yet paid, by due date, with how many days each is overdue (nought when it is not due yet).',
    'Stock value': 'Units in stock and their value at cost, by product group.',
    'Product list': 'The products you still sell, by group and code, with their list prices.',
    'Purchases on order': 'Purchase orders sent and not yet received, by the day they are expected.',
  },
  reports: {
    'Product catalogue': 'The price list: every product you still sell by group, with its unit and list price.',
    'Monthly sales': 'One line a month with units and sales, and the totals.',
    'Orders by customer': 'Each customer\'s orders, the newest first, with a total for each customer.',
    'Invoices due': 'Unpaid invoices by customer with their due dates, days overdue and a total owed: the list for chasing payment.',
    'Stock to reorder': 'What to order, on a page for each supplier.',
  },
  documents: {
    Invoice: 'The tax invoice sent when an order ships: the business\'s details at the head, the customer\'s address, each line with its VAT rate, and the goods, delivery, VAT and total due. Numbered INV-2026-0042.',
    'Delivery note': 'Goes in the box with a shipped order: what was sent and how many, without prices, for the customer to check and sign.',
    'Purchase order': 'Sent to a supplier: what to send, at what cost, and when it is expected. Numbered PO-2026-0042.',
    Statement: 'Sent to a customer each month: their invoices, which are paid and when, and the balance owing. It prints for every company; print it for a customer from the Companies form.',
  },
  tutorial: [
    { title: 'Take an order', steps: [
      'Open the Orders form and press New. Choose the customer in Customer and yourself in Employee; type today in Order Date and the day they need it in Required Date.',
      'Choose the carrier in Shipper and type any delivery charge in Shipping Fee. Leave Status as New.',
      'Open the Order Lines form and add a line for each product: choose the Order you just made, the Product, and type the Quantity. Type the Unit Price from the product\'s List Price, any Discount (0.05 is five percent), and the Tax Rate from the product\'s group in Categories.',
      'Run the query Orders to ship to see it waiting, with the days left before the customer needs it.',
    ] },
    { title: 'Ship the order and send its delivery note', steps: [
      'Pick the goods, then open the Orders form, find the order and type the Shipped Date. Change Status to Shipped.',
      'For each product sent, open the Stock Movements form and add a Sold movement with the quantity below nought (-12 for twelve sent) and Order and its number in Reference; then lower Units in Stock on the Products form by the same.',
      'Open Print templates, choose Delivery note, choose the order and press Print. It goes in the box for the customer to check and sign.',
    ] },
    { title: 'Invoice it', steps: [
      'Open the Invoices form and press New. Type the next number in Number (INV-2026-0301 after INV-2026-0300), choose the Order and the Customer, and type today in Invoice Date.',
      'Type the Due Date: the invoice date plus the customer\'s Payment Days from the Companies form.',
      'Type Net (the goods after discounts, with the delivery charge), Tax (the VAT on them) and Amount (the two together). The query Order totals shows the goods for the order.',
      'Open the Order Lines form and choose the new invoice in Invoice on each of the order\'s lines, so the invoice lists them.',
      'Open Print templates, choose Invoice, choose the new invoice and press Save as PDF to e-mail it, or Print.',
    ] },
    { title: 'Record a payment and see who owes what', steps: [
      'When a customer pays, open the Invoices form, find the invoice and type the day in Paid On. Then open the Orders form and change the order\'s Status to Closed.',
      'Run the query Unpaid invoices for every invoice not yet paid, by due date, with how many days each is overdue.',
      'Open the report Invoices due for the same list a page a customer, with what each owes in total.',
      'At the month\'s end, open Print templates, choose Statement and print it for each customer who owes money.',
    ] },
    { title: 'Reorder stock', steps: [
      'Run the query Products to reorder: products at or below their Reorder Level, by supplier, with how many to order.',
      'Open the Purchase Orders form and press New: type the next number (PO-2026-0120), choose the Supplier and yourself, today in Ordered On and the day it should come in Expected. Status is Submitted.',
      'Open the Purchase Lines form and add a line for each product to order, with the Quantity and the Unit Cost.',
      'Open Print templates, choose Purchase order and send it to the supplier as a PDF.',
      'When the goods come, type the day in Received On and change Status to Received; add a Received movement in Stock Movements for each line and raise Units in Stock by the same.',
    ] },
    { title: 'See how the business is doing', steps: [
      'Open the report Monthly sales for sales and units each month, and run Sales by category, Sales by employee and Top customers for who and what sells.',
      'Run Stock value for what the stock in the warehouse is worth at cost.',
      'Run Purchases on order for what is coming in and when.',
    ] },
  ],
  notes: [
    'Units in Stock on a product is the sum of its Stock Movements: received above nought, sold and corrected below. When you change one, change the other, or count the stock and add an Adjusted movement to put it right.',
    'An order line keeps its own Unit Price, Discount and Tax Rate, so a later price change never alters an order already taken or invoiced.',
    'Invoice and purchase order numbers are typed by you, a new run each year. Number them in date order and never use one twice; Data Files refuses a number already used.',
    'Customers, suppliers and carriers share the Companies table. Kind says which a company is, and most queries choose by it.',
    'VAT is a rate on each order line, taken from the product\'s group in Categories: food at 0, kitchen and cleaning supplies at 0.2. Delivery is charged at the standard rate. Check the rates for your own country and goods.',
  ],
  faq: [
    { q: 'How do I print one invoice?', a: 'Open Print templates, choose Invoice and choose the invoice. Press Print, or Save as PDF to send it by e-mail. From the Invoices form, the invoice on screen prints with the Print button.' },
    { q: 'Why does a product show below its reorder level when a delivery is due?', a: 'Products to reorder looks at Units in Stock only. Run Purchases on order to see what is already coming before ordering again.' },
    { q: 'How do I give one customer a discount on everything?', a: 'Type the discount on each of their order lines. To make it automatic, add a Discount field to Companies and copy it onto each line as you take the order.' },
    { q: 'Where do I change the name and bank details on invoices?', a: 'In the Business table: its one record is printed at the head and foot of every document.' },
    { q: 'Can a customer also be a supplier?', a: 'Yes. Enter the company once with the Kind it is most, and choose it as either; queries that choose by Kind will show it under that one.' },
    { q: 'How do I start again with my own records?', a: 'Make the sample again from the Data Files start screen and choose Empty: the same tables, queries, forms, reports and print templates, with nothing in them.' },
  ],
  empty: 'Start with Business, your own name, address and payment details, which every document prints. Then Categories and Companies (your suppliers and carriers first, then your customers), then Employees and Products. Orders, purchase orders, stock movements and invoices come in as you trade. If you already hold stock, enter it as one Received movement per product and set Units in Stock to match.',
  adapt: [
    'Sell in more than one currency: add a Currency field to Companies and a rate to Orders, and a calculated column for the amount in your own currency.',
    'Keep several warehouses: add a Warehouses table, point Stock Movements at it, and group the stock queries by warehouse.',
    'Record part payments: add a Payments table pointing at Invoices, and an Unpaid invoices query that sums what has been paid against each one.',
    'Give customers their own prices: add a Price Lists table (customer, product, price) and use it when taking an order.',
    'Track batches and best-before dates: add a Batch and a Best Before field to Purchase Lines and Stock Movements.',
    'Add sales tax: add a Tax Rate field to Products and a calculated column for the tax on each line.',
  ],
};

/** What is sold: [category, description, [name, unit, cost]...]. Generic goods; no brand. */
const RANGE = [
  ['Tea and Coffee', 'Loose and bagged teas, beans and ground coffee', [['Assam tea, loose', '1 kg bag', 9.2], ['Darjeeling tea, loose', '500 g tin', 11.5], ['Green tea, sencha', '500 g bag', 8.4], ['Mint tea bags', 'box of 100', 4.1], ['Chai spice blend', '500 g tin', 7.3], ['Arabica beans, medium roast', '1 kg bag', 13.8], ['Espresso beans, dark roast', '1 kg bag', 14.6], ['Ground coffee, filter', '500 g bag', 6.9]]],
  ['Spices', 'Whole and ground spices and blends', [['Cumin seed', '1 kg bag', 6.2], ['Ground turmeric', '1 kg bag', 5.4], ['Black pepper, whole', '500 g tub', 6.8], ['Cardamom, green', '250 g tub', 9.9], ['Saffron threads', '5 g tin', 12.5], ['Smoked paprika', '500 g tub', 5.1], ['Garam masala', '500 g tub', 5.9], ['Cinnamon sticks', '250 g bag', 4.4]]],
  ['Rice and Grains', 'Rice, couscous, lentils and flour', [['Basmati rice, aged', '10 kg sack', 16.5], ['Jasmine rice', '10 kg sack', 14.2], ['Red lentils', '5 kg sack', 7.8], ['Chickpeas, dried', '5 kg sack', 8.1], ['Couscous, medium', '5 kg sack', 7.2], ['Bulgur wheat', '5 kg sack', 6.6], ['Chapati flour', '10 kg sack', 9.4]]],
  ['Dried Fruit and Nuts', 'Dates, apricots, nuts and seeds', [['Medjool dates', '5 kg box', 24.5], ['Dried apricots', '2 kg bag', 9.6], ['Almonds, whole', '1 kg bag', 10.4], ['Pistachios, roasted', '1 kg bag', 17.9], ['Cashews, whole', '1 kg bag', 12.8], ['Walnut halves', '1 kg bag', 11.2], ['Sultanas', '2 kg bag', 6.3]]],
  ['Oils and Vinegars', 'Cooking oils, olive oils and vinegars', [['Extra virgin olive oil', '5 l tin', 29.5], ['Sunflower oil', '10 l drum', 17.4], ['Sesame oil, toasted', '500 ml bottle', 4.9], ['Balsamic vinegar', '500 ml bottle', 5.6], ['Cider vinegar', '1 l bottle', 2.8], ['Rapeseed oil, cold pressed', '1 l bottle', 4.2]]],
  ['Sauces and Preserves', 'Chutneys, pastes, jams and honey', [['Mango chutney', '1.2 kg jar', 6.1], ['Harissa paste', '1 kg tub', 7.4], ['Tahini', '1 kg jar', 6.7], ['Raw honey', '1 kg jar', 9.8], ['Apricot jam', '1.2 kg jar', 5.3], ['Tomato passata', 'case of 12', 10.6], ['Chilli sauce', '1 l bottle', 4.6]]],
  ['Baking', 'Sugar, flour, yeast and baking goods', [['Caster sugar', '5 kg bag', 5.2], ['Bread flour, strong', '16 kg sack', 13.1], ['Dried yeast', '500 g pack', 4.3], ['Cocoa powder', '1 kg bag', 8.9], ['Vanilla extract', '500 ml bottle', 14.2], ['Baking powder', '1 kg tub', 3.9]]],
  ['Household', 'Kitchen and cleaning supplies for the trade', [['Paper napkins', 'case of 2000', 11.4], ['Kraft takeaway boxes', 'case of 300', 21.6], ['Compostable cups, 12 oz', 'case of 1000', 34.5], ['Dish soap, concentrate', '5 l can', 7.6], ['Food wrap film', '300 m roll', 6.2], ['Baking parchment', '50 m roll', 4.8]]],
];
const KINDS_OF_CUSTOMER = ['Pantry', 'Deli', 'Market', 'Stores', 'Kitchen', 'Hospitality', 'Cafe', 'Hotels', 'Grocers'];
const KINDS_OF_SUPPLIER = ['Provisions', 'Imports', 'Merchants', 'Wholesale', 'Supplies', 'Foods'];

/** Fill a Northbank Trading database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 540 }) {
  const insert = (table, values) => base.insert(table, values);
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return insert(table, values); };

  add('Business', businessRecord({
    name: 'Northbank Trading', domain: 'northbank.example', phone: '020 7946 0100',
    address: { Address: '14 Quarry Road', City: 'London', Postcode: 'E14 9QR', Country: 'United Kingdom' },
    footer: 'Thank you for your order. Goods remain ours until paid for in full. Shortages or damage to be reported within 48 hours of delivery.',
  }));

  // Staff.
  const roles = ['Managing Director', 'Sales Manager', 'Sales Representative', 'Sales Representative', 'Sales Representative', 'Sales Representative', 'Purchasing Officer', 'Warehouse Lead', 'Accounts Clerk'];
  const staff = roles.map((role, i) => {
    const p = person(r, { domain: 'northbank.example' });
    return { id: add('Employees', { ...p, 'Job Title': role, Phone: `020 7946 0${String(100 + i).padStart(3, '0')}`, Hired: cal.day(-r.int(400, 3600)) }), role };
  });
  const reps = staff.filter((s) => /Sales/.test(s.role)).map((s) => s.id);
  const buyer = staff.find((s) => s.role === 'Purchasing Officer').id;

  // Companies: suppliers, carriers, customers.
  const supplierNames = companies(r, 12, KINDS_OF_SUPPLIER);
  const suppliers = supplierNames.map((name) => {
    const a = address(r);
    const id = add('Companies', { Name: name, Kind: 'Supplier', ...a, Phone: phoneFor(r, a.Country), 'E-mail': `orders@${name.toLowerCase().replace(/[^a-z]+/g, '')}.example`, 'Payment Days': r.pick([30, 45, 60]) });
    add('Contacts', { Company: id, ...person(r, { domain: `${name.toLowerCase().replace(/[^a-z]+/g, '')}.example` }), 'Job Title': r.pick(['Sales Manager', 'Account Manager', 'Export Manager']), Phone: phoneFor(r, a.Country) });
    return id;
  });
  const shippers = ['Swift Freight', 'Cargo Line Couriers', 'Overland Parcel'].map((name) => {
    const a = address(r, PLACES.slice(0, 1));
    return add('Companies', { Name: name, Kind: 'Shipper', ...a, Phone: phoneFor(r, a.Country), 'E-mail': `bookings@${name.toLowerCase().replace(/[^a-z]+/g, '')}.example` });
  });
  const customerNames = companies(r, 48, KINDS_OF_CUSTOMER).filter((n) => !supplierNames.includes(n));
  const customers = customerNames.map((name) => {
    const a = address(r);
    const size = r.weighted([['small', 5], ['medium', 3], ['large', 1]]);
    const terms = r.pick([14, 30, 30, 30, 45]);
    const id = add('Companies', {
      Name: name, Kind: 'Customer', ...a, Phone: phoneFor(r, a.Country), 'E-mail': `accounts@${name.toLowerCase().replace(/[^a-z]+/g, '')}.example`,
      'Payment Days': terms, 'Credit Limit': { small: 2000, medium: 6000, large: 15000 }[size],
    });
    for (let k = 0; k < r.int(1, 2); k++) add('Contacts', { Company: id, ...person(r, { domain: `${name.toLowerCase().replace(/[^a-z]+/g, '')}.example` }), 'Job Title': r.pick(['Owner', 'Buyer', 'Head Chef', 'Manager', 'Purchasing']), Phone: phoneFor(r, a.Country) });
    return { id, size, weight: { small: 1, medium: 2.5, large: 6 }[size], discount: size === 'large' ? 0.05 : 0, rep: r.pick(reps), days: terms };
  });

  // Categories and products.
  const products = [];
  RANGE.forEach(([cat, desc, items], ci) => {
    // Food carries no VAT; kitchen and cleaning supplies carry the standard rate.
    const rate = cat === 'Household' ? 0.2 : 0;
    const catId = add('Categories', { Name: cat, Description: desc, 'Tax Rate': rate });
    const supplierOf = suppliers[ci % suppliers.length];
    const second = suppliers[(ci + 5) % suppliers.length];
    items.forEach(([name, unit, cost], k) => {
      const list = money(cost * (1.35 + r.next() * 0.25));
      const level = r.int(8, 24);
      const p = { code: `${cat.slice(0, 2).toUpperCase()}${String(101 + k).padStart(3, '0')}`, name, unit, cost, list, rate, level, qty: level * 6, stock: 0, held: 0, supplier: k % 3 === 2 ? second : supplierOf, demand: 0.4 + r.next() * 1.6, discontinued: false };
      p.id = add('Products', { Code: p.code, Name: name, Category: catId, Supplier: p.supplier, Unit: unit, 'Cost Price': cost, 'List Price': list, 'Units in Stock': 0, 'Reorder Level': level, 'Reorder Quantity': p.qty });
      products.push(p);
    });
  });
  // Two lines no longer sold.
  for (const p of r.some(products, 2)) p.discontinued = true;

  const move = (p, when, kind, quantity, reference) => { p.stock += quantity; add('Stock Movements', { Product: p.id, When: when, Kind: kind, Quantity: quantity, Reference: reference }); };

  // Eighteen months from the first of a month, so each month's sales are a whole month's.
  let start = -days;
  while (cal.at(start).getDate() !== 1) start++;

  // Purchase orders numbered by the year they are placed in: PO-2026-0042.
  const poCount = {};
  const poNumber = (d) => { const y = cal.at(d).getFullYear(); poCount[y] = (poCount[y] || 0) + 1; return `PO-${y}-${String(poCount[y]).padStart(4, '0')}`; };

  // Opening stock, received before the first order: one purchase order to each supplier.
  for (const supplier of suppliers) {
    const mine = products.filter((p) => p.supplier === supplier);
    if (!mine.length) continue;
    const opening = add('Purchase Orders', { Number: poNumber(start - 20), Supplier: supplier, Employee: buyer, 'Ordered On': cal.day(start - 20), Expected: cal.day(start - 6), 'Received On': cal.day(start - 6), Status: 'Received' });
    for (const p of mine) {
      const n = p.qty * 2;
      add('Purchase Lines', { 'Purchase Order': opening, Product: p.id, Quantity: n, 'Unit Cost': p.cost });
      move(p, cal.day(start - 6), 'Received', n, `PO ${opening}`);
    }
  }

  const pending = []; // purchase orders not yet in: { id, at, lines }
  const invoices = [];
  for (let d = start; d <= 0; d++) {
    const date = cal.day(d);
    const wd = cal.weekday(d);
    if (wd === 0 || wd === 6) continue;
    // Deliveries arriving today.
    for (const po of pending.filter((x) => x.at === d)) {
      base.update('Purchase Orders', po.id, { 'Received On': date, Status: 'Received' });
      for (const l of po.lines) move(l.p, date, 'Received', l.n, `PO ${po.id}`);
    }
    for (let i = pending.length - 1; i >= 0; i--) if (pending[i].at === d) pending.splice(i, 1);

    // Orders: more of them in the weeks before the year ends.
    const month = cal.at(d).getMonth();
    const season = month === 10 || month === 11 ? 1.6 : month === 0 ? 0.7 : 1;
    const n = Math.round((0.8 + r.next() * 1.6) * season);
    for (let k = 0; k < n; k++) {
      const c = r.weighted(customers.map((x) => [x, x.weight]));
      // What is asked for, cut to what is on the shelf: nothing is sold that is not there.
      const lines = [];
      for (const p of r.some(products.filter((x) => !x.discontinued || d < -200), r.int(1, c.size === 'large' ? 7 : 4))) {
        const want = Math.max(1, Math.round(r.int(2, 12) * p.demand * (c.size === 'large' ? 2.5 : 1)));
        const qty = Math.min(want, p.stock - p.held);
        if (qty > 0) { lines.push({ p, qty }); p.held += qty; }
      }
      if (!lines.length) continue;
      const cancelled = r.chance(0.03);
      const shipIn = r.int(1, 4);
      const shipped = !cancelled && d + shipIn <= 0 ? cal.day(cal.workday(d + shipIn)) : null;
      const fee = c.size === 'large' ? 0 : r.pick([0, 8.5, 12, 15]);
      const order = add('Orders', {
        Customer: c.id, Employee: c.rep, 'Order Date': date, 'Required Date': cal.day(cal.workday(d + 7)), 'Shipped Date': shipped,
        Shipper: r.pick(shippers), 'Shipping Fee': fee, Status: cancelled ? 'Cancelled' : shipped ? 'Shipped' : 'New',
      });
      // Net and VAT as the invoice prints them: each line rounded, then added up; delivery at the standard rate.
      let net = fee;
      let vat = money(fee * 0.2);
      const lineIds = [];
      for (const { p, qty } of lines) {
        const discount = c.discount || (qty >= 20 ? 0.05 : 0);
        lineIds.push(add('Order Lines', { Order: order, Product: p.id, Quantity: qty, 'Unit Price': p.list, Discount: discount, 'Tax Rate': p.rate }));
        const lineNet = money(qty * p.list * (1 - discount));
        net += lineNet;
        vat += money(lineNet * p.rate);
        p.held -= qty;
        if (shipped) move(p, shipped, 'Sold', -qty, `Order ${order}`);
        // An order not shipped yet keeps its goods set aside.
        else if (!cancelled) p.held += qty;
      }
      if (shipped) invoices.push({ order, customer: c.id, lineIds, date: shipped, offset: cal.workday(d + shipIn), net: money(net), vat: money(vat), days: c.days, slow: c.size === 'small' && r.chance(0.25) });
    }

    // Each Monday and Thursday the buyer orders what is at or below its reorder level.
    if (wd === 1 || wd === 4) {
      const bySupplier = new Map();
      for (const p of products) {
        if (p.discontinued || p.stock - p.held > p.level || pending.some((po) => po.lines.some((l) => l.p === p))) continue;
        if (!bySupplier.has(p.supplier)) bySupplier.set(p.supplier, []);
        bySupplier.get(p.supplier).push({ p, n: p.qty + Math.max(0, p.level - p.stock) });
      }
      for (const [supplier, lines] of bySupplier) {
        const lead = r.int(5, 12);
        const id = add('Purchase Orders', { Number: poNumber(d), Supplier: supplier, Employee: buyer, 'Ordered On': date, Expected: cal.day(cal.workday(d + lead)), Status: 'Submitted' });
        for (const l of lines) add('Purchase Lines', { 'Purchase Order': id, Product: l.p.id, Quantity: l.n, 'Unit Cost': l.p.cost });
        pending.push({ id, at: cal.workday(d + lead), lines });
      }
    }
  }

  // Invoices for what shipped, numbered in date order a year at a time: most paid in time, some late, the newest not yet.
  invoices.sort((a, b) => a.offset - b.offset || a.order - b.order);
  const invCount = {};
  for (const inv of invoices) {
    const y = cal.at(inv.offset).getFullYear();
    invCount[y] = (invCount[y] || 0) + 1;
    const due = inv.offset + inv.days;
    const paidAt = inv.offset + (inv.slow ? inv.days + r.int(5, 40) : r.int(3, inv.days));
    const paid = paidAt <= 0 ? cal.day(paidAt) : null;
    const invoice = add('Invoices', { Number: `INV-${y}-${String(invCount[y]).padStart(4, '0')}`, Order: inv.order, Customer: inv.customer, 'Invoice Date': cal.day(inv.offset), 'Due Date': cal.day(due), Net: inv.net, Tax: inv.vat, Amount: money(inv.net + inv.vat), 'Paid On': paid });
    for (const id of inv.lineIds) base.update('Order Lines', id, { Invoice: invoice });
    if (paid) base.update('Orders', inv.order, { Status: 'Closed' });
  }

  // A stock count: a few corrections, as a real count finds.
  for (const p of r.some(products, 4)) move(p, cal.day(-r.int(20, 120)), 'Adjusted', -r.int(1, 3), 'Stock count');

  for (const p of products) base.update('Products', p.id, { 'Units in Stock': p.stock, Discontinued: p.discontinued });
  return counts;
}
