// The shop: Larkspur Corner Store, an invented neighbourhood shop. What it
// sells in categories, who supplies it, its till (each sale and its lines,
// paid by cash, card or a mobile wallet), the loyalty customers, deliveries
// in, stock, the staff on the till, and what the shop costs to run.
//
// Its records are a year of trading a day at a time: busier at
// lunchtime and on Saturdays, deliveries on each supplier's days for what
// has run low, rent and wages each month. Its queries are the ones a
// shopkeeper asks at closing time and at the month's end: today's takings,
// what sells, what earns, what to order.

import { companies, money, person, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const LINE = 'Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)';
const DAYS_AGO = 'DateDiff("d", Format([When], "yyyy-mm-dd"), Date())';

export const spec = {
  tables: [
    { name: 'Categories', fields: [{ name: 'Name', type: 'text', required: true, unique: true }] },
    { name: 'Suppliers', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Contact', type: 'text' }, { name: 'Phone', type: 'text' },
      { name: 'E-mail', type: 'text' }, { name: 'Delivers On', type: 'text' },
    ] },
    { name: 'Products', fields: [
      { name: 'Barcode', type: 'text', unique: true }, { name: 'Name', type: 'text', required: true }, { name: 'Category', ...ref('Categories') },
      { name: 'Supplier', ...ref('Suppliers') }, { name: 'Cost Price', type: 'currency' }, { name: 'Sale Price', type: 'currency' },
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
      { name: 'When', type: 'datetime' }, { name: 'Staff', ...ref('Staff') }, { name: 'Customer', ...ref('Customers') },
      { name: 'Payment', type: 'text', default: 'Card' }, { name: 'Refunded', type: 'yesno' },
    ] },
    { name: 'Sale Lines', fields: [
      { name: 'Sale', ...ref('Sales'), required: true }, { name: 'Product', ...ref('Products'), required: true },
      { name: 'Quantity', type: 'integer', default: 1 }, { name: 'Unit Price', type: 'currency' }, { name: 'Discount', type: 'number', default: 0 },
    ] },
    { name: 'Deliveries', fields: [
      { name: 'Supplier', ...ref('Suppliers'), required: true }, { name: 'Received', type: 'date' }, { name: 'Invoice Number', type: 'text' }, { name: 'Paid', type: 'yesno' },
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
    'Profit by category': { tables: ['Sale Lines', 'Products', 'Categories', 'Sales'], columns: [
      { table: 'Categories', field: 'Name', as: 'Category', total: 'group' }, { expr: LINE, as: 'Takings', total: 'sum' },
      { expr: 'Round([Quantity] * ([Unit Price] * (1 - Nz([Discount])) - [Cost Price]), 2)', as: 'Profit', total: 'sum', sort: 'desc' },
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
  },
  reports: [
    { name: 'Takings by month', query: 'Monthly takings', columns: ['Month', { field: 'Takings', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }] },
    { name: 'Last 30 days', title: 'Takings, the last 30 days', query: 'Daily takings', columns: [{ field: 'Day', type: 'date' }, { field: 'Items', type: 'integer' }, { field: 'Takings', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }, { field: 'Takings', fn: 'avg' }] },
    { name: 'Profit by category', query: 'Profit by category', columns: ['Category', { field: 'Takings', type: 'currency' }, { field: 'Profit', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }, { field: 'Profit', fn: 'sum' }] },
    { name: 'Shopping list', title: 'To order', query: 'Low stock', group: 'Supplier', columns: ['Product', { field: 'In Stock', type: 'integer' }, { field: 'Reorder Level', type: 'integer' }] },
    { name: 'Stock take sheet', table: 'Products', group: 'Category', sort: 'Name', columns: ['Name', 'Barcode', 'In Stock'], totals: [{ field: 'In Stock', fn: 'sum' }] },
    { name: 'Running costs', query: 'Costs by month', group: 'Month', columns: ['Kind', { field: 'Amount', type: 'currency' }], totals: [{ field: 'Amount', fn: 'sum' }] },
    { name: 'Deliveries to pay', query: 'Unpaid deliveries', group: 'Supplier', columns: ['Delivery', { field: 'Received', type: 'date' }, 'Invoice Number', { field: 'Owed', type: 'currency' }], totals: [{ field: 'Owed', fn: 'sum' }] },
  ],
};

export const about = {
  name: 'Larkspur Corner Store',
  kind: 'Shop',
  line: 'Products, the till, stock, deliveries, loyalty customers and running costs for a neighbourhood shop.',
  summary: [
    'Larkspur Corner Store is an invented neighbourhood shop open from seven in the morning to nine at night, every day of the week. It sells bread, milk, fruit and vegetables, drinks, snacks, groceries, household goods and newspapers, from about forty lines bought from six suppliers.',
    'The sample holds a year of trading made up a day at a time. The till rings up a sale every few minutes, more at lunchtime and after school and more on Saturdays. Each supplier delivers on its own days for what has run low, so stock goes down between deliveries and back up again. At each month end the rent, wages, power and card fees are paid, so the takings can be set against what the shop costs to run.',
    'Use it to see how a till, stock and suppliers fit together in one database, or start from it empty for a shop of your own.',
  ],
  suits: ['A convenience store or minimarket', 'A farm shop', 'A newsagent', 'A gift shop', 'A pharmacy counter for the goods it sells off the shelf', 'A market stall'],
  tables: {
    Categories: { about: 'The sections of the shop, used to group products on reports.', fields: { Name: 'The section, such as Bakery or Drinks. Each name is used once.' } },
    Suppliers: { about: 'The wholesalers and producers the shop buys from.', fields: {
      Name: 'The supplier\'s name.', Contact: 'The person you deal with there.', Phone: 'Their phone number.', 'E-mail': 'Where orders and queries go.',
      'Delivers On': 'The days the supplier comes, so you know when an order will arrive.',
    } },
    Products: { about: 'Everything on the shelves, one record per line you sell.', fields: {
      Barcode: 'The number on the pack the till scans. Each barcode is used once.', Name: 'The product as it shows on the till and the receipt.',
      Category: 'The section it belongs to.', Supplier: 'Who you buy it from.', 'Cost Price': 'What one costs you to buy.',
      'Sale Price': 'What one sells for.', 'In Stock': 'How many are on the shelf and in the back now.',
      'Reorder Level': 'When stock falls to this number or below, it goes on the shopping list.', Active: 'Untick a product you no longer sell; it stays in the old sales.',
    } },
    Staff: { about: 'The people who work the till.', fields: {
      'First Name': 'Their first name.', 'Last Name': 'Their family name.', Role: 'Owner, shop assistant or weekend assistant.',
      Phone: 'A number to reach them.', Started: 'The day they started.',
    } },
    Customers: { about: 'Regulars who signed up for the loyalty card.', fields: {
      'First Name': 'Their first name.', 'Last Name': 'Their family name.', Phone: 'Their mobile number.', 'E-mail': 'For offers, if they asked for them.',
      Joined: 'The day they signed up.', Points: 'Loyalty points earned: one for each whole amount spent.',
    } },
    Sales: { about: 'Each time the till is rung up: one record per receipt.', fields: {
      When: 'The date and time of the sale.', Staff: 'Who was on the till.', Customer: 'The loyalty customer, when they showed their card.',
      Payment: 'Card, Cash or Mobile.', Refunded: 'Ticked when the whole sale was refunded; refunded sales are left out of the takings.',
    } },
    'Sale Lines': { about: 'The items on each receipt.', fields: {
      Sale: 'The receipt the line belongs to.', Product: 'What was sold.', Quantity: 'How many.',
      'Unit Price': 'The price of one at the time, kept so a later price change does not alter old sales.', Discount: 'Any discount, as a fraction: 0.1 is ten percent off.',
    } },
    Deliveries: { about: 'Each delivery from a supplier, with its invoice.', fields: {
      Supplier: 'Who delivered.', Received: 'The day it came.', 'Invoice Number': 'The number on the supplier\'s invoice.', Paid: 'Ticked once the invoice is paid.',
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
    'Profit by category': 'For each section of the shop, what it took and what it earned after the cost of the goods, the best earner first.',
    'Takings by staff': 'What each person rang up over the last 30 days.',
    'Low stock': 'Products at or below their reorder level, by supplier, with the days each supplier delivers. This is the shopping list.',
    'Stock value': 'What the stock on the shelves is worth by section, at cost and at sale price.',
    'Costs by month': 'What the shop paid out each month, by kind of cost.',
    'Best customers': 'Loyalty customers by what they have spent, with their points.',
    'Unpaid deliveries': 'Supplier invoices not yet paid, oldest first, with what each one comes to.',
  },
  reports: {
    'Takings by month': 'One line a month with the takings and a total for the year.',
    'Last 30 days': 'Takings and items for each of the last 30 days, with the total and the average day.',
    'Profit by category': 'Takings and profit by section, with totals: the page to look at when deciding what to stock more of.',
    'Shopping list': 'What to order, on a page for each supplier, with what is left and the reorder level.',
    'Stock take sheet': 'Every product by section with its barcode and what the database says is in stock, to count against.',
    'Running costs': 'Each month\'s costs by kind, with a total for the month.',
    'Deliveries to pay': 'Supplier invoices still to pay, by supplier, with a total for each.',
  },
  empty: 'Start with Categories and Suppliers, then add your Products with their prices and reorder levels, then your Staff. Sales and Deliveries come in as you trade. Put your running costs in Expenses as you pay them.',
  adapt: [
    'Sell by weight: add a Unit field to Products (each, kg, 100 g) and make Quantity in Sale Lines a Number rather than a whole number.',
    'Track sell-by dates: add a Best Before field to Delivery Lines and a query for lines within three days of it.',
    'Add VAT or sales tax: add a Tax Rate field to Categories and a calculated column in the takings queries for the tax in each line.',
    'Keep more than one till or shop: add a Tills table, point Sales at it, and group the takings queries by till.',
    'Reward points differently: change Points to count one for every two spent, or add a Points Redeemed field to Sales.',
    'Order from the shopping list: copy the Low stock query into Worksheets and send it to the supplier, or add an Orders table for what you asked for and when.',
  ],
};

/** The shelves: [category, [name, cost, price, units a day]...]. Generic goods; no brand. */
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

export function fill(base, { r, cal, days = 365 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };

  const staff = ['Owner', 'Shop Assistant', 'Shop Assistant', 'Weekend Assistant'].map((role, i) => {
    const p = person(r);
    return add('Staff', { 'First Name': p['First Name'], 'Last Name': p['Last Name'], Role: role, Phone: `07700 900${String(400 + i).padStart(3, '0')}`, Started: cal.day(-r.int(600, 2400)) });
  });

  const taken = new Set(['Larkspur']);
  const suppliers = ROUNDS.map(([sections, on, tail]) => {
    let name;
    do name = companies(r, 1, [tail])[0]; while (taken.has(name.split(' ')[0]));
    taken.add(name.split(' ')[0]);
    const id = add('Suppliers', {
      Name: name, Contact: (({ 'First Name': f, 'Last Name': l }) => `${f} ${l}`)(person(r)),
      Phone: phoneFor(r, 'United Kingdom'), 'E-mail': `orders@${name.toLowerCase().replace(/[^a-z]+/g, '')}.example`,
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
      const p = { name, cost, price, demand, par, level, stock: par, supplier: s };
      p.id = add('Products', { Barcode: String(code += r.int(7, 97)), Name: name, Category: catId, Supplier: s.id, 'Cost Price': cost, 'Sale Price': price, 'In Stock': p.stock, 'Reorder Level': level });
      products.push(p);
    }
  }

  const customers = [];
  for (let i = 0; i < 60; i++) {
    const p = person(r);
    customers.push({ id: add('Customers', { ...p, Phone: `07700 900${String(i * 11 + 3).padStart(3, '0')}`, Joined: cal.day(-r.int(30, days)) }), points: 0 });
  }

  // A year from the first of a month, so each month in the costs is a whole one.
  let start = -days;
  while (cal.at(start).getDate() !== 1) start++;
  let cardTakings = 0;
  // Two lines the suppliers have been short of for the last fortnight: there is always something to chase.
  const short = r.some(products.filter((p) => p.demand >= 3 && p.supplier.sections[0] !== 'Newsagent'), 2);

  for (let d = start; d <= 0; d++) {
    const wd = cal.weekday(d);
    // Deliveries: each supplier on its days, topping up what has run low.
    for (const s of suppliers) {
      if (!s.on.includes(wd)) continue;
      const low = products.filter((p) => p.supplier === s && p.stock < p.par * 0.8 && !(d > -14 && short.includes(p)));
      if (!low.length) continue;
      const delivery = add('Deliveries', { Supplier: s.id, Received: cal.day(d), 'Invoice Number': `INV-${r.int(10000, 99999)}`, Paid: d < -14 || r.chance(0.3) });
      for (const p of low) {
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
      for (let k = 0; k < n; k++) {
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
        const sale = add('Sales', {
          When: cal.time(d, h, r.int(0, 59)), Staff: wd === 0 || wd === 6 ? r.pick([staff[0], staff[3]]) : r.pick(staff.slice(0, 3)), Customer: loyal?.id ?? null,
          Payment: payment, Refunded: refunded,
        });
        let spent = 0;
        for (const { p, q } of items) {
          // Ten percent off now and then: a short date or a multibuy.
          const discount = r.chance(0.04) ? 0.1 : 0;
          add('Sale Lines', { Sale: sale, Product: p.id, Quantity: q, 'Unit Price': p.price, Discount: discount });
          if (!refunded) p.stock -= q;
          spent += money(q * p.price * (1 - discount));
        }
        if (loyal && !refunded) loyal.points += Math.floor(spent);
        if (payment !== 'Cash' && !refunded) cardTakings += spent;
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

  for (const p of products) base.update('Products', p.id, { 'In Stock': p.stock });
  for (const c of customers) base.update('Customers', c.id, { Points: c.points });
  return counts;
}
