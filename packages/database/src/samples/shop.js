// The shop: Larkspur Corner Store, an invented neighbourhood shop. What it
// sells in categories, who supplies it, its till (each sale and its lines,
// paid by cash, card or a mobile wallet), the loyalty customers, deliveries
// in, stock, the staff on the till, and what the shop costs to run.
//
// Its records are a year of trading a day at a time: busier at
// lunchtime and on Saturdays, deliveries when stock runs low, rent and wages
// each month. Its queries are the ones a shopkeeper asks at closing time and
// at the month's end: today's takings, what sells, what earns, what to order.

import { companies, money, person, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const LINE = 'Round([Quantity] * [Unit Price] * (1 - Nz([Discount])), 2)';

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
      { expr: 'DateDiff("d", Format([When], "yyyy-mm-dd"), Date())', as: 'Days Ago', total: 'where', criteria: ['0'] },
      { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Daily takings': { tables: ['Sales', 'Sale Lines'], columns: [
      { expr: 'Format([When], "yyyy-mm-dd")', as: 'Day', total: 'group', sort: 'desc' }, { expr: LINE, as: 'Takings', total: 'sum' },
      { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Monthly takings': { tables: ['Sales', 'Sale Lines'], columns: [
      { expr: 'Format([When], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { expr: LINE, as: 'Takings', total: 'sum' },
      { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Best sellers': { tables: ['Sale Lines', 'Products', 'Sales'], columns: [
      { table: 'Products', field: 'Name', as: 'Product', total: 'group' }, { table: 'Sale Lines', field: 'Quantity', as: 'Sold', total: 'sum', sort: 'desc' },
      { expr: LINE, as: 'Takings', total: 'sum' }, { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Profit by category': { tables: ['Sale Lines', 'Products', 'Categories', 'Sales'], columns: [
      { table: 'Categories', field: 'Name', as: 'Category', total: 'group' }, { expr: LINE, as: 'Takings', total: 'sum' },
      { expr: 'Round([Quantity] * ([Unit Price] * (1 - Nz([Discount])) - [Cost Price]), 2)', as: 'Profit', total: 'sum', sort: 'desc' },
      { table: 'Sales', field: 'Refunded', total: 'where', criteria: ['No'] },
    ] },
    'Low stock': { tables: ['Products', 'Suppliers'], columns: [
      { table: 'Suppliers', field: 'Name', as: 'Supplier', sort: 'asc' }, { table: 'Products', field: 'Name', as: 'Product', sort: 'asc' },
      { table: 'Products', field: 'In Stock' }, { table: 'Products', field: 'Reorder Level' },
      { expr: '[In Stock] - [Reorder Level]', as: 'Spare', criteria: ['<= 0'], show: false }, { table: 'Products', field: 'Active', criteria: ['Yes'], show: false },
    ] },
    'Costs by month': { tables: ['Expenses'], columns: [
      { expr: 'Format([Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Expenses', field: 'Kind', total: 'group' },
      { table: 'Expenses', field: 'Amount', total: 'sum' },
    ] },
    'Best customers': { tables: ['Customers', 'Sales', 'Sale Lines'], columns: [
      { expr: '[First Name] & " " & [Last Name]', as: 'Customer', total: 'group' }, { expr: LINE, as: 'Spent', total: 'sum', sort: 'desc' },
      { table: 'Customers', field: 'Points', total: 'max' },
    ] },
    'Unpaid deliveries': { tables: ['Deliveries', 'Suppliers', 'Delivery Lines'], columns: [
      { table: 'Deliveries', field: 'ID', as: 'Delivery', total: 'group' }, { table: 'Suppliers', field: 'Name', as: 'Supplier', total: 'group' },
      { table: 'Deliveries', field: 'Received', total: 'group', sort: 'asc' }, { table: 'Deliveries', field: 'Invoice Number', total: 'group' },
      { expr: 'Round([Quantity] * [Unit Cost], 2)', as: 'Owed', total: 'sum' }, { table: 'Deliveries', field: 'Paid', total: 'where', criteria: ['No'] },
    ] },
  },
  reports: [
    { name: 'Takings by month', query: 'Monthly takings', columns: ['Month', { field: 'Takings', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }] },
    { name: 'Profit by category', query: 'Profit by category', columns: ['Category', { field: 'Takings', type: 'currency' }, { field: 'Profit', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }, { field: 'Profit', fn: 'sum' }] },
    { name: 'Shopping list', title: 'To order', query: 'Low stock', group: 'Supplier', columns: ['Product', { field: 'In Stock', type: 'integer' }, { field: 'Reorder Level', type: 'integer' }] },
    { name: 'Running costs', query: 'Costs by month', group: 'Month', columns: ['Kind', { field: 'Amount', type: 'currency' }], totals: [{ field: 'Amount', fn: 'sum' }] },
  ],
};

/** The shelves: [category, [name, cost, price, daily demand]...]. Generic goods. */
const SHELVES = [
  ['Bakery', [['White sliced loaf', 0.62, 1.25, 9], ['Wholemeal loaf', 0.75, 1.45, 5], ['Croissants, pack of 4', 1.1, 2.2, 3], ['Flatbreads, pack of 6', 0.8, 1.6, 4]]],
  ['Dairy and Eggs', [['Milk, 2 litres', 0.98, 1.65, 14], ['Free-range eggs, 6', 1.05, 1.9, 6], ['Butter, 250 g', 1.3, 2.25, 4], ['Natural yoghurt, 500 g', 0.7, 1.35, 4], ['Cheddar, 400 g', 2.1, 3.6, 3]]],
  ['Fruit and Veg', [['Bananas, bunch', 0.55, 1.1, 8], ['Apples, bag of 6', 0.9, 1.75, 5], ['Onions, 1 kg', 0.45, 0.95, 5], ['Potatoes, 2.5 kg', 0.9, 1.85, 4], ['Tomatoes, 6', 0.7, 1.4, 4]]],
  ['Drinks', [['Bottled water, 500 ml', 0.18, 0.75, 16], ['Orange juice, 1 litre', 0.85, 1.6, 5], ['Cola, 330 ml can', 0.32, 0.95, 12], ['Instant coffee, 100 g', 1.9, 3.5, 2], ['Tea bags, 80', 1.2, 2.3, 3]]],
  ['Snacks', [['Crisps, single bag', 0.25, 0.85, 14], ['Chocolate bar', 0.38, 0.99, 15], ['Biscuits, packet', 0.55, 1.2, 6], ['Mixed nuts, 200 g', 1.1, 2.25, 2]]],
  ['Groceries', [['Pasta, 500 g', 0.42, 0.95, 5], ['Basmati rice, 1 kg', 1.15, 2.2, 3], ['Chopped tomatoes, tin', 0.32, 0.7, 6], ['Baked beans, tin', 0.36, 0.8, 6], ['Cooking oil, 1 litre', 1.05, 1.95, 2]]],
  ['Household', [['Washing-up liquid', 0.6, 1.3, 2], ['Toilet roll, 9 pack', 2.4, 4.5, 3], ['Bin bags, roll of 20', 0.9, 1.9, 2], ['Laundry liquid, 1 litre', 2.1, 4.0, 1]]],
  ['Newsagent', [['Daily newspaper', 0.9, 1.2, 10], ['Greeting card', 0.7, 2.5, 2], ['Phone top-up voucher', 9.5, 10, 2], ['Postage stamps, book of 4', 2.6, 2.6, 1]]],
];

export function fill(base, { r, cal, days = 365 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };

  const staff = ['Owner', 'Shop Assistant', 'Shop Assistant', 'Weekend Assistant'].map((role, i) => {
    const p = person(r);
    return add('Staff', { 'First Name': p['First Name'], 'Last Name': p['Last Name'], Role: role, Phone: `07700 900${String(400 + i).padStart(3, '0')}`, Started: cal.day(-r.int(600, 2400)) });
  });

  const supplierNames = companies(r, 6, ['Wholesale', 'Supplies', 'Provisions', 'Foods']);
  const deliveryDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  const suppliers = supplierNames.map((name, i) => add('Suppliers', {
    Name: name, Contact: (({ 'First Name': f, 'Last Name': l }) => `${f} ${l}`)(person(r)),
    Phone: phoneFor(r, 'United Kingdom'), 'E-mail': `orders@${name.toLowerCase().replace(/[^a-z]+/g, '')}.example`, 'Delivers On': deliveryDays[i % 5],
  }));

  const products = [];
  let code = 5012345000000;
  SHELVES.forEach(([cat, items], ci) => {
    const catId = add('Categories', { Name: cat });
    for (const [name, cost, price, demand] of items) {
      const level = Math.max(4, Math.round(demand * 2));
      const p = { name, cost, price, demand, level, stock: level * 3, supplier: suppliers[ci % suppliers.length] };
      p.id = add('Products', { Barcode: String(code += r.int(7, 97)), Name: name, Category: catId, Supplier: p.supplier, 'Cost Price': cost, 'Sale Price': price, 'In Stock': p.stock, 'Reorder Level': level });
      products.push(p);
    }
  });

  const customers = [];
  for (let i = 0; i < 40; i++) {
    const p = person(r);
    customers.push({ id: add('Customers', { ...p, Phone: `07700 900${String(i * 11 + 3).padStart(3, '0')}`, Joined: cal.day(-r.int(30, days)) }), points: 0 });
  }

  for (let d = -days; d <= 0; d++) {
    const wd = cal.weekday(d);
    // Deliveries: each supplier on its day, for what has run low.
    for (const [si, supplier] of suppliers.entries()) {
      if (wd !== (si % 5) + 1) continue;
      const low = products.filter((p) => p.supplier === supplier && p.stock <= p.level * 1.5);
      if (!low.length) continue;
      const delivery = add('Deliveries', { Supplier: supplier, Received: cal.day(d), 'Invoice Number': `INV-${r.int(10000, 99999)}`, Paid: d < -14 || r.chance(0.3) });
      for (const p of low) {
        const n = p.level * 3;
        add('Delivery Lines', { Delivery: delivery, Product: p.id, Quantity: n, 'Unit Cost': p.cost });
        p.stock += n;
      }
    }
    // The till: open 7 to 21, busier at lunch and on Saturdays, quieter on Sundays.
    const dayFactor = wd === 6 ? 1.35 : wd === 0 ? 0.7 : 1;
    const hours = d === 0 ? Math.min(21, new Date().getHours() || 21) : 21;
    for (let h = 7; h < hours; h++) {
      const busy = (h >= 12 && h <= 13 ? 2 : h >= 16 && h <= 18 ? 1.6 : 1) * dayFactor;
      const n = Math.round(r.next() * 1.6 * busy);
      for (let k = 0; k < n; k++) {
        const items = [];
        for (const p of r.some(products, r.int(1, 4))) {
          if (p.stock <= 0 || !r.chance(Math.min(0.95, p.demand / 10 + 0.15))) continue;
          const q = Math.min(p.stock, r.chance(0.8) ? 1 : r.int(2, 3));
          items.push({ p, q });
        }
        if (!items.length) continue;
        const loyal = r.chance(0.3) ? r.pick(customers) : null;
        const sale = add('Sales', {
          When: cal.time(d, h, r.int(0, 59)), Staff: wd === 0 || wd === 6 ? staff[3] : r.pick(staff.slice(0, 3)), Customer: loyal?.id ?? null,
          Payment: r.weighted([['Card', 6], ['Cash', 3], ['Mobile', 2]]), Refunded: r.chance(0.004),
        });
        let spent = 0;
        for (const { p, q } of items) {
          // Two for one on snacks, now and then.
          const discount = r.chance(0.04) ? 0.1 : 0;
          add('Sale Lines', { Sale: sale, Product: p.id, Quantity: q, 'Unit Price': p.price, Discount: discount });
          p.stock -= q;
          spent += q * p.price;
        }
        if (loyal) loyal.points += Math.floor(spent);
      }
    }
    // Month's end: rent, wages, power, and now and then a repair.
    if (cal.at(d + 1).getDate() === 1) {
      const date = cal.day(d);
      add('Expenses', { Date: date, Kind: 'Rent', Description: 'Shop rent', Amount: 1450 });
      add('Expenses', { Date: date, Kind: 'Wages', Description: 'Staff wages', Amount: money(3900 + r.next() * 400) });
      add('Expenses', { Date: date, Kind: 'Utilities', Description: 'Electricity and water', Amount: money(310 + r.next() * 140) });
      add('Expenses', { Date: date, Kind: 'Card fees', Description: 'Card terminal fees', Amount: money(85 + r.next() * 30) });
      if (r.chance(0.25)) add('Expenses', { Date: date, Kind: 'Repairs', Description: r.pick(['Fridge service', 'Door lock', 'Shelving', 'Lighting']), Amount: money(60 + r.next() * 240) });
    }
  }

  for (const p of products) base.update('Products', p.id, { 'In Stock': p.stock });
  for (const c of customers) base.update('Customers', c.id, { Points: c.points });
  return counts;
}
