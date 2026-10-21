// The repair shop: Ironbridge Repairs, an invented shop that mends phones,
// tablets, laptops and small appliances, with three technicians at the bench
// and a counter where work is booked in and collected. Customers and the
// devices they bring, the staff, every job from booking in to collection,
// the parts on the shelf and those ordered in, the parts and labour each job
// used, and what was paid.
//
// Its records come from fifteen months of work, run a day at a time: jobs
// are booked in, diagnosed, quoted and approved or turned down, wait for a
// part when the shelf has none, are repaired and then collected (most within
// days, a few left for weeks). Phones come in more in the summer holidays,
// laptops when term starts, kitchen appliances before the year ends. A few
// repairs come back under warranty. The questions a repair shop asks at the
// counter and at the bench are its saved queries; its reports print them.

import { companies, money, person, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const CUSTOMER = '[Customers].[First Name] & " " & [Customers].[Last Name]';
const TECHNICIAN = '[Staff].[First Name] & " " & [Staff].[Last Name]';
const DEVICE = '[Make] & " " & [Model]';
// The job board's order, as the shop works through it, rather than A to Z.
const STAGE = 'IIf([Status] = "Booked in", "1 Booked in", IIf([Status] = "Diagnosing", "2 Diagnosing", IIf([Status] = "Awaiting approval", "3 Awaiting approval", '
  + 'IIf([Status] = "Awaiting parts", "4 Awaiting parts", IIf([Status] = "Repairing", "5 Repairing", IIf([Status] = "Ready", "6 Ready", "7 " & [Status]))))))';

export const spec = {
  // Where a value below nought is right: a job already past its promised
  // date, and a part that was fitted free under warranty more than it was sold.
  negative: { 'Job board': ['Days Left'], 'Parts margin': ['Margin'] },
  tables: [
    { name: 'Customers', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Company', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Address', type: 'text' }, { name: 'City', type: 'text' },
      { name: 'Postcode', type: 'text' }, { name: 'Contact By', type: 'text', default: 'Text message' }, { name: 'On Account', type: 'yesno' },
      { name: 'Joined', type: 'date' },
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
    { name: 'Jobs', fields: [
      { name: 'Device', ...ref('Devices'), required: true }, { name: 'Technician', ...ref('Staff') }, { name: 'Booked In', type: 'date' },
      { name: 'Fault', type: 'text' }, { name: 'Left With It', type: 'text' }, { name: 'Promised', type: 'date' },
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
    { name: 'Job Parts', fields: [
      { name: 'Job', ...ref('Jobs'), required: true }, { name: 'Part', ...ref('Parts'), required: true },
      { name: 'Quantity', type: 'integer', default: 1 }, { name: 'Unit Price', type: 'currency' }, { name: 'Fitted On', type: 'date' },
    ] },
    { name: 'Labour', fields: [
      { name: 'Job', ...ref('Jobs'), required: true }, { name: 'Done On', type: 'date' }, { name: 'Description', type: 'text' },
      { name: 'Hours', type: 'number' }, { name: 'Rate', type: 'currency' },
    ] },
    { name: 'Payments', fields: [
      { name: 'Job', ...ref('Jobs'), required: true }, { name: 'Paid On', type: 'date' }, { name: 'Amount', type: 'currency' },
      { name: 'Method', type: 'text', default: 'Card' }, { name: 'Kind', type: 'text', default: 'On collection' },
    ] },
  ],
  queries: {
    'Job board': { tables: ['Jobs', 'Devices', 'Customers', 'Staff'], columns: [
      { expr: STAGE, as: 'Stage', sort: 'asc' }, { table: 'Jobs', field: 'ID', as: 'Job' }, { table: 'Jobs', field: 'Booked In' },
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
      { table: 'Jobs', field: 'ID', as: 'Job' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' }, { expr: DEVICE, as: 'Device' },
      { table: 'Parts', field: 'Name', as: 'Part' }, { table: 'Suppliers', field: 'Name', as: 'Supplier' }, { table: 'Part Orders', field: 'Ordered On' },
      { table: 'Part Orders', field: 'Expected', sort: 'asc' }, { table: 'Jobs', field: 'Promised' },
      { table: 'Part Orders', field: 'Received On', criteria: ['Is Null'], show: false }, { table: 'Jobs', field: 'Status', criteria: ['Awaiting parts'], show: false },
    ] },
    'Past promised date': { tables: ['Jobs', 'Devices', 'Customers', 'Staff'], columns: [
      { table: 'Jobs', field: 'ID', as: 'Job' }, { table: 'Jobs', field: 'Status' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' },
      { expr: DEVICE, as: 'Device' }, { expr: TECHNICIAN, as: 'Technician' }, { table: 'Jobs', field: 'Promised' },
      { expr: 'DateDiff("d", [Promised], Date())', as: 'Days Late', sort: 'desc', criteria: ['> 0'] },
      { table: 'Jobs', field: 'Finished', criteria: ['Is Null'], show: false },
    ] },
    'Ready for collection': { tables: ['Jobs', 'Devices', 'Customers'], columns: [
      { table: 'Jobs', field: 'ID', as: 'Job' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' },
      { table: 'Customers', field: 'Contact By' }, { expr: DEVICE, as: 'Device' }, { table: 'Jobs', field: 'Outcome' }, { table: 'Jobs', field: 'Finished' },
      { expr: 'DateDiff("d", [Finished], Date())', as: 'Days Waiting', sort: 'desc' }, { expr: 'Round(Nz([Charge]) - Nz([Deposit]), 2)', as: 'To Pay' },
      { table: 'Jobs', field: 'Finished', criteria: ['Is Not Null'], show: false }, { table: 'Jobs', field: 'Collected', criteria: ['Is Null'], show: false },
    ] },
    'Uncollected 14 days or more': { tables: ['Jobs', 'Devices', 'Customers'], columns: [
      { table: 'Jobs', field: 'ID', as: 'Job' }, { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' }, { table: 'Customers', field: 'E-mail' },
      { expr: DEVICE, as: 'Device' }, { table: 'Jobs', field: 'Finished' },
      { expr: 'DateDiff("d", [Finished], Date())', as: 'Days Waiting', sort: 'desc', criteria: ['>= 14'] }, { expr: 'Round(Nz([Charge]) - Nz([Deposit]), 2)', as: 'To Pay' },
      { table: 'Jobs', field: 'Notes' }, { table: 'Jobs', field: 'Collected', criteria: ['Is Null'], show: false },
    ] },
    'Quotes awaiting approval': { tables: ['Jobs', 'Devices', 'Customers'], columns: [
      { table: 'Jobs', field: 'ID', as: 'Job' }, { table: 'Jobs', field: 'Quoted On', sort: 'asc' },
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
      { table: 'Jobs', field: 'ID', as: 'Job' }, { table: 'Jobs', field: 'Booked In', sort: 'desc' }, { table: 'Jobs', field: 'Warranty Of', as: 'First Job' },
      { expr: CUSTOMER, as: 'Customer' }, { expr: DEVICE, as: 'Device' }, { table: 'Jobs', field: 'Fault' }, { expr: TECHNICIAN, as: 'Technician' },
      { table: 'Jobs', field: 'Status' }, { table: 'Jobs', field: 'Warranty Of', criteria: ['Is Not Null'], show: false },
    ] },
    'Job sheets': { tables: ['Jobs', 'Devices', 'Customers', 'Staff'], columns: [
      { expr: TECHNICIAN, as: 'Technician', sort: 'asc' }, { table: 'Jobs', field: 'ID', as: 'Job', sort: 'asc' }, { table: 'Jobs', field: 'Booked In' },
      { expr: CUSTOMER, as: 'Customer' }, { table: 'Customers', field: 'Phone' }, { expr: DEVICE, as: 'Device' }, { table: 'Devices', field: 'Serial' },
      { table: 'Jobs', field: 'Fault' }, { table: 'Jobs', field: 'Left With It' }, { table: 'Jobs', field: 'Status' }, { table: 'Jobs', field: 'Quote' },
      { table: 'Jobs', field: 'Promised' }, { expr: 'DateDiff("d", [Booked In], Date())', as: 'Days Ago', criteria: ['Between 0 And 14'], show: false },
    ] },
    'Parts margin': { tables: ['Job Parts', 'Parts'], columns: [
      { table: 'Parts', field: 'Fits', total: 'group', sort: 'asc' }, { table: 'Parts', field: 'Name', as: 'Part', total: 'group', sort: 'asc' },
      { table: 'Job Parts', field: 'Quantity', as: 'Fitted', total: 'sum' }, { expr: 'Round([Quantity] * [Unit Price], 2)', as: 'Charged', total: 'sum' },
      { expr: 'Round([Quantity] * ([Unit Price] - [Cost]), 2)', as: 'Margin', total: 'sum' },
      { table: 'Job Parts', field: 'Fitted On', total: 'where', criteria: ['Is Not Null'] },
    ] },
    'Accounts to chase': { tables: ['Jobs', 'Devices', 'Customers'], columns: [
      { table: 'Customers', field: 'Company', sort: 'asc' }, { expr: CUSTOMER, as: 'Contact' }, { table: 'Customers', field: 'E-mail' },
      { table: 'Jobs', field: 'ID', as: 'Job' }, { table: 'Jobs', field: 'Collected', sort: 'asc' }, { table: 'Jobs', field: 'Charge' },
      { expr: 'DateDiff("d", [Collected], Date())', as: 'Days Since Collected' },
      { table: 'Jobs', field: 'Collected', criteria: ['Is Not Null'], show: false }, { table: 'Jobs', field: 'Paid', criteria: ['No'], show: false },
      { table: 'Jobs', field: 'Charge', criteria: ['> 0'], show: false },
    ] },
  },
  reports: [
    { name: 'Job board', query: 'Job board', group: 'Stage', sort: 'Promised', columns: ['Job', { field: 'Booked In', type: 'date' }, 'Customer', 'Device', 'Fault', 'Technician', { field: 'Promised', type: 'date' }, { field: 'Days Left', type: 'integer' }], totals: [{ field: 'Job', fn: 'count' }] },
    { name: 'Ready for collection', title: 'Ready for collection: to phone', query: 'Ready for collection', sort: 'Days Waiting', desc: true, columns: ['Job', 'Customer', 'Phone', 'Contact By', 'Device', 'Outcome', { field: 'Finished', type: 'date' }, { field: 'Days Waiting', type: 'integer' }, { field: 'To Pay', type: 'currency' }], totals: [{ field: 'Job', fn: 'count' }, { field: 'To Pay', fn: 'sum' }] },
    { name: 'Parts to order', query: 'Parts to reorder', group: 'Supplier', columns: ['Code', 'Part', { field: 'In Stock', type: 'integer' }, { field: 'Reorder Level', type: 'integer' }, { field: 'Reorder Quantity', type: 'integer' }, { field: 'Cost', type: 'currency' }, { field: 'Order Value', type: 'currency' }], totals: [{ field: 'Order Value', fn: 'sum' }] },
    { name: 'Takings by month', query: 'Takings by month', columns: ['Month', { field: 'Payments', type: 'integer' }, { field: 'Takings', type: 'currency' }], totals: [{ field: 'Payments', fn: 'sum' }, { field: 'Takings', fn: 'sum' }] },
    { name: 'Job sheets', title: 'Job sheets: the last two weeks', query: 'Job sheets', group: 'Technician', sort: 'Job', columns: ['Job', { field: 'Booked In', type: 'date' }, 'Customer', 'Phone', 'Device', 'Serial', 'Fault', 'Left With It', 'Status', { field: 'Quote', type: 'currency' }, { field: 'Promised', type: 'date' }], totals: [{ field: 'Job', fn: 'count' }] },
    { name: 'Technician turnaround', query: 'Technician turnaround', columns: ['Technician', { field: 'Jobs Repaired', type: 'integer' }, { field: 'Average Days', type: 'number' }, { field: 'Charged', type: 'currency' }], totals: [{ field: 'Jobs Repaired', fn: 'sum' }, { field: 'Charged', fn: 'sum' }] },
  ],
};

export const about = {
  name: 'Ironbridge Repairs',
  kind: 'Repair shop',
  line: 'Jobs, devices, quotes, parts and payments for a shop that mends phones, tablets, laptops and small appliances.',
  summary: 'Ironbridge Repairs is an invented repair shop on a high street. Three technicians work at the bench and one assistant runs the counter, where customers book in a phone with a cracked screen or a coffee machine that has stopped pumping, and come back later to collect it.\n\n'
    + 'Every job is one record. It starts as Booked in, is diagnosed, gets a quote, waits for the customer to approve it, waits for a part if the shelf has none, is repaired and is marked Ready, then Collected. A job can also end as Unrepairable, or Ready with the quote turned down and a small diagnosis charge. The parts and labour each job used are kept beside it, so the charge can be explained line by line, and every payment, deposits included, is recorded against its job.\n\n'
    + 'The records cover about fifteen months up to the day the file is made. Phones come in more in the summer holidays and laptops when the school year starts; kitchen appliances pick up before the year ends. Some jobs are waiting on parts today, some customers have not picked up their devices after weeks, a few repairs have come back under warranty, and the parts shelf has items at or below their reorder level, one of them held up because the supplier has it on back order.',
  suits: [
    'A phone and tablet repair kiosk',
    'A computer support firm that takes machines in',
    'A bike shop workshop, with bikes as the devices',
    'A watch and jewellery repairer',
    'An appliance engineer who also visits homes',
    'A musical instrument or audio repair bench',
  ],
  tables: {
    'Customers': { about: 'One record for each person or business that brings work in. Business customers on account are invoiced and pay later; everyone else pays at the counter.', fields: {
      'First Name': 'The customer\'s first name, or the contact\'s at a business.',
      'Last Name': 'The family name. Needed for every customer.',
      'Company': 'The business name for customers on account. Empty for members of the public.',
      'Phone': 'The number the shop rings or texts when the device is ready.',
      'E-mail': 'Where receipts and reminders go, when the customer gave one.',
      'Address': 'The street address, used for invoices and for home visits.',
      'City': 'The town or city.',
      'Postcode': 'The postcode.',
      'Contact By': 'How the customer wants to hear that the job is ready: Text message, Phone call or E-mail.',
      'On Account': 'Yes for a business that is invoiced and pays by bank transfer after collection.',
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
      'Name': 'What the part is, in plain words.',
      'Fits': 'The kind of device it is for: Phone, Tablet, Laptop or an appliance.',
      'Supplier': 'Who it is bought from.',
      'Cost': 'What one costs the shop.',
      'Price': 'What one is charged at on a job.',
      'In Stock': 'How many are on the shelf now, including any already set aside for a job.',
      'Reorder Level': 'When stock falls to this number the part is ordered again.',
      'Reorder Quantity': 'How many to order at a time.',
      'Stocked': 'Yes for parts kept on the shelf. No for parts ordered in for each job, such as main boards and model-specific keyboards.',
      'Bin': 'Where it is kept on the parts wall.',
    } },
    'Devices': { about: 'Each device a customer has brought in. One customer can have several, and a device that comes back keeps its record, so its history is in one place.', fields: {
      'Customer': 'Who owns it.',
      'Kind': 'Phone, Tablet, Laptop, Coffee machine, Vacuum cleaner, Food mixer or Sewing machine.',
      'Make': 'The maker\'s name (invented makes in this sample).',
      'Model': 'The model name or number.',
      'Serial': 'The serial number, checked when the device is handed back.',
      'Colour': 'Its colour, which helps find it on the shelf.',
    } },
    'Jobs': { about: 'One record for each repair, from booking in to collection. The record number is the job number on the ticket.', fields: {
      'Device': 'The device being repaired, and through it the customer.',
      'Technician': 'Who the job is booked to.',
      'Booked In': 'The day the device was handed in.',
      'Fault': 'The fault as the customer described it at the counter.',
      'Left With It': 'Anything left with the device, such as a charger, case or water tank.',
      'Promised': 'The day the customer was told it would be ready.',
      'Status': 'Booked in, Diagnosing, Awaiting approval, Awaiting parts, Repairing, Ready, Collected or Unrepairable.',
      'Quote': 'The price given for the repair. Fixed-price jobs such as screens are quoted at the counter; others after diagnosis.',
      'Quoted On': 'The day the quote was given.',
      'Quote Approved': 'Yes once the customer has agreed to the quote.',
      'Deposit': 'Money taken up front when a part has to be ordered in.',
      'Outcome': 'How the job ended: Repaired, Quote declined or Beyond economic repair.',
      'Charge': 'What the job finally costs the customer. Nothing for a device that cannot be repaired ("no fix, no fee") or a warranty return.',
      'Finished': 'The day the work was done and the customer could be told.',
      'Collected': 'The day the customer took the device away.',
      'Paid': 'Yes once everything charged has been paid.',
      'Warranty Until': 'The last day of the 90-day warranty on the repair.',
      'Warranty Of': 'For a warranty return, the number of the first job it came back from.',
      'Notes': 'Anything else the counter or the bench needs to know, such as reminders sent.',
    } },
    'Part Orders': { about: 'Parts ordered from suppliers: stock to refill the shelf, or one part for one job.', fields: {
      'Supplier': 'Who the order went to.',
      'Part': 'The part ordered.',
      'Job': 'The job the part is for, when it was ordered for one. Empty for shelf stock.',
      'Quantity': 'How many were ordered.',
      'Unit Cost': 'The price paid for each.',
      'Ordered On': 'The day the order was placed.',
      'Expected': 'The day the supplier said it would arrive.',
      'Received On': 'The day it actually arrived. Empty while it is still on its way.',
      'Status': 'On order or Received.',
    } },
    'Job Parts': { about: 'The parts used on each job, at the price charged.', fields: {
      'Job': 'The job.',
      'Part': 'The part fitted or waiting to be fitted.',
      'Quantity': 'How many.',
      'Unit Price': 'The price charged for each (nothing on a warranty return).',
      'Fitted On': 'The day it was fitted. Empty while the job is waiting for it.',
    } },
    'Labour': { about: 'The bench time on each job: diagnosis, the repair itself, and warranty work.', fields: {
      'Job': 'The job.',
      'Done On': 'The day the work was done.',
      'Description': 'What was done.',
      'Hours': 'How long it took, in hours.',
      'Rate': 'The hourly rate charged. Nought for warranty work and for checks on devices that could not be repaired.',
    } },
    'Payments': { about: 'Every payment taken, each against its job.', fields: {
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
    'Parts on order': 'Everything ordered and not yet arrived, with the expected date and the job it is for when there is one.',
    'Takings by month': 'Money taken each month, deposits and account payments included, and how many payments made it up.',
    'Takings last 30 days': 'The last thirty days of takings split by payment method, for checking against the card machine and the bank statement.',
    'Technician workload': 'How many unfinished jobs each technician has, how many are stuck on parts or quotes, and their oldest job.',
    'Technician turnaround': 'For repairs finished in the last year: how many each technician did, the average days from booking in to finished, and what they were charged at.',
    'Work by device kind': 'Jobs and money by kind of device over the whole period, with the average days each kind takes.',
    'Warranty returns': 'Jobs that came back under warranty, with the number of the first job, so repeat faults on a part or a technician\'s work show up.',
    'Job sheets': 'Every job booked in over the last fourteen days with its full detail: serial, fault, what was left with it, quote and promised date.',
    'Parts margin': 'Each part fitted over the period: how many, what they were charged at, and the margin over cost. Warranty fittings count against the margin.',
    'Accounts to chase': 'Business customers on account who have collected work and not yet paid for it, oldest first.',
  },
  reports: {
    'Job board': 'The job board printed for the bench, one section per stage, each job with its customer, device, fault, technician and promised date.',
    'Ready for collection': 'A calling list for the counter: each device waiting to be collected, the customer\'s number, how they like to be contacted and what they owe, longest waiting first.',
    'Parts to order': 'The parts to reorder, one section per supplier, with the cost of each order. Take it to the phone or the supplier\'s website.',
    'Takings by month': 'Monthly takings with a grand total, for the accountant or the owner\'s month-end look.',
    'Job sheets': 'The last two weeks of jobs, one section per technician, with every detail taken at the counter. Print it to check work in progress against the shelf.',
    'Technician turnaround': 'Each technician\'s finished repairs over the last year, their average turnaround and the value of their work.',
  },
  empty: 'Start with Staff, then Suppliers, then Parts (each part pointing at its supplier, with Stocked set to No for parts you only order in). When a customer comes to the counter, add them to Customers, add their device to Devices, then open a Job for that device with the fault, the technician and the promised date. As the job moves on, change its Status and fill in the quote, the parts used (Job Parts) and the time spent (Labour). Record each payment in Payments as it is taken, and put Finished and Collected on the job when they happen.',
  adapt: [
    'For a bike workshop, use Devices for bikes (Kind as Road, Mountain, Hybrid, E-bike; Serial as the frame number) and add a Booked For date for the day the customer wants it back.',
    'For an engineer who visits homes, add Visit Date and Visit Address to Jobs and a Call-out Fee to the charge, and make a query of visits in the next 7 days ordered by postcode.',
    'Add a Courier Reference to Part Orders, and a query of orders past their Expected date, to chase late suppliers.',
    'Add a status such as Sent to maker for repairs sent away, and add it to the Stage calculation in the Job board query so it sorts in the right place.',
    'Add a Disposal Date to Jobs for devices never collected, and a query of devices uncollected for 90 days or more, if your terms say you may then recycle or sell them.',
    'Give Staff an hourly rate and add a query that sets Labour hours against it, to see what each technician earns the shop.',
    'Add Passcode Given (Yes/No) and Data Backed Up (Yes/No) to Jobs, so the bench knows what it can test before starting.',
  ],
};

// The bench rate per hour, and the charge for checking a device when its quote is turned down.
const RATE = 45;
const DIAGNOSIS = 0.5;
const WARRANTY_DAYS = 90;

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
  // job, at the end: one insert each rather than an insert and updates.
  const later = { 'Part Orders': [], 'Job Parts': [], Labour: [], Payments: [] };
  const keep = (table, values) => { later[table].push(values); return values; };
  // The shop opens Monday to Saturday; anything due on a Sunday happens on the Monday.
  const opened = (d) => (cal.weekday(d) === 0 ? d + 1 : d);
  const slug = (s) => s.toLowerCase().replace(/[^a-z]+/g, '');

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
    const p = { code, cost, price, level, qty, stocked: level > 0, supplier: suppliers[s], stock: 0, reserved: 0, onOrder: false };
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

  // Customers: members of the public as they walk in, and a few businesses on account.
  const area = () => ({ Address: `${r.int(1, 220)} ${r.pick(['Mill Lane', 'Station Road', 'Orchard Way', 'Church Street', 'Canal Walk', 'Quarry Road', 'Cedar Avenue', 'Linden Close', 'Harbour Street'])}`, City: 'Bristol', Postcode: `BS${r.int(1, 16)} ${r.int(1, 9)}${r.pick(['AB', 'DE', 'FG', 'HJ', 'NP', 'QR', 'TU'])}` });
  const walkins = [];
  const newCustomer = (d) => {
    const p = person(r, { domain: r.pick(['mail.example', 'post.example', 'example.com']) });
    const email = r.chance(0.7);
    const c = { devices: [], account: false };
    c.id = add('Customers', {
      'First Name': p['First Name'], 'Last Name': p['Last Name'], Phone: phoneFor(r, 'United Kingdom'), 'E-mail': email ? p['E-mail'] : null,
      ...(r.chance(0.5) ? area() : {}), 'Contact By': r.weighted([['Text message', 6], ['Phone call', 3], ...(email ? [['E-mail', 2]] : [])]), 'On Account': false, Joined: cal.day(d),
    });
    walkins.push(c);
    return c;
  };
  const accounts = companies(r, 5, ['Lettings', 'Surveyors', 'Accountants', 'Dental', 'Architects', 'Solicitors']).map((name) => {
    const p = person(r, { domain: `${slug(name)}.example` });
    const c = { devices: [], account: true };
    c.id = add('Customers', {
      'First Name': p['First Name'], 'Last Name': p['Last Name'], Company: name, Phone: phoneFor(r, 'United Kingdom'), 'E-mail': `accounts@${slug(name)}.example`,
      ...area(), 'Contact By': 'E-mail', 'On Account': true, Joined: cal.day(-days - r.int(30, 700)),
    });
    return c;
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
      fault = `Warranty: ${RETURNS[part?.code] || 'the same fault has come back'}`;
      hours = 0.5; work = 'Warranty repair'; quoted = false; beyond = 0; approve = 1;
    } else {
      const f = chosen || r.weighted(FAULTS[dev.kind].map((x) => [x, x[0]]));
      [, fault, part, hours, work, quoted, beyond, approve] = f;
      part = part ? parts.get(partFor(part, dev)) : null;
    }
    const price = warrantyOf ? 0 : money((part ? part.price : 0) + hours * RATE);
    const promise = opened(d + (PROMISE[dev.kind] || 5) + (quoted ? 1 : 0));
    const job = {
      dev, c, d, part, hours, work, quoted, beyond, approve, price, warrantyOf, tech: warrantyOf ? warrantyOf.tech : technicianFor(dev.kind),
      status: 'Booked in', quote: null, quotedOn: null, approved: false, deposit: null, outcome: null, charge: null,
      finished: null, collected: null, paid: false, warrantyUntil: warrantyOf ? warrantyOf.warrantyUntil : null, notes: [],
      step: 'start', at: d + (r.chance(0.65) ? 0 : 1),
    };
    if (!quoted) { job.quote = price; job.quotedOn = d; job.approved = true; }
    if (r.chance(0.08)) job.notes.push(r.pick(['Customer has a backup of the data.', 'Passcode not given: display tested only.', 'Collect after 5 pm.', 'Second visit for this fault.', 'Customer wants the old part back.']));
    job.row = { Device: dev.id, Technician: job.tech, 'Booked In': cal.day(d), Fault: fault, 'Left With It': r.weighted(LEFT_WITH[dev.kind]), Promised: cal.day(promise) };
    jobs.push(job);
    active.push(job);
    return job;
  };

  const labour = (job, d, description, hours, rate) => keep('Labour', { Job: job, 'Done On': cal.day(d), Description: description, Hours: hours, Rate: rate });
  const collectIn = (job, d) => {
    const wait = r.weighted([[0, 30], [1, 18], [2, 14], [3, 10], [5, 8], [8, 7], [12, 4], [r.int(14, 45), 5], [Infinity, 0.6]]);
    job.step = 'collect';
    // A device never fetched is kept a few months, then the customer comes for it after the final reminder.
    job.at = wait !== Infinity ? opened(d + wait) : d > -100 ? Infinity : opened(d + r.int(60, 95));
  };
  const finish = (job, d, status, outcome, charge) => {
    job.status = status; job.outcome = outcome; job.charge = charge; job.finished = d;
    collectIn(job, d);
  };

  /** The customer has said yes (or the price was fixed): set the part aside, or order it. */
  const goAhead = (job, d) => {
    job.approved = true;
    if (!job.part) { job.status = 'Repairing'; job.step = 'repair'; job.at = d + repairDelay(job.dev.kind); return; }
    const p = job.part;
    job.jobPart = keep('Job Parts', { Job: job, Part: p.id, Quantity: 1, 'Unit Price': job.warrantyOf ? 0 : p.price, 'Fitted On': null });
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
      keep('Payments', { Job: job, 'Paid On': cal.day(d), Amount: p.price, Method: r.chance(0.8) ? 'Card' : 'Cash', Kind: 'Deposit' });
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
      job.step = 'decide'; job.at = job.slow ? d + r.int(12, 16) : d + r.weighted([[0, 10], [1, 25], [2, 22], [3, 15], [r.int(4, 6), 15], [r.int(7, 14), 13]]);
    },
    decide(job, d) {
      if (r.chance(job.approve)) return goAhead(job, d);
      labour(job, job.diagnosed, 'Diagnosis and quote', DIAGNOSIS, RATE);
      return finish(job, d, 'Ready', 'Quote declined', money(DIAGNOSIS * RATE));
    },
    repair(job, d) {
      const p = job.part;
      if (p) { p.stock -= 1; p.reserved -= 1; job.jobPart['Fitted On'] = cal.day(d); }
      const rate = job.warrantyOf ? 0 : RATE;
      if (job.quoted) { labour(job, job.diagnosed, 'Diagnosis', 0.25, rate); labour(job, d, job.work, job.hours - 0.25, rate); }
      else labour(job, d, job.work, job.hours, rate);
      if (!job.warrantyOf) job.warrantyUntil = d + WARRANTY_DAYS;
      finish(job, d, 'Ready', 'Repaired', job.warrantyOf ? 0 : job.price);
    },
    collect(job, d) {
      job.collected = d;
      if (job.status !== 'Unrepairable') job.status = 'Collected';
      const due = money((job.charge || 0) - (job.deposit || 0));
      if (due <= 0) job.paid = true;
      else if (job.c.account) transfers.push({ job, at: cal.workday(d + r.int(10, 45)), amount: due });
      else {
        keep('Payments', { Job: job, 'Paid On': cal.day(d), Amount: due, Method: r.weighted([['Card', 72], ['Cash', 24], ['Bank transfer', 4]]), Kind: 'On collection' });
        job.paid = true;
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
    // Business customers paying their invoices.
    for (let i = transfers.length - 1; i >= 0; i--) {
      const t = transfers[i];
      if (t.at > d) continue;
      transfers.splice(i, 1);
      keep('Payments', { Job: t.job, 'Paid On': cal.day(d), Amount: t.amount, Method: 'Bank transfer', Kind: 'On account' });
      t.job.paid = true;
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
        c = r.pick(accounts);
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
  for (const [table, rows] of Object.entries(later)) for (const row of rows) add(table, row.Job ? { ...row, Job: row.Job.id } : row);
  for (const p of parts.values()) base.update('Parts', p.id, { 'In Stock': p.stock });
  return counts;
}
