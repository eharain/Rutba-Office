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

import { companies, FIRST, LAST, money } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const CLIENT = '[Clients].[First Name] & " " & [Clients].[Last Name]';
const STYLIST = '[Staff].[First Name] & " " & [Staff].[Last Name]';
const WHEN = 'Left([Start], 10) & " " & Mid([Start], 12, 5)';

export const spec = {
  tables: [
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
      { name: 'Patch Test', type: 'yesno' }, { name: 'Active', type: 'yesno', default: true },
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
      { name: 'Sale Price', type: 'currency' }, { name: 'In Stock', type: 'integer', default: 0 },
      { name: 'Reorder Level', type: 'integer', default: 4 }, { name: 'Reorder Quantity', type: 'integer', default: 12 },
      { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Product Sales', fields: [
      { name: 'Date', type: 'date' }, { name: 'Staff', ...ref('Staff') }, { name: 'Client', ...ref('Clients') },
      { name: 'Product', ...ref('Products'), required: true }, { name: 'Quantity', type: 'integer', default: 1 },
      { name: 'Unit Price', type: 'currency' },
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
  },
  reports: [
    { name: 'The column', title: 'The column: the next 7 days by stylist', query: 'Bookings next 7 days', group: 'Stylist', sort: 'When', columns: ['When', 'Client', 'Phone', { field: 'Minutes', type: 'integer' }, { field: 'Deposit', type: 'currency' }], totals: [{ field: 'Minutes', fn: 'sum' }] },
    { name: 'Commission statement', title: 'Commission statement: last month', query: 'Service commission last month', group: 'Stylist', columns: [{ field: 'Day', type: 'date' }, 'Client', 'Service', { field: 'Price', type: 'currency' }, { field: 'Commission', type: 'currency' }], totals: [{ field: 'Price', fn: 'sum' }, { field: 'Commission', fn: 'sum' }] },
    { name: 'Retail commission', title: 'Retail commission: last month', query: 'Retail commission last month', columns: ['Stylist', { field: 'Items Sold', type: 'integer' }, { field: 'Retail Sales', type: 'currency' }, { field: 'Rate', type: 'number' }, { field: 'Commission', type: 'currency' }], totals: [{ field: 'Retail Sales', fn: 'sum' }, { field: 'Commission', fn: 'sum' }] },
    { name: 'Takings by month', query: 'Service takings by month', columns: ['Month', { field: 'Services Done', type: 'integer' }, { field: 'Takings', type: 'currency' }], totals: [{ field: 'Services Done', fn: 'sum' }, { field: 'Takings', fn: 'sum' }] },
    { name: 'Products to order', query: 'Products to order', group: 'Supplier', columns: ['Product', 'Kind', 'Size', { field: 'In Stock', type: 'integer' }, { field: 'Reorder Level', type: 'integer' }, { field: 'Reorder Quantity', type: 'integer' }, { field: 'Order Cost', type: 'currency' }], totals: [{ field: 'Order Cost', fn: 'sum' }] },
    { name: 'Clients to win back', query: 'Clients to win back', columns: ['Client', 'Phone', 'E-mail', { field: 'Last Visit', type: 'date' }, { field: 'Visits', type: 'integer' }, { field: 'Days Since', type: 'integer' }], totals: [{ field: 'Client', fn: 'count' }] },
    { name: 'Price list', query: 'Service menu', group: 'Category', columns: ['Service', { field: 'Minutes', type: 'integer' }, { field: 'Junior Price', type: 'currency' }, { field: 'Stylist Price', type: 'currency' }, { field: 'Senior Price', type: 'currency' }] },
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
  ].join('\n\n'),
  suits: ['A barber shop', 'A nail bar', 'A beauty or brow and lash studio', 'A massage or spa room', 'A dog groomer', 'A mobile hairdresser'],
  tables: {
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
      'Patch Test': 'Yes where a skin test is needed in the six months before, such as colour and lash lifts.', 'Active': 'No takes a service off the menu without losing its history.',
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
      'Cost Price': 'What one unit costs the salon.', 'Sale Price': 'The retail price; empty for back-bar products.',
      'In Stock': 'Units on the shelf now.', 'Reorder Level': 'At or below this many, it goes on the next order.',
      'Reorder Quantity': 'How many to order at a time.', 'Active': 'No for a product no longer stocked.',
    } },
    'Product Sales': { about: 'Retail sold at the desk, one record per product per sale.', fields: {
      'Date': 'The day of the sale.', 'Staff': 'Who sold it, for retail commission.', 'Client': 'Who bought it.',
      'Product': 'What was sold.', 'Quantity': 'How many.', 'Unit Price': 'The price of one, as charged.',
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
  },
  reports: {
    'The column': 'The appointment book for the next seven days, a section for each stylist, with times and phone numbers. Print it for the desk.',
    'Commission statement': 'Last month\'s services for each person on commission, with a total of takings and commission each. For the owner and for payroll.',
    'Retail commission': 'Retail commission due for last month, one line a person.',
    'Takings by month': 'Service takings month by month with a grand total.',
    'Products to order': 'The order list for the week, a section for each supplier, with the cost of each line.',
    'Clients to win back': 'A call or mailing list of lapsed clients who agreed to marketing.',
    'Price list': 'The service menu, printed by category, with the junior, stylist and senior prices.',
  },
  empty: 'Start with Suppliers, then Staff with their level and commission rates, then the Services menu with a price for each level. Add your Products next, with their supplier and reorder level. Clients can be added as they book. Then make Bookings and add the services for each in Booking Services, changing the price where you charge something different. Record Product Sales and Stock Deliveries as they happen, and Gift Vouchers when you sell them.',
  adapt: [
    'A barber can drop the beauty services, keep Junior and Senior prices and add a Walk-in choice to Booked Via.',
    'A dog groomer can rename Clients to Owners and add a Pets table pointing at it (name, breed, size, temperament), then make each booking point at a pet.',
    'To pay commission only above a target, add a Target field to Staff and a calculated column in Takings by stylist last month that takes it off the takings.',
    'Add a Reminder Sent yes/no to Bookings and tick it when the reminder goes, so nobody sends two.',
    'For a nail bar or brow studio, add a Station or Room field to Bookings to show which chair or room each appointment uses.',
    'Change the patch test rule from six months to another period by editing 180 in the Patch tests due query.',
    'Add a Tips field to Bookings and a query that totals tips by stylist each month.',
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
      const id = add('Services', { Name: name, Category: cat, Minutes: minutes, 'Junior Price': junior, 'Stylist Price': stylist, 'Senior Price': senior, 'Patch Test': patch });
      services.set(name, { id, name, minutes, prices: { Junior: junior, Stylist: stylist, Senior: senior, Therapist: stylist }, patch, hair: !['Nails', 'Brows and Lashes', 'Skin and Body'].includes(cat) });
    }
  }
  const S = (name) => services.get(name);

  // The stock.
  const products = [];
  const byName = new Map();
  for (const [kind, cat, si, items] of STOCK) {
    for (const [name, size, cost, price, level, qty, demand] of items) {
      const p = { name, kind, cost, price, level, qty, demand, supplier: suppliers[si], stock: 0, use: 0, pending: false };
      p.id = add('Products', { Name: name, Kind: kind, Category: cat, Supplier: p.supplier.id, Size: size, 'Cost Price': cost, 'Sale Price': price, 'In Stock': 0, 'Reorder Level': level, 'Reorder Quantity': qty });
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

        const price = (x) => money((x.prices[s.level] ?? x.prices.Stylist) * (c.fresh ? 0.8 : 1));
        const bill = money(done.reduce((n, x) => n + price(x), 0));
        let status = 'Booked';
        let cancelledOn = null;
        if (d < 0) status = r.weighted([['Completed', 92.5], ['Cancelled', 5], ['No-show', c.unreliable ? 12 : 1.6]]);
        else if (r.chance(0.025)) status = 'Cancelled';
        if (status === 'Cancelled') cancelledOn = r.int(bookedOn, Math.min(d, 0));
        // A deposit for new clients and long appointments, and for anyone who has missed one before.
        const deposit = bill >= 80 ? 20 : isNew || c.missed ? 10 : 0;
        const late = status === 'No-show' || (status === 'Cancelled' && d - cancelledOn < 2);
        let paidBy = null;
        let voucher = null;
        if (status === 'Completed') {
          const i = voucherQueue.findIndex((x) => x.at <= d && x.v.expires >= d && x.v.Value <= bill);
          if (i >= 0) { voucher = voucherQueue.splice(i, 1)[0].v; paidBy = 'Gift voucher'; } else paidBy = r.weighted([['Card', 8], ['Cash', 1.4]]);
        }
        const [hh, mm] = hhmm(t);
        const booking = add('Bookings', {
          Client: c.id, Staff: s.id, Start: cal.time(d, hh, mm), Minutes: minutes, Status: status, 'Booked On': cal.day(bookedOn),
          'Booked Via': !isNew && d - bookedOn > 20 && r.chance(0.6) ? 'In salon' : r.weighted([['Online', 5], ['Phone', 4]]),
          Deposit: deposit, 'Deposit Refunded': status === 'Cancelled' && deposit > 0 && !late, 'Cancelled On': cancelledOn === null ? null : cal.day(cancelledOn),
          'Paid By': paidBy, Notes: r.chance(isNew ? 0.2 : 0.015) ? (isNew ? BOOKING_NOTES[4] : r.pick(BOOKING_NOTES.slice(0, 4))) : null,
        });
        for (const x of done) add('Booking Services', { Booking: booking, Service: x.id, Price: price(x) });
        if (voucher) { voucher['Redeemed On'] = date; voucher.Booking = booking; }

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
              add('Product Sales', { Date: date, Staff: s.id, Client: c.id, Product: p.id, Quantity: q, 'Unit Price': p.price });
              p.stock -= q;
            }
          }
          if (r.chance(c.unreliable ? 0.08 : 0.035)) c.lapsed = true;
        }
        c.next = status === 'Cancelled' ? d + r.int(4, 14) : status === 'No-show' ? d + r.int(10, 35) : d + Math.round(r.int(c.cadence[0], c.cadence[1]) * (0.9 + r.next() * 0.3));
        t += minutes;
      }
    }
  }

  // Vouchers in the order they were sold.
  for (const v of vouchers) {
    const { expires, ...values } = v;
    add('Gift Vouchers', values);
  }
  // Each client's latest patch test, and the stock as it stands.
  for (const c of allClients) if (c.retested) base.update('Clients', c.id, { 'Patch Test Date': cal.day(c.patch) });
  for (const p of products) base.update('Products', p.id, { 'In Stock': p.stock });
  return counts;
}
