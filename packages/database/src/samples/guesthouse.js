// The guest house: Gullhaven Guest House, an invented bed and breakfast of
// ten rooms on the seafront of Saltby Bay, an invented seaside town. Room
// types with a rate for each season, weeknight and weekend, with or without
// breakfast; the rooms and their housekeeping state; guests and their
// bookings, taken direct, by phone, at the door and through three invented
// booking sites that take a commission; a diary of every room and night;
// what is charged to each room; deposits and payments; the guest's bill; and
// a monthly statement to each booking site of the commission owed.
//
// Its records are fifteen months of letting rooms and six months of
// bookings ahead, made a booking at a time in the order they were taken:
// more people want rooms in summer and at weekends, so those nights fill
// and the rest do not; a room goes to whoever books first, and a cancelled
// booking gives its nights back to the diary. Stays are charged night by
// night with breakfast, dinner, drinks, parking and the rest, billed when
// the guest leaves. The questions a guest house asks are its saved queries,
// and what it prints are its reports and print templates.

import { address, BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, FIRST, LAST, money, PLACES, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const GUEST = '[First Name] & " " & [Last Name]';
const ROOM = '[Rooms].[Number] & " " & [Rooms].[Name]';
const AMOUNT = 'Round([Quantity] * [Unit Price], 2)';
const BOOKED = '[Status] = "Booked"';
// The business's own details at the head of every document.
const LETTERHEAD = { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Website'] };

export const spec = {
  negative: {},
  tables: [
    BUSINESS_TABLE,
    { name: 'Seasons', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Months', type: 'text' },
      { name: 'Minimum Weekend Nights', type: 'integer', default: 1 }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Room Types', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Description', type: 'text' },
      { name: 'Sleeps', type: 'integer', default: 2 }, { name: 'Breakfast Allowance', type: 'currency' },
    ] },
    { name: 'Rates', fields: [
      { name: 'Room Type', ...ref('Room Types'), required: true }, { name: 'Season', ...ref('Seasons'), required: true },
      { name: 'Plan', type: 'text', default: 'Bed and breakfast' }, { name: 'Weeknight Rate', type: 'currency' }, { name: 'Weekend Rate', type: 'currency' },
    ] },
    { name: 'Rooms', fields: [
      { name: 'Number', type: 'integer', required: true, unique: true }, { name: 'Name', type: 'text', required: true },
      { name: 'Room Type', ...ref('Room Types'), required: true }, { name: 'Floor', type: 'text' }, { name: 'Outlook', type: 'text' },
      { name: 'Beds', type: 'text' }, { name: 'Dogs Allowed', type: 'yesno' }, { name: 'Housekeeping', type: 'text', default: 'Vacant clean' },
      { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Booking Sources', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Code', type: 'text', unique: true },
      { name: 'Kind', type: 'text', default: 'Direct' }, { name: 'Commission Rate', type: 'number', default: 0 },
      { name: 'Account Ref', type: 'text' }, { name: 'Contact', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' },
      { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' }, { name: 'Country', type: 'text' },
      { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Guests', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
      { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' }, { name: 'Country', type: 'text' },
      { name: 'Nationality', type: 'text' }, { name: 'Passport Number', type: 'text' }, { name: 'Passport Issued By', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Car Registration', type: 'text' },
      { name: 'Mailing List', type: 'yesno' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Extras', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Kind', type: 'text' }, { name: 'Price', type: 'currency' },
      { name: 'Per', type: 'text' }, { name: 'Tax Rate', type: 'number', default: 0.2 }, { name: 'Available', type: 'yesno', default: true },
    ] },
    { name: 'Commission Statements', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Source', ...ref('Booking Sources'), required: true },
      { name: 'Month', type: 'text' }, { name: 'Statement Date', type: 'date' }, { name: 'Stays', type: 'integer', default: 0 },
      { name: 'Room Takings', type: 'currency' }, { name: 'Commission', type: 'currency' }, { name: 'Paid On', type: 'date' },
    ] },
    // Bookings point at Guests before Rooms and Sources, so a query on bookings and guests joins the guest who booked.
    { name: 'Bookings', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Guest', ...ref('Guests'), required: true }, { name: 'Room', ...ref('Rooms'), required: true },
      { name: 'Source', ...ref('Booking Sources'), required: true }, { name: 'Site Reference', type: 'text' }, { name: 'Booked On', type: 'date' },
      { name: 'Arrival', type: 'date' }, { name: 'Departure', type: 'date' }, { name: 'Nights', type: 'integer', default: 1 },
      { name: 'Adults', type: 'integer', default: 2 }, { name: 'Children', type: 'integer', default: 0 }, { name: 'Rate Plan', type: 'text', default: 'Bed and breakfast' },
      { name: 'Room Total', type: 'currency' }, { name: 'Commission Rate', type: 'number', default: 0 }, { name: 'Commission', type: 'currency', default: 0 },
      { name: 'Deposit Due', type: 'currency', default: 0 }, { name: 'Deposit Paid', type: 'currency', default: 0 },
      { name: 'Status', type: 'text', default: 'Confirmed' }, { name: 'Expected Arrival', type: 'text' },
      { name: 'Checked In', type: 'datetime' }, { name: 'Checked Out', type: 'datetime' }, { name: 'Cancelled On', type: 'date' },
      { name: 'Car Registration', type: 'text' }, { name: 'Statement', ...ref('Commission Statements') }, { name: 'Requests', type: 'longtext' },
    ] },
    { name: 'Room Nights', fields: [
      { name: 'Room', ...ref('Rooms'), required: true }, { name: 'Night', type: 'date', required: true }, { name: 'Weekday', type: 'text' },
      { name: 'Status', type: 'text', default: 'Free' }, { name: 'Booking', ...ref('Bookings') }, { name: 'Rate', type: 'currency' },
    ] },
    { name: 'Bills', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Booking', ...ref('Bookings'), required: true }, { name: 'Bill Date', type: 'date' },
      { name: 'Total', type: 'currency' }, { name: 'VAT', type: 'currency' }, { name: 'Deposit', type: 'currency', default: 0 },
      { name: 'Balance', type: 'currency' }, { name: 'Paid By', type: 'text' },
    ] },
    { name: 'Charges', fields: [
      { name: 'Booking', ...ref('Bookings'), required: true }, { name: 'Bill', ...ref('Bills') }, { name: 'Date', type: 'date' },
      { name: 'Kind', type: 'text' }, { name: 'Extra', ...ref('Extras') }, { name: 'Description', type: 'text' },
      { name: 'Quantity', type: 'integer', default: 1 }, { name: 'Unit Price', type: 'currency' }, { name: 'Tax Rate', type: 'number', default: 0.2 },
    ] },
    { name: 'Payments', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Booking', ...ref('Bookings'), required: true }, { name: 'Date', type: 'date' },
      { name: 'Kind', type: 'text' }, { name: 'Method', type: 'text' }, { name: 'Amount', type: 'currency' },
      { name: 'Refunded On', type: 'date' }, { name: 'Refund', type: 'currency', default: 0 },
    ] },
  ],
  queries: {
    'Arrivals in the next 7 days': { tables: ['Bookings', 'Guests', 'Rooms', 'Booking Sources'], columns: [
      { table: 'Bookings', field: 'Arrival', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Arrival])', as: 'Days Until', criteria: ['Between 0 And 7'] },
      { table: 'Rooms', field: 'Number', sort: 'asc', show: false },
      { table: 'Bookings', field: 'Number', as: 'Booking No' }, { expr: GUEST, as: 'Guest' }, { expr: ROOM, as: 'Room' },
      { table: 'Bookings', field: 'Nights' }, { table: 'Bookings', field: 'Adults' }, { table: 'Bookings', field: 'Children' },
      { table: 'Booking Sources', field: 'Name', as: 'Source' }, { table: 'Bookings', field: 'Expected Arrival' },
      { table: 'Bookings', field: 'Deposit Paid' }, { expr: 'Round([Room Total] - Nz([Deposit Paid]), 2)', as: 'Balance' },
      { table: 'Bookings', field: 'Requests' }, { table: 'Bookings', field: 'Status', criteria: ['Confirmed'], show: false },
    ] },
    'Departures in the next 7 days': { tables: ['Bookings', 'Guests', 'Rooms'], columns: [
      { table: 'Bookings', field: 'Departure', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Departure])', as: 'Days Until', criteria: ['Between 0 And 7'] },
      { table: 'Rooms', field: 'Number', sort: 'asc', show: false },
      { table: 'Bookings', field: 'Number', as: 'Booking No' }, { expr: GUEST, as: 'Guest' }, { expr: ROOM, as: 'Room' },
      { table: 'Bookings', field: 'Arrival' }, { table: 'Bookings', field: 'Nights' }, { table: 'Guests', field: 'Phone' },
      { table: 'Bookings', field: 'Room Total' }, { table: 'Bookings', field: 'Deposit Paid' },
      { expr: 'Round([Room Total] - Nz([Deposit Paid]), 2)', as: 'Balance on Room' },
      { table: 'Bookings', field: 'Status', criteria: ['In ("Confirmed", "Checked in")'] },
    ] },
    'Rooms free tonight and this weekend': { tables: ['Room Nights', 'Rooms', 'Room Types'], columns: [
      { table: 'Room Nights', field: 'Night', sort: 'asc' }, { table: 'Room Nights', field: 'Weekday', criteria: ['', 'In ("Fri", "Sat")'] },
      { table: 'Rooms', field: 'Number', as: 'Room No', sort: 'asc' }, { table: 'Rooms', field: 'Name', as: 'Room' },
      { table: 'Room Types', field: 'Name', as: 'Room Type' }, { table: 'Room Types', field: 'Sleeps' }, { table: 'Rooms', field: 'Outlook' },
      { expr: 'DateDiff("d", Date(), [Night])', as: 'Days Ahead', criteria: ['0', 'Between 0 And 6'], show: false },
      { table: 'Room Nights', field: 'Status', criteria: ['Free', 'Free'], show: false },
    ] },
    'Occupancy by month': { tables: ['Room Nights'], columns: [
      { expr: 'Format([Night], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { table: 'Room Nights', field: 'ID', as: 'Rooms Available', total: 'count' },
      { expr: `IIf(${BOOKED}, 1, 0)`, as: 'Nights Sold', total: 'sum' },
      { expr: `IIf(${BOOKED}, 100, 0)`, as: 'Occupancy %', total: 'avg' },
      { expr: `IIf(${BOOKED}, [Rate], Null)`, as: 'Average Rate', total: 'avg' },
      { expr: `IIf(${BOOKED}, Nz([Rate]), 0)`, as: 'Revenue per Room', total: 'avg' },
      { expr: `IIf(${BOOKED}, Nz([Rate]), 0)`, as: 'Room Revenue', total: 'sum' },
      { table: 'Room Nights', field: 'Status', total: 'where', criteria: ['<> "Out of order"'] },
      { expr: 'DateDiff("m", [Night], Date())', as: 'Months Ago', total: 'where', criteria: ['Between -2 And 12'] },
    ] },
    'Revenue by source': { tables: ['Bookings', 'Booking Sources'], columns: [
      { table: 'Booking Sources', field: 'Name', as: 'Source', total: 'group' }, { table: 'Booking Sources', field: 'Kind', total: 'group' },
      { table: 'Bookings', field: 'ID', as: 'Bookings', total: 'count' }, { table: 'Bookings', field: 'Nights', as: 'Nights', total: 'sum' },
      { table: 'Bookings', field: 'Room Total', as: 'Room Takings', total: 'sum' }, { table: 'Bookings', field: 'Commission', as: 'Commission', total: 'sum' },
      { expr: 'Round([Room Total] - Nz([Bookings].[Commission]), 2)', as: 'After Commission', total: 'sum', sort: 'desc' },
      { table: 'Bookings', field: 'Status', total: 'where', criteria: ['In ("Checked in", "Checked out")'] },
      { expr: 'DateDiff("d", [Arrival], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Deposits outstanding': { tables: ['Bookings', 'Guests'], columns: [
      { table: 'Bookings', field: 'Arrival', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Arrival])', as: 'Days to Arrival' },
      { table: 'Bookings', field: 'Number', as: 'Booking No' }, { expr: GUEST, as: 'Guest' }, { table: 'Guests', field: 'Phone' }, { table: 'Guests', field: 'E-mail' },
      { table: 'Bookings', field: 'Booked On' }, { table: 'Bookings', field: 'Deposit Due' }, { table: 'Bookings', field: 'Deposit Paid' },
      { expr: 'Round([Deposit Due] - Nz([Deposit Paid]), 2)', as: 'Owed', criteria: ['> 0'] },
      { table: 'Bookings', field: 'Status', criteria: ['Confirmed'], show: false },
    ] },
    'Returning guests': { tables: ['Guests', 'Bookings'], columns: [
      { table: 'Guests', field: 'ID', total: 'group', show: false }, { expr: GUEST, as: 'Guest', total: 'group' },
      { table: 'Guests', field: 'City', total: 'group' }, { table: 'Guests', field: 'Country', total: 'group' }, { table: 'Guests', field: 'E-mail', total: 'group' },
      { table: 'Bookings', field: 'ID', as: 'Stays', total: 'count', criteria: ['> 1'], sort: 'desc' },
      { table: 'Bookings', field: 'Arrival', as: 'First Stay', total: 'min' }, { table: 'Bookings', field: 'Departure', as: 'Last Stay', total: 'max' },
      { table: 'Bookings', field: 'Nights', as: 'Nights', total: 'sum' }, { table: 'Bookings', field: 'Room Total', as: 'Spent on Rooms', total: 'sum' },
      { table: 'Bookings', field: 'Status', total: 'where', criteria: ['Checked out'] },
    ] },
    'Housekeeping list for today': { tables: ['Bookings', 'Rooms', 'Guests'], columns: [
      { table: 'Rooms', field: 'Number', as: 'Room No', sort: 'asc' }, { table: 'Rooms', field: 'Name', as: 'Room' },
      { expr: 'IIf(DateDiff("d", Date(), [Departure]) = 0, 1, IIf(DateDiff("d", [Arrival], Date()) = 0, 3, 2))', as: 'Order', sort: 'asc', show: false },
      { expr: 'IIf(DateDiff("d", Date(), [Departure]) = 0, "Departure: strip the beds and clean in full", IIf(DateDiff("d", [Arrival], Date()) = 0, "Arrival: make up, check and put out the welcome tray", "Stayover: tidy, fresh towels, refill the tray"))', as: 'Task' },
      { table: 'Rooms', field: 'Housekeeping' }, { expr: GUEST, as: 'Guest' }, { table: 'Bookings', field: 'Adults' }, { table: 'Bookings', field: 'Children' },
      { table: 'Bookings', field: 'Expected Arrival' }, { table: 'Bookings', field: 'Requests' },
      { expr: 'DateDiff("d", [Arrival], Date())', as: 'Since Arrival', criteria: ['>= 0'], show: false },
      { expr: 'DateDiff("d", Date(), [Departure])', as: 'To Departure', criteria: ['>= 0'], show: false },
      { table: 'Bookings', field: 'Status', criteria: ['In ("Confirmed", "Checked in")'], show: false },
    ] },
    'Room status': { tables: ['Rooms', 'Room Types'], columns: [
      { table: 'Rooms', field: 'Number', as: 'Room No', sort: 'asc' }, { table: 'Rooms', field: 'Name', as: 'Room' }, { table: 'Room Types', field: 'Name', as: 'Room Type' },
      { table: 'Rooms', field: 'Floor' }, { table: 'Rooms', field: 'Outlook' }, { table: 'Rooms', field: 'Beds' }, { table: 'Room Types', field: 'Sleeps' },
      { table: 'Rooms', field: 'Housekeeping' }, { table: 'Rooms', field: 'Notes' },
    ] },
    'Guest register': { tables: ['Bookings', 'Guests', 'Rooms'], columns: [
      { expr: 'Format([Arrival], "yyyy-mm")', as: 'Month' }, { table: 'Bookings', field: 'Arrival', sort: 'asc' }, { table: 'Rooms', field: 'Number', sort: 'asc', show: false },
      { expr: GUEST, as: 'Guest' }, { table: 'Guests', field: 'Nationality' }, { table: 'Guests', field: 'Passport Number' }, { table: 'Guests', field: 'Passport Issued By' },
      { expr: '[Guests].[Address] & ", " & [Guests].[City] & IIf(Nz([Guests].[Postcode], "") = "", "", " " & [Guests].[Postcode]) & ", " & [Guests].[Country]', as: 'Home Address' },
      { expr: ROOM, as: 'Room' }, { table: 'Bookings', field: 'Departure' }, { table: 'Bookings', field: 'Car Registration' },
      { table: 'Bookings', field: 'Status', criteria: ['In ("Checked in", "Checked out")'], show: false },
      { expr: 'DateDiff("d", [Arrival], Date())', as: 'Days Ago', criteria: ['Between 0 And 365'], show: false },
    ] },
    'Guests in the house': { tables: ['Bookings', 'Guests', 'Rooms', 'Charges'], columns: [
      { table: 'Rooms', field: 'Number', as: 'Room No', total: 'group', sort: 'asc' }, { table: 'Rooms', field: 'Name', as: 'Room', total: 'group' },
      { table: 'Bookings', field: 'Number', as: 'Booking No', total: 'group' }, { expr: GUEST, as: 'Guest', total: 'group' },
      { table: 'Bookings', field: 'Arrival', total: 'group' }, { table: 'Bookings', field: 'Departure', total: 'group' },
      { expr: AMOUNT, as: 'Charged So Far', total: 'sum' }, { table: 'Bookings', field: 'Deposit Paid', as: 'Deposit Paid', total: 'max' },
      { table: 'Bookings', field: 'Status', total: 'where', criteria: ['Checked in'] },
    ] },
    'Extras sold by month': { tables: ['Charges', 'Extras'], columns: [
      { expr: 'Format([Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Extras', field: 'Name', as: 'Extra', total: 'group', sort: 'asc' },
      { table: 'Charges', field: 'Quantity', as: 'Sold', total: 'sum' }, { expr: AMOUNT, as: 'Takings', total: 'sum' },
      { expr: 'DateDiff("d", [Date], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Cancellations and no-shows': { tables: ['Bookings', 'Guests', 'Booking Sources'], columns: [
      { table: 'Bookings', field: 'Arrival', sort: 'desc' }, { table: 'Bookings', field: 'Number', as: 'Booking No' }, { expr: GUEST, as: 'Guest' },
      { table: 'Booking Sources', field: 'Name', as: 'Source' }, { table: 'Bookings', field: 'Status', criteria: ['In ("Cancelled", "No-show")'] },
      { table: 'Bookings', field: 'Nights' }, { table: 'Bookings', field: 'Room Total', as: 'Value' }, { table: 'Bookings', field: 'Cancelled On' },
      { expr: 'DateDiff("d", [Cancelled On], [Arrival])', as: 'Days Notice' }, { table: 'Bookings', field: 'Deposit Paid', as: 'Deposit Kept' },
      { expr: 'DateDiff("d", [Arrival], Date())', as: 'Days Ago', criteria: ['Between -180 And 365'], show: false },
    ] },
    'Commission by month': { tables: ['Commission Statements', 'Booking Sources'], columns: [
      { table: 'Commission Statements', field: 'Statement Date', sort: 'desc' }, { table: 'Commission Statements', field: 'Number', as: 'Statement' },
      { table: 'Commission Statements', field: 'Month' }, { table: 'Booking Sources', field: 'Name', as: 'Site', sort: 'asc' },
      { table: 'Commission Statements', field: 'Stays' }, { table: 'Commission Statements', field: 'Room Takings' },
      { table: 'Commission Statements', field: 'Commission' }, { table: 'Commission Statements', field: 'Paid On' },
      { expr: 'DateDiff("d", [Statement Date], Date())', as: 'Days Ago', criteria: ['Between 0 And 400'], show: false },
    ] },
    Tariff: { tables: ['Rates', 'Room Types', 'Seasons'], columns: [
      { table: 'Room Types', field: 'ID', sort: 'asc', show: false }, { table: 'Room Types', field: 'Name', as: 'Room Type' },
      { table: 'Seasons', field: 'ID', sort: 'asc', show: false }, { table: 'Seasons', field: 'Name', as: 'Season' }, { table: 'Seasons', field: 'Months' },
      { table: 'Rates', field: 'Plan', sort: 'asc' }, { table: 'Rates', field: 'Weeknight Rate' }, { table: 'Rates', field: 'Weekend Rate' },
    ] },
  },
  reports: [
    { name: 'Arrivals list', title: 'Arrivals in the next 7 days', query: 'Arrivals in the next 7 days', group: 'Arrival', columns: ['Booking No', 'Guest', 'Room', { field: 'Nights', type: 'integer' }, { field: 'Adults', type: 'integer' }, { field: 'Children', type: 'integer' }, 'Source', 'Expected Arrival', { field: 'Balance', type: 'currency' }], totals: [{ field: 'Nights', fn: 'sum' }, { field: 'Balance', fn: 'sum' }] },
    { name: 'Departures list', title: 'Departures in the next 7 days', query: 'Departures in the next 7 days', group: 'Departure', columns: ['Booking No', 'Guest', 'Room', { field: 'Nights', type: 'integer' }, 'Phone', { field: 'Room Total', type: 'currency' }, { field: 'Deposit Paid', type: 'currency' }, { field: 'Balance on Room', type: 'currency' }], totals: [{ field: 'Balance on Room', fn: 'sum' }] },
    { name: 'Housekeeping sheet', title: 'Housekeeping for today', query: 'Housekeeping list for today', columns: [{ field: 'Room No', type: 'integer' }, 'Room', 'Task', 'Housekeeping', 'Guest', 'Expected Arrival', 'Requests'], totals: [{ field: 'Room No', fn: 'count' }] },
    { name: 'Monthly occupancy', title: 'Occupancy and room rate by month', query: 'Occupancy by month', columns: ['Month', { field: 'Rooms Available', type: 'integer' }, { field: 'Nights Sold', type: 'integer' }, { field: 'Occupancy %', type: 'number' }, { field: 'Average Rate', type: 'currency' }, { field: 'Revenue per Room', type: 'currency' }, { field: 'Room Revenue', type: 'currency' }], totals: [{ field: 'Nights Sold', fn: 'sum' }, { field: 'Room Revenue', fn: 'sum' }] },
    { name: 'Takings by source', title: 'Room takings by source, last 12 months', query: 'Revenue by source', columns: ['Source', 'Kind', { field: 'Bookings', type: 'integer' }, { field: 'Nights', type: 'integer' }, { field: 'Room Takings', type: 'currency' }, { field: 'Commission', type: 'currency' }, { field: 'After Commission', type: 'currency' }], totals: [{ field: 'Room Takings', fn: 'sum' }, { field: 'Commission', fn: 'sum' }, { field: 'After Commission', fn: 'sum' }] },
    { name: 'Deposits to chase', title: 'Deposits not yet paid', query: 'Deposits outstanding', columns: [{ field: 'Arrival', type: 'date' }, 'Booking No', 'Guest', 'Phone', 'E-mail', { field: 'Booked On', type: 'date' }, { field: 'Deposit Due', type: 'currency' }, { field: 'Owed', type: 'currency' }], totals: [{ field: 'Owed', fn: 'sum' }] },
    { name: 'Guest register', title: 'Register of guests', query: 'Guest register', group: 'Month', columns: [{ field: 'Arrival', type: 'date' }, 'Guest', 'Nationality', 'Passport Number', 'Passport Issued By', 'Home Address', 'Room', { field: 'Departure', type: 'date' }, 'Car Registration'], totals: [{ field: 'Guest', fn: 'count' }] },
    { name: 'Tariff card', title: 'Room rates', query: 'Tariff', group: 'Room Type', columns: ['Season', 'Months', 'Plan', { field: 'Weeknight Rate', type: 'currency' }, { field: 'Weekend Rate', type: 'currency' }] },
    { name: 'Commission by site', title: 'Commission statements by booking site', query: 'Commission by month', group: 'Site', columns: ['Statement', 'Month', { field: 'Stays', type: 'integer' }, { field: 'Room Takings', type: 'currency' }, { field: 'Commission', type: 'currency' }, { field: 'Paid On', type: 'date' }], totals: [{ field: 'Room Takings', fn: 'sum' }, { field: 'Commission', fn: 'sum' }] },
  ],
  documents: [
    { name: 'Booking confirmation', table: 'Bookings', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Booking confirmation {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Guest', label: 'Guest' }, { field: 'Number', label: 'Booking' },
        { field: 'Guest.Address', label: 'Address' }, { field: 'Booked On', label: 'Booked on' },
        { field: 'Guest.City', label: 'Town' }, { field: 'Arrival', label: 'Arriving' },
        { field: 'Guest.Postcode', label: 'Postcode' }, { field: 'Departure', label: 'Leaving' },
        { field: 'Guest.Country', label: 'Country' }, { field: 'Nights', label: 'Nights' },
        { field: 'Room.Number', label: 'Room' }, { field: 'Room.Name', label: 'Room name' },
        { field: 'Room.Room Type.Name', label: 'Room type' }, { field: 'Rate Plan', label: 'Rate' },
        { field: 'Adults', label: 'Adults' }, { field: 'Children', label: 'Children' },
      ] },
      { type: 'totals', rows: [
        { label: 'Room charge for the stay', expr: '[Room Total]' }, { label: 'Deposit asked for', expr: 'Nz([Deposit Due])' },
        { label: 'Deposit received', expr: 'Nz([Deposit Paid])' }, { label: 'Balance due on departure', expr: '[Room Total] - Nz([Deposit Paid])', strong: true },
      ] },
      { type: 'text', text: 'Dear {Guest.First Name},' },
      { type: 'text', text: 'Thank you for booking with us. This letter confirms room {Room.Number}, {Room.Name}, from {Arrival} to {Departure}, {Nights} nights, at the rate shown above. Our rates include VAT. Where the rate is bed and breakfast, breakfast is included for the guests booked; on a room only rate it can be added on the morning.' },
      { type: 'text', text: 'Deposit. To hold a room booked with us directly we ask for a deposit of the first night. What we have received so far is shown above. If it is not yet paid, please pay it within three days of this letter, by card on {Business.Phone} or by bank transfer to {Business.Bank Details}, quoting {Number}. The balance and anything charged to the room are paid when you leave.' },
      { type: 'text', text: 'Arriving and leaving. Check-in is from 3 pm to 8 pm. If you will be later than 8 pm, please ring us and we will leave your key in the lock box by the front door. On the day you leave the room is needed by 10.30 am, and you are welcome to leave your bags with us. A late check-out until 1 pm can be arranged for a small charge when the room is not needed.' },
      { type: 'text', text: 'Cancelling. If you cancel seven days or more before you arrive, we return your deposit in full. If you cancel later than that, we keep the deposit unless we let the room to someone else. If you do not arrive and have not told us, the first night is charged.' },
      { type: 'text', text: 'Finding us. We are at {Business.Address}, {Business.City} {Business.Postcode}, the white house with blue window frames facing the bandstand. By car, follow the signs for the seafront and turn left along the Esplanade. We have four parking spaces behind the house, to be booked with your room, and there is a long-stay car park on Station Road two minutes away. From the railway station it is a ten-minute walk downhill towards the sea.' },
      { type: 'text', text: 'Breakfast is served from 8 am to 9.30 am, and from 9 am to 10 am on Sundays. Please tell us of any allergy or diet before you come. If there is anything else we can do, write to {Business.E-mail} or ring us. We look forward to welcoming you.' },
      { type: 'spacer', size: 24 },
      { type: 'signatures', labels: ['With best wishes, for the guest house'] },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    { name: 'Guest bill', table: 'Bills', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Bill and VAT invoice {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Booking.Guest', label: 'Guest' }, { field: 'Number', label: 'Invoice number' },
        { field: 'Booking.Guest.Address', label: 'Address' }, { field: 'Bill Date', label: 'Date' },
        { field: 'Booking.Guest.City', label: 'Town' }, { field: 'Booking.Number', label: 'Booking' },
        { field: 'Booking.Guest.Postcode', label: 'Postcode' }, { field: 'Booking.Room.Number', label: 'Room' },
        { field: 'Booking.Guest.Country', label: 'Country' }, { field: 'Booking.Arrival', label: 'Arrived' },
        { field: 'Booking.Rate Plan', label: 'Rate' }, { field: 'Booking.Departure', label: 'Left' },
      ] },
      { type: 'lines', table: 'Charges', link: 'Bill', sort: 'Date', columns: [
        { field: 'Date' }, { field: 'Description' }, { field: 'Quantity', type: 'integer' }, { field: 'Unit Price', label: 'Price', type: 'currency' },
        { expr: AMOUNT, as: 'Line Amount', label: 'Amount', type: 'currency' }, { expr: 'Nz([Tax Rate]) * 100', as: 'VAT Rate', label: 'VAT %', type: 'number' },
        { expr: 'Round([Line Amount] - [Line Amount] / (1 + Nz([Tax Rate])), 2)', as: 'Line VAT', label: 'VAT within', type: 'currency' },
      ] },
      { type: 'totals', rows: [
        { label: 'Total for your stay', expr: '[Line Amount]' }, { label: 'Of which VAT', expr: '[Line VAT]' },
        { label: 'Total before VAT', expr: '[Total for your stay] - [Of which VAT]' }, { label: 'Less deposit paid', expr: 'Nz([Deposit])' },
        { label: 'Balance paid on departure', expr: '[Total for your stay] - Nz([Deposit])', strong: true },
      ] },
      { type: 'text', text: 'Balance paid by {Paid By}. Prices include VAT at the rate shown on each line. VAT registration {Business.Tax Number}.', small: true },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    { name: 'Receipt', table: 'Payments', paper: 'receipt-80', letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone'] }, blocks: [
      { type: 'heading', text: 'Receipt {Number}' },
      { type: 'fields', columns: 1, items: [
        { field: 'Date' }, { field: 'Booking.Guest', label: 'Received from' }, { field: 'Booking.Number', label: 'Booking' },
        { field: 'Booking.Room.Number', label: 'Room' }, { field: 'Booking.Arrival', label: 'Arrival' }, { field: 'Booking.Departure', label: 'Departure' },
        { field: 'Kind', label: 'For' }, { field: 'Method', label: 'Paid by' }, { field: 'Refunded On', label: 'Refunded on' },
      ] },
      { type: 'totals', rows: [{ label: 'Amount received', expr: '[Amount]', strong: true }, { label: 'Refunded', expr: 'Nz([Refund])' }] },
      { type: 'text', text: 'Thank you. Your VAT invoice is the bill you are given when you leave.', small: true },
    ] },
    { name: 'Registration card', table: 'Bookings', paper: 'a5', letterhead: { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone'] }, blocks: [
      { type: 'heading', text: 'Registration card' },
      { type: 'fields', columns: 2, items: [
        { field: 'Number', label: 'Booking' }, { field: 'Room.Number', label: 'Room' },
        { field: 'Arrival', label: 'Arriving' }, { field: 'Departure', label: 'Leaving' },
        { field: 'Adults', label: 'Adults' }, { field: 'Children', label: 'Children' },
      ] },
      { type: 'fields', columns: 2, items: [
        { field: 'Guest', label: 'Full name' }, { field: 'Guest.Nationality', label: 'Nationality' },
        { field: 'Guest.Passport Number', label: 'Passport number' }, { field: 'Guest.Passport Issued By', label: 'Passport issued by' },
        { field: 'Guest.Address', label: 'Home address' }, { field: 'Guest.City', label: 'Town' },
        { field: 'Guest.Postcode', label: 'Postcode' }, { field: 'Guest.Country', label: 'Country' },
        { field: 'Guest.Phone', label: 'Phone' }, { field: 'Guest.E-mail', label: 'E-mail' },
        { field: 'Car Registration', label: 'Car registration' }, { field: 'Expected Arrival', label: 'Expected at' },
      ] },
      { type: 'text', text: 'Passport details are filled in only for guests from countries where we are asked to record them, and left blank for others. Please write the name and nationality of any other guest over 16 sharing the room on the line below.' },
      { type: 'text', text: 'Other guests: ..................................................................................' },
      { type: 'text', text: 'I confirm these details are correct. Breakfast is from 8 am to 9.30 am, rooms are needed by 10.30 am on the day you leave, and we ask for quiet after 11 pm.', small: true },
      { type: 'spacer', size: 16 },
      { type: 'signatures', labels: ['Guest signature', 'Date'] },
    ] },
    { name: 'Commission statement', table: 'Commission Statements', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Commission statement {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Source.Name', label: 'To' }, { field: 'Number', label: 'Statement' },
        { field: 'Source.Address', label: 'Address' }, { field: 'Statement Date', label: 'Date' },
        { field: 'Source.City', label: 'City' }, { field: 'Month', label: 'Stays ending in' },
        { field: 'Source.Postcode', label: 'Postcode' }, { field: 'Source.Account Ref', label: 'Our account' },
        { field: 'Source.Country', label: 'Country' }, { field: 'Paid On', label: 'Paid on' },
      ] },
      { type: 'lines', table: 'Bookings', link: 'Statement', sort: 'Departure', columns: [
        { field: 'Number', label: 'Booking' }, { field: 'Site Reference', label: 'Your reference' }, { field: 'Guest' },
        { field: 'Arrival' }, { field: 'Departure' }, { field: 'Nights', type: 'integer' },
        { field: 'Room Total', label: 'Room takings', type: 'currency' }, { expr: 'Nz([Commission Rate]) * 100', as: 'Rate %', type: 'number' },
        { field: 'Commission', type: 'currency' },
      ] },
      { type: 'totals', rows: [{ label: 'Room takings', expr: '[Room Total]' }, { label: 'Commission owed', expr: '[Commission]', strong: true }] },
      { type: 'text', text: 'These are the stays booked through you that ended in {Month}, with the commission on each at the rate we agreed. We will pay the commission owed by bank transfer within 14 days of this statement. Please tell us of any difference before then, quoting {Number}.' },
      { type: 'text', text: '{Business.Name}, {Business.Address}, {Business.City} {Business.Postcode}. {Business.E-mail}', small: true },
    ] },
  ],
};

export const about = {
  name: 'Gullhaven Guest House',
  kind: 'Guest house or B and B',
  line: 'Rooms, rates, guests, bookings, deposits, extras, bills and booking-site commission for a small guest house.',
  summary: [
    'Gullhaven Guest House is an invented bed and breakfast of ten rooms on the seafront of Saltby Bay, an invented seaside town in the United Kingdom. It has single, double, twin, family and superior sea-view rooms, each type with a rate for the low, mid and high season, a weeknight rate and a dearer one for Friday and Saturday nights, with breakfast or room only.',
    'The sample holds fifteen months of stays and six months of bookings ahead. Guests book on the website or by e-mail, by phone, at the door, or through three invented booking sites that each take a commission. The rooms fill in July and August and at weekends, and stand half empty on winter weeknights. A booking direct asks for the first night as a deposit. Some guests cancel, early enough to have the deposit back or too late, and a few never arrive. Each night of a stay is charged to the room with the extras taken: breakfast on a room only rate, an evening meal, drinks, parking, a dog, a late check-out. The guest pays the bill on leaving.',
    'A diary, the Room Nights table, holds one record for every room and every night, marked Booked, Free or Out of order. It is the paper booking chart of a guest house in a table: rooms free tonight, and the occupancy and average rate of every month, come from it. At the end of each month each booking site is sent a statement of the stays it sent and the commission owed.',
  ],
  suits: ['A bed and breakfast or guest house', 'A small hotel or inn with rooms', 'A pub with letting rooms', 'Holiday cottages let by the night or week', 'A hostel with private rooms', 'A farm stay or glamping site', 'Serviced rooms let to contractors'],
  tables: {
    Business: BUSINESS_ABOUT,
    Seasons: { about: 'The seasons the rates follow: low, mid and high.', fields: {
      Name: 'The season\'s name. Each name is used once.', Months: 'The months it covers, in words.',
      'Minimum Weekend Nights': 'The fewest nights a stay that takes in a Saturday may be: 2 means a Saturday night is let only with the Friday or Sunday.',
      Notes: 'Anything else about the season: bank holidays, school holidays, Christmas.',
    } },
    'Room Types': { about: 'The kinds of room you let. Each room is one of these, and the rates are set for each type.', fields: {
      Name: 'The type\'s name, such as Double or Superior sea view.', Description: 'What the room has, as you would put it on your website.',
      Sleeps: 'The most people the room takes.', 'Breakfast Allowance': 'What the room only rate is below the bed and breakfast rate for this type.',
    } },
    Rates: { about: 'The tariff: what one night costs for each room type, season and plan, VAT included.', fields: {
      'Room Type': 'The type of room.', Season: 'The season the rate is for.', Plan: 'Bed and breakfast, or Room only.',
      'Weeknight Rate': 'The price of a Sunday to Thursday night.', 'Weekend Rate': 'The price of a Friday or Saturday night.',
    } },
    Rooms: { about: 'Each room in the house, one record each.', fields: {
      Number: 'The room\'s number on the door. Each number is used once.', Name: 'The room\'s name, as many guest houses give them.',
      'Room Type': 'The type of room, which sets its rates.', Floor: 'The floor it is on.', Outlook: 'What the window looks onto: Sea view, Garden or Town.',
      Beds: 'The beds in the room.', 'Dogs Allowed': 'Yes for a room where guests may bring a dog.',
      Housekeeping: 'How the room stands today: Vacant clean, Inspected (checked and ready for an arrival), Occupied, Due out, Vacant dirty or Out of order.',
      Notes: 'Anything about the room: a step down to the shower, a work in hand.',
    } },
    'Booking Sources': { about: 'Where bookings come from: your own website and e-mail, the phone, guests at the door, and each booking site you sell rooms through.', fields: {
      Name: 'The source\'s name. Each name is used once.', Code: 'A short code, used in the commission statement numbers.',
      Kind: 'Direct, or Booking site.', 'Commission Rate': 'The commission the site takes as a fraction of the room charge: 0.15 is fifteen percent. Nought for direct bookings.',
      'Account Ref': 'Your account or property number with the site (invented here).', Contact: 'The person you deal with.',
      'E-mail': 'Where statements go.', Phone: 'Their phone number.', Address: 'The street address statements are sent to.',
      City: 'The town or city.', Postcode: 'The postcode.', Country: 'The country.', Notes: 'Terms agreed: how they pay out, cancellation rules on their rates.',
    } },
    Guests: { about: 'Everyone who has booked a room, one record each. A guest who comes back is the same record with another booking.', fields: {
      'First Name': 'Their first name.', 'Last Name': 'Their family name.', Address: 'Their home street address.', City: 'Their home town or city.',
      Postcode: 'Their postcode.', Country: 'The country they live in.', Nationality: 'Their nationality, as a guest register records it.',
      'Passport Number': 'The passport number, kept only for guests from countries where it must be recorded and empty for the rest. The sample\'s numbers are invented (PP-).',
      'Passport Issued By': 'The country that issued the passport, kept with the number.', Phone: 'Their phone number.', 'E-mail': 'Their e-mail address.',
      'Car Registration': 'The car they usually come in. The sample\'s registrations all carry 00, which is never issued.',
      'Mailing List': 'Yes when they have asked to hear of offers and the new season\'s rates.', Notes: 'Preferences: a favourite room, an allergy, a dog\'s name.',
    } },
    Extras: { about: 'What can be charged to a room besides the night itself, with its price including VAT.', fields: {
      Name: 'The item as it appears on the bill. Each name is used once.', Kind: 'Breakfast, Dinner, Bar, Parking, Late check-out, Packed lunch or Dog.',
      Price: 'The price of one, VAT included.', Per: 'What one is: per person, per night, each or per stay.',
      'Tax Rate': 'The VAT rate as a fraction: 0.2 is twenty percent.', Available: 'Yes while it is offered.',
    } },
    'Commission Statements': { about: 'One statement for each booking site and month: the stays it sent that ended that month, and the commission owed on them.', fields: {
      Number: 'The statement\'s number, in date order, a new run each year: COM-2026-0012.', Source: 'The booking site.',
      Month: 'The month the stays ended in.', 'Statement Date': 'The day it was made, the first of the month after.',
      Stays: 'How many stays it lists.', 'Room Takings': 'What those stays paid for their rooms.', Commission: 'The commission owed on them.',
      'Paid On': 'The day you paid the site; empty until then.',
    } },
    Bookings: { about: 'Each booking of a room for a run of nights, from the day it is taken to the day the guest leaves, or cancels.', fields: {
      Number: 'The booking\'s number, given in the order bookings are taken, a new run each year: BK-2026-0042.', Guest: 'The guest who booked.',
      Room: 'The room booked.', Source: 'Where the booking came from.', 'Site Reference': 'The booking site\'s own reference, for its bookings (invented here).',
      'Booked On': 'The day the booking was taken.', Arrival: 'The day the guest arrives.', Departure: 'The day they leave.',
      Nights: 'How many nights.', Adults: 'How many adults.', Children: 'How many children.', 'Rate Plan': 'Bed and breakfast, or Room only.',
      'Room Total': 'The room charge for all the nights, VAT included, from the tariff.', 'Commission Rate': 'The site\'s commission rate when the booking was taken.',
      Commission: 'The commission owed to the site on the stay; nought for direct bookings and for cancellations.',
      'Deposit Due': 'The deposit asked for: the first night, for a direct booking taken two days or more ahead.',
      'Deposit Paid': 'The deposit held now: empty until paid, back to nought when refunded.',
      Status: 'Confirmed, Checked in, Checked out, Cancelled or No-show.', 'Expected Arrival': 'The time the guest expects to arrive.',
      'Checked In': 'When they arrived.', 'Checked Out': 'When they left.', 'Cancelled On': 'The day they cancelled.',
      'Car Registration': 'The car for this stay, for the car park and the register.', Statement: 'The commission statement the stay is on, for a booking site\'s stays once the month is done.',
      Requests: 'What the guest asked for: a ground-floor room, a cot, an early breakfast.',
    } },
    'Room Nights': { about: 'The diary: one record for every room and every night, fifteen months back and six months ahead, as a booking chart would show it.', fields: {
      Room: 'The room.', Night: 'The night, by the date it begins.', Weekday: 'The day of the week, Mon to Sun, so weekends can be picked out.',
      Status: 'Booked, Free, No-show (held for a guest who did not come) or Out of order.', Booking: 'The booking that has the room that night.',
      Rate: 'The room charge for the night, VAT included.',
    } },
    Bills: { about: 'The bill given to a guest when they leave, which is also the VAT invoice. A late cancellation or a no-show gets a bill too.', fields: {
      Number: 'The bill\'s number, in date order, a new run each year: INV-2026-0042.', Booking: 'The stay billed.', 'Bill Date': 'The day it was made.',
      Total: 'What the stay came to, VAT included: the charges on it added up.', VAT: 'The VAT within the total, each line\'s VAT rounded and then added up.',
      Deposit: 'The deposit already paid, taken off the total.', Balance: 'What was left to pay on leaving.', 'Paid By': 'How the balance was paid.',
    } },
    Charges: { about: 'Everything charged to a room: each night, and each extra on the day it was taken.', fields: {
      Booking: 'The stay it is charged to.', Bill: 'The bill it went on; empty while the guest is still in the house.',
      Date: 'The day of the charge: the night for a room charge, the morning for breakfast.', Kind: 'Accommodation, Breakfast, Dinner, Bar, Parking, Dog, Late check-out, Packed lunch, Cancellation or No-show.',
      Extra: 'The item from the Extras list; empty for the room itself and for cancellation charges.', Description: 'The words on the bill.',
      Quantity: 'How many.', 'Unit Price': 'The price of one, VAT included.', 'Tax Rate': 'The VAT rate as a fraction; nought for cancellation and no-show charges in the sample.',
    } },
    Payments: { about: 'Money taken from guests: deposits, the balance on leaving, and the charge for a guest who did not come.', fields: {
      Number: 'The receipt number, in date order, a new run each year: RCP-2026-0042.', Booking: 'The booking it was paid against.', Date: 'The day it was paid.',
      Kind: 'Deposit, Balance or No-show charge.', Method: 'Card, Cash or Bank transfer.', Amount: 'How much was paid.',
      'Refunded On': 'The day a deposit was paid back, when the guest cancelled in time.', Refund: 'How much was paid back.',
    } },
  },
  queries: {
    'Arrivals in the next 7 days': 'Who is coming in the next week, day by day and room by room, with their source, expected time, requests and what they will owe. Look at it every morning.',
    'Departures in the next 7 days': 'Who leaves in the next week, with the room charge still to pay after the deposit.',
    'Rooms free tonight and this weekend': 'The rooms still free tonight, and on the coming Friday and Saturday nights, for the phone call that asks "have you a room?".',
    'Occupancy by month': 'For each month from a year ago to two months ahead: room nights there were to sell (rooms out of order left out), how many were sold, the occupancy, the average rate for a night sold, the room revenue for each room night there was, and the room revenue. Months ahead show what is on the books so far.',
    'Revenue by source': 'Room takings for stays begun in the last twelve months, by where they were booked, with the commission each site took and what is left after it.',
    'Deposits outstanding': 'Bookings to come whose deposit is not yet paid, with the guest\'s phone and e-mail, soonest arrival first: the list for chasing.',
    'Returning guests': 'Guests who have stayed more than once, with their first and last stay and what they spent on rooms, the most loyal first.',
    'Housekeeping list for today': 'Every room in use today and what it needs: departures to strip and clean first, then stayovers, then arrivals to make ready.',
    'Room status': 'Every room with its type and its housekeeping state now.',
    'Guest register': 'The register of guests for the last twelve months: arrival, name, nationality, passport where it is kept, home address, room, departure and car.',
    'Guests in the house': 'Guests staying now, with what has been charged to their room so far and the deposit held, so a bill can be ready when they leave.',
    'Extras sold by month': 'What the extras took each month over the last year, item by item.',
    'Cancellations and no-shows': 'Cancelled bookings and guests who did not come, from a year ago to six months ahead, with the notice given and the deposit kept.',
    'Commission by month': 'Every commission statement of the last year, by date, with what it listed, what was owed and when it was paid.',
    Tariff: 'The rate card: each room type, season and plan with its weeknight and weekend rate.',
  },
  reports: {
    'Arrivals list': 'The coming week\'s arrivals on a page, a group for each day, for the front desk.',
    'Departures list': 'The coming week\'s departures, a group for each day, with the balances to take.',
    'Housekeeping sheet': 'Today\'s list for whoever cleans the rooms: room, task, guest and requests.',
    'Monthly occupancy': 'Occupancy, average rate and room revenue for each month, with the totals.',
    'Takings by source': 'Room takings over the last year by source, commission and what is left, with the totals.',
    'Deposits to chase': 'The deposits not yet paid, with how to reach each guest.',
    'Guest register': 'The register of guests month by month, as a guest house keeps it for inspection.',
    'Tariff card': 'The room rates, a group for each room type, to print or copy to the website.',
    'Commission by site': 'Each booking site\'s statements over the last year, with the total commission paid to it.',
  },
  documents: {
    'Booking confirmation': 'The letter sent when a booking is taken: the dates, the room, the rate, the deposit received and the balance, check-in and check-out times, the cancellation terms and how to find the house. Print it from the Bookings form for the booking on screen; the template prints every booking when run alone, including those from booking sites.',
    'Guest bill': 'The bill and VAT invoice given on leaving: each night and each extra as a line with the VAT within it, the total, the deposit taken off and the balance paid. Numbered INV-2026-0042.',
    Receipt: 'A till-roll receipt for each payment: a deposit, a balance or a no-show charge, with the refund if a deposit was paid back.',
    'Registration card': 'The card a guest fills in and signs on arrival: name, nationality, passport where it must be recorded and blank otherwise, home address, car registration and signature. Print it from the Bookings form for the day\'s arrivals; run alone it prints every booking.',
    'Commission statement': 'Sent to each booking site at the start of the month: the stays it sent that ended last month, with their references, and the commission owed. Numbered COM-2026-0012.',
  },
  tutorial: [
    { title: 'Take a booking on the phone', steps: [
      'Open the query Rooms free tonight and this weekend, or the Room Nights table filtered on the dates, to see which rooms are free. The Tariff query gives the rate for the season, the night of the week and the plan.',
      'Open the Guests form. Look the caller up by name; if they have not stayed before, press New and type their name, address, phone and e-mail.',
      'Open the Bookings form and press New. Type the next Number (BK- and the year and a count), choose the Guest, the Room and Direct: phone in Source, and type Booked On, Arrival, Departure, Nights, Adults and Rate Plan.',
      'Add up the nights from the tariff into Room Total, and type the first night\'s rate in Deposit Due. Leave Status at Confirmed.',
      'In the Room Nights table, set each night of the stay for that room to Booked, choose the booking in Booking and type the night\'s rate in Rate.',
      'Open the print template Booking confirmation from Print templates, with the booking on screen in the Bookings form, and press Print or Save as PDF to post or e-mail it.',
    ] },
    { title: 'Take a deposit and give a receipt', steps: [
      'Open the Payments form and press New. Type the next Number (RCP- and the year and a count), choose the Booking, and type the Date, Deposit in Kind, the Method and the Amount.',
      'Open the Bookings form for the same booking and type the amount in Deposit Paid.',
      'Open the print template Receipt with the payment on screen and press Print. It prints on a till roll, or Save as PDF to e-mail it.',
      'Each morning, open the query Deposits outstanding (or print the report Deposits to chase) and ring or write to anyone who has not paid.',
    ] },
    { title: 'Start the day', steps: [
      'Print the report Housekeeping sheet for whoever cleans the rooms: departures first, then stayovers, then arrivals.',
      'Open the query Arrivals in the next 7 days, or print the report Arrivals list, to see who comes today, at what time and with what requests.',
      'Print the Registration card for each of today\'s arrivals from the Bookings form, so they are ready on the desk.',
      'Open the Rooms form and set Housekeeping for each room as the cleaning is done: Vacant clean, Inspected, Occupied, Due out or Out of order. The query Room status shows them all.',
    ] },
    { title: 'Check a guest in', steps: [
      'Hand the guest their Registration card to check and sign. Keep the signed cards for at least a year.',
      'In the Bookings form, set Status to Checked in, type the time in Checked In and the car in Car Registration. If they did not pay a deposit, nothing more is owed until they leave.',
      'In the Rooms form, set the room\'s Housekeeping to Occupied.',
    ] },
    { title: 'Charge nights and extras to the room', steps: [
      'Each night, open the Charges form and press New for each room in use: choose the Booking, type the Date, Accommodation in Kind, the room and plan in Description, 1 in Quantity, the night\'s rate in Unit Price and 0.2 in Tax Rate.',
      'When a guest takes something more (breakfast on a room only rate, an evening meal, drinks, parking, a late check-out), press New in the Charges form, choose the Booking and the item in Extra, and type its Description, Quantity and Unit Price from the Extras table.',
      'Leave Bill empty until the guest leaves. The query Guests in the house shows what each room has run up so far.',
    ] },
    { title: 'Check a guest out and give the bill', steps: [
      'Open the query Departures in the next 7 days to see who leaves today, and the query Guests in the house for what each room owes.',
      'Open the Bills form and press New. Type the next Number (INV- and the year and a count), choose the Booking and type the Bill Date.',
      'In the Charges table, choose the new bill in Bill on every charge of the stay. Add them up into Total, add up the VAT within them into VAT, type the deposit in Deposit and what is left in Balance.',
      'Open the print template Guest bill with the bill on screen and press Print. Take the balance, enter it in the Payments form with Balance in Kind, type the method in Paid By on the bill, and print a Receipt if the guest wants one.',
      'In the Bookings form set Status to Checked out and type the time in Checked Out; in the Rooms form set Housekeeping to Vacant dirty until the room is cleaned.',
    ] },
    { title: 'A cancellation or a guest who does not come', steps: [
      'Find the booking in the Bookings form. Set Status to Cancelled and type the day in Cancelled On, or set it to No-show the morning after a guest did not arrive.',
      'In the Room Nights table set the booking\'s nights back to Free, empty Booking and Rate, so the rooms can be let again. For a no-show, set the first night to No-show.',
      'Cancelled seven days or more before arrival: refund the deposit, type the day in Refunded On and the amount in Refund on its payment in the Payments form, and set Deposit Paid on the booking to 0.',
      'Cancelled later, or a no-show: add a charge in the Charges form (Cancellation or No-show in Kind, 0 in Tax Rate) and make a bill for it as at check-out. For a booking-site guest who did not come, take the first night from the card the site passed on and enter it in Payments as a No-show charge.',
    ] },
    { title: 'Month end: commission statements and the numbers', steps: [
      'Open the Commission Statements form and press New for each booking site: type the next Number (COM- and the year and a count), choose the Source, type the Month and the Statement Date.',
      'In the Bookings table, find the site\'s stays that ended last month with Status Checked out, and choose the new statement in Statement on each. Type how many there are in Stays and add up their Room Total and Commission into the statement.',
      'Open the print template Commission statement with the statement on screen and press Print or Save as PDF to send to the site. When you have paid it, type the day in Paid On.',
      'Print the reports Monthly occupancy and Takings by source to see how the month went against the last, and which sources earn their commission.',
    ] },
  ],
  notes: [
    'The diary, Room Nights, is what the free-room and occupancy queries read. A booking entered in Bookings does not mark its nights by itself: set each night to Booked with its booking and rate, and back to Free if it is cancelled. The sample\'s diary runs six months ahead; add more nights as the year goes on (ten rooms, one record each a night).',
    'Every price in the sample includes VAT at twenty percent, as prices shown to guests do: rates, extras and the lines on a bill. The bill shows the VAT within each line, rounded to the penny, and the VAT on the bill is those added up.',
    'A bill is made when the guest leaves, so charges for a guest still in the house have an empty Bill. Room charges are posted night by night for nights already past; tonight\'s is posted tonight.',
    'Deposit Paid is the deposit held now: it goes back to nought when a deposit is refunded. Direct bookings taken two days or more ahead are asked for the first night; booking-site guests and walk-ins pay on leaving.',
    'Commission is worked out on the room charge alone, not on extras, at the rate held on the booking. Cancelled bookings and no-shows carry none in the sample. Booking sites\' own terms differ; set the rate and the rules to match yours.',
    'Cancellation and no-show charges carry no VAT in the sample. How they are taxed depends on your country and your terms: ask your accountant.',
    'The passport number is kept only for guests whose nationality calls for it and is left empty for the rest, as the registration card is printed. Every name, address, passport number, car registration and booking reference in the sample is invented.',
    'Print templates print one document for every record of their table. Print a single one from its form (the booking, bill or payment on screen), or choose the ones you want.',
  ],
  faq: [
    { q: 'How do I find a room for a caller?', a: 'Open the query Rooms free tonight and this weekend. For other dates, open the Room Nights table and filter it on Night and Status Free; the Tariff query gives the price.' },
    { q: 'Why is a room free in the diary when it shows a booking in Bookings?', a: 'The booking was entered without marking its nights in Room Nights, or it was cancelled and the nights were set back to Free. The diary is what the queries read, so keep the two in step.' },
    { q: 'How is occupancy worked out?', a: 'From Room Nights: the nights marked Booked as a share of all the room nights in the month, leaving out rooms out of order. The average rate is the room revenue divided by the nights sold.' },
    { q: 'How do I print a bill for an earlier stay?', a: 'Open the Bills form, find the bill by its number or the guest, and print the Guest bill template with that bill on screen.' },
    { q: 'What do I owe each booking site?', a: 'The query Commission by month lists every statement and whether it is paid. For stays not yet on a statement, the Commission field on each booking-site booking holds what will be owed.' },
    { q: 'Do I have to keep passport numbers?', a: 'That depends on the law where your guest house is. In the sample the field is filled only for guests whose nationality calls for it. Check what your own country asks, and keep no more than it asks.' },
    { q: 'How do I start using it for my own house?', a: 'Open it empty, enter your Business record, then Seasons, Room Types, Rates, Rooms, Booking Sources and Extras. Then make the Room Nights for each room for the months ahead, all Free, and take bookings from there.' },
  ],
  empty: 'Start with Business, your own name, address, VAT number and bank details, which every document prints. Then the lookups: Seasons, Room Types, Rates for each type, season and plan, your Rooms, Booking Sources (your own website, the phone, the door and each site you sell through, with its commission rate) and Extras. Then make the diary: one Room Nights record for each room and night for the months ahead, all Free. Guests and Bookings come in as people book; Charges, Bills and Payments as they stay; Commission Statements at the end of each month.',
  adapt: [
    'Register every guest, not only the one who booked: add a Party table pointing at Bookings with each person\'s name and nationality, and list it on the registration card as its lines.',
    'Sell packages: add a Packages table (two nights with dinner, a walking weekend) with a price, and point Bookings at it in place of the plan.',
    'Keep the car park: add a Parking Spaces table and a Space field on Charges for parking, and a query of spaces taken each night.',
    'Take payments in part: the Payments table already holds several payments for a booking; add a query that adds them up against each bill to show what is still owed by guests on account.',
    'Track the booking sites\' own payouts: where a site takes the guest\'s money and pays you less its commission, add a Payouts table pointing at Commission Statements.',
    'Price by the length of stay: add a Nights From field to Rates for a cheaper rate on stays of three nights or more.',
    'Let rooms by the week as holiday lets: change Rates to weekly prices and make a booking a run of Saturdays.',
  ],
};

/** The room types: name, description, sleeps, B and B rates [low weeknight, low weekend, mid..., high...], what room only takes off. */
const TYPES = [
  ['Single', 'Single bed and shower room, for one', 1, [62, 68, 72, 78, 82, 88], 10],
  ['Double', 'Double bed and en-suite shower room', 2, [86, 96, 102, 116, 122, 138], 20],
  ['Twin', 'Two single beds and en-suite shower room', 2, [86, 96, 102, 116, 122, 138], 20],
  ['Superior sea view', 'King-size bed, bath and shower, looking out to sea', 2, [112, 128, 132, 152, 158, 178], 20],
  ['Family', 'A double and two single beds, sea view, for up to four', 4, [118, 134, 138, 158, 168, 188], 20],
];
/** The rooms: number, name, type, floor, outlook, beds, dogs, notes. */
const ROOMS = [
  [1, 'Lighthouse', 'Superior sea view', 'First floor', 'Sea view', 'King-size bed', false, 'Bay window with two chairs.'],
  [2, 'Harbour', 'Superior sea view', 'First floor', 'Sea view', 'King-size bed', false, ''],
  [3, 'Dune', 'Double', 'First floor', 'Town', 'Double bed', false, ''],
  [4, 'Kittiwake', 'Twin', 'First floor', 'Garden', 'Two single beds', false, 'The beds can be pushed together.'],
  [5, 'Rockpool', 'Family', 'Second floor', 'Sea view', 'Double and two single beds', false, 'A cot can be added.'],
  [6, 'Samphire', 'Double', 'Second floor', 'Town', 'Double bed', false, ''],
  [7, 'Driftwood', 'Twin', 'Second floor', 'Garden', 'Two single beds', false, 'Redecorated last January.'],
  [8, 'Lookout', 'Superior sea view', 'Second floor', 'Sea view', 'King-size bed and window seat', false, ''],
  [9, 'Puffin', 'Single', 'Top floor', 'Town', 'Single bed', false, 'Sloping ceiling.'],
  [10, 'Garden', 'Double', 'Ground floor', 'Garden', 'Double bed', true, 'Step-free, its own door to the garden. Dogs welcome.'],
];
/** Where bookings come from: code, name, kind, commission, share of requests, lead times [[from, to, weight]...]. */
const SOURCES = [
  { code: 'WEB', name: 'Direct: website and e-mail', kind: 'Direct', rate: 0, weight: 22, leads: [[2, 6, 2], [7, 30, 3], [31, 90, 3], [91, 170, 2], [0, 1, 1]] },
  { code: 'PHN', name: 'Direct: phone', kind: 'Direct', rate: 0, weight: 14, leads: [[2, 6, 2], [7, 30, 3], [31, 120, 4], [121, 175, 2], [0, 1, 1]] },
  { code: 'WLK', name: 'Walk-in', kind: 'Direct', rate: 0, weight: 3, leads: [[0, 0, 1]] },
  { code: 'TWR', name: 'Tidewater Rooms', kind: 'Booking site', rate: 0.15, weight: 30, leads: [[0, 3, 3], [4, 14, 3], [15, 45, 3], [46, 120, 2]], place: { Address: '40 Canal Walk', City: 'Amsterdam', Postcode: '1016 KL', Country: 'Netherlands' } },
  { code: 'RTR', name: 'Roomtrail', kind: 'Booking site', rate: 0.17, weight: 20, leads: [[0, 3, 3], [4, 14, 4], [15, 45, 2], [46, 100, 1]], place: { Address: '7 Harbour Street', City: 'Dublin', Postcode: 'D02 X418', Country: 'Ireland' } },
  { code: 'NLB', name: 'Northlight Breaks', kind: 'Booking site', rate: 0.12, weight: 11, leads: [[3, 14, 2], [15, 60, 3], [61, 150, 2]], place: { Address: '118 Station Road', City: 'Manchester', Postcode: 'M3 4HJ', Country: 'United Kingdom' } },
];
/** What can be charged to a room: name, kind, price with VAT, per. */
const EXTRAS = [
  ['Breakfast', 'Breakfast', 12.5, 'per person'], ['Evening meal, two courses', 'Dinner', 26, 'per person'],
  ['Glass of wine', 'Bar', 6.5, 'each'], ['Bottle of house wine', 'Bar', 22, 'each'], ['Bottled ale', 'Bar', 5.2, 'each'],
  ['Soft drink', 'Bar', 2.8, 'each'], ['Parking space', 'Parking', 8, 'per night'], ['Late check-out until 1 pm', 'Late check-out', 20, 'per stay'],
  ['Packed lunch', 'Packed lunch', 9.5, 'per person'], ['Dog', 'Dog', 15, 'per night'],
];
// How many want a room, by the month they arrive (January first) and the day of the week (Sunday first).
const MONTH_DEMAND = [0.35, 0.45, 0.55, 0.7, 0.8, 0.9, 1.15, 1.2, 0.85, 0.7, 0.45, 0.5];
const DAY_DEMAND = [0.55, 0.8, 0.7, 0.75, 0.95, 1.9, 1.0];
const SEASON_OF = [0, 0, 0, 1, 1, 1, 2, 2, 1, 1, 0, 0];
const WEEKDAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
// Guests' home countries: [index in PLACES, nationality, weight for direct, weight for booking sites, passport recorded].
const HOMES = [
  [0, 'British', 84, 62, false], [3, 'Irish', 3, 5, false], [6, 'German', 3, 8, true], [7, 'Dutch', 2, 7, true],
  [1, 'American', 3, 7, true], [2, 'Canadian', 2, 3, false], [8, 'Australian', 2, 4, false], [4, 'Pakistani', 0.6, 1.5, false],
  [5, 'Emirati', 0.2, 1.5, true], [9, 'Nigerian', 0.4, 1, false],
];
const REQUESTS = ['Ground floor if possible, please.', 'Celebrating our wedding anniversary.', 'Arriving by train.', 'Vegetarian breakfast, please.', 'Gluten-free breakfast, please.', 'A quiet room, please.', 'Early breakfast at 7.30 am on the last morning.', 'Feather-free pillows, please.'];

/** Fill a Gullhaven Guest House database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 450, ahead = 180 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };
  // The bill's VAT worked out as SQLite works it in the print template, so the two agree to the penny.
  const roundSql = base.db.prepare('SELECT ROUND(?, 2) AS v');
  const vatSql = base.db.prepare('SELECT ROUND(? - CAST(? AS REAL) / (1 + ?), 2) AS v');
  const lineAmount = (q, p) => roundSql.get(q * p).v;
  const vatWithin = (amount, rate) => (rate ? vatSql.get(amount, amount, rate).v : 0);

  add('Business', businessRecord({
    name: 'Gullhaven Guest House', domain: 'gullhaven.example', phone: '07700 900412',
    address: { Address: '12 The Esplanade', City: 'Saltby Bay', Postcode: 'SB7 2QT', Country: 'United Kingdom' },
    terms: 'The balance of your stay and anything charged to the room are paid on departure.',
    footer: 'Check-in 3 pm to 8 pm. Check-out by 10.30 am. Breakfast 8 am to 9.30 am, Sundays 9 am to 10 am. Thank you for staying with us.',
  }));

  // Seasons, room types, the tariff, the rooms.
  const seasonIds = [
    add('Seasons', { Name: 'Low', Months: 'January to March, November and December', 'Minimum Weekend Nights': 1, Notes: 'Christmas and New Year at mid-season rates by arrangement.' }),
    add('Seasons', { Name: 'Mid', Months: 'April to June, September and October', 'Minimum Weekend Nights': 2, Notes: 'Bank holiday weekends at high-season rates.' }),
    add('Seasons', { Name: 'High', Months: 'July and August', 'Minimum Weekend Nights': 2, Notes: 'School summer holidays.' }),
  ];
  const minWeekend = [1, 2, 2];
  const types = {};
  for (const [name, desc, sleeps, rates, off] of TYPES) {
    const id = add('Room Types', { Name: name, Description: desc, Sleeps: sleeps, 'Breakfast Allowance': off });
    types[name] = { id, name, sleeps, rates, off };
    seasonIds.forEach((season, s) => {
      add('Rates', { 'Room Type': id, Season: season, Plan: 'Bed and breakfast', 'Weeknight Rate': rates[s * 2], 'Weekend Rate': rates[s * 2 + 1] });
      add('Rates', { 'Room Type': id, Season: season, Plan: 'Room only', 'Weeknight Rate': rates[s * 2] - off, 'Weekend Rate': rates[s * 2 + 1] - off });
    });
  }
  const rooms = ROOMS.map(([number, name, type, floor, outlook, beds, dogs, notes], i) => ({
    i, number, name, type: types[type], dogs,
    id: add('Rooms', { Number: number, Name: name, 'Room Type': types[type].id, Floor: floor, Outlook: outlook, Beds: beds, 'Dogs Allowed': dogs, Notes: notes || null }),
  }));
  const ofType = (t) => rooms.filter((x) => x.type.name === t);

  const sources = SOURCES.map((s) => {
    const slug = s.name.toLowerCase().replace(/[^a-z]+/g, '');
    const site = s.kind === 'Booking site';
    const id = add('Booking Sources', {
      Name: s.name, Code: s.code, Kind: s.kind, 'Commission Rate': s.rate,
      'Account Ref': site ? `GH-${s.code}-0412` : null, Contact: site ? `${r.pick(FIRST)} ${r.pick(LAST)}` : null,
      'E-mail': site ? `partners@${slug}.example` : null, Phone: site ? phoneFor(r, s.place.Country) : null, ...(site ? s.place : {}),
      Notes: site ? 'Guests pay us on departure. Commission invoiced monthly on stays completed; none on cancellations.' : null,
    });
    return { ...s, id, site };
  });
  const extras = {};
  for (const [name, kind, price, per] of EXTRAS) extras[name] = { name, kind, price, id: add('Extras', { Name: name, Kind: kind, Price: price, Per: per, 'Tax Rate': 0.2, Available: true }) };

  // The diary as the allocation sees it: for each room, each night from the start to six months ahead.
  let start = -days;
  while (cal.at(start).getDate() !== 1) start++;
  const N = ahead - start + 1;
  const OUT = 'out of order';
  const occ = rooms.map(() => new Array(N).fill(null));
  // Driftwood was redecorated for ten nights in January; Kittiwake's shower is being retiled tonight and tomorrow.
  let jan = start;
  while (!(cal.at(jan).getMonth() === 0 && cal.at(jan).getDate() === 8)) jan++;
  for (let k = 0; k < 10; k++) occ[6][jan + k - start] = OUT;
  for (let k = 0; k < 2; k++) occ[3][k - start] = OUT;

  const rateFor = (type, night, plan) => {
    const s = SEASON_OF[cal.at(night).getMonth()];
    const wd = cal.weekday(night);
    const base0 = type.rates[s * 2 + (wd === 5 || wd === 6 ? 1 : 0)];
    return plan === 'Room only' ? base0 - type.off : base0;
  };
  const leadOf = (s) => { const [lo, hi] = r.weighted(s.leads.map((l) => [l, l[2]])); return r.int(lo, hi); };

  // Who asks for a room, arrival day by arrival day: more in summer and for Friday nights.
  const requests = [];
  for (let a = start; a <= ahead; a++) {
    const month = cal.at(a).getMonth();
    const wd = cal.weekday(a);
    const want = 5.6 * MONTH_DEMAND[month] * DAY_DEMAND[wd];
    const n = Math.floor(want * (0.7 + 0.6 * r.next()) + r.next());
    const season = SEASON_OF[month];
    for (let k = 0; k < n; k++) {
      const source = r.weighted(sources.map((s) => [s, s.weight]));
      if (source.code === 'WLK' && a > 0) continue;
      const lead = leadOf(source);
      const bookedOn = a - lead;
      if (bookedOn > 0) continue;
      let nights = season === 2 ? r.weighted([[2, 3], [3, 3], [4, 2], [5, 1], [7, 1]])
        : wd === 5 ? r.weighted([[2, 7], [3, 2], [1, season === 0 ? 2 : 0]])
          : wd === 6 ? r.weighted([[1, 2], [2, 3]]) : r.weighted([[1, 4], [2, 4], [3, 3], [4, 1], [5, 1]]);
      if (source.code === 'WLK') nights = Math.min(nights, 2);
      // A stay that takes in a Saturday night is at least the season's weekend minimum.
      if (nights < minWeekend[season] && Array.from({ length: nights }, (_, j) => cal.weekday(a + j)).includes(6)) nights = minWeekend[season];
      const type = r.weighted([['Single', 8], ['Double', 34], ['Twin', 14], ['Superior sea view', 30], ['Family', season === 2 ? 22 : 9]]);
      const adults = type === 'Single' ? 1 : type === 'Family' ? 2 : r.chance(0.92) ? 2 : 1;
      const children = type === 'Family' ? r.int(1, 2) : 0;
      const plan = source.site ? (r.chance(0.55) ? 'Room only' : 'Bed and breakfast') : r.chance(0.07) ? 'Room only' : 'Bed and breakfast';
      requests.push({ a, nights, bookedOn, source, want: type, adults, children, plan, order: r.next() });
    }
  }
  requests.sort((x, y) => x.bookedOn - y.bookedOn || x.order - y.order);

  // Rooms go to whoever books first. A cancellation gives its nights back on the day it is made.
  // Over the coming week one room is always kept back, as a late cancellation would leave one.
  const ALT = { Single: ['Double', 'Twin'], Double: ['Superior sea view', 'Twin'], Twin: ['Double'], 'Superior sea view': ['Double'], Family: [] };
  const taken = (d) => rooms.reduce((n, x) => n + (occ[x.i][d - start] !== null ? 1 : 0), 0);
  const free = (room, a, nights) => {
    for (let d = a; d < a + nights; d++) {
      if (d < start || d > ahead || occ[room.i][d - start] !== null) return false;
      if (d >= 0 && d <= 6 && taken(d) >= rooms.length - 1) return false;
    }
    return true;
  };
  const bookings = [];
  const cancels = [];
  const release = (b) => { for (let d = b.a; d < b.a + b.nights; d++) if (occ[b.room.i][d - start] === b) occ[b.room.i][d - start] = null; };
  for (const q of requests) {
    for (let i = cancels.length - 1; i >= 0; i--) if (cancels[i].cancelled <= q.bookedOn) { release(cancels[i]); cancels.splice(i, 1); }
    if (q.a < start) continue;
    let room = null;
    for (const t of [q.want, ...ALT[q.want]]) {
      const open = ofType(t).filter((x) => free(x, q.a, q.nights));
      if (open.length) { room = r.pick(open); break; }
    }
    if (!room) continue;
    const b = { ...q, room, type: room.type };
    for (let d = q.a; d < q.a + q.nights; d++) occ[room.i][d - start] = b;
    // Some cancel between booking and arriving; one decided on a day still to come has not happened yet.
    if (q.a - q.bookedOn >= 1 && q.source.code !== 'WLK' && r.chance(q.source.site ? 0.13 : 0.05)) {
      const c = q.bookedOn + r.int(0, q.a - q.bookedOn - 1);
      if (c <= 0) { b.cancelled = c; cancels.push(b); }
    }
    if (b.cancelled === undefined && q.a <= -1 && q.source.code !== 'WLK' && r.chance(q.source.site ? 0.03 : 0.015)) b.noShow = true;
    bookings.push(b);
  }
  for (const b of cancels) release(b);
  // Whatever the day, someone stayed last night and someone arrives this week: a quiet spell gets a walk-in or a phone booking.
  const forced = (a, nights, bookedOn, code) => {
    const source = sources.find((s) => s.code === code);
    const room = rooms.find((x) => x.type.name !== 'Family' && free(x, a, nights));
    if (!room) return;
    const b = { a, nights, bookedOn, source, want: room.type.name, room, type: room.type, adults: 2, children: 0, plan: 'Bed and breakfast', order: 0 };
    for (let d = a; d < a + nights; d++) occ[room.i][d - start] = b;
    bookings.push(b);
  };
  const staying = (from, to) => bookings.some((b) => b.cancelled === undefined && !b.noShow && b.a <= to && b.a + b.nights - 1 >= from);
  if (!staying(-1, -1)) forced(-1, 2, -1, 'WLK');
  if (!bookings.some((b) => b.cancelled === undefined && b.a >= 0 && b.a <= 7)) forced(2, 2, -3, 'PHN');

  // Guests: new ones, and some who have stayed before and come back.
  const guests = [];
  const usedNames = new Set();
  const letter = () => r.pick('ABCDEFGHJKLMNOPRSTUVWXY'.split(''));
  const newGuest = (site) => {
    const [pi, nationality, , , passport] = r.weighted(HOMES.map((h) => [h, site ? h[3] : h[2]]));
    let first;
    let last;
    for (let k = 0; k < 30; k++) { first = r.pick(FIRST); last = r.pick(LAST); if (!usedNames.has(`${first} ${last}`)) break; }
    usedNames.add(`${first} ${last}`);
    const a = address(r, [PLACES[pi]]);
    const car = r.chance(pi === 0 ? 0.75 : 0.2) ? `${letter()}${letter()}00 ${letter()}${letter()}${letter()}` : null;
    const g = { lastDep: -Infinity, car };
    g.id = add('Guests', {
      'First Name': first, 'Last Name': last, ...a, Nationality: nationality,
      'Passport Number': passport ? `PP-${String(r.int(100000, 999999))}` : null, 'Passport Issued By': passport ? a.Country : null,
      Phone: phoneFor(r, a.Country), 'E-mail': `${first}.${last}@mail.example`.toLowerCase(),
      'Car Registration': car, 'Mailing List': r.chance(0.4),
      Notes: r.chance(0.06) ? r.pick(['Likes a sea-view room.', 'Allergic to nuts.', 'Prefers the ground floor.', 'Comes for the walking festival.']) : null,
    });
    guests.push(g);
    return g;
  };

  // Each booking's charges, deposits, bill and payments, worked out in the order the bookings were taken.
  bookings.sort((x, y) => x.bookedOn - y.bookedOn || x.a - y.a);
  const payments = [];
  const bills = [];
  let firstOfMonth = 0;
  while (cal.at(firstOfMonth).getDate() !== 1) firstOfMonth--;
  for (const b of bookings) {
    const dep = b.a + b.nights;
    b.dep = dep;
    const site = b.source.site;
    // A guest who has stayed before and has no stay to come.
    const back = guests.filter((g) => g.lastDep < b.bookedOn);
    b.guest = back.length && r.chance(site ? 0.06 : b.source.code === 'WLK' ? 0.05 : 0.35) ? r.pick(back) : newGuest(site);
    if (b.cancelled === undefined) b.guest.lastDep = Math.max(b.guest.lastDep, dep);
    b.rates = Array.from({ length: b.nights }, (_, k) => rateFor(b.type, b.a + k, b.plan));
    b.total = money(b.rates.reduce((n, x) => n + x, 0));
    b.status = b.cancelled !== undefined ? 'Cancelled' : b.noShow ? 'No-show' : dep <= -1 ? 'Checked out' : b.a <= -1 ? 'Checked in' : 'Confirmed';
    b.commission = site && ['Confirmed', 'Checked in', 'Checked out'].includes(b.status) ? money(b.total * b.source.rate) : 0;
    b.expected = r.pick(['15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30', '19:00', '20:00', '21:30']);
    b.car = b.guest.car && r.chance(0.9) ? b.guest.car : null;
    b.dog = b.room.dogs && r.chance(0.4);
    b.requests = b.dog ? 'Bringing a dog.' : b.children && r.chance(0.3) ? 'A cot, please.' : r.chance(0.12) ? r.pick(REQUESTS) : null;

    // A deposit of the first night for a direct booking taken two days or more ahead.
    b.depositDue = b.source.kind === 'Direct' && b.source.code !== 'WLK' && b.a - b.bookedOn >= 2 ? b.rates[0] : 0;
    b.depositHeld = 0;
    if (b.depositDue) {
      // Most pay the day they book; some take a few days; a few never send it and pay on leaving.
      const payDay = b.bookedOn + r.weighted([[0, 5], [1, 2], [2, 2], [3, 1.5], [5, 1], [8, 0.7], [999, 0.8]]);
      const paid = payDay <= 0 && payDay < b.a && (b.cancelled === undefined || payDay <= b.cancelled);
      if (paid) {
        const pay = { b, date: payDay, kind: 'Deposit', method: r.weighted([['Card', 7], ['Bank transfer', 3]]), amount: b.depositDue };
        payments.push(pay);
        // Cancelled seven days or more ahead: the deposit goes back. Later: it is kept.
        if (b.cancelled !== undefined && b.a - b.cancelled >= 7) { pay.refundedOn = b.cancelled; pay.refund = b.depositDue; } else b.depositHeld = b.depositDue;
      }
    }

    // What is charged to the room: nights already past and the extras taken with them.
    b.charges = [];
    const charge = (date, kind, description, quantity, price, extra = null, rate = 0.2) => b.charges.push({ date, kind, description, quantity, price, extra, rate });
    if (b.status === 'Checked out' || b.status === 'Checked in') {
      const parking = b.car && r.chance(0.45);
      for (let k = 0; k < b.nights; k++) {
        const night = b.a + k;
        if (night > -1) break;
        charge(night, 'Accommodation', `Room ${b.room.number} ${b.room.name}, ${b.plan.toLowerCase()}`, 1, b.rates[k]);
        if (parking) charge(night, 'Parking', 'Parking space', 1, extras['Parking space'].price, extras['Parking space']);
        if (b.dog) charge(night, 'Dog', 'Dog', 1, extras.Dog.price, extras.Dog);
        if (r.chance(SEASON_OF[cal.at(night).getMonth()] === 0 ? 0.16 : 0.1)) charge(night, 'Dinner', 'Evening meal, two courses', b.adults, extras['Evening meal, two courses'].price, extras['Evening meal, two courses']);
        if (r.chance(0.3)) {
          for (const name of r.some(['Glass of wine', 'Bottle of house wine', 'Bottled ale', 'Soft drink'], r.int(1, 2))) charge(night, 'Bar', name, name === 'Bottle of house wine' ? 1 : r.int(1, 3), extras[name].price, extras[name]);
        }
        const morning = night + 1;
        if (morning <= -1) {
          if (b.plan === 'Room only') {
            const eaters = Array.from({ length: b.adults }, () => r.chance(0.55)).filter(Boolean).length;
            if (eaters) charge(morning, 'Breakfast', 'Breakfast', eaters, extras.Breakfast.price, extras.Breakfast);
          }
          if (r.chance(0.04)) charge(morning, 'Packed lunch', 'Packed lunch', b.adults, extras['Packed lunch'].price, extras['Packed lunch']);
        }
      }
      b.late = b.status === 'Checked out' && r.chance(0.07);
      if (b.late) charge(dep, 'Late check-out', 'Late check-out until 1 pm', 1, extras['Late check-out until 1 pm'].price, extras['Late check-out until 1 pm']);
    }
    // A no-show is charged the first night; a late cancellation keeps its deposit. Neither carries VAT here.
    let billDate = null;
    if (b.status === 'Checked out') billDate = dep;
    if (b.status === 'No-show') { billDate = b.a + 1; charge(b.a, 'No-show', 'Did not arrive: first night charged', 1, b.rates[0], null, 0); }
    if (b.status === 'Cancelled' && b.depositHeld) { billDate = b.cancelled; charge(b.cancelled, 'Cancellation', 'Cancelled within 7 days of arrival: deposit kept', 1, b.depositHeld, null, 0); }
    if (billDate !== null) {
      let total = 0;
      let vat = 0;
      for (const c of b.charges) { c.amount = lineAmount(c.quantity, c.price); total += c.amount; vat += vatWithin(c.amount, c.rate); }
      total = money(total);
      vat = money(vat);
      const balance = money(total - b.depositHeld);
      const method = balance > 0 ? (b.status === 'No-show' ? 'Card' : r.weighted([['Card', 75], ['Cash', 12], ['Bank transfer', 13]])) : 'Deposit';
      const bill = { b, date: billDate, total, vat, deposit: b.depositHeld, balance, method };
      bills.push(bill);
      b.bill = bill;
      if (balance > 0) payments.push({ b, date: billDate, kind: b.status === 'No-show' ? 'No-show charge' : 'Balance', method, amount: balance });
    }
  }

  // A deposit or two is always waiting to come in: if none is, the newest one asked for has not arrived yet.
  const waiting = (b) => b.status === 'Confirmed' && b.depositDue > 0;
  if (!bookings.some((b) => waiting(b) && !b.depositHeld)) {
    const b = bookings.filter(waiting).at(-1);
    if (b) { b.depositHeld = 0; for (let i = payments.length - 1; i >= 0; i--) if (payments[i].b === b) payments.splice(i, 1); }
  }

  // Commission statements: each booking site, each month gone, for the stays that ended in it.
  const statements = new Map();
  for (const b of bookings) {
    if (!b.source.site || b.status !== 'Checked out' || b.dep >= firstOfMonth) continue;
    const at = cal.at(b.dep);
    const key = `${b.source.code}|${at.getFullYear()}-${at.getMonth()}`;
    if (!statements.has(key)) {
      let next = b.dep + 1;
      while (cal.at(next).getDate() !== 1) next++;
      statements.set(key, { source: b.source, month: `${MONTHS[at.getMonth()]} ${at.getFullYear()}`, date: next, list: [] });
    }
    statements.get(key).list.push(b);
  }
  const yearly = {};
  const numbered = (prefix, d) => { const y = cal.at(d).getFullYear(); const k = `${prefix}${y}`; yearly[k] = (yearly[k] || 0) + 1; return `${prefix}-${y}-${String(yearly[k]).padStart(4, '0')}`; };
  for (const s of [...statements.values()].sort((x, y) => x.date - y.date || x.source.id - y.source.id)) {
    const paid = s.date + r.int(8, 16);
    s.id = add('Commission Statements', {
      Number: numbered('COM', s.date), Source: s.source.id, Month: s.month, 'Statement Date': cal.day(s.date), Stays: s.list.length,
      'Room Takings': money(s.list.reduce((n, b) => n + b.total, 0)), Commission: money(s.list.reduce((n, b) => n + b.commission, 0)),
      'Paid On': paid <= 0 ? cal.day(paid) : null,
    });
    for (const b of s.list) b.statement = s.id;
  }

  // The bookings, numbered in the order they were taken.
  for (const b of bookings) {
    const inAt = b.status === 'Checked out' || b.status === 'Checked in' ? cal.time(b.a, r.int(15, 20), r.pick([0, 10, 20, 30, 40, 50])) : null;
    const outAt = b.status === 'Checked out' ? (b.late ? cal.time(b.dep, 12, r.pick([30, 45, 55])) : cal.time(b.dep, r.int(8, 10), r.pick([0, 15, 30, 45]))) : null;
    b.id = add('Bookings', {
      Number: numbered('BK', b.bookedOn), Guest: b.guest.id, Room: b.room.id, Source: b.source.id,
      'Site Reference': b.source.site ? `${b.source.code}${r.int(1000000, 9999999)}` : null, 'Booked On': cal.day(b.bookedOn),
      Arrival: cal.day(b.a), Departure: cal.day(b.dep), Nights: b.nights, Adults: b.adults, Children: b.children, 'Rate Plan': b.plan,
      'Room Total': b.total, 'Commission Rate': b.source.rate, Commission: b.commission, 'Deposit Due': b.depositDue, 'Deposit Paid': b.depositHeld,
      Status: b.status, 'Expected Arrival': b.expected, 'Checked In': inAt, 'Checked Out': outAt,
      'Cancelled On': b.cancelled !== undefined ? cal.day(b.cancelled) : null, 'Car Registration': b.car, Statement: b.statement ?? null, Requests: b.requests,
    });
  }

  // The diary, night by night and room by room.
  for (let d = start; d <= ahead; d++) {
    for (const room of rooms) {
      const v = occ[room.i][d - start];
      const row = { Room: room.id, Night: cal.day(d), Weekday: WEEKDAYS[cal.weekday(d)], Status: 'Free', Booking: null, Rate: null };
      if (v === OUT) row.Status = 'Out of order';
      else if (v && v.noShow) { if (d === v.a) Object.assign(row, { Status: 'No-show', Booking: v.id }); }
      else if (v) Object.assign(row, { Status: 'Booked', Booking: v.id, Rate: v.rates[d - v.a] });
      add('Room Nights', row);
    }
  }

  // Bills in date order, then every charge (on its bill once there is one), then the payments in date order.
  bills.sort((x, y) => x.date - y.date || x.b.id - y.b.id);
  for (const bill of bills) {
    bill.id = add('Bills', { Number: numbered('INV', bill.date), Booking: bill.b.id, 'Bill Date': cal.day(bill.date), Total: bill.total, VAT: bill.vat, Deposit: bill.deposit, Balance: bill.balance, 'Paid By': bill.method });
  }
  for (const b of bookings) {
    for (const c of b.charges) {
      add('Charges', { Booking: b.id, Bill: b.bill ? b.bill.id : null, Date: cal.day(c.date), Kind: c.kind, Extra: c.extra ? c.extra.id : null, Description: c.description, Quantity: c.quantity, 'Unit Price': c.price, 'Tax Rate': c.rate });
    }
  }
  payments.sort((x, y) => x.date - y.date || x.b.id - y.b.id);
  for (const p of payments) {
    add('Payments', { Number: numbered('RCP', p.date), Booking: p.b.id, Date: cal.day(p.date), Kind: p.kind, Method: p.method, Amount: p.amount, 'Refunded On': p.refundedOn !== undefined ? cal.day(p.refundedOn) : null, Refund: p.refund || 0 });
  }

  // How each room stands this morning.
  for (const room of rooms) {
    const tonight = occ[room.i][-start];
    const lastNight = occ[room.i][-1 - start];
    const stay = (v) => v && v !== OUT && !v.noShow;
    let state = 'Vacant clean';
    if (tonight === OUT) state = 'Out of order';
    else if (stay(lastNight) && lastNight.dep === 0) state = 'Due out';
    else if (stay(lastNight)) state = 'Occupied';
    else if (stay(tonight) && tonight.a === 0) state = 'Inspected';
    base.update('Rooms', room.id, { Housekeeping: state });
  }
  return counts;
}
