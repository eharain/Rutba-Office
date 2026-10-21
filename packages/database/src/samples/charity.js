// The charity: Rowanmere Community Trust, an invented small community charity
// in Glasgow with a club side. Members pay a yearly fee, supporters give once
// or every month, volunteers give their time, the trust runs events with
// tickets, and a few grants come from invented funders. One People table
// holds members, donors, guests and volunteers alike, with what each has
// agreed to be contacted by and whether they have made the tax declaration
// that lets the trust claim tax back on their gifts (as the UK's Gift Aid
// scheme does).
//
// Its records come from eighteen months run a day at a time: memberships run
// for a year and most are renewed, some late and some never; appeals open
// and close and draw gifts while they are open; standing orders arrive on the
// same day each month; each quarter the claimable gifts go into a tax claim
// that is paid some weeks later; events are held every other Saturday with
// bookings taken weeks ahead; volunteers turn up on their usual day. Every
// payment in is also written in the Income table, as a treasurer's cash book
// keeps it, so income can be counted by source and month in one query.

import { FIRST, LAST, money, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const NAME = '[First Name] & " " & [Last Name]';
// The share of a gift the tax claim brings back: a quarter, as the UK's scheme pays at the basic rate.
const CLAIM = 'Round([Amount] * 0.25, 2)';
const AHEAD = (field) => `DateDiff("d", Date(), [${field}])`;
const AGO = (field) => `DateDiff("d", [${field}], Date())`;

export const spec = {
  // Where a value below nought is right: a grant report already late.
  negative: { 'Grant reports due': ['Days Left'] },
  tables: [
    { name: 'People', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
      { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' }, { name: 'Country', type: 'text', default: 'United Kingdom' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Joined', type: 'date' },
      { name: 'Contact by Post', type: 'yesno' }, { name: 'Contact by E-mail', type: 'yesno' }, { name: 'Contact by Phone', type: 'yesno' },
      { name: 'Tax Declaration', type: 'yesno' }, { name: 'Declaration Date', type: 'date' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Membership Types', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Fee', type: 'currency' }, { name: 'Months', type: 'integer', default: 12 },
      { name: 'Description', type: 'text' },
    ] },
    { name: 'Memberships', fields: [
      { name: 'Person', ...ref('People'), required: true }, { name: 'Membership Type', ...ref('Membership Types'), required: true },
      { name: 'Starts', type: 'date' }, { name: 'Ends', type: 'date' }, { name: 'Amount', type: 'currency' }, { name: 'Paid On', type: 'date' },
      { name: 'Method', type: 'text' }, { name: 'Renewed', type: 'yesno' },
    ] },
    { name: 'Campaigns', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Opens', type: 'date' }, { name: 'Closes', type: 'date' },
      { name: 'Target', type: 'currency' }, { name: 'Description', type: 'text' },
    ] },
    { name: 'Tax Claims', fields: [
      { name: 'Reference', type: 'text', unique: true }, { name: 'Gifts Up To', type: 'date' }, { name: 'Submitted On', type: 'date' },
      { name: 'Amount Claimed', type: 'currency' }, { name: 'Received On', type: 'date' }, { name: 'Amount Received', type: 'currency' },
    ] },
    { name: 'Donations', fields: [
      { name: 'Person', ...ref('People'), required: true }, { name: 'Campaign', ...ref('Campaigns'), required: true },
      { name: 'Gift Date', type: 'date' }, { name: 'Amount', type: 'currency' }, { name: 'Method', type: 'text' },
      { name: 'Regular', type: 'yesno' }, { name: 'Claimable', type: 'yesno' }, { name: 'Tax Claim', ...ref('Tax Claims') },
    ] },
    { name: 'Events', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Event Date', type: 'date' }, { name: 'Start Time', type: 'text' },
      { name: 'Venue', type: 'text' }, { name: 'Capacity', type: 'integer' }, { name: 'Ticket Price', type: 'currency', default: 0 },
      { name: 'Costs', type: 'currency', default: 0 }, { name: 'Status', type: 'text', default: 'Planned' },
    ] },
    { name: 'Event Bookings', fields: [
      { name: 'Event', ...ref('Events'), required: true }, { name: 'Person', ...ref('People'), required: true },
      { name: 'Booked On', type: 'date' }, { name: 'Tickets', type: 'integer', default: 1 }, { name: 'Amount Paid', type: 'currency', default: 0 },
      { name: 'Cancelled', type: 'yesno' }, { name: 'Attended', type: 'yesno' },
    ] },
    { name: 'Volunteer Roles', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Usual Day', type: 'text' }, { name: 'Hour Value', type: 'currency' },
      { name: 'Description', type: 'text' },
    ] },
    { name: 'Volunteer Shifts', fields: [
      { name: 'Person', ...ref('People'), required: true }, { name: 'Role', ...ref('Volunteer Roles'), required: true },
      { name: 'Shift Date', type: 'date' }, { name: 'Start Time', type: 'text' }, { name: 'Hours', type: 'number' },
    ] },
    { name: 'Funders', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Kind', type: 'text' }, { name: 'Contact Name', type: 'text' },
      { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' },
    ] },
    { name: 'Grants', fields: [
      { name: 'Funder', ...ref('Funders'), required: true }, { name: 'Purpose', type: 'text' }, { name: 'Applied On', type: 'date' },
      { name: 'Amount Asked', type: 'currency' }, { name: 'Status', type: 'text', default: 'Applied' }, { name: 'Decided On', type: 'date' },
      { name: 'Amount Awarded', type: 'currency' }, { name: 'Paid On', type: 'date' }, { name: 'Period Ends', type: 'date' },
      { name: 'Report Due', type: 'date' }, { name: 'Report Sent', type: 'date' },
    ] },
    { name: 'Income', fields: [
      { name: 'Received On', type: 'date' }, { name: 'Source', type: 'text' }, { name: 'Description', type: 'text' },
      { name: 'Amount', type: 'currency' }, { name: 'Method', type: 'text' }, { name: 'Reference', type: 'text' },
    ] },
  ],
  queries: {
    // A membership is current when it has started and has not ended; a life membership has no end.
    'Current members': { tables: ['Memberships', 'People', 'Membership Types'], columns: [
      { table: 'Membership Types', field: 'Name', as: 'Type', sort: 'asc' }, { expr: NAME, as: 'Member', sort: 'asc' },
      { table: 'Memberships', field: 'Starts' }, { table: 'Memberships', field: 'Ends' },
      { table: 'People', field: 'E-mail' }, { table: 'People', field: 'Phone' },
      { expr: AGO('Starts'), as: 'Days In', criteria: ['>= 0', '>= 0'], show: false },
      { expr: AHEAD('Ends'), as: 'Days Left', criteria: ['>= 0'], show: false },
      { table: 'Memberships', field: 'Ends', as: 'No End', criteria: ['', 'Is Null'], show: false },
    ] },
    'Members by type': { tables: ['Memberships', 'Membership Types'], columns: [
      { table: 'Membership Types', field: 'Name', as: 'Type', total: 'group', sort: 'asc' },
      { table: 'Memberships', field: 'ID', as: 'Members', total: 'count' }, { table: 'Memberships', field: 'Amount', as: 'Fees Paid', total: 'sum' },
      { expr: AGO('Starts'), as: 'Days In', total: 'where', criteria: ['>= 0', '>= 0'] },
      { expr: AHEAD('Ends'), as: 'Days Left', total: 'where', criteria: ['>= 0'] },
      { table: 'Memberships', field: 'Ends', total: 'where', criteria: ['', 'Is Null'] },
    ] },
    'Renewals due': { tables: ['Memberships', 'People', 'Membership Types'], columns: [
      { table: 'Memberships', field: 'Ends', sort: 'asc' }, { expr: AHEAD('Ends'), as: 'Days Left', criteria: ['Between 0 And 30'] },
      { expr: NAME, as: 'Member' }, { table: 'Membership Types', field: 'Name', as: 'Type' }, { table: 'Membership Types', field: 'Fee', as: 'Renewal Fee' },
      { table: 'People', field: 'Address' }, { table: 'People', field: 'City' }, { table: 'People', field: 'Postcode' }, { table: 'People', field: 'E-mail' },
      { table: 'People', field: 'Contact by E-mail' }, { table: 'Memberships', field: 'Renewed', criteria: ['No'], show: false },
    ] },
    'Lapsed members': { tables: ['Memberships', 'People', 'Membership Types'], columns: [
      { expr: NAME, as: 'Member' }, { table: 'Membership Types', field: 'Name', as: 'Type' }, { table: 'Memberships', field: 'Ends', as: 'Ended', sort: 'desc' },
      { expr: AGO('Ends'), as: 'Days Lapsed', criteria: ['Between 1 And 365'] }, { table: 'People', field: 'E-mail' }, { table: 'People', field: 'Phone' },
      { table: 'People', field: 'Contact by E-mail' }, { table: 'Memberships', field: 'Renewed', criteria: ['No'], show: false },
    ] },
    'Appeals against target': { tables: ['Campaigns', 'Donations'], columns: [
      { table: 'Campaigns', field: 'Name', as: 'Campaign', total: 'group' }, { table: 'Campaigns', field: 'Opens', total: 'group', sort: 'asc' },
      { table: 'Campaigns', field: 'Closes', total: 'group' }, { table: 'Campaigns', field: 'Target', as: 'Target', total: 'max' },
      { table: 'Donations', field: 'Amount', as: 'Given', total: 'sum' }, { table: 'Donations', field: 'ID', as: 'Gifts', total: 'count' },
    ] },
    'Gifts to claim': { tables: ['Donations', 'People'], columns: [
      { table: 'Donations', field: 'Gift Date', sort: 'asc' }, { table: 'People', field: 'First Name' }, { table: 'People', field: 'Last Name' },
      { table: 'People', field: 'Address' }, { table: 'People', field: 'Postcode' }, { table: 'Donations', field: 'Amount' },
      { expr: CLAIM, as: 'Claim Value' }, { table: 'Donations', field: 'Claimable', criteria: ['Yes'], show: false },
      { table: 'Donations', field: 'Tax Claim', criteria: ['Is Null'], show: false },
    ] },
    'Tax claims made': { tables: ['Tax Claims', 'Donations'], columns: [
      { table: 'Tax Claims', field: 'Reference', total: 'group', sort: 'asc' }, { table: 'Tax Claims', field: 'Gifts Up To', total: 'group' },
      { table: 'Tax Claims', field: 'Submitted On', total: 'group' }, { table: 'Donations', field: 'ID', as: 'Gifts', total: 'count' },
      { table: 'Donations', field: 'Amount', as: 'Gifts Total', total: 'sum' }, { table: 'Tax Claims', field: 'Amount Claimed', as: 'Claimed', total: 'max' },
      { table: 'Tax Claims', field: 'Received On', total: 'group' }, { table: 'Tax Claims', field: 'Amount Received', as: 'Received', total: 'max' },
    ] },
    'Top supporters': { tables: ['People', 'Donations'], columns: [
      { expr: NAME, as: 'Supporter', total: 'group' }, { table: 'People', field: 'City', total: 'group' },
      { table: 'Donations', field: 'Amount', as: 'Given', total: 'sum', sort: 'desc' }, { table: 'Donations', field: 'ID', as: 'Gifts', total: 'count' },
      { table: 'People', field: 'Tax Declaration', total: 'group' },
      { expr: AGO('Gift Date'), as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Regular givers': { tables: ['People', 'Donations'], columns: [
      { expr: NAME, as: 'Supporter', total: 'group', sort: 'asc' }, { table: 'Donations', field: 'Amount', as: 'Monthly Gift', total: 'max' },
      { table: 'Donations', field: 'Gift Date', as: 'Last Gift', total: 'max' }, { table: 'People', field: 'Tax Declaration', total: 'group' },
      { table: 'Donations', field: 'Regular', total: 'where', criteria: ['Yes'] },
      { expr: AGO('Gift Date'), as: 'Days Ago', total: 'where', criteria: ['Between 0 And 45'] },
    ] },
    'Volunteer hours by month': { tables: ['Volunteer Shifts', 'Volunteer Roles'], columns: [
      { expr: 'Format([Shift Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { table: 'Volunteer Roles', field: 'Name', as: 'Role', total: 'group', sort: 'asc' },
      { table: 'Volunteer Shifts', field: 'Hours', as: 'Hours', total: 'sum' }, { expr: 'Round([Hours] * [Hour Value], 2)', as: 'Value', total: 'sum' },
      { expr: AGO('Shift Date'), as: 'Days Ago', total: 'where', criteria: ['>= 0'] },
    ] },
    'Volunteers': { tables: ['Volunteer Shifts', 'People'], columns: [
      { expr: NAME, as: 'Volunteer', total: 'group' }, { table: 'People', field: 'Phone', total: 'group' },
      { table: 'Volunteer Shifts', field: 'ID', as: 'Shifts', total: 'count' }, { table: 'Volunteer Shifts', field: 'Hours', as: 'Hours', total: 'sum', sort: 'desc' },
      { table: 'Volunteer Shifts', field: 'Shift Date', as: 'Last Shift', total: 'max' },
      { expr: AGO('Shift Date'), as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Volunteer rota': { tables: ['Volunteer Shifts', 'Volunteer Roles', 'People'], columns: [
      { table: 'Volunteer Shifts', field: 'Shift Date', sort: 'asc' }, { table: 'Volunteer Shifts', field: 'Start Time', sort: 'asc' },
      { table: 'Volunteer Roles', field: 'Name', as: 'Role', sort: 'asc' }, { expr: NAME, as: 'Volunteer' }, { table: 'People', field: 'Phone' },
      { table: 'Volunteer Shifts', field: 'Hours' }, { expr: AHEAD('Shift Date'), as: 'Days Ahead', criteria: ['Between 0 And 14'], show: false },
    ] },
    'Upcoming events': { tables: ['Events', 'Event Bookings'], columns: [
      { table: 'Events', field: 'Event Date', total: 'group', sort: 'asc' }, { table: 'Events', field: 'Name', as: 'Event', total: 'group' },
      { table: 'Events', field: 'Start Time', total: 'group' }, { table: 'Events', field: 'Venue', total: 'group' },
      { table: 'Events', field: 'Capacity', as: 'Capacity', total: 'max' }, { table: 'Event Bookings', field: 'Tickets', as: 'Booked', total: 'sum' },
      { table: 'Event Bookings', field: 'Amount Paid', as: 'Taken', total: 'sum' },
      { expr: AHEAD('Event Date'), as: 'Days Ahead', total: 'where', criteria: ['Between 0 And 60'] },
      { table: 'Event Bookings', field: 'Cancelled', total: 'where', criteria: ['No'] },
    ] },
    'Event attendance': { tables: ['Events', 'Event Bookings'], columns: [
      { table: 'Events', field: 'Event Date', total: 'group', sort: 'desc' }, { table: 'Events', field: 'Name', as: 'Event', total: 'group' },
      { table: 'Events', field: 'Capacity', as: 'Capacity', total: 'max' }, { table: 'Event Bookings', field: 'Tickets', as: 'Booked', total: 'sum' },
      { expr: 'IIf([Attended] = 1, [Tickets], 0)', as: 'Came', total: 'sum' },
      { table: 'Event Bookings', field: 'Amount Paid', as: 'Ticket Income', total: 'sum' }, { table: 'Events', field: 'Costs', as: 'Costs', total: 'max' },
      { table: 'Events', field: 'Status', total: 'where', criteria: ['Held'] }, { table: 'Event Bookings', field: 'Cancelled', total: 'where', criteria: ['No'] },
    ] },
    'Guests booked': { tables: ['Event Bookings', 'Events', 'People'], columns: [
      { expr: '[Events].[Name] & ", " & Format([Event Date], "yyyy-mm-dd")', as: 'Event' }, { table: 'Events', field: 'Event Date', sort: 'asc' },
      { table: 'Events', field: 'Name', as: 'Event Name', sort: 'asc', show: false }, { table: 'People', field: 'Last Name', sort: 'asc', show: false },
      { expr: NAME, as: 'Guest' }, { table: 'Event Bookings', field: 'Tickets' }, { table: 'Event Bookings', field: 'Amount Paid' },
      { table: 'People', field: 'Phone' }, { expr: AHEAD('Event Date'), as: 'Days Ahead', criteria: ['Between 0 And 60'], show: false },
      { table: 'Event Bookings', field: 'Cancelled', criteria: ['No'], show: false },
    ] },
    'Grant reports due': { tables: ['Grants', 'Funders'], columns: [
      { table: 'Funders', field: 'Name', as: 'Funder' }, { table: 'Grants', field: 'Purpose' }, { table: 'Grants', field: 'Amount Awarded' },
      { table: 'Grants', field: 'Period Ends' }, { table: 'Grants', field: 'Report Due', sort: 'asc' }, { expr: AHEAD('Report Due'), as: 'Days Left' },
      { table: 'Funders', field: 'Contact Name' }, { table: 'Grants', field: 'Status', criteria: ['Awarded'], show: false },
      { table: 'Grants', field: 'Report Sent', criteria: ['Is Null'], show: false },
    ] },
    'Grants by funder': { tables: ['Grants', 'Funders'], columns: [
      { table: 'Funders', field: 'Name', as: 'Funder', total: 'group', sort: 'asc' }, { table: 'Funders', field: 'Kind', total: 'group' },
      { table: 'Grants', field: 'ID', as: 'Applications', total: 'count' }, { table: 'Grants', field: 'Amount Asked', as: 'Asked', total: 'sum' },
      { expr: 'Nz([Amount Awarded])', as: 'Awarded', total: 'sum' },
    ] },
    'Income by source by month': { tables: ['Income'], columns: [
      { expr: 'Format([Received On], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Income', field: 'Source', total: 'group', sort: 'asc' },
      { table: 'Income', field: 'Amount', as: 'Received', total: 'sum' }, { table: 'Income', field: 'ID', as: 'Payments', total: 'count' },
    ] },
    'Income by source': { tables: ['Income'], columns: [
      { table: 'Income', field: 'Source', total: 'group' }, { table: 'Income', field: 'Amount', as: 'Received', total: 'sum', sort: 'desc' },
      { table: 'Income', field: 'ID', as: 'Payments', total: 'count' },
      { expr: AGO('Received On'), as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Mailing list': { tables: ['People'], columns: [
      { table: 'People', field: 'Last Name', sort: 'asc' }, { table: 'People', field: 'First Name', sort: 'asc' },
      { table: 'People', field: 'Address' }, { table: 'People', field: 'City' }, { table: 'People', field: 'Postcode' },
      { table: 'People', field: 'Contact by Post', criteria: ['Yes'], show: false },
    ] },
  },
  reports: [
    { name: 'Membership list', query: 'Current members', group: 'Type', columns: ['Member', { field: 'Starts', type: 'date' }, { field: 'Ends', type: 'date' }, 'E-mail', 'Phone'], totals: [{ field: 'Member', fn: 'count' }] },
    { name: 'Renewal reminders', query: 'Renewals due', sort: 'Ends', columns: ['Member', 'Address', 'City', 'Postcode', 'Type', { field: 'Ends', type: 'date' }, { field: 'Renewal Fee', type: 'currency' }], totals: [{ field: 'Renewal Fee', fn: 'sum' }] },
    { name: 'Donations by campaign', query: 'Appeals against target', columns: ['Campaign', { field: 'Opens', type: 'date' }, { field: 'Closes', type: 'date' }, { field: 'Target', type: 'currency' }, { field: 'Gifts', type: 'integer' }, { field: 'Given', type: 'currency' }], totals: [{ field: 'Given', fn: 'sum' }, { field: 'Gifts', fn: 'sum' }] },
    { name: 'Claim schedule', title: 'Tax claim schedule', query: 'Gifts to claim', sort: 'Gift Date', columns: ['First Name', 'Last Name', 'Address', 'Postcode', { field: 'Gift Date', type: 'date' }, { field: 'Amount', type: 'currency' }, { field: 'Claim Value', type: 'currency' }], totals: [{ field: 'Amount', fn: 'sum' }, { field: 'Claim Value', fn: 'sum' }] },
    { name: 'Volunteer hours', query: 'Volunteer hours by month', group: 'Month', columns: ['Role', { field: 'Hours', type: 'number' }, { field: 'Value', type: 'currency' }], totals: [{ field: 'Hours', fn: 'sum' }, { field: 'Value', fn: 'sum' }] },
    { name: 'Event guest list', query: 'Guests booked', group: 'Event', columns: ['Guest', { field: 'Tickets', type: 'integer' }, { field: 'Amount Paid', type: 'currency' }, 'Phone'], totals: [{ field: 'Tickets', fn: 'sum' }, { field: 'Amount Paid', fn: 'sum' }] },
    { name: 'Income by month', query: 'Income by source by month', group: 'Month', columns: ['Source', { field: 'Payments', type: 'integer' }, { field: 'Received', type: 'currency' }], totals: [{ field: 'Received', fn: 'sum' }] },
    { name: 'Grant reports', title: 'Grant reports to write', query: 'Grant reports due', columns: ['Funder', 'Purpose', { field: 'Amount Awarded', type: 'currency' }, { field: 'Report Due', type: 'date' }, { field: 'Days Left', type: 'integer' }], totals: [{ field: 'Amount Awarded', fn: 'sum' }] },
  ],
};

export const about = {
  name: 'Rowanmere Community Trust',
  kind: 'Charity or club',
  line: 'Members, supporters, gifts, tax claims, events, volunteers and grants for a small community charity or club.',
  summary: [
    'For a small charity or club run by a committee and a part-time treasurer: one database for the people who belong, give, come to events and help out, and for the money that comes in.',
    'Rowanmere Community Trust is an invented community charity in Glasgow. It has about two hundred members who pay a yearly fee, regular supporters who give by standing order, appeals for things like a minibus or a new garden path, a hall where it holds a quiz, a ceilidh or a coffee morning most fortnights, a few dozen volunteers, and grants from a handful of invented trusts and the local council.',
    'The records cover the last eighteen months and run into the coming weeks: memberships due for renewal, events with bookings already taken, the volunteer rota and grant reports still to write. Gifts from people who have made a tax declaration are gathered into a claim each quarter, as a UK charity claims Gift Aid; the current quarter\'s gifts are waiting for the next claim. Every payment in is also written once in the Income table, so the treasurer can see income by source and by month.',
  ].join('\n\n'),
  suits: [
    'A sports or social club with yearly subscriptions',
    'A local history, gardening or arts society',
    'A parent teacher association',
    'A place of worship with regular givers and volunteers',
    'A community centre or village hall committee',
    'A small animal rescue or food bank run by volunteers',
  ],
  tables: {
    'People': { about: 'One record per person the trust knows: members, donors, event guests and volunteers are all here, once each. What someone is to the trust comes from their other records (a membership, a gift, a shift).', fields: {
      'First Name': 'The given name, used on letters and lists.',
      'Last Name': 'The family name; lists are sorted by it.',
      'Address': 'House number and street, for letters and for a tax claim, which needs the house and postcode of each giver.',
      'City': 'The town or city.',
      'Postcode': 'The postcode or ZIP code.',
      'Country': 'The country, United Kingdom unless you enter another.',
      'Phone': 'A phone number, used for the volunteer rota and event changes.',
      'E-mail': 'An e-mail address, left empty for people who have none.',
      'Joined': 'The day the person first came onto the records, by joining, giving or booking.',
      'Contact by Post': 'Yes when the person has agreed to letters and the newsletter by post. The mailing list uses it.',
      'Contact by E-mail': 'Yes when the person has agreed to e-mail from the trust.',
      'Contact by Phone': 'Yes when the person has agreed to be phoned.',
      'Tax Declaration': 'Yes when the person has declared that they pay enough tax for the trust to claim tax back on their gifts (in the UK, a Gift Aid declaration). Many countries have a scheme like it; where yours has none, ignore it.',
      'Declaration Date': 'The day the declaration was made. Gifts from that day on are claimable.',
      'Notes': 'Anything worth remembering, such as a gift made in memory of someone or a request for large print.',
    } },
    'Membership Types': { about: 'The kinds of membership on offer and what each costs. Enter these first.', fields: {
      'Name': 'The name of the membership, such as Adult or Family.',
      'Fee': 'What the membership costs now. A renewal is charged at this fee.',
      'Months': 'How long it lasts, usually 12. Nought means it never ends, as for a life membership.',
      'Description': 'Who the membership is for.',
    } },
    'Memberships': { about: 'One record for each year of membership paid. A renewal is a new record starting the day after the last one ends, so a member\'s history stays.', fields: {
      'Person': 'The member.',
      'Membership Type': 'The kind of membership.',
      'Starts': 'The first day of the membership.',
      'Ends': 'The last day of the membership; empty for a life membership.',
      'Amount': 'What was paid for it.',
      'Paid On': 'The day the fee was paid.',
      'Method': 'How it was paid: standing order, card, cash or bank transfer.',
      'Renewed': 'Yes once a later membership has been taken out, so this one no longer shows as due or lapsed.',
    } },
    'Campaigns': { about: 'Appeals for a stated purpose, each with a target, and the General Fund for gifts given to no appeal in particular. Every gift belongs to one.', fields: {
      'Name': 'The name of the appeal or fund.',
      'Opens': 'The day the appeal started.',
      'Closes': 'The day it ended or ends; empty for a fund that stays open.',
      'Target': 'How much the appeal set out to raise; empty for the General Fund.',
      'Description': 'What the money is for.',
    } },
    'Tax Claims': { about: 'Each claim made to the tax office for tax back on gifts, made each quarter. The gifts in a claim point at it.', fields: {
      'Reference': 'The trust\'s own reference for the claim, such as TC-004.',
      'Gifts Up To': 'The last gift date the claim covers.',
      'Submitted On': 'The day the claim was sent.',
      'Amount Claimed': 'The total claimed: a quarter of the gifts in it, the rate the UK scheme pays.',
      'Received On': 'The day the money arrived; empty while it is awaited.',
      'Amount Received': 'What was paid.',
    } },
    'Donations': { about: 'One record per gift, whether one-off or a monthly standing order.', fields: {
      'Person': 'Who gave it.',
      'Campaign': 'The appeal or fund it was given to.',
      'Gift Date': 'The day it was received.',
      'Amount': 'How much was given.',
      'Method': 'How: standing order, card, online, cash, cheque or bank transfer.',
      'Regular': 'Yes for a gift that comes every month.',
      'Claimable': 'Yes when the giver had a tax declaration in place on the day of the gift.',
      'Tax Claim': 'The claim the gift went into; empty until it is claimed.',
    } },
    'Events': { about: 'Everything the trust holds: quiz nights, ceilidhs, coffee mornings, talks, the summer fair and the AGM, past and planned.', fields: {
      'Name': 'What the event is.',
      'Event Date': 'The day it is held.',
      'Start Time': 'When it starts, as 19:30.',
      'Venue': 'Where it is held.',
      'Capacity': 'How many people the venue takes.',
      'Ticket Price': 'The price of one ticket; nought for a free event.',
      'Costs': 'What the event cost to put on: hall hire, a band, food.',
      'Status': 'Planned, Held or Cancelled.',
    } },
    'Event Bookings': { about: 'One record per booking: a person and the tickets they took for one event.', fields: {
      'Event': 'The event booked.',
      'Person': 'Who booked.',
      'Booked On': 'The day the booking was made and paid.',
      'Tickets': 'How many tickets.',
      'Amount Paid': 'What was paid for them; nought once a booking is cancelled and refunded.',
      'Cancelled': 'Yes when the booking was cancelled, or the event was.',
      'Attended': 'Yes when the party came on the day, ticked at the door.',
    } },
    'Volunteer Roles': { about: 'The jobs volunteers do, with a value per hour so their time can be counted in money for funders and annual reports.', fields: {
      'Name': 'The role.',
      'Usual Day': 'The day of the week the role usually runs.',
      'Hour Value': 'What an hour of the role is worth if it were paid for, at a rate you choose.',
      'Description': 'What the role involves.',
    } },
    'Volunteer Shifts': { about: 'One record for each turn a volunteer does, past and on the rota ahead.', fields: {
      'Person': 'The volunteer.',
      'Role': 'The role done.',
      'Shift Date': 'The day of the shift.',
      'Start Time': 'When it starts.',
      'Hours': 'How long it lasts, in hours.',
    } },
    'Funders': { about: 'The trusts, foundations and councils the trust applies to for grants.', fields: {
      'Name': 'The funder.',
      'Kind': 'Charitable trust, community foundation, council or other.',
      'Contact Name': 'The grants officer the trust deals with.',
      'E-mail': 'Their e-mail address.',
      'Phone': 'Their phone number.',
    } },
    'Grants': { about: 'Each application for a grant, from the day it was sent to the report the funder wants at the end.', fields: {
      'Funder': 'Who the application went to.',
      'Purpose': 'What the money is for.',
      'Applied On': 'The day it was sent.',
      'Amount Asked': 'How much was asked for.',
      'Status': 'Applied (waiting for a decision), Awarded, Declined, or Completed once the report is sent.',
      'Decided On': 'The day the funder said yes or no.',
      'Amount Awarded': 'How much was given; empty when declined.',
      'Paid On': 'The day the money arrived.',
      'Period Ends': 'The day the funded work should be finished.',
      'Report Due': 'The day the funder wants a report on how the money was spent.',
      'Report Sent': 'The day the report went; empty while it is still to write.',
    } },
    'Income': { about: 'The cash book: one line for every payment received, whatever it was for, as the treasurer banks it. Fees, gifts, tickets, grants, tax claims and collections are all here, so income can be counted in one place.', fields: {
      'Received On': 'The day the money came in.',
      'Source': 'What kind of income: Membership, Donations, Events, Collections, Grants or Tax claims.',
      'Description': 'What it was for, such as the appeal or event.',
      'Amount': 'How much.',
      'Method': 'How it was paid.',
      'Reference': 'The record it came from, such as Gift 123 or Booking 45, so it can be traced.',
    } },
  },
  queries: {
    'Current members': 'Everyone whose membership has started and not yet ended, and life members, by type. Use it to check who may vote at the AGM or use members\' rates.',
    'Members by type': 'How many current members there are of each type and what they paid. A quick count for the committee.',
    'Renewals due': 'Memberships ending in the next 30 days that have not been renewed, with addresses and the fee to ask for. Run it each week and send the reminders.',
    'Lapsed members': 'Members whose membership ended in the last year and who have not renewed. Worth a friendly call or e-mail; some simply forgot.',
    'Appeals against target': 'Each appeal and fund with its target, what has been given to it and how many gifts. Shows which appeals reached their target.',
    'Gifts to claim': 'Claimable gifts not yet put into a tax claim, with the giver\'s name, house and postcode and what the claim on each brings back. This is the next claim.',
    'Tax claims made': 'Each tax claim with the gifts in it, what was claimed and when the money came; a claim still awaited has no date received.',
    'Top supporters': 'The people who gave most in the last 12 months, with the number of gifts. Use it for thank-you letters.',
    'Regular givers': 'Supporters with a monthly gift that arrived in the last 45 days, and the amount. A giver missing from it may have stopped their standing order.',
    'Volunteer hours by month': 'Hours given each month in each role, and their value at the role\'s hourly rate. Funders and annual reports often ask for this.',
    'Volunteers': 'Each volunteer\'s shifts and hours in the last 12 months and their last shift, most hours first.',
    'Volunteer rota': 'Who is on which shift in the next 14 days, with a phone number, for the rota sheet on the noticeboard.',
    'Upcoming events': 'Events in the next 60 days with tickets booked so far against the capacity and the money taken.',
    'Event attendance': 'Each event held, with tickets booked, how many came and the ticket income against the costs.',
    'Guests booked': 'Every booking for events in the next 60 days, by event, for the door list.',
    'Grant reports due': 'Grants awarded whose report to the funder is not yet sent, soonest first; a negative Days Left means the report is late.',
    'Grants by funder': 'For each funder, how many applications were made, how much was asked and how much awarded.',
    'Income by source by month': 'All income each month by source: membership, donations, events, collections, grants and tax claims.',
    'Income by source': 'Income in the last 12 months by source, largest first, for the treasurer\'s report.',
    'Mailing list': 'Everyone who agreed to post, with their address, sorted by name, for the newsletter labels.',
  },
  reports: {
    'Membership list': 'The current members by type, with dates and contact details, and a count of each type. For the membership secretary.',
    'Renewal reminders': 'Members due to renew in the next 30 days with their address and fee, to write or print the reminders from.',
    'Donations by campaign': 'Each appeal and fund with its target, the number of gifts and the total given. For the trustees\' meeting.',
    'Claim schedule': 'The gifts for the next tax claim with each giver\'s name, house and postcode, date and amount, and the total to claim. Keep it with the claim.',
    'Volunteer hours': 'Volunteer hours and their value by month and role, with totals. For the annual report and grant reports.',
    'Event guest list': 'The guests booked for each event in the next 60 days with their tickets, for the door.',
    'Income by month': 'Income by source for each month with the month\'s total. For the treasurer.',
    'Grant reports': 'Grant reports still to write, with the amount and the date each is due.',
  },
  empty: 'Start with the lookup tables: Membership Types (your fees), Campaigns (at least a General Fund for gifts given to nothing in particular), Volunteer Roles and Funders. Then enter People as you meet them, with what they have agreed to be contacted by. After that, record Memberships, Donations, Events and their Bookings, Volunteer Shifts and Grants as they happen, and write each payment received in Income too. Make a Tax Claim when you send one and set each gift\'s Tax Claim to it.',
  adapt: [
    'If your country has no scheme for claiming tax on gifts, leave Tax Declaration empty and delete the Tax Claims table, the Gifts to claim query and the Claim schedule report.',
    'If your scheme pays a different share than a quarter, change the 0.25 in the Claim Value column of the Gifts to claim query.',
    'For a club with monthly subscriptions, add a Monthly membership type with Months set to 1, and the renewal queries work as before.',
    'Add a Spending table (date, kind, description, amount) to see income against costs each month.',
    'Add an Event field to Volunteer Shifts, pointing at Events, to see who stewarded each event.',
    'For a parent teacher association, rename Memberships to Families and add a field for the children\'s classes.',
    'Add a Gift in Memory Of field to Donations if people often give in memory of someone.',
    'Add a Safeguarding Check date to People, and a query of volunteers whose check is over three years old.',
  ],
};

/** The kinds of membership: [name, fee, months, description, weight among new members]. */
const TYPES = [
  ['Adult', 30, 12, 'One adult', 50],
  ['Concession', 18, 12, 'Over 65, student, or on benefits', 22],
  ['Family', 50, 12, 'Two adults and their children at one address', 18],
  ['Junior', 12, 12, 'Under 18, for the junior sports sessions', 8],
  ['Life', 300, 0, 'One adult, for life, paid once', 0],
];

/** Appeals as [name, opens, closes, target, description, pull]: offsets in days from today; pull is how many gifts a day it draws while open. */
const APPEALS = [
  ['Minibus Appeal', -520, -400, 9000, 'Towards a minibus for outings and the community transport service', 1.5],
  ['Hall Roof Appeal', -370, -280, 6000, 'Repairs to the leaking hall roof and gutters', 1.4],
  ['Hardship Fund', -240, -150, 5000, 'Small grants to local families facing a crisis', 1.3],
  ['Holiday Club Appeal', -170, -100, 4000, 'A free holiday club for local children in the school holidays', 1.0],
  ['Garden Path Appeal', -45, 45, 5000, 'A level path through the community garden for wheelchairs and buggies', 1.6],
];

/** Volunteer roles: [name, usual day (0 Sunday), start time, hour value, description, volunteers]. */
const ROLES = [
  ['Coffee morning team', 2, '09:30', 12.5, 'Setting up, serving and clearing the Tuesday coffee morning', 8],
  ['Gardening group', 4, '13:00', 12.5, 'Looking after the community garden on Thursday afternoons', 7],
  ['Junior sports coach', 6, '09:30', 16, 'Coaching the Saturday junior football and netball sessions', 6],
  ['Community driver', 3, '09:00', 14, 'Driving the minibus for shopping trips and appointments', 4],
  ['Office and admin', 1, '10:00', 13, 'Membership records, banking and the newsletter', 3],
  ['Event steward', null, '18:00', 12.5, 'Tickets, setting up and clearing at events', 0],
];

/** Events held in turn on Saturdays: [name, start, venue, capacity, price, costs]. */
const EVENTS = [
  ['Quiz Night', '19:30', 'Main Hall', 80, 5, 60],
  ['Ceilidh', '19:00', 'Main Hall', 120, 12, 420],
  ['Curry Night', '19:00', 'Main Hall', 70, 15, 520],
  ['Film Night', '19:00', 'Upper Room', 50, 4, 40],
  ['Local History Talk', '14:00', 'Upper Room', 50, 3, 30],
  ['Craft Fair', '10:00', 'Main Hall', 200, 1, 80],
  ['Charity Walk', '10:00', 'Riverside Park', 150, 10, 150],
  ['Games Afternoon', '14:00', 'Upper Room', 40, 2, 20],
];

const FUNDERS = [
  ['Kestrel Charitable Trust', 'Charitable trust'],
  ['Riverbend District Council', 'Council'],
  ['Hollybrook Community Foundation', 'Community foundation'],
  ['Amberly Family Trust', 'Charitable trust'],
  ['Granite Hill Sports Fund', 'Sports fund'],
  ['Willow Bank Arts Fund', 'Arts fund'],
];

/** Grants as [funder, purpose, applied, asked, decided, outcome, awarded, paid, period ends, report due, report sent]: offsets in days. */
const GRANTS = [
  [0, 'Minibus purchase', -540, 15000, -480, 'Awarded', 12000, -470, -120, -90, -96],
  [2, 'Kitchen refit', -400, 8000, -350, 'Declined'],
  [1, 'Junior sports coaching', -390, 5000, -350, 'Awarded', 4500, -340, 10, 40, null],
  [3, 'Holiday club', -260, 3500, -220, 'Awarded', 3000, -210, -30, -6, null],
  [4, 'Sports kit and equipment', -210, 2000, -180, 'Awarded', 1500, -172, 30, 62, null],
  [0, 'Running costs', -160, 6000, -120, 'Awarded', 5000, -112, 250, 280, null],
  [5, 'Concert series', -130, 2500, -95, 'Awarded', 2000, -85, 120, 150, null],
  [3, 'Garden tools', -300, 900, -270, 'Declined'],
  [2, 'Volunteer training', -50, 2400, null, 'Applied'],
  [1, 'Hall heating', -25, 7500, null, 'Applied'],
];

const STREETS = ['Thistle Street', 'Kelvin Way', 'Station Road', 'Burnside Avenue', 'Church Lane', 'Mill Road', 'Park Terrace', 'Hillhead Crescent', 'Main Street', 'Birch Grove', 'Clyde View', 'Rowan Drive', 'Elm Court', 'Victoria Road'];
// Towns around the trust and their postcode districts: [[town, districts], weight].
const TOWNS = [[['Glasgow', ['G11', 'G12', 'G13', 'G14', 'G20', 'G31', 'G41', 'G42', 'G44']], 10], [['Paisley', ['PA1', 'PA2', 'PA3']], 2], [['Rutherglen', ['G73']], 1.5], [['Clydebank', ['G81']], 1], [['Bearsden', ['G61']], 1]];
const NOTES = ['Prefers large print', 'Gives in memory of her late husband', 'Gives in memory of his mother', 'Happy to help at the summer fair', 'No calls after 6pm', 'Former trustee', 'Wheelchair user: ground floor seating', 'Bakes for the coffee morning'];

/** Fill a Rowanmere Community Trust database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 540 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };
  const zero = cal.at(0).getTime();
  const offsetOf = (date) => Math.round((date.getTime() - zero) / 86400000);
  /** The last day of a membership of `months` months starting on day `start`. */
  const endOf = (start, months) => { const a = cal.at(start); return offsetOf(new Date(a.getFullYear(), a.getMonth() + months, a.getDate() - 1)); };
  const income = (d, source, description, amount, method, reference) => {
    if (d < -days || d > 0 || !amount) return;
    add('Income', { 'Received On': cal.day(d), Source: source, Description: description, Amount: money(amount), Method: method, Reference: reference });
  };

  // People: invented names, each once, at invented addresses around Glasgow.
  const taken = new Set();
  const people = [];
  const newPerson = (joined, { declare = 0.5 } = {}) => {
    let first;
    let last;
    do { first = r.pick(FIRST); last = r.pick(LAST); } while (taken.has(`${first} ${last}`));
    taken.add(`${first} ${last}`);
    const mail = r.chance(0.82) ? `${first}.${last}@${r.pick(['example.com', 'example.org', 'example.net', 'mail.example'])}`.toLowerCase() : '';
    const declared = r.chance(declare);
    const later = joined + r.int(20, 200);
    const declOffset = declared ? (r.chance(0.8) || later > 0 ? joined : later) : null;
    const [town, districts] = r.weighted(TOWNS);
    const p = { joined, declared, declOffset, weight: 0.3 + r.next() * 2 };
    p.id = add('People', {
      'First Name': first, 'Last Name': last, Address: `${r.int(1, 180)} ${r.pick(STREETS)}`, City: town,
      Postcode: `${r.pick(districts)} ${r.int(1, 9)}${r.pick(['AB', 'DE', 'FG', 'HJ', 'NP', 'QR', 'TU', 'XY'])}`, Country: 'United Kingdom',
      Phone: r.chance(0.85) ? phoneFor(r, 'United Kingdom') : '', 'E-mail': mail, Joined: cal.day(joined),
      'Contact by Post': r.chance(0.6), 'Contact by E-mail': Boolean(mail) && r.chance(0.75), 'Contact by Phone': r.chance(0.3),
      'Tax Declaration': declared, 'Declaration Date': declOffset == null ? null : cal.day(declOffset),
      Notes: r.chance(0.05) ? r.pick(NOTES) : null,
    });
    people.push(p);
    return p;
  };
  const claimable = (p, d) => p.declared && d >= p.declOffset;
  /** A draw from `list` by each person's weight, as r.weighted does, found by halving: a thousand people take no longer than ten. */
  const drawFrom = (list) => {
    const sums = [];
    let total = 0;
    for (const x of list) sums.push(total += x.weight);
    return () => {
      const at = r.next() * total;
      let lo = 0;
      let hi = sums.length - 1;
      while (lo < hi) { const mid = (lo + hi) >> 1; if (sums[mid] > at) hi = mid; else lo = mid + 1; }
      return list[lo];
    };
  };

  // Lookups.
  const types = TYPES.map(([name, fee, months, description, weight]) => ({ name, fee, months, weight, id: add('Membership Types', { Name: name, Fee: fee, Months: months, Description: description }) }));
  const general = add('Campaigns', { Name: 'General Fund', Opens: cal.day(-days), Description: 'Gifts for the trust\'s work in general' });
  const appeals = APPEALS.map(([name, opens, closes, target, description, pull]) => ({ opens, closes, pull, name, id: add('Campaigns', { Name: name, Opens: cal.day(opens), Closes: cal.day(closes), Target: target, Description: description }) }));
  const roles = ROLES.map(([name, day, start, value, description, n]) => ({ day, start, n, name, id: add('Volunteer Roles', { Name: name, 'Usual Day': day == null ? 'Event days' : ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][day], 'Hour Value': value, Description: description }) }));
  const funders = FUNDERS.map(([name, kind]) => {
    const slug = name.toLowerCase().replace(/[^a-z]+/g, '');
    return add('Funders', { Name: name, Kind: kind, 'Contact Name': `${r.pick(FIRST)} ${r.pick(LAST)}`, 'E-mail': `grants@${slug}.example`, Phone: phoneFor(r, 'United Kingdom') });
  });

  // Memberships. Each runs its months; most are renewed, some days before the end, some weeks after; the rest lapse.
  const renewals = new Map(); // day -> memberships renewed that day
  const pay = () => r.weighted([['Standing order', 4], ['Card', 4], ['Cash', 2], ['Bank transfer', 2]]);
  const members = [];
  const isMember = new Set();
  const join = (p, type, start, paid, method = pay()) => {
    const end = type.months ? endOf(start, type.months) : null;
    const id = add('Memberships', { Person: p.id, 'Membership Type': type.id, Starts: cal.day(start), Ends: end == null ? null : cal.day(end), Amount: type.fee, 'Paid On': cal.day(paid), Method: method, Renewed: false });
    income(paid, 'Membership', `${type.name} membership`, type.fee, method, `Membership ${id}`);
    if (!isMember.has(p)) { isMember.add(p); members.push(p); }
    if (end == null) return;
    // A standing order renews on the day; others renew around the reminder, or never.
    const keep = method === 'Standing order' ? 0.95 : type.name === 'Junior' ? 0.65 : 0.78;
    if (!r.chance(keep)) return;
    const at = method === 'Standing order' ? end + 1 : end + r.weighted([[r.int(-21, 0), 6], [r.int(1, 14), 3], [r.int(15, 45), 1]]);
    if (!renewals.has(at)) renewals.set(at, []);
    renewals.get(at).push({ id, p, type, end, method });
  };
  const pickType = () => r.weighted(types.filter((t) => t.weight).map((t) => [t, t.weight]));
  // Members on the books before the records begin, and a few life members.
  for (let i = 0; i < 185; i++) {
    const start = -days - r.int(1, 365);
    const p = newPerson(start - r.int(0, 2500), { declare: 0.45 });
    join(p, pickType(), start, start);
  }
  for (let i = 0; i < 8; i++) {
    const start = -days - r.int(400, 4000);
    join(newPerson(start - r.int(0, 1500), { declare: 0.6 }), types[4], start, start, 'Cheque');
  }

  // Supporters who are not members, and the regular givers among everyone.
  for (let i = 0; i < 140; i++) newPerson(-days - r.int(1, 2500), { declare: 0.55 });
  const regular = r.some(people, 52).map((p) => ({ p, dom: r.int(1, 28), amount: r.pick([5, 5, 10, 10, 10, 15, 20, 25]), from: r.chance(0.7) ? -days : -r.int(30, days - 30), to: r.chance(0.12) ? -r.int(40, 400) : 1 }));
  for (let i = 0; i < 8; i++) {
    const from = -r.int(60, days - 30);
    regular.push({ p: newPerson(from, { declare: 0.7 }), dom: r.int(1, 28), amount: r.pick([10, 15, 20]), from, to: 1 });
  }

  // Events every other Saturday, the summer fair in July, a winter concert in December and the AGM in May, until 90 days ahead.
  const events = [];
  let turn = 0;
  let firstSat = -days;
  while (cal.weekday(firstSat) !== 6) firstSat++;
  for (let d = firstSat; d <= 90; d += 14) {
    const month = cal.at(d).getMonth();
    let e;
    if (month === 6 && !events.some((x) => x.name === 'Summer Fair' && Math.abs(x.d - d) < 60)) e = ['Summer Fair', '11:00', 'Community Garden', 300, 2, 350];
    else if (month === 11 && !events.some((x) => x.name === 'Winter Concert' && Math.abs(x.d - d) < 60)) e = ['Winter Concert', '19:00', 'Main Hall', 120, 8, 300];
    else e = EVENTS[turn++ % EVENTS.length];
    events.push({ d, name: e[0], start: e[1], venue: e[2], capacity: e[3], price: e[4], costs: e[5] });
    // The AGM on a Tuesday evening in May.
    if (month === 4 && !events.some((x) => x.name === 'Annual General Meeting' && Math.abs(x.d - d) < 60)) events.push({ d: d + 3, name: 'Annual General Meeting', start: '19:00', venue: 'Upper Room', capacity: 80, price: 0, costs: 25 });
  }
  // A coffee morning with a raffle on the first Tuesday of each month.
  for (let d = -days; d <= 90; d++) if (cal.weekday(d) === 2 && cal.at(d).getDate() <= 7) events.push({ d, name: 'Coffee Morning', start: '10:00', venue: 'Upper Room', capacity: 50, price: 2, costs: 15 });
  events.sort((a, b) => a.d - b.d);
  // One evening called off for snow.
  const off = r.pick(events.filter((e) => e.d < -30 && e.d > -days + 60 && e.name !== 'Annual General Meeting' && [0, 1, 11].includes(cal.at(e.d).getMonth())));
  const bookings = new Map(); // booking day -> events to take a booking for
  for (const e of events) {
    e.status = e === off ? 'Cancelled' : e.d <= 0 ? 'Held' : 'Planned';
    e.id = add('Events', { Name: e.name, 'Event Date': cal.day(e.d), 'Start Time': e.start, Venue: e.venue, Capacity: e.capacity, 'Ticket Price': e.price, Costs: e.status === 'Cancelled' ? 0 : e.costs, Status: e.status });
    // How full it will be, taken in bookings over the weeks before.
    const tickets = Math.round(e.capacity * (0.45 + r.next() * 0.5));
    e.left = tickets;
    let guard = 0;
    while (guard++ < 200 && e.left > 0) {
      const n = Math.min(e.left, r.weighted([[1, 3], [2, 5], [3, 1.5], [4, 1.5], [6, 0.4]]));
      e.left -= n;
      const at = e.d - r.weighted([[r.int(0, 7), 4], [r.int(8, 21), 4], [r.int(22, 50), 2]]);
      if (at > 0 || at < -days) continue;
      if (!bookings.has(at)) bookings.set(at, []);
      bookings.get(at).push({ e, n });
    }
  }

  // The day by day: memberships, gifts, bookings, claims.
  const unclaimed = [];
  let claimNo = 0;
  for (let d = -days - 50; d <= 0; d++) {
    for (const m of renewals.get(d) || []) {
      base.update('Memberships', m.id, { Renewed: true });
      join(m.p, m.type, m.end + 1, d, m.method);
    }
    if (d < -days) continue;
    const date = cal.at(d);
    const month = date.getMonth();
    // Who the trust knows by today, for gifts and bookings from people already on the records.
    let pool = null;
    const known = () => pool || (pool = drawFrom(people.filter((x) => x.joined <= d)));

    // New members: more in September as the season starts, and in January.
    const joinRate = month === 8 ? 0.35 : month === 0 ? 0.22 : 0.1;
    if (r.chance(joinRate)) {
      const p = r.chance(0.3) ? r.pick(people.filter((x) => x.joined <= d && !isMember.has(x))) : newPerson(d, { declare: 0.45 });
      join(p, pickType(), d, d);
    }

    const gift = (p, campaign, amount, method, isRegular, label) => {
      const ok = claimable(p, d);
      const id = add('Donations', { Person: p.id, Campaign: campaign, 'Gift Date': cal.day(d), Amount: amount, Method: method, Regular: isRegular, Claimable: ok });
      income(d, 'Donations', label, amount, method, `Gift ${id}`);
      if (ok) unclaimed.push({ id, d, amount });
    };
    // Standing orders on their day of the month.
    for (const g of regular) if (date.getDate() === g.dom && d >= g.from && d < g.to) gift(g.p, general, g.amount, 'Standing order', true, 'General Fund');

    // One-off gifts: to an open appeal when there is one, more in December.
    const open = appeals.filter((a) => d >= a.opens && d <= a.closes);
    const rate = (0.55 + open.reduce((n, a) => n + a.pull, 0)) * (month === 11 ? 1.8 : 1);
    const n = Math.floor(rate) + (r.chance(rate - Math.floor(rate)) ? 1 : 0);
    for (let k = 0; k < n; k++) {
      const p = r.chance(0.2) ? newPerson(d) : known()();
      const a = open.length && r.chance(0.75) ? r.weighted(open.map((x) => [x, x.pull])) : null;
      const amount = r.weighted([[5, 8], [10, 20], [15, 6], [20, 18], [25, 12], [30, 5], [50, 14], [75, 3], [100, 8], [150, 2], [250, 2], [500, 0.6], [1000, 0.25]]);
      gift(p, a ? a.id : general, amount, r.weighted([['Card', 4], ['Online', 5], ['Cash', 2], ['Cheque', 1.5], ['Bank transfer', 1]]), false, a ? a.name : 'General Fund');
    }

    // Bookings taken today, paid when booked. Most from people already known, some new faces.
    for (const { e, n: tickets } of bookings.get(d) || []) {
      const p = r.chance(0.18) ? newPerson(d, { declare: 0.3 }) : known()();
      const dropped = e.status === 'Cancelled' || r.chance(0.04);
      const paid = dropped ? 0 : money(tickets * e.price);
      const id = add('Event Bookings', { Event: e.id, Person: p.id, 'Booked On': cal.day(d), Tickets: tickets, 'Amount Paid': paid, Cancelled: dropped, Attended: !dropped && e.status === 'Held' && r.chance(0.92) });
      income(d, 'Events', `${e.name} tickets`, paid, r.weighted([['Card', 3], ['Online', 4], ['Cash', 3]]), `Booking ${id}`);
    }
    // The bucket at the door on the night.
    for (const e of events) if (e.d === d && e.status === 'Held') income(d, 'Collections', `${e.name} collection`, r.int(20, Math.max(40, Math.round(e.capacity * 1.2))), 'Cash', `Event ${e.id}`);

    // On the 12th of January, April, July and October: claim tax on the claimable gifts up to the end of last month.
    if (date.getDate() === 12 && month % 3 === 0 && unclaimed.length) {
      const upTo = d - 12;
      const these = unclaimed.filter((g) => g.d <= upTo);
      if (these.length) {
        const claimed = money(these.reduce((s, g) => s + money(g.amount * 0.25), 0));
        const back = d + r.int(16, 35);
        const ref = `TC-${String(++claimNo).padStart(3, '0')}`;
        const id = add('Tax Claims', { Reference: ref, 'Gifts Up To': cal.day(upTo), 'Submitted On': cal.day(d), 'Amount Claimed': claimed, 'Received On': back <= 0 ? cal.day(back) : null, 'Amount Received': back <= 0 ? claimed : null });
        for (const g of these) { base.update('Donations', g.id, { 'Tax Claim': id }); unclaimed.splice(unclaimed.indexOf(g), 1); }
        income(back, 'Tax claims', `Tax claim ${ref}`, claimed, 'Bank transfer', ref);
      }
    }
  }

  // Volunteers: each has a role and a usual day, comes most weeks, and is on the rota for the next three weeks.
  const volunteers = [];
  for (const role of roles.filter((x) => x.day != null)) {
    for (let i = 0; i < role.n; i++) {
      const p = r.chance(0.75) ? r.pick(members.filter((x) => !volunteers.some((v) => v.p === x))) : newPerson(-r.int(60, days), { declare: 0.4 });
      volunteers.push({ p, role, from: Math.max(p.joined, r.chance(0.75) ? -days : -r.int(60, days)), to: r.chance(0.12) ? -r.int(30, 300) : 99, come: 0.65 + r.next() * 0.3, hours: r.pick([2, 2.5, 3, 3, 4]) });
    }
  }
  for (let d = -days; d <= 21; d++) {
    const wd = cal.weekday(d);
    for (const v of volunteers) {
      if (v.role.day !== wd || d < v.from || d > v.to) continue;
      if (d <= 0 ? !r.chance(v.come) : !r.chance(0.85)) continue;
      add('Volunteer Shifts', { Person: v.p.id, Role: v.role.id, 'Shift Date': cal.day(d), 'Start Time': v.role.start, Hours: v.hours });
    }
  }
  // Stewards at each event held or coming in the next three weeks.
  const steward = roles.find((x) => x.day == null);
  for (const e of events) {
    if (e.status === 'Cancelled' || e.d > 21) continue;
    const able = volunteers.filter((v) => e.d >= v.from && e.d <= v.to);
    for (const v of r.some(able, Math.min(able.length, e.capacity >= 100 ? r.int(4, 7) : r.int(2, 4)))) {
      add('Volunteer Shifts', { Person: v.p.id, Role: steward.id, 'Shift Date': cal.day(e.d), 'Start Time': e.start < '12:00' ? '09:00' : e.start < '18:00' ? '13:00' : '18:00', Hours: r.pick([3, 4, 4, 5]) });
    }
  }

  // Grants: applications, decisions, money in, and the reports funders want.
  for (const [f, purpose, applied, asked, decided, outcome, awarded, paid, ends, due, sent] of GRANTS) {
    const done = outcome === 'Awarded' && sent != null;
    add('Grants', {
      Funder: funders[f], Purpose: purpose, 'Applied On': cal.day(applied), 'Amount Asked': asked, Status: done ? 'Completed' : outcome,
      'Decided On': decided == null ? null : cal.day(decided), 'Amount Awarded': outcome === 'Awarded' ? awarded : null,
      'Paid On': paid == null ? null : cal.day(paid), 'Period Ends': ends == null ? null : cal.day(ends), 'Report Due': due == null ? null : cal.day(due),
      'Report Sent': sent == null ? null : cal.day(sent),
    });
    if (outcome === 'Awarded') income(paid, 'Grants', `${FUNDERS[f][0]}: ${purpose}`, awarded, 'Bank transfer', `Grant from ${FUNDERS[f][0]}`);
  }
  return counts;
}
