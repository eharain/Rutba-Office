// The lettings business: Hollybrook Lettings, an invented landlord that owns
// and manages a dozen buildings in one city (houses, two blocks of flats, two
// converted houses, a parade of shops with flats above and a small yard of
// workshops), with thirty-eight units let to tenants. The buildings and their
// units, the tenants, each lease, a rent charge for every month of every
// lease, the contractors, repair requests, inspections and safety checks,
// and one ledger of every sum of money in or out for each building.
//
// Its records are eighteen months of letting. Each unit's leases are laid
// out first, counted back from today: who is in now and when they leave,
// who renewed, how long a unit stood empty between tenants. Then the months
// are run a day at a time: rent falls due and is paid on time, late or in
// part; things break, more in winter, and contractors are booked to mend
// them; inspections and safety checks come round; the building's bills are
// paid. The questions a landlord asks are its saved queries; its reports
// print them.

import { money, person, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const TENANT = 'Nz([Business Name], [First Name] & " " & [Last Name])';
const UNIT = '[Units].[Name] & ", " & [Properties].[Name]';
const OWED = 'Round([Amount] - Nz([Amount Paid]), 2)';
// The postal address of a unit: a house let whole is its street address alone.
const POSTAL = 'IIf([Properties].[Name] = [Address], IIf([Units].[Name] = "Whole house", [Address], [Units].[Name] & ", " & [Address]), [Units].[Name] & ", " & [Properties].[Name] & ", " & [Address]) & ", " & [Postcode]';
const NET = 'Round(Nz([Money In]) - Nz([Money Out]), 2)';
const LAST_YEAR = { expr: 'DateDiff("d", [Date], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] };

export const spec = {
  // Where a value below nought is right: a check already overdue, a building
  // or a month that cost more than it took in.
  negative: {
    'Inspections due': ['Days Left'], 'Safety certificates due': ['Days Left'],
    'Income and expenditure by property': ['Net'], 'Money in and out by month': ['Net'],
  },
  tables: [
    { name: 'Properties', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Kind', type: 'text' },
      { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' },
      { name: 'Bought', type: 'date' }, { name: 'Purchase Price', type: 'currency' }, { name: 'Current Value', type: 'currency' },
      { name: 'Valued On', type: 'date' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Units', fields: [
      { name: 'Property', ...ref('Properties'), required: true }, { name: 'Name', type: 'text', required: true }, { name: 'Kind', type: 'text' },
      { name: 'Bedrooms', type: 'integer', default: 0 }, { name: 'Market Rent', type: 'currency' }, { name: 'Has Gas', type: 'yesno' },
      { name: 'Furnished', type: 'yesno' }, { name: 'Status', type: 'text', default: 'Vacant' }, { name: 'Vacant Since', type: 'date' },
      { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Tenants', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Business Name', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Emergency Contact', type: 'text' },
      { name: 'Emergency Phone', type: 'text' }, { name: 'References Checked', type: 'yesno' }, { name: 'ID Checked', type: 'yesno' },
      { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Leases', fields: [
      { name: 'Unit', ...ref('Units'), required: true }, { name: 'Tenant', ...ref('Tenants'), required: true }, { name: 'Type', type: 'text', default: 'Residential' },
      { name: 'Start', type: 'date' }, { name: 'End', type: 'date' }, { name: 'Monthly Rent', type: 'currency' },
      { name: 'Rent Day', type: 'integer', default: 1 }, { name: 'Pays By', type: 'text', default: 'Standing order' },
      { name: 'Deposit', type: 'currency' }, { name: 'Deposit Ref', type: 'text' }, { name: 'Renewal', type: 'yesno' },
      { name: 'Status', type: 'text', default: 'Active' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Rent Charges', fields: [
      { name: 'Lease', ...ref('Leases'), required: true }, { name: 'Period', type: 'text' }, { name: 'Due Date', type: 'date' },
      { name: 'Amount', type: 'currency' }, { name: 'Amount Paid', type: 'currency', default: 0 },
    ] },
    { name: 'Contractors', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Trade', type: 'text' }, { name: 'Contact', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Maintenance Requests', fields: [
      { name: 'Unit', ...ref('Units'), required: true }, { name: 'Reported', type: 'date' }, { name: 'Reported By', type: 'text' },
      { name: 'Problem', type: 'text', required: true }, { name: 'Priority', type: 'text', default: 'Normal' }, { name: 'Status', type: 'text', default: 'Reported' },
      { name: 'Contractor', ...ref('Contractors') }, { name: 'Booked For', type: 'date' }, { name: 'Completed', type: 'date' },
      { name: 'Cost', type: 'currency' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Inspections', fields: [
      { name: 'Unit', ...ref('Units'), required: true }, { name: 'Kind', type: 'text' }, { name: 'Due', type: 'date' },
      { name: 'Done', type: 'date' }, { name: 'Inspector', type: 'text' }, { name: 'Outcome', type: 'text' },
      { name: 'Certificate Number', type: 'text' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Transactions', fields: [
      { name: 'Property', ...ref('Properties'), required: true }, { name: 'Date', type: 'date' }, { name: 'Kind', type: 'text' },
      { name: 'Description', type: 'text' }, { name: 'Charge', ...ref('Rent Charges') }, { name: 'Payee or Payer', type: 'text' },
      { name: 'Method', type: 'text' }, { name: 'Money In', type: 'currency', default: 0 }, { name: 'Money Out', type: 'currency', default: 0 },
    ] },
  ],
  queries: {
    'Rent roll': { tables: ['Leases', 'Units', 'Properties', 'Tenants'], columns: [
      { table: 'Properties', field: 'Name', as: 'Property', sort: 'asc' }, { table: 'Units', field: 'Name', as: 'Unit', sort: 'asc' },
      { expr: TENANT, as: 'Tenant' }, { table: 'Leases', field: 'Type' }, { table: 'Leases', field: 'Start' }, { table: 'Leases', field: 'End' },
      { table: 'Leases', field: 'Monthly Rent' }, { table: 'Leases', field: 'Deposit' }, { table: 'Leases', field: 'Status', criteria: ['Active'], show: false },
    ] },
    'Rent arrears by tenant': { tables: ['Rent Charges', 'Leases', 'Tenants', 'Units', 'Properties'], columns: [
      { expr: TENANT, as: 'Tenant', total: 'group' }, { expr: UNIT, as: 'Unit', total: 'group' }, { table: 'Tenants', field: 'Phone', total: 'group' },
      { table: 'Rent Charges', field: 'ID', as: 'Months Owing', total: 'count' },
      { expr: OWED, as: 'Owed', total: 'sum', sort: 'desc' }, { expr: 'DateDiff("d", [Due Date], Date())', as: 'Days Late', total: 'max' },
      { expr: OWED, as: 'Unpaid', total: 'where', criteria: ['> 0'] }, { expr: 'DateDiff("d", [Due Date], Date())', as: 'Overdue', total: 'where', criteria: ['> 0'] },
    ] },
    'Unpaid rent': { tables: ['Rent Charges', 'Leases', 'Tenants', 'Units', 'Properties'], columns: [
      { expr: TENANT, as: 'Tenant', sort: 'asc' }, { expr: POSTAL, as: 'Address' },
      { table: 'Rent Charges', field: 'Period' }, { table: 'Rent Charges', field: 'Due Date', sort: 'asc' }, { table: 'Rent Charges', field: 'Amount' },
      { table: 'Rent Charges', field: 'Amount Paid' }, { expr: OWED, as: 'Owed', criteria: ['> 0'] },
      { expr: 'DateDiff("d", [Due Date], Date())', as: 'Days Late', criteria: ['> 0'] },
    ] },
    'Rent due in the next 14 days': { tables: ['Rent Charges', 'Leases', 'Tenants', 'Units', 'Properties'], columns: [
      { table: 'Rent Charges', field: 'Due Date', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Due Date])', as: 'Days Until', criteria: ['Between 0 And 14'] },
      { expr: TENANT, as: 'Tenant' }, { expr: UNIT, as: 'Unit' }, { table: 'Leases', field: 'Pays By' }, { table: 'Rent Charges', field: 'Amount' },
      { expr: OWED, as: 'Owed', criteria: ['> 0'], show: false },
    ] },
    'Leases ending in the next 60 days': { tables: ['Leases', 'Units', 'Properties', 'Tenants'], columns: [
      { table: 'Properties', field: 'Name', as: 'Property' }, { table: 'Units', field: 'Name', as: 'Unit' }, { expr: TENANT, as: 'Tenant' },
      { table: 'Tenants', field: 'Phone' }, { table: 'Leases', field: 'End', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [End])', as: 'Days Left', criteria: ['Between 0 And 60'] },
      { table: 'Leases', field: 'Monthly Rent' }, { table: 'Units', field: 'Market Rent' }, { table: 'Leases', field: 'Status', criteria: ['Active'], show: false },
    ] },
    'Vacant units': { tables: ['Units', 'Properties'], columns: [
      { table: 'Properties', field: 'Name', as: 'Property', sort: 'asc' }, { table: 'Units', field: 'Name', as: 'Unit' }, { table: 'Units', field: 'Kind' },
      { table: 'Units', field: 'Bedrooms' }, { table: 'Units', field: 'Market Rent' }, { table: 'Units', field: 'Vacant Since' },
      { expr: 'DateDiff("d", [Vacant Since], Date())', as: 'Days Empty' },
      { expr: 'Round([Market Rent] * 12 / 365 * DateDiff("d", [Vacant Since], Date()), 2)', as: 'Rent Lost' },
      { table: 'Units', field: 'Status', criteria: ['Vacant'], show: false },
    ] },
    'Open repairs': { tables: ['Maintenance Requests', 'Units', 'Properties'], columns: [
      { table: 'Maintenance Requests', field: 'ID', as: 'Job' }, { table: 'Maintenance Requests', field: 'Priority' },
      { expr: 'IIf([Priority] = "Urgent", 1, IIf([Priority] = "High", 2, IIf([Priority] = "Normal", 3, 4)))', as: 'Rank', sort: 'asc', show: false },
      { expr: UNIT, as: 'Unit' }, { table: 'Maintenance Requests', field: 'Problem' }, { table: 'Maintenance Requests', field: 'Reported', sort: 'asc' },
      { expr: 'DateDiff("d", [Reported], Date())', as: 'Days Open' }, { table: 'Maintenance Requests', field: 'Status', criteria: ['Not In ("Done", "Cancelled")'] },
    ] },
    'Repairs booked': { tables: ['Maintenance Requests', 'Units', 'Properties', 'Contractors'], columns: [
      { table: 'Contractors', field: 'Name', as: 'Contractor', sort: 'asc' }, { table: 'Contractors', field: 'Phone' },
      { table: 'Maintenance Requests', field: 'ID', as: 'Job' }, { table: 'Maintenance Requests', field: 'Booked For', sort: 'asc' },
      { table: 'Maintenance Requests', field: 'Priority' }, { expr: POSTAL, as: 'Address' },
      { table: 'Maintenance Requests', field: 'Problem' }, { table: 'Maintenance Requests', field: 'Status', criteria: ['Booked'], show: false },
    ] },
    'Repair costs by contractor': { tables: ['Maintenance Requests', 'Contractors'], columns: [
      { table: 'Contractors', field: 'Name', as: 'Contractor', total: 'group' }, { table: 'Contractors', field: 'Trade', total: 'group' },
      { table: 'Maintenance Requests', field: 'ID', as: 'Jobs', total: 'count' }, { table: 'Maintenance Requests', field: 'Cost', as: 'Spent', total: 'sum', sort: 'desc' },
      { table: 'Maintenance Requests', field: 'Status', total: 'where', criteria: ['Done'] },
      { expr: 'DateDiff("d", [Completed], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Inspections due': { tables: ['Inspections', 'Units', 'Properties'], columns: [
      { table: 'Inspections', field: 'Due', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Due])', as: 'Days Left', criteria: ['<= 30'] },
      { table: 'Inspections', field: 'Kind' }, { expr: UNIT, as: 'Unit' }, { table: 'Inspections', field: 'Notes' },
      { table: 'Inspections', field: 'Done', criteria: ['Is Null'], show: false },
    ] },
    'Safety certificates due': { tables: ['Inspections', 'Units', 'Properties'], columns: [
      { table: 'Inspections', field: 'Due', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Due])', as: 'Days Left', criteria: ['<= 60'] },
      { table: 'Inspections', field: 'Kind', criteria: ['In ("Gas safety", "Electrical safety")'] }, { expr: UNIT, as: 'Unit' },
      { table: 'Inspections', field: 'Done', criteria: ['Is Null'], show: false },
    ] },
    'Deposits held': { tables: ['Leases', 'Tenants', 'Units', 'Properties'], columns: [
      { table: 'Properties', field: 'Name', as: 'Property', sort: 'asc' }, { table: 'Units', field: 'Name', as: 'Unit', sort: 'asc' }, { expr: TENANT, as: 'Tenant' },
      { table: 'Leases', field: 'Type' }, { table: 'Leases', field: 'Deposit' }, { table: 'Leases', field: 'Deposit Ref' },
      { table: 'Leases', field: 'Status', criteria: ['Active'], show: false }, { table: 'Leases', field: 'Deposit', as: 'Held', criteria: ['> 0'], show: false },
    ] },
    'Deposits to protect': { tables: ['Leases', 'Tenants', 'Units', 'Properties'], columns: [
      { expr: TENANT, as: 'Tenant' }, { expr: UNIT, as: 'Unit' }, { table: 'Leases', field: 'Start', sort: 'asc' }, { table: 'Leases', field: 'Status' },
      { table: 'Leases', field: 'Deposit' }, { expr: '30 - DateDiff("d", [Start], Date())', as: 'Days to Protect' },
      { table: 'Leases', field: 'Type', criteria: ['Residential'], show: false }, { table: 'Leases', field: 'Deposit Ref', criteria: ['Is Null'], show: false },
      { table: 'Leases', field: 'Status', as: 'State', criteria: ['In ("Active", "Upcoming")'], show: false },
    ] },
    'Current tenants': { tables: ['Leases', 'Tenants', 'Units', 'Properties'], columns: [
      { table: 'Properties', field: 'Name', as: 'Property', sort: 'asc' }, { table: 'Units', field: 'Name', as: 'Unit', sort: 'asc' }, { expr: TENANT, as: 'Tenant' },
      { table: 'Tenants', field: 'Phone' }, { table: 'Tenants', field: 'E-mail' }, { table: 'Tenants', field: 'Emergency Contact' }, { table: 'Tenants', field: 'Emergency Phone' },
      { table: 'Leases', field: 'Status', criteria: ['Active'], show: false },
    ] },
    'Income and expenditure by property': { tables: ['Transactions', 'Properties'], columns: [
      { table: 'Properties', field: 'Name', as: 'Property', total: 'group', sort: 'asc' }, { table: 'Transactions', field: 'Money In', as: 'Money In', total: 'sum' },
      { table: 'Transactions', field: 'Money Out', as: 'Money Out', total: 'sum' }, { expr: NET, as: 'Net', total: 'sum' }, LAST_YEAR,
    ] },
    'Costs by kind': { tables: ['Transactions'], columns: [
      { table: 'Transactions', field: 'Kind', total: 'group' }, { table: 'Transactions', field: 'ID', as: 'Payments', total: 'count' },
      { table: 'Transactions', field: 'Money Out', as: 'Spent', total: 'sum', sort: 'desc' },
      { table: 'Transactions', field: 'Money Out', as: 'Paid Out', total: 'where', criteria: ['> 0'] }, LAST_YEAR,
    ] },
    'Money in and out by month': { tables: ['Transactions'], columns: [
      { expr: 'Format([Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Transactions', field: 'Money In', as: 'Money In', total: 'sum' },
      { table: 'Transactions', field: 'Money Out', as: 'Money Out', total: 'sum' }, { expr: NET, as: 'Net', total: 'sum' },
      { expr: 'DateDiff("m", [Date], Date())', as: 'Months Ago', total: 'where', criteria: ['Between 0 And 12'] },
    ] },
    'Gross yield by property': { tables: ['Properties', 'Units'], columns: [
      { table: 'Properties', field: 'Name', as: 'Property', total: 'group' }, { table: 'Properties', field: 'Kind', total: 'group' },
      { table: 'Units', field: 'ID', as: 'Units', total: 'count' }, { table: 'Properties', field: 'Current Value', as: 'Value', total: 'max' },
      { expr: 'Round([Market Rent] * 12, 2)', as: 'Annual Rent', total: 'sum' },
      { expr: 'Round([Market Rent] * 1200 / [Current Value], 2)', as: 'Gross Yield %', total: 'sum', sort: 'desc' },
    ] },
  },
  reports: [
    { name: 'Rent roll', query: 'Rent roll', group: 'Property', columns: ['Unit', 'Tenant', { field: 'Start', type: 'date' }, { field: 'End', type: 'date' }, { field: 'Monthly Rent', type: 'currency' }, { field: 'Deposit', type: 'currency' }], totals: [{ field: 'Monthly Rent', fn: 'sum' }, { field: 'Deposit', fn: 'sum' }] },
    { name: 'Arrears letters', title: 'Rent arrears', query: 'Unpaid rent', group: 'Tenant', columns: ['Address', 'Period', { field: 'Due Date', type: 'date' }, { field: 'Days Late', type: 'integer' }, { field: 'Amount', type: 'currency' }, { field: 'Amount Paid', type: 'currency' }, { field: 'Owed', type: 'currency' }], totals: [{ field: 'Owed', fn: 'sum' }] },
    { name: 'Repairs for contractors', title: 'Jobs booked', query: 'Repairs booked', group: 'Contractor', columns: ['Job', { field: 'Booked For', type: 'date' }, 'Priority', 'Address', 'Problem'] },
    { name: 'Income and expenditure', title: 'Income and expenditure, last 12 months', query: 'Income and expenditure by property', columns: ['Property', { field: 'Money In', type: 'currency' }, { field: 'Money Out', type: 'currency' }, { field: 'Net', type: 'currency' }], totals: [{ field: 'Money In', fn: 'sum' }, { field: 'Money Out', fn: 'sum' }, { field: 'Net', fn: 'sum' }] },
    { name: 'Leases ending', title: 'Leases ending in the next 60 days', query: 'Leases ending in the next 60 days', columns: ['Property', 'Unit', 'Tenant', 'Phone', { field: 'End', type: 'date' }, { field: 'Days Left', type: 'integer' }, { field: 'Monthly Rent', type: 'currency' }, { field: 'Market Rent', type: 'currency' }], totals: [{ field: 'Monthly Rent', fn: 'sum' }] },
    { name: 'Empty units', title: 'Vacant units and rent lost', query: 'Vacant units', columns: ['Property', 'Unit', 'Kind', { field: 'Market Rent', type: 'currency' }, { field: 'Vacant Since', type: 'date' }, { field: 'Days Empty', type: 'integer' }, { field: 'Rent Lost', type: 'currency' }], totals: [{ field: 'Rent Lost', fn: 'sum' }, { field: 'Market Rent', fn: 'sum' }] },
    { name: 'Checks due', title: 'Inspections and safety checks due', query: 'Inspections due', group: 'Kind', columns: [{ field: 'Due', type: 'date' }, { field: 'Days Left', type: 'integer' }, 'Unit', 'Notes'] },
    { name: 'Monthly accounts', query: 'Money in and out by month', columns: ['Month', { field: 'Money In', type: 'currency' }, { field: 'Money Out', type: 'currency' }, { field: 'Net', type: 'currency' }], totals: [{ field: 'Money In', fn: 'sum' }, { field: 'Money Out', fn: 'sum' }, { field: 'Net', fn: 'sum' }] },
  ],
};

export const about = {
  name: 'Hollybrook Lettings',
  kind: 'Property rentals',
  line: 'Buildings, units, tenants, leases, rent, repairs, safety checks and the accounts for a landlord.',
  summary: [
    'Hollybrook Lettings is an invented letting business in Leeds that owns and manages twelve buildings: six houses, two blocks of flats, two houses split into flats, a parade of three shops with two flats above, and a yard of two workshops bought about a year ago and refitted. Between them they hold thirty-eight units let to tenants on residential and commercial leases.',
    'The sample holds eighteen months of letting. Each unit has its run of leases: tenants who renewed, tenants who left, and the weeks a unit stood empty between them. Rent is charged once a month on each lease\'s rent day. Most tenants pay on the day by standing order, some pay late, and a few have fallen behind and pay part of what is due, so there are arrears to chase. Things break and are reported, more of them in winter, and a contractor is booked to mend each one. Homes are inspected twice a year and shops and workshops once a year; gas appliances are checked every year and the wiring every five years. Insurance, loan interest, cleaning of shared areas, letting costs and the bills for empty units are paid out of each building\'s account.',
    'Money for each building goes through one Transactions table, rent in and costs out, so income and expenditure, net income and the month-by-month accounts come from one place. A rent payment also points at the month\'s rent charge it pays, and the charge keeps the total paid against it, which is how arrears are found.',
  ],
  suits: ['A private landlord with a few houses or flats', 'A small letting agent managing homes for owners', 'Holiday lets, with weekly lets in place of leases', 'Student houses let by the room', 'Storage units or lock-up garages', 'Allotments or moorings let by the year', 'A small commercial landlord with shops and workshops'],
  tables: {
    Properties: { about: 'Each building you own or manage, one record each.', fields: {
      Name: 'What you call the building, such as Linden Court or 22 Mill Lane. Each name is used once.',
      Kind: 'House, Block of flats, Converted house, Shops with flats or Workshop yard.',
      Address: 'The street address.', City: 'The town or city.', Postcode: 'The postcode.',
      Bought: 'The day it was bought.', 'Purchase Price': 'What it cost.',
      'Current Value': 'What it is worth now, from the last valuation. The yield query divides rent by this.',
      'Valued On': 'When it was last valued.', Notes: 'Anything worth knowing about the building: a lease on the freehold, a loan, access.',
    } },
    Units: { about: 'What is let: a house, a flat, a shop or a workshop. A house let whole is one unit; a block of flats has one unit for each flat.', fields: {
      Property: 'The building it is in.', Name: 'The unit\'s name within the building, such as Flat 3, Shop 2 or Whole house.',
      Kind: 'House, Flat, Shop or Workshop.', Bedrooms: 'How many bedrooms; nought for a shop or workshop.',
      'Market Rent': 'What the unit would let for each month now. Used for empty units and the yield.',
      'Has Gas': 'Yes when the unit has gas appliances, so it needs a gas safety check every year.',
      Furnished: 'Yes when it is let with furniture.', Status: 'Let or Vacant, as it stands today.',
      'Vacant Since': 'The first day it stood empty; empty while it is let.', Notes: 'Anything about the unit: the parking space, the meter numbers, what the inventory says.',
    } },
    Tenants: { about: 'Everyone who rents from you, past and present, one record each. A tenant who renews stays one record with several leases.', fields: {
      'First Name': 'Their first name, or the first name of the person who signs for a business.', 'Last Name': 'Their family name.',
      'Business Name': 'The business that rents a shop or workshop; empty for a home.', Phone: 'Their phone number.',
      'E-mail': 'Their e-mail address.', 'Emergency Contact': 'Who to call if you cannot reach them, and how they are related.',
      'Emergency Phone': 'That person\'s phone number.', 'References Checked': 'Yes once their references and credit check have come back.',
      'ID Checked': 'Yes once you have seen proof of who they are and their right to rent.', Notes: 'Anything worth knowing: pets, a guarantor, a forwarding address.',
    } },
    Leases: { about: 'Each agreement to let a unit to a tenant for a term. A renewal is a new lease for the same tenant and unit.', fields: {
      Unit: 'The unit let.', Tenant: 'Who it is let to.', Type: 'Residential or Commercial.',
      Start: 'The first day of the lease.', End: 'The last day of the fixed term.', 'Monthly Rent': 'The rent each month.',
      'Rent Day': 'The day of the month rent falls due.', 'Pays By': 'Standing order, Bank transfer, Card or Cash.',
      Deposit: 'The deposit taken: five weeks\' rent for a home, three months\' for a shop or workshop.',
      'Deposit Ref': 'The reference the deposit protection scheme gave for a home\'s deposit (invented here). Empty until it is protected.',
      Renewal: 'Yes when the lease follows an earlier one for the same tenant.', Status: 'Upcoming, Active or Ended, as it stands today.',
      Notes: 'Anything agreed that is not in the fields: a break clause, a rent review, a pet clause.',
    } },
    'Rent Charges': { about: 'The rent due on a lease for one month: one record for each lease and month, made as each month comes round, and a month ahead.', fields: {
      Lease: 'The lease the rent is due on.', Period: 'The month it is for, as 2026-10.', 'Due Date': 'The day it is due.',
      Amount: 'How much is due.', 'Amount Paid': 'How much has been paid against it so far. Add each payment here as you enter it in Transactions; what is left is owed.',
    } },
    Contractors: { about: 'The trades you call on for repairs and safety checks.', fields: {
      Name: 'The firm\'s name.', Trade: 'What they do, such as Plumbing and heating or Electrical.', Contact: 'The person you deal with.',
      Phone: 'Their phone number.', 'E-mail': 'Where to send job details.', Notes: 'Rates, insurance, call-out hours.',
    } },
    'Maintenance Requests': { about: 'Each repair asked for or found, from the day it is reported to the day it is done.', fields: {
      Unit: 'The unit where the problem is.', Reported: 'The day you heard of it.', 'Reported By': 'Tenant, Inspection or Landlord.',
      Problem: 'What is wrong, in a few words.', Priority: 'Urgent (same or next day), High (within three days), Normal or Low.',
      Status: 'Reported, Booked, Done or Cancelled.', Contractor: 'Who is doing the work; empty until someone is booked.',
      'Booked For': 'The day the contractor is due.', Completed: 'The day the work was finished.',
      Cost: 'What the work cost. The same sum goes into Transactions as money out.', Notes: 'Access arrangements, parts on order, what was found.',
    } },
    Inspections: { about: 'Routine inspections of each unit and the safety checks the law asks for, each one due on a date. When one is done, the next one is entered with its due date.', fields: {
      Unit: 'The unit inspected.', Kind: 'Routine inspection, Gas safety or Electrical safety.', Due: 'When it must be done by.',
      Done: 'The day it was done; empty until then.', Inspector: 'Who did it: one of your staff, or the contractor who did the check.',
      Outcome: 'What was found, such as Good, Action needed, Pass or Unsatisfactory.', 'Certificate Number': 'The number on the safety certificate (invented here).',
      Notes: 'Anything to know before going, or what happened.',
    } },
    Transactions: { about: 'Every sum of money in or out for a building: rent received, and every cost paid. Each building\'s accounts are made from these.', fields: {
      Property: 'The building the money is for.', Date: 'The day it was paid or received.',
      Kind: 'Rent, Repairs, Insurance, Loan interest, Shared areas, Safety checks, Letting costs, Agent fees or Empty unit costs.',
      Description: 'What it was for.', Charge: 'For rent, the month\'s rent charge it pays; empty for a cost.',
      'Payee or Payer': 'Who paid you, or who you paid.', Method: 'How it was paid.',
      'Money In': 'What came in; nought for a cost.', 'Money Out': 'What went out; nought for rent.',
    } },
  },
  queries: {
    'Rent roll': 'Every lease running today, by building and unit, with the tenant, the term, the rent and the deposit.',
    'Rent arrears by tenant': 'Each tenant who owes rent that is past due, with how many months are unpaid, how much they owe and how many days late the oldest month is. Former tenants who still owe are on it too.',
    'Unpaid rent': 'Every month\'s rent past its due day and not paid in full, by tenant, with the address to write to.',
    'Rent due in the next 14 days': 'Rent falling due in the next two weeks and not yet paid, with how each tenant pays, so you know what to expect in the bank.',
    'Leases ending in the next 60 days': 'Leases whose term ends within two months, with the rent now and what the unit would let for, so you can offer a renewal or plan to relet.',
    'Vacant units': 'Units standing empty, how long for, and the rent lost at the market rent since they emptied.',
    'Open repairs': 'Repairs not yet done or cancelled, the most urgent first and then the oldest, with how many days each has been open.',
    'Repairs booked': 'Jobs booked with a contractor and not yet done, by contractor and day, with the address and the problem.',
    'Repair costs by contractor': 'What each contractor has been paid for finished jobs in the last twelve months, and how many jobs.',
    'Inspections due': 'Inspections and safety checks due in the next 30 days or already overdue (a Days Left below nought is overdue).',
    'Safety certificates due': 'Gas and electrical safety checks due in the next 60 days or overdue, so the engineer can be booked in time.',
    'Deposits held': 'The deposit taken on every lease running today, with its protection reference.',
    'Deposits to protect': 'Home deposits not yet protected, with how many days are left of the 30 allowed from the start of the lease.',
    'Current tenants': 'Everyone living or trading in your units today, with their phone, e-mail and emergency contact.',
    'Income and expenditure by property': 'For each building over the last twelve months: rent received, costs paid and the net income.',
    'Costs by kind': 'What was spent in the last twelve months on each kind of cost, the biggest first.',
    'Money in and out by month': 'Money in, money out and the net for each month of the last year.',
    'Gross yield by property': 'Each building\'s rent for a year at market rents as a percentage of its current value, the highest first.',
  },
  reports: {
    'Rent roll': 'The rent roll: each building\'s units, tenants, terms, rents and deposits, with totals for each building and for all of them. The page a lender or buyer asks for.',
    'Arrears letters': 'Each tenant in arrears on their own section, with every unpaid month and the total owed: the list to write the reminder letters from.',
    'Repairs for contractors': 'A page for each contractor listing the jobs booked with them, the day, the address and the problem, to send or hand over.',
    'Income and expenditure': 'Money in, money out and the net for each building over the last twelve months, with the totals: the figures for your accounts.',
    'Leases ending': 'Leases ending in the next 60 days with the tenant\'s phone, the rent now and the market rent.',
    'Empty units': 'The units standing empty, how long for and the rent lost on each.',
    'Checks due': 'Inspections and safety checks due in the next 30 days or overdue, grouped by kind, to plan the visits.',
    'Monthly accounts': 'One line a month for the last year with money in, money out and the net, and the totals.',
  },
  empty: 'Start with Properties, then the Units in each one, then Contractors. Enter each Tenant and their Lease, and for every unit the next inspection and safety checks due. Each month add a Rent Charge for every running lease; when rent comes in, enter it in Transactions pointing at the charge, and add it to the charge\'s Amount Paid. Enter repairs in Maintenance Requests as they are reported, and every bill you pay in Transactions with its building.',
  adapt: [
    'Manage for owners: add an Owners table with their management fee, point Properties at it, and add a query that takes your fee off each owner\'s rent to show what you pay them.',
    'Let holiday homes: use a weekly Rent Charge in place of a monthly one, rename Leases to Bookings with arrival and departure dates, and add a cleaning cost to each booking.',
    'Let rooms in shared houses: make each room a unit, and add a field to Units for the shared bathroom and kitchen it uses.',
    'Let storage units or garages: drop Bedrooms and Has Gas from Units, add a Size field, and keep the rest as it is.',
    'Keep the energy rating: add an Energy Rating and an Energy Certificate Expires field to Units, and a query for certificates running out.',
    'Record rent reviews on commercial leases: add a Next Review date to Leases and a query for reviews due in the next 90 days.',
    'Recharge damage to tenants: add a Recharge field to Maintenance Requests, and take it from the deposit when the lease ends.',
  ],
};

// The buildings: [name, kind, street, gas in the homes, [unit name, kind, bedrooms, market rent]...].
const BUILDINGS = [
  ['Linden Court', 'Block of flats', '14 Beech Grove', false, [
    ['Flat 1', 'Flat', 1, 750], ['Flat 2', 'Flat', 2, 895], ['Flat 3', 'Flat', 1, 750], ['Flat 4', 'Flat', 2, 910], ['Flat 5', 'Flat', 1, 765], ['Flat 6', 'Flat', 2, 920],
    ['Flat 7', 'Flat', 1, 770], ['Flat 8', 'Flat', 2, 925], ['Flat 9', 'Flat', 1, 775], ['Flat 10', 'Flat', 2, 935], ['Flat 11', 'Flat', 1, 785], ['Flat 12', 'Flat', 2, 950]]],
  ['Cedar House', 'Block of flats', '2 Cedar Avenue', true, [
    ['Flat A', 'Flat', 1, 710], ['Flat B', 'Flat', 2, 840], ['Flat C', 'Flat', 1, 715], ['Flat D', 'Flat', 2, 850], ['Flat E', 'Flat', 1, 720], ['Flat F', 'Flat', 2, 860],
    ['Flat G', 'Flat', 1, 730], ['Flat H', 'Flat', 2, 875]]],
  ['48 Orchard Way', 'Converted house', '48 Orchard Way', true, [['Ground floor flat', 'Flat', 2, 795], ['First floor flat', 'Flat', 1, 680], ['Top floor flat', 'Flat', 1, 650]]],
  ['7 Linden Close', 'Converted house', '7 Linden Close', true, [['Lower flat', 'Flat', 2, 780], ['Upper flat', 'Flat', 1, 670]]],
  ['22 Mill Lane', 'House', '22 Mill Lane', true, [['Whole house', 'House', 3, 1150]]],
  ['31 Mill Lane', 'House', '31 Mill Lane', true, [['Whole house', 'House', 2, 925]]],
  ['9 Quarry Road', 'House', '9 Quarry Road', true, [['Whole house', 'House', 3, 1095]]],
  ['64 Station Road', 'House', '64 Station Road', true, [['Whole house', 'House', 4, 1395]]],
  ['15 Canal Walk', 'House', '15 Canal Walk', true, [['Whole house', 'House', 3, 1225]]],
  ['3 Church Street', 'House', '3 Church Street', true, [['Whole house', 'House', 2, 950]]],
  ['Market Parade', 'Shops with flats', '110 Harbour Street', true, [
    ['Shop 1', 'Shop', 0, 1350], ['Shop 2', 'Shop', 0, 1200], ['Shop 3', 'Shop', 0, 1550], ['Flat over 1', 'Flat', 1, 675], ['Flat over 3', 'Flat', 1, 695]]],
  ['Riverside Yard', 'Workshop yard', 'Riverside Drive', false, [['Workshop 1', 'Workshop', 0, 950], ['Workshop 2', 'Workshop', 0, 1075]]],
];
const MORTGAGED = ['Linden Court', 'Cedar House', '64 Station Road', 'Riverside Yard'];
const SHARED_AREAS = ['Linden Court', 'Cedar House', 'Market Parade'];

// What goes wrong: [problem, trade, priority, cost from, cost to, where (h homes, s shops and workshops), winter weight].
const PROBLEMS = [
  ['Boiler not firing, no heating or hot water', 'Plumbing and heating', 'Urgent', 140, 680, 'h', 3],
  ['Radiators cold at the top', 'Plumbing and heating', 'Normal', 60, 150, 'h', 2.5],
  ['Dripping tap', 'Plumbing and heating', 'Low', 55, 120, 'hs', 1],
  ['Leak under the bathroom sink', 'Plumbing and heating', 'High', 75, 230, 'h', 1],
  ['Toilet will not flush', 'Plumbing and heating', 'High', 70, 190, 'hs', 1],
  ['Shower pump failed', 'Plumbing and heating', 'High', 180, 420, 'h', 1],
  ['Blocked drain at the back', 'Plumbing and heating', 'High', 90, 260, 'hs', 1.5],
  ['Light fitting not working', 'Electrical', 'Normal', 60, 150, 'hs', 1],
  ['Socket scorched in the kitchen', 'Electrical', 'Urgent', 85, 220, 'hs', 1],
  ['Smoke alarm keeps sounding', 'Electrical', 'High', 45, 95, 'h', 1],
  ['Extractor fan not working', 'Electrical', 'Normal', 90, 240, 'h', 1.5],
  ['Front door lock sticking', 'Locksmith', 'High', 70, 180, 'hs', 1],
  ['Keys lost, locks changed', 'Locksmith', 'Urgent', 110, 240, 'hs', 1],
  ['Roller shutter jammed', 'General repairs', 'High', 120, 380, 's', 1],
  ['Damp patch on the bedroom wall', 'Roofing and damp', 'Normal', 150, 900, 'h', 2.5],
  ['Slipped roof tiles after high wind', 'Roofing and damp', 'High', 180, 750, 'hs', 2.5],
  ['Gutter overflowing', 'Roofing and damp', 'Normal', 80, 220, 'hs', 2],
  ['Kitchen cupboard door hanging off', 'Joinery', 'Low', 45, 120, 'h', 1],
  ['Window will not close', 'Joinery', 'High', 70, 260, 'hs', 1.5],
  ['Mould around the bathroom window', 'Decorating', 'Normal', 90, 320, 'h', 2.5],
  ['Mice in the kitchen', 'Pest control', 'High', 95, 210, 'hs', 1.5],
  ['Fence panel blown down', 'General repairs', 'Normal', 85, 260, 'h', 2],
];

// The firms called on: [name, trade]. Invented names.
const CONTRACTORS = [
  ['Meridian Plumbing and Heating', 'Plumbing and heating'], ['Cobalt Heating Services', 'Plumbing and heating'],
  ['Silverline Electrical', 'Electrical'], ['Kestrel Locksmiths', 'Locksmith'], ['Sandstone Roofing and Damp', 'Roofing and damp'],
  ['Riverbend Joinery', 'Joinery'], ['Crescent Decorators', 'Decorating'], ['Bramble Pest Control', 'Pest control'],
  ['Marlow General Repairs', 'General repairs'], ['Greenacre Grounds', 'Grounds'],
];

// Businesses that rent the shops and workshops. Invented names.
const BUSINESSES = ['Amberly Florist', 'Copperfield Barbers', 'Larkspur Books', 'Juniper Nail Studio', 'Granite Print Shop', 'Willow Dry Cleaners',
  'Linden Cards and Gifts', 'Alder Phone Repairs', 'Saffron Takeaway', 'Orchard Upholstery', 'Northgate Motor Repairs', 'Kestrel Cycle Repairs'];

const STAFF_DOMAIN = 'hollybrook.example';
const slug = (s) => s.toLowerCase().replace(/[^a-z]+/g, '');
const round5 = (n) => Math.round(n / 5) * 5;

/** Fill a Hollybrook Lettings database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 540 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };

  // Dates as day offsets from today. Month arithmetic is done on the
  // calendar's own dates, so the 14th stays the 14th from month to month.
  const b0 = cal.at(0);
  const zero = Date.UTC(b0.getFullYear(), b0.getMonth(), b0.getDate());
  const offsetOf = (y, m, d) => Math.round((Date.UTC(y, m, d) - zero) / 864e5);
  const dom = (off) => cal.at(off).getDate();
  const addMonths = (off, n) => {
    const d = cal.at(off);
    const y = d.getFullYear();
    const m = d.getMonth() + n;
    const last = offsetOf(y, m + 1, 1) - offsetOf(y, m, 1);
    return offsetOf(y, m, Math.min(d.getDate(), last));
  };
  // Leases start on the 1st to the 28th, so the rent day is the same every month.
  const endBefore = (off) => { while (dom(off + 1) > 28) off--; return off; };
  const startFrom = (off) => { while (dom(off) > 28) off++; return off; };

  // The firm's own staff, who do the routine inspections.
  const staff = [0, 1].map(() => { const p = person(r); return `${p['First Name']} ${p['Last Name']}`; });

  // Contractors.
  const trades = new Map();
  for (const [name, trade] of CONTRACTORS) {
    const p = person(r, { domain: `${slug(name)}.example` });
    const id = add('Contractors', { Name: name, Trade: trade, Contact: `${p['First Name']} ${p['Last Name']}`, Phone: phoneFor(r, 'United Kingdom'), 'E-mail': `jobs@${slug(name)}.example`, Notes: r.chance(0.3) ? 'Call-out fee waived for jobs booked a day ahead.' : null });
    if (!trades.has(trade)) trades.set(trade, []);
    trades.get(trade).push({ id, name });
  }
  const contractorFor = (trade) => r.pick(trades.get(trade));
  const gasFirm = trades.get('Plumbing and heating')[0];
  const sparks = trades.get('Electrical')[0];

  // Buildings and their units.
  const props = [];
  const units = [];
  const streets = ['LS6', 'LS4', 'LS7', 'LS11', 'LS12', 'LS8'];
  for (const [name, kind, street, gas, list] of BUILDINGS) {
    const recent = kind === 'Workshop yard';
    const boughtAt = recent ? -r.int(330, 400) : -r.int(4 * 365, 17 * 365);
    const annual = list.reduce((n, u) => n + u[3] * 12, 0);
    const value = Math.round(annual / (0.064 + r.next() * 0.02) / 5000) * 5000;
    const years = -boughtAt / 365;
    const paid = Math.round(value / Math.pow(1.045, years) / 1000) * 1000;
    const p = { name, kind, recent, boughtAt, units: [] };
    p.id = add('Properties', {
      Name: name, Kind: kind, Address: street, City: 'Leeds', Postcode: `${r.pick(streets)} ${r.int(1, 9)}${r.pick(['AB', 'DE', 'FG', 'HJ', 'NP', 'QR', 'TU'])}`,
      Bought: cal.day(boughtAt), 'Purchase Price': paid, 'Current Value': value, 'Valued On': cal.day(-r.int(30, 300)),
      Notes: MORTGAGED.includes(name) ? 'Bought with a buy-to-let loan; the interest is paid monthly.' : recent ? 'Bought empty and refitted before letting.' : null,
    });
    for (const [uname, ukind, beds, rent] of list) {
      const commercial = ukind === 'Shop' || ukind === 'Workshop';
      const u = { prop: p, name: uname, kind: ukind, beds, market: rent, commercial, gas: gas && !commercial, label: `${uname}, ${name}`, leases: [] };
      u.id = add('Units', { Property: p.id, Name: uname, Kind: ukind, Bedrooms: beds, 'Market Rent': rent, 'Has Gas': u.gas, Furnished: !commercial && ukind === 'Flat' && r.chance(0.3), Status: 'Let' });
      units.push(u);
      p.units.push(u);
    }
    props.push(p);
  }

  // Which units stand empty today, and which leases end in the next two
  // months: chosen, so every day's file has some of each.
  const homes = units.filter((u) => !u.commercial && !u.prop.recent);
  const [emptyA, emptyB, emptyC, ...rest] = r.some(homes, homes.length);
  const emptyShop = r.pick(units.filter((u) => u.kind === 'Shop'));
  emptyA.vacant = true; emptyB.vacant = true; emptyC.vacant = true; emptyShop.vacant = true;
  emptyA.upcoming = true; // signed up, moving in soon
  for (const u of rest.slice(0, 5)) u.endingSoon = true;
  for (const u of r.some(rest.slice(5), 3)) u.struggling = true;

  // Each unit's leases, laid out backwards from today.
  const tenants = [];
  const newTenant = (u) => { const t = { commercial: u.commercial, profile: r.weighted([['prompt', 7], ['late', 3]]), leases: [] }; tenants.push(t); return t; };
  const rentAt = (u, start) => round5(u.market * Math.pow(u.commercial ? 1.02 : 1.035, Math.min(0, start) / 365));
  const lease = (u, t, S, E, renewal) => {
    const l = { unit: u, tenant: t, S, E, rent: rentAt(u, S), renewal, term: 0 };
    u.leases.unshift(l);
    t.leases.push(l);
    return l;
  };
  for (const u of units) {
    if (u.prop.recent) {
      // Let once the refit was done.
      const S = startFrom(u.prop.boughtAt + r.int(55, 110));
      lease(u, newTenant(u), S, addMonths(S, 60) - 1, false);
      continue;
    }
    let E;
    if (u.vacant) {
      E = endBefore(u.commercial ? -r.int(60, 140) : -r.int(8, 45));
      u.vacantSince = E + 1;
      if (u.upcoming) { const S = startFrom(r.int(4, 18)); lease(u, newTenant(u), S, addMonths(S, 12) - 1, false); }
    } else if (u.endingSoon) E = endBefore(r.int(8, 58));
    else E = endBefore(u.commercial ? r.int(120, 1400) : r.int(62, 358));
    let t = newTenant(u);
    for (let first = !u.vacant; ; first = false) {
      let term = u.commercial ? r.pick([36, 60]) : r.chance(0.15) ? 6 : 12;
      // The lease running today began on or before today.
      if (first && addMonths(E + 1, -term) > 0) term = u.commercial ? 60 : 12;
      const S = addMonths(E + 1, -term);
      const l = lease(u, t, S, E, false);
      if (u.struggling && l === u.leases[u.leases.length - 1]) t.struggling = true;
      if (S <= -days) break;
      if (!u.commercial && r.chance(0.55)) { l.renewal = true; E = S - 1; continue; }
      E = endBefore(S - (u.commercial ? r.int(30, 120) : r.int(7, 40)) - 1);
      t = newTenant(u);
    }
    // A struggling tenant must have been in long enough to fall behind.
    const now = u.leases.find((l) => l.S <= 0 && l.E >= 0);
    if (now && now.tenant.struggling && now.S > -100 && !now.renewal) now.tenant.struggling = false;
  }

  // Tenants and leases, oldest first.
  const all = units.flatMap((u) => u.leases).sort((a, b) => a.S - b.S || a.unit.id - b.unit.id);
  const relations = ['mother', 'father', 'sister', 'brother', 'partner', 'friend', 'daughter', 'son'];
  let business = 0;
  for (const l of all) {
    const t = l.tenant;
    if (!t.id) {
      const p = person(r, { domain: 'mail.example' });
      const kin = person(r);
      const name = t.commercial ? BUSINESSES[business++ % BUSINESSES.length] : null;
      t.id = add('Tenants', {
        'First Name': p['First Name'], 'Last Name': p['Last Name'], 'Business Name': name, Phone: phoneFor(r, 'United Kingdom'),
        'E-mail': name ? `hello@${slug(name)}.example` : p['E-mail'],
        'Emergency Contact': `${kin['First Name']} ${kin['Last Name']} (${t.commercial ? 'business partner' : r.pick(relations)})`, 'Emergency Phone': phoneFor(r, 'United Kingdom'),
        'References Checked': l.S > -days ? true : r.chance(0.85), 'ID Checked': l.S > 0 ? false : true,
        Notes: t.commercial ? null : r.chance(0.12) ? r.pick(['Has a small dog, agreed in writing.', 'Guarantor: a parent, details on file.', 'Works nights; call after noon.']) : null,
      });
      t.display = name || `${p['First Name']} ${p['Last Name']}`;
      t.method = t.commercial ? 'Bank transfer' : r.weighted([['Standing order', 7], ['Bank transfer', 2], ['Card', 0.7], ['Cash', 0.3]]);
    }
    const prev = l.renewal ? t.last : null;
    l.deposit = prev ? prev.deposit : l.unit.commercial ? l.rent * 3 : Math.floor((l.rent * 12) / 52 * 5);
    l.ref = prev ? prev.ref : l.unit.commercial || l.S > 0 || (l.S > -12 && r.chance(0.5)) ? null : `DEP-${cal.day(l.S).slice(0, 4)}-${String(r.int(10000, 99999))}`;
    const status = l.S > 0 ? 'Upcoming' : l.E < 0 ? 'Ended' : 'Active';
    l.id = add('Leases', {
      Unit: l.unit.id, Tenant: t.id, Type: l.unit.commercial ? 'Commercial' : 'Residential', Start: cal.day(l.S), End: cal.day(l.E),
      'Monthly Rent': l.rent, 'Rent Day': dom(l.S), 'Pays By': t.method, Deposit: l.deposit, 'Deposit Ref': l.ref, Renewal: l.renewal, Status: status,
      Notes: l.unit.commercial ? 'Tenant repairs the shop front and pays the business rates.' : null,
    });
    t.last = l;
  }

  // The ledger, collected and written in date order at the end.
  const ledger = [];
  const pay = (p, at, kind, description, out, who, method = 'Bank transfer') => {
    if (at < -days || at > 0) return;
    ledger.push({ at, values: { Property: p.id, Date: cal.day(at), Kind: kind, Description: description, 'Payee or Payer': who, Method: method, 'Money In': 0, 'Money Out': money(out) } });
  };

  // Rent: a charge each month of each lease, from when the records begin to a
  // month ahead, and the payments made against it.
  for (const l of all) {
    const t = l.tenant;
    let first = true;
    for (let k = 0; ; k++) {
      const due = addMonths(l.S, k);
      if (due > l.E || due > 31) break;
      if (due < -days) continue;
      const period = cal.day(due).slice(0, 7);
      const payments = [];
      if (due <= 0) {
        if (t.struggling && due >= -100) {
          // Fallen behind: part of the rent, some months none.
          const share = first ? 0.5 : r.pick([0, 0.5, 0.6, 0.75, 1]);
          first = false;
          const at = due + r.int(3, 14);
          if (share > 0 && at <= 0) payments.push([at, money(l.rent * share)]);
        } else {
          const lag = t.profile === 'late' && r.chance(0.6) ? r.int(3, 24) : t.method === 'Standing order' ? 0 : r.int(0, 3);
          const at = due + lag;
          if (at <= 0) {
            if (r.chance(0.04)) {
              // Paid in two halves.
              const half = money(l.rent / 2);
              payments.push([at, half]);
              const at2 = at + r.int(5, 12);
              if (at2 <= 0) payments.push([at2, money(l.rent - half)]);
            } else payments.push([at, l.rent]);
          }
        }
      }
      const paid = money(payments.reduce((n, [, a]) => n + a, 0));
      const charge = add('Rent Charges', { Lease: l.id, Period: period, 'Due Date': cal.day(due), Amount: l.rent, 'Amount Paid': paid });
      for (const [at, amount] of payments) {
        ledger.push({ at, values: { Property: l.unit.prop.id, Date: cal.day(at), Kind: 'Rent', Description: `Rent ${period}, ${l.unit.label}`, Charge: charge, 'Payee or Payer': t.display, Method: t.method, 'Money In': amount, 'Money Out': 0 } });
      }
    }
    // Finding a new tenant costs something; a renewal costs a little paperwork.
    const at = l.S - r.int(3, 12);
    if (l.unit.commercial && !l.renewal) pay(l.unit.prop, at, 'Agent fees', `Commercial letting fee, ${l.unit.label}`, l.rent * 12 * 0.1, 'Commercial letting agent');
    else if (!l.renewal) pay(l.unit.prop, at, 'Letting costs', `Advertising, referencing and inventory, ${l.unit.label}`, r.int(180, 290), 'Letting costs');
    else pay(l.unit.prop, at, 'Letting costs', `Renewal paperwork, ${l.unit.label}`, r.int(45, 75), 'Letting costs');
  }
  // When each unit is let, as [start, end] pairs, and the empty spells between.
  const letOn = (u, d) => u.leases.some((l) => l.S <= d && d <= l.E);
  const voids = [];
  for (const u of units) {
    for (let i = 0; i < u.leases.length; i++) {
      const a = u.leases[i].E + 1;
      const b = i + 1 < u.leases.length ? u.leases[i + 1].S - 1 : 0;
      if (a <= 0 && b >= a && b >= -days) voids.push({ u, a: Math.max(a, -days), b: Math.min(b, 0), last: i + 1 === u.leases.length || u.leases[i + 1].S > 0 });
    }
    if (u.prop.recent) voids.push({ u, a: Math.max(u.prop.boughtAt, -days), b: u.leases[0].S - 1, last: false, refit: true });
  }

  // Repairs: the jobs that come up between tenants, and when units stand empty.
  const jobs = [];
  for (const v of voids) {
    // Council tax and power while a unit is empty, a month at a time.
    for (let a = v.a; a <= v.b; a += 30) {
      const n = Math.min(30, v.b - a + 1);
      pay(v.u.prop, a + n - 1, 'Empty unit costs', `${v.u.commercial ? 'Business rates and power' : 'Council tax and power'} while empty, ${v.u.label}`, n * (v.u.commercial ? 14 : 6.2), v.u.commercial ? 'City council and energy supplier' : 'City council and energy supplier', 'Direct debit');
    }
    if (v.refit) {
      jobs.push({ u: v.u, at: v.a + 5, problem: 'Refit before first letting', trade: 'General repairs', priority: 'Normal', cost: r.int(4200, 7800), book: v.a + 12, done: v.b - r.int(3, 10), by: 'Landlord' });
    } else if (v.last && v.u.vacant) {
      // Empty now: work booked to get it ready to let.
      const problem = v.u.commercial ? 'Shop front and shutter repairs before reletting' : 'Redecorate before reletting';
      jobs.push({ u: v.u, at: Math.max(v.a, -days), problem, trade: v.u.commercial ? 'General repairs' : 'Decorating', priority: 'Normal', cost: r.int(900, 2400), book: r.int(1, 12), done: null, by: 'Landlord' });
    } else if (!v.u.commercial && v.b - v.a >= 12 && r.chance(0.45)) {
      jobs.push({ u: v.u, at: v.a, problem: 'Redecorate between tenants', trade: 'Decorating', priority: 'Normal', cost: r.int(650, 1900), book: v.a + r.int(2, 5), done: Math.min(v.b, v.a + r.int(7, 11)), by: 'Landlord' });
    }
  }

  // Inspections and safety checks: each unit's routine visit every six months
  // (a year for shops and workshops), gas every year where there is gas, the
  // wiring every five years. Spread evenly across the units, so some fall due
  // in any month.
  const checks = [];
  const fromInspection = new Map(); // day -> [{ u, problem }]
  const later = (d, item) => { if (!fromInspection.has(d)) fromInspection.set(d, []); fromInspection.get(d).push(item); };
  const homesAll = units.filter((u) => !u.commercial);
  const shopsAll = units.filter((u) => u.commercial);
  const gasUnits = units.filter((u) => u.gas);
  const schedule = (u, kind, first, months, doneAt) => {
    // A building bought lately was checked once, just before its first tenant.
    if (u.prop.recent && first < u.leases[0].S) first = u.leases[0].S - r.int(5, 15);
    let due = first;
    for (;;) {
      const done = due > 0 ? null : doneAt(due);
      if (done == null || done > 0) { checks.push({ u, kind, due, done: null }); break; }
      const c = { u, kind, due, done };
      checks.push(c);
      if (kind === 'Routine inspection') {
        c.inspector = r.pick(staff);
        c.outcome = r.weighted([['Good', 6], ['Fair: minor wear', 2.5], ['Action needed', 1.5]]);
        if (c.outcome === 'Action needed') later(done, { u, problem: r.pick(PROBLEMS.filter((p) => p[5].includes(u.commercial ? 's' : 'h') && p[2] !== 'Urgent')) });
        due = addMonths(done, months);
      } else if (kind === 'Gas safety') {
        c.inspector = gasFirm.name;
        c.outcome = r.chance(0.94) ? 'Pass' : 'Pass with advice';
        c.number = `GAS-${cal.day(done).slice(0, 4)}-${String(r.int(1000, 9999))}`;
        pay(u.prop, done, 'Safety checks', `Gas safety check, ${u.label}`, r.int(68, 88), gasFirm.name);
        due = addMonths(due, months);
      } else {
        c.inspector = sparks.name;
        c.outcome = r.chance(0.85) ? 'Satisfactory' : 'Unsatisfactory: remedial work';
        c.number = `ELEC-${cal.day(done).slice(0, 4)}-${String(r.int(1000, 9999))}`;
        pay(u.prop, done, 'Safety checks', `Electrical safety check, ${u.label}`, r.int(165, 240), sparks.name);
        if (c.outcome !== 'Satisfactory') later(done, { u, problem: ['Remedial electrical work after the safety check', 'Electrical', 'High', 180, 650, 'hs', 1] });
        due = addMonths(due, months);
      }
    }
  };
  homesAll.forEach((u, i) => schedule(u, 'Routine inspection', -days + Math.round((i * 182) / homesAll.length), 6, (due) => due + r.int(-6, 6)));
  shopsAll.forEach((u, i) => schedule(u, 'Routine inspection', -days + Math.round((i * 365) / shopsAll.length), 12, (due) => due + r.int(-6, 6)));
  gasUnits.forEach((u, i) => schedule(u, 'Gas safety', -days + Math.round((i * 365) / gasUnits.length), 12, (due) => due - r.int(0, 21)));
  units.forEach((u, i) => schedule(u, 'Electrical safety', -days + Math.round((i * 1826) / units.length), 60, (due) => due - r.int(0, 30)));
  // One routine visit put off: the tenant asked to rearrange.
  const putOff = checks.filter((c) => c.kind === 'Routine inspection' && !c.done && c.due > 0 && c.due < 60 && letOn(c.u, 0));
  if (putOff.length) { const c = r.pick(putOff); c.due = -r.int(3, 12); c.notes = 'Tenant asked to rearrange; new day to be agreed.'; }

  // Repairs reported day by day: more in winter, fewer when a unit is empty.
  for (let d = -days; d <= 0; d++) {
    const month = cal.at(d).getMonth();
    const winter = month >= 10 || month <= 1;
    for (const u of units) {
      if (d < u.prop.boughtAt) continue;
      const occupied = letOn(u, d);
      if (r.chance((occupied ? 1 / 135 : 1 / 400) * (winter ? 1.35 : 0.85))) {
        const p = r.weighted(PROBLEMS.filter((x) => x[5].includes(u.commercial ? 's' : 'h')).map((x) => [x, winter ? x[6] : 1]));
        jobs.push({ u, at: d, problem: p[0], trade: p[1], priority: p[2], cost: r.int(p[3], p[4]), by: occupied ? 'Tenant' : 'Landlord' });
      }
    }
    for (const { u, problem: p } of fromInspection.get(d) || []) jobs.push({ u, at: d, problem: p[0], trade: p[1], priority: p[2], cost: r.int(p[3], p[4]), by: 'Inspection' });
  }

  // Each job: booked with a contractor of its trade, as fast as its priority
  // asks, done on the day or after a second visit; a few are cancelled.
  jobs.sort((a, b) => a.at - b.at);
  const lead = { Urgent: [0, 1], High: [1, 3], Normal: [3, 12], Low: [5, 25] };
  for (const j of jobs) {
    let status;
    let contractor = null;
    let book = j.book ?? null;
    let done = j.done ?? null;
    if (j.book == null) {
      if (j.at >= -1 && r.chance(0.6)) status = 'Reported';
      else if (r.chance(0.04)) status = 'Cancelled';
      else {
        const [lo, hi] = lead[j.priority];
        book = cal.workday(j.at + r.int(lo, hi));
        done = book + (r.chance(0.15) ? r.int(1, 7) : 0);
      }
    }
    if (!status) {
      contractor = contractorFor(j.trade);
      status = done != null && done <= 0 ? 'Done' : 'Booked';
      if (status === 'Booked') done = null;
    }
    const id = add('Maintenance Requests', {
      Unit: j.u.id, Reported: cal.day(j.at), 'Reported By': j.by, Problem: j.problem, Priority: j.priority, Status: status,
      Contractor: contractor?.id ?? null, 'Booked For': book != null && status !== 'Cancelled' ? cal.day(book) : null,
      Completed: status === 'Done' ? cal.day(done) : null, Cost: status === 'Done' ? j.cost : null,
      Notes: status === 'Cancelled' ? r.pick(['Tenant fixed it themselves.', 'Reported twice; see the other job.', 'Came right on its own.']) : status === 'Booked' && j.u.vacant ? 'Unit empty; keys in the office.' : null,
    });
    if (status === 'Done') pay(j.u.prop, done, 'Repairs', `Job ${id}: ${j.problem}, ${j.u.label}`, j.cost, contractor.name);
  }

  // Inspections, in the order they fall due.
  checks.sort((a, b) => a.due - b.due || a.u.id - b.u.id);
  for (const c of checks) {
    add('Inspections', {
      Unit: c.u.id, Kind: c.kind, Due: cal.day(c.due), Done: c.done != null ? cal.day(c.done) : null, Inspector: c.inspector ?? null,
      Outcome: c.outcome ?? null, 'Certificate Number': c.number ?? null,
      Notes: c.notes ?? (c.done == null && c.u.vacant ? 'Unit empty; keys in the office.' : null),
    });
  }

  // The buildings' running costs: insurance once a year, loan interest each
  // month, cleaning and lighting of shared areas, grounds in the growing months.
  for (const p of props) {
    const units = p.units.length;
    const premium = p.units[0].commercial || p.kind === 'Shops with flats' ? 1150 + units * 120 : units === 1 ? r.int(340, 520) : 260 * units;
    const from = Math.max(-days, p.boughtAt);
    for (let at = from + r.int(0, 364); at <= 0; at = addMonths(at, 12)) pay(p, at, 'Insurance', `Buildings and landlord insurance, ${p.name}`, premium * (1 + r.next() * 0.08), 'Landlord insurer', 'Direct debit');
    const loan = MORTGAGED.includes(p.name) ? Math.round(p.units.reduce((n, u) => n + u.market, 0) * 0.32) : 0;
    for (let d = -days; d <= 0; d++) {
      if (dom(d) !== 1) continue;
      const month = cal.at(d).getMonth();
      if (loan && d >= p.boughtAt) pay(p, d, 'Loan interest', `Buy-to-let loan interest, ${p.name}`, loan, 'Lender', 'Direct debit');
      if (SHARED_AREAS.includes(p.name)) {
        pay(p, d + 2, 'Shared areas', `Cleaning of halls and stairs, ${p.name}`, 140 + units * 6, 'Cleaning contractor');
        if (month % 3 === 0) pay(p, d + 9, 'Shared areas', `Lighting and power for shared areas, ${p.name}`, 90 + units * 11 + r.int(0, 40), 'Energy supplier', 'Direct debit');
        if (month >= 3 && month <= 9) pay(p, d + 14, 'Shared areas', `Grounds and bins, ${p.name}`, 75 + r.int(0, 30), 'Greenacre Grounds');
      }
    }
  }

  // The ledger, in date order.
  ledger.sort((a, b) => a.at - b.at);
  for (const x of ledger) add('Transactions', x.values);

  // Each unit as it stands today.
  for (const u of units) {
    const let_ = letOn(u, 0);
    base.update('Units', u.id, { Status: let_ ? 'Let' : 'Vacant', 'Vacant Since': let_ ? null : cal.day(u.vacantSince ?? 0), Notes: u.upcoming ? 'New tenant signed; moving in soon.' : null });
  }
  return counts;
}
