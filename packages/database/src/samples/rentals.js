// The lettings business: Hollybrook Lettings, an invented landlord that owns
// and manages a dozen buildings in one city (houses, two blocks of flats, two
// converted houses, a parade of shops with flats above and a small yard of
// workshops), with thirty-eight units let to tenants. The buildings and their
// units, the tenants, each lease, the deposits taken, a rent invoice for every
// month of every lease, the reminders sent when rent is late, the
// contractors, repair requests, inspections and safety checks, and one ledger
// of every sum of money in or out for each building.
//
// Its records are eighteen months of letting. Each unit's leases are laid
// out first, counted back from today: who is in now and when they leave,
// who renewed, how long a unit stood empty between tenants. Then the months
// are run a day at a time: rent is invoiced two weeks ahead and paid on
// time, late or in part, and a reminder goes out when it is a week late;
// things break, more in winter, and contractors are booked to mend them;
// inspections and safety checks come round, each with a notice to the
// tenant; the building's bills are paid. The questions a landlord asks are
// its saved queries; its reports and print templates print them.

import { address, BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, money, person, phoneFor, PLACES } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const TENANT = 'Nz([Business Name], [First Name] & " " & [Last Name])';
const UNIT = '[Units].[Name] & ", " & [Properties].[Name]';
const OWED = 'Round([Amount] - Nz([Amount Paid]), 2)';
// The postal address of a unit, with its postcode.
const POSTAL = '[Units].[Postal Address] & ", " & [Properties].[Postcode]';
const NET = 'Round(Nz([Money In]) - Nz([Money Out]), 2)';
const LAST_YEAR = { expr: 'DateDiff("d", [Date], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] };
// The business's own details at the head of every document.
const LETTERHEAD = { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Website'] };

export const spec = {
  // Where a value below nought is right: a check already overdue, a building
  // or a month that cost more than it took in.
  negative: {
    'Inspections due': ['Days Left'], 'Safety certificates due': ['Days Left'],
    'Income and expenditure by property': ['Net'], 'Money in and out by month': ['Net'],
  },
  tables: [
    BUSINESS_TABLE,
    { name: 'Properties', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Kind', type: 'text' },
      { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' },
      { name: 'Bought', type: 'date' }, { name: 'Purchase Price', type: 'currency' }, { name: 'Current Value', type: 'currency' },
      { name: 'Valued On', type: 'date' }, { name: 'VAT on Rent', type: 'yesno' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Units', fields: [
      { name: 'Property', ...ref('Properties'), required: true }, { name: 'Name', type: 'text', required: true }, { name: 'Kind', type: 'text' },
      { name: 'Postal Address', type: 'text' }, { name: 'Bedrooms', type: 'integer', default: 0 }, { name: 'Market Rent', type: 'currency' },
      { name: 'Has Gas', type: 'yesno' }, { name: 'Furnished', type: 'yesno' }, { name: 'Rent Includes', type: 'text' },
      { name: 'Status', type: 'text', default: 'Vacant' }, { name: 'Vacant Since', type: 'date' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Tenants', fields: [
      { name: 'Account', type: 'text', unique: true }, { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
      { name: 'Business Name', type: 'text' }, { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Emergency Contact', type: 'text' },
      { name: 'Emergency Phone', type: 'text' }, { name: 'References Checked', type: 'yesno' }, { name: 'ID Checked', type: 'yesno' },
      { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Leases', fields: [
      { name: 'Unit', ...ref('Units'), required: true }, { name: 'Tenant', ...ref('Tenants'), required: true }, { name: 'Type', type: 'text', default: 'Residential' },
      { name: 'Signed On', type: 'date' }, { name: 'Start', type: 'date' }, { name: 'End', type: 'date' }, { name: 'Monthly Rent', type: 'currency' },
      { name: 'Tax Rate', type: 'number', default: 0 }, { name: 'Rent Day', type: 'integer', default: 1 }, { name: 'Pays By', type: 'text', default: 'Standing order' },
      { name: 'Deposit', type: 'currency' }, { name: 'Deposit Ref', type: 'text' }, { name: 'Renewal', type: 'yesno' },
      { name: 'Pets', type: 'text' }, { name: 'Permitted Use', type: 'text' }, { name: 'Rent Review', type: 'date' },
      { name: 'Status', type: 'text', default: 'Active' }, { name: 'Notes', type: 'longtext' },
    ] },
    // Lease before Tenant: a query that has Rent Charges, Leases and Tenants joins the tenant through the lease.
    { name: 'Rent Charges', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Lease', ...ref('Leases'), required: true }, { name: 'Tenant', ...ref('Tenants') },
      { name: 'Period', type: 'text' }, { name: 'Invoice Date', type: 'date' }, { name: 'Due Date', type: 'date' }, { name: 'Period End', type: 'date' },
      { name: 'Net', type: 'currency' }, { name: 'Tax Rate', type: 'number', default: 0 }, { name: 'Tax', type: 'currency', default: 0 },
      { name: 'Amount', type: 'currency' }, { name: 'Amount Paid', type: 'currency', default: 0 }, { name: 'Last Paid On', type: 'date' },
    ] },
    { name: 'Deposits', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Lease', ...ref('Leases'), required: true }, { name: 'Received On', type: 'date' },
      { name: 'Amount', type: 'currency' }, { name: 'Paid By', type: 'text', default: 'Bank transfer' }, { name: 'Held By', type: 'text' },
      { name: 'Returned On', type: 'date' }, { name: 'Amount Returned', type: 'currency' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Rent Reminders', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Lease', ...ref('Leases'), required: true }, { name: 'Sent On', type: 'date' },
      { name: 'Stage', type: 'text', default: 'First reminder' }, { name: 'Amount Owed', type: 'currency' }, { name: 'Months Owing', type: 'integer' },
      { name: 'Since', type: 'date' }, { name: 'Pay By', type: 'date' }, { name: 'Next Step', type: 'longtext' }, { name: 'Outcome', type: 'text', default: 'Waiting' },
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
    // Unit before Lease: the inspection queries join a unit through the inspection's own Unit.
    { name: 'Inspections', fields: [
      { name: 'Unit', ...ref('Units'), required: true }, { name: 'Lease', ...ref('Leases') }, { name: 'Kind', type: 'text' }, { name: 'Due', type: 'date' },
      { name: 'Visit Time', type: 'text' }, { name: 'Notice Sent', type: 'date' },
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
      { expr: TENANT, as: 'Tenant', sort: 'asc' }, { expr: POSTAL, as: 'Address' }, { table: 'Rent Charges', field: 'Number', as: 'Invoice' },
      { table: 'Rent Charges', field: 'Period' }, { table: 'Rent Charges', field: 'Due Date', sort: 'asc' }, { table: 'Rent Charges', field: 'Amount' },
      { table: 'Rent Charges', field: 'Amount Paid' }, { expr: OWED, as: 'Owed', criteria: ['> 0'] },
      { expr: 'DateDiff("d", [Due Date], Date())', as: 'Days Late', criteria: ['> 0'] },
    ] },
    'Rent due in the next 14 days': { tables: ['Rent Charges', 'Leases', 'Tenants', 'Units', 'Properties'], columns: [
      { table: 'Rent Charges', field: 'Due Date', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Due Date])', as: 'Days Until', criteria: ['Between 0 And 14'] },
      { table: 'Rent Charges', field: 'Number', as: 'Invoice' }, { expr: TENANT, as: 'Tenant' }, { expr: UNIT, as: 'Unit' }, { table: 'Leases', field: 'Pays By' },
      { table: 'Rent Charges', field: 'Amount' }, { expr: OWED, as: 'Owed', criteria: ['> 0'], show: false },
    ] },
    'Rent invoices raised': { tables: ['Rent Charges', 'Leases', 'Tenants', 'Units', 'Properties'], columns: [
      { table: 'Rent Charges', field: 'Number', as: 'Invoice', sort: 'asc' }, { table: 'Rent Charges', field: 'Invoice Date' }, { table: 'Rent Charges', field: 'Due Date' },
      { expr: TENANT, as: 'Tenant' }, { expr: UNIT, as: 'Unit' }, { table: 'Rent Charges', field: 'Net' }, { table: 'Rent Charges', field: 'Tax', as: 'VAT' },
      { table: 'Rent Charges', field: 'Amount' }, { table: 'Rent Charges', field: 'Amount Paid' },
      { expr: 'DateDiff("d", [Invoice Date], Date())', as: 'Days Ago', criteria: ['Between 0 And 31'], show: false },
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
      { table: 'Inspections', field: 'Kind' }, { expr: UNIT, as: 'Unit' }, { table: 'Inspections', field: 'Notice Sent' }, { table: 'Inspections', field: 'Notes' },
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
    'Arrears letters sent': { tables: ['Rent Reminders', 'Leases', 'Tenants', 'Units', 'Properties'], columns: [
      { table: 'Rent Reminders', field: 'Number', as: 'Letter' }, { table: 'Rent Reminders', field: 'Sent On', sort: 'desc' }, { expr: TENANT, as: 'Tenant' },
      { expr: UNIT, as: 'Unit' }, { table: 'Rent Reminders', field: 'Stage' }, { table: 'Rent Reminders', field: 'Amount Owed' },
      { table: 'Rent Reminders', field: 'Pay By' }, { table: 'Rent Reminders', field: 'Outcome' },
      { expr: 'DateDiff("d", [Sent On], Date())', as: 'Days Ago', criteria: ['Between 0 And 90'], show: false },
    ] },
    'Current tenants': { tables: ['Leases', 'Tenants', 'Units', 'Properties'], columns: [
      { table: 'Properties', field: 'Name', as: 'Property', sort: 'asc' }, { table: 'Units', field: 'Name', as: 'Unit', sort: 'asc' }, { expr: TENANT, as: 'Tenant' },
      { table: 'Tenants', field: 'Account' }, { table: 'Tenants', field: 'Phone' }, { table: 'Tenants', field: 'E-mail' }, { table: 'Tenants', field: 'Emergency Contact' },
      { table: 'Tenants', field: 'Emergency Phone' }, { table: 'Leases', field: 'Status', criteria: ['Active'], show: false },
    ] },
    'Income and expenditure by property': { tables: ['Transactions', 'Properties'], columns: [
      { table: 'Properties', field: 'Name', as: 'Property', total: 'group', sort: 'asc' }, { table: 'Transactions', field: 'Money In', as: 'Money In', total: 'sum' },
      { table: 'Transactions', field: 'Money Out', as: 'Money Out', total: 'sum' }, { expr: NET, as: 'Net', total: 'sum' }, LAST_YEAR,
    ] },
    'Costs by kind': { tables: ['Transactions'], columns: [
      { table: 'Transactions', field: 'Kind', total: 'group', criteria: ['<> "VAT"'] }, { table: 'Transactions', field: 'ID', as: 'Payments', total: 'count' },
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
    { name: 'Arrears letters', title: 'Rent arrears', query: 'Unpaid rent', group: 'Tenant', columns: ['Address', 'Invoice', 'Period', { field: 'Due Date', type: 'date' }, { field: 'Days Late', type: 'integer' }, { field: 'Amount', type: 'currency' }, { field: 'Amount Paid', type: 'currency' }, { field: 'Owed', type: 'currency' }], totals: [{ field: 'Owed', fn: 'sum' }] },
    { name: 'Repairs for contractors', title: 'Jobs booked', query: 'Repairs booked', group: 'Contractor', columns: ['Job', { field: 'Booked For', type: 'date' }, 'Priority', 'Address', 'Problem'] },
    { name: 'Income and expenditure', title: 'Income and expenditure, last 12 months', query: 'Income and expenditure by property', columns: ['Property', { field: 'Money In', type: 'currency' }, { field: 'Money Out', type: 'currency' }, { field: 'Net', type: 'currency' }], totals: [{ field: 'Money In', fn: 'sum' }, { field: 'Money Out', fn: 'sum' }, { field: 'Net', fn: 'sum' }] },
    { name: 'Leases ending', title: 'Leases ending in the next 60 days', query: 'Leases ending in the next 60 days', columns: ['Property', 'Unit', 'Tenant', 'Phone', { field: 'End', type: 'date' }, { field: 'Days Left', type: 'integer' }, { field: 'Monthly Rent', type: 'currency' }, { field: 'Market Rent', type: 'currency' }], totals: [{ field: 'Monthly Rent', fn: 'sum' }] },
    { name: 'Empty units', title: 'Vacant units and rent lost', query: 'Vacant units', columns: ['Property', 'Unit', 'Kind', { field: 'Market Rent', type: 'currency' }, { field: 'Vacant Since', type: 'date' }, { field: 'Days Empty', type: 'integer' }, { field: 'Rent Lost', type: 'currency' }], totals: [{ field: 'Rent Lost', fn: 'sum' }, { field: 'Market Rent', fn: 'sum' }] },
    { name: 'Checks due', title: 'Inspections and safety checks due', query: 'Inspections due', group: 'Kind', columns: [{ field: 'Due', type: 'date' }, { field: 'Days Left', type: 'integer' }, 'Unit', { field: 'Notice Sent', type: 'date' }, 'Notes'] },
    { name: 'Monthly accounts', query: 'Money in and out by month', columns: ['Month', { field: 'Money In', type: 'currency' }, { field: 'Money Out', type: 'currency' }, { field: 'Net', type: 'currency' }], totals: [{ field: 'Money In', fn: 'sum' }, { field: 'Money Out', fn: 'sum' }, { field: 'Net', fn: 'sum' }] },
    { name: 'Rent invoices this month', title: 'Rent invoices raised in the last 31 days', query: 'Rent invoices raised', columns: ['Invoice', { field: 'Invoice Date', type: 'date' }, { field: 'Due Date', type: 'date' }, 'Tenant', 'Unit', { field: 'Net', type: 'currency' }, { field: 'VAT', type: 'currency' }, { field: 'Amount', type: 'currency' }], totals: [{ field: 'Net', fn: 'sum' }, { field: 'VAT', fn: 'sum' }, { field: 'Amount', fn: 'sum' }] },
  ],
  // The print templates: a document for each record of their table, the business's details at the head.
  documents: [
    { name: 'Rent invoice', table: 'Rent Charges', paper: 'a4', letterhead: { ...LETTERHEAD, fields: [...LETTERHEAD.fields, 'Tax Number'] }, blocks: [
      { type: 'heading', text: 'Rent invoice {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Tenant', label: 'Invoice to' }, { field: 'Number', label: 'Invoice number' },
        { field: 'Tenant.Business Name', label: 'Trading as' }, { field: 'Invoice Date', label: 'Date' },
        { field: 'Tenant.Address', label: 'Address' }, { field: 'Due Date', label: 'Rent due' },
        { field: 'Tenant.City', label: 'City' }, { field: 'Period End', label: 'Rent paid up to' },
        { field: 'Tenant.Postcode', label: 'Postcode' }, { field: 'Tenant.Account', label: 'Rent account' },
      ] },
      { type: 'text', text: 'Rent for {Lease.Unit.Postal Address}, {Lease.Unit.Property.City} {Lease.Unit.Property.Postcode}, for the month from {Due Date} to {Period End}, due in advance. You pay by: {Lease.Pays By}.' },
      { type: 'text', text: 'Payments received against this invoice:', small: true },
      { type: 'lines', table: 'Transactions', link: 'Charge', sort: 'Date', columns: [
        { field: 'Date', label: 'Received' }, { field: 'Method' }, { field: 'Money In', label: 'Amount', type: 'currency' },
      ] },
      { type: 'totals', rows: [
        { label: 'Rent', expr: '[Net]' }, { label: 'VAT', expr: 'Nz([Tax])' }, { label: 'Total for the month', expr: '[Amount]' },
        { label: 'Already paid', expr: 'Nz([Amount Paid])' }, { label: 'Balance to pay', expr: 'Round([Amount] - Nz([Amount Paid]), 2)', strong: true },
      ] },
      { type: 'text', text: '{Business.Payment Terms} Bank transfer to {Business.Bank Details}, quoting {Number}. A standing order pays this on the due date by itself: there is nothing more to do.', small: true },
      { type: 'text', text: 'VAT registration {Business.Tax Number}. VAT is added only to the rent of a shop or workshop in a building the landlord has opted to tax; rent on a home carries none.', small: true },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    { name: 'Rent statement', table: 'Tenants', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Rent statement: {First Name} {Last Name}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Business Name', label: 'Trading as' }, { field: 'Account', label: 'Rent account' },
        { field: 'Address', label: 'Address' }, { field: 'Phone', label: 'Phone' },
        { field: 'City', label: 'City' }, { field: 'E-mail', label: 'E-mail' }, { field: 'Postcode', label: 'Postcode' },
      ] },
      { type: 'lines', table: 'Rent Charges', link: 'Tenant', sort: 'Due Date', columns: [
        { field: 'Number', label: 'Invoice' }, { field: 'Period', label: 'Month' }, { field: 'Due Date', label: 'Due' },
        { field: 'Amount', label: 'Charged', type: 'currency' }, { field: 'Amount Paid', label: 'Paid', type: 'currency' }, { field: 'Last Paid On', label: 'Paid on' },
        { expr: 'Round([Amount] - Nz([Amount Paid]), 2)', as: 'Owing', type: 'currency' },
      ] },
      { type: 'totals', rows: [
        { label: 'Rent charged', expr: '[Amount]' }, { label: 'Rent paid', expr: '[Amount Paid]' }, { label: 'Balance owing', expr: '[Owing]', strong: true },
      ] },
      { type: 'text', text: 'Please pay any balance owing to {Business.Bank Details}, quoting your rent account {Account}. If you have paid since this statement was made, thank you. If anything here looks wrong, call us on {Business.Phone} and we will look into it.', small: true },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    { name: 'Tenancy agreement', table: 'Leases', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Tenancy agreement for a home' },
      { type: 'fields', columns: 2, items: [
        { field: 'Tenant', label: 'Tenant' }, { field: 'Signed On', label: 'Date of this agreement' },
        { field: 'Unit.Postal Address', label: 'The property' }, { field: 'Start', label: 'Tenancy begins' },
        { field: 'Unit.Property.City', label: 'City' }, { field: 'End', label: 'Fixed term ends' },
        { field: 'Unit.Property.Postcode', label: 'Postcode' }, { field: 'Rent Day', label: 'Rent day each month' },
        { field: 'Unit.Bedrooms', label: 'Bedrooms' }, { field: 'Pays By', label: 'Rent paid by' },
        { field: 'Tenant.Account', label: 'Rent account' }, { field: 'Deposit Ref', label: 'Deposit protection reference' },
      ] },
      { type: 'totals', rows: [{ label: 'Rent each month', expr: '[Monthly Rent]' }, { label: 'Deposit', expr: '[Deposit]' }] },
      { type: 'text', text: 'This agreement is made on {Signed On} between {Business.Name} of {Business.Address}, {Business.City} {Business.Postcode} ("the Landlord") and {Tenant} ("the Tenant"). The Landlord lets the property named above ("the Property") to the Tenant as a home, on the terms below.' },
      { type: 'text', text: '1. The property. The Property is let together with the use of any shared entrance, halls, stairs, paths, garden and bin store that serve it, and with the contents listed in the inventory. It is to be lived in by the Tenant and the members of the Tenant\'s household named on the application, and no one else without the Landlord\'s written consent.' },
      { type: 'text', text: '2. The term. The tenancy begins on {Start} and its fixed term ends on {End}. If neither side has ended it by then, it carries on from month to month on the same terms until it is ended as clause 11 says.' },
      { type: 'text', text: '3. The rent. The rent is {Monthly Rent} a month, paid in advance on day {Rent Day} of each month, the first payment on {Start}. The Tenant pays it by the method shown above into the Landlord\'s account, {Business.Bank Details}, quoting the rent account {Tenant.Account}. The Landlord sends a rent invoice before each rent day. The rent does not go up during the fixed term. After it, the Landlord may raise it no more than once a year, by giving at least two months\' notice in writing.' },
      { type: 'text', text: '4. What the rent includes. The rent includes: {Unit.Rent Includes}. The Tenant pays for everything else used at the Property, among them council tax, gas, electricity, water and sewerage, telephone and broadband.' },
      { type: 'text', text: '5. The deposit. The Tenant pays a deposit of {Deposit} before the tenancy begins. Within 30 days of receiving it, the Landlord protects it in a government-approved tenancy deposit scheme and gives the Tenant the scheme\'s details; its reference is shown above once it has been given. At the end of the tenancy the deposit is returned within ten days of both sides agreeing what, if anything, is kept back for unpaid rent, damage beyond fair wear and tear, or cleaning needed to bring the Property back to its condition at the start. If they cannot agree, either of them may ask the scheme to decide.' },
      { type: 'text', text: '6. The Tenant agrees: to pay the rent on time; to keep the Property clean and in good order, and to heat and air it so that damp and mould do not build up; to tell the Landlord promptly of any repair needed or damage done; to test the smoke alarms once a month; not to make alterations, redecorate or change the locks without the Landlord\'s written consent; not to cause a nuisance or annoyance to the neighbours; not to sublet the Property or take in a lodger without the Landlord\'s written consent; and to give back all the keys at the end.' },
      { type: 'text', text: '7. The Landlord agrees: to let the Tenant live in the Property in peace, without needless interruption; to keep the structure and outside in repair, including the roof, walls, windows, gutters and drains; to keep the supplies of water, gas and electricity, and the heating and hot water, in repair and working order; to have any gas appliances checked by a registered engineer every year and the wiring every five years, and to give the Tenant a copy of each certificate; to fit smoke alarms, and carbon monoxide alarms where they are needed; and to insure the building.' },
      { type: 'text', text: '8. Repairs and access. The Tenant reports repairs to the Landlord on {Business.Phone} or at {Business.E-mail}, and the Landlord arranges them in a reasonable time, the most urgent first. The Landlord, or a contractor sent by the Landlord, may come in at a reasonable time of day to inspect the Property or carry out repairs after giving the Tenant at least 24 hours\' notice in writing. In an emergency, such as a gas leak, a flood or a fire, they may come in at once.' },
      { type: 'text', text: '9. Pets and smoking. Pets agreed for this tenancy: {Pets}. Any other pet needs the Landlord\'s written consent first, which will not be refused without good reason. No one may smoke or vape inside the Property.' },
      { type: 'text', text: '10. The inventory. An inventory and record of condition, with photographs, is made at the start of the tenancy and signed by both sides. At the end, the Property is compared with it, allowing for fair wear and tear.' },
      { type: 'text', text: '11. Ending the tenancy. The Tenant may end the tenancy at the end of the fixed term, or at any time after it, by giving at least one month\'s notice in writing, ending the day before a rent day. The Landlord may end it only on the grounds, and with the notice, that the law allows, and always in writing. Before the fixed term ends, the Landlord will ask the Tenant whether they would like to renew.' },
      { type: 'text', text: '12. Notices. Notices to the Landlord go to {Business.Address}, {Business.City} {Business.Postcode} or {Business.E-mail}. Notices to the Tenant go to the Property, or to {Tenant.E-mail}.' },
      { type: 'text', text: 'Each side keeps a signed copy of this agreement and of the inventory.', small: true },
      { type: 'spacer', size: 24 },
      { type: 'signatures', labels: ['Signed for the Landlord', 'Date', 'Signed by the Tenant', 'Date'] },
    ] },
    { name: 'Commercial lease summary', table: 'Leases', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Commercial lease: summary of terms' },
      { type: 'fields', columns: 2, items: [
        { field: 'Tenant.Business Name', label: 'Tenant' }, { field: 'Signed On', label: 'Date of the lease' },
        { field: 'Tenant', label: 'Signed for the tenant by' }, { field: 'Start', label: 'Term begins' },
        { field: 'Unit.Postal Address', label: 'The premises' }, { field: 'End', label: 'Term ends' },
        { field: 'Unit.Property.City', label: 'City' }, { field: 'Rent Review', label: 'Rent review' },
        { field: 'Unit.Property.Postcode', label: 'Postcode' }, { field: 'Rent Day', label: 'Rent day each month' },
        { field: 'Permitted Use', label: 'Permitted use' }, { field: 'Tenant.Account', label: 'Rent account' },
      ] },
      { type: 'totals', rows: [
        { label: 'Rent each month', expr: '[Monthly Rent]' }, { label: 'VAT on the rent', expr: 'Round([Monthly Rent] * Nz([Tax Rate]), 2)' },
        { label: 'Each month with VAT', expr: '[Rent each month] + [VAT on the rent]', strong: true }, { label: 'Rent for a year', expr: '[Monthly Rent] * 12' },
        { label: 'Deposit held', expr: '[Deposit]' },
      ] },
      { type: 'text', text: 'This summary sets out the main terms agreed between {Business.Name} ("the Landlord") and {Tenant.Business Name} ("the Tenant") for the lease of the premises named above. It goes with the lease and does not replace it: where the two differ, the lease is what counts.' },
      { type: 'text', text: '1. The premises. The unit named above, with the right to use the shared paths, service yard and bin store in common with the other occupiers.' },
      { type: 'text', text: '2. Use. The premises may be used for the permitted use shown above and for nothing else without the Landlord\'s written consent, which will not be refused without good reason.' },
      { type: 'text', text: '3. Term. The lease runs from {Start} to {End}. The Tenant has no right to end it early except as the lease says.' },
      { type: 'text', text: '4. Rent. The rent is {Monthly Rent} a month, paid in advance by bank transfer on day {Rent Day} of each month, quoting the rent account. VAT is added at the rate shown above when the Landlord has opted to tax the building; otherwise none is added. Where a rent review date is shown, the rent is reviewed on that date to the rent the premises would then let for on the open market, and it does not go down at review.' },
      { type: 'text', text: '5. Deposit. The Tenant pays the deposit shown above, three months\' rent, before the term begins. The Landlord holds it until the lease ends and returns it within 28 days, less any rent or other sums then owed and the cost of any repairs the Tenant should have done.' },
      { type: 'text', text: '6. Outgoings. The Tenant pays the business rates, water, power and telephone for the premises. The Landlord insures the building; the premium is included in the rent.' },
      { type: 'text', text: '7. Repairs. The Tenant keeps the inside of the premises in good repair and decoration, together with the shop front or roller door, the doors and the glass, and the Tenant\'s own fittings. The Landlord keeps the roof, the structure, the outside walls and the shared parts in repair. Notes agreed for this lease: {Notes}' },
      { type: 'text', text: '8. Access. The Landlord may come in to inspect or repair after giving at least 48 hours\' notice, and at once in an emergency. The premises are inspected once a year.' },
      { type: 'text', text: '9. Alterations, signs and assignment. Alterations other than minor ones inside, and any sign outside, need the Landlord\'s written consent. The Tenant may not assign the lease or sublet without the Landlord\'s written consent, which will not be refused without good reason.' },
      { type: 'text', text: '10. The end of the lease. At the end the Tenant gives the premises back clean, in the repair the lease asks for, with the Tenant\'s own fittings and signs taken away if the Landlord asks. The Landlord will talk to the Tenant about a new lease at least six months before the term ends.' },
      { type: 'spacer', size: 24 },
      { type: 'signatures', labels: ['For the Landlord', 'Date', 'For the Tenant', 'Date'] },
    ] },
    { name: 'Deposit receipt', table: 'Deposits', paper: 'a5', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Deposit receipt {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Lease.Tenant', label: 'Received from' }, { field: 'Number', label: 'Receipt number' },
        { field: 'Lease.Tenant.Business Name', label: 'Trading as' }, { field: 'Received On', label: 'Received on' },
        { field: 'Lease.Unit.Postal Address', label: 'For the tenancy of' }, { field: 'Paid By', label: 'Paid by' },
        { field: 'Lease.Start', label: 'Tenancy begins' }, { field: 'Held By', label: 'Held by' },
        { field: 'Lease.Tenant.Account', label: 'Rent account' }, { field: 'Lease.Deposit Ref', label: 'Protection reference' },
      ] },
      { type: 'totals', rows: [{ label: 'Deposit received', expr: '[Amount]', strong: true }] },
      { type: 'text', text: 'This deposit is held as security for the tenant\'s side of the agreement. It is not rent and is not used to pay rent while the tenancy runs. The deposit on a home is protected in a tenancy deposit scheme within 30 days, and the scheme\'s reference is sent to the tenant; the deposit on a shop or workshop is held by the landlord until the lease ends. At the end it is returned, less anything owed, as the agreement says.', small: true },
      { type: 'spacer', size: 16 },
      { type: 'signatures', labels: ['Received for the Landlord'] },
    ] },
    { name: 'Rent arrears letter', table: 'Rent Reminders', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'fields', columns: 2, items: [
        { field: 'Lease.Tenant', label: 'To' }, { field: 'Sent On', label: 'Date' },
        { field: 'Lease.Tenant.Address', label: 'Address' }, { field: 'Number', label: 'Our reference' },
        { field: 'Lease.Tenant.City', label: 'City' }, { field: 'Lease.Tenant.Account', label: 'Rent account' },
        { field: 'Lease.Tenant.Postcode', label: 'Postcode' }, { field: 'Since', label: 'Unpaid since' },
      ] },
      { type: 'heading', text: 'Rent arrears: {Stage}' },
      { type: 'text', text: 'Dear {Lease.Tenant},' },
      { type: 'text', text: 'Our records show that rent of {Amount Owed} is unpaid on your tenancy of {Lease.Unit.Postal Address}. The oldest unpaid month fell due on {Since}; the number of months unpaid is {Months Owing}.' },
      { type: 'totals', rows: [{ label: 'Rent unpaid', expr: '[Amount Owed]', strong: true }] },
      { type: 'text', text: 'Please pay the full amount by {Pay By}, by bank transfer to {Business.Bank Details}, quoting your rent account {Lease.Tenant.Account}, or by card by calling us on {Business.Phone}.' },
      { type: 'text', text: '{Next Step}' },
      { type: 'text', text: 'If you have paid in the last few days, thank you, and please ignore this letter. If you cannot pay it all by that date, please call us before then: we would much rather agree a plan with you to clear it in steps than take things further. Free and confidential advice on money and debt is available from advice services near you.' },
      { type: 'text', text: 'Yours sincerely,' },
      { type: 'spacer', size: 24 },
      { type: 'text', text: 'Rent accounts, {Business.Name}' },
    ] },
    { name: 'Notice of inspection', table: 'Inspections', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'fields', columns: 2, items: [
        { field: 'Lease.Tenant', label: 'To' }, { field: 'Notice Sent', label: 'Date' },
        { field: 'Unit.Postal Address', label: 'Address' }, { field: 'Kind', label: 'Visit for' },
        { field: 'Unit.Property.City', label: 'City' }, { field: 'Due', label: 'Day of the visit' },
        { field: 'Unit.Property.Postcode', label: 'Postcode' }, { field: 'Visit Time', label: 'Time' },
      ] },
      { type: 'heading', text: 'Notice of a visit to {Unit.Postal Address}' },
      { type: 'text', text: 'Dear {Lease.Tenant},' },
      { type: 'text', text: 'We are writing to let you know that we will visit {Unit.Postal Address} on {Due}, {Visit Time}. The visit is for: {Kind}.' },
      { type: 'text', text: 'A routine inspection takes about half an hour. We look at the general condition of the property, test the smoke and carbon monoxide alarms, look for any damp, leaks or damage, and note anything that needs mending so that we can arrange it. A gas or electrical safety check is made by a registered engineer from our contractor, takes about an hour, and needs the boiler, the meters and the fuse board to be easy to reach; you will be given a copy of the certificate.' },
      { type: 'text', text: 'You do not need to be in, though you are welcome to be, and it is a good time to tell us about anything that needs attention. If you are out, we will let ourselves in with our key, lock up when we leave, and leave a card to say that we have been.' },
      { type: 'text', text: 'If the day or time does not suit you, please call us on {Business.Phone} or write to {Business.E-mail}, and we will find another time in the next two weeks. This letter gives you at least 24 hours\' notice in writing, as your agreement asks.' },
      { type: 'text', text: 'Yours sincerely,' },
      { type: 'spacer', size: 24 },
      { type: 'text', text: 'Property management, {Business.Name}' },
    ] },
  ],
};

export const about = {
  name: 'Hollybrook Lettings',
  kind: 'Property rentals',
  line: 'Buildings, units, tenants, leases, rent invoices, deposits, repairs, safety checks and the accounts for a landlord.',
  summary: [
    'Hollybrook Lettings is an invented letting business in Leeds that owns and manages twelve buildings: six houses, two blocks of flats, two houses split into flats, a parade of three shops with two flats above, and a yard of two workshops bought about a year ago and refitted. Between them they hold thirty-eight units let to tenants on residential and commercial leases.',
    'The sample holds eighteen months of letting. Each unit has its run of leases: tenants who renewed, tenants who left, and the weeks a unit stood empty between them. A rent invoice goes out two weeks before each rent day, numbered RENT-2026-0042. Most tenants pay on the day by standing order, some pay late, and a few have fallen behind and pay part of what is due; a reminder letter goes to anyone a week behind, and every fortnight after while the rent is still unpaid. Things break and are reported, more of them in winter, and a contractor is booked to mend each one. Homes are inspected twice a year and shops and workshops once a year, gas appliances are checked every year and the wiring every five years, and the tenant gets a written notice before each visit. Insurance, loan interest, cleaning of shared areas, letting costs and the bills for empty units are paid out of each building\'s account.',
    'Rent on a home carries no VAT. The workshop yard is a building the landlord has opted to tax, so its rent invoices add VAT at twenty percent, and the VAT collected is paid over to the tax office each quarter.',
    'Money for each building goes through one Transactions table, rent in and costs out, so income and expenditure, net income and the month-by-month accounts come from one place. A rent payment also points at the rent invoice it pays, and the invoice keeps the total paid against it, which is how arrears are found. Seven print templates make the papers a landlord sends: the rent invoice, the rent statement, the tenancy agreement for a home, the summary of a commercial lease, the deposit receipt, the arrears letter and the notice of an inspection.',
  ],
  suits: ['A private landlord with a few houses or flats', 'A small letting agent managing homes for owners', 'Holiday lets, with weekly lets in place of leases', 'Student houses let by the room', 'Storage units or lock-up garages', 'Allotments or moorings let by the year', 'A small commercial landlord with shops and workshops'],
  tables: {
    Business: BUSINESS_ABOUT,
    Properties: { about: 'Each building you own or manage, one record each.', fields: {
      Name: 'What you call the building, such as Linden Court or 22 Mill Lane. Each name is used once.',
      Kind: 'House, Block of flats, Converted house, Shops with flats or Workshop yard.',
      Address: 'The street address.', City: 'The town or city.', Postcode: 'The postcode.',
      Bought: 'The day it was bought.', 'Purchase Price': 'What it cost.',
      'Current Value': 'What it is worth now, from the last valuation. The yield query divides rent by this.',
      'Valued On': 'When it was last valued.',
      'VAT on Rent': 'Yes when you have opted to tax the building, so VAT is added to the rent of its shops and workshops. Homes never carry VAT.',
      Notes: 'Anything worth knowing about the building: a lease on the freehold, a loan, access.',
    } },
    Units: { about: 'What is let: a house, a flat, a shop or a workshop. A house let whole is one unit; a block of flats has one unit for each flat.', fields: {
      Property: 'The building it is in.', Name: 'The unit\'s name within the building, such as Flat 3, Shop 2 or Whole house.',
      Kind: 'House, Flat, Shop or Workshop.',
      'Postal Address': 'The unit\'s address as a letter is addressed, without the town and postcode: Flat 3, Linden Court, 14 Beech Grove. Agreements and notices print it.',
      Bedrooms: 'How many bedrooms; nought for a shop or workshop.',
      'Market Rent': 'What the unit would let for each month now. Used for empty units and the yield.',
      'Has Gas': 'Yes when the unit has gas appliances, so it needs a gas safety check every year.',
      Furnished: 'Yes when it is let with furniture.',
      'Rent Includes': 'What the rent pays for besides the use of the unit, in words the tenancy agreement prints: buildings insurance, cleaning of the shared halls, the furniture.',
      Status: 'Let or Vacant, as it stands today.',
      'Vacant Since': 'The first day it stood empty; empty while it is let.', Notes: 'Anything about the unit: the parking space, the meter numbers, what the inventory says.',
    } },
    Tenants: { about: 'Everyone who rents from you, past and present, one record each. A tenant who renews stays one record with several leases.', fields: {
      Account: 'The tenant\'s rent account, quoted on every payment and printed on every invoice and statement: TEN-0042. Each is used once.',
      'First Name': 'Their first name, or the first name of the person who signs for a business.', 'Last Name': 'Their family name.',
      'Business Name': 'The business that rents a shop or workshop; empty for a home.',
      Address: 'Where letters to them go: the unit they rent while they live there, a forwarding address after they leave.',
      City: 'The town or city of that address.', Postcode: 'Its postcode.', Phone: 'Their phone number.',
      'E-mail': 'Their e-mail address.', 'Emergency Contact': 'Who to call if you cannot reach them, and how they are related.',
      'Emergency Phone': 'That person\'s phone number.', 'References Checked': 'Yes once their references and credit check have come back.',
      'ID Checked': 'Yes once you have seen proof of who they are and their right to rent.', Notes: 'Anything worth knowing: pets, a guarantor, a forwarding address.',
    } },
    Leases: { about: 'Each agreement to let a unit to a tenant for a term. A renewal is a new lease for the same tenant and unit. The Tenancy agreement and the Commercial lease summary are printed from here.', fields: {
      Unit: 'The unit let.', Tenant: 'Who it is let to.', Type: 'Residential or Commercial.',
      'Signed On': 'The day both sides signed it, the date the agreement prints.',
      Start: 'The first day of the lease.', End: 'The last day of the fixed term.', 'Monthly Rent': 'The rent each month, before any VAT.',
      'Tax Rate': 'The VAT rate on the rent as a fraction: 0.2 for a shop or workshop in a building opted to tax, 0 for everything else. Each rent invoice takes its rate from here.',
      'Rent Day': 'The day of the month rent falls due.', 'Pays By': 'Standing order, Bank transfer, Card or Cash.',
      Deposit: 'The deposit agreed: five weeks\' rent for a home, three months\' for a shop or workshop. The money itself is recorded in Deposits.',
      'Deposit Ref': 'The reference the deposit protection scheme gave for a home\'s deposit (invented here). Empty until it is protected. A renewal keeps the same one.',
      Renewal: 'Yes when the lease follows an earlier one for the same tenant.',
      Pets: 'The pets agreed for a home, in words the agreement prints, or None.',
      'Permitted Use': 'For a shop or workshop, what the tenant may use it for, as the lease says.',
      'Rent Review': 'For a commercial lease, the day the rent is next reviewed; empty when the lease has no review.',
      Status: 'Upcoming, Active or Ended, as it stands today.',
      Notes: 'Anything agreed that is not in the fields: a break clause, who repairs what. The commercial lease summary prints it.',
    } },
    'Rent Charges': { about: 'The rent invoices: one record for each lease and month, raised two weeks before the rent day (on the day of signing for a lease\'s first month). The Rent invoice is printed from here and the Rent statement lists them.', fields: {
      Number: 'The invoice\'s number, given in the order invoices go out and once only: RENT-2026-0042, a new run each year.',
      Lease: 'The lease the rent is due on.', Tenant: 'The tenant charged, the same as the lease\'s; kept here so a tenant\'s statement can list their invoices.',
      Period: 'The month it is for, as 2026-10.', 'Invoice Date': 'The day the invoice was raised.', 'Due Date': 'The day the rent is due, the first day the month\'s rent covers.',
      'Period End': 'The last day the month\'s rent covers.',
      Net: 'The rent before VAT.', 'Tax Rate': 'The VAT rate, taken from the lease when the invoice is raised.', Tax: 'The VAT: the rent times the rate, rounded to the penny.',
      Amount: 'What the invoice comes to: rent and VAT together.',
      'Amount Paid': 'How much has been paid against it so far. Add each payment here as you enter it in Transactions; what is left is owed.',
      'Last Paid On': 'The day the last payment against it came in; empty while nothing has been paid.',
    } },
    Deposits: { about: 'Each deposit taken, one record for each new tenancy: a renewal keeps the deposit already held. The Deposit receipt is printed from here.', fields: {
      Number: 'The receipt\'s number, given in the order deposits come in: DRC-2026-0042.', Lease: 'The lease the deposit is for.',
      'Received On': 'The day it was paid, usually the day the agreement was signed.', Amount: 'How much was paid.',
      'Paid By': 'How it was paid.', 'Held By': 'Deposit protection scheme for a home, Landlord for a shop or workshop.',
      'Returned On': 'The day it was paid back after the tenancy ended; empty while it is held.',
      'Amount Returned': 'How much was paid back. Anything kept back is explained in Notes.', Notes: 'What was kept back and why, or anything agreed about it.',
    } },
    'Rent Reminders': { about: 'Each letter sent about rent that is late: the first a week after the rent day, then one a fortnight while it is still unpaid. The Rent arrears letter is printed from here.', fields: {
      Number: 'The letter\'s reference: ARR-2026-0042.', Lease: 'The lease the rent is owed on.', 'Sent On': 'The day the letter went out.',
      Stage: 'First reminder, Second reminder or Final notice.', 'Amount Owed': 'The rent unpaid on the day the letter was written.',
      'Months Owing': 'How many months of rent that covers.', Since: 'The due date of the oldest unpaid month.',
      'Pay By': 'The date the letter asks the tenant to pay by, fourteen days after it is sent.',
      'Next Step': 'The paragraph saying what happens if the rent is not paid, which the letter prints. It is firmer at each stage; change it for one letter if you need to.',
      Outcome: 'Waiting until the pay-by date, then Paid in full, Part paid or Not paid.',
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
    Inspections: { about: 'Routine inspections of each unit and the safety checks the law asks for, each one due on a date. When one is done, the next one is entered with its due date. The Notice of inspection is printed from here.', fields: {
      Unit: 'The unit inspected.', Lease: 'The lease running on the day of the visit, whose tenant gets the notice; empty when the unit is empty.',
      Kind: 'Routine inspection, Gas safety or Electrical safety.', Due: 'When it must be done by, and the day the visit is booked for.',
      'Visit Time': 'The part of the day the visit is booked for, in words the notice prints: between 10am and 1pm.',
      'Notice Sent': 'The day the notice went to the tenant; empty until it has.',
      Done: 'The day it was done; empty until then.', Inspector: 'Who did it: one of your staff, or the contractor who did the check.',
      Outcome: 'What was found, such as Good, Action needed, Pass or Unsatisfactory.', 'Certificate Number': 'The number on the safety certificate (invented here).',
      Notes: 'Anything to know before going, or what happened.',
    } },
    Transactions: { about: 'Every sum of money in or out for a building: rent received, and every cost paid. Each building\'s accounts are made from these.', fields: {
      Property: 'The building the money is for.', Date: 'The day it was paid or received.',
      Kind: 'Rent, Repairs, Insurance, Loan interest, Shared areas, Safety checks, Letting costs, Agent fees, Empty unit costs or VAT.',
      Description: 'What it was for.', Charge: 'For rent, the rent invoice it pays; empty for a cost.',
      'Payee or Payer': 'Who paid you, or who you paid.', Method: 'How it was paid.',
      'Money In': 'What came in; nought for a cost.', 'Money Out': 'What went out; nought for rent.',
    } },
  },
  queries: {
    'Rent roll': 'Every lease running today, by building and unit, with the tenant, the term, the rent and the deposit.',
    'Rent arrears by tenant': 'Each tenant who owes rent that is past due, with how many months are unpaid, how much they owe and how many days late the oldest month is. Former tenants who still owe are on it too.',
    'Unpaid rent': 'Every month\'s rent past its due day and not paid in full, by tenant, with its invoice number and the address to write to.',
    'Rent due in the next 14 days': 'Rent falling due in the next two weeks and not yet paid, with how each tenant pays, so you know what to expect in the bank.',
    'Rent invoices raised': 'The rent invoices raised in the last 31 days, in number order, with the rent, the VAT and what has been paid.',
    'Leases ending in the next 60 days': 'Leases whose term ends within two months, with the rent now and what the unit would let for, so you can offer a renewal or plan to relet.',
    'Vacant units': 'Units standing empty, how long for, and the rent lost at the market rent since they emptied.',
    'Open repairs': 'Repairs not yet done or cancelled, the most urgent first and then the oldest, with how many days each has been open.',
    'Repairs booked': 'Jobs booked with a contractor and not yet done, by contractor and day, with the address and the problem.',
    'Repair costs by contractor': 'What each contractor has been paid for finished jobs in the last twelve months, and how many jobs.',
    'Inspections due': 'Inspections and safety checks due in the next 30 days or already overdue (a Days Left below nought is overdue), with the day the notice went out.',
    'Safety certificates due': 'Gas and electrical safety checks due in the next 60 days or overdue, so the engineer can be booked in time.',
    'Deposits held': 'The deposit taken on every lease running today, with its protection reference.',
    'Deposits to protect': 'Home deposits not yet protected, with how many days are left of the 30 allowed from the start of the lease.',
    'Arrears letters sent': 'The reminder letters sent in the last 90 days, the newest first, with the stage, the amount, the pay-by date and what happened.',
    'Current tenants': 'Everyone living or trading in your units today, with their rent account, phone, e-mail and emergency contact.',
    'Income and expenditure by property': 'For each building over the last twelve months: rent received, costs paid and the net income.',
    'Costs by kind': 'What was spent in the last twelve months on each kind of cost, the biggest first. VAT paid over is left out, as it is not a cost.',
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
    'Rent invoices this month': 'The rent invoices raised in the last 31 days with the rent, VAT and totals: the month\'s invoicing at a glance.',
  },
  documents: {
    'Rent invoice': 'The invoice each tenant gets two weeks before the rent day: the month it covers, the rent, any VAT, the payments already received against it and the balance to pay, with the bank details to pay to. Numbered RENT-2026-0042. It prints for every rent invoice, so choose the ones to print, or print from the Rent Charges form.',
    'Rent statement': 'A tenant\'s account: every rent invoice with what was charged, paid and when, and the balance owing. Sent when a tenant asks, with an arrears letter, or when a tenancy ends. It prints for every tenant; choose the one you want, or print from the Tenants form.',
    'Tenancy agreement': 'The agreement for a home, filled from the lease, the tenant, the unit and the building: the parties, the property, the term, the rent and rent day, the deposit and its protection, what the rent includes, each side\'s duties, repairs and access, pets and smoking, ending the tenancy, and places to sign. It is a plainly worded starting point of our own: have it checked against the law where your homes are before you use it. Choose a residential lease to print it.',
    'Commercial lease summary': 'A summary of the main terms of a lease of a shop or workshop: the premises, the use allowed, the term, the rent with any VAT and its review, the deposit, outgoings, repairs, access and the end of the lease. It goes with the full lease, which a solicitor draws up; have both checked for your own country. Choose a commercial lease to print it.',
    'Deposit receipt': 'Given when a new tenant pays the deposit: how much, on what day, for which tenancy, and who holds it. Numbered DRC-2026-0042.',
    'Rent arrears letter': 'The reminder sent when rent is a week late, and each fortnight after while it is unpaid: polite, clear about the amount and the date to pay by, and firmer at each stage, ending with the final notice. Have the final notice\'s wording checked for your own country before you rely on it.',
    'Notice of inspection': 'The letter that tells a tenant, at least a week ahead, the day and time of a routine inspection or safety check, what happens at it and how to change the day.',
  },
  tutorial: [
    { title: 'Let a unit to a new tenant', steps: [
      'Open the query Vacant units to see what is empty and what it would let for.',
      'Open the Tenants form and press New. Give the next rent account in Account (TEN- and the next number), then the name, Address (the unit\'s address, where letters will go), City, Postcode, phone, e-mail and emergency contact. Tick References Checked and ID Checked when they are done.',
      'Open the Leases form and press New. Choose the Unit and the Tenant, set Type to Residential or Commercial, and type Signed On, Start, End, Monthly Rent, Rent Day, Pays By and Deposit. For a home type the Pets agreed, or None; for a shop or workshop type the Permitted Use, the Rent Review date and a Tax Rate of 0.2 if the building is opted to tax. Set Status to Upcoming or Active.',
      'Open the print template Tenancy agreement (or Commercial lease summary for a shop or workshop) from Print templates, choose the new lease and press Print. Both sides sign two copies.',
      'Open the Units form, find the unit, set Status to Let and clear Vacant Since.',
    ] },
    { title: 'Take and protect a deposit', steps: [
      'When the deposit comes in, open the Deposits form and press New. Give the next receipt number in Number (DRC-, the year and the next number), choose the Lease, and type Received On, Amount, Paid By and Held By.',
      'Open the print template Deposit receipt, choose the new deposit and press Print or Save as PDF, and give it to the tenant.',
      'For a home, protect the deposit with your scheme within 30 days. Open the query Deposits to protect to see which are still waiting and how many days are left.',
      'When the scheme gives you its reference, open the Leases form, find the lease and type it into Deposit Ref.',
    ] },
    { title: 'Raise the rent invoices', steps: [
      'Each working day, open the query Rent due in the next 14 days, or the Leases form, to see whose rent day is two weeks away.',
      'Open the Rent Charges form and press New for each one. Give the next number in Number (RENT-, the year and the next number), choose the Lease and the Tenant, and type Period (as 2026-11), Invoice Date (today), Due Date (the rent day) and Period End (the day before the next rent day).',
      'Type Net (the lease\'s Monthly Rent) and Tax Rate (the lease\'s Tax Rate), then Tax (the rent times the rate, to the penny) and Amount (Net and Tax together). Leave Amount Paid at nought.',
      'Open the print template Rent invoice from Print templates, choose the new invoices and press Print or Save as PDF to send them.',
      'At the end of the month, open the report Rent invoices this month to check the month\'s invoicing.',
    ] },
    { title: 'Record a rent payment', steps: [
      'When rent reaches the bank, open the Transactions form and press New. Choose the Property, type the Date, set Kind to Rent, choose the rent invoice in Charge, type the tenant in Payee or Payer, the Method, and the sum in Money In.',
      'Open the Rent Charges form, find the same invoice, add the payment to Amount Paid and type the day in Last Paid On.',
      'Open the query Unpaid rent: the invoice drops off it once Amount Paid reaches Amount.',
    ] },
    { title: 'Chase rent that is late', steps: [
      'Open the query Rent arrears by tenant, or the report Arrears letters for every unpaid month by tenant.',
      'For each tenant a week or more behind, open the Rent Reminders form and press New. Give the next reference in Number (ARR-, the year and the next number), choose the Lease, and type Sent On, Stage, Amount Owed, Months Owing, Since and Pay By (fourteen days on). Copy the Next Step wording from an earlier letter of the same stage.',
      'Open the print template Rent arrears letter, choose the new letter and press Print. Print the tenant\'s Rent statement from Print templates too, choosing the tenant, and send it with the letter.',
      'When the pay-by date comes, set Outcome to Paid in full, Part paid or Not paid. Open the query Arrears letters sent to see the last three months of letters and what came of them.',
    ] },
    { title: 'Arrange an inspection or safety check', steps: [
      'Open the query Inspections due, or the report Checks due, to see what falls due in the next 30 days. For gas and electrical checks, open the query Safety certificates due, which looks 60 days ahead.',
      'Agree a day with the inspector or the engineer. Open the Inspections form, find the visit, set Due to the day, type the Visit Time, choose the running Lease, and type today in Notice Sent.',
      'Open the print template Notice of inspection, choose the visit and press Print, and send it to the tenant at least a week ahead.',
      'After the visit, type Done, Inspector, Outcome and, for a safety check, the Certificate Number. Then press New and enter the next one with its due date. If something needs mending, add it to Maintenance Requests with Reported By set to Inspection.',
    ] },
    { title: 'Deal with a repair', steps: [
      'When a tenant reports a problem, open the Maintenance Requests form and press New. Choose the Unit, type Reported, Reported By, the Problem and the Priority.',
      'Book a contractor: choose them in Contractor, type Booked For and set Status to Booked. Open the report Repairs for contractors to print each firm\'s list of jobs.',
      'When the work is done, type Completed and Cost and set Status to Done. Open the Transactions form and enter the bill as money out with Kind set to Repairs and the building in Property.',
      'Open the query Open repairs each morning to see what is still waiting, the most urgent first.',
    ] },
    { title: 'Close the month', steps: [
      'Open the report Monthly accounts for money in and out each month, and the report Income and expenditure for each building over the last year.',
      'Open the query Costs by kind to see where the money went.',
      'At the end of each VAT quarter, add up the Tax on the quarter\'s paid invoices for the opted building and pay it over; enter it in Transactions with Kind set to VAT.',
      'Open the report Rent roll for the units, tenants and rents as they stand, and the report Leases ending to plan renewals.',
    ] },
  ],
  notes: [
    'A rent payment is recorded twice on purpose: once in Transactions, where the money is, and once in the invoice\'s Amount Paid, which is how arrears are found. When you enter one, change the other.',
    'Monthly Rent on a lease is the rent before VAT. A rent invoice\'s Amount is the rent and VAT together, and that is what Amount Paid is measured against.',
    'Only the workshop yard is opted to tax, so only its two workshops have a Tax Rate of 0.2. The shops in Market Parade are commercial lets without VAT, which is how a building is until its landlord opts to tax it.',
    'The Status of units and leases (Let, Vacant, Upcoming, Active, Ended) is as it stood on the day the sample was made. Change it by hand when a tenancy starts or ends; the queries for current tenants and the rent roll read it.',
    'A print template prints a document for every record of its table: every rent invoice, every lease, every inspection. Choose the records you want before you print, or print the one on screen from its form. The Tenancy agreement is worded for a home and the Commercial lease summary for a shop or workshop, so choose a lease of the right Type.',
    'An arrears letter keeps the amount owed on the day it was written. It is a record of what the tenant was told, so it does not change when they pay.',
    'A renewal is a new lease for the same tenant and unit. It keeps the deposit and its protection reference from the lease before, so it has no Deposits record of its own.',
    'Tenants keep a correspondence address. For a tenant living in one of your units it is that unit; when they leave, change it to their forwarding address so the final statement and the deposit letter reach them.',
  ],
  faq: [
    { q: 'How do I print one tenant\'s rent invoice?', a: 'Open the Rent Charges form, find the invoice by its Number, and print the template Rent invoice from there. Or open Print templates, choose Rent invoice, pick the invoice from the list and press Print.' },
    { q: 'How do I see who owes rent?', a: 'Open the query Rent arrears by tenant for a line a tenant with what they owe and how late the oldest month is, or the report Arrears letters for every unpaid month. The query Unpaid rent lists each unpaid invoice.' },
    { q: 'Why do the workshop invoices have VAT and the others not?', a: 'Rent on a home never carries VAT. A commercial landlord adds VAT only after opting to tax the building, and here only Riverside Yard is opted (its VAT on Rent is ticked), so its leases have a Tax Rate of 0.2 and its invoices add VAT. The VAT collected is paid over each quarter and shows in Transactions as VAT.' },
    { q: 'How do I renew a tenancy?', a: 'Open the Leases form, press New, and enter the same Unit and Tenant with the new Start, End and rent, tick Renewal and copy the Deposit and Deposit Ref from the old lease. Set the old lease\'s Status to Ended on its last day. Print a new Tenancy agreement for the new lease.' },
    { q: 'What happens to the deposit at the end?', a: 'Agree any deductions with the tenant against the inventory, then return the rest. Open the Deposits form, find the deposit and type Returned On, Amount Returned and, in Notes, what was kept back and why.' },
    { q: 'How do I print older documents, or many at once?', a: 'Every record of a print template\'s table can print. Open Print templates, choose the template, and choose the records you want in its list; for a month\'s rent invoices, choose those whose Invoice Date falls in that month. The query Rent invoices raised lists the last month\'s.' },
    { q: 'Can I use this for flats I manage for other owners?', a: 'Yes, with an Owners table: see the first idea under adapting it. The Properties, Units, Leases and rent invoices stay as they are.' },
  ],
  empty: 'Start with Business, your own name, address, VAT number and bank details, which every invoice, statement and letter prints. Then enter Properties, then the Units in each one (with the Postal Address and what the rent includes), then Contractors. Enter each Tenant with their rent account, their Lease, and the Deposit taken, and for every unit the next inspection and safety checks due. Two weeks before each rent day, add a Rent Charge for each running lease; when rent comes in, enter it in Transactions pointing at the charge, and add it to the charge\'s Amount Paid. Enter repairs in Maintenance Requests as they are reported, reminder letters in Rent Reminders as you send them, and every bill you pay in Transactions with its building.',
  adapt: [
    'Manage for owners: add an Owners table with their management fee, point Properties at it, and add a query that takes your fee off each owner\'s rent to show what you pay them.',
    'Let holiday homes: use a weekly Rent Charge in place of a monthly one, rename Leases to Bookings with arrival and departure dates, and add a cleaning cost to each booking.',
    'Let rooms in shared houses: make each room a unit, and add a field to Units for the shared bathroom and kitchen it uses.',
    'Let storage units or garages: drop Bedrooms and Has Gas from Units, add a Size field, and keep the rest as it is.',
    'Keep the energy rating: add an Energy Rating and an Energy Certificate Expires field to Units, and a query for certificates running out.',
    'List rent reviews coming up: add a query on Leases with a calculated Days to Review column, DateDiff("d", Date(), [Rent Review]), and the criteria Between 0 And 90.',
    'Recharge damage to tenants: add a Recharge field to Maintenance Requests, and take it from the deposit when the lease ends.',
    'Charge a service charge on the shops: add a Service Charge line to Rent Charges as a second amount, and print it on the Rent invoice as its own total row.',
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
// The building the landlord has opted to tax: VAT is added to its rents.
const OPTED = ['Riverside Yard'];

// What the rent pays for besides the unit, by the kind of building, as the agreement prints it.
const INCLUDES = {
  'Block of flats': 'the buildings insurance, and the cleaning, lighting and upkeep of the shared halls, stairs, grounds and bin store',
  'Shops with flats': 'the buildings insurance, and the cleaning and lighting of the shared entrance and stairs',
  'Converted house': 'the buildings insurance, and the upkeep of the shared entrance, garden and bins',
  House: 'the buildings insurance',
  'Workshop yard': 'the buildings insurance, and the upkeep of the shared yard and its gates',
};

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

// Businesses that rent the shops and the workshops, with what their lease lets them use it for. Invented names.
const SHOP_TENANTS = [
  ['Amberly Florist', 'Retail sale of cut flowers, plants and gifts'], ['Copperfield Barbers', 'Barber shop and men\'s grooming'],
  ['Larkspur Books', 'Retail sale of new and second-hand books'], ['Juniper Nail Studio', 'Nail and beauty treatments'],
  ['Willow Dry Cleaners', 'Taking in and handing out dry cleaning and laundry, and clothing alterations'],
  ['Linden Cards and Gifts', 'Retail sale of cards, gifts and stationery'], ['Alder Phone Repairs', 'Repair and sale of phones and small electronics'],
  ['Saffron Takeaway', 'Hot food takeaway, open no later than 11pm'],
];
const YARD_TENANTS = [
  ['Granite Print Shop', 'Printing, copying and print finishing'], ['Orchard Upholstery', 'Upholstery and furniture repair'],
  ['Northgate Motor Repairs', 'Light vehicle repairs and servicing, no paint spraying'], ['Kestrel Cycle Repairs', 'Cycle repairs, servicing and sales'],
];
const PETS = ['One small dog', 'One cat', 'Two cats', 'One rabbit, kept indoors'];
const VISIT_TIMES = ['between 9am and 12 noon', 'between 10am and 1pm', 'between 1pm and 4pm'];

// What an arrears letter says will happen next, by its stage.
const NEXT_STEP = {
  'First reminder': 'It may simply have been missed, so this is a friendly reminder. If something has gone wrong with your standing order or your bank, let us know and we will help to put it right.',
  'Second reminder': 'We wrote to you about this two weeks ago and have not yet received the payment or heard from you. If it is not paid by the date above, or you have not been in touch, we will need to consider what to do next.',
  'Final notice': 'This is our final letter before we take formal steps to recover the arrears, which may include giving notice to end your tenancy and asking the court to order payment. We do not want to do that, and it can still be avoided if you pay, or contact us to agree a plan, by the date above.',
};

const slug = (s) => s.toLowerCase().replace(/[^a-z]+/g, '');
const round5 = (n) => Math.round(n / 5) * 5;
const pad4 = (n) => String(n).padStart(4, '0');

/** Fill a Hollybrook Lettings database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 540 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };

  add('Business', businessRecord({
    name: 'Hollybrook Lettings', domain: 'hollybrook.example', phone: '020 7946 0450',
    address: { Address: '5 Market Square', City: 'Leeds', Postcode: 'LS1 4DE', Country: 'United Kingdom' },
    terms: 'Rent is due in advance on the rent day shown. Please pay by standing order or bank transfer, quoting your invoice number or rent account.',
    footer: 'Hollybrook Lettings: homes, shops and workshops to let in Leeds. Office open Monday to Friday, 9am to 5pm. For an emergency out of hours, such as a water leak or a break-in, call the office number and choose option 1.',
  }));

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
  // Numbers given in date order, a new run each year: RENT-2026-0042.
  const numberer = (prefix) => { const n = {}; return (off) => { const y = cal.at(off).getFullYear(); n[y] = (n[y] || 0) + 1; return `${prefix}-${y}-${pad4(n[y])}`; }; };
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
    const postcode = `${r.pick(streets)} ${r.int(1, 9)}${r.pick(['AB', 'DE', 'FG', 'HJ', 'NP', 'QR', 'TU'])}`;
    const p = { name, kind, street, postcode, recent, boughtAt, vat: OPTED.includes(name), units: [] };
    p.id = add('Properties', {
      Name: name, Kind: kind, Address: street, City: 'Leeds', Postcode: postcode,
      Bought: cal.day(boughtAt), 'Purchase Price': paid, 'Current Value': value, 'Valued On': cal.day(-r.int(30, 300)), 'VAT on Rent': p.vat,
      Notes: MORTGAGED.includes(name) ? 'Bought with a buy-to-let loan; the interest is paid monthly.' : recent ? 'Bought empty and refitted before letting. Opted to tax, so VAT is added to the rents.' : null,
    });
    for (const [uname, ukind, beds, rent] of list) {
      const commercial = ukind === 'Shop' || ukind === 'Workshop';
      // A house let whole is its street address; a flat or shop is named within its building.
      const postal = name === street ? (uname === 'Whole house' ? street : `${uname}, ${street}`) : `${uname}, ${name}, ${street}`;
      const furnished = !commercial && ukind === 'Flat' && r.chance(0.3);
      const u = { prop: p, name: uname, kind: ukind, beds, market: rent, commercial, postal, gas: gas && !commercial, label: `${uname}, ${name}`, leases: [] };
      u.id = add('Units', {
        Property: p.id, Name: uname, Kind: ukind, 'Postal Address': postal, Bedrooms: beds, 'Market Rent': rent, 'Has Gas': u.gas, Furnished: furnished,
        'Rent Includes': `${INCLUDES[kind]}${furnished ? ', and the use of the furniture listed in the inventory' : ''}`, Status: 'Let',
      });
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

  // Tenants and leases, oldest first. A new tenant signs, and pays the
  // deposit, a week or two before moving in; a renewal is signed a month or
  // so before the old term ends. Nobody signs in the future.
  const all = units.flatMap((u) => u.leases).sort((a, b) => a.S - b.S || a.unit.id - b.unit.id);
  for (const l of all) {
    l.signedAt = l.renewal ? l.S - r.int(20, 40) : l.S - r.int(4, 16);
    if (l.signedAt > 0) l.signedAt = -r.int(1, 8);
  }
  const relations = ['mother', 'father', 'sister', 'brother', 'partner', 'friend', 'daughter', 'son'];
  let shops = 0;
  let yards = 0;
  let accounts = 0;
  for (const l of all) {
    const t = l.tenant;
    if (!t.id) {
      const p = person(r, { domain: 'mail.example' });
      const kin = person(r);
      const [name, use] = !t.commercial ? [null, null] : l.unit.kind === 'Shop' ? SHOP_TENANTS[shops++ % SHOP_TENANTS.length] : YARD_TENANTS[yards++ % YARD_TENANTS.length];
      // Letters go to the unit while they rent it, to a forwarding address after they leave.
      const current = t.leases.some((x) => x.E >= 0);
      const away = address(r, PLACES.slice(0, 1));
      const post = current ? { Address: l.unit.postal, City: 'Leeds', Postcode: l.unit.prop.postcode } : { Address: away.Address, City: away.City, Postcode: away.Postcode };
      t.pet = !t.commercial && r.chance(0.1) ? r.pick(PETS) : null;
      t.account = `TEN-${pad4(++accounts)}`;
      t.id = add('Tenants', {
        Account: t.account, 'First Name': p['First Name'], 'Last Name': p['Last Name'], 'Business Name': name, ...post, Phone: phoneFor(r, 'United Kingdom'),
        'E-mail': name ? `hello@${slug(name)}.example` : p['E-mail'],
        'Emergency Contact': `${kin['First Name']} ${kin['Last Name']} (${t.commercial ? 'business partner' : r.pick(relations)})`, 'Emergency Phone': phoneFor(r, 'United Kingdom'),
        'References Checked': l.S > -days ? true : r.chance(0.85), 'ID Checked': l.S > 0 ? false : true,
        Notes: t.commercial ? null : t.pet ? `${t.pet}, agreed in writing.` : r.chance(0.08) ? r.pick(['Guarantor: a parent, details on file.', 'Works nights; call after noon.']) : current ? null : 'Moved out; forwarding address given.',
      });
      t.display = name || `${p['First Name']} ${p['Last Name']}`;
      t.use = use;
      t.method = t.commercial ? 'Bank transfer' : r.weighted([['Standing order', 7], ['Bank transfer', 2], ['Card', 0.7], ['Cash', 0.3]]);
    }
    const prev = l.renewal ? t.last : null;
    l.deposit = prev ? prev.deposit : l.unit.commercial ? l.rent * 3 : Math.floor((l.rent * 12) / 52 * 5);
    l.ref = prev ? prev.ref : l.unit.commercial || l.S > 0 || (l.S > -12 && r.chance(0.5)) ? null : `DEP-${cal.day(l.S).slice(0, 4)}-${String(r.int(10000, 99999))}`;
    l.rate = l.unit.prop.vat ? 0.2 : 0;
    const months = Math.round((l.E - l.S + 1) / 30.44);
    const status = l.S > 0 ? 'Upcoming' : l.E < 0 ? 'Ended' : 'Active';
    l.id = add('Leases', {
      Unit: l.unit.id, Tenant: t.id, Type: l.unit.commercial ? 'Commercial' : 'Residential', 'Signed On': cal.day(l.signedAt), Start: cal.day(l.S), End: cal.day(l.E),
      'Monthly Rent': l.rent, 'Tax Rate': l.rate, 'Rent Day': dom(l.S), 'Pays By': t.method, Deposit: l.deposit, 'Deposit Ref': l.ref, Renewal: l.renewal,
      Pets: l.unit.commercial ? null : t.pet || 'None', 'Permitted Use': t.use,
      'Rent Review': l.unit.commercial && months >= 59 ? cal.day(addMonths(l.S, 36)) : null, Status: status,
      Notes: l.unit.kind === 'Shop' ? 'The tenant repairs the shop front and shutter, and pays the business rates.' : l.unit.kind === 'Workshop' ? 'The tenant repairs the roller door and the inside of the unit, and pays the business rates.' : null,
    });
    t.last = l;
  }

  // Deposits: one for each new tenancy, in the order they came in. Paid back
  // a week or so after the tenant leaves, now and then less something for
  // cleaning or damage.
  const receipt = numberer('DRC');
  for (const l of all.filter((x) => !x.renewal).sort((a, b) => a.signedAt - b.signedAt || a.unit.id - b.unit.id)) {
    const left = Math.max(...l.tenant.leases.map((x) => x.E));
    const back = left < 0 ? left + r.int(4, 12) : null;
    const returned = back != null && back <= 0;
    const kept = returned && r.chance(0.25) ? Math.min(l.deposit, r.int(6, 30) * 10) : 0;
    add('Deposits', {
      Number: receipt(l.signedAt), Lease: l.id, 'Received On': cal.day(l.signedAt), Amount: l.deposit, 'Paid By': 'Bank transfer',
      'Held By': l.unit.commercial ? 'Landlord' : 'Deposit protection scheme', 'Returned On': returned ? cal.day(back) : null,
      'Amount Returned': returned ? l.deposit - kept : null,
      Notes: kept ? `${kept} kept back for ${r.pick(['cleaning at the end of the tenancy', 'a damaged carpet', 'a broken kitchen cupboard', 'redecorating a marked wall'])}, agreed with the tenant.` : null,
    });
  }

  // The ledger, collected and written in date order at the end.
  const ledger = [];
  const pay = (p, at, kind, description, out, who, method = 'Bank transfer') => {
    if (at < -days || at > 0) return;
    ledger.push({ at, values: { Property: p.id, Date: cal.day(at), Kind: kind, Description: description, 'Payee or Payer': who, Method: method, 'Money In': 0, 'Money Out': money(out) } });
  };

  // Rent: an invoice for each month of each lease, from when the records
  // begin, raised two weeks before the rent day (the first on the day the
  // lease is signed), and the payments made against it.
  const charges = [];
  for (const l of all) {
    const t = l.tenant;
    const net = l.rent;
    const tax = money(net * l.rate);
    const total = money(net + tax);
    let first = true;
    for (let k = 0; ; k++) {
      const due = addMonths(l.S, k);
      if (due > l.E) break;
      const issued = Math.max(l.signedAt, due - 14);
      if (issued > 0) break;
      if (due < -days) continue;
      const payments = [];
      if (due <= 0) {
        if (t.struggling && due >= -100) {
          // Fallen behind: part of the rent, some months none.
          const share = first ? 0.5 : r.pick([0, 0.5, 0.6, 0.75, 1]);
          first = false;
          const at = due + r.int(3, 14);
          if (share > 0 && at <= 0) payments.push([at, money(total * share)]);
        } else {
          const lag = t.profile === 'late' && r.chance(0.6) ? r.int(3, 24) : t.method === 'Standing order' ? 0 : r.int(0, 3);
          const at = due + lag;
          if (at <= 0) {
            if (r.chance(0.04)) {
              // Paid in two halves.
              const half = money(total / 2);
              payments.push([at, half]);
              const at2 = at + r.int(5, 12);
              if (at2 <= 0) payments.push([at2, money(total - half)]);
            } else payments.push([at, total]);
          }
        }
      }
      charges.push({ l, due, issued, end: addMonths(l.S, k + 1) - 1, net, tax, total, payments });
    }
    // Finding a new tenant costs something; a renewal costs a little paperwork.
    const at = l.S - r.int(3, 12);
    if (l.unit.commercial && !l.renewal) pay(l.unit.prop, at, 'Agent fees', `Commercial letting fee, ${l.unit.label}`, l.rent * 12 * 0.1, 'Commercial letting agent');
    else if (!l.renewal) pay(l.unit.prop, at, 'Letting costs', `Advertising, referencing and inventory, ${l.unit.label}`, r.int(180, 290), 'Letting costs');
    else pay(l.unit.prop, at, 'Letting costs', `Renewal paperwork, ${l.unit.label}`, r.int(45, 75), 'Letting costs');
  }
  // Written in the order the invoices went out, so their numbers run in date order.
  charges.sort((a, b) => a.issued - b.issued || a.due - b.due || a.l.unit.id - b.l.unit.id);
  const rentNumber = numberer('RENT');
  const vatDue = new Map(); // the quarter's last day -> { prop, vat }
  for (const c of charges) {
    const { l } = c;
    const t = l.tenant;
    const period = cal.day(c.due).slice(0, 7);
    const paid = money(c.payments.reduce((n, [, a]) => n + a, 0));
    c.number = rentNumber(c.issued);
    c.id = add('Rent Charges', {
      Number: c.number, Lease: l.id, Tenant: t.id, Period: period, 'Invoice Date': cal.day(c.issued), 'Due Date': cal.day(c.due), 'Period End': cal.day(c.end),
      Net: c.net, 'Tax Rate': l.rate, Tax: c.tax, Amount: c.total, 'Amount Paid': paid,
      'Last Paid On': c.payments.length ? cal.day(Math.max(...c.payments.map(([at]) => at))) : null,
    });
    for (const [at, amount] of c.payments) {
      ledger.push({ at, values: { Property: l.unit.prop.id, Date: cal.day(at), Kind: 'Rent', Description: `Rent ${period}, ${l.unit.label}, ${c.number}`, Charge: c.id, 'Payee or Payer': t.display, Method: t.method, 'Money In': amount, 'Money Out': 0 } });
      // The VAT within a payment, paid over to the tax office after the quarter it came in.
      if (c.tax > 0) {
        const d = cal.at(at);
        const q = offsetOf(d.getFullYear(), Math.floor(d.getMonth() / 3) * 3 + 3, 0);
        if (!vatDue.has(q)) vatDue.set(q, { prop: l.unit.prop, vat: 0 });
        vatDue.get(q).vat += amount * c.tax / c.total;
      }
    }
  }
  for (const [q, { prop, vat }] of vatDue) {
    const d = cal.at(q);
    pay(prop, offsetOf(d.getFullYear(), d.getMonth() + 2, 7), 'VAT', `VAT on rents for the quarter to ${cal.day(q).slice(0, 7)}, paid to the tax office`, vat, 'Tax office');
  }

  // Reminders: each working day, a letter to any tenant with rent a week or
  // more late, and another each fortnight while it stays unpaid, firmer each time.
  const owedOn = (c, d) => money(c.total - c.payments.reduce((n, [at, a]) => n + (at <= d ? a : 0), 0));
  const reminders = [];
  for (const t of tenants) {
    const mine = charges.filter((c) => c.l.tenant === t);
    if (!mine.length) continue;
    let lastAt = null;
    let stage = 0;
    for (let d = -days + 7; d <= 0; d++) {
      const wd = cal.weekday(d);
      if (wd === 0 || wd === 6) continue;
      const late = mine.filter((c) => c.due < d && owedOn(c, d) > 0);
      if (!late.length) { stage = 0; lastAt = null; continue; }
      if (!late.some((c) => c.due <= d - 7) || (lastAt != null && d - lastAt < 14)) continue;
      const step = ['First reminder', 'Second reminder', 'Final notice'][Math.min(stage, 2)];
      reminders.push({ at: d, t, late, step, owed: money(late.reduce((n, c) => n + owedOn(c, d), 0)), lease: late[late.length - 1].l });
      stage++;
      lastAt = d;
    }
  }
  reminders.sort((a, b) => a.at - b.at || a.t.id - b.t.id);
  const letterNumber = numberer('ARR');
  for (const x of reminders) {
    const payBy = x.at + 14;
    const left = payBy <= 0 ? money(x.late.reduce((n, c) => n + owedOn(c, payBy), 0)) : null;
    add('Rent Reminders', {
      Number: letterNumber(x.at), Lease: x.lease.id, 'Sent On': cal.day(x.at), Stage: x.step, 'Amount Owed': x.owed, 'Months Owing': x.late.length,
      Since: cal.day(Math.min(...x.late.map((c) => c.due))), 'Pay By': cal.day(payBy), 'Next Step': NEXT_STEP[x.step],
      Outcome: left == null ? 'Waiting' : left <= 0 ? 'Paid in full' : left < x.owed ? 'Part paid' : 'Not paid',
    });
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
      pay(v.u.prop, a + n - 1, 'Empty unit costs', `${v.u.commercial ? 'Business rates and power' : 'Council tax and power'} while empty, ${v.u.label}`, n * (v.u.commercial ? 14 : 6.2), 'City council and energy supplier', 'Direct debit');
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

  // Inspections, in the order they fall due. A tenant gets a written notice
  // a week or so before each visit: sent already for the visits of the next
  // fortnight, not yet for those further off.
  checks.sort((a, b) => a.due - b.due || a.u.id - b.u.id);
  for (const c of checks) {
    const visit = c.done ?? c.due;
    const running = c.u.leases.find((l) => l.S <= visit && visit <= l.E) || null;
    const sent = visit - r.int(7, 12);
    const notice = running && running.signedAt <= sent && sent <= 0 && !c.notes ? sent : null;
    add('Inspections', {
      Unit: c.u.id, Lease: running ? running.id : null, Kind: c.kind, Due: cal.day(c.due), 'Visit Time': running ? r.pick(VISIT_TIMES) : null,
      'Notice Sent': notice != null ? cal.day(notice) : null, Done: c.done != null ? cal.day(c.done) : null, Inspector: c.inspector ?? null,
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
