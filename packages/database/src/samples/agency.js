// The studio: Bramble Studio, an invented design and consultancy studio of
// six people (two partners, two designers, a developer and a project
// manager) working for small and mid-sized clients. Its clients and their
// contacts, the staff with what an hour of each costs and what it is charged
// at, projects (fixed fee or by the hour) broken into tasks, every hour
// logged against a task, expenses on projects, invoices and their lines, and
// the payments that settle them.
//
// Its records are sixteen months of work, run a day at a time: proposals
// come in, most are won and some lost, won projects start, people log their
// hours to the tasks most behind, a few projects overrun or go on hold, and
// the studio is quieter in August and shut over Christmas. Fixed-fee work is
// invoiced in stages, hourly work at the start of each month for the month
// before, and clients pay on time, late or in two parts. The questions a
// studio asks are its saved queries; its reports print them.
//
// It prints what a studio sends: the invoice, the proposal for work quoted,
// a statement of what a client owes, each person's weekly time sheet, and a
// letter of engagement for the client to sign before the work starts.

import { address, BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, companies, money, person, phoneFor, PLACES } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const PERSON = '[First Name] & " " & [Last Name]';
const COST = 'Round([Hours] * [Hourly Cost], 2)';
const AT_RATE = 'Round([Hours] * [Charge-out Rate], 2)';
const LINE = 'Round([Quantity] * [Unit Price], 2)';
// Each line's VAT rounded to pence, then added up: the invoice's VAT is the sum of these.
const LINE_VAT = 'Round(Round([Quantity] * [Unit Price], 2) * Nz([Tax Rate]), 2)';
const NOT_INTERNAL = '<> "Internal"';
// The studio's own details at the head of every document it sends.
const LETTERHEAD = { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Website'] };

export const spec = {
  // Where a value below nought is right: an invoice not yet due, a project already late.
  negative: { 'Unpaid invoices': ['Days Overdue'], 'Projects due': ['Days Left'] },
  tables: [
    BUSINESS_TABLE,
    { name: 'Clients', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Sector', type: 'text' },
      { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' }, { name: 'Country', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Payment Days', type: 'integer', default: 30 },
      { name: 'VAT Rate', type: 'number', default: 0.2 }, { name: 'Client Since', type: 'date' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Contacts', fields: [
      { name: 'Client', ...ref('Clients'), required: true }, { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
      { name: 'Job Title', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'Main Contact', type: 'yesno' },
    ] },
    { name: 'Staff', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Role', type: 'text' },
      { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'Hourly Cost', type: 'currency' },
      { name: 'Charge-out Rate', type: 'currency' }, { name: 'Hours per Week', type: 'number', default: 37.5 },
      { name: 'Started', type: 'date' }, { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Projects', fields: [
      { name: 'Code', type: 'text', unique: true }, { name: 'Name', type: 'text', required: true }, { name: 'Client', ...ref('Clients'), required: true },
      { name: 'Kind', type: 'text', default: 'Fixed fee' }, { name: 'Work Type', type: 'text' }, { name: 'Budget Hours', type: 'number' },
      { name: 'Fee', type: 'currency' }, { name: 'Proposed On', type: 'date' }, { name: 'Start Date', type: 'date' }, { name: 'Due Date', type: 'date' },
      { name: 'Finished On', type: 'date' }, { name: 'Status', type: 'text', default: 'Proposal' }, { name: 'Notes', type: 'longtext' },
      { name: 'Quote No', type: 'text', unique: true }, { name: 'Client PO', type: 'text' }, { name: 'Contact', ...ref('Contacts') },
    ] },
    { name: 'Tasks', fields: [
      { name: 'Project', ...ref('Projects'), required: true }, { name: 'Name', type: 'text', required: true }, { name: 'Role', type: 'text' },
      { name: 'Estimated Hours', type: 'number' }, { name: 'Done', type: 'yesno' },
    ] },
    // One a person a week, which their hours are logged on and which is approved before invoicing.
    { name: 'Time Sheets', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Staff', ...ref('Staff'), required: true }, { name: 'Week Commencing', type: 'date' },
      { name: 'Status', type: 'text', default: 'Open' }, { name: 'Submitted On', type: 'date' }, { name: 'Approved On', type: 'date' },
    ] },
    // Staff, then Project, then Task: a query that puts Tasks straight after
    // Time Entries joins each hour to its own task, not to every task of the project.
    { name: 'Time Entries', fields: [
      { name: 'Date', type: 'date', required: true }, { name: 'Staff', ...ref('Staff'), required: true }, { name: 'Project', ...ref('Projects'), required: true },
      { name: 'Task', ...ref('Tasks') }, { name: 'Hours', type: 'number' }, { name: 'Billable', type: 'yesno', default: true },
      { name: 'Note', type: 'text' }, { name: 'Invoice', ...ref('Invoices') }, { name: 'Time Sheet', ...ref('Time Sheets') },
    ] },
    { name: 'Expenses', fields: [
      { name: 'Date', type: 'date' }, { name: 'Project', ...ref('Projects'), required: true }, { name: 'Staff', ...ref('Staff') },
      { name: 'Category', type: 'text' }, { name: 'Description', type: 'text' }, { name: 'Amount', type: 'currency' },
      { name: 'Rechargeable', type: 'yesno', default: true }, { name: 'Invoice', ...ref('Invoices') },
    ] },
    { name: 'Invoices', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Project', ...ref('Projects'), required: true }, { name: 'Client', ...ref('Clients') },
      { name: 'Invoice Date', type: 'date' }, { name: 'Due Date', type: 'date' }, { name: 'Net', type: 'currency' }, { name: 'VAT', type: 'currency' },
      { name: 'Total', type: 'currency' }, { name: 'Paid', type: 'currency', default: 0 }, { name: 'Status', type: 'text', default: 'Sent' },
    ] },
    { name: 'Invoice Lines', fields: [
      { name: 'Invoice', ...ref('Invoices'), required: true }, { name: 'Description', type: 'text' }, { name: 'Kind', type: 'text' },
      { name: 'Quantity', type: 'number', default: 1 }, { name: 'Unit Price', type: 'currency' }, { name: 'Tax Rate', type: 'number', default: 0.2 },
    ] },
    { name: 'Payments', fields: [
      { name: 'Invoice', ...ref('Invoices'), required: true }, { name: 'Received On', type: 'date' }, { name: 'Amount', type: 'currency' },
      { name: 'Method', type: 'text', default: 'Bank transfer' }, { name: 'Reference', type: 'text' },
    ] },
  ],
  queries: {
    'Unbilled time': { tables: ['Time Entries', 'Projects', 'Clients', 'Staff'], columns: [
      { table: 'Clients', field: 'Name', as: 'Client', total: 'group', sort: 'asc' }, { table: 'Projects', field: 'Name', as: 'Project', total: 'group', sort: 'asc' },
      { expr: PERSON, as: 'Person', total: 'group' }, { table: 'Time Entries', field: 'Date', as: 'Earliest', total: 'min' },
      { table: 'Time Entries', field: 'Hours', as: 'Hours', total: 'sum' }, { expr: AT_RATE, as: 'Value', total: 'sum' },
      { table: 'Projects', field: 'Kind', total: 'where', criteria: ['Hourly'] },
      { table: 'Time Entries', field: 'Billable', total: 'where', criteria: ['Yes'] }, { table: 'Time Entries', field: 'Invoice', total: 'where', criteria: ['Is Null'] },
    ] },
    'Unbilled expenses': { tables: ['Expenses', 'Projects', 'Clients'], columns: [
      { table: 'Clients', field: 'Name', as: 'Client', sort: 'asc' }, { table: 'Projects', field: 'Name', as: 'Project' },
      { table: 'Expenses', field: 'Date', sort: 'asc' }, { table: 'Expenses', field: 'Category' }, { table: 'Expenses', field: 'Description' },
      { table: 'Expenses', field: 'Amount' },
      { table: 'Expenses', field: 'Rechargeable', criteria: ['Yes'], show: false }, { table: 'Expenses', field: 'Invoice', criteria: ['Is Null'], show: false },
    ] },
    'Project budgets': { tables: ['Time Entries', 'Projects', 'Clients', 'Staff'], columns: [
      { table: 'Clients', field: 'Name', as: 'Client', total: 'group', sort: 'asc' }, { table: 'Projects', field: 'Code', total: 'group' },
      { table: 'Projects', field: 'Name', as: 'Project', total: 'group' }, { table: 'Projects', field: 'Kind', total: 'group', criteria: [NOT_INTERNAL] },
      { table: 'Projects', field: 'Due Date', total: 'group', sort: 'asc' }, { table: 'Projects', field: 'Budget Hours', total: 'group' },
      { table: 'Time Entries', field: 'Hours', as: 'Hours Used', total: 'sum' }, { table: 'Projects', field: 'Fee', total: 'group' },
      { expr: AT_RATE, as: 'Time at Rates', total: 'sum' }, { expr: COST, as: 'Cost of Time', total: 'sum' },
      { table: 'Projects', field: 'Status', total: 'where', criteria: ['In ("Active", "On hold")'] },
    ] },
    'Task hours against estimate': { tables: ['Time Entries', 'Tasks', 'Projects'], columns: [
      { table: 'Projects', field: 'Code', total: 'group', sort: 'asc' }, { table: 'Projects', field: 'Name', as: 'Project', total: 'group' },
      { table: 'Tasks', field: 'Name', as: 'Task', total: 'group' }, { table: 'Tasks', field: 'Role', total: 'group' },
      { table: 'Tasks', field: 'Estimated Hours', total: 'group' }, { table: 'Time Entries', field: 'Hours', as: 'Hours Used', total: 'sum' },
      { table: 'Projects', field: 'Status', total: 'where', criteria: ['In ("Active", "On hold")'] }, { table: 'Projects', field: 'Kind', total: 'where', criteria: [NOT_INTERNAL] },
    ] },
    'Hours by person, last 30 days': { tables: ['Time Entries', 'Staff'], columns: [
      { expr: PERSON, as: 'Person', total: 'group', sort: 'asc' }, { table: 'Staff', field: 'Role', total: 'group' },
      { table: 'Time Entries', field: 'Hours', as: 'Hours', total: 'sum' }, { expr: 'IIf([Billable], [Hours], 0)', as: 'Billable Hours', total: 'sum' },
      { expr: 'Round(IIf([Billable], [Hours] * [Charge-out Rate], 0), 2)', as: 'Billable Value', total: 'sum' },
      { expr: 'DateDiff("d", [Date], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 30'] },
    ] },
    'Time sheet': { tables: ['Time Entries', 'Tasks', 'Projects', 'Clients', 'Staff'], columns: [
      { expr: PERSON, as: 'Person', sort: 'asc' }, { table: 'Time Entries', field: 'Date', sort: 'asc' },
      { table: 'Clients', field: 'Name', as: 'Client' }, { table: 'Projects', field: 'Name', as: 'Project' }, { table: 'Tasks', field: 'Name', as: 'Task' },
      { table: 'Time Entries', field: 'Hours' }, { table: 'Time Entries', field: 'Billable' }, { table: 'Time Entries', field: 'Note' },
      { expr: 'DateDiff("d", [Date], Date())', as: 'Days Ago', criteria: ['Between 0 And 13'], show: false },
    ] },
    'Overdue invoices': { tables: ['Invoices', 'Projects', 'Clients'], columns: [
      { table: 'Invoices', field: 'Number', as: 'Invoice' }, { table: 'Clients', field: 'Name', as: 'Client' }, { table: 'Projects', field: 'Name', as: 'Project' },
      { table: 'Invoices', field: 'Invoice Date' }, { table: 'Invoices', field: 'Due Date' }, { table: 'Invoices', field: 'Total' }, { table: 'Invoices', field: 'Paid' },
      { expr: 'Round([Total] - [Paid], 2)', as: 'Outstanding', criteria: ['> 0'] },
      { expr: 'DateDiff("d", [Invoices].[Due Date], Date())', as: 'Days Overdue', criteria: ['> 0'], sort: 'desc' },
      { table: 'Clients', field: 'Phone', as: 'Client Phone' },
    ] },
    'Unpaid invoices': { tables: ['Invoices', 'Projects', 'Clients'], columns: [
      { table: 'Clients', field: 'Name', as: 'Client', sort: 'asc' }, { table: 'Invoices', field: 'Number', as: 'Invoice' }, { table: 'Projects', field: 'Name', as: 'Project' },
      { table: 'Invoices', field: 'Invoice Date' }, { table: 'Invoices', field: 'Due Date', sort: 'asc' }, { table: 'Invoices', field: 'Total' },
      { expr: 'Round([Total] - [Paid], 2)', as: 'Outstanding', criteria: ['> 0'] }, { expr: 'DateDiff("d", [Invoices].[Due Date], Date())', as: 'Days Overdue' },
    ] },
    'Revenue by client': { tables: ['Invoices', 'Projects', 'Clients'], columns: [
      { table: 'Clients', field: 'Name', as: 'Client', total: 'group' }, { table: 'Clients', field: 'Sector', total: 'group' },
      { expr: 'Round(IIf(Year([Invoice Date]) = Year(Date()) - 1, [Net], 0), 2)', as: 'Last Year', total: 'sum' },
      { expr: 'Round(IIf(Year([Invoice Date]) = Year(Date()), [Net], 0), 2)', as: 'This Year', total: 'sum', sort: 'desc' },
      { table: 'Invoices', field: 'ID', as: 'Invoices', total: 'count' },
    ] },
    'Revenue by month': { tables: ['Invoices'], columns: [
      { expr: 'Format([Invoice Date], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Invoices', field: 'ID', as: 'Invoices', total: 'count' },
      { table: 'Invoices', field: 'Net', as: 'Net', total: 'sum' }, { table: 'Invoices', field: 'VAT', as: 'VAT', total: 'sum' },
      { table: 'Invoices', field: 'Total', as: 'Total', total: 'sum' },
    ] },
    'Income by kind': { tables: ['Invoice Lines', 'Invoices'], columns: [
      { expr: 'Format([Invoice Date], "yyyy")', as: 'Year', total: 'group', sort: 'asc' }, { table: 'Invoice Lines', field: 'Kind', total: 'group', sort: 'asc' },
      { table: 'Invoice Lines', field: 'ID', as: 'Lines', total: 'count' }, { expr: LINE, as: 'Amount', total: 'sum' },
    ] },
    'Cash received by month': { tables: ['Payments'], columns: [
      { expr: 'Format([Received On], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Payments', field: 'ID', as: 'Payments', total: 'count' },
      { table: 'Payments', field: 'Amount', as: 'Received', total: 'sum' },
    ] },
    'Project profit': { tables: ['Time Entries', 'Projects', 'Clients', 'Staff'], columns: [
      { table: 'Projects', field: 'Finished On', total: 'group', sort: 'desc' }, { table: 'Projects', field: 'Code', total: 'group' },
      { table: 'Projects', field: 'Name', as: 'Project', total: 'group' }, { table: 'Clients', field: 'Name', as: 'Client', total: 'group' },
      { table: 'Projects', field: 'Kind', total: 'group', criteria: [NOT_INTERNAL] }, { table: 'Projects', field: 'Fee', total: 'group' },
      { table: 'Time Entries', field: 'Hours', as: 'Hours', total: 'sum' }, { expr: AT_RATE, as: 'Time at Rates', total: 'sum' },
      { expr: COST, as: 'Cost of Time', total: 'sum' }, { table: 'Projects', field: 'Status', total: 'where', criteria: ['Completed'] },
    ] },
    'Projects due': { tables: ['Projects', 'Clients'], columns: [
      { table: 'Projects', field: 'Due Date', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Due Date])', as: 'Days Left', criteria: ['<= 30'] },
      { table: 'Projects', field: 'Code' }, { table: 'Projects', field: 'Name', as: 'Project' }, { table: 'Clients', field: 'Name', as: 'Client' },
      { table: 'Projects', field: 'Kind' }, { table: 'Projects', field: 'Start Date' }, { table: 'Projects', field: 'Status', criteria: ['In ("Active", "On hold")'] },
    ] },
    'Proposals waiting': { tables: ['Projects', 'Clients'], columns: [
      { table: 'Projects', field: 'Proposed On', sort: 'asc' }, { expr: 'DateDiff("d", [Proposed On], Date())', as: 'Days Waiting' },
      { table: 'Projects', field: 'Code' }, { table: 'Projects', field: 'Name', as: 'Project' }, { table: 'Clients', field: 'Name', as: 'Client' },
      { table: 'Projects', field: 'Kind' }, { table: 'Projects', field: 'Budget Hours' }, { table: 'Projects', field: 'Fee' },
      { table: 'Projects', field: 'Start Date', as: 'Could Start' }, { table: 'Projects', field: 'Status', criteria: ['Proposal'], show: false },
    ] },
    'Recent invoice lines': { tables: ['Invoice Lines', 'Invoices', 'Projects', 'Clients'], columns: [
      { table: 'Invoices', field: 'Number', as: 'Invoice', sort: 'asc' }, { table: 'Invoices', field: 'Invoice Date' }, { table: 'Clients', field: 'Name', as: 'Client' },
      { table: 'Projects', field: 'Name', as: 'Project' }, { table: 'Invoice Lines', field: 'Description' }, { table: 'Invoice Lines', field: 'Quantity' },
      { table: 'Invoice Lines', field: 'Unit Price' }, { expr: LINE, as: 'Amount' },
      { expr: 'DateDiff("d", [Invoice Date], Date())', as: 'Days Ago', criteria: ['Between 0 And 45'], show: false },
    ] },
    'Client contacts': { tables: ['Contacts', 'Clients'], columns: [
      { table: 'Clients', field: 'Name', as: 'Client', sort: 'asc' }, { table: 'Contacts', field: 'Main Contact', sort: 'desc' },
      { table: 'Contacts', field: 'First Name' }, { table: 'Contacts', field: 'Last Name' }, { table: 'Contacts', field: 'Job Title' },
      { table: 'Contacts', field: 'E-mail' }, { table: 'Contacts', field: 'Phone' },
    ] },
  },
  // What the studio prints: one document per record of the table named, its lines from the records that point at it.
  documents: [
    { name: 'Invoice', table: 'Invoices', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Invoice {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Client', label: 'Bill to' }, { field: 'Number', label: 'Invoice' },
        { field: 'Project.Contact', label: 'For the attention of' }, { field: 'Invoice Date', label: 'Date' },
        { field: 'Client.Address', label: 'Address' }, { field: 'Due Date', label: 'Payment due' },
        { field: 'Client.City', label: 'City' }, { field: 'Project.Code', label: 'Our reference' },
        { field: 'Client.Postcode', label: 'Postcode' }, { field: 'Project.Client PO', label: 'Your order number' },
        { field: 'Client.Country', label: 'Country' }, { field: 'Project.Name', label: 'Project' },
      ] },
      { type: 'lines', table: 'Invoice Lines', columns: [
        { field: 'Description' }, { field: 'Quantity', label: 'Hours or qty', type: 'number' }, { field: 'Unit Price', label: 'Rate', type: 'currency' },
        { expr: LINE, as: 'Amount', type: 'currency' }, { field: 'Tax Rate', label: 'VAT rate', type: 'number' }, { expr: LINE_VAT, as: 'VAT', type: 'currency' },
      ] },
      { type: 'totals', rows: [
        { label: 'Fees and expenses', expr: '[Amount]' }, { label: 'VAT', expr: '[VAT]' },
        { label: 'Total', expr: '[Fees and expenses] + [VAT]', strong: true }, { label: 'Paid', expr: '[Paid]' },
        { label: 'Balance due', expr: 'Round([Total] - [Paid], 2)', strong: true },
      ] },
      { type: 'text', text: 'Please pay by {Due Date}. {Business.Payment Terms} Pay to {Business.Bank Details}, quoting {Number}. VAT registration {Business.Tax Number}.', small: true },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    { name: 'Proposal', table: 'Projects', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Proposal {Quote No}: {Name}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Client', label: 'For' }, { field: 'Quote No', label: 'Proposal' },
        { field: 'Contact', label: 'For the attention of' }, { field: 'Proposed On', label: 'Date' },
        { field: 'Client.Address', label: 'Address' }, { field: 'Code', label: 'Our reference' },
        { field: 'Client.City', label: 'City' }, { field: 'Work Type', label: 'The work' },
        { field: 'Client.Postcode', label: 'Postcode' }, { field: 'Kind', label: 'Charged as' },
        { field: 'Client.Country', label: 'Country' }, { field: 'Start Date', label: 'Could start' },
      ] },
      { type: 'text', text: 'Thank you for asking us to quote. Below is the work as we would plan it, in its parts, with the hours we allow for each. We could start on {Start Date} and finish by {Due Date}, given your comments at each stage within a week of asking.' },
      { type: 'lines', table: 'Tasks', columns: [{ field: 'Name', label: 'Part of the work' }, { field: 'Role', label: 'Who does it' }, { field: 'Estimated Hours', label: 'Hours', type: 'number' }] },
      { type: 'totals', rows: [{ label: 'Hours planned', expr: '[Estimated Hours]', type: 'number' }, { label: 'Fee, before VAT', expr: '[Fee]', strong: true }] },
      { type: 'text', text: 'A fixed fee covers the work above with two rounds of changes at each stage, and is invoiced in stages as the work goes. Work charged by the hour is invoiced monthly for the time spent, and the fee above is then our estimate. VAT is added for clients in the UK. Printing, licences, hosting and travel are passed on at cost. This proposal holds for 60 days from its date.' },
      { type: 'text', text: 'To go ahead, reply to this proposal and we will send a letter of engagement to sign. {Business.Document Footer}', small: true },
    ] },
    { name: 'Statement', table: 'Clients', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Statement of account' },
      { type: 'fields', columns: 2, items: [{ field: 'Name', label: 'Client' }, { field: 'Address' }, { field: 'City' }, { field: 'Postcode' }, { field: 'Country' }, { field: 'Payment Days', label: 'Payment terms, days' }] },
      { type: 'lines', table: 'Invoices', link: 'Client', sort: 'Invoice Date', columns: [
        { field: 'Number', label: 'Invoice' }, { field: 'Invoice Date', label: 'Date' }, { field: 'Due Date', label: 'Due' }, { field: 'Project.Name', label: 'Project' },
        { field: 'Total', label: 'Invoiced', type: 'currency' }, { field: 'Paid', type: 'currency' }, { expr: 'Round([Total] - [Paid], 2)', as: 'Owing', type: 'currency' },
        { expr: 'IIf([Owing] > 0 And DateDiff("d", [Due Date], Date()) > 0, DateDiff("d", [Due Date], Date()), 0)', as: 'Days Overdue', label: 'Days overdue', type: 'integer' },
      ] },
      { type: 'totals', rows: [{ label: 'Invoiced', expr: '[Total]' }, { label: 'Received', expr: '[Paid]' }, { label: 'Balance owing', expr: '[Owing]', strong: true }] },
      { type: 'text', text: 'Invoices with nothing owing are shown for your records. Please pay the balance owing to {Business.Bank Details}, quoting the invoice numbers. If you think anything here is wrong, call us on {Business.Phone}.', small: true },
    ] },
    { name: 'Weekly time sheet', table: 'Time Sheets', paper: 'a4', letterhead: { table: 'Business', fields: ['Name'] }, blocks: [
      { type: 'heading', text: 'Time sheet {Number}' },
      { type: 'fields', columns: 3, items: [
        { field: 'Staff', label: 'Name' }, { field: 'Staff.Role', label: 'Role' }, { field: 'Week Commencing', label: 'Week commencing' },
        { field: 'Staff.Hours per Week', label: 'Contracted hours' }, { field: 'Status' }, { field: 'Approved On', label: 'Approved on' },
      ] },
      { type: 'lines', table: 'Time Entries', sort: 'Date', columns: [
        { field: 'Date' }, { field: 'Project.Client', label: 'Client' }, { field: 'Project.Code', label: 'Code' }, { field: 'Project.Name', label: 'Project' },
        { field: 'Task.Name', label: 'Task' }, { field: 'Note' }, { field: 'Hours', type: 'number' }, { expr: 'IIf([Billable], [Hours], 0)', as: 'Billable', type: 'number' },
      ] },
      { type: 'totals', rows: [
        { label: 'Hours worked', expr: '[Hours]', type: 'number' }, { label: 'Billable hours', expr: '[Billable]', type: 'number' },
        { label: 'Billable share, percent', expr: 'IIf([Hours] > 0, Round(100 * [Billable] / [Hours], 0), 0)', type: 'number' },
      ] },
      { type: 'text', text: 'I confirm these are the hours I worked this week.', small: true },
      { type: 'signatures', labels: ['Signed', 'Approved by'] },
    ] },
    { name: 'Letter of engagement', table: 'Projects', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Letter of engagement: {Name}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Client', label: 'Client' }, { field: 'Code', label: 'Our reference' },
        { field: 'Contact', label: 'For the attention of' }, { field: 'Quote No', label: 'Proposal' },
        { field: 'Client.Address', label: 'Address' }, { field: 'Start Date', label: 'Starting' },
        { field: 'Client.City', label: 'City' }, { field: 'Due Date', label: 'Finishing by' },
        { field: 'Client.Postcode', label: 'Postcode' }, { field: 'Work Type', label: 'The work' },
      ] },
      { type: 'text', text: 'Dear {Contact.First Name},' },
      { type: 'text', text: 'Thank you for asking us to do this work. This letter sets out what we will do, what it costs, how we invoice it, who owns the result and how either of us can bring it to an end, so that we both know where we stand. Please sign a copy and send it back to us before the work starts.' },
      { type: 'text', text: '1. The work. We will carry out the {Work Type} work set out in our proposal {Quote No}, which goes with this letter. We plan to start on {Start Date} and to finish by {Due Date}. Those dates depend on your comments reaching us when the plan says; if they are late, or you ask us to pause, the finish date moves on by the same time.' },
      { type: 'text', text: '2. The fee. This work is charged as: {Kind}. It is planned at {Budget Hours} hours and the fee is {Fee}, plus VAT where it applies. A fixed fee covers the work in the proposal with two rounds of changes at each stage; changes beyond those, or work you add, we quote for before we do them. Work by the hour is charged in quarters of an hour for the time we spend, at the rates in the proposal; its fee is our estimate, and we will tell you before the time used goes past it.' },
      { type: 'text', text: '3. Costs we pay for you. Printing, licences for photographs and fonts, hosting, couriers and travel to see you are passed on at what they cost us and shown on the invoice. We will ask you before taking on any single large cost.' },
      { type: 'text', text: '4. How we invoice. A fixed fee is invoiced in stages: the first part when the work starts, the next as the proposal sets out, and the last on completion. Work by the hour is invoiced at the start of each month for the month before. Each invoice is due {Client.Payment Days} days after its date, paid by bank transfer to {Business.Bank Details}, quoting the invoice number. If an invoice is more than 30 days overdue we may pause the work until it is paid.' },
      { type: 'text', text: '5. Who owns the work. Once every invoice for this work has been paid in full, the rights in the final work we deliver pass to you. Until then the work stays ours and you may use it only to review it. Ideas and drafts you did not choose stay ours. Photographs, fonts and software made by others are licensed to you on their makers\' terms, which we will hand over with the files. We may show the finished work in our portfolio unless you ask us not to.' },
      { type: 'text', text: '6. Keeping things private. We keep to ourselves anything you tell us about your business that is not public, and ask the same of you for what we tell you about our methods and prices.' },
      { type: 'text', text: '7. Ending the work. Either of us may end this engagement by writing to the other with 14 days\' notice. You then pay for the work done up to the day it ends, at our hourly rates and, for fixed-fee work, never more than the fee, and for any costs already taken on for you. The rights pass to you in the work that has been paid for.' },
      { type: 'text', text: '8. Keeping in touch. We will name one person to look after the work day to day, and ask you to do the same. Approvals are best given by e-mail, so that there is a record of them.' },
      { type: 'text', text: 'If this is as you understand it, please sign and return a copy. We look forward to working with you.' },
      { type: 'text', text: 'Yours sincerely, for {Business.Name}' },
      { type: 'spacer', size: 24 },
      { type: 'signatures', labels: ['For the studio', 'For the client'] },
    ] },
  ],
  reports: [
    { name: 'Time sheet', title: 'Time sheet, last 14 days', query: 'Time sheet', group: 'Person', sort: 'Date', columns: [{ field: 'Date', type: 'date' }, 'Client', 'Project', 'Task', { field: 'Hours', type: 'number' }, 'Note'], totals: [{ field: 'Hours', fn: 'sum' }] },
    { name: 'Work in progress', title: 'Work in progress to invoice', query: 'Unbilled time', group: 'Client', columns: ['Project', 'Person', { field: 'Earliest', type: 'date' }, { field: 'Hours', type: 'number' }, { field: 'Value', type: 'currency' }], totals: [{ field: 'Hours', fn: 'sum' }, { field: 'Value', fn: 'sum' }] },
    { name: 'Invoices due', query: 'Unpaid invoices', group: 'Client', columns: ['Invoice', 'Project', { field: 'Invoice Date', type: 'date' }, { field: 'Due Date', type: 'date' }, { field: 'Days Overdue', type: 'integer' }, { field: 'Total', type: 'currency' }, { field: 'Outstanding', type: 'currency' }], totals: [{ field: 'Outstanding', fn: 'sum' }] },
    { name: 'Project budget report', query: 'Project budgets', group: 'Client', columns: ['Code', 'Project', 'Kind', { field: 'Due Date', type: 'date' }, { field: 'Budget Hours', type: 'number' }, { field: 'Hours Used', type: 'number' }, { field: 'Fee', type: 'currency' }, { field: 'Time at Rates', type: 'currency' }, { field: 'Cost of Time', type: 'currency' }], totals: [{ field: 'Hours Used', fn: 'sum' }, { field: 'Fee', fn: 'sum' }, { field: 'Time at Rates', fn: 'sum' }, { field: 'Cost of Time', fn: 'sum' }] },
    { name: 'Revenue by client', query: 'Revenue by client', columns: ['Client', 'Sector', { field: 'Invoices', type: 'integer' }, { field: 'Last Year', type: 'currency' }, { field: 'This Year', type: 'currency' }], totals: [{ field: 'Last Year', fn: 'sum' }, { field: 'This Year', fn: 'sum' }] },
    { name: 'Project profit', title: 'Completed projects: fee against cost of time', query: 'Project profit', columns: [{ field: 'Finished On', type: 'date' }, 'Code', 'Project', 'Client', 'Kind', { field: 'Hours', type: 'number' }, { field: 'Fee', type: 'currency' }, { field: 'Cost of Time', type: 'currency' }], totals: [{ field: 'Fee', fn: 'sum' }, { field: 'Cost of Time', fn: 'sum' }] },
    { name: 'Invoices issued', title: 'Invoices issued, last 45 days', query: 'Recent invoice lines', group: 'Invoice', columns: [{ field: 'Invoice Date', type: 'date' }, 'Client', 'Description', { field: 'Quantity', type: 'number' }, { field: 'Unit Price', type: 'currency' }, { field: 'Amount', type: 'currency' }], totals: [{ field: 'Amount', fn: 'sum' }] },
  ],
};

export const about = {
  name: 'Bramble Studio',
  kind: 'Consultancy or agency',
  line: 'Clients, projects, time sheets, expenses and invoices for a small studio or consultancy.',
  summary: [
    'For a small firm that sells its time: a design studio, a consultancy, a bookkeeper or anyone who quotes for work, logs hours against it and invoices for it. It answers the questions such a firm asks each week: what time has not been billed yet, which projects are eating their budget, who is owed what and for how long.',
    'Bramble Studio is invented. Six people work there: two partners, a senior designer, a designer, a developer and a project manager. They make brand identities, websites, annual reports, packaging and campaigns for about two dozen small and mid-sized clients, most in the UK and a few in Ireland and the US. Some work is for a fixed fee paid in stages; some is by the hour, invoiced at the start of each month for the month before, including three retainers that renew.',
    'The records cover the last sixteen months, counted back from the day the file is made. Proposals come in every few days and about two in three are won, fewer when the studio is already busy. Every hour is logged to a task, including the studio\'s own time on admin and new business, so the billable share of each person\'s week shows. Some projects overrun, a few are paused by the client for a while, some invoices are paid late or in two parts, and the studio is quiet in August and shut over Christmas.',
    'It prints what a studio sends out: the invoice, with time, fee stages and expenses on separate lines and VAT worked out line by line; a proposal listing the parts of the work and the hours allowed for each; a statement of a client\'s invoices and what is still owed; each person\'s weekly time sheet to sign; and a letter of engagement setting out the work, the fee, payment and who owns the result.',
  ].join('\n\n'),
  suits: [
    'A freelance designer or developer', 'A marketing or branding agency', 'An architect\'s practice', 'IT or management consultants',
    'An accountant or bookkeeper who bills by the hour', 'A small law practice', 'A translation or copywriting business', 'A video or photography studio',
  ],
  tables: {
    'Business': BUSINESS_ABOUT,
    'Clients': { about: 'One record per client the studio works for or has quoted for. The studio itself is kept here too, as the client of its own internal projects.', fields: {
      'Name': 'The client\'s business name, as it goes on an invoice.',
      'Sector': 'The line of business the client is in, such as Charity or Retail; Internal for the studio itself.',
      'Address': 'The street address invoices are sent to.', 'City': 'The town or city of the address.', 'Postcode': 'The postcode or ZIP of the address.', 'Country': 'The country of the address.',
      'Phone': 'The client\'s main phone number.', 'E-mail': 'Where invoices are sent, usually an accounts address.',
      'Payment Days': 'How many days after the invoice date payment is due: the client\'s terms.',
      'VAT Rate': 'The VAT charged on this client\'s invoices as a fraction: 0.2 for UK clients, 0 for clients abroad.',
      'Client Since': 'When the client first came to the studio.', 'Notes': 'Anything worth knowing, such as a client who pays late.',
    } },
    'Contacts': { about: 'The people the studio deals with at each client. A client can have several.', fields: {
      'Client': 'The client the person works for.', 'First Name': 'Their first name.', 'Last Name': 'Their family name.',
      'Job Title': 'What they do there, so you know who signs off work and who pays.', 'E-mail': 'Their e-mail address.', 'Phone': 'Their direct phone number.',
      'Main Contact': 'Yes for the person to speak to first about the work.',
    } },
    'Staff': { about: 'Everyone who logs time. What each hour costs and what it is charged at turn hours into money in the queries.', fields: {
      'First Name': 'Their first name.', 'Last Name': 'Their family name.', 'Role': 'Their job in the studio.',
      'E-mail': 'Their work e-mail address.', 'Phone': 'Their work mobile.',
      'Hourly Cost': 'What an hour of this person costs the studio, with pay, pension and a share of overheads.',
      'Charge-out Rate': 'What an hour of this person is charged to a client on hourly work.',
      'Hours per Week': 'Their contracted hours, for checking time sheets against.', 'Started': 'The day they joined.', 'Active': 'No once they have left; their time stays in the records.',
    } },
    'Projects': { about: 'One record per piece of work quoted for, from proposal to completion or loss. The studio\'s own time (admin, new business, training) is kept on three internal projects.', fields: {
      'Code': 'A short reference for the project, used on time sheets and in conversation.', 'Name': 'What the work is, as the client knows it.',
      'Client': 'Who the work is for.', 'Kind': 'Fixed fee, Hourly or Internal. It decides how the work is invoiced.',
      'Work Type': 'The kind of job, such as Brand identity or Website design and build, for comparing like with like.',
      'Budget Hours': 'The hours the fee or estimate was worked out from.',
      'Fee': 'For fixed-fee work, the agreed fee. For hourly work, the estimate quoted; once the project is completed it becomes what the billable time came to.',
      'Proposed On': 'When the proposal went to the client.', 'Start Date': 'When the work starts, or could start if the proposal is won.',
      'Due Date': 'When the work is due to be finished. A pause moves it on.', 'Finished On': 'The day the last of the work was done.',
      'Status': 'Proposal, Active, On hold, Completed or Lost.', 'Notes': 'Why a proposal was lost, why work is paused, and the like.',
      'Quote No': 'The number of the proposal sent for the work, given in the order proposals go out, a new run each year: QUO-2026-0042. Printed on the proposal and the letter of engagement.',
      'Client PO': 'The client\'s own purchase order number, for clients who pay only against one. Printed on every invoice for the project.',
      'Contact': 'The person at the client the work is for. Invoices, proposals and the letter of engagement are addressed to them.',
    } },
    'Tasks': { about: 'The parts a project is broken into, each with the hours estimated for it, so you can see which part ran over.', fields: {
      'Project': 'The project the task belongs to.', 'Name': 'What the task is, such as Concepts or Front-end build.',
      'Role': 'Who does it: Partner, Design, Development or Project management.', 'Estimated Hours': 'The hours allowed for it in the budget.',
      'Done': 'Yes when the task is finished.',
    } },
    'Time Sheets': { about: 'One for each person for each week they log time. Their time entries for the week point at it, it is printed for them to sign, and it is approved before the month\'s invoices go out.', fields: {
      'Number': 'The time sheet\'s number, given as each week\'s sheets are opened, a new run each year: TS-2026-0042.', 'Staff': 'Whose week it is.',
      'Week Commencing': 'The Monday the week starts on.', 'Status': 'Open while the week is being filled in, Submitted once the person has sent it in, Approved once checked.',
      'Submitted On': 'The day the person sent it in, usually the Friday or the Monday after.', 'Approved On': 'The day the project manager or a partner approved it.',
    } },
    'Time Entries': { about: 'Every stretch of time someone worked, logged to a project and one of its tasks. The heart of the database.', fields: {
      'Date': 'The day the work was done.', 'Staff': 'Who did it.', 'Project': 'The project it was for.',
      'Task': 'The task of that project it was for.', 'Hours': 'How long, in hours and quarters of an hour.',
      'Billable': 'Yes for client work; No for the studio\'s own time and for corrections done at the studio\'s cost.',
      'Note': 'A word on what was done.',
      'Invoice': 'The invoice the hours were billed on. Empty means not billed yet. On fixed-fee work it is the stage invoice that covered them.',
      'Time Sheet': 'The person\'s time sheet for the week the work was done in.',
    } },
    'Expenses': { about: 'Money spent on a project: licences, printing, travel, hosting. Rechargeable ones are passed on to the client at cost.', fields: {
      'Date': 'When the cost was run up.', 'Project': 'The project it was for.', 'Staff': 'Who spent it.',
      'Category': 'The kind of cost, such as Printing or Travel.', 'Description': 'What it was.', 'Amount': 'What it cost, before VAT.',
      'Rechargeable': 'Yes if the client pays for it on their next invoice.', 'Invoice': 'The invoice it was recharged on. Empty means not recharged yet.',
    } },
    'Invoices': { about: 'Every invoice sent, with what was paid on it so far.', fields: {
      'Number': 'The invoice number, given in date order and once only, a new run each year: INV-2026-0042.', 'Project': 'The project billed.',
      'Client': 'Who the invoice is addressed to: the project\'s client. Kept on the invoice so that a client\'s statement can list its invoices.',
      'Invoice Date': 'The date on the invoice.', 'Due Date': 'When payment is due, from the client\'s payment days.',
      'Net': 'The total of the lines before VAT.', 'VAT': 'The VAT on each line, rounded to the penny, added up: the same figure the printed invoice shows.', 'Total': 'Net plus VAT: what the client owes.',
      'Paid': 'What has been received against it so far, from Payments.', 'Status': 'Sent, Part paid or Paid.',
    } },
    'Invoice Lines': { about: 'The lines of each invoice: hours by role on hourly work, a stage of the fee on fixed-fee work, and expenses passed on.', fields: {
      'Invoice': 'The invoice the line is on.', 'Description': 'What the line says, as printed.', 'Kind': 'Time, Fee or Expense.',
      'Quantity': 'Hours for a Time line; 1 for a fee stage or an expense.', 'Unit Price': 'The rate per hour, or the amount of the stage or expense.',
      'Tax Rate': 'The VAT rate of the line as a fraction, taken from the client when the invoice is raised: 0.2 for UK clients, 0 for clients abroad.',
    } },
    'Payments': { about: 'Money received from clients. An invoice can be paid in more than one payment.', fields: {
      'Invoice': 'The invoice the money was for.', 'Received On': 'The day it arrived.', 'Amount': 'How much arrived.',
      'Method': 'Bank transfer or Card.', 'Reference': 'The reference the client put on the payment, usually the invoice number.',
    } },
  },
  queries: {
    'Unbilled time': 'Billable hours on hourly projects not yet on an invoice, by client, project and person, with what they are worth at charge-out rates. Run it before the monthly invoices go out.',
    'Unbilled expenses': 'Rechargeable costs not yet passed on to the client, oldest first, so nothing is forgotten on the next invoice.',
    'Project budgets': 'Each project in hand: its budget hours against the hours used so far, and its fee against the time used at charge-out rates and at cost. A project whose time at rates is near its fee needs a word with the client.',
    'Task hours against estimate': 'For each project in hand, each task\'s estimated hours against the hours logged to it, to see which part of the work is running over.',
    'Hours by person, last 30 days': 'Each person\'s hours over the last 30 days, how many of them were billable and what those were worth. Use it to see who is stretched and who has room.',
    'Time sheet': 'Every entry of the last 14 days, by person and day, with the client, project and task. For checking time sheets before invoicing.',
    'Overdue invoices': 'Invoices past their due date with money still owed, the longest overdue first, with the client\'s phone number for chasing.',
    'Unpaid invoices': 'Every invoice with money still owed, by client, with the days it is overdue (below nought when it is not due yet).',
    'Revenue by client': 'Each client\'s invoiced fees before VAT, last year and this year, the biggest this year first. Shows who the studio depends on.',
    'Revenue by month': 'What was invoiced each month, net, VAT and total, for seeing the busy and quiet months.',
    'Income by kind': 'Invoiced amounts each year split into fixed-fee stages, hourly time and recharged expenses.',
    'Cash received by month': 'The money that actually came in each month, which is what pays the wages.',
    'Project profit': 'Completed projects, newest first: the fee earned against the cost of the time spent, and the time at charge-out rates. Fee less cost of time is what the project made towards the studio\'s other costs.',
    'Projects due': 'Projects in hand due within the next 30 days, and any already late (Days Left below nought), soonest first.',
    'Proposals waiting': 'Proposals sent and not yet answered, the oldest first, with how long each has waited. For a follow-up call.',
    'Recent invoice lines': 'The lines of every invoice sent in the last 45 days, for checking what went out.',
    'Client contacts': 'The people at each client, the main contact first, with their e-mail and phone.',
  },
  reports: {
    'Time sheet': 'The last 14 days of time, one section per person with their total hours. For each person to check, and for the project manager before invoicing.',
    'Work in progress': 'Unbilled hourly time by client, with hours and value, and a total for the studio: what the next round of invoices will bring in.',
    'Invoices due': 'Money owed by each client, invoice by invoice, with the days overdue and the total outstanding. For the partner who chases payment.',
    'Project budget report': 'Projects in hand by client, with budget hours, hours used, fee, time at rates and cost of time. For the weekly project meeting.',
    'Revenue by client': 'One line per client with last year\'s and this year\'s invoiced fees, and the totals.',
    'Project profit': 'Completed projects with their fee and the cost of the time they took, and the totals for the period.',
    'Invoices issued': 'Each invoice of the last 45 days with its lines and its total before VAT.',
  },
  documents: {
    'Invoice': 'The VAT invoice for a record in Invoices: the client and the person it is for, the project, the client\'s order number where they have one, each line (hours by grade, a fee stage or an expense) with its VAT, then the total, what has been paid and the balance due, and how to pay. Numbered INV-2026-0042.',
    'Proposal': 'The quote for a record in Projects: the parts of the work with the hours allowed for each, the total hours and the fee before VAT, when it could start and how it is charged. Print it from the Projects form for a project whose Status is Proposal; the template prints for any project you choose.',
    'Statement': 'A statement for a record in Clients: every invoice sent to the client with what was invoiced, what has been paid and what is still owing on each, the days overdue, and the balance owing. Paid invoices are listed too, for the client\'s records.',
    'Weekly time sheet': 'One person\'s week from Time Sheets: each day\'s entries with client, project, task and hours, the hours worked and the billable share, and lines for the person to sign and for the approver.',
    'Letter of engagement': 'The letter that goes to the client with a won proposal, for a record in Projects: the work and its dates, the fee and how it is charged, costs passed on, how and when invoices are paid, who owns the work once paid for, confidentiality and how either side can end it, with two signature lines. It is a plain starting point written for this sample, not legal advice: have it checked for your own country and trade before you use it.',
  },
  tutorial: [
    { title: 'Quote for new work and send the proposal', steps: [
      'Open the Projects form and press New. Type the Name, choose the Client and the Contact the work is for, and set Kind to Fixed fee or Hourly. Leave Status as Proposal.',
      'Type the Work Type, the Budget Hours you have estimated, the Fee, Proposed On (today) and the Start Date and Due Date you could manage. Give it a Quote No in the next number of the year\'s run, such as QUO-2026-0064.',
      'In the Tasks below, add each part of the work with its Role and Estimated Hours. Their hours should add up to the Budget Hours.',
      'Open the print template Proposal from Print templates, choose the project, and press Print or Save as PDF to send it.',
      'Each week, run the query Proposals waiting to see which proposals have had no answer and how long they have waited.',
    ] },
    { title: 'Win the work and send the letter of engagement', steps: [
      'When the client says yes, open the project in the Projects form and change Status from Proposal to Active. If the client gave you a purchase order, type it in Client PO.',
      'Open the print template Letter of engagement, choose the project, and press Print or Save as PDF. Send it with the proposal for the client to sign and return.',
      'For a fixed fee, raise the first stage invoice on the day the work starts (see Invoice the month\'s work).',
    ] },
    { title: 'Log time and print the week\'s time sheet', steps: [
      'On Monday, open the Time Sheets form and press New for each person: choose the Staff, type Week Commencing (the Monday) and the next Number, such as TS-2026-0252. Status is Open.',
      'Each day, open the Time Entries form and press New for each stretch of work: Date, Staff, Project, Task, Hours (in quarters, such as 2.25), Billable, a Note, and the person\'s Time Sheet for the week.',
      'Time on the studio\'s own work goes to one of the internal projects (INT-1 New business and proposals, INT-2 Studio admin, INT-3 Training and development) with Billable set to No.',
      'On Friday, open the print template Weekly time sheet, choose the person\'s sheet, and print it for them to check and sign. Set Status to Submitted and type Submitted On.',
      'Once checked, the project manager sets Status to Approved and types Approved On. The report Time sheet shows the last 14 days for everyone at once.',
    ] },
    { title: 'Invoice the month\'s work', steps: [
      'In the first days of the month, run the query Unbilled time (or print the report Work in progress) to see the hourly time not yet invoiced, and Unbilled expenses for costs to pass on.',
      'Open the Invoices form and press New. Type the next Number of the year (INV-2026-0111), choose the Project and its Client, and type the Invoice Date and the Due Date (the invoice date plus the client\'s Payment Days).',
      'In the Invoice Lines below, add a line for each grade of person (such as "Designer time, September 2026", Quantity in hours, Unit Price their rate), a fee stage for fixed-fee work, and a line for each expense. Set Tax Rate to the client\'s VAT Rate.',
      'Type the invoice\'s Net (the lines added up), VAT (each line\'s VAT added up) and Total. Then put the invoice in the Invoice field of each time entry and expense it covers, so they drop out of the unbilled queries.',
      'Open the print template Invoice, choose the invoice, and press Print or Save as PDF.',
    ] },
    { title: 'Record a payment and chase what is owed', steps: [
      'When money arrives, open the Payments form and press New: choose the Invoice, type Received On, Amount, Method and the client\'s Reference.',
      'Open the invoice in the Invoices form and bring Paid up to what has been received, and Status to Part paid or Paid.',
      'Run the query Overdue invoices: it lists what is past due, the longest first, with the client\'s phone number.',
      'Open the print template Statement, choose the client, and send it: it lists every invoice with what is owing and the balance.',
    ] },
    { title: 'Keep projects on budget', steps: [
      'Before the weekly meeting, print the report Project budget report: each project in hand with its budget hours, hours used, fee and the cost of the time so far.',
      'For a project near its budget, run the query Task hours against estimate to see which part of the work is running over.',
      'Run Projects due for anything due in the next 30 days or already late.',
      'When a project is finished, set its Status to Completed and type Finished On. The report Project profit then shows its fee against the cost of its time.',
    ] },
  ],
  notes: [
    'The studio is in Clients too, as Bramble Studio with the Sector Internal, so that its own time (admin, new business, training) can be logged to three internal projects. Leave it out of anything you send: it has no invoices.',
    'Numbers are typed, not made for you: invoices INV-, proposals QUO- and time sheets TS-, each a new run each year. Each number can be used only once.',
    'An invoice\'s Net, VAT and Total are kept on the invoice as it was sent. If you change a line afterwards, change these too, or the totals block on the printed invoice and the figures in the revenue queries will differ.',
    'VAT comes from the client: 0.2 for clients in the UK, 0 for clients abroad. Each invoice line carries its own Tax Rate so an old invoice keeps the rate it was sent at.',
    'An hourly project\'s Fee is the estimate while the work goes on; when it is completed the Fee becomes what its billable time came to, so Project profit compares like with like.',
    'Each print template prints for every record of its table unless you choose one: choose the invoice, project, client or time sheet you want before pressing Print. The Proposal and the Letter of engagement are for projects whose Status is Proposal or newly Active.',
    'The Statement lists every invoice the client has had, paid or not; the balance owing is the sum of the Owing column.',
  ],
  faq: [
    { q: 'How do I see what has not been invoiced yet?', a: 'Run the query Unbilled time for hourly work, or print the report Work in progress. Unbilled expenses lists costs to pass on. Fixed-fee work is invoiced in stages instead, so its hours never show there.' },
    { q: 'How do I print an old invoice again?', a: 'Open the print template Invoice and choose the invoice by its number, or open the invoice in the Invoices form and print from there. The template prints whatever is in the record, paid or not.' },
    { q: 'Why does a time entry need a Time Sheet?', a: 'So the week\'s hours can be printed, signed and approved together. Each entry points at its person\'s sheet for that week; open a new sheet each Monday for each person.' },
    { q: 'A client pays in two parts. What do I do?', a: 'Add a Payments record for each part. Bring the invoice\'s Paid up to the total received and set Status to Part paid until the rest arrives, then Paid.' },
    { q: 'How do I see who is busy and who has room?', a: 'Run Hours by person, last 30 days. It shows each person\'s hours, how many were billable and what those were worth, against their Hours per Week in Staff.' },
    { q: 'A client needs their order number on our invoices. Where does it go?', a: 'Type it in Client PO on the project. The Invoice prints it as Your order number on every invoice for that project.' },
    { q: 'Can I use the letter of engagement as it is?', a: 'Treat it as a draft. It is written plainly for this sample and is not legal advice; have it checked for your own country and kind of work, then change the words in the print template Letter of engagement.' },
  ],
  empty: 'Start with Business: your own name, address, VAT number, bank details and the words for the foot of your documents, which every print template uses. Then Staff: each person with their hourly cost and charge-out rate. Then Clients and their Contacts, adding your own firm as a client if you want to record internal time. Add a project for each piece of work, with its kind, budget hours and fee, and break it into Tasks. From then on, log Time Entries and Expenses as the work goes. When you invoice, add the Invoice and its Invoice Lines, and put the invoice in the Invoice field of the time entries and expenses it covers, so they drop out of the unbilled queries. Record each Payment when the money arrives and bring the invoice\'s Paid and Status up to date.',
  adapt: [
    'If rates differ by client, add a Rate field to Projects and use it instead of the person\'s charge-out rate in the Unbilled time query.',
    'A law or accounting practice can rename Projects to Matters and Tasks to Stages, and add a Matter Type and a Fee Earner field.',
    'Add a Payment Link field to Invoices for clients who pay by card online, and a line in the Invoice template that prints it under the bank details.',
    'Add a Leave table, or an internal Holiday project, so time off shows on the time sheet and the hours per week add up.',
    'Add a Status of Invoiced to fixed-fee projects whose last stage has gone out but whose work is not signed off yet.',
    'Add a Mileage expense category with a Miles field, and a calculated column at your mileage rate in the Unbilled expenses query.',
    'Make a query of Time Entries by Work Type to see which kinds of job overrun most, before quoting for the next one.',
  ],
};

/** The team: role, the grade an invoice names, what kind of task each takes, an hour's cost and charge-out rate, and the share of a day spent on client work. */
const TEAM = [
  { role: 'Partner, Creative Director', grade: 'Partner', kind: 'Partner', cost: 52, rate: 145, util: 0.55, day: 8, since: [2900, 3300] },
  { role: 'Partner, Strategy', grade: 'Partner', kind: 'Partner', cost: 52, rate: 145, util: 0.55, day: 8, since: [2900, 3300] },
  { role: 'Senior Designer', grade: 'Senior designer', kind: 'Design', cost: 34, rate: 95, util: 0.85, day: 7.5, since: [1100, 2200] },
  { role: 'Designer', grade: 'Designer', kind: 'Design', cost: 26, rate: 80, util: 0.88, day: 7.5, since: [300, 900] },
  { role: 'Developer', grade: 'Developer', kind: 'Development', cost: 36, rate: 100, util: 0.85, day: 7.5, since: [600, 1600] },
  { role: 'Project Manager', grade: 'Project manager', kind: 'Project management', cost: 30, rate: 80, util: 0.65, day: 7.5, since: [700, 1800] },
];
/** What an hour of each kind of task is charged at when a fee is worked out. */
const RATE = { Partner: 145, Design: 88, Development: 100, 'Project management': 80 };

/**
 * The work the studio sells: its billing, hours, weeks, how often it is asked
 * for and by which sectors (all when not said), its names, its tasks [name,
 * role, share of the hours], its stage payments, and the costs it runs up
 * [description, category, low, high, rechargeable].
 */
const WORK = [
  { type: 'Brand identity', kind: 'Fixed fee', hours: [90, 170], weeks: [7, 12], weight: 4,
    names: ['Brand identity', 'Rebrand', 'Brand refresh', 'Visual identity'],
    tasks: [['Discovery and research', 'Partner', 0.12], ['Concepts', 'Design', 0.3], ['Design development', 'Design', 0.22], ['Brand guidelines', 'Design', 0.14], ['Creative direction', 'Partner', 0.08], ['Project management', 'Project management', 0.14]],
    stages: [0.4, 0.3, 0.3],
    costs: [['Stock photography licence', 'Licences', 60, 240, true], ['Font licence', 'Licences', 80, 400, true], ['Printed proofs', 'Printing', 25, 90, true], ['Train fare to client workshop', 'Travel', 30, 160, true]] },
  { type: 'Website design and build', kind: 'Fixed fee', hours: [200, 380], weeks: [10, 18], weight: 4,
    names: ['New website', 'Website redesign', 'Online shop', 'Booking website'],
    tasks: [['Discovery', 'Partner', 0.06], ['Wireframes and user journeys', 'Design', 0.1], ['Visual design', 'Design', 0.16], ['Creative direction', 'Partner', 0.04], ['Front-end build', 'Development', 0.32], ['Content and CMS set-up', 'Development', 0.12], ['Testing and launch', 'Development', 0.08], ['Project management', 'Project management', 0.12]],
    stages: [0.3, 0.4, 0.3],
    costs: [['Stock photography licence', 'Licences', 60, 300, true], ['Web hosting, first year', 'Hosting', 180, 420, true], ['Plugin licence', 'Licences', 40, 180, true], ['Device testing service, one month', 'Software', 30, 60, false]] },
  { type: 'Annual report', kind: 'Fixed fee', hours: [60, 120], weeks: [5, 8], weight: 2, sectors: ['Charity', 'Education', 'Professional services', 'Health', 'Property'],
    names: ['Annual report', 'Impact report', 'Annual review'],
    tasks: [['Layout and design', 'Design', 0.48], ['Artwork and proofs', 'Design', 0.3], ['Editorial direction', 'Partner', 0.08], ['Print management', 'Project management', 0.14]],
    stages: [0.5, 0.5],
    costs: [['Printed proofs', 'Printing', 30, 120, true], ['Courier to printer', 'Couriers', 15, 45, true], ['Stock photography licence', 'Licences', 60, 240, true]] },
  { type: 'Packaging design', kind: 'Fixed fee', hours: [60, 130], weeks: [6, 10], weight: 2, sectors: ['Food and drink', 'Retail', 'Health'],
    names: ['Packaging range', 'Label redesign', 'Gift range packaging'],
    tasks: [['Shelf research', 'Partner', 0.1], ['Concepts', 'Design', 0.35], ['Artwork', 'Design', 0.35], ['Printer liaison', 'Project management', 0.2]],
    stages: [0.5, 0.5],
    costs: [['Printed mock-ups', 'Printing', 60, 260, true], ['Courier', 'Couriers', 12, 40, true], ['Font licence', 'Licences', 80, 300, true]] },
  { type: 'Campaign design', kind: 'Fixed fee', hours: [45, 100], weeks: [4, 7], weight: 2,
    names: ['campaign'],
    tasks: [['Campaign idea', 'Partner', 0.12], ['Key visuals', 'Design', 0.38], ['Adaptations and resizes', 'Design', 0.38], ['Project management', 'Project management', 0.12]],
    stages: [0.5, 0.5],
    costs: [['Stock photography licence', 'Licences', 60, 300, true], ['Printed proofs', 'Printing', 25, 90, true]] },
  { type: 'UX review', kind: 'Fixed fee', hours: [30, 60], weeks: [3, 5], weight: 2, sectors: ['Technology', 'Retail', 'Hospitality', 'Education', 'Property'],
    names: ['Website UX review', 'App UX review', 'Checkout review'],
    tasks: [['User interviews', 'Partner', 0.3], ['Review and findings', 'Design', 0.45], ['Report and workshop', 'Partner', 0.25]],
    stages: [0.5, 0.5],
    costs: [['Interview participant vouchers', 'Research', 100, 300, true], ['Train fare to client', 'Travel', 30, 160, false]] },
  { type: 'Strategy consultancy', kind: 'Hourly', hours: [30, 90], weeks: [4, 10], weight: 3,
    names: ['Brand strategy', 'Marketing strategy', 'Customer research', 'Positioning workshops'],
    tasks: [['Workshops', 'Partner', 0.5], ['Research and analysis', 'Partner', 0.3], ['Write-up and reporting', 'Project management', 0.2]],
    costs: [['Train fare to client', 'Travel', 30, 180, true], ['Workshop room hire', 'Venue', 120, 380, true], ['Workshop materials', 'Materials', 15, 60, false]] },
];
/** Retainers: hourly work paid for each month, renewed when they end. */
const RETAINERS = [
  { type: 'Design support', kind: 'Hourly', monthly: [26, 40], months: 6, weight: 0,
    tasks: [['Design requests', 'Design', 0.85], ['Account management', 'Project management', 0.15]],
    costs: [['Stock photography licence', 'Licences', 40, 160, true], ['Courier', 'Couriers', 12, 40, true]] },
  { type: 'Website support', kind: 'Hourly', monthly: [14, 24], months: 12, weight: 0,
    tasks: [['Fixes and updates', 'Development', 0.85], ['Account management', 'Project management', 0.15]],
    costs: [['Plugin licence renewal', 'Licences', 40, 180, true]] },
  { type: 'Advisory retainer', kind: 'Hourly', monthly: [10, 18], months: 6, weight: 0,
    tasks: [['Advice and planning', 'Partner', 0.8], ['Account management', 'Project management', 0.2]],
    costs: [['Train fare to client', 'Travel', 30, 180, true]] },
];
/** Sectors the clients are in: [sector, name endings, how many]. */
const SECTORS = [
  ['Charity', ['Trust', 'Foundation'], 4], ['Food and drink', ['Bakery', 'Brewing', 'Coffee Roasters', 'Kitchen'], 4], ['Technology', ['Software', 'Labs', 'Systems'], 3],
  ['Education', ['College', 'Academy'], 2], ['Hospitality', ['Hotels', 'Inns'], 2], ['Retail', ['Home', 'Interiors', 'Outfitters'], 3],
  ['Professional services', ['Partners', 'Accountants', 'Surveyors'], 3], ['Property', ['Homes', 'Estates'], 2], ['Health', ['Clinic', 'Dental', 'Physio'], 2],
];
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const SEASONS = ['Winter', 'Winter', 'Spring', 'Spring', 'Spring', 'Summer', 'Summer', 'Summer', 'Autumn', 'Autumn', 'Autumn', 'Winter'];
/** How busy each month is for new work: quiet in August and December. */
const DEMAND = [1.1, 1.1, 1.1, 1, 1, 0.9, 0.85, 0.6, 1.2, 1.2, 1.1, 0.7];
const NOTES = {
  Partner: ['Client call', 'Workshop prep', 'Review with the team', 'Presentation to client', ''],
  Design: ['Revisions after feedback', 'First round', 'Internal review', 'Final files', '', ''],
  Development: ['Bug fixes', 'Templates', 'Content load', 'Deploy to staging', '', ''],
  'Project management': ['Status report', 'Client call', 'Schedule update', '', ''],
};
const LOST = ['Went with a studio nearer to them.', 'Budget cut for this year.', 'Put back to next year.', 'Chose to do it in house.'];

/** Fill a Bramble Studio database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 480 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };
  const quarter = (h) => Math.round(h * 4) / 4;
  const half = (h) => Math.round(h * 2) / 2;
  const between = ([lo, hi]) => lo + r.next() * (hi - lo);
  const slug = (name) => name.toLowerCase().replace(/[^a-z]+/g, '');
  // The studio is shut at weekends and from Christmas Eve to New Year's Day.
  const closed = (d) => {
    const w = cal.weekday(d);
    if (w === 0 || w === 6) return true;
    const a = cal.at(d);
    return (a.getMonth() === 11 && a.getDate() >= 24) || (a.getMonth() === 0 && a.getDate() === 1);
  };
  const open = (d) => { let x = d; while (closed(x)) x++; return x; };

  // The studio's own details, which every document it prints carries.
  add('Business', businessRecord({
    name: 'Bramble Studio', domain: 'bramble.example', phone: '020 7946 0410',
    address: { Address: '14 Canal Walk', City: 'Bristol', Postcode: 'BS1 4QR', Country: 'United Kingdom' },
    terms: 'Payment by the due date shown, by bank transfer.',
    footer: 'Thank you for working with us. The rights in the work we deliver pass to you once its invoices are paid in full. Studio open Monday to Friday, 9 to 5.30.',
  }));

  // The team.
  const used = new Set();
  const team = TEAM.map((t, i) => {
    let p;
    do p = person(r, { domain: 'bramble.example' }); while (used.has(p['Last Name']));
    used.add(p['Last Name']);
    const id = add('Staff', {
      ...p, Role: t.role, Phone: `07700 900${String(410 + i * 7)}`, 'Hourly Cost': t.cost, 'Charge-out Rate': t.rate,
      'Hours per Week': t.day * 5, Started: cal.day(-r.int(...t.since)), Active: true,
    });
    return { ...t, id };
  });

  // Clients, one of them the studio itself for its own time.
  const studio = add('Clients', {
    Name: 'Bramble Studio', Sector: 'Internal', Address: '14 Canal Walk', City: 'Bristol', Postcode: 'BS1 4QR', Country: 'United Kingdom',
    Phone: '020 7946 0410', 'E-mail': 'hello@bramble.example', 'Payment Days': null, 'VAT Rate': 0, Notes: 'Our own time: admin, new business and training.',
  });
  const clients = [];
  const names = new Set();
  for (const [sector, tails, n] of SECTORS) {
    for (let k = 0; k < n; k++) {
      let name;
      do name = companies(r, 1, tails)[0]; while (names.has(name));
      names.add(name);
      const where = r.weighted([[0, 8], [3, 1], [1, 1]]);
      const a = address(r, [PLACES[where]]);
      const terms = r.weighted([[30, 6], [14, 2], [45, 1]]);
      const slow = r.chance(0.18);
      const id = add('Clients', {
        Name: name, Sector: sector, ...a, Phone: phoneFor(r, a.Country), 'E-mail': `accounts@${slug(name)}.example`, 'Payment Days': terms,
        'VAT Rate': a.Country === 'United Kingdom' ? 0.2 : 0, Notes: slow ? 'Pays late; chase at 30 days.' : '',
      });
      let main = null;
      for (let c = 0; c < r.int(1, 3); c++) {
        const contact = add('Contacts', {
          Client: id, ...person(r, { domain: `${slug(name)}.example` }), Phone: phoneFor(r, a.Country), 'Main Contact': c === 0,
          'Job Title': c === 0 ? r.pick(['Marketing Manager', 'Head of Communications', 'Founder', 'Managing Director']) : r.pick(['Finance Manager', 'Brand Manager', 'Operations Director', 'Digital Lead']),
        });
        if (c === 0) main = contact;
      }
      clients.push({ id, name, sector, main, terms, slow, vat: a.Country === 'United Kingdom' ? 0.2 : 0, weight: 0.5 + r.next() * 2, first: null });
    }
  }

  // Projects and their tasks. Each project keeps its plan and its progress here.
  const projects = [];
  let code = 100;
  const makeProject = (p) => {
    const id = add('Projects', {
      Code: p.code || `P-${++code}`, Name: p.name, Client: p.client ? p.client.id : studio, Contact: p.client ? p.client.main : null, Kind: p.kind, 'Work Type': p.work ? p.work.type : 'Internal',
      'Budget Hours': p.budget, Fee: p.fee, 'Proposed On': p.proposed == null ? null : cal.day(p.proposed),
      'Start Date': p.start == null ? null : cal.day(p.start), 'Due Date': p.due == null ? null : cal.day(p.due), Status: p.status || 'Proposal',
    });
    p.id = id;
    p.target = p.tasks.reduce((n, t) => n + (t.target || 0), 0);
    let before = 0;
    for (const t of p.tasks) {
      t.id = add('Tasks', { Project: id, Name: t.name, Role: t.role, 'Estimated Hours': t.est });
      // A task can be started once the work before it is well under way; managing it runs throughout.
      t.opensAt = t.role === 'Project management' || p.retainer ? 0 : before * 0.55;
      before += t.share;
    }
    if (p.client && (p.client.first == null || p.proposed < p.client.first)) p.client.first = p.proposed;
    projects.push(p);
    return p;
  };
  const tasksFor = (work, budget, overrun) => work.tasks.map(([name, role, share]) => {
    const est = half(budget * share);
    return { name, role, share, est, target: est * overrun * (0.85 + r.next() * 0.3), spent: 0 };
  });
  const feeFor = (work, budget) => Math.round(work.tasks.reduce((n, [, role, share]) => n + budget * share * RATE[role], 0) * (0.92 + r.next() * 0.12) / 50) * 50;

  // The studio's own projects, for the time that is not a client's.
  const internal = {};
  for (const [key, name, task] of [['business', 'New business and proposals', 'Proposals and pitches'], ['admin', 'Studio admin', 'Admin and planning'], ['training', 'Training and development', 'Courses and reading']]) {
    internal[key] = makeProject({ code: `INT-${Object.keys(internal).length + 1}`, name, kind: 'Internal', status: 'Active', start: -days - 60, tasks: [{ name: task, role: '', share: 1, est: null }] });
    internal[key].started = true;
  }

  /**
   * A proposal for `client` made on day `d`: won (and when it starts), lost,
   * or still waiting. When the people the work needs have `queue` days of work
   * already in hand, fewer are won and those won start later, as a busy
   * studio quotes longer lead times.
   */
  const propose = (client, work, d, { decide = true, queue = 0 } = {}) => {
    const budget = Math.round(between(work.hours) / 5) * 5;
    const overrun = work.kind === 'Hourly' ? 0.8 + r.next() * 0.4 : r.weighted([[0.85 + r.next() * 0.1, 3], [0.95 + r.next() * 0.15, 5], [1.1 + r.next() * 0.25, 2]]);
    const decided = d + r.int(7, 28);
    // A proposal still out could start a little after the answer, and never before today.
    const start = open((decide ? decided : Math.max(decided, 0)) + r.int(4, 14) + Math.round(queue * 0.8));
    const due = open(start + Math.round(between(work.weeks)) * 7);
    const year = cal.at(start).getFullYear();
    const month = cal.at(start).getMonth();
    let name = r.pick(work.names);
    if (work.type === 'Campaign design') name = `${r.pick([SEASONS[month], 'Launch', 'Recruitment', 'Open day'])} campaign`;
    // A report covers the financial year to March: one made from June on is for the year just ended.
    if (work.type === 'Annual report') name = month >= 5 ? `${name} ${year - 1}-${String(year % 100).padStart(2, '0')}` : `${name} ${year - 2}-${String((year - 1) % 100).padStart(2, '0')}`;
    const waiting = decide && decided > -1;
    const won = decide && !waiting && r.chance(queue < 30 ? 0.75 : queue < 55 ? 0.6 : 0.35);
    const p = makeProject({
      client, work, kind: work.kind, name, budget, fee: feeFor(work, budget), proposed: d, start, due, overrun,
      status: waiting || !decide ? 'Proposal' : won ? 'Active' : 'Lost', tasks: tasksFor(work, budget, overrun), won,
    });
    if (!won && !waiting) base.update('Projects', p.id, { Notes: r.pick(LOST) });
    // One project in fifteen stops for a while part way through, the client's doing.
    if (won && work.kind === 'Fixed fee' && r.chance(0.07)) p.hold = { at: 0.3 + r.next() * 0.3, length: r.int(20, 70) };
    return p;
  };
  const pickWork = () => r.weighted(WORK.map((w) => [w, w.weight]));
  // Returning clients are asked first: most work comes from people who have used the studio.
  const pickClient = (work) => r.weighted(clients.filter((c) => !work.sectors || work.sectors.includes(c.sector)).map((c) => [c, c.weight * (c.first != null ? 1.6 : 1)]));

  // Retainers, one of each to begin with, renewed as they end.
  const retainer = (client, kind, start) => {
    const monthly = r.int(...kind.monthly);
    const budget = monthly * kind.months;
    const due = open(start + Math.round(kind.months * 30.4));
    const year = cal.at(start).getFullYear();
    return makeProject({
      client, work: kind, kind: 'Hourly', retainer: true, name: `${kind.type} ${MONTHS[cal.at(start).getMonth()].slice(0, 3)} ${year}`, budget,
      fee: Math.round(kind.tasks.reduce((n, [, role, share]) => n + budget * share * RATE[role], 0) / 50) * 50, proposed: start - r.int(10, 25),
      start, due, status: 'Active', won: true, tasks: kind.tasks.map(([name, role, share]) => ({ name, role, share, est: half(budget * share), target: budget * share * (0.8 + r.next() * 0.2), spent: 0 })),
    });
  };
  const retainerClients = r.some(clients.filter((c) => c.sector !== 'Health'), 3);
  retainerClients[1].slow = true; // the website support client is one who pays late
  retainer(retainerClients[0], RETAINERS[0], open(-days + r.int(0, 4)));
  retainer(retainerClients[1], RETAINERS[1], open(-days + r.int(10, 30)));
  retainer(retainerClients[2], RETAINERS[2], open(-days + r.int(40, 90)));

  // Work already won when the records begin, then a proposal every few days.
  for (let k = 0; k < 6; k++) {
    const work = pickWork();
    const p = propose(pickClient(work), work, -days - r.int(25, 45));
    p.status = 'Active'; p.won = true; p.start = open(-days + r.int(0, 12)); p.due = open(p.start + Math.round(between(p.work.weeks)) * 7);
    base.update('Projects', p.id, { Status: 'Active', 'Start Date': cal.day(p.start), 'Due Date': cal.day(p.due), Notes: null });
  }
  const asked = new Set();
  for (let d = open(-days + 2); d < 0;) {
    asked.add(d);
    d = open(d + Math.max(2, Math.round(r.int(3, 5) / DEMAND[cal.at(d).getMonth()])));
  }
  // Client hours a day each kind of task gets, and the days of won work waiting for the busiest kind `work` needs.
  const capacity = {};
  for (const s of team) capacity[s.kind] = (capacity[s.kind] || 0) + s.day * s.util;
  const queueFor = (work) => Math.max(...work.tasks.filter(([, , share]) => share >= 0.1).map(([, role]) => {
    let hours = 0;
    for (const p of projects) if (p.won && p.finished == null && !p.retainer) for (const t of p.tasks) if (t.role === role) hours += Math.max(0, t.target - t.spent);
    return hours / capacity[role];
  }));

  const raised = []; // invoices made, with the payments planned on each; those up to yesterday are written at the end
  // Invoices numbered in date order, a new run each year: INV-2026-0042.
  const invoiceCount = {};
  const invoiceNumber = (d) => { const y = cal.at(d).getFullYear(); invoiceCount[y] = (invoiceCount[y] || 0) + 1; return `INV-${y}-${String(invoiceCount[y]).padStart(4, '0')}`; };

  /** Raise an invoice for project `p` on day `d` with `lines`, and plan when the client pays it. */
  const invoice = (p, d, lines) => {
    if (!lines.length) return null;
    // Each line's VAT rounded to pence and then added up, as the printed invoice works it out.
    const net = money(lines.reduce((n, l) => n + money(l.qty * l.price), 0));
    const vat = money(lines.reduce((n, l) => n + money(money(l.qty * l.price) * p.client.vat), 0));
    const total = money(net + vat);
    const terms = p.client.terms;
    const number = invoiceNumber(d);
    const id = add('Invoices', { Number: number, Project: p.id, Client: p.client.id, 'Invoice Date': cal.day(d), 'Due Date': cal.day(d + terms), Net: net, VAT: vat, Total: total, Paid: 0, Status: 'Sent' });
    for (const l of lines) add('Invoice Lines', { Invoice: id, Description: l.text, Kind: l.kind, Quantity: l.qty, 'Unit Price': l.price, 'Tax Rate': p.client.vat });
    // Most pay inside their terms; a slow payer some weeks after; now and then in two parts.
    const first = d + (p.client.slow && r.chance(0.8) ? terms + r.int(6, 45) : r.int(Math.max(6, terms - 16), terms + 2));
    const parts = total > 1500 && r.chance(0.08) ? [money(total / 2), money(total - money(total / 2))] : [total];
    const inv = { id, number, total, due: d + terms, payments: [] };
    let at = first;
    for (const amount of parts) {
      const day = open(at);
      inv.payments.push({ day, amount, method: r.chance(0.9) ? 'Bank transfer' : 'Card' });
      at = day + r.int(14, 35);
    }
    raised.push(inv);
    return id;
  };
  /** The project's rechargeable costs from before day `upTo` not yet invoiced. */
  const costsToBill = (p, upTo) => (p.costs || []).filter((e) => e.recharge && !e.invoice && e.d < upTo);
  /** An invoice of `lines` and `costs`, and the hours (`items`) and costs it covers marked with it. */
  const bill = (p, d, lines, items, costs) => {
    const all = [...lines, ...costs.map((e) => ({ text: `Expense: ${e.text}`, kind: 'Expense', qty: 1, price: e.amount }))];
    const id = invoice(p, d, all);
    if (!id) return;
    for (const t of items) { t.invoice = id; base.update('Time Entries', t.id, { Invoice: id }); }
    for (const e of costs) { e.invoice = id; base.update('Expenses', e.id, { Invoice: id }); }
  };
  /** A stage payment on a fixed-fee project, with its costs, and the hours it covers marked as billed. */
  const stage = (p, d) => {
    const i = p.stagesDone || 0;
    const n = p.work.stages.length;
    const amount = i === n - 1 ? money(p.fee - p.billed) : money(p.fee * p.work.stages[i]);
    p.billed = money((p.billed || 0) + amount);
    p.stagesDone = i + 1;
    const words = i === 0 ? 'on starting' : i === n - 1 ? 'on completion' : 'at the half-way review';
    const items = (p.log || []).filter((t) => t.billable && !t.invoice);
    bill(p, d, [{ text: `${p.name}: stage ${i + 1} of ${n}, ${words}`, kind: 'Fee', qty: 1, price: amount }], items, costsToBill(p, d + 1));
  };
  const spend = (p, d) => {
    const c = r.pick(p.work.costs);
    const amount = money(between([c[2], c[3]]));
    const recharge = c[4] && r.chance(0.9);
    const who = r.pick(team.filter((s) => s.kind !== 'Development' || p.work.type.startsWith('Website')));
    const e = { p, d, text: c[0], amount, recharge };
    e.id = add('Expenses', { Date: cal.day(d), Project: p.id, Staff: who.id, Category: c[1], Description: c[0], Amount: amount, Rechargeable: recharge });
    (p.costs ||= []).push(e);
  };
  // A time sheet for each person each week they log time, numbered in the order they are opened: TS-2026-0042.
  const sheets = new Map();
  const sheetCount = {};
  const sheetFor = (s, d) => {
    const monday = d - ((cal.weekday(d) + 6) % 7);
    const key = `${s.id}|${monday}`;
    if (!sheets.has(key)) {
      const y = cal.at(monday).getFullYear();
      sheetCount[y] = (sheetCount[y] || 0) + 1;
      const id = add('Time Sheets', { Number: `TS-${y}-${String(sheetCount[y]).padStart(4, '0')}`, Staff: s.id, 'Week Commencing': cal.day(monday), Status: 'Open' });
      sheets.set(key, { id, monday });
    }
    return sheets.get(key).id;
  };
  const log = (s, p, t, d, hours, billable, note) => {
    const id = add('Time Entries', { Date: cal.day(d), Staff: s.id, Project: p.id, Task: t.id, Hours: hours, Billable: billable, Note: note || null, 'Time Sheet': sheetFor(s, d) });
    (p.log ||= []).push({ id, s, d, hours, billable, invoice: null });
    t.spent += hours;
    p.spent = (p.spent || 0) + hours;
    p.today = (p.today || 0) + hours;
  };
  const finish = (p, d) => {
    p.finished = d;
    p.status = 'Completed';
    if (p.work.type === 'Annual report') {
      const e = { p, d, text: 'Print run, 500 copies', amount: money(between([1200, 3600])), recharge: true };
      e.id = add('Expenses', { Date: cal.day(d), Project: p.id, Staff: team[5].id, Category: 'Printing', Description: e.text, Amount: e.amount, Rechargeable: true });
      (p.costs ||= []).push(e);
    }
    if (p.kind === 'Fixed fee') while (p.stagesDone < p.work.stages.length) stage(p, d);
    // A retainer that ends is mostly renewed for another term.
    if (p.retainer && (p.work === RETAINERS[0] || r.chance(0.85))) retainer(p.client, p.work, open(p.due + 1));
  };

  // The days, one at a time.
  let month = -1;
  let dayOfMonth = 0;
  for (let d = -days; d <= -1; d++) {
    if (closed(d)) continue;
    const at = cal.at(d);
    if (at.getMonth() !== month) { month = at.getMonth(); dayOfMonth = 0; }
    dayOfMonth++;

    // On the third working day of a month, hourly work is invoiced for the months before it.
    if (dayOfMonth === 3) {
      const monthStart = d - (at.getDate() - 1);
      const last = cal.at(monthStart - 1);
      const period = `${MONTHS[last.getMonth()]} ${last.getFullYear()}`;
      for (const p of projects.filter((x) => x.kind === 'Hourly' && x.started)) {
        const items = (p.log || []).filter((t) => t.billable && !t.invoice && t.d < monthStart);
        // One line for each grade of person: the two partners' hours go on one line.
        const lines = [];
        for (const s of team) {
          const hours = items.filter((t) => t.s.grade === s.grade).reduce((n, t) => n + t.hours, 0);
          if (hours > 0 && !lines.some((l) => l.grade === s.grade)) lines.push({ grade: s.grade, text: `${s.grade} time, ${period}`, kind: 'Time', qty: hours, price: s.rate });
        }
        bill(p, d, lines, items, costsToBill(p, monthStart));
      }
    }

    // Work that starts today; fixed fees take their first stage then.
    for (const p of projects) {
      if (!p.won || p.started || p.start > d) continue;
      p.started = true;
      p.status = 'Active';
      if (p.kind === 'Fixed fee') stage(p, d);
    }
    // Projects on hold, and coming back: the due date moves with the pause.
    for (const p of projects) {
      if (!p.started || p.finished != null || !p.hold) continue;
      if (p.hold.from == null && p.spent >= p.hold.at * p.budget) { p.hold.from = d; p.hold.to = d + p.hold.length; p.status = 'On hold'; }
      else if (p.hold.from != null && p.hold.to === d) { p.status = 'Active'; p.due = open(p.due + p.hold.length); }
      else if (p.hold.from != null && p.hold.to < d && p.status === 'On hold') { p.status = 'Active'; p.due = open(p.due + p.hold.length); }
    }
    if (asked.has(d)) { const work = pickWork(); propose(pickClient(work), work, d, { queue: queueFor(work) }); }
    const live = projects.filter((p) => p.started && p.finished == null && p.status === 'Active' && p.kind !== 'Internal');
    // However quiet the studio, a project moves at the pace its client can answer: so many hours a day at most.
    for (const p of live) { p.today = 0; p.cap ??= Math.max(4, (2.5 * p.budget * 7) / (5 * Math.max(14, p.due - p.start))); }

    // August and school holidays empty the studio a little.
    const away = at.getMonth() === 7 ? 0.2 : at.getMonth() === 3 || at.getMonth() === 9 ? 0.09 : 0.05;
    for (const s of team) {
      if (r.chance(s.kind === 'Partner' ? away * 0.7 : away)) continue;
      const left = quarter(s.day * s.util * (0.85 + r.next() * 0.3));
      // The tasks this person can do, most behind first.
      const can = [];
      for (const p of live) {
        for (const t of p.tasks) {
          if (t.role !== s.kind || t.target - t.spent < 0.25 || p.spent < t.opensAt * p.target || p.cap - p.today < 0.5) continue;
          if (p.retainer && t.spent >= t.target * Math.min(1, (d - p.start + 3) / Math.max(1, p.due - p.start))) continue;
          can.push({ p, t, pace: p.due < d ? (t.target - t.spent) * (1 + (d - p.due) / 10) : (t.target - t.spent) / Math.max(1, ((p.due - d) * 5) / 7) });
        }
      }
      can.sort((a, b) => b.pace - a.pace);
      const chosen = can.slice(0, r.int(1, 3));
      const pace = chosen.reduce((n, c) => n + c.pace, 0);
      let worked = 0;
      for (const c of chosen) {
        const share = chosen.length === 1 ? left : left * Math.max(0.25, c.pace / pace);
        const hours = quarter(Math.min(share, c.t.target - c.t.spent + 0.25, left - worked, c.p.cap - c.p.today));
        if (hours < 0.5) continue;
        // Now and then a correction is done at the studio's cost.
        const ours = r.chance(0.04);
        log(s, c.p, c.t, d, hours, !ours, ours ? 'Correction at our cost' : r.pick(NOTES[s.kind]));
        worked += hours;
      }
      // The rest of the day is the studio's own: new business for the partners, admin, a course now and then.
      const rest = quarter(s.day - worked - r.next() * 0.75);
      if (rest >= 0.5) {
        const kind = s.kind === 'Partner' ? r.weighted([['business', 6], ['admin', 4]]) : r.weighted([['admin', 6], ['business', 2], ['training', 1]]);
        log(s, internal[kind], internal[kind].tasks[0], d, rest, false, kind === 'business' ? r.pick(['Proposal writing', 'Pitch', 'Meeting a prospect', '']) : '');
      }
    }

    // Costs run up on the projects in hand.
    for (const p of live) if (r.chance(p.retainer ? 0.03 : 0.06)) spend(p, d);

    // Work finished today, the stage at half-way, and retainers at the end of their term.
    for (const p of live) {
      if (p.kind === 'Fixed fee' && p.work.stages.length === 3 && p.stagesDone === 1 && p.spent >= p.budget * 0.5) stage(p, d);
      if (p.retainer ? d >= p.due : p.tasks.every((t) => t.target - t.spent < 0.25)) finish(p, d);
    }
  }

  // Always a few proposals out waiting for an answer.
  while (projects.filter((p) => p.status === 'Proposal').length < 3) { const work = pickWork(); propose(pickClient(work), work, open(-r.int(3, 14)), { decide: false, queue: queueFor(work) }); }
  // Clients were clients from before their first proposal.
  for (const c of clients) base.update('Clients', c.id, { 'Client Since': cal.day((c.first ?? -r.int(30, 200)) - r.int(0, 900)) });

  // Something always falls due within the month: a deadline brought forward if need be.
  const active = projects.filter((p) => p.status === 'Active' && p.kind !== 'Internal' && p.due != null);
  if (!active.some((p) => p.due >= 0 && p.due <= 30) && active.length) {
    const p = active.filter((x) => x.started).sort((a, b) => a.due - b.due).find((x) => x.due > 30) || active[0];
    p.due = cal.workday(r.int(6, 26));
  }

  // What has been paid by yesterday. If no invoice is overdue, the client of a recent one has gone quiet.
  const late = (inv) => inv.due < -1 && inv.payments.some((x) => x.day > -1);
  if (!raised.some(late)) {
    const quiet = raised.filter((inv) => inv.due >= -45 && inv.due <= -3).pop();
    if (quiet) for (const x of quiet.payments) x.day = 0;
  }
  const received = raised.flatMap((inv) => inv.payments.filter((x) => x.day <= -1).map((x) => ({ inv, ...x }))).sort((a, b) => a.day - b.day);
  for (const x of received) {
    add('Payments', { Invoice: x.inv.id, 'Received On': cal.day(x.day), Amount: x.amount, Method: x.method, Reference: x.inv.number });
    x.inv.paid = money((x.inv.paid || 0) + x.amount);
  }
  for (const inv of raised) {
    const paid = inv.paid || 0;
    base.update('Invoices', inv.id, { Paid: paid, Status: paid >= inv.total ? 'Paid' : paid > 0 ? 'Part paid' : 'Sent' });
  }

  // Where each project stands today. An hourly project's fee becomes what its time came to once it is done.
  for (const p of projects) {
    if (p.kind === 'Internal') continue;
    const values = { Status: p.status, 'Due Date': cal.day(p.due), 'Finished On': p.finished == null ? null : cal.day(p.finished) };
    if (p.status === 'On hold') values.Notes = 'Paused by the client; restart date to be agreed.';
    if (p.kind === 'Hourly' && p.finished != null) values.Fee = money((p.log || []).filter((t) => t.billable).reduce((n, t) => n + t.hours * t.s.rate, 0));
    base.update('Projects', p.id, values);
    for (const t of p.tasks) if (p.finished != null || (p.started && t.target - t.spent < 0.25)) base.update('Tasks', t.id, { Done: true });
  }

  // Proposals numbered in the order they went out, a run each year: QUO-2026-0042.
  const quoteCount = {};
  for (const p of projects.filter((x) => x.kind !== 'Internal').sort((a, b) => a.proposed - b.proposed || a.id - b.id)) {
    const y = cal.at(p.proposed).getFullYear();
    quoteCount[y] = (quoteCount[y] || 0) + 1;
    const values = { 'Quote No': `QUO-${y}-${String(quoteCount[y]).padStart(4, '0')}` };
    // Charities, colleges and clinics pay only against their own purchase order, quoted on every invoice.
    if (['Charity', 'Education', 'Health'].includes(p.client.sector) && p.won) values['Client PO'] = `PO-${r.int(10000, 89999)}`;
    base.update('Projects', p.id, values);
  }

  // Each week's sheet is sent in on the Friday or the Monday after, and approved a day or two later.
  for (const sh of sheets.values()) {
    const sent = open(sh.monday + (r.chance(0.7) ? 4 : 7));
    if (sent > -1) continue;
    const approved = open(sent + r.int(1, 2));
    base.update('Time Sheets', sh.id, approved > -1 ? { Status: 'Submitted', 'Submitted On': cal.day(sent) } : { Status: 'Approved', 'Submitted On': cal.day(sent), 'Approved On': cal.day(approved) });
  }
  return counts;
}
