// The clinic: Cedar Lane Clinic, an invented small private clinic with two
// doctors, a practice nurse and two physiotherapists, and a reception desk
// that books for them. Its patients and the insurers some of them claim
// from, the practitioners and the hours they work, the services it offers
// with their length and price, every appointment and what came of it, a
// short note for each visit, the invoices and the payments against them, and
// the recalls that bring patients back for a check or a review.
//
// Its records are a year of the clinic run a day at a time. Patients
// ring or book online, mostly a few days ahead; the doctors are busiest in
// winter, the nurse in the autumn flu weeks, the physiotherapists in spring.
// A physiotherapy assessment starts a course of weekly sessions, a doctor
// sometimes books a follow-up or a blood test, and a health check sets a
// recall a year on. Some patients cancel, a few at the last minute, and a
// few do not come; the clinic charges a fee for some of those. Insurers pay
// in their own time, one of them slowly. The notes are invented and kept to
// plain, everyday words.
//
// What it prints: an invoice for each visit (to the patient, or to the
// insurer with the policy reference), a receipt for each payment, a letter
// confirming each appointment, a recall letter, and a statement of claims
// for each insurer. Medical care is exempt from VAT, so the Business record
// carries a rate of nought and the invoices say so.

import { BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, money, person } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const PATIENT = '[Patients].[First Name] & " " & [Patients].[Last Name]';
const DAY = 'Format([Start], "yyyy-mm-dd")';
const AHEAD = `DateDiff("d", Date(), ${DAY})`;
const AGO = `DateDiff("d", ${DAY}, Date())`;
const OWED_BY = `IIf([Bill To] = "Insurer", [Insurers].[Name], ${PATIENT})`;
const WEEKDAY = `DateDiff("d", #2024-01-01#, ${DAY}) Mod 7`;
const LETTERHEAD = { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Website'] };

export const spec = {
  // Where a value below nought is right: an invoice not due yet is "overdue" by minus some days.
  negative: { 'Unpaid invoices': ['Days Overdue'] },
  tables: [
    BUSINESS_TABLE,
    { name: 'Insurers', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Kind', type: 'text', default: 'Insurer' },
      { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Payment Days', type: 'integer', default: 30 },
      { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Patients', fields: [
      { name: 'Patient No', type: 'text', unique: true }, { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
      { name: 'Date of Birth', type: 'date' }, { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Insurer', ...ref('Insurers'), required: true },
      { name: 'Policy Reference', type: 'text' }, { name: 'Registered', type: 'date' }, { name: 'Reminders', type: 'yesno', default: true },
      { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Practitioners', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Known As', type: 'text', required: true },
      { name: 'Role', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Phone', type: 'text' }, { name: 'Works', type: 'text' },
      { name: 'Hours per Week', type: 'number' }, { name: 'Started', type: 'date' },
    ] },
    { name: 'Services', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Kind', type: 'text' }, { name: 'Given By', type: 'text' },
      { name: 'Minutes', type: 'integer', default: 15 }, { name: 'Price', type: 'currency' }, { name: 'Insurer Pays', type: 'yesno' },
      { name: 'Recall Months', type: 'integer' },
    ] },
    { name: 'Appointments', fields: [
      { name: 'Patient', ...ref('Patients'), required: true }, { name: 'Practitioner', ...ref('Practitioners'), required: true },
      { name: 'Service', ...ref('Services'), required: true }, { name: 'Start', type: 'datetime', required: true },
      { name: 'Status', type: 'text', default: 'Booked' }, { name: 'Booked On', type: 'date' }, { name: 'Booked Via', type: 'text' },
      { name: 'Cancelled On', type: 'date' },
    ] },
    { name: 'Visit Notes', fields: [
      { name: 'Appointment', ...ref('Appointments'), required: true }, { name: 'Written', type: 'datetime' },
      { name: 'Reason', type: 'text' }, { name: 'Note', type: 'longtext' }, { name: 'Plan', type: 'text' },
    ] },
    // An invoice keeps the name and address it was sent to, and points at the insurer when one is
    // billed, so each insurer's statement lists its own claims.
    { name: 'Invoices', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Appointment', ...ref('Appointments'), required: true },
      { name: 'Invoice Date', type: 'date' }, { name: 'Due Date', type: 'date' },
      { name: 'Bill To', type: 'text', default: 'Patient' }, { name: 'Insurer', ...ref('Insurers') }, { name: 'Bill To Name', type: 'text' },
      { name: 'Bill To Address', type: 'text' }, { name: 'Bill To City', type: 'text' }, { name: 'Bill To Postcode', type: 'text' },
      { name: 'Policy Reference', type: 'text' }, { name: 'For', type: 'text', default: 'Visit' }, { name: 'Amount', type: 'currency' },
      { name: 'Amount Paid', type: 'currency', default: 0 }, { name: 'Paid On', type: 'date' },
    ] },
    { name: 'Payments', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Invoice', ...ref('Invoices'), required: true }, { name: 'Paid On', type: 'date' }, { name: 'Amount', type: 'currency' },
      { name: 'Method', type: 'text' }, { name: 'Reference', type: 'text' },
    ] },
    { name: 'Recalls', fields: [
      { name: 'Patient', ...ref('Patients'), required: true }, { name: 'Service', ...ref('Services'), required: true },
      { name: 'Due Date', type: 'date' }, { name: 'Reason', type: 'text' }, { name: 'Sent On', type: 'date' }, { name: 'Booked On', type: 'date' },
    ] },
  ],
  queries: {
    'Next 7 days': { tables: ['Appointments', 'Practitioners', 'Services'], columns: [
      { expr: DAY, as: 'Day', total: 'group', sort: 'asc' }, { table: 'Practitioners', field: 'Known As', as: 'Practitioner', total: 'group', sort: 'asc' },
      { table: 'Appointments', field: 'ID', as: 'Appointments', total: 'count' }, { expr: '[Minutes] / 60', as: 'Booked Hours', total: 'sum' },
      { expr: 'Mid([Start], 12, 5)', as: 'First Start', total: 'min' }, { expr: 'Mid([Start], 12, 5)', as: 'Last Start', total: 'max' },
      { table: 'Appointments', field: 'Status', total: 'where', criteria: ['Booked'] },
      { expr: AHEAD, as: 'Days Ahead', total: 'where', criteria: ['Between 0 And 6'] },
    ] },
    'Day sheet': { tables: ['Appointments', 'Patients', 'Practitioners', 'Services', 'Insurers'], columns: [
      { expr: `${DAY} & " " & [Known As]`, as: 'Sheet', sort: 'asc' }, { expr: 'Mid([Start], 12, 5)', as: 'Time', sort: 'asc' },
      { expr: PATIENT, as: 'Patient' }, { table: 'Patients', field: 'Patient No' }, { table: 'Patients', field: 'Phone' },
      { table: 'Services', field: 'Name', as: 'Service' }, { table: 'Services', field: 'Minutes' }, { table: 'Insurers', field: 'Name', as: 'Paid By' },
      { table: 'Patients', field: 'Notes', as: 'Patient Notes' },
      { table: 'Appointments', field: 'Status', criteria: ['Booked'], show: false }, { expr: AHEAD, as: 'Days Ahead', criteria: ['Between 0 And 6'], show: false },
    ] },
    'Diary next 14 days': { tables: ['Appointments', 'Practitioners', 'Services'], columns: [
      { table: 'Practitioners', field: 'Known As', as: 'Practitioner', total: 'group', sort: 'asc' }, { table: 'Practitioners', field: 'Role', total: 'group' },
      { table: 'Appointments', field: 'ID', as: 'Appointments', total: 'count' }, { expr: '[Minutes] / 60', as: 'Booked Hours', total: 'sum' },
      { expr: '[Hours per Week] * 2', as: 'Open Hours', total: 'max' },
      { expr: 'Round([Minutes] * 100 / ([Hours per Week] * 120), 2)', as: 'Percent Full', total: 'sum' },
      { table: 'Appointments', field: 'Status', total: 'where', criteria: ['Booked'] },
      { expr: AHEAD, as: 'Days Ahead', total: 'where', criteria: ['Between 0 And 13'] },
    ] },
    'No-shows and late cancellations': { tables: ['Appointments', 'Patients', 'Practitioners', 'Services'], columns: [
      { table: 'Appointments', field: 'Start', sort: 'desc' }, { expr: PATIENT, as: 'Patient' }, { table: 'Patients', field: 'Phone' },
      { table: 'Practitioners', field: 'Known As', as: 'Practitioner' }, { table: 'Services', field: 'Name', as: 'Service' },
      { table: 'Appointments', field: 'Status', criteria: ['No-show', 'Cancelled'] }, { table: 'Appointments', field: 'Cancelled On' },
      // A cancellation the day before or on the day is a late one; a no-show has no notice at all.
      { expr: `DateDiff("d", [Cancelled On], ${DAY})`, as: 'Days Notice', criteria: ['', '<= 1'] },
      { expr: AGO, as: 'Days Ago', criteria: ['Between 0 And 90', 'Between 0 And 90'], show: false },
    ] },
    'Unpaid invoices': { tables: ['Invoices', 'Appointments', 'Patients', 'Insurers'], columns: [
      { table: 'Invoices', field: 'Number', as: 'Invoice' }, { expr: OWED_BY, as: 'Owed By', sort: 'asc' }, { table: 'Invoices', field: 'Bill To' },
      { expr: PATIENT, as: 'Patient' }, { table: 'Invoices', field: 'For' }, { table: 'Invoices', field: 'Invoice Date' },
      { table: 'Invoices', field: 'Due Date', sort: 'asc' }, { table: 'Invoices', field: 'Amount' }, { table: 'Invoices', field: 'Amount Paid', as: 'Paid So Far' },
      { expr: 'Round([Amount] - Nz([Amount Paid]), 2)', as: 'Balance' }, { expr: 'DateDiff("d", [Due Date], Date())', as: 'Days Overdue' },
      { table: 'Invoices', field: 'Paid On', criteria: ['Is Null'], show: false },
    ] },
    'Money owed': { tables: ['Invoices', 'Appointments', 'Patients', 'Insurers'], columns: [
      { expr: OWED_BY, as: 'Owed By', total: 'group' }, { table: 'Invoices', field: 'Bill To', total: 'group' },
      { table: 'Invoices', field: 'ID', as: 'Invoices', total: 'count' }, { expr: 'Round([Amount] - Nz([Amount Paid]), 2)', as: 'Balance', total: 'sum', sort: 'desc' },
      { table: 'Invoices', field: 'Due Date', as: 'Oldest Due', total: 'min' }, { table: 'Invoices', field: 'Paid On', total: 'where', criteria: ['Is Null'] },
    ] },
    'Takings by month': { tables: ['Payments'], columns: [
      { expr: 'Format([Paid On], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { expr: 'Round(IIf([Method] = "Insurer", 0, [Amount]), 2)', as: 'From Patients', total: 'sum' },
      { expr: 'Round(IIf([Method] = "Insurer", [Amount], 0), 2)', as: 'From Insurers', total: 'sum' },
      { table: 'Payments', field: 'Amount', as: 'Takings', total: 'sum' },
    ] },
    'Takings by practitioner': { tables: ['Payments', 'Invoices', 'Appointments', 'Practitioners'], columns: [
      { table: 'Practitioners', field: 'Known As', as: 'Practitioner', total: 'group' }, { table: 'Practitioners', field: 'Role', total: 'group' },
      { table: 'Payments', field: 'ID', as: 'Payments', total: 'count' }, { table: 'Payments', field: 'Amount', as: 'Takings', total: 'sum', sort: 'desc' },
      { expr: 'DateDiff("d", [Payments].[Paid On], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Fees by service': { tables: ['Invoices', 'Appointments', 'Services'], columns: [
      { table: 'Services', field: 'Name', as: 'Service', total: 'group' }, { table: 'Services', field: 'Kind', total: 'group' },
      { table: 'Invoices', field: 'ID', as: 'Visits', total: 'count' }, { table: 'Invoices', field: 'Amount', as: 'Invoiced', total: 'sum', sort: 'desc' },
      { table: 'Invoices', field: 'For', total: 'where', criteria: ['Visit'] },
      { expr: 'DateDiff("d", [Invoice Date], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Busy days': { tables: ['Appointments', 'Services'], columns: [
      // Counted in days from Monday 1 January 2024: the remainder after sevens is the day of the week.
      { expr: `${WEEKDAY} + 1`, as: 'Day No', total: 'group', sort: 'asc' },
      { expr: `Mid("MonTueWedThuFriSatSun", (${WEEKDAY}) * 3 + 1, 3)`, as: 'Day', total: 'group' },
      { table: 'Appointments', field: 'ID', as: 'Appointments', total: 'count' }, { expr: '[Minutes] / 60', as: 'Hours Booked', total: 'sum' },
      { table: 'Appointments', field: 'Status', total: 'where', criteria: ['In ("Attended", "No-show")'] },
      { expr: AGO, as: 'Days Ago', total: 'where', criteria: ['Between 1 And 365'] },
    ] },
    'Busy hours': { tables: ['Appointments', 'Services'], columns: [
      { expr: 'Mid([Start], 12, 2) & ":00"', as: 'Hour', total: 'group', sort: 'asc' },
      { table: 'Appointments', field: 'ID', as: 'Appointments', total: 'count' }, { expr: '[Minutes] / 60', as: 'Hours Booked', total: 'sum' },
      { table: 'Appointments', field: 'Status', total: 'where', criteria: ['In ("Attended", "No-show")'] },
      { expr: AGO, as: 'Days Ago', total: 'where', criteria: ['Between 1 And 365'] },
    ] },
    'Recalls due': { tables: ['Recalls', 'Patients', 'Services'], columns: [
      { table: 'Recalls', field: 'Due Date', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Due Date])', as: 'Days Left', criteria: ['Between 0 And 30'] },
      { expr: PATIENT, as: 'Patient' }, { table: 'Patients', field: 'Phone' }, { table: 'Patients', field: 'E-mail' },
      { table: 'Services', field: 'Name', as: 'Recall For' }, { table: 'Recalls', field: 'Reason' },
      { table: 'Patients', field: 'Reminders', as: 'Letters Allowed' }, { table: 'Recalls', field: 'Sent On' },
      { table: 'Recalls', field: 'Booked On', criteria: ['Is Null'], show: false },
    ] },
    'Recall letters to send': { tables: ['Recalls', 'Patients', 'Services'], columns: [
      { table: 'Services', field: 'Name', as: 'Recall For', sort: 'asc' }, { expr: PATIENT, as: 'Patient' },
      { table: 'Patients', field: 'Address' }, { table: 'Patients', field: 'City' }, { table: 'Patients', field: 'Postcode' },
      { table: 'Recalls', field: 'Due Date', sort: 'asc' }, { expr: 'DateDiff("d", Date(), [Due Date])', as: 'Days Left', criteria: ['Between 0 And 30'] },
      { table: 'Patients', field: 'Reminders', criteria: ['Yes'], show: false }, { table: 'Recalls', field: 'Sent On', criteria: ['Is Null'], show: false },
      { table: 'Recalls', field: 'Booked On', criteria: ['Is Null'], show: false },
    ] },
    'New patients by month': { tables: ['Patients', 'Insurers'], columns: [
      { expr: 'Format([Registered], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' },
      { table: 'Patients', field: 'ID', as: 'New Patients', total: 'count' }, { expr: 'IIf([Kind] = "Insurer", 1, 0)', as: 'Insured', total: 'sum' },
      { expr: 'DateDiff("d", [Registered], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Not seen for six months': { tables: ['Patients', 'Appointments'], columns: [
      { expr: PATIENT, as: 'Patient', total: 'group' }, { table: 'Patients', field: 'Patient No', total: 'group' }, { table: 'Patients', field: 'Phone', total: 'group' },
      { expr: DAY, as: 'Last Visit', total: 'max', sort: 'asc' }, { expr: AGO, as: 'Days Since', total: 'min', criteria: ['> 180'] },
      { table: 'Appointments', field: 'Status', total: 'where', criteria: ['Attended'] },
    ] },
    'Patient list': { tables: ['Patients', 'Insurers'], columns: [
      { table: 'Patients', field: 'Last Name', sort: 'asc' }, { table: 'Patients', field: 'First Name', sort: 'asc' }, { table: 'Patients', field: 'Patient No' },
      { table: 'Patients', field: 'Date of Birth' }, { table: 'Patients', field: 'Phone' }, { table: 'Patients', field: 'E-mail' },
      { table: 'Insurers', field: 'Name', as: 'Insurer' }, { table: 'Patients', field: 'Policy Reference' }, { table: 'Patients', field: 'Reminders' },
      { table: 'Patients', field: 'Registered' },
    ] },
    'Visit history': { tables: ['Visit Notes', 'Appointments', 'Patients', 'Practitioners', 'Services'], columns: [
      { table: 'Patients', field: 'Last Name', sort: 'asc', show: false }, { table: 'Patients', field: 'First Name', sort: 'asc', show: false },
      { expr: PATIENT, as: 'Patient' }, { table: 'Patients', field: 'Patient No' }, { expr: DAY, as: 'Seen', sort: 'desc' },
      { table: 'Practitioners', field: 'Known As', as: 'Practitioner' }, { table: 'Services', field: 'Name', as: 'Service' },
      { table: 'Visit Notes', field: 'Reason' }, { table: 'Visit Notes', field: 'Plan' },
    ] },
  },
  // The print templates: one document for each record of the table named.
  documents: [
    { name: 'Invoice', table: 'Invoices', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Invoice {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Bill To Name', label: 'Invoice to' }, { field: 'Number', label: 'Invoice number' },
        { field: 'Bill To Address', label: 'Address' }, { field: 'Invoice Date', label: 'Date' },
        { field: 'Bill To City', label: 'Town' }, { field: 'Due Date', label: 'Payment due' },
        { field: 'Bill To Postcode', label: 'Postcode' }, { field: 'For', label: 'Charge for' },
      ] },
      { type: 'fields', columns: 2, items: [
        { field: 'Appointment.Patient', label: 'Patient' }, { field: 'Appointment.Service.Name', label: 'Service' },
        { field: 'Appointment.Patient.Patient No', label: 'Patient number' }, { field: 'Appointment.Start', label: 'Appointment' },
        { field: 'Appointment.Patient.Date of Birth', label: 'Date of birth' }, { field: 'Appointment.Practitioner.Known As', label: 'Seen by' },
        { field: 'Policy Reference', label: 'Policy reference' }, { field: 'Appointment.Service.Minutes', label: 'Minutes' },
      ] },
      { type: 'text', text: 'Payments received against this invoice:', small: true },
      { type: 'lines', table: 'Payments', link: 'Invoice', sort: 'Paid On', columns: [
        { field: 'Number', label: 'Receipt' }, { field: 'Paid On', label: 'Paid on' }, { field: 'Method' }, { field: 'Reference' },
        { field: 'Amount', as: 'Received', type: 'currency' },
      ] },
      { type: 'totals', rows: [
        { label: 'Fee', expr: '[Amount]' }, { label: 'VAT', expr: 'Round([Amount] * Nz([Business].[Tax Rate]), 2)' },
        { label: 'Total', expr: '[Fee] + [VAT]' }, { label: 'Paid', expr: '[Received]' },
        { label: 'Balance due', expr: 'Round([Total] - [Received], 2)', strong: true },
      ] },
      { type: 'text', text: 'Medical care given at the clinic is exempt from VAT, so no VAT is charged. {Business.Payment Terms} Please pay by card at the desk, or by transfer to {Business.Bank Details}, quoting {Number}.', small: true },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    { name: 'Payment receipt', table: 'Payments', paper: 'a5', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Receipt {Number}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Invoice.Bill To Name', label: 'Received from' }, { field: 'Paid On', label: 'Date' },
        { field: 'Invoice', label: 'For invoice' }, { field: 'Method', label: 'Paid by' },
        { field: 'Invoice.Appointment.Patient', label: 'Patient' }, { field: 'Reference', label: 'Reference' },
        { field: 'Invoice.Appointment.Service.Name', label: 'Service' }, { field: 'Invoice.Appointment.Start', label: 'Appointment' },
      ] },
      { type: 'totals', rows: [{ label: 'Amount received', expr: '[Amount]', strong: true }] },
      { type: 'fields', columns: 2, items: [{ field: 'Invoice.Amount', label: 'Invoice total' }, { field: 'Invoice.Amount Paid', label: 'Paid to date' }] },
      { type: 'text', text: 'Thank you. Medical care is exempt from VAT, so this receipt shows none.', small: true },
      { type: 'text', text: '{Business.Document Footer}', small: true },
    ] },
    { name: 'Appointment letter', table: 'Appointments', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Your appointment' },
      { type: 'fields', columns: 2, items: [
        { field: 'Patient', label: 'To' }, { field: 'Patient.Patient No', label: 'Patient number' },
        { field: 'Patient.Address', label: 'Address' }, { field: 'Booked On', label: 'Booked on' },
        { field: 'Patient.City', label: 'Town' }, { field: 'Patient.Postcode', label: 'Postcode' },
      ] },
      { type: 'text', text: 'Dear {Patient.First Name},' },
      { type: 'text', text: 'Thank you for booking with us. This letter confirms your appointment.' },
      { type: 'fields', columns: 2, items: [
        { field: 'Start', label: 'Date and time' }, { field: 'Service.Name', label: 'Appointment for' },
        { field: 'Practitioner.Known As', label: 'With' }, { field: 'Service.Minutes', label: 'Length in minutes' },
        { field: 'Patient.Insurer', label: 'Paid by' }, { field: 'Patient.Policy Reference', label: 'Policy reference' },
      ] },
      { type: 'text', text: 'Where: {Business.Name}, {Business.Address}, {Business.City} {Business.Postcode}. Please come to reception five minutes before your time. The front door has step-free access and every treatment room is on the ground floor.' },
      { type: 'text', text: 'If your insurer is paying, please check that your policy covers this appointment and bring your policy reference with you. If you pay for yourself, the fee of {Service.Price} is paid at the desk on the day, by card or cash.' },
      { type: 'text', text: 'If you need to cancel or move the appointment, please ring us on {Business.Phone} or write to {Business.E-mail} at least 24 hours before. We may charge half the fee for an appointment that is missed, or cancelled with less notice, because we cannot offer the time to someone else.' },
      { type: 'text', text: 'We look forward to seeing you.' },
      { type: 'text', text: 'Yours sincerely,' },
      { type: 'spacer', size: 24 },
      { type: 'text', text: 'Reception, {Business.Name}' },
    ] },
    { name: 'Recall letter', table: 'Recalls', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'A reminder from {Business.Name}' },
      { type: 'fields', columns: 2, items: [
        { field: 'Patient', label: 'To' }, { field: 'Patient.Patient No', label: 'Patient number' },
        { field: 'Patient.Address', label: 'Address' }, { field: 'Patient.City', label: 'Town' }, { field: 'Patient.Postcode', label: 'Postcode' },
      ] },
      { type: 'text', text: 'Dear {Patient.First Name},' },
      { type: 'text', text: 'This is a friendly reminder that you are due to come back and see us for your {Service.Name}, around {Due Date}. It takes about {Service.Minutes} minutes.' },
      { type: 'text', text: 'To book a time that suits you, ring us on {Business.Phone}, write to {Business.E-mail}, or book online at {Business.Website}.' },
      { type: 'text', text: 'If you have already booked, or have had this done elsewhere, please ignore this letter, or tell us and we will update our records. If you would rather not have reminders from us, just let us know.' },
      { type: 'text', text: 'With best wishes,' },
      { type: 'spacer', size: 24 },
      { type: 'text', text: 'The team at {Business.Name}' },
    ] },
    { name: 'Insurer statement', table: 'Insurers', paper: 'a4', letterhead: LETTERHEAD, blocks: [
      { type: 'heading', text: 'Statement of claims' },
      { type: 'fields', columns: 2, items: [
        { field: 'Name', label: 'Insurer' }, { field: 'Payment Days', label: 'Terms (days)' },
        { field: 'Address', label: 'Address' }, { field: 'E-mail', label: 'Claims e-mail' },
        { field: 'City', label: 'Town' }, { field: 'Postcode', label: 'Postcode' },
      ] },
      { type: 'lines', table: 'Invoices', link: 'Insurer', sort: 'Invoice Date', columns: [
        { field: 'Number', label: 'Invoice' }, { field: 'Invoice Date', label: 'Date' }, { field: 'Appointment.Patient', label: 'Patient' },
        { field: 'Policy Reference', label: 'Policy' }, { field: 'Appointment.Service.Name', label: 'Service' }, { field: 'Due Date', label: 'Due' },
        { field: 'Amount', label: 'Claimed', type: 'currency' }, { field: 'Amount Paid', label: 'Paid', type: 'currency' },
        { expr: 'IIf(Nz([Paid On], "") = "", Round([Amount] - Nz([Amount Paid]), 2), 0)', as: 'Owing', type: 'currency' },
      ] },
      { type: 'totals', rows: [
        { label: 'Claimed', expr: '[Amount]' }, { label: 'Paid', expr: '[Amount Paid]' }, { label: 'Balance owing', expr: '[Owing]', strong: true },
      ] },
      { type: 'text', text: 'Please pay the balance owing to {Business.Bank Details}, quoting each invoice number, and send the remittance list to {Business.E-mail}. If you have paid since this statement was made, thank you.', small: true },
    ] },
  ],
  reports: [
    { name: 'Day sheet', title: 'Day sheets for the week ahead', query: 'Day sheet', group: 'Sheet', columns: ['Time', 'Patient', 'Patient No', 'Phone', 'Service', { field: 'Minutes', type: 'integer' }, 'Paid By'], totals: [{ field: 'Patient', fn: 'count' }, { field: 'Minutes', fn: 'sum' }] },
    { name: 'Invoices due', query: 'Unpaid invoices', group: 'Owed By', columns: ['Invoice', 'Patient', 'For', { field: 'Invoice Date', type: 'date' }, { field: 'Due Date', type: 'date' }, { field: 'Days Overdue', type: 'integer' }, { field: 'Balance', type: 'currency' }], totals: [{ field: 'Balance', fn: 'sum' }] },
    { name: 'Takings by month', query: 'Takings by month', columns: ['Month', { field: 'From Patients', type: 'currency' }, { field: 'From Insurers', type: 'currency' }, { field: 'Takings', type: 'currency' }], totals: [{ field: 'From Patients', fn: 'sum' }, { field: 'From Insurers', fn: 'sum' }, { field: 'Takings', fn: 'sum' }] },
    { name: 'Takings by practitioner', title: 'Takings by practitioner, last 12 months', query: 'Takings by practitioner', columns: ['Practitioner', 'Role', { field: 'Payments', type: 'integer' }, { field: 'Takings', type: 'currency' }], totals: [{ field: 'Takings', fn: 'sum' }] },
    { name: 'Recall letters', title: 'Recall letters to send', query: 'Recall letters to send', group: 'Recall For', columns: ['Patient', 'Address', 'City', 'Postcode', { field: 'Due Date', type: 'date' }], totals: [{ field: 'Patient', fn: 'count' }] },
    { name: 'Patient list', query: 'Patient list', group: 'Insurer', columns: ['Last Name', 'First Name', 'Patient No', { field: 'Date of Birth', type: 'date' }, 'Phone', 'Policy Reference'], totals: [{ field: 'Patient No', fn: 'count' }] },
  ],
};

export const about = {
  name: 'Cedar Lane Clinic',
  kind: 'Clinic',
  line: 'Patients, appointments, visit notes, invoices and recalls for a small private clinic.',
  summary: [
    'For a small clinic that books appointments, writes a short note at each visit and bills patients or their insurers. Cedar Lane Clinic is invented: two doctors, a practice nurse and two physiotherapists, with a reception desk that takes bookings by phone and online.',
    'The records cover the year up to the day the file is made, and the weeks after it. Patients book a few days ahead. Some cancel and a few do not turn up, and the clinic charges a fee for some of those. A physiotherapy assessment leads to a course of weekly sessions, a health check sets a recall for a year later, and the doctors are busier in winter while the nurse is busiest in the autumn when flu vaccinations are given.',
    'Each visit has an invoice, made out to the patient or, when the policy covers the service, to the insurer. Most patients pay at the desk; insurers pay within their terms, one of them slowly, so there is always something to chase. The saved queries answer what a practice manager asks each day and each month: who is booked this week, how full each diary is, who missed an appointment, who owes what, what came in, and which patients are due to be recalled.',
    'It prints what the desk sends out: invoices numbered INV-2026-0042 (to an insurer with the patient\'s policy reference, or to the patient), receipts for payments, a letter confirming each appointment, a recall letter, and a statement of claims for each insurer. Medical care is exempt from VAT, so no VAT is charged and the invoices say so.',
    'Every patient and every note in it is made up. The notes are short and ordinary on purpose; a real clinic keeps its own wording and its own rules on who may read them.',
  ].join('\n\n'),
  suits: [
    'A family doctor or GP practice that sees private patients',
    'A physiotherapy or sports injury practice',
    'A chiropractic or osteopathy clinic',
    'A dental surgery, with treatments in place of services',
    'A podiatry or foot care clinic',
    'A veterinary practice, with pets as patients and owners for contact details',
    'A counselling or talking therapy practice',
  ],
  tables: {
    Business: BUSINESS_ABOUT,
    Insurers: {
      about: 'The health insurers whose policies the clinic accepts, and one record called Self-pay for patients who pay for themselves. Every patient points at one of these, so a patient list can always say who pays.',
      fields: {
        Name: 'The insurer as it appears on claims and on the patient list. The Self-pay record stands for no insurer.',
        Kind: 'Insurer for a real insurer, Self-pay for the record that stands for paying yourself. Queries use it to tell the two apart.',
        Address: 'The street address claims and statements are posted to.',
        City: 'The town or city of the claims address.',
        Postcode: 'The postcode of the claims address.',
        Phone: 'The claims line the reception desk rings about a claim.',
        'E-mail': 'Where claims and questions about them are sent.',
        'Payment Days': 'How many days the insurer takes to pay a claim under its agreement with the clinic. An invoice to the insurer falls due this many days after the visit.',
        Notes: 'Anything reception should know: what a policy usually covers, or that an insurer wants a referral letter first.',
      },
    },
    Patients: {
      about: 'One record per patient registered with the clinic, with contact details, who pays for their care and whether they agree to reminders.',
      fields: {
        'Patient No': 'The clinic\'s own number for the patient, printed on day sheets and letters. It is not a national health number.',
        'First Name': 'The patient\'s first name.',
        'Last Name': 'The patient\'s family name. Lists are sorted by it.',
        'Date of Birth': 'Used to tell two patients with the same name apart and to see who is due a check for their age.',
        Address: 'The street address, for recall letters and invoices.',
        City: 'The town or city.',
        Postcode: 'The postcode, for letters.',
        Phone: 'The number reception rings to confirm or move an appointment.',
        'E-mail': 'Where appointment reminders and invoices are sent, if the patient has given one.',
        Insurer: 'Who pays: the patient\'s insurer, or Self-pay.',
        'Policy Reference': 'The patient\'s policy or membership reference with the insurer, quoted on every claim. Empty for self-pay patients.',
        Registered: 'The day the patient first registered. The new patients query counts these by month.',
        Reminders: 'Ticked when the patient agrees to reminders and recall letters. Patients without it are rung instead, or not contacted.',
        Notes: 'Practical things reception should know before a visit, such as needing step-free access. Not clinical notes, which go in Visit Notes.',
      },
    },
    Practitioners: {
      about: 'The people who see patients: the doctors, the nurse and the physiotherapists. Reception staff are not here because they do not hold a diary.',
      fields: {
        'First Name': 'The practitioner\'s first name.',
        'Last Name': 'The practitioner\'s family name.',
        'Known As': 'The name used on day sheets and letters, such as Dr and the name, or the name and Physio.',
        Role: 'Doctor, Nurse or Physiotherapist. Services say which role gives them.',
        'E-mail': 'The practitioner\'s clinic e-mail address.',
        Phone: 'The practitioner\'s extension or direct line.',
        Works: 'The days and hours the practitioner sees patients, in words.',
        'Hours per Week': 'The hours a week the practitioner has for appointments. The diary query measures bookings against it.',
        Started: 'The day the practitioner started at the clinic.',
      },
    },
    Services: {
      about: 'What the clinic offers, how long each takes in the diary and what it costs.',
      fields: {
        Name: 'The service as it appears in the diary and on invoices.',
        Kind: 'A group of services: Consultation, Follow-up, Health check, Nursing, Vaccination or Physiotherapy.',
        'Given By': 'The role that gives it: Doctor, Nurse or Physiotherapist.',
        Minutes: 'How long the appointment lasts. The diary is kept in quarter hours, so these are multiples of 15.',
        Price: 'The fee for one appointment.',
        'Insurer Pays': 'Ticked when insurers usually cover the service. Health checks and vaccinations are paid by the patient even when insured.',
        'Recall Months': 'How many months after this service the patient should come back for it, such as 12 for a yearly health check. Empty when there is no recall.',
      },
    },
    Appointments: {
      about: 'Every booking in the diary, past and future, with what came of it.',
      fields: {
        Patient: 'Who the appointment is for.',
        Practitioner: 'Whose diary it is in.',
        Service: 'What the appointment is for, which sets its length and price.',
        Start: 'The date and time the appointment starts.',
        Status: 'Booked until the day, then Attended or No-show. Cancelled when the patient or the clinic called it off.',
        'Booked On': 'The day the booking was made.',
        'Booked Via': 'How it was booked: Phone, Online, At the desk, or Recall when it came from a recall letter or call.',
        'Cancelled On': 'The day the appointment was cancelled. A cancellation the day before or on the day counts as late.',
      },
    },
    'Visit Notes': {
      about: 'A short note written by the practitioner after each attended appointment. In this sample they are invented and kept to everyday words.',
      fields: {
        Appointment: 'The visit the note is about, which gives the patient, the practitioner and the date.',
        Written: 'When the note was written, usually just after the appointment ended.',
        Reason: 'Why the patient came, in a few words.',
        Note: 'What was found and what was done or advised.',
        Plan: 'What happens next: nothing, a follow-up, a test, a course of sessions or a recall.',
      },
    },
    Invoices: {
      about: 'What the clinic charges, one invoice per visit, and for a missed appointment or a late cancellation when a fee is charged. Each invoice keeps the name and address it was made out to, so a reprint reads as the one that was sent.',
      fields: {
        Number: 'The invoice\'s number, given in date order and once only, a new run each year: INV-2026-0042. Quote it on claims and payments.',
        Appointment: 'The appointment the invoice is for.',
        'Invoice Date': 'The day the invoice was raised: the day of the visit, or of the missed or cancelled appointment.',
        'Due Date': 'When payment is due: 14 days for patients, the insurer\'s payment days for claims.',
        'Bill To': 'Patient or Insurer. An insurer is billed when the patient is insured and the service is one insurers pay for.',
        Insurer: 'The insurer the claim is made to. Empty when the patient is billed. The insurer\'s statement lists the invoices that point at it.',
        'Bill To Name': 'Who the invoice is made out to: the insurer\'s name, or the patient\'s.',
        'Bill To Address': 'The street address the invoice went to.',
        'Bill To City': 'The town or city the invoice went to.',
        'Bill To Postcode': 'The postcode the invoice went to.',
        'Policy Reference': 'The patient\'s policy reference, copied on to a claim so the insurer can match it. Empty on an invoice to the patient.',
        For: 'Visit, Missed appointment or Late cancellation.',
        Amount: 'The amount charged.',
        'Amount Paid': 'How much has been paid so far, the sum of the payments against it.',
        'Paid On': 'The day the invoice was paid in full. Empty while anything is still owed.',
      },
    },
    Payments: {
      about: 'Money received against invoices. An invoice can have more than one payment when a patient pays in parts. Each payment has a receipt.',
      fields: {
        Number: 'The receipt\'s number, given in date order, a new run each year: RCP-2026-0042.',
        Invoice: 'The invoice the payment is for.',
        'Paid On': 'The day the money came in.',
        Amount: 'How much was paid.',
        Method: 'Card, Cash or Bank transfer from a patient, or Insurer for a claim paid by an insurer.',
        Reference: 'The insurer\'s remittance reference or the reference on a bank transfer.',
      },
    },
    Recalls: {
      about: 'Patients the clinic should ask back: for a yearly health check, a blood pressure review, the autumn flu vaccination or a review after physiotherapy.',
      fields: {
        Patient: 'Who to recall.',
        Service: 'The service they are being asked back for.',
        'Due Date': 'When the recall falls due.',
        Reason: 'Why, in words for the letter or the call.',
        'Sent On': 'The day a recall letter or reminder went out. Letters go each Monday for recalls due within two weeks.',
        'Booked On': 'The day the patient booked in response. Empty while the recall is still open.',
      },
    },
  },
  queries: {
    'Next 7 days': 'For each day of the coming week and each practitioner, how many appointments are booked and the hours they fill, with the first and last start times. Reception uses it to see the shape of the week at a glance.',
    'Day sheet': 'Every booked appointment for the next seven days, by day and practitioner, with the patient\'s number, phone, who pays and any practical notes. It is what each practitioner and the desk work from each morning.',
    'Diary next 14 days': 'How full each practitioner\'s diary is for the next two weeks: appointments, hours booked, hours open and the percentage full. Use it to see who can take more patients or where to offer a cancellation slot.',
    'No-shows and late cancellations': 'Appointments in the last 90 days the patient did not attend, or cancelled the day before or on the day, with the phone number. Use it to decide on fees and to talk to patients who miss often.',
    'Unpaid invoices': 'Every invoice not yet paid in full, with who owes it, what has been paid and how many days overdue it is. A minus figure means it is not due yet.',
    'Money owed': 'The unpaid balance added up for each insurer and each patient, with the number of invoices and the oldest due date. Start here when chasing payments.',
    'Takings by month': 'Money received each month, split into what came from patients and what came from insurers.',
    'Takings by practitioner': 'Money received in the last twelve months for each practitioner\'s appointments, and how many payments it came in.',
    'Fees by service': 'For each service, how many visits were invoiced in the last twelve months and the total charged. It shows which services earn the most.',
    'Busy days': 'Appointments attended or missed in the last year by day of the week, with the hours they took. Use it to plan staff hours.',
    'Busy hours': 'The same appointments by the hour they started, which shows the busy mornings and late afternoons.',
    'Recalls due': 'Open recalls falling due in the next 30 days, with contact details and whether a letter has gone out or the patient has agreed to letters at all.',
    'Recall letters to send': 'Recalls due in the next 30 days for patients who agree to letters and have not been sent one yet, with their address. It feeds the recall letters report.',
    'New patients by month': 'How many patients registered each month over the last year, and how many of them are insured.',
    'Not seen for six months': 'Patients whose last attended appointment was more than six months ago, with the date of that visit. Use it for a "we have not seen you for a while" letter or to tidy the list.',
    'Patient list': 'All patients by family name with their date of birth, contact details, insurer and policy reference.',
    'Visit history': 'Every visit note by patient, newest first, with the practitioner, the service, the reason and the plan.',
  },
  reports: {
    'Day sheet': 'A page for each practitioner and day of the week ahead, listing each appointment\'s time, patient, phone, service and who pays, with the number of patients and minutes booked. Printed each morning for the practitioner and the desk.',
    'Invoices due': 'Unpaid invoices grouped by who owes them, each insurer and each patient, with the balance and days overdue and a total for each. For whoever chases payments.',
    'Takings by month': 'A month by month list of money received from patients and from insurers, with totals. For the owners and the accountant.',
    'Takings by practitioner': 'Money received in the last twelve months for each practitioner, with a total. Useful when practitioners are paid a share of their takings.',
    'Recall letters': 'The names and addresses of patients to send a recall letter to, grouped by what they are recalled for. Print it, or use it for a mail merge.',
    'Patient list': 'Patients grouped by insurer, with patient number, date of birth, phone and policy reference, and a count for each insurer.',
  },
  documents: {
    Invoice: 'The invoice for a visit, a missed appointment or a late cancellation, made out to the insurer with the patient\'s policy reference, or to the patient. It shows the visit, any payments received, and the balance due, and says that medical care is exempt from VAT. Numbered INV-2026-0042.',
    'Payment receipt': 'Given to a patient who pays at the desk, or sent when a transfer arrives: who paid, how much, for which invoice and visit. Numbered RCP-2026-0042.',
    'Appointment letter': 'Sent when an appointment is booked: the date and time, who it is with, where to come, who pays, and how to cancel or move it. Every appointment can print one, past ones too, so print it from the Appointments form for the one on screen.',
    'Recall letter': 'A short, friendly letter asking a patient back for a check that is due, with how to book. Print it for the patients in the Recall letters to send query, from the Recalls form. Have the wording checked against the rules where you work.',
    'Insurer statement': 'Sent to each insurer: every claim made to it, what it has paid and the balance owing. The Self-pay record prints an empty one, which can be left out.',
  },
  tutorial: [
    { title: 'Book an appointment and send the letter', steps: [
      'Open the Patients form and find the patient with the search box, or press New and fill in their name, date of birth, address, phone, Insurer and Policy Reference.',
      'Open the Diary next 14 days query to see whose diary has room.',
      'Open the Appointments form and press New. Choose the Patient, the Practitioner and the Service, type the Start date and time, and set Booked On to today and Booked Via to Phone, Online or At the desk. Status stays Booked.',
      'Open the print template Appointment letter from Print templates, choose the appointment you just made, and press Print, or Save as PDF to e-mail it.',
    ] },
    { title: 'Run the day', steps: [
      'Print the Day sheet report in the morning: a page for each practitioner and day of the week ahead.',
      'As each patient arrives and is seen, open their appointment in the Appointments form and set Status to Attended. If they do not come, set it to No-show.',
      'After the visit the practitioner opens the Visit Notes form, presses New, chooses the appointment and writes the Reason, the Note and the Plan.',
      'If the patient cancels, set Status to Cancelled and type the Cancelled On date. The No-shows and late cancellations query lists those with a day or less of notice.',
    ] },
    { title: 'Invoice a visit and take payment', steps: [
      'Open the Invoices form and press New. Type the next Number (INV, the year, and the next number in the run), choose the Appointment, and type the Invoice Date.',
      'For an insured patient on a service insurers pay for, set Bill To to Insurer, choose the Insurer, and copy the insurer\'s name and address and the patient\'s Policy Reference into the Bill To fields. Otherwise set Bill To to Patient and copy the patient\'s name and address. Set Due Date 14 days on for a patient, or the insurer\'s Payment Days on.',
      'Type the Amount from the service\'s Price, or half of it for a missed appointment or a late cancellation, and set For to match.',
      'When the patient pays, open the Payments form, press New, give it the next RCP number, choose the Invoice, and type Paid On, Amount and Method. Then add the amount to the invoice\'s Amount Paid, and type Paid On on the invoice once it is paid in full.',
      'Print the Payment receipt template for that payment, and the Invoice template for the invoice when it is sent.',
    ] },
    { title: 'Chase what is owed', steps: [
      'Open the Money owed query: the unpaid balance for each insurer and each patient, the largest first.',
      'Print the Invoices due report for the list with days overdue.',
      'Once a month, print the Insurer statement template for each insurer and send it with a covering note.',
      'When an insurer pays, record one payment for each invoice on its remittance list, Method Insurer, with the remittance reference in Reference.',
    ] },
    { title: 'Send the recall letters', steps: [
      'Each Monday, open the Recall letters to send query: open recalls due in the next 30 days for patients who agree to letters.',
      'Open the Recalls form, find each of those recalls, and print the Recall letter template for it.',
      'Type today\'s date in Sent On on each recall you sent a letter for.',
      'When the patient books, type the date in Booked On. The Recalls due query lists the ones still open, including patients who do not want letters, for a phone call instead.',
    ] },
    { title: 'Look at the month', steps: [
      'Print the Takings by month report for money received from patients and from insurers.',
      'Open the Fees by service query to see which services earn the most, and Takings by practitioner if practitioners are paid a share.',
      'Open the Busy days and Busy hours queries to plan staff hours, and New patients by month to see how the list is growing.',
    ] },
  ],
  notes: [
    'An invoice\'s Amount Paid and Paid On are not worked out by themselves: when you add a payment, add its amount to the invoice\'s Amount Paid, and type Paid On when the invoice is paid in full. The unpaid invoice queries read those two fields.',
    'The Bill To fields are copied on to each invoice when it is made, so a patient who moves house, or an insurer that changes its address, does not change invoices already sent.',
    'Medical care is exempt from VAT, so Business has a Tax Rate of nought and the invoice\'s VAT row comes to nothing. A clinic that also sells things VAT is charged on would put its rate in Business and keep those sales on invoices of their own.',
    'Visit Notes holds short notes in everyday words. A real clinic decides who may open that table, and keeps clinical records under the rules where it works.',
    'Each print template prints every record of its table: every appointment has a letter, every insurer a statement. Print the one on screen from its form, or choose the records to print.',
    'The Self-pay record in Insurers stands for patients who pay for themselves. Do not delete it: every patient must point at an insurer or at Self-pay.',
  ],
  faq: [
    { q: 'Why does the invoice show no VAT?', a: 'Medical care is exempt from VAT in the UK, so the clinic charges none and is not registered. The invoice says so in its small print. If your services are taxed, put the rate in Business, Tax Rate, and the invoice adds it.' },
    { q: 'How do I bill an insurer instead of the patient?', a: 'On the invoice, set Bill To to Insurer, choose the Insurer, copy the insurer\'s claims address into the Bill To fields and the patient\'s Policy Reference into Policy Reference. The claim then appears on that insurer\'s statement.' },
    { q: 'A patient paid half now and half later. How do I record it?', a: 'Add two records in Payments, each pointing at the same invoice, each with its own receipt number. Add both to the invoice\'s Amount Paid; type Paid On on the invoice only when the second arrives.' },
    { q: 'How do I print last month\'s invoices?', a: 'Open the Invoices form, choose the invoices dated last month (sort by Invoice Date), and print the Invoice template for those chosen.' },
    { q: 'Can I change what the appointment letter says?', a: 'Yes. Open the Appointment letter template from Print templates and edit its text blocks. Words in curly brackets, such as {Start} or {Business.Phone}, are filled from the appointment and from the Business record.' },
    { q: 'How do I stop sending recall letters to someone?', a: 'Untick Reminders on their record in Patients. The Recall letters to send query leaves them out, and Recalls due still lists them so reception can ring instead.' },
    { q: 'What is the cancellation fee?', a: 'The sample charges half the fee for some missed appointments and cancellations with a day or less of notice, as an invoice with For set to Missed appointment or Late cancellation. Set your own rule and put it in the Business record\'s Document Footer, which every document prints.' },
  ],
  empty: 'Start with Business: the clinic\'s name, address, phone and bank details, which every invoice, receipt and letter prints. Then Insurers: add a Self-pay record first (Kind Self-pay), then each insurer you accept with its payment days. Then add your Practitioners with their hours, then your Services with their minutes and prices. Once those are in, register Patients, pointing each at an insurer or at Self-pay, and book Appointments. Write a Visit Note after each attended appointment, raise an Invoice for it, and record Payments as they arrive, updating the invoice\'s Amount Paid and Paid On. Add Recalls as visits call for them.',
  adapt: [
    'For a dental surgery, rename Services to Treatments and add check-ups, scale and polish, fillings and crowns with their minutes and fees, then add a Tooth field to Visit Notes.',
    'For a vet, add a Pets table pointing at the owner in Patients, with species and breed, and point Appointments at the pet instead of the owner.',
    'Add a Room field to Appointments, or a Rooms table, if the treatment rooms are shared and you need to see which is free.',
    'Add a Referred By field to Patients and a query that counts new patients by where they heard of you.',
    'If you take deposits for a course of treatment, add a Deposit For field to Payments and a status of Deposit paid to Invoices.',
    'Add an Excess field to Patients for insurers that leave the patient to pay the first part of a claim, and raise a second invoice to the patient for it.',
    'Add a Consent Signed date to Patients if your practice asks for written consent before treatment, and a query listing patients booked this week who have none.',
  ],
};

// The practitioners and the hours they see patients: weekday (0 Sunday) to [from, to] in whole hours.
const TEAM = [
  { role: 'Doctor', title: 'Dr', works: 'Mon, Wed and Fri 9 to 1, Tue 2 to 6', pattern: { 1: [9, 13], 2: [14, 18], 3: [9, 13], 5: [9, 13] } },
  { role: 'Doctor', title: 'Dr', works: 'Tue and Thu, 9 to 5', pattern: { 2: [9, 17], 4: [9, 17] } },
  { role: 'Nurse', title: 'Nurse', works: 'Mon, Wed and Thu, 8 to 12', pattern: { 1: [8, 12], 3: [8, 12], 4: [8, 12] } },
  { role: 'Physiotherapist', works: 'Mon and Thu 9 to 5, Sat 9 to 1', pattern: { 1: [9, 17], 4: [9, 17], 6: [9, 13] } },
  { role: 'Physiotherapist', works: 'Wed and Fri, 12 to 7', pattern: { 3: [12, 19], 5: [12, 19] } },
];

/** The services: [name, kind, given by, minutes, price, insurer pays, recall months]. */
const SERVICES = [
  ['GP consultation', 'Consultation', 'Doctor', 15, 85, true, null],
  ['Long GP consultation', 'Consultation', 'Doctor', 30, 145, true, null],
  ['GP follow-up', 'Follow-up', 'Doctor', 15, 65, true, null],
  ['Blood pressure review', 'Follow-up', 'Doctor', 15, 60, true, 6],
  ['Annual health check', 'Health check', 'Doctor', 45, 195, false, 12],
  ['Blood test', 'Nursing', 'Nurse', 15, 45, true, null],
  ['Wound care', 'Nursing', 'Nurse', 15, 40, true, null],
  ['Flu vaccination', 'Vaccination', 'Nurse', 15, 22, false, 12],
  ['Travel vaccination', 'Vaccination', 'Nurse', 15, 65, false, null],
  ['Nurse health check', 'Health check', 'Nurse', 30, 95, false, 12],
  ['Physiotherapy assessment', 'Physiotherapy', 'Physiotherapist', 45, 80, true, null],
  ['Physiotherapy session', 'Physiotherapy', 'Physiotherapist', 30, 60, true, null],
  ['Physiotherapy review', 'Physiotherapy', 'Physiotherapist', 30, 60, true, null],
  ['Sports massage', 'Physiotherapy', 'Physiotherapist', 45, 65, false, null],
];

/** Invented insurers: [name, short code for policy references, payment days, notes, claims address]. The last pays slowly. */
const INSURERS = [
  ['Kestrel Health Cover', 'KHC', 30, 'Covers consultations, tests and physiotherapy. Wants the policy reference on every claim.', ['Claims Office, 40 Canal Walk', 'Leeds', 'LS1 4AB']],
  ['Alder Medical Assurance', 'AMA', 28, 'Physiotherapy limited to ten sessions a year on most policies.', ['PO Box 7100, 12 Market Square', 'Manchester', 'M2 6DE']],
  ['Granite Health Plan', 'GHP', 30, 'Pays by bank transfer each week with a remittance list.', ['Granite House, 9 Riverside Drive', 'Cardiff', 'CF10 3FG']],
  ['Lantern Care Insurance', 'LCI', 45, 'Often slow to pay; chase claims older than 45 days by phone.', ['Claims Team, 221 Station Road', 'Birmingham', 'B4 7HJ']],
];

// What each kind of visit is for, and what the note says. Ordinary complaints only.
const GP_REASONS = ['Cough for two weeks', 'Sore throat', 'Lower back pain', 'Itchy skin rash', 'Feeling tired', 'Headaches', 'Ear ache', 'Knee pain', 'Trouble sleeping', 'Hay fever', 'Indigestion', 'Sore shoulder', 'Stiff neck', 'Mild eczema', 'Check-up after a cold'];
const GP_FOUND = ['Examined, nothing of concern found.', 'Chest clear, temperature normal.', 'Mild, likely to settle on its own.', 'Symptoms already easing.', 'Talked through what has helped so far.'];
const GP_ADVICE = ['Advised rest, fluids and simple pain relief from the pharmacy.', 'Gave a leaflet on gentle stretches.', 'Suggested a sleep diary for two weeks.', 'Suggested a soothing cream from the pharmacy.', 'Advised to keep active within comfort.'];
const PHYSIO_REASONS = ['Lower back pain', 'Stiff shoulder', 'Sprained ankle', 'Neck pain from desk work', 'Knee pain after running', 'Sore elbow from racket sport', 'Hip stiffness', 'Hamstring strain'];
const PATIENT_NOTES = ['Prefers morning appointments', 'Needs step-free access', 'Hard of hearing, please speak clearly', 'Usually comes with a carer', 'Prefers appointments after 4 pm', 'Prefers to be called rather than e-mailed', 'Nervous about injections'];
const STREETS = ['Cedar Lane', 'Mill Road', 'Station Street', 'Orchard Way', 'Church Walk', 'Linden Close', 'Beech Grove', 'Quarry Hill', 'Canal Side', 'Market Row', 'Willow Crescent', 'Park View'];
const TOWNS = [['Bristol', 10], ['Keynsham', 3], ['Portishead', 2], ['Clevedon', 1]];
const HOUR_LIKED = { 8: 1.3, 9: 1.4, 10: 1.1, 11: 1, 12: 1.1, 13: 1, 14: 0.8, 15: 0.9, 16: 1.2, 17: 1.4, 18: 1.3 };

/** How much busier each role is in each month (January first). */
const SEASON = {
  Doctor: [1.15, 1.1, 1.0, 0.95, 0.9, 0.9, 0.85, 0.8, 0.95, 1.05, 1.1, 0.95],
  Nurse: [0.95, 0.95, 1, 1, 1.05, 1.05, 1, 0.85, 1.1, 1.3, 1.2, 0.85],
  Physiotherapist: [0.9, 0.95, 1.05, 1.1, 1.1, 1.05, 1.0, 0.9, 1.05, 1.0, 0.95, 0.8],
};

/** Fill a Cedar Lane Clinic database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal, days = 365 }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };
  const pad = (n, w = 2) => String(n).padStart(w, '0');

  // Days of the week and closed days, worked out once for each day.
  const horizon = 70;
  const calendarDays = new Map();
  const dayOf = (d) => {
    if (!calendarDays.has(d)) {
      const a = cal.at(d);
      const m = a.getMonth();
      const n = a.getDate();
      calendarDays.set(d, { w: a.getDay(), m, closed: (m === 11 && (n === 25 || n === 26)) || (m === 0 && n === 1) });
    }
    return calendarDays.get(d);
  };
  const wd = (d) => dayOf(d).w;
  const monthOf = (d) => dayOf(d).m;
  const closed = (d) => dayOf(d).closed;

  // The clinic's own details, printed at the head of every invoice, receipt and letter. Medical
  // care is exempt from VAT, so its rate is nought and it has no VAT number.
  add('Business', {
    ...businessRecord({
      name: 'Cedar Lane Clinic', domain: 'cedarlane.example', phone: '020 7946 0150', taxRate: 0,
      address: { Address: '2 Cedar Lane', City: 'Bristol', Postcode: 'BS6 5AB', Country: 'United Kingdom' },
      terms: 'Patients pay at the desk on the day or within 14 days of the invoice. Insurers pay within the days agreed with the clinic.',
      footer: 'Please give us 24 hours notice to cancel or move an appointment; a missed appointment or a late cancellation may be charged at half the fee. Open Monday to Friday from 8 am, and Saturday mornings.',
    }),
    'Tax Number': 'Not registered for VAT: medical care is exempt',
  });

  // Insurers, with Self-pay first so it is easy to find.
  const selfPay = { id: add('Insurers', { Name: 'Self-pay', Kind: 'Self-pay', 'Payment Days': 14, Notes: 'Patients who pay for themselves, at the desk or within 14 days.' }), real: false, days: 14 };
  const insurers = INSURERS.map(([name, code, payDays, notes, [street, city, postcode]], i) => ({
    id: add('Insurers', { Name: name, Kind: 'Insurer', Address: street, City: city, Postcode: postcode, Phone: `020 7946 0${pad(310 + i * 7, 3)}`, 'E-mail': `claims@${name.toLowerCase().replace(/[^a-z]+/g, '')}.example`, 'Payment Days': payDays, Notes: notes }),
    name, street, city, postcode, code, days: payDays, real: true, slow: i === INSURERS.length - 1,
  }));

  // Practitioners, with the quarter hours each can be booked on each weekday.
  const team = TEAM.map((t, i) => {
    const p = person(r, { domain: 'cedarlane.example' });
    const slots = {};
    let hours = 0;
    for (const [w, [from, to]] of Object.entries(t.pattern)) {
      // A whole day has a lunch hour from 1 to 2; a half day or a late shift does not.
      const lunch = from <= 10 && to >= 16;
      const list = [];
      for (let s = from * 4; s < to * 4; s++) if (!(lunch && s >= 52 && s < 56)) list.push(s);
      slots[w] = { list, set: new Set(list) };
      hours += list.length / 4;
    }
    const known = t.title ? `${t.title} ${p['First Name']} ${p['Last Name']}` : `${p['First Name']} ${p['Last Name']} (Physio)`;
    const id = add('Practitioners', {
      'First Name': p['First Name'], 'Last Name': p['Last Name'], 'Known As': known, Role: t.role, 'E-mail': p['E-mail'],
      Phone: `020 7946 0${pad(200 + i, 3)}`, Works: t.works, 'Hours per Week': hours, Started: cal.day(-r.int(500, 4000)),
    });
    return { id, i, role: t.role, slots, hours, leave: new Set() };
  });
  // Leave: a fortnight in July or August and a week some other time, all in the past.
  for (const p of team) {
    const summer = [];
    for (let d = -days; d < -20; d++) if ([6, 7].includes(monthOf(d)) && wd(d) === 1) summer.push(d);
    if (summer.length) { const s = r.pick(summer); for (let k = 0; k < 14; k++) p.leave.add(s + k); }
    const s = r.int(-days, -30);
    if (!p.leave.has(s)) for (let k = 0; k < 7; k++) p.leave.add(s + k);
  }
  const byRole = (role) => team.filter((p) => p.role === role);
  const works = (p, d) => !closed(d) && !p.leave.has(d) && p.slots[wd(d)];

  // Services.
  const services = {};
  for (const [name, kind, by, minutes, price, covered, recall] of SERVICES) {
    services[name] = { name, kind, by, minutes, price, covered, recall, n: minutes / 15 };
    services[name].id = add('Services', { Name: name, Kind: kind, 'Given By': by, Minutes: minutes, Price: price, 'Insurer Pays': covered, 'Recall Months': recall });
  }
  const S = (name) => services[name];

  // Patients: those registered before the records begin, then new ones as they come.
  const patients = [];
  const pool = []; // each patient once to three times over: some come far more often than others
  let patientNo = 0;
  // Each patient has a number of their own from the range kept for fiction, and an e-mail address
  // no one else has: two patients with the same name get a number in theirs.
  const mobiles = r.some(Array.from({ length: 1000 }, (_, i) => i), 1000);
  const phone = () => (mobiles.length ? `07700 900${pad(mobiles.pop(), 3)}` : `020 7946 0${r.int(400, 999)}`);
  const mails = new Set();
  const newPatient = (registered) => {
    const p = person(r, { domain: r.pick(['example.com', 'example.net', 'mail.example']) });
    for (let k = 2; mails.has(p['E-mail']); k++) p['E-mail'] = p['E-mail'].replace(/^([^@]*?)\d*@/, `$1${k}@`);
    mails.add(p['E-mail']);
    const age = r.weighted([[r.int(1, 17), 15], [r.int(18, 64), 60], [r.int(65, 90), 25]]);
    const born = registered - age * 365 - r.int(0, 364);
    const ins = r.chance(0.55) ? r.weighted([[insurers[0], 4], [insurers[1], 3], [insurers[2], 2], [insurers[3], 2]]) : selfPay;
    const town = r.weighted(TOWNS);
    const pat = { born, ins, reminders: r.chance(0.8), recalls: new Map(), name: `${p['First Name']} ${p['Last Name']}`, city: town };
    // The address is kept on the patient too, for the invoices made out to them.
    pat.street = `${r.int(1, 180)} ${r.pick(STREETS)}`;
    pat.postcode = `BS${r.int(1, 49)} ${r.int(1, 9)}${r.pick(['AB', 'DE', 'FG', 'HJ', 'NP', 'QR', 'TU', 'WX'])}`;
    pat.id = add('Patients', {
      'Patient No': `CL-${pad(++patientNo, 5)}`, 'First Name': p['First Name'], 'Last Name': p['Last Name'], 'Date of Birth': cal.day(born),
      Address: pat.street, City: town, Postcode: pat.postcode,
      Phone: phone(), 'E-mail': r.chance(0.85) ? p['E-mail'] : null,
      Insurer: ins.id, 'Policy Reference': (pat.policy = ins.real ? `${ins.code}-POL-${r.int(10000, 99999)}` : null), Registered: cal.day(registered),
      Reminders: pat.reminders, Notes: r.chance(0.06) ? r.pick(PATIENT_NOTES) : null,
    });
    patients.push(pat);
    const copies = r.weighted([[1, 5], [2, 3], [3, 1]]);
    for (let k = 0; k < copies; k++) pool.push(pat);
    return pat;
  };
  const established = [];
  for (let k = 0; k < 650; k++) established.push(-days - r.int(1, 3650));
  established.sort((a, b) => a - b).forEach((d) => newPatient(d));
  const ageOn = (pat, d) => (d - pat.born) / 365.25;
  const fits = (svc, pat, d) => {
    if (svc.name === 'Flu vaccination') return ageOn(pat, d) >= 50 || r.chance(0.2);
    if (svc.kind === 'Health check' || svc.kind === 'Physiotherapy') return ageOn(pat, d) >= 16;
    return true;
  };
  const pickPatient = (svc, d) => {
    if (r.chance(0.05)) return newPatient(d);
    let pat;
    for (let k = 0; k < 6; k++) { pat = r.pick(pool); if (fits(svc, pat, d)) break; }
    return pat;
  };

  // The diary: for each practitioner and day, each quarter hour taken. A taken slot holds
  // Infinity; a cancelled one holds the day it was given up, from which it can be booked again.
  const diary = new Map();
  const byDay = new Map();
  const appts = [];
  const findSlot = (p, d, n, b) => {
    const ws = p.slots[wd(d)];
    const taken = diary.get(`${p.i}:${d}`);
    const free = (s) => !taken || !taken.has(s) || taken.get(s) <= b;
    const starts = [];
    for (const s of ws.list) {
      let ok = true;
      for (let k = 0; k < n && ok; k++) ok = ws.set.has(s + k) && free(s + k);
      if (ok) starts.push([s, HOUR_LIKED[Math.floor(s / 4)] || 1]);
    }
    return starts.length ? r.weighted(starts) : null;
  };
  const LEAD = {
    'GP consultation': () => r.weighted([[0, 25], [1, 25], [r.int(2, 4), 30], [r.int(5, 10), 20]]),
    'Long GP consultation': () => r.int(2, 10), 'GP follow-up': () => r.int(7, 21), 'Blood pressure review': () => r.int(3, 21),
    'Annual health check': () => r.int(7, 30), 'Blood test': () => r.int(1, 7), 'Wound care': () => r.int(0, 3),
    'Flu vaccination': () => r.int(2, 14), 'Travel vaccination': () => r.int(5, 21), 'Nurse health check': () => r.int(7, 28),
    'Physiotherapy assessment': () => r.int(1, 10), 'Physiotherapy session': () => 7, 'Physiotherapy review': () => r.int(3, 14),
    'Sports massage': () => r.int(2, 14),
  };

  /** Book `svc` for `pat`, asked for on day `b`, with practitioner `want` if given. Answers the appointment or null. */
  const book = (b, pat, svc, { want = null, lead = LEAD[svc.name](), via = null, extra = {} } = {}) => {
    const givers = byRole(svc.by);
    const anyone = r.some(givers, givers.length);
    const target = b + lead;
    for (let t = target; t <= target + 10 && t <= horizon; t++) {
      // The practitioner asked for is tried alone for the first week, then anyone who gives the service.
      for (const p of want && t - target < 7 ? [want] : anyone) {
        if (!works(p, t)) continue;
        const s = findSlot(p, t, svc.n, b);
        if (s == null) continue;
        const key = `${p.i}:${t}`;
        if (!diary.has(key)) diary.set(key, new Map());
        const a = { pat, p, svc, d: t, s, bookedOn: b, via: via || r.weighted([['Phone', 5], ['Online', 4], ['At the desk', 2]]), status: 'Booked', cancelOn: null, ...extra };
        // Some bookings are cancelled: most with notice, some the day before or on the day.
        const cancelChance = svc.kind === 'Physiotherapy' ? 0.09 : 0.07;
        let freeFrom = Infinity;
        if (r.chance(cancelChance)) {
          const c = r.chance(0.3) || t - b < 2 ? Math.max(b, t - r.int(0, 1)) : b + r.int(0, t - b - 2);
          if (c <= 0) { a.status = 'Cancelled'; a.cancelOn = c; freeFrom = c; }
        }
        for (let k = 0; k < svc.n; k++) diary.get(key).set(s + k, freeFrom);
        appts.push(a);
        if (!byDay.has(t)) byDay.set(t, []);
        byDay.get(t).push(a);
        return a;
      }
    }
    return null;
  };

  // Recalls: what the clinic will ask patients back for.
  const recalls = [];
  const RECALL_REASON = {
    'Annual health check': 'Yearly health check due', 'Nurse health check': 'Yearly nurse health check due',
    'Blood pressure review': 'Six-monthly blood pressure review due', 'Flu vaccination': 'Seasonal flu vaccination due',
    'Physiotherapy review': 'Review after a course of physiotherapy',
  };
  const recall = (pat, svc, due, made) => {
    const open = pat.recalls.get(svc.name);
    if (open && !open.bookedOn) return open;
    const x = { pat, svc, due, made, sentOn: null, bookedOn: null };
    pat.recalls.set(svc.name, x);
    recalls.push(x);
    return x;
  };
  // Recalls already running when the records begin.
  for (const pat of patients) {
    if (!r.chance(0.2)) continue;
    const svc = S(r.weighted([['Annual health check', 3], ['Nurse health check', 2], ['Blood pressure review', 3], ['Flu vaccination', 2]]));
    if (fits(svc, pat, -days)) recall(pat, svc, r.int(-days + 7, 150), -days);
  }
  const responses = new Map(); // day: [{ pat, svc, x }]
  const respond = (d, item) => { if (!responses.has(d)) responses.set(d, []); responses.get(d).push(item); };

  // Per practitioner, fresh bookings a day: enough to fill about four fifths of the diary.
  const freshSlots = { Doctor: 1.45, Nurse: 1.2, Physiotherapist: 7.5 };
  const MIX = {
    Doctor: () => [['GP consultation', 70], ['Long GP consultation', 12], ['Blood pressure review', 6], ['Annual health check', 5]],
    Nurse: (m) => [['Blood test', 30], ['Wound care', 12], ['Flu vaccination', [8, 9, 10].includes(m) ? 70 : m === 11 ? 15 : 1], ['Travel vaccination', [3, 4, 5, 6].includes(m) ? 16 : 6], ['Nurse health check', 8]],
    Physiotherapist: () => [['Physiotherapy assessment', 75], ['Sports massage', 25]],
  };
  for (const p of team) p.rate = (p.hours * 4 * 0.5) / freshSlots[p.role] / 6;

  const invoices = [];
  const notes = [];
  const pay = (inv, d, amount, method, reference = null) => {
    if (d > 0 || amount <= 0) return;
    inv.payments.push({ d, amount: money(amount), method, reference });
  };
  const invoice = (a, d, billTo, kind, amount) => {
    const ins = a.pat.ins;
    const inv = { a, d, due: d + (billTo === 'Insurer' ? ins.days : 14), billTo, kind, amount: money(amount), payments: [] };
    invoices.push(inv);
    return inv;
  };

  // What came of a day's appointments: attended or not, the note, the invoice, what is booked next.
  const outcome = (a) => {
    const { pat, svc, d } = a;
    if (a.status === 'Cancelled') {
      // A late cancellation is charged half the fee, some of the time.
      if (a.cancelOn >= d - 1 && r.chance(0.6)) {
        const inv = invoice(a, a.cancelOn, 'Patient', 'Late cancellation', svc.price / 2);
        if (r.chance(0.55)) pay(inv, a.cancelOn + r.int(0, 20), inv.amount, 'Card');
      }
      return;
    }
    if (r.chance(0.035)) {
      a.status = 'No-show';
      if (r.chance(0.7)) {
        const inv = invoice(a, d, 'Patient', 'Missed appointment', svc.price / 2);
        if (r.chance(0.5)) pay(inv, d + r.int(1, 25), inv.amount, r.pick(['Card', 'Bank transfer']));
      }
      return;
    }
    a.status = 'Attended';

    // The bill: to the insurer when the policy covers it, otherwise to the patient.
    const ins = pat.ins;
    if (ins.real && svc.covered && r.chance(0.96)) {
      const inv = invoice(a, d, 'Insurer', 'Visit', svc.price);
      let at = ins.slow ? d + ins.days + r.int(-10, 35) : d + r.int(10, ins.days + 3);
      if (r.chance(0.03)) at += r.int(40, 90); // a claim the insurer queried
      pay(inv, at, inv.amount, 'Insurer', `REM-${r.int(100000, 999999)}`);
    } else {
      const inv = invoice(a, d, 'Patient', 'Visit', svc.price);
      const u = r.next();
      if (u < 0.78) {
        if (inv.amount >= 140 && r.chance(0.25)) {
          // Paid in two parts: half at the desk, the rest later.
          const half = money(inv.amount / 2);
          pay(inv, d, half, 'Card');
          pay(inv, d + r.int(7, 45), inv.amount - half, r.pick(['Card', 'Bank transfer']), null);
        } else pay(inv, d, inv.amount, r.weighted([['Card', 8], ['Cash', 1.5]]));
      } else if (u < 0.94) {
        const method = r.pick(['Card', 'Bank transfer']);
        pay(inv, d + r.int(1, 14), inv.amount, method, method === 'Bank transfer' ? `BT-${r.int(10000, 99999)}` : null);
      } else pay(inv, d + r.int(15, 75), inv.amount, 'Bank transfer', `BT-${r.int(10000, 99999)}`);
    }

    // Any open recall for this service is answered by this visit.
    const open = pat.recalls.get(svc.name);
    if (open && !open.bookedOn) open.bookedOn = a.bookedOn;

    // The note, and what comes next.
    let reason;
    let note;
    let plan = 'No follow-up needed';
    const p = a.p;
    if (svc.by === 'Doctor') {
      if (svc.name === 'Blood pressure review') {
        reason = 'Blood pressure review';
        note = r.pick(['Blood pressure recorded, within the target range.', 'Blood pressure a little high today, repeated after a rest and nearer target.', 'Readings from home look steady.']) + ' Talked about salt, exercise and sleep.';
        plan = 'Review again in six months';
        recall(pat, svc, d + 183, d);
      } else if (svc.name === 'Annual health check') {
        reason = 'Yearly health check';
        note = 'Height, weight, blood pressure and pulse recorded. ' + r.pick(['All in the normal range.', 'Weight a little up on last year; talked about diet.', 'All in the normal range. Asked to keep up the walking.']) + ' Lifestyle advice given.';
        plan = 'Recall in twelve months';
        recall(pat, svc, d + 365, d);
      } else {
        reason = svc.name === 'GP follow-up' ? `Review of ${a.reason ? a.reason.toLowerCase() : 'earlier visit'}` : a.reason || r.pick(GP_REASONS);
        note = svc.name === 'GP follow-up'
          ? r.pick(['Much better since the last visit.', 'Improving slowly.', 'Settled, no further concerns.']) + ' ' + r.pick(GP_ADVICE)
          : `${r.pick(GP_FOUND)} ${r.pick(GP_ADVICE)}`;
        if (svc.name !== 'GP follow-up') {
          const next = r.next();
          if (next < 0.12) { if (book(d, pat, S('GP follow-up'), { want: p, via: 'At the desk', extra: { reason } })) plan = 'Follow-up booked'; }
          else if (next < 0.2) { if (book(d, pat, S('Blood test'), { via: 'At the desk' })) plan = 'Blood test with the nurse'; }
          else if (next < 0.26) { if (book(d, pat, S('Physiotherapy assessment'), { via: 'At the desk' })) plan = 'Referred to physiotherapy'; }
          else plan = r.pick(['No follow-up needed', 'Return if not better in two weeks', 'Return if not better in two weeks']);
        }
      }
    } else if (svc.by === 'Nurse') {
      const N = {
        'Blood test': ['Blood test', 'Blood sample taken without difficulty and sent to the lab.', 'Results to the doctor'],
        'Wound care': ['Dressing for a small cut', r.pick(['Dressing changed, healing well.', 'Cleaned and redressed, healing as expected.']), 'Change dressing again if needed'],
        'Flu vaccination': ['Seasonal flu vaccination', 'Vaccination given in the upper arm. No reaction while waiting afterwards.', 'Due again next autumn'],
        'Travel vaccination': ['Travel vaccinations', 'Vaccinations given for a trip abroad. Advice on food, water and insect bites given.', 'No follow-up needed'],
        'Nurse health check': ['Nurse health check', 'Weight, blood pressure and pulse recorded, all normal. Talked about activity and diet.', 'Recall in twelve months'],
      }[svc.name];
      [reason, note, plan] = N;
      if (svc.recall) recall(pat, svc, d + Math.round(svc.recall * 30.44), d);
    } else {
      // Physiotherapy: an assessment starts a course; each session books the next a week on.
      if (svc.name === 'Physiotherapy assessment') {
        reason = r.pick(PHYSIO_REASONS);
        const course = r.int(2, 5);
        note = `Assessed movement and strength. Explained the likely cause. Gave ${r.pick(['three', 'four'])} home exercises.`;
        plan = `Course of ${course} sessions`;
        book(d, pat, S('Physiotherapy session'), { want: p, via: 'At the desk', extra: { reason, left: course - 1 } });
      } else if (svc.name === 'Physiotherapy session') {
        reason = a.reason;
        note = `${r.pick(['Good progress since last time.', 'Some progress since last time.', 'Slow progress, exercises adjusted.'])} Hands-on treatment and home exercises reviewed.`;
        if (a.left > 0) {
          plan = 'Next session in a week';
          book(d, pat, svc, { want: p, via: 'At the desk', extra: { reason, left: a.left - 1 } });
        } else {
          plan = 'Course finished';
          if (r.chance(0.4)) { plan = 'Course finished, review in three months'; recall(pat, S('Physiotherapy review'), d + 91, d); }
        }
      } else if (svc.name === 'Physiotherapy review') {
        reason = 'Review after a course of physiotherapy';
        note = r.pick(['Keeping up the exercises, no pain at rest.', 'Mostly better, the odd twinge after exercise.', 'Back to normal activity.']);
      } else {
        reason = 'Sports massage';
        note = `Sports massage to ${r.pick(['legs and lower back', 'shoulders and upper back', 'calves and hamstrings'])} after training.`;
      }
    }
    const end = a.s * 15 + svc.minutes + r.int(0, 10);
    notes.push({ a, written: cal.time(d, Math.floor(end / 60), end % 60), reason, note, plan });
  };

  // The clinic a day at a time.
  // The loop starts three weeks early so the diary already holds advance bookings on
  // the first day of the records; what happens in those weeks is not kept.
  for (let b = -days - 21; b <= 0; b++) {
    const w = wd(b);
    // Each Monday, letters for recalls due within two weeks; some patients book, some ring anyway.
    if (w === 1) {
      for (const x of recalls) {
        if (x.bookedOn || x.sentOn || x.due > b + 14 || x.due < b - 60 || x.made > b) continue;
        if (x.pat.reminders) {
          x.sentOn = b;
          if (r.chance(0.6)) respond(b + r.int(1, 10), x);
        } else if (!x.asked) {
          x.asked = true;
          if (r.chance(0.3)) respond(Math.max(b + 1, x.due - r.int(0, 10)), x);
        }
      }
    }
    for (const x of responses.get(b) || []) {
      if (x.bookedOn) continue;
      if (book(b, x.pat, x.svc, { via: 'Recall' })) x.bookedOn = b;
    }
    // Fresh bookings for each practitioner: fewer at the weekend, when only the website takes them.
    const m = monthOf(b);
    const dayFactor = w === 0 ? 0.4 : w === 6 ? 0.6 : 1;
    for (const p of team) {
      const n = Math.floor(p.rate * SEASON[p.role][m] * dayFactor * (0.75 + r.next() * 0.5) + r.next());
      for (let k = 0; k < n; k++) {
        const svc = S(r.weighted(MIX[p.role](m)));
        const pat = pickPatient(svc, b);
        book(b, pat, svc, { want: r.chance(0.75) ? p : null, via: w === 0 || w === 6 ? 'Online' : null });
      }
    }
    if (b < 0 && b >= -days) for (const a of byDay.get(b) || []) outcome(a);
  }

  // Into the file: appointments in diary order, then their notes, invoices and payments.
  const kept = appts.filter((a) => a.d >= -days).sort((x, y) => x.d - y.d || x.s - y.s || x.p.i - y.p.i);
  for (const a of kept) {
    a.id = add('Appointments', {
      Patient: a.pat.id, Practitioner: a.p.id, Service: a.svc.id, Start: cal.time(a.d, Math.floor(a.s / 4), (a.s % 4) * 15),
      Status: a.status, 'Booked On': cal.day(a.bookedOn), 'Booked Via': a.via, 'Cancelled On': a.cancelOn == null ? null : cal.day(a.cancelOn),
    });
  }
  notes.sort((x, y) => x.a.id - y.a.id);
  for (const n of notes) add('Visit Notes', { Appointment: n.a.id, Written: n.written, Reason: n.reason, Note: n.note, Plan: n.plan });
  // Invoices and receipts are numbered in date order, a new run each year: INV-2026-0042, RCP-2026-0042.
  const runs = {};
  const numbered = (prefix, d) => {
    const key = `${prefix}${cal.at(d).getFullYear()}`;
    runs[key] = (runs[key] || 0) + 1;
    return `${prefix}-${cal.at(d).getFullYear()}-${pad(runs[key], 4)}`;
  };
  invoices.sort((x, y) => x.d - y.d || x.a.id - y.a.id);
  const payments = [];
  for (const inv of invoices) {
    const paid = money(inv.payments.reduce((t, x) => t + x.amount, 0));
    const last = inv.payments.length ? Math.max(...inv.payments.map((x) => x.d)) : null;
    // Made out to the insurer, with the patient's policy reference, or to the patient at home.
    const pat = inv.a.pat;
    const ins = inv.billTo === 'Insurer' ? pat.ins : null;
    inv.id = add('Invoices', {
      Number: numbered('INV', inv.d), Appointment: inv.a.id, 'Invoice Date': cal.day(inv.d), 'Due Date': cal.day(inv.due), 'Bill To': inv.billTo,
      Insurer: ins ? ins.id : null, 'Bill To Name': ins ? ins.name : pat.name, 'Bill To Address': ins ? ins.street : pat.street,
      'Bill To City': ins ? ins.city : pat.city, 'Bill To Postcode': ins ? ins.postcode : pat.postcode, 'Policy Reference': ins ? pat.policy : null,
      For: inv.kind, Amount: inv.amount, 'Amount Paid': paid, 'Paid On': paid >= inv.amount && last != null ? cal.day(last) : null,
    });
    for (const x of inv.payments) payments.push({ ...x, inv });
  }
  payments.sort((x, y) => x.d - y.d || x.inv.id - y.inv.id);
  for (const x of payments) add('Payments', { Number: numbered('RCP', x.d), Invoice: x.inv.id, 'Paid On': cal.day(x.d), Amount: x.amount, Method: x.method, Reference: x.reference });
  recalls.sort((x, y) => x.made - y.made || x.due - y.due);
  for (const x of recalls) {
    add('Recalls', {
      Patient: x.pat.id, Service: x.svc.id, 'Due Date': cal.day(x.due), Reason: RECALL_REASON[x.svc.name],
      'Sent On': x.sentOn == null ? null : cal.day(x.sentOn), 'Booked On': x.bookedOn == null ? null : cal.day(x.bookedOn),
    });
  }
  return counts;
}
