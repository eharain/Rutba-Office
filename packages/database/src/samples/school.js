// The learning centre: Lantern Learning Centre, an invented tuition and
// training centre with about a hundred and twenty students. After school and
// on Saturday mornings it teaches children maths, English, science and
// coding in small groups; in the evenings it runs courses for adults in
// languages and bookkeeping; in the school holidays it runs daytime
// workshops. Families pay the fees, one invoice a term; an adult learner is
// their own family.
//
// Its records come from fourteen to eighteen months of terms and holidays, run one
// block of weeks at a time and a lesson at a time inside it: children move
// up a school year each September and leave after Year 11, some families
// stay for years and some for a term, classes are opened to fit how many
// want a course, a register is taken at every lesson that has happened, tests
// are marked twice a term, invoices go out three weeks before a term and are
// paid on time, in instalments, late or not yet. The lessons of the current
// and the next block are already planned. Children's records are kept to what
// a centre needs: a name, a date of birth, a school year, a short medical
// note where there is one, and whether photographs may be taken.
//
// What the centre prints comes from its own records: a fee invoice to each
// family each term with a line for each child's class, a receipt for each
// payment, a letter confirming each place, a progress report for each
// student at the end of each term, a certificate for each adult course and
// holiday workshop finished, and a pay statement for each teacher each
// month. Tuition is exempt from VAT, so the fees carry none.

import { BUSINESS_ABOUT, BUSINESS_TABLE, businessRecord, FIRST, LAST, money, person, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const FEE = 'Round([Fee per Lesson] * [Lessons Billed] * (1 - Nz([Discount])), 2)';
const STUDENT = '[Students].[First Name] & " " & [Students].[Last Name]';
const TEACHER = '[Teachers].[First Name] & " " & [Teachers].[Last Name]';
const OWED = 'Round([Amount] - Nz([Paid So Far]), 2)';
// Every document is headed with the centre's own details.
const LETTERHEAD = { table: 'Business', fields: ['Name', 'Address', 'City', 'Postcode', 'Phone', 'E-mail', 'Website'] };
// Days from today to a birthday, counting each month as 31 days: near enough
// to find the birthdays in the coming month without a date built from parts.
const TO_BIRTHDAY = '(Month([Date of Birth]) * 31 + Day([Date of Birth]) - Month(Date()) * 31 - Day(Date()) + 372) Mod 372';

export const spec = {
  negative: {},
  tables: [
    BUSINESS_TABLE,
    { name: 'Terms', fields: [
      { name: 'Name', type: 'text', required: true, unique: true }, { name: 'Kind', type: 'text', default: 'Term' },
      { name: 'Starts', type: 'date', required: true }, { name: 'Ends', type: 'date', required: true },
    ] },
    { name: 'Families', fields: [
      { name: 'Name', type: 'text', required: true }, { name: 'Contact', type: 'text' }, { name: 'Relationship', type: 'text' },
      { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' }, { name: 'Second Contact', type: 'text' }, { name: 'Second Phone', type: 'text' },
      { name: 'Address', type: 'text' }, { name: 'City', type: 'text' }, { name: 'Postcode', type: 'text' },
      { name: 'Pays By', type: 'text', default: 'Bank transfer' }, { name: 'Notes', type: 'longtext' },
    ] },
    { name: 'Students', fields: [
      { name: 'Family', ...ref('Families'), required: true }, { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true },
      { name: 'Date of Birth', type: 'date' }, { name: 'School Year', type: 'text' }, { name: 'Medical Note', type: 'text' },
      { name: 'Photo Consent', type: 'yesno' }, { name: 'Joined', type: 'date' }, { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Teachers', fields: [
      { name: 'First Name', type: 'text' }, { name: 'Last Name', type: 'text', required: true }, { name: 'Subjects', type: 'text' },
      { name: 'Hourly Rate', type: 'currency' }, { name: 'Phone', type: 'text' }, { name: 'E-mail', type: 'text' },
      { name: 'Started', type: 'date' }, { name: 'Background Check', type: 'date' }, { name: 'Active', type: 'yesno', default: true },
    ] },
    { name: 'Courses', fields: [
      { name: 'Code', type: 'text', unique: true }, { name: 'Name', type: 'text', required: true }, { name: 'Subject', type: 'text' },
      { name: 'Level', type: 'text' }, { name: 'For', type: 'text', default: 'Children' }, { name: 'Holiday Course', type: 'yesno' },
      { name: 'Minutes', type: 'integer', default: 60 }, { name: 'Fee per Lesson', type: 'currency' }, { name: 'Capacity', type: 'integer', default: 10 },
      { name: 'What to Bring', type: 'text' },
    ] },
    { name: 'Classes', fields: [
      { name: 'Code', type: 'text', required: true }, { name: 'Course', ...ref('Courses'), required: true }, { name: 'Teacher', ...ref('Teachers') },
      { name: 'Term', ...ref('Terms'), required: true }, { name: 'Weekday', type: 'text' }, { name: 'Start Time', type: 'text' },
      { name: 'Room', type: 'text' }, { name: 'Capacity', type: 'integer', default: 10 }, { name: 'Lessons Planned', type: 'integer' },
    ] },
    { name: 'Enrolments', fields: [
      { name: 'Student', ...ref('Students'), required: true }, { name: 'Class', ...ref('Classes'), required: true },
      { name: 'Enrolled On', type: 'date' }, { name: 'Status', type: 'text', default: 'Enrolled' },
      { name: 'Lessons Billed', type: 'integer' }, { name: 'Discount', type: 'number', default: 0 }, { name: 'Left On', type: 'date' },
      // The first lesson the student comes to, and the term's invoice the place is billed on: a line of it.
      { name: 'First Lesson', type: 'date' }, { name: 'Invoice', ...ref('Fee Invoices') },
    ] },
    { name: 'Lessons', fields: [
      { name: 'Class', ...ref('Classes'), required: true }, { name: 'Teacher', ...ref('Teachers') }, { name: 'Lesson Date', type: 'date', required: true },
      { name: 'Minutes', type: 'integer' }, { name: 'Topic', type: 'text' }, { name: 'Status', type: 'text', default: 'Planned' },
      // The month's pay statement a taught lesson was paid on: a line of it.
      { name: 'Pay Statement', ...ref('Pay Statements') },
    ] },
    { name: 'Attendance', fields: [
      { name: 'Lesson', ...ref('Lessons'), required: true }, { name: 'Student', ...ref('Students'), required: true },
      { name: 'Mark', type: 'text', default: 'Present' }, { name: 'Reason', type: 'text' },
    ] },
    { name: 'Assessments', fields: [
      { name: 'Class', ...ref('Classes'), required: true }, { name: 'Assessment Date', type: 'date' }, { name: 'Title', type: 'text' },
      { name: 'Out Of', type: 'integer' },
    ] },
    { name: 'Results', fields: [
      { name: 'Assessment', ...ref('Assessments'), required: true }, { name: 'Student', ...ref('Students'), required: true },
      { name: 'Score', type: 'integer' }, { name: 'Comment', type: 'text' },
    ] },
    { name: 'Fee Invoices', fields: [
      { name: 'Number', type: 'text', unique: true },
      { name: 'Family', ...ref('Families'), required: true }, { name: 'Term', ...ref('Terms'), required: true },
      { name: 'Invoice Date', type: 'date' }, { name: 'Due Date', type: 'date' }, { name: 'Amount', type: 'currency' }, { name: 'Paid So Far', type: 'currency', default: 0 },
    ] },
    { name: 'Payments', fields: [
      { name: 'Number', type: 'text', unique: true },
      { name: 'Invoice', ...ref('Fee Invoices'), required: true }, { name: 'Paid On', type: 'date' }, { name: 'Amount', type: 'currency' },
      { name: 'Method', type: 'text' }, { name: 'Balance After', type: 'currency' },
    ] },
    // The report sent home at the end of each term: one for each student, a line for each class.
    { name: 'Term Reports', fields: [
      { name: 'Student', ...ref('Students'), required: true }, { name: 'Term', ...ref('Terms'), required: true },
      { name: 'School Year', type: 'text' }, { name: 'Written On', type: 'date' }, { name: 'Comment', type: 'longtext' },
    ] },
    { name: 'Report Lines', fields: [
      { name: 'Report', ...ref('Term Reports'), required: true }, { name: 'Class', ...ref('Classes'), required: true },
      { name: 'Lessons', type: 'integer' }, { name: 'Present', type: 'integer' }, { name: 'Late', type: 'integer' }, { name: 'Absent', type: 'integer' },
      { name: 'Half-term %', type: 'integer' }, { name: 'End of Term %', type: 'integer' }, { name: 'Comment', type: 'text' },
    ] },
    { name: 'Certificates', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Enrolment', ...ref('Enrolments'), required: true },
      { name: 'Issued On', type: 'date' }, { name: 'Lessons Attended', type: 'integer' }, { name: 'Hours', type: 'number' },
    ] },
    // A teacher's pay for a month: the lessons taught are its lines.
    { name: 'Pay Statements', fields: [
      { name: 'Number', type: 'text', unique: true }, { name: 'Teacher', ...ref('Teachers'), required: true }, { name: 'Month', type: 'text' },
      { name: 'Period Start', type: 'date' }, { name: 'Period End', type: 'date' }, { name: 'Paid On', type: 'date' },
      { name: 'Hourly Rate', type: 'currency' }, { name: 'Lessons', type: 'integer' }, { name: 'Hours', type: 'number' }, { name: 'Gross', type: 'currency' },
    ] },
  ],
  queries: {
    'Class lists': { tables: ['Enrolments', 'Students', 'Classes', 'Courses', 'Terms', 'Teachers', 'Families'], columns: [
      { table: 'Terms', field: 'Starts', sort: 'asc', show: false },
      { expr: '[Classes].[Code] & ", " & [Terms].[Name]', as: 'Class', sort: 'asc' }, { table: 'Courses', field: 'Name', as: 'Course' },
      { table: 'Classes', field: 'Weekday' }, { table: 'Classes', field: 'Start Time' }, { table: 'Classes', field: 'Room' }, { expr: TEACHER, as: 'Teacher' },
      { table: 'Students', field: 'Last Name', sort: 'asc', show: false }, { expr: STUDENT, as: 'Student' }, { table: 'Students', field: 'School Year' },
      { table: 'Students', field: 'Medical Note' }, { table: 'Students', field: 'Photo Consent' }, { table: 'Families', field: 'Contact' }, { table: 'Families', field: 'Phone' },
      { table: 'Enrolments', field: 'Status', criteria: ['Enrolled'], show: false },
      { expr: 'DateDiff("d", Date(), [Ends])', as: 'Days to End', criteria: ['>= 0'], show: false },
    ] },
    'Timetable, next 7 days': { tables: ['Lessons', 'Classes', 'Courses', 'Teachers'], columns: [
      { table: 'Lessons', field: 'Lesson Date', sort: 'asc' }, { table: 'Classes', field: 'Weekday' }, { table: 'Classes', field: 'Start Time', sort: 'asc' },
      { table: 'Courses', field: 'Name', as: 'Course' }, { table: 'Classes', field: 'Code', as: 'Class' }, { table: 'Classes', field: 'Room' },
      { expr: TEACHER, as: 'Teacher' }, { table: 'Lessons', field: 'Minutes' }, { table: 'Lessons', field: 'Topic' },
      { expr: 'DateDiff("d", Date(), [Lesson Date])', as: 'Days Ahead', criteria: ['Between 0 And 7'], show: false },
      { table: 'Lessons', field: 'Status', criteria: ['<> "Cancelled"'], show: false },
    ] },
    'Attendance below 85 percent': { tables: ['Attendance', 'Students', 'Lessons', 'Families'], columns: [
      { table: 'Students', field: 'ID', total: 'group', show: false }, { expr: STUDENT, as: 'Student', total: 'group' },
      { table: 'Students', field: 'School Year', total: 'group' }, { table: 'Families', field: 'Contact', total: 'group' }, { table: 'Families', field: 'Phone', total: 'group' },
      { table: 'Attendance', field: 'ID', as: 'Lessons', total: 'count', criteria: ['>= 4'] },
      { expr: 'IIf([Mark] = "Absent", 1, 0)', as: 'Absences', total: 'sum' },
      { expr: 'IIf([Mark] = "Absent", 0, 100)', as: 'Attendance %', total: 'avg', criteria: ['< 85'], sort: 'asc' },
      { expr: 'DateDiff("d", [Lesson Date], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 1 And 90'] },
    ] },
    'Fees outstanding by family': { tables: ['Fee Invoices', 'Families'], columns: [
      { table: 'Families', field: 'ID', total: 'group', show: false }, { table: 'Families', field: 'Name', as: 'Family', total: 'group' },
      { table: 'Families', field: 'Contact', total: 'group' }, { table: 'Families', field: 'Phone', total: 'group' },
      { table: 'Fee Invoices', field: 'ID', as: 'Invoices', total: 'count' }, { table: 'Fee Invoices', field: 'Due Date', as: 'Oldest Due', total: 'min' },
      { expr: OWED, as: 'Outstanding', total: 'sum', sort: 'desc' }, { expr: OWED, as: 'Balance', total: 'where', criteria: ['> 0'] },
    ] },
    'Overdue fees': { tables: ['Fee Invoices', 'Families', 'Terms'], columns: [
      { table: 'Families', field: 'Name', as: 'Family', sort: 'asc' }, { table: 'Families', field: 'Contact' }, { table: 'Families', field: 'Address' },
      { table: 'Families', field: 'City' }, { table: 'Families', field: 'Postcode' }, { table: 'Fee Invoices', field: 'Number', as: 'Invoice' },
      { table: 'Terms', field: 'Name', as: 'Term' }, { table: 'Fee Invoices', field: 'Due Date', sort: 'asc' }, { table: 'Fee Invoices', field: 'Amount' },
      { expr: 'DateDiff("d", [Due Date], Date())', as: 'Days Overdue', criteria: ['> 0'] }, { expr: OWED, as: 'Outstanding', criteria: ['> 0'] },
    ] },
    'Classes with free places': { tables: ['Classes', 'Courses', 'Terms', 'Teachers', 'Enrolments'], columns: [
      { table: 'Terms', field: 'Starts', total: 'group', sort: 'asc', show: false }, { table: 'Classes', field: 'ID', total: 'group', show: false },
      { table: 'Classes', field: 'Code', as: 'Class', total: 'group', sort: 'asc' }, { table: 'Terms', field: 'Name', as: 'Term', total: 'group' },
      { table: 'Courses', field: 'Name', as: 'Course', total: 'group' }, { table: 'Classes', field: 'Weekday', total: 'group' },
      { table: 'Classes', field: 'Start Time', total: 'group' }, { expr: TEACHER, as: 'Teacher', total: 'group' },
      { table: 'Classes', field: 'Capacity', as: 'Places', total: 'max' }, { table: 'Enrolments', field: 'ID', as: 'Enrolled', total: 'count' },
      // Each student enrolled fills 100 / Capacity percent of the class; the sum is how full it is.
      { expr: '100 / [Classes].[Capacity]', as: 'Percent Full', total: 'sum', criteria: ['< 100'] },
      { table: 'Enrolments', field: 'Status', total: 'where', criteria: ['Enrolled'] },
      { expr: 'DateDiff("d", Date(), [Ends])', as: 'Days to End', total: 'where', criteria: ['>= 0'] },
    ] },
    'Results by course': { tables: ['Results', 'Assessments', 'Classes', 'Courses'], columns: [
      { table: 'Courses', field: 'Name', as: 'Course', total: 'group', sort: 'asc' }, { table: 'Courses', field: 'Subject', total: 'group' },
      { table: 'Results', field: 'ID', as: 'Results', total: 'count' },
      { expr: 'Round([Score] * 100 / [Out Of], 0)', as: 'Average %', total: 'avg' },
      { expr: 'Round([Score] * 100 / [Out Of], 0)', as: 'Highest %', total: 'max' },
      { expr: 'Round([Score] * 100 / [Out Of], 0)', as: 'Lowest %', total: 'min' },
      { expr: 'DateDiff("d", [Assessment Date], Date())', as: 'Days Ago', total: 'where', criteria: ['Between 0 And 365'] },
    ] },
    'Progress by student': { tables: ['Results', 'Students', 'Assessments', 'Classes', 'Courses'], columns: [
      { table: 'Students', field: 'Last Name', sort: 'asc', show: false }, { table: 'Students', field: 'First Name', sort: 'asc', show: false },
      { expr: `${STUDENT} & " (" & [School Year] & ")"`, as: 'Student' }, { table: 'Courses', field: 'Name', as: 'Course' },
      { table: 'Assessments', field: 'Title', as: 'Assessment' }, { table: 'Assessments', field: 'Assessment Date', sort: 'asc' },
      { table: 'Results', field: 'Score' }, { table: 'Assessments', field: 'Out Of' }, { expr: 'Round([Score] * 100 / [Out Of], 0)', as: 'Percent' },
      { table: 'Results', field: 'Comment' },
      { expr: 'DateDiff("d", [Assessment Date], Date())', as: 'Days Ago', criteria: ['Between 0 And 365'], show: false },
    ] },
    'Enrolments by course and term': { tables: ['Enrolments', 'Classes', 'Courses', 'Terms'], columns: [
      { table: 'Terms', field: 'Starts', total: 'group', sort: 'asc' }, { table: 'Terms', field: 'Name', as: 'Term', total: 'group' },
      { table: 'Courses', field: 'Name', as: 'Course', total: 'group', sort: 'asc' }, { table: 'Enrolments', field: 'ID', as: 'Enrolments', total: 'count' },
      { expr: FEE, as: 'Fees', total: 'sum' },
    ] },
    'Teacher hours last month': { tables: ['Lessons', 'Teachers'], columns: [
      { expr: 'Format([Lesson Date], "yyyy-mm")', as: 'Month', total: 'group' }, { table: 'Teachers', field: 'ID', total: 'group', show: false },
      { expr: TEACHER, as: 'Teacher', total: 'group', sort: 'asc' }, { table: 'Teachers', field: 'Hourly Rate', total: 'group' },
      { table: 'Lessons', field: 'ID', as: 'Lessons', total: 'count' }, { expr: '[Minutes] / 60', as: 'Hours', total: 'sum' },
      { expr: 'Round([Minutes] / 60 * [Hourly Rate], 2)', as: 'Pay', total: 'sum' },
      { table: 'Lessons', field: 'Status', total: 'where', criteria: ['Taught'] },
      { expr: 'DateDiff("m", [Lesson Date], Date())', as: 'Months Ago', total: 'where', criteria: ['1'] },
    ] },
    'Fees received by month': { tables: ['Payments'], columns: [
      { expr: 'Format([Paid On], "yyyy-mm")', as: 'Month', total: 'group', sort: 'asc' }, { table: 'Payments', field: 'ID', as: 'Payments', total: 'count' },
      { table: 'Payments', field: 'Amount', as: 'Received', total: 'sum' },
    ] },
    'Birthdays coming up': { tables: ['Students'], columns: [
      { expr: TO_BIRTHDAY, as: 'Days Away', criteria: ['Between 0 And 30'], sort: 'asc', show: false },
      { expr: '[First Name] & " " & [Last Name]', as: 'Student' }, { table: 'Students', field: 'School Year' },
      { expr: 'Format([Date of Birth], "dd") & "/" & Format([Date of Birth], "mm")', as: 'Birthday' },
      { expr: 'Year(Date()) - Year([Date of Birth]) + IIf(Month([Date of Birth]) * 31 + Day([Date of Birth]) < Month(Date()) * 31 + Day(Date()), 1, 0)', as: 'Turns' },
      { table: 'Students', field: 'Active', criteria: ['Yes'], show: false },
    ] },
    'Background checks due': { tables: ['Teachers'], columns: [
      { expr: '[First Name] & " " & [Last Name]', as: 'Teacher' }, { table: 'Teachers', field: 'Subjects' },
      { table: 'Teachers', field: 'Background Check', sort: 'asc' }, { expr: 'DateDiff("m", [Background Check], Date())', as: 'Months Old', criteria: ['>= 33'] },
      { table: 'Teachers', field: 'Active', criteria: ['Yes'], show: false },
    ] },
  },
  reports: [
    { name: 'Class register', query: 'Class lists', group: 'Class', columns: ['Student', 'School Year', 'Medical Note', 'Photo Consent', 'Contact', 'Phone'], totals: [{ field: 'Student', fn: 'count' }] },
    { name: 'Timetable', title: 'Timetable for the next 7 days', query: 'Timetable, next 7 days', group: 'Lesson Date', sort: 'Start Time', columns: ['Start Time', 'Course', 'Class', 'Room', 'Teacher', { field: 'Minutes', type: 'integer' }], totals: [{ field: 'Class', fn: 'count' }] },
    { name: 'Fees due letters', title: 'Fees overdue', query: 'Overdue fees', group: 'Family', columns: ['Contact', 'Address', 'City', 'Postcode', 'Invoice', 'Term', { field: 'Due Date', type: 'date' }, { field: 'Days Overdue', type: 'integer' }, { field: 'Outstanding', type: 'currency' }], totals: [{ field: 'Outstanding', fn: 'sum' }] },
    { name: 'Progress report', query: 'Progress by student', group: 'Student', sort: 'Assessment Date', columns: ['Course', 'Assessment', { field: 'Assessment Date', type: 'date' }, { field: 'Score', type: 'integer' }, { field: 'Out Of', type: 'integer' }, { field: 'Percent', type: 'integer' }, 'Comment'], totals: [{ field: 'Percent', fn: 'avg' }] },
    { name: 'Teacher pay', query: 'Teacher hours last month', columns: ['Month', 'Teacher', { field: 'Hourly Rate', type: 'currency' }, { field: 'Lessons', type: 'integer' }, { field: 'Hours', type: 'number' }, { field: 'Pay', type: 'currency' }], totals: [{ field: 'Hours', fn: 'sum' }, { field: 'Pay', fn: 'sum' }] },
    { name: 'Attendance concerns', query: 'Attendance below 85 percent', columns: ['Student', 'School Year', 'Contact', 'Phone', { field: 'Lessons', type: 'integer' }, { field: 'Absences', type: 'integer' }, { field: 'Attendance %', type: 'number' }] },
    { name: 'Enrolments by course', query: 'Enrolments by course and term', group: 'Course', sort: 'Starts', columns: ['Term', { field: 'Starts', type: 'date' }, { field: 'Enrolments', type: 'integer' }, { field: 'Fees', type: 'currency' }], totals: [{ field: 'Enrolments', fn: 'sum' }, { field: 'Fees', fn: 'sum' }] },
    { name: 'Fees received', query: 'Fees received by month', columns: ['Month', { field: 'Payments', type: 'integer' }, { field: 'Received', type: 'currency' }], totals: [{ field: 'Payments', fn: 'sum' }, { field: 'Received', fn: 'sum' }] },
  ],
  // The print templates: one document for each record of the table named.
  documents: [
    { name: 'Fee invoice', table: 'Fee Invoices', paper: 'a4', letterhead: LETTERHEAD,
      blocks: [
        { type: 'heading', text: 'Fee invoice {Number}' },
        { type: 'fields', columns: 2, items: [
          { field: 'Family.Contact', label: 'To' }, { field: 'Number', label: 'Invoice' },
          { field: 'Family.Address', label: 'Address' }, { field: 'Invoice Date' },
          { field: 'Family.City', label: 'City' }, { field: 'Due Date', label: 'Please pay by' },
          { field: 'Family.Postcode', label: 'Postcode' }, { field: 'Term', label: 'For' },
          { field: 'Family', label: 'Account' },
        ] },
        { type: 'lines', table: 'Enrolments', sort: 'Student', columns: [
          { field: 'Student' }, { field: 'Class.Course.Name', as: 'Course' }, { field: 'Class.Weekday', as: 'Day' }, { field: 'Class.Start Time', as: 'Time' },
          { field: 'Lessons Billed', as: 'Lessons', type: 'integer' }, { field: 'Class.Course.Fee per Lesson', as: 'Per Lesson', type: 'currency' },
          { expr: 'Round([Per Lesson] * [Lessons], 2)', as: 'Fees', type: 'currency' },
          { expr: 'Round([Fees] - Round([Per Lesson] * [Lessons] * (1 - Nz([Discount])), 2), 2)', as: 'Less', label: 'Discount', type: 'currency' },
          { expr: 'Round([Per Lesson] * [Lessons] * (1 - Nz([Discount])), 2)', as: 'To Pay', type: 'currency' },
        ] },
        { type: 'totals', rows: [
          { label: 'Fees', expr: '[Fees]' }, { label: 'Less discounts', expr: '[Less]' },
          { label: 'Total for the term', expr: '[To Pay]', strong: true },
          { label: 'Paid so far', expr: 'Nz([Paid So Far])' },
          { label: 'Balance due', expr: 'Round([Total for the term] - [Paid so far], 2)', strong: true },
        ] },
        { type: 'text', text: 'Brothers and sisters after the first child booked in a term have ten percent off; a bursary place is half fees. {Business.Payment Terms}', small: true },
        { type: 'text', text: 'Pay by bank transfer to {Business.Bank Details}, quoting {Number}, or by card or cash at the office.', small: true },
        { type: 'text', text: '{Business.Document Footer}', small: true },
      ] },
    { name: 'Payment receipt', table: 'Payments', paper: 'a5', letterhead: LETTERHEAD,
      blocks: [
        { type: 'heading', text: 'Receipt {Number}' },
        { type: 'fields', columns: 2, items: [
          { field: 'Number', label: 'Receipt' }, { field: 'Paid On', label: 'Received on' },
          { field: 'Invoice.Family.Contact', label: 'Received from' }, { field: 'Method', label: 'Paid by' },
          { field: 'Invoice', label: 'For invoice' }, { field: 'Invoice.Term', label: 'Term' },
        ] },
        { type: 'totals', rows: [
          { label: 'Invoice total', expr: 'Round([Amount] + [Balance After], 2)', type: 'currency' },
          { label: 'Received with thanks', expr: '[Amount]', strong: true },
          { label: 'Left to pay on the invoice', expr: '[Balance After]' },
        ] },
        { type: 'text', text: 'The invoice total here is what was owing on the invoice before this payment.', small: true },
        { type: 'text', text: '{Business.Document Footer}', small: true },
        { type: 'signatures', labels: ['Received for the centre'] },
      ] },
    { name: 'Enrolment letter', table: 'Enrolments', paper: 'a4', letterhead: LETTERHEAD,
      blocks: [
        { type: 'fields', columns: 2, items: [
          { field: 'Student.Family.Contact', label: 'To' }, { field: 'Enrolled On', label: 'Date' },
          { field: 'Student.Family.Address', label: 'Address' }, { field: 'Student.Family.City', label: 'City' },
          { field: 'Student.Family.Postcode', label: 'Postcode' }, { field: 'Class', label: 'Our reference' },
        ] },
        { type: 'heading', text: 'Your place is confirmed: {Student.First Name} on {Class.Course.Name}' },
        { type: 'text', text: 'Dear {Student.Family.Contact},' },
        { type: 'text', text: 'Thank you for booking with us. This letter confirms the place for {Student} and tells you what to expect before the first lesson.' },
        { type: 'text', text: '1. The class. {Class.Course.Name} meets every {Class.Weekday} at {Class.Start Time} in {Class.Room}, and each lesson lasts {Class.Course.Minutes} minutes. The teacher is {Class.Teacher}.' },
        { type: 'text', text: '2. When it starts. The first lesson for {Student.First Name} is on {First Lesson}. {Class.Term} runs from {Class.Term.Starts} to {Class.Term.Ends}, and the class has {Class.Lessons Planned} lessons in that time.' },
        { type: 'text', text: '3. What to bring. {Class.Course.What to Bring}' },
        { type: 'text', text: '4. Fees. The place is charged for {Lessons Billed} lessons, on the invoice for the term. Fees are not refunded part way through a term, but if a lesson is cancelled by us it is made up or taken off the next invoice.' },
        { type: 'text', text: '5. Arriving and going home. Please arrive five minutes before the start. Children are signed in at the front desk and wait inside until they are collected, so please tell the office if someone new is coming for them. Adult learners go straight to the room.' },
        { type: 'text', text: '6. If you cannot come. Ring {Business.Phone} or write to {Business.E-mail} before the lesson, so the register can be marked and work kept back.' },
        { type: 'text', text: 'We look forward to seeing {Student.First Name}.' },
        { type: 'text', text: 'Yours sincerely,' },
        { type: 'spacer', size: 24 },
        { type: 'text', text: 'The office, {Business.Name}' },
      ] },
    { name: 'Term report', table: 'Term Reports', paper: 'a4', letterhead: LETTERHEAD,
      blocks: [
        { type: 'heading', text: 'Progress report, {Term}' },
        { type: 'fields', columns: 2, items: [
          { field: 'Student' }, { field: 'School Year' }, { field: 'Term' }, { field: 'Written On', label: 'Date' },
          { field: 'Student.Family.Contact', label: 'Sent to' },
        ] },
        { type: 'lines', table: 'Report Lines', columns: [
          { field: 'Class.Course.Name', as: 'Course' }, { field: 'Class.Teacher', as: 'Teacher' },
          { field: 'Lessons', type: 'integer' }, { field: 'Late', type: 'integer' }, { field: 'Absent', type: 'integer' },
          { expr: 'Round(([Lessons] - [Absent]) * 100 / [Lessons], 0)', as: 'Attended %', type: 'integer' },
          { field: 'Half-term %', type: 'integer' }, { field: 'End of Term %', type: 'integer' }, { field: 'Comment', as: 'Teacher\'s comment' },
        ] },
        { type: 'totals', rows: [
          { label: 'Lessons this term', expr: '[Lessons]', type: 'integer' },
          { label: 'Missed', expr: '[Absent]', type: 'integer' },
          { label: 'Attendance, percent', expr: 'Round(([Lessons] - [Absent]) * 100 / [Lessons], 0)', type: 'integer', strong: true },
        ] },
        { type: 'text', text: '{Comment}' },
        { type: 'text', text: 'Marks are percentages of the test. A blank mark means the test was missed or the course has none.', small: true },
        { type: 'signatures', labels: ['Centre director'] },
      ] },
    { name: 'Certificate of completion', table: 'Certificates', paper: 'a4', letterhead: LETTERHEAD,
      blocks: [
        { type: 'spacer', size: 32 },
        { type: 'heading', text: 'Certificate of completion' },
        { type: 'text', text: 'This is to certify that' },
        { type: 'heading', text: '{Enrolment.Student}' },
        { type: 'text', text: 'has completed {Enrolment.Class.Course.Name} ({Enrolment.Class.Course.Level}) at {Business.Name}, {Enrolment.Class.Term}, taught by {Enrolment.Class.Teacher}, attending {Lessons Attended} lessons and {Hours} hours of teaching.' },
        { type: 'spacer', size: 24 },
        { type: 'fields', columns: 2, items: [{ field: 'Number', label: 'Certificate' }, { field: 'Issued On' }] },
        { type: 'spacer', size: 24 },
        { type: 'signatures', labels: ['Teacher', 'Centre director'] },
      ] },
    { name: 'Pay statement', table: 'Pay Statements', paper: 'a4', letterhead: LETTERHEAD,
      blocks: [
        { type: 'heading', text: 'Pay statement {Number}' },
        { type: 'fields', columns: 2, items: [
          { field: 'Teacher' }, { field: 'Number', label: 'Statement' },
          { field: 'Month' }, { field: 'Paid On' },
          { field: 'Period Start', label: 'From' }, { field: 'Period End', label: 'To' },
          { field: 'Hourly Rate' },
        ] },
        { type: 'lines', table: 'Lessons', sort: 'Lesson Date', columns: [
          { field: 'Lesson Date', as: 'Date' }, { field: 'Class', as: 'Class' }, { field: 'Class.Course.Name', as: 'Course' },
          { field: 'Minutes', type: 'integer' }, { expr: 'Round([Minutes] / 60, 2)', as: 'Hours', type: 'number' },
          { field: 'Teacher.Hourly Rate', as: 'Rate', type: 'currency' },
          { expr: 'Round([Minutes] / 60 * [Rate], 2)', as: 'Pay', type: 'currency' },
        ] },
        { type: 'totals', rows: [
          { label: 'Lessons taught', expr: '[Lessons]', type: 'integer' },
          { label: 'Hours', expr: '[Hours]', type: 'number' },
          { label: 'Gross pay', expr: '[Pay]', strong: true },
        ] },
        { type: 'text', text: 'Pay is for lessons taught, at your hourly rate; cancelled lessons are not paid. Tax and National Insurance are taken by the payroll and shown on the payslip it sends.', small: true },
      ] },
  ],
};

export const about = {
  name: 'Lantern Learning Centre',
  kind: 'School or training centre',
  line: 'Students, families, classes, registers, test results and termly fees for a tuition or training centre.',
  summary: [
    'For a centre that teaches groups in terms and bills by the term: a tuition business, a language or music school, a dance studio, a sports academy. It keeps who comes, who pays, which class each student is in, who was there at each lesson, how they did in their tests and what is still owed.',
    'Lantern Learning Centre is an invented tuition centre in Leeds with about a hundred and twenty students. After school and on Saturday mornings it teaches children from Year 3 to Year 11 (ages seven to sixteen) in groups of eight to ten: maths, English and science for each stage up to the GCSE exams, and coding. On weekday evenings adults come for Spanish, French and Mandarin, and for bookkeeping in the computer room. In the school holidays it runs daytime workshops: a maths booster, GCSE revision days, a coding club, a reading and writing club and science lab days. Ten teachers are paid by the hour.',
    'The records go back to the start of a term fourteen to eighteen months before the day the file is made: three terms a year with the holidays between, run to the end of the current one, with the next term or holiday planned when it starts within six weeks. Each class has its lessons on one weekday at one time, a register for every lesson that has happened, and tests twice a term for the children. Families get one invoice a term, three weeks before it starts, with ten percent off for a second or third child; most pay before the first lesson, some pay in three instalments, some pay late and a few still owe from earlier terms.',
    'It prints what such a centre sends out: a fee invoice to each family each term, with a line for each child\'s class, the discounts and what has been paid; a receipt for every payment; a letter confirming each place, with the day, time, room, first lesson and what to bring; a progress report for each student at the end of each term, with attendance and marks for each class; a certificate for each adult course and holiday workshop finished; and a pay statement for each teacher each month. Tuition is exempt from VAT here, so no VAT is charged.',
    'Children\'s records hold only what a centre needs: a name, the date of birth, the school year, a short note such as an allergy or an inhaler, and whether photographs may be taken. Addresses and phone numbers belong to the family, not the child. All of it is invented.',
  ].join('\n\n'),
  suits: [
    'A tutoring business with group or one-to-one lessons',
    'A language school with evening classes',
    'A music or dance school that teaches in terms',
    'A driving school, with a class for each instructor\'s block of lessons',
    'A sports academy or swimming club with weekly sessions',
    'A small private or weekend school',
    'An adult education or evening class programme',
  ],
  tables: {
    Business: BUSINESS_ABOUT,
    'Terms': { about: 'The blocks of weeks the centre teaches in: three terms a year and the school holidays between them, when holiday workshops run. Every class belongs to one, and fees are invoiced by it. Enter these first.', fields: {
      'Name': 'The name families know the block by, such as Autumn 2026 or Easter holiday 2026.',
      'Kind': 'Term for a teaching term, Holiday for the holiday workshops between terms.',
      'Starts': 'The first day of the block. Invoices go out three weeks before it.',
      'Ends': 'The last day of the block. A class whose term has not ended is current or coming, and the class lists and free places look only at those.',
    } },
    'Families': { about: 'One record for each household that pays: the parent or carer who is the centre\'s contact, and where to send letters. Brothers and sisters share one family. An adult learner is their own family, with Relationship set to Self.', fields: {
      'Name': 'The name the account goes by, such as Okafor family, or the adult learner\'s own name.',
      'Contact': 'The person the centre talks to and sends invoices to.',
      'Relationship': 'What the contact is to the students: Parent, Grandparent, Carer, or Self for an adult learner.',
      'Phone': 'The contact\'s phone number, used when a child is unwell or a class is cancelled.',
      'E-mail': 'Where invoices and reports are sent.',
      'Second Contact': 'Another adult who may collect the children or be called, when there is one.',
      'Second Phone': 'The second contact\'s phone number.',
      'Address': 'House number and street, for letters.',
      'City': 'The town or city.',
      'Postcode': 'The postcode.',
      'Pays By': 'How the family usually pays: Bank transfer, Card or Cash.',
      'Notes': 'Anything the office should know, such as a bursary place or who collects the children.',
    } },
    'Students': { about: 'One record per learner, child or adult. Keep a child\'s record to what the centre needs; contact details are on the family.', fields: {
      'Family': 'The family who pays and is contacted for this student.',
      'First Name': 'The student\'s first name, as used in class.',
      'Last Name': 'The family name; lists are sorted by it.',
      'Date of Birth': 'For children, to place them in the right school year and for the birthdays list. Left empty for adults.',
      'School Year': 'The school year the child is in this year, Year 3 to Year 11, or Adult. Change it each September as children move up.',
      'Medical Note': 'A short note for the teacher, such as a nut allergy or an inhaler kept in the bag. Empty for most students.',
      'Photo Consent': 'Yes when the family has agreed that photographs of the student may be taken at the centre.',
      'Joined': 'The day the student first enrolled.',
      'Active': 'Yes while the student is in a class now or was in one in the last two months. Set it to No when they leave; their records stay.',
    } },
    'Teachers': { about: 'The people who teach, paid by the hour for the lessons they take.', fields: {
      'First Name': 'The teacher\'s first name.',
      'Last Name': 'The teacher\'s family name.',
      'Subjects': 'What they teach, such as Maths, Science. Used to find cover when a teacher is away.',
      'Hourly Rate': 'What the centre pays them for an hour of teaching.',
      'Phone': 'The teacher\'s phone number.',
      'E-mail': 'The teacher\'s e-mail address at the centre.',
      'Started': 'The day they started teaching at the centre.',
      'Background Check': 'The date of their latest background check for working with children. The centre renews them every three years.',
      'Active': 'Yes while they teach at the centre.',
    } },
    'Courses': { about: 'What the centre teaches: one record for each subject at each level, with its length and price. A course runs as one or more classes each term.', fields: {
      'Code': 'A short code, such as MA56 for maths for Years 5 and 6. Class codes start with it.',
      'Name': 'The course name on the timetable and invoices.',
      'Subject': 'Maths, English, Science, Coding, Spanish, French, Mandarin or Bookkeeping. Teachers are matched to classes by it.',
      'Level': 'Who it is for: the school years, GCSE, Beginner or Improver.',
      'For': 'Children or Adults.',
      'Holiday Course': 'Yes for the workshops that run in the school holidays.',
      'Minutes': 'How long each lesson lasts.',
      'Fee per Lesson': 'The price of one lesson. A term\'s fee is this times the lessons billed.',
      'Capacity': 'The most students a class of this course takes.',
      'What to Bring': 'What a student brings to each lesson, printed on the enrolment letter.',
    } },
    'Classes': { about: 'One group of students taking a course in one term, on one weekday at one time, in one room, with one teacher.', fields: {
      'Code': 'The course code, the weekday and the start time, such as MA56 Tue 16:00.',
      'Course': 'The course this class teaches.',
      'Teacher': 'The class\'s own teacher. A lesson taken by someone else records them on the lesson.',
      'Term': 'The term or holiday the class runs in.',
      'Weekday': 'The day of the week it meets.',
      'Start Time': 'When the lesson starts, as 16:00.',
      'Room': 'Where it meets: Room 1, Room 2, Room 3 or the Computer Room.',
      'Capacity': 'How many students it takes, set by the course and the room.',
      'Lessons Planned': 'How many lessons it has this term, leaving out the days the centre is closed.',
    } },
    'Enrolments': { about: 'One record for each student in each class: the place booked and what is charged for it.', fields: {
      'Student': 'Who is enrolled.',
      'Class': 'The class they are enrolled in.',
      'Enrolled On': 'The day the place was booked.',
      'Status': 'Enrolled for a class still running or about to start, Completed when the term has ended, Withdrawn when the student left part way through.',
      'Lessons Billed': 'How many lessons are charged for: all of them, or those left for a student who joined after the start.',
      'Discount': 'A share off the fee, as 0.1 for ten percent: brothers and sisters after the first get 0.1, a bursary place 0.5.',
      'Left On': 'The day a withdrawn student left. Fees are not refunded part way through a term.',
      'First Lesson': 'The first lesson the student comes to: the class\'s first, or a later one for a student who joined after the start. The enrolment letter gives it.',
      'Invoice': 'The family\'s fee invoice for the term that charges for this place. Each place is a line of that invoice; it is empty until the invoice is sent.',
    } },
    'Lessons': { about: 'Every lesson of every class: those taught, those planned for the rest of the term and the next, and the few cancelled.', fields: {
      'Class': 'The class the lesson is for.',
      'Teacher': 'Who took the lesson, which is the class\'s teacher unless someone covered. Teacher pay is worked out from it.',
      'Lesson Date': 'The day of the lesson.',
      'Minutes': 'How long it lasted, which is what the teacher is paid for.',
      'Topic': 'What the lesson covered, or is planned to cover.',
      'Status': 'Taught, Planned or Cancelled. Only taught lessons are paid.',
      'Pay Statement': 'The month\'s pay statement the lesson was paid on, set when the month\'s statements are made. Empty for planned and cancelled lessons and for this month\'s.',
    } },
    'Attendance': { about: 'The register: one mark for each student at each lesson taught.', fields: {
      'Lesson': 'The lesson the mark is for.',
      'Student': 'The student marked.',
      'Mark': 'Present, Late or Absent.',
      'Reason': 'Why the student was absent or late, when the family said.',
    } },
    'Assessments': { about: 'The tests and marked pieces of work set to a class. Children\'s classes have a test at half term and one at the end of term; adult courses have one at the end.', fields: {
      'Class': 'The class that sat it.',
      'Assessment Date': 'The day it was sat or handed in.',
      'Title': 'What it was, such as Half-term check or End of term test.',
      'Out Of': 'The highest mark possible.',
    } },
    'Results': { about: 'Each student\'s mark in each assessment they sat. A student away that day has no result.', fields: {
      'Assessment': 'The assessment the mark is for.',
      'Student': 'Who sat it.',
      'Score': 'The mark, out of the assessment\'s Out Of.',
      'Comment': 'The teacher\'s short comment, when there is one; it goes on the progress report.',
    } },
    'Fee Invoices': { about: 'One invoice for each family for each term or holiday they book, covering every class their students take in it. Its lines are the Enrolments that point at it.', fields: {
      'Number': 'The invoice number, given in date order and once only, a new run each year: INV-2026-0042. Families quote it when they pay.',
      'Family': 'Who the invoice is for.',
      'Term': 'The term or holiday it is for.',
      'Invoice Date': 'The day it was sent, three weeks before the term starts or when a late place was booked.',
      'Due Date': 'When it should be paid: by the first day of the term.',
      'Amount': 'The total of the family\'s fees for the term, after discounts.',
      'Paid So Far': 'What has been paid against it. Keep it equal to the total of its payments; what is left is outstanding.',
    } },
    'Payments': { about: 'Each payment received against an invoice. A family paying in instalments has several. Each one is a receipt.', fields: {
      'Number': 'The receipt number, in date order, a new run each year: RCP-2026-0042.',
      'Invoice': 'The invoice the payment is for.',
      'Paid On': 'The day the money arrived.',
      'Amount': 'How much was paid.',
      'Method': 'Bank transfer, Card or Cash.',
      'Balance After': 'What was left to pay on the invoice once this payment was in, printed on the receipt. Nought when the invoice is paid in full.',
    } },
    'Term Reports': { about: 'The progress report sent home at the end of each term: one for each student who was in a class that term. Its lines are in Report Lines.', fields: {
      'Student': 'Who the report is about.',
      'Term': 'The term it reports on.',
      'School Year': 'The school year the student was in that term, kept as it was, since the student\'s own record moves up each September.',
      'Written On': 'The day it was written, in the week after the term ended.',
      'Comment': 'The centre\'s comment on the term as a whole, printed under the lines.',
    } },
    'Report Lines': { about: 'One line of a term report for each class the student took that term: the register added up and the test marks, with the teacher\'s comment.', fields: {
      'Report': 'The term report the line belongs to.',
      'Class': 'The class it reports on; the course and teacher come from it.',
      'Lessons': 'How many lessons the student was marked at that term.',
      'Present': 'How many of them they were present at on time.',
      'Late': 'How many they came to late.',
      'Absent': 'How many they missed.',
      'Half-term %': 'The half-term check mark as a percentage; empty for adult courses, which have no half-term check, or when the student missed it.',
      'End of Term %': 'The end of term test mark as a percentage; empty when it was missed.',
      'Comment': 'The teacher\'s short comment for the class.',
    } },
    'Certificates': { about: 'A certificate of completion for each place in an adult course or a holiday workshop finished with at least three lessons in four attended.', fields: {
      'Number': 'The certificate number, in the order they are issued, a new run each year: CRT-2026-0042.',
      'Enrolment': 'The place it is for: the student, the class, its course and term come from it.',
      'Issued On': 'The day it was given, a few days after the course ended.',
      'Lessons Attended': 'How many lessons the student came to, on time or late.',
      'Hours': 'The hours of teaching those lessons came to.',
    } },
    'Pay Statements': { about: 'A pay statement for each teacher for each month that has ended: the lessons they taught are its lines (the Lessons that point at it), with the hours and gross pay.', fields: {
      'Number': 'The statement number, in month order, a new run each year: PAY-2026-0042.',
      'Teacher': 'Who is paid.',
      'Month': 'The month it pays for, as 2026-09.',
      'Period Start': 'The first day of that month.',
      'Period End': 'The last day of that month.',
      'Paid On': 'The day the pay goes out, the seventh of the month after or the next working day; empty until then.',
      'Hourly Rate': 'The teacher\'s hourly rate for the month.',
      'Lessons': 'How many lessons they taught that month.',
      'Hours': 'The hours those lessons came to.',
      'Gross': 'The pay before tax and National Insurance: each lesson\'s pay rounded to the penny, then added up, as the statement prints it.',
    } },
  },
  queries: {
    'Class lists': 'Every student in every class that is running now or about to start, with the class\'s day, time, room and teacher, any medical note, photo consent and the family\'s phone number. Print it as the register before a lesson.',
    'Timetable, next 7 days': 'Every lesson from today to a week ahead, in date and time order, with the room, the teacher and the planned topic. Use it to plan rooms and cover.',
    'Attendance below 85 percent': 'Students who have missed more than fifteen in a hundred of their lessons in the last 90 days, with at least four lessons in that time, and the family to call.',
    'Fees outstanding by family': 'Each family that still owes something, with how many invoices are open, the oldest due date and the total left to pay, largest first.',
    'Overdue fees': 'Each invoice past its due date that is not fully paid, with its number, the family\'s address, the term, how many days late it is and what is left to pay. The fees letters are written from it.',
    'Classes with free places': 'Classes running now or next with room for more students: places, how many are enrolled and how full the class is. Use it when a family asks for a place.',
    'Results by course': 'For each course, how many test results there were in the last twelve months and the average, highest and lowest marks as percentages.',
    'Progress by student': 'Every test result of the last twelve months, student by student, with the mark, the percentage and the teacher\'s comment. The progress report prints it.',
    'Enrolments by course and term': 'How many places were booked on each course in each term and holiday, and the fees charged for them.',
    'Teacher hours last month': 'For each teacher, the lessons they taught last month, the hours and what they are owed at their hourly rate. Run it at the start of each month for the pay run.',
    'Fees received by month': 'Money received each month from fees, and how many payments made it up.',
    'Birthdays coming up': 'Active students with a birthday in about the next month, with the age they turn. Adults have no date of birth here, so only children appear.',
    'Background checks due': 'Teachers whose background check is 33 months old or more, so it can be renewed before three years are up.',
  },
  reports: {
    'Class register': 'A page for each class running now or next: the students with their school year, medical note, photo consent and the family\'s contact and phone. Teachers take it into the lesson.',
    'Timetable': 'The lessons of the coming week, day by day and in time order, with course, class, room and teacher. Pin it up in the office.',
    'Fees due letters': 'Overdue invoices family by family, with the address to write to and the total owed. Use it to send reminder letters.',
    'Progress report': 'Each student\'s test results over the last year with the teacher\'s comments and their average percentage, to send home at the end of term.',
    'Teacher pay': 'Hours and pay for each teacher for last month, with the totals, for the pay run.',
    'Attendance concerns': 'The students whose attendance has fallen below 85 percent, with the family to call.',
    'Enrolments by course': 'For each course, the places booked and the fees charged term by term, with totals.',
    'Fees received': 'Fee income month by month, with the total.',
  },
  documents: {
    'Fee invoice': 'Sent to each family about three weeks before a term or holiday starts: a line for each child\'s class with the lessons, the price per lesson and any discount, then the total, what has been paid so far and the balance due. Numbered INV-2026-0042. Print it again later and it shows the payments since.',
    'Payment receipt': 'Given or sent for each payment, on A5: the receipt number, who paid, how, which invoice it was for, and what is left to pay on it.',
    'Enrolment letter': 'Sent when a place is booked: the class, its day, time and room, the teacher, the first lesson, the term\'s dates, what to bring and what happens if the student cannot come. It prints for every enrolment, so choose the ones just booked before printing. Have the wording checked for your own centre and country.',
    'Term report': 'The progress report sent home after each term: for each class, the lessons, late marks and absences, the attendance percentage, the half-term and end of term marks and the teacher\'s comment, then the attendance for the term and the centre\'s comment.',
    'Certificate of completion': 'Given to each adult learner who finishes a course and each child who finishes a holiday workshop with three lessons in four attended: the course, its level, the term, the teacher, the lessons and hours, signed by the teacher and the director.',
    'Pay statement': 'Given to each teacher each month: every lesson they taught with its class, minutes and pay at their hourly rate, then the hours and the gross pay. Tax and National Insurance are worked out by the payroll, not here.',
  },
  empty: 'Start with Business, the centre\'s own name, address and bank details, which every document prints. Then Terms (this year\'s terms and holidays), then Teachers and Courses (what you teach, how long and at what price). Then set up the Classes for the first term, each with its course, teacher, day, time and room. Enter Families and then their Students as they book, and an Enrolment for each place. Add the Lessons of each class (one for each week of the term), take the register in Attendance as the lessons happen, and record Assessments and Results when tests are marked. Raise a Fee Invoice for each family each term and set each of their enrolments\' Invoice to it, then enter Payments as they come in, adding each to the invoice\'s Paid So Far. At the end of a term write a Term Report for each student, with a Report Line for each class; at the end of a month make each teacher\'s Pay Statement and set it on the lessons they taught.',
  adapt: [
    'For one-to-one tutoring, set each course\'s Capacity to 1 and make a class for each student\'s regular slot; the timetable and teacher pay work as they are.',
    'For a music or dance school, use the subjects for instruments or dance styles and the levels for grades; add a Grade field to Results for exam grades.',
    'If you charge by the month instead of by the term, make each month a Term record and raise the invoices monthly.',
    'Add a Waiting status to Enrolments for full classes, and a query listing waiting students by class, oldest first.',
    'For a driving school, make each learner\'s block of lessons a class with Capacity 1 and the instructor as the teacher; add a Test Date field to Students.',
    'To keep a fuller register, add a Minutes Late field to Attendance, or a Homework Done yes/no field.',
    'If teachers are paid a fixed amount a lesson instead of by the hour, add a Rate per Lesson field to Teachers and change the Pay column of Teacher hours last month to use it.',
    'Add a Rooms table with each room\'s size if you have many rooms, and point Classes at it instead of typing the room.',
    'If your centre is registered for VAT and its courses are not exempt, add a Tax Rate field to Courses and a VAT column to the Fee invoice print template, as the trading sample\'s invoice has.',
  ],
  tutorial: [
    { title: 'Book a new family and their child into a class', steps: [
      'Open the query Classes with free places to find a class of the right course with room in it.',
      'Open the Families form and press New. Type the family\'s Name (such as Okafor family), the Contact, Phone, E-mail and address.',
      'Open the Students form and press New. Choose the family in Family, type the child\'s names, Date of Birth and School Year, any Medical Note, and tick Photo Consent if the family agrees.',
      'Open the Enrolments form and press New. Choose the student and the Class, type today in Enrolled On, Status Enrolled, the Lessons Billed (all of the class\'s lessons, or those left), the Discount if any, and the First Lesson they will come to.',
      'Open the print template Enrolment letter from Print templates, choose the new enrolment, and press Print or Save as PDF.',
    ] },
    { title: 'Send the term\'s fee invoices', steps: [
      'About three weeks before a term, open the Fee Invoices form and press New for each family with places booked. The Number is the next in the year\'s run (INV-2026-0043 after INV-2026-0042). Choose the Family and Term, and type the Invoice Date and the Due Date (the first day of term).',
      'Type the Amount: the sum of the family\'s places, each the course\'s Fee per Lesson times the Lessons Billed, less its Discount.',
      'Open the Enrolments form, find each of the family\'s places for the term and choose the new invoice in Invoice. These are the invoice\'s lines.',
      'Open the print template Fee invoice, choose the invoices just made, and press Print or Save as PDF. Check that the Total for the term matches the Amount you typed.',
    ] },
    { title: 'Record a payment and give a receipt', steps: [
      'Open the Payments form and press New. Type the next receipt Number (RCP-2026-0042), choose the Invoice, and type Paid On, the Amount and the Method.',
      'Type in Balance After what is still owed on the invoice once this payment is in.',
      'Open the Fee Invoices form, find the invoice and add the payment to Paid So Far.',
      'Open the print template Payment receipt, choose the payment and press Print. It fits on A5.',
    ] },
    { title: 'Take the register and mark a test', steps: [
      'Print the report Class register before the lessons: a page for each class with medical notes and the family\'s phone number.',
      'After the lesson, open the Lessons form, find the lesson by its date and class, and set Status to Taught. Set Teacher to whoever took it if someone covered.',
      'Open the Attendance form and press New for each student: choose the Lesson and Student and set Mark to Present, Late or Absent, with a Reason if the family gave one.',
      'On a test day, open the Assessments form and press New: choose the Class, type the Assessment Date, Title and Out Of. Then add a Results record for each student who sat it, with their Score.',
    ] },
    { title: 'Write and print the term reports', steps: [
      'In the week after a term ends, open the query Progress by student to see each student\'s marks, and the report Attendance concerns for anyone who has missed a lot.',
      'Open the Term Reports form and press New for each student: choose the Student and Term, type the School Year they were in, the Written On date and a short Comment on the term.',
      'In the Report Lines below, add a line for each class they took: the Class, the Lessons, Present, Late and Absent from the register, the Half-term % and End of Term % from their results, and the teacher\'s comment.',
      'Open the print template Term report, choose the term\'s reports, and press Print or Save as PDF to send home.',
      'For adult courses and holiday workshops finished, open the Certificates form and press New for each learner who came to three lessons in four: choose the Enrolment, type the Number, Issued On, Lessons Attended and Hours. Print them with the print template Certificate of completion.',
    ] },
    { title: 'Pay the teachers for the month', steps: [
      'On the first working day of the month, run the query Teacher hours last month, or print the report Teacher pay: the lessons, hours and pay for each teacher for the month just ended.',
      'Open the Pay Statements form and press New for each teacher: the Number (PAY-2026-0042), Teacher, Month (2026-09), Period Start and Period End, the Hourly Rate, and the Lessons, Hours and Gross from the query.',
      'Open the Lessons form and set Pay Statement on each lesson they taught that month, so the statement lists them.',
      'Open the print template Pay statement, choose the month\'s statements, and press Print. Type Paid On when the pay has gone out.',
    ] },
    { title: 'Chase fees that are overdue', steps: [
      'Run the query Fees outstanding by family to see who owes what, largest first.',
      'Print the report Fees due letters: each overdue invoice, family by family, with the address and what is left to pay.',
      'Print the family\'s Fee invoice again from Print templates: it shows the payments made so far and the balance due.',
    ] },
  ],
  notes: [
    'An invoice\'s lines are the enrolments that point at it, so an enrolment with no Invoice is not on any invoice yet. Set it when you send the invoice; a place booked after the invoice was sent goes on a new invoice of its own.',
    'Paid So Far on an invoice and Balance After on a payment are typed, not worked out. Keep Paid So Far equal to the invoice\'s payments added up, or the balance printed on the invoice will be wrong.',
    'A print template prints every record of its table, so the Enrolment letter, the Fee invoice and the others print the whole history unless you choose the records first. Choose the ones you want (the new enrolments, this term\'s invoices) before pressing Print.',
    'Term reports and certificates are records of their own, written once at the end of a term. Changing the register later does not change a report already written: update its line by hand if a mark was wrong.',
    'A pay statement\'s lines are the lessons that point at it. A lesson covered by another teacher has that teacher in its Teacher field, so it is paid on their statement, not the class teacher\'s.',
    'Tuition here is exempt from VAT, so the Business record\'s Tax Rate is 0 and no document shows VAT. If yours is not, see the last of the ways to adapt the sample.',
    'School Year on a student is this year\'s and moves up each September; the term report keeps the year the student was in at the time.',
  ],
  faq: [
    { q: 'How do I print one family\'s invoice?', a: 'Open the print template Fee invoice, choose that family\'s invoice from the list (or open it in the Fee Invoices form and print from there), and press Print or Save as PDF.' },
    { q: 'How is the discount for brothers and sisters worked out?', a: 'The first child booked in a term pays the full fee; each brother or sister booked after them has 0.1 in Discount on their enrolments, ten percent off. A bursary place has 0.5. The invoice shows the fee, the discount taken off and what is left to pay on each line.' },
    { q: 'A child joins half way through the term. What do I charge?', a: 'Set Lessons Billed on the enrolment to the lessons left, and First Lesson to the first one they will come to. The invoice charges for those lessons only.' },
    { q: 'Why does the invoice say nothing about VAT?', a: 'Private tuition and education are exempt from VAT in many countries, and the sample is set up that way. If you must charge it, add a tax rate as the ways to adapt the sample describe, and ask your accountant first.' },
    { q: 'How do I print last month\'s pay statements?', a: 'Open the print template Pay statement and choose the statements whose Month is last month, then press Print. Each lists the lessons the teacher taught, the hours and the gross pay.' },
    { q: 'Who gets a certificate?', a: 'In the sample, adult learners who finish a course and children who finish a holiday workshop, when they came to at least three lessons in four. Children on term courses get their term report instead. Change the rule to suit your centre: a certificate is simply a record in Certificates.' },
    { q: 'Can I change the wording of the enrolment letter?', a: 'Yes. Open Print templates, choose Enrolment letter and edit its text blocks. The words in curly brackets, such as {Class.Weekday}, are filled from the enrolment and its class; keep them as they are or the letter will print them as blanks.' },
  ],
};

// The rooms and how many each holds. Coding and bookkeeping are taught in the computer room.
const ROOMS = [['Room 1', 10], ['Room 2', 10], ['Room 3', 16], ['Computer Room', 8]];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

// The teachers: subjects and hourly rate.
const TEACHERS = [
  ['Maths', 32], ['Maths, Science', 30], ['English', 30], ['English, Maths', 26], ['Science', 31],
  ['Coding', 34], ['Coding, Maths', 28], ['Spanish, French', 27], ['Mandarin', 29], ['Bookkeeping', 35],
];

// What the centre teaches: code, name, subject, level, for, minutes, fee per lesson, capacity, school years.
const COURSES = [
  ['MA34', 'Maths, Years 3 and 4', 'Maths', 'Years 3-4', 'Children', 60, 16, 10, [3, 4]],
  ['MA56', 'Maths, Years 5 and 6', 'Maths', 'Years 5-6', 'Children', 60, 17, 10, [5, 6]],
  ['MA79', 'Maths, Years 7 to 9', 'Maths', 'Years 7-9', 'Children', 60, 18, 10, [7, 9]],
  ['MAGC', 'GCSE Maths', 'Maths', 'GCSE', 'Children', 90, 26, 8, [10, 11]],
  ['EN34', 'English, Years 3 and 4', 'English', 'Years 3-4', 'Children', 60, 16, 10, [3, 4]],
  ['EN56', 'English, Years 5 and 6', 'English', 'Years 5-6', 'Children', 60, 17, 10, [5, 6]],
  ['EN79', 'English, Years 7 to 9', 'English', 'Years 7-9', 'Children', 60, 18, 10, [7, 9]],
  ['ENGC', 'GCSE English', 'English', 'GCSE', 'Children', 90, 26, 8, [10, 11]],
  ['SC56', 'Science, Years 5 and 6', 'Science', 'Years 5-6', 'Children', 60, 17, 10, [5, 6]],
  ['SC79', 'Science, Years 7 to 9', 'Science', 'Years 7-9', 'Children', 60, 18, 10, [7, 9]],
  ['SCGC', 'GCSE Science', 'Science', 'GCSE', 'Children', 90, 26, 8, [10, 11]],
  ['CD1', 'Coding for beginners', 'Coding', 'Years 5-8', 'Children', 90, 20, 8, [5, 8]],
  ['CD2', 'Coding, next steps', 'Coding', 'Years 7-11', 'Children', 90, 22, 8, [7, 11]],
  ['ES1', 'Spanish for beginners', 'Spanish', 'Beginner', 'Adults', 120, 15, 16, null],
  ['ES2', 'Spanish, improvers', 'Spanish', 'Improver', 'Adults', 120, 15, 10, null],
  ['FR1', 'French for beginners', 'French', 'Beginner', 'Adults', 120, 15, 16, null],
  ['ZH1', 'Mandarin for beginners', 'Mandarin', 'Beginner', 'Adults', 120, 16.5, 10, null],
  ['BK1', 'Bookkeeping, part 1', 'Bookkeeping', 'Beginner', 'Adults', 120, 19.5, 8, null],
  ['BK2', 'Bookkeeping, part 2', 'Bookkeeping', 'Improver', 'Adults', 120, 21, 8, null],
  // Holiday workshops, on the weekdays given, so every weekday of a holiday has something on.
  ['HMB', 'Holiday maths booster', 'Maths', 'Years 5-9', 'Children', 120, 24, 10, [5, 9], [1, 3]],
  ['HGM', 'GCSE maths revision day', 'Maths', 'GCSE', 'Children', 120, 26, 8, [10, 11], [2]],
  ['HGE', 'GCSE English revision day', 'English', 'GCSE', 'Children', 120, 26, 8, [10, 11], [4]],
  ['HRW', 'Holiday reading and writing club', 'English', 'Years 3-6', 'Children', 120, 22, 10, [3, 6], [3]],
  ['HCC', 'Holiday coding club', 'Coding', 'Years 5-9', 'Children', 120, 24, 8, [5, 9], [2, 5]],
  ['HSL', 'Holiday science lab', 'Science', 'Years 5-8', 'Children', 120, 24, 10, [5, 8], [4]],
];

const TOPICS = {
  Maths: ['Place value', 'Times tables', 'Fractions', 'Decimals', 'Percentages', 'Long multiplication', 'Long division', 'Angles', 'Area and perimeter', 'Ratio', 'Algebra: expressions', 'Solving equations', 'Straight-line graphs', 'Probability', 'Averages and charts', 'Problem solving', 'Practice paper'],
  English: ['Reading for meaning', 'Spelling patterns', 'Punctuation', 'Paragraphs', 'Persuasive writing', 'Descriptive writing', 'Poetry', 'Comprehension', 'Grammar: clauses', 'Story planning', 'Speaking and listening', 'Set text', 'Practice paper'],
  Science: ['Living things', 'Forces', 'Electricity', 'Light', 'The Earth and space', 'States of matter', 'Cells', 'Chemical reactions', 'Energy', 'Waves', 'Working scientifically', 'Practice paper'],
  Coding: ['Loops', 'Variables', 'Conditions', 'Making a game', 'Lists', 'Functions', 'Finding bugs', 'Web pages', 'Data and files', 'Project work', 'Showcase'],
  Bookkeeping: ['Double entry', 'Day books', 'Bank reconciliation', 'Sales and purchase ledgers', 'Trial balance', 'Payroll basics', 'Sales tax returns', 'Year end', 'Spreadsheet practice', 'Review and test'],
  Language: ['Greetings', 'Numbers and dates', 'Family', 'Food and ordering', 'Directions', 'Shopping', 'Daily routine', 'Travel', 'The past tense', 'Conversation practice', 'Review'],
};
const OUT_OF = { Maths: 40, English: 30, Science: 40, Coding: 20, Bookkeeping: 100, Language: 50 };
const LANGUAGES = new Set(['Spanish', 'French', 'Mandarin']);
const MEDICAL = ['Nut allergy', 'Asthma, inhaler in bag', 'Hay fever in summer', 'Wears glasses for reading', 'Dairy free', 'Mild eczema'];
const COMMENTS = ['Good progress', 'Show your working', 'Well prepared', 'Revise the topics marked', 'Big step up since the last test', 'Careless slips, check answers', 'Excellent'];
// What each subject's students bring, for the enrolment letter.
const BRING = {
  Maths: 'A pencil, a ruler, a calculator from Year 6 up, and any homework from school the student would like help with.',
  English: 'A pen, a pencil and the book the student is reading at school, if there is one.',
  Science: 'A pen, a pencil, a ruler and a calculator. Aprons and safety glasses for practical work are kept at the centre.',
  Coding: 'Nothing but a water bottle: the computers are the centre\'s. A USB stick is useful for taking projects home.',
  Bookkeeping: 'A pen and a notepad. The computers and the course workbook are provided.',
  Language: 'A pen and a notebook. The course book is lent for the term and handed back at the end.',
};
// What a teacher writes on a report line, by how the term went.
const REPORT_COMMENTS = {
  high: ['Excellent work all term', 'A very strong term', 'Confident and accurate', 'Ready for harder work'],
  mid: ['Steady progress', 'Good effort in class', 'Doing well; keep practising at home', 'Improving week by week'],
  low: ['Needs more practice at home', 'Finds this hard but keeps trying', 'Would benefit from coming every week', 'Revise the topics on the test'],
};
const STREETS = ['Harehills Lane', 'Mill Lane', 'Station Road', 'Orchard Way', 'Church Street', 'Canal Walk', 'Beech Grove', 'Riverside Drive', 'Quarry Road', 'Cedar Avenue', 'Linden Close', 'Moor View', 'Park Terrace'];

/** Fill a Lantern Learning Centre database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };

  add('Business', businessRecord({
    name: 'Lantern Learning Centre', domain: 'lantern.example', phone: '020 7946 0420',
    address: { Address: '27 Harehills Lane', City: 'Leeds', Postcode: 'LS8 4DE', Country: 'United Kingdom' },
    terms: 'Fees are due by the first day of term. Tuition is exempt from VAT, so no VAT is charged.', taxRate: 0,
    footer: 'Office open Monday to Friday 15:00 to 19:30 and Saturday 09:00 to 13:00. If a student is not coming back next term, please tell us two weeks before it starts.',
  }));

  // Days counted from today, for a calendar date (month from 0), worked out without a clock.
  const t0 = cal.at(0);
  const day0 = Date.UTC(t0.getFullYear(), t0.getMonth(), t0.getDate());
  const off = (y, m, d) => Math.round((Date.UTC(y, m, d) - day0) / 86400000);
  const mondayFrom = (o) => { while (cal.weekday(o) !== 1) o++; return o; };
  const saturdayTo = (o) => { while (cal.weekday(o) !== 6) o--; return o; };
  // The school year a day falls in, named by the year it starts in September.
  const yearOf = (o) => { const d = cal.at(o); return d.getMonth() >= 8 ? d.getFullYear() : d.getFullYear() - 1; };
  // Closed on Sundays, Christmas Eve to Boxing Day, and New Year's Eve and Day.
  const closed = (o) => {
    const d = cal.at(o);
    const m = d.getMonth();
    const n = d.getDate();
    return d.getDay() === 0 || (m === 11 && [24, 25, 26, 31].includes(n)) || (m === 0 && n === 1);
  };

  // The blocks of weeks: three terms a year and the holidays between them.
  const blocks = [];
  const now = yearOf(0);
  for (let a = now - 2; a <= now + 1; a++) {
    const autumn = mondayFrom(off(a, 8, 6));
    const autumnEnd = saturdayTo(off(a, 11, 20));
    const spring = mondayFrom(off(a + 1, 0, 4));
    const springEnd = spring + 12 * 7 - 2;
    const summer = springEnd + 16;
    const summerEnd = saturdayTo(off(a + 1, 6, 18));
    const nextAutumn = mondayFrom(off(a + 1, 8, 6));
    blocks.push(
      { name: `Autumn ${a}`, kind: 'Term', start: autumn, end: autumnEnd },
      { name: `Christmas holiday ${a}`, kind: 'Holiday', start: autumnEnd + 1, end: spring - 1 },
      { name: `Spring ${a + 1}`, kind: 'Term', start: spring, end: springEnd },
      { name: `Easter holiday ${a + 1}`, kind: 'Holiday', start: springEnd + 1, end: summer - 1 },
      { name: `Summer ${a + 1}`, kind: 'Term', start: summer, end: summerEnd },
      { name: `Summer holiday ${a + 1}`, kind: 'Holiday', start: summerEnd + 1, end: nextAutumn - 1 },
    );
  }
  // Back to a term that began fourteen to eighteen months ago, and the next block when it starts within six weeks.
  const kept = blocks.filter((b) => b.start >= -540 && b.start <= 40);
  while (kept[0].kind !== 'Term') kept.shift();

  // Teachers.
  const teachers = TEACHERS.map(([subjects, rate], i) => {
    const p = person(r, { domain: 'lantern.example' });
    // Two checks are nearly three years old, so the renewal list has someone on it.
    const check = i === 2 || i === 7 ? -r.int(1010, 1070) : -r.int(40, 900);
    const id = add('Teachers', {
      ...p, Subjects: subjects, 'Hourly Rate': rate, Phone: `07700 900${String(310 + i * 7).padStart(3, '0')}`,
      Started: cal.day(Math.min(check, -r.int(500, 3200))), 'Background Check': cal.day(check), Active: true,
    });
    return { id, subjects: subjects.split(', '), rate, load: 0 };
  });

  // Courses.
  const courses = COURSES.map(([code, name, subject, level, audience, minutes, fee, capacity, years, days]) => {
    const id = add('Courses', { Code: code, Name: name, Subject: subject, Level: level, For: audience, 'Holiday Course': Boolean(days), Minutes: minutes, 'Fee per Lesson': fee, Capacity: capacity, 'What to Bring': BRING[LANGUAGES.has(subject) ? 'Language' : subject] });
    return { id, code, name, subject, audience, minutes, fee, capacity, years, days: days || null };
  });
  const course = Object.fromEntries(courses.map((c) => [c.code, c]));

  // Families and students are made up as they come, and written when they first book a place.
  const families = [];
  const students = [];
  const familyAddress = () => ({ Address: `${r.int(1, 220)} ${r.pick(STREETS)}`, City: 'Leeds', Postcode: `LS${r.int(1, 17)} ${r.int(1, 9)}${r.pick(['AB', 'DE', 'FG', 'HJ', 'NP', 'QR', 'TU'])}` });
  const newFamily = (adult) => {
    const p = person(r);
    const f = {
      id: null, last: p['Last Name'], kids: 0,
      pays: r.weighted([['prompt', 70], ['instalments', 12], ['late', 12], ['poor', 6]]),
      method: r.weighted([['Bank transfer', 5], ['Card', 4], ['Cash', 1]]),
      bursary: !adult && r.chance(0.03),
    };
    const second = !adult && r.chance(0.5) ? person(r) : null;
    f.values = {
      Name: adult ? `${p['First Name']} ${p['Last Name']}` : `${p['Last Name']} family`, Contact: `${p['First Name']} ${p['Last Name']}`,
      Relationship: adult ? 'Self' : r.weighted([['Parent', 12], ['Grandparent', 1], ['Carer', 0.5]]),
      Phone: phoneFor(r, 'United Kingdom'), 'E-mail': p['E-mail'],
      'Second Contact': second ? `${second['First Name']} ${r.chance(0.7) ? p['Last Name'] : second['Last Name']}` : null,
      'Second Phone': second ? phoneFor(r, 'United Kingdom') : null,
      ...familyAddress(), 'Pays By': f.method,
      Notes: f.bursary ? 'Bursary place: half fees' : !adult && r.chance(0.08) ? r.pick(['Grandparent collects on Fridays', 'Children walk home together', 'Prefers e-mail to phone calls']) : null,
    };
    families.push(f);
    return f;
  };
  const newChild = (f, year, a, joinedBy) => {
    const subjects = [];
    if (r.chance(0.68)) subjects.push('Maths');
    if (r.chance(0.45)) subjects.push('English');
    if (year >= 5 && r.chance(0.22)) subjects.push('Science');
    if (year >= 5 && r.chance(0.2)) subjects.push('Coding');
    if (!subjects.length) subjects.push(r.pick(['Maths', 'English']));
    const cohort = a - year - 5;
    const s = {
      id: null, family: f, adult: false, cohort, first: r.pick(FIRST), last: r.chance(0.9) ? f.last : r.pick(LAST),
      dob: cal.day(off(cohort, 8, 1) + r.int(0, 364)), subjects: r.some(subjects, subjects.length), coded: 0,
      attend: r.weighted([[0.97, 70], [0.93, 22], [0.8, 8]]), ability: {}, joinedBy,
      medical: r.chance(0.12) ? r.pick(MEDICAL) : null, consent: r.chance(0.82), last_on: -9999,
    };
    f.kids++;
    students.push(s);
    return s;
  };
  const newAdult = (joinedBy) => {
    const f = newFamily(true);
    const s = {
      id: null, family: f, adult: true, first: f.values.Contact.split(' ')[0], last: f.last, dob: null,
      course: r.weighted([['ES1', 30], ['FR1', 20], ['ZH1', 15], ['BK1', 25], ['ES2', 10]]), concession: r.chance(0.1),
      attend: r.weighted([[0.95, 62], [0.9, 28], [0.8, 10]]), ability: {}, joinedBy, medical: null, consent: r.chance(0.7), last_on: -9999,
    };
    students.push(s);
    return s;
  };
  // A new family with one child or more, the eldest in `year`.
  const newChildren = (a, year, joinedBy) => {
    const f = newFamily(false);
    const kids = [newChild(f, year, a, joinedBy)];
    for (let k = 1; k < r.weighted([[1, 70], [2, 25], [3, 5]]); k++) {
      const y = year - r.int(1, 4);
      if (y >= 3) kids.push(newChild(f, y, a, joinedBy));
    }
    return kids;
  };
  const YEARS = [[3, 1], [4, 1.2], [5, 1.6], [6, 1.6], [7, 1.1], [8, 1], [9, 1.1], [10, 1.5], [11, 1.4]];

  // A student's record, and their family's, written the first time they book.
  const save = (s, on) => {
    if (s.id) return s.id;
    const f = s.family;
    if (!f.id) f.id = add('Families', f.values);
    s.id = add('Students', {
      Family: f.id, 'First Name': s.first, 'Last Name': s.last, 'Date of Birth': s.dob, 'School Year': s.adult ? 'Adult' : '',
      'Medical Note': s.medical, 'Photo Consent': s.consent, Joined: cal.day(Math.min(on, s.joinedBy ?? on)), Active: false,
    });
    return s.id;
  };

  const courseFor = (s, subject, year) => {
    const band = year <= 4 ? '34' : year <= 6 ? '56' : year <= 9 ? '79' : 'GC';
    if (subject === 'Maths') return course[`MA${band}`];
    if (subject === 'English') return course[`EN${band}`];
    if (subject === 'Science') return year >= 5 ? course[`SC${band}`] : null;
    if (subject === 'Coding') return year <= 6 || (year <= 8 && s.coded < 2) ? course.CD1 : course.CD2;
    return null;
  };

  // Slots: weekday, start time, the longest lesson that fits before the next slot, and who it is for.
  const SLOTS = [];
  for (let wd = 1; wd <= 5; wd++) SLOTS.push([wd, '16:00', 60, 'Children'], [wd, '17:15', 90, 'Children']);
  for (let wd = 1; wd <= 4; wd++) SLOTS.push([wd, '19:00', 120, 'Adults']);
  SLOTS.push([6, '09:30', 90, 'Children'], [6, '11:15', 90, 'Children']);
  const roomsFor = (c) => ROOMS.filter(([name, size]) => (c.subject === 'Coding' || c.subject === 'Bookkeeping' ? name === 'Computer Room' : name !== 'Computer Room' && size >= c.capacity)).sort((x, y) => x[1] - y[1]);

  const invoices = [];
  const reportLines = []; // { termId, b, a, s, classId, st } for each student in each class of an ended term
  const finished = []; // { b, k, enrolment, st }: places in ended adult courses and holiday workshops
  const taught = []; // { id, t, o, minutes }: lessons taught, for the pay statements
  let pool = []; // students in the last term
  let lastTerm = null;
  const absentWhy = (s) => (s.adult ? r.weighted([['Working late', 3], ['Ill', 2], ['Away', 2], ['No reason given', 2]]) : r.weighted([['Ill', 4], ['Family holiday', 2], ['School event', 2], ['Medical appointment', 1], ['Family event', 1], ['No reason given', 2]]));

  for (const b of kept) {
    const a = yearOf(b.start);
    const termId = add('Terms', { Name: b.name, Kind: b.kind, Starts: cal.day(b.start), Ends: cal.day(b.end) });
    const ended = b.end < 0;
    const yearIn = (s) => a - s.cohort - 5;
    // Who wants what this block: [{ s, c, on }].
    const wants = [];
    const roomBusy = new Set();
    const teacherBusy = new Set();
    const classes = [];
    for (const t of teachers) t.load = 0;
    const kids = pool.filter((s) => !s.adult && !s.gone && yearIn(s) <= 11);
    const fits = (s, c) => { const y = yearIn(s); return y >= c.years[0] && y <= c.years[1]; };
    const on = () => b.start - r.int(1, 30);
    const dates = (wd) => { const out = []; for (let o = b.start; o <= b.end; o++) if (cal.weekday(o) === wd && !closed(o)) out.push(o); return out; };
    const open = (c, wd, time, room, size) => {
      const ds = dates(wd);
      if (!ds.length) return null;
      const t = teachers.filter((x) => x.subjects.includes(c.subject) && !teacherBusy.has(`${x.id}|${wd}|${time}`)).sort((x, y) => x.load - y.load)[0];
      if (!t) return null;
      t.load++;
      teacherBusy.add(`${t.id}|${wd}|${time}`);
      roomBusy.add(`${room}|${wd}|${time}`);
      const k = { c, wd, time, room, t, dates: ds, cap: Math.min(c.capacity, size), places: [] };
      classes.push(k);
      return k;
    };

    if (b.kind === 'Term') {
      // Who comes back: children until they finish Year 11, adults for a term or a few.
      const children = pool.filter((s) => !s.adult && !s.gone && yearIn(s) <= 11 && r.chance(lastTerm ? 0.88 : 1));
      const adults = pool.filter((s) => s.adult && !s.gone && r.chance(0.6));
      for (const s of pool) if (!children.includes(s) && !adults.includes(s)) s.gone = true;
      const first = !lastTerm;
      // New families, more of them in the autumn.
      const target = 86 + r.int(-4, 6) + (b.name.startsWith('Autumn') ? 4 : 0);
      while (children.length < target) {
        const joinedBy = first ? b.start - r.int(30, 900) : null;
        const kids = newChildren(a, r.weighted(YEARS), joinedBy);
        const late = !first && r.chance(0.2);
        for (const s of kids) { s.isNew = !first; s.late = late; children.push(s); }
      }
      while (adults.length < 30 + r.int(-3, 4)) { const s = newAdult(first ? b.start - r.int(30, 600) : null); s.isNew = !first; s.late = !first && r.chance(0.15); adults.push(s); }
      // Adults who finished part 1 often go on to the next course.
      for (const s of adults) if (!s.isNew && ((s.course === 'ES1' && r.chance(0.4)) || (s.course === 'BK1' && r.chance(0.5)))) s.course = s.course === 'ES1' ? 'ES2' : 'BK2';
      const when = (s) => (s.isNew ? (s.late ? b.start + r.int(1, 13) : b.start - r.int(0, 28)) : b.start - r.int(10, 35));
      for (const s of children) {
        const on = when(s);
        const taken = [];
        for (const subject of s.subjects) {
          const c = courseFor(s, subject, yearIn(s));
          if (c && taken.length < 2) { taken.push(c); wants.push({ s, c, on }); }
        }
      }
      for (const s of adults) wants.push({ s, c: course[s.course], on: when(s) });
      pool = [...children, ...adults];
      for (const s of pool) { s.isNew = false; s.late = false; }
      lastTerm = b;

      // Open as many classes of each course as its students need, each in a free slot with a free teacher.
      for (const c of courses.filter((x) => !x.days)) {
        const n = wants.filter((w) => w.c === c).length;
        if (!n || (c.audience === 'Adults' && n < 4)) continue;
        const sections = Math.ceil(n / c.capacity);
        const used = new Set();
        for (let k = 0; k < sections; k++) {
          const slots = r.some(SLOTS.filter(([, , max, who]) => who === c.audience && max >= c.minutes), SLOTS.length);
          slots.sort((x, y) => Number(used.has(x[0])) - Number(used.has(y[0])));
          let made = null;
          for (const [wd, time] of slots) {
            for (const [room, size] of roomsFor(c)) {
              if (roomBusy.has(`${room}|${wd}|${time}`)) continue;
              made = open(c, wd, time, room, size);
              if (made) break;
            }
            if (made) { used.add(wd); break; }
          }
        }
      }
    } else {
      // Holiday workshops on their weekdays, mornings and afternoons, for children in last term.
      let morning = true;
      for (const c of courses.filter((x) => x.days)) {
        for (const wd of c.days) {
          const time = morning ? '10:00' : '13:00';
          morning = !morning;
          const room = roomsFor(c).find(([name]) => !roomBusy.has(`${name}|${wd}|${time}`));
          if (room) open(c, wd, time, room[0], room[1]);
        }
      }
      for (const s of kids) {
        if (!r.chance(b.name.startsWith('Christmas') ? 0.2 : 0.32)) continue;
        const choice = courses.filter((c) => c.days && fits(s, c) && (s.subjects.includes(c.subject) || r.chance(0.15)));
        if (choice.length) wants.push({ s, c: r.pick(choice), on: on() });
      }
    }

    // Places: each student into the class of their course with most room, never two at once.
    const busy = new Map();
    for (const w of r.some(wants, wants.length)) {
      const mine = busy.get(w.s) || new Set();
      const k = classes.filter((x) => x.c === w.c && x.places.length < x.cap && !mine.has(`${x.wd}|${x.time}`)).sort((x, y) => x.places.length - y.places.length)[0];
      if (!k) continue;
      mine.add(`${k.wd}|${k.time}`);
      busy.set(w.s, mine);
      k.places.push({ s: w.s, on: Math.min(w.on, k.dates[k.dates.length - 1]) });
    }
    // A holiday workshop runs with three at least: the office rings round to fill it.
    if (b.kind === 'Holiday') for (const k of classes) {
      const free = (x) => fits(x, k.c) && !k.places.some((p) => p.s === x) && !busy.get(x)?.has(`${k.wd}|${k.time}`);
      for (const s of r.some(kids.filter(free), Math.max(0, 3 - k.places.length))) k.places.push({ s, on: on() });
    }

    // Write the classes, their places, their lessons and registers, and their tests.
    const billed = new Map(); // family -> [{ amount, on }]
    for (const k of classes) {
      const code = `${k.c.code} ${DAYS[k.wd].slice(0, 3)} ${k.time}`;
      const classId = add('Classes', { Code: code, Course: k.c.id, Teacher: k.t.id, Term: termId, Weekday: DAYS[k.wd], 'Start Time': k.time, Room: k.room, Capacity: k.cap, 'Lessons Planned': k.dates.length });
      const places = [];
      for (const p of k.places.sort((x, y) => x.on - y.on)) {
        if (p.on > 0) continue; // not booked yet
        const s = p.s;
        const lessons = k.dates.filter((o) => o >= p.on).length;
        // Brothers and sisters after the first child booked this term get ten percent off.
        const seen = (s.family.seen ??= new Map());
        if (!seen.has(termId)) seen.set(termId, []);
        if (!seen.get(termId).includes(s)) seen.get(termId).push(s);
        const nth = seen.get(termId).indexOf(s);
        const discount = s.family.bursary ? 0.5 : s.adult ? (s.concession ? 0.2 : 0) : b.kind === 'Term' && nth > 0 ? 0.1 : 0;
        // A few leave part way through a term; their fees stand.
        const left = b.kind === 'Term' && k.dates.length > 4 && r.chance(0.03) ? k.dates[r.int(2, k.dates.length - 2)] : null;
        const gone = left !== null && left <= 0;
        if (gone) { s.gone = true; s.withdrew = true; }
        const enrolment = add('Enrolments', {
          Student: save(s, p.on), Class: classId, 'Enrolled On': cal.day(p.on), Status: gone ? 'Withdrawn' : ended ? 'Completed' : 'Enrolled',
          'Lessons Billed': lessons, Discount: discount, 'Left On': gone ? cal.day(left) : null, 'First Lesson': cal.day(k.dates.find((o) => o >= p.on)),
        });
        s.last_on = Math.max(s.last_on, gone ? left : b.end);
        const bill = billed.get(s.family) || [];
        bill.push({ amount: money(k.c.fee * lessons * (1 - discount)), on: p.on, enrolment });
        billed.set(s.family, bill);
        // The register and the tests, added up for the term report and the certificate.
        const st = { lessons: 0, present: 0, late: 0, absent: 0, half: null, end: null, comment: null };
        places.push({ s, on: p.on, left: gone ? left : null, st });
        if (ended && !gone) {
          if (b.kind === 'Term') reportLines.push({ termId, b, a, s, classId, st });
          if (k.c.audience === 'Adults' || k.c.days) finished.push({ b, k, enrolment, st });
        }
      }

      const subject = LANGUAGES.has(k.c.subject) ? 'Language' : k.c.subject;
      const topics = TOPICS[subject];
      // Each class works through the topics from its own starting point.
      const from = r.int(0, topics.length - 1);
      const tests = k.c.days ? [] : k.c.audience === 'Adults' ? [k.dates.length - 1] : [Math.floor(k.dates.length / 2) - 1, k.dates.length - 1];
      k.dates.forEach((o, i) => {
        let teacher = k.t;
        let status = o < 0 ? 'Taught' : 'Planned';
        if (o < 0 && r.chance(0.05)) {
          // The teacher is away: another who teaches the subject covers, or the lesson is cancelled.
          const cover = teachers.find((x) => x !== k.t && x.subjects.includes(k.c.subject) && !teacherBusy.has(`${x.id}|${k.wd}|${k.time}`));
          if (cover && r.chance(0.8)) teacher = cover;
          else status = 'Cancelled';
        }
        const topic = tests.includes(i) ? (k.c.audience === 'Adults' ? 'End of term assessment' : i === tests[0] && tests.length > 1 ? 'Half-term check' : 'End of term test') : topics[(i + from) % topics.length];
        const lesson = add('Lessons', { Class: classId, Teacher: teacher.id, 'Lesson Date': cal.day(o), Minutes: k.c.minutes, Topic: topic, Status: status });
        if (status !== 'Taught') return;
        taught.push({ id: lesson, t: teacher, o, minutes: k.c.minutes });
        const here = [];
        for (const p of places) {
          if (p.on > o || (p.left !== null && o >= p.left)) continue;
          const x = r.next();
          const m = x > p.s.attend ? 'Absent' : x > p.s.attend - 0.03 ? 'Late' : 'Present';
          add('Attendance', { Lesson: lesson, Student: p.s.id, Mark: m, Reason: m === 'Absent' ? absentWhy(p.s) : m === 'Late' && r.chance(0.3) ? r.pick(['Traffic', 'Bus late', 'Came from a school club']) : null });
          p.st.lessons++;
          p.st[m.toLowerCase()]++;
          if (m !== 'Absent') here.push(p);
        }
        if (tests.includes(i)) {
          const outOf = OUT_OF[subject];
          const test = add('Assessments', { Class: classId, 'Assessment Date': cal.day(o), Title: topic, 'Out Of': outOf });
          for (const p of here) {
            const s = p.s;
            s.ability[subject] ??= 0.45 + r.next() * 0.45;
            const score = Math.round(outOf * Math.min(1, Math.max(0.15, s.ability[subject] + (r.next() - 0.5) * 0.2)));
            const comment = r.chance(0.25) ? r.pick(COMMENTS) : null;
            add('Results', { Assessment: test, Student: s.id, Score: score, Comment: comment });
            // The percentage as the queries round it.
            p.st[i === tests[tests.length - 1] ? 'end' : 'half'] = Math.round((score * 100) / outOf);
            if (comment) p.st.comment = comment;
          }
        }
      });
    }
    // A term of lessons moves everyone on a little.
    if (b.kind === 'Term') for (const s of pool) {
      for (const k of Object.keys(s.ability)) s.ability[k] = Math.min(0.97, s.ability[k] + 0.01 + r.next() * 0.02);
      if (s.subjects?.includes('Coding')) s.coded++;
    }

    // One invoice a family a block, sent three weeks before it starts or when the last place was booked.
    for (const [f, bill] of billed) {
      const on = Math.max(b.start - 21, ...bill.map((x) => x.on));
      if (on > 0) continue;
      invoices.push({ f, termId, on, due: Math.max(b.start, on + 7), amount: money(bill.reduce((n, x) => n + x.amount, 0)), enrolments: bill.map((x) => x.enrolment) });
    }
  }

  // Invoices and what has come in against them: most on time, some in three parts, some late.
  // Numbered in date order, a new run each year: INV-2026-0042; receipts the same: RCP-2026-0042.
  const numbered = (prefix) => { const n = {}; return (o) => { const y = cal.at(o).getFullYear(); n[y] = (n[y] || 0) + 1; return `${prefix}-${y}-${String(n[y]).padStart(4, '0')}`; }; };
  const invoiceNumber = numbered('INV');
  const received = [];
  invoices.sort((x, y) => x.on - y.on);
  for (const inv of invoices) {
    const plan = [];
    const { amount, on, due } = inv;
    if (inv.f.pays === 'prompt') plan.push([on + r.int(0, due - on), amount]);
    else if (inv.f.pays === 'instalments') {
      const part = money(amount / 3);
      plan.push([due - r.int(0, 3), part], [due + 28, part], [due + 56, money(amount - part * 2)]);
    } else if (inv.f.pays === 'late') plan.push([due + r.int(8, 60), amount]);
    else {
      const part = money(amount / 2);
      // Half paid late; the rest later still, or not at all from a family who left owing.
      plan.push([due + r.int(10, 40), part]);
      if (r.chance(0.7)) plan.push([due + r.int(60, 240), money(amount - part)]);
    }
    let paid = 0;
    const payments = plan.filter(([o, x]) => o <= 0 && x > 0);
    for (const [, x] of payments) paid = money(paid + x);
    const id = add('Fee Invoices', { Number: invoiceNumber(on), Family: inv.f.id, Term: inv.termId, 'Invoice Date': cal.day(on), 'Due Date': cal.day(due), Amount: amount, 'Paid So Far': paid });
    // Each place is a line of its family's invoice.
    for (const e of inv.enrolments) base.update('Enrolments', e, { Invoice: id });
    let left = amount;
    for (const [o, x] of payments.sort((p, q) => p[0] - q[0])) {
      left = money(left - x);
      received.push({ id, o: Math.max(o, on), amount: x, left, method: r.chance(0.9) ? inv.f.method : r.pick(['Card', 'Bank transfer']) });
    }
  }
  const receiptNumber = numbered('RCP');
  received.sort((x, y) => x.o - y.o || x.id - y.id);
  for (const p of received) add('Payments', { Number: receiptNumber(p.o), Invoice: p.id, 'Paid On': cal.day(p.o), Amount: p.amount, Method: p.method, 'Balance After': p.left });

  // Term reports, written in the week after a term ends: a line for each class, from its register and tests.
  const reports = new Map();
  for (const x of reportLines) {
    if (!x.st.lessons) continue;
    const key = `${x.termId}|${x.s.id}`;
    if (!reports.has(key)) reports.set(key, { ...x, lines: [] });
    reports.get(key).lines.push(x);
  }
  for (const rep of reports.values()) {
    const { s, b, a } = rep;
    const lessons = rep.lines.reduce((n, x) => n + x.st.lessons, 0);
    const absent = rep.lines.reduce((n, x) => n + x.st.absent, 0);
    const marks = rep.lines.flatMap((x) => [x.st.half, x.st.end]).filter((v) => v !== null);
    const mean = marks.length ? marks.reduce((n, v) => n + v, 0) / marks.length : null;
    const name = s.first;
    const overall = [
      mean === null ? `${name} missed the tests this term, so there are no marks to report.` : mean >= 75 ? `${name} has had an excellent term and is working above the level of the group.` : mean >= 55 ? `${name} has made steady progress this term.` : `${name} has found the work hard this term and would gain from more practice at home.`,
      absent / lessons > 0.15 ? `Attendance has been patchy (${lessons - absent} of ${lessons} lessons), and coming every week would make a real difference.` : absent === 0 ? 'Full attendance this term: thank you.' : 'Attendance has been good.',
    ].join(' ');
    const id = add('Term Reports', { Student: s.id, Term: rep.termId, 'School Year': s.adult ? 'Adult' : `Year ${a - s.cohort - 5}`, 'Written On': cal.day(Math.min(b.end + r.int(2, 6), 0)), Comment: overall });
    for (const x of rep.lines) {
      const m = x.st.end ?? x.st.half;
      const band = m === null ? 'mid' : m >= 75 ? 'high' : m >= 55 ? 'mid' : 'low';
      add('Report Lines', {
        Report: id, Class: x.classId, Lessons: x.st.lessons, Present: x.st.present, Late: x.st.late, Absent: x.st.absent,
        'Half-term %': x.st.half, 'End of Term %': x.st.end, Comment: x.st.comment || r.pick(REPORT_COMMENTS[band]),
      });
    }
  }

  // Certificates for adult courses and holiday workshops finished with three lessons in four attended.
  const certificateNumber = numbered('CRT');
  const certificates = finished
    .filter((x) => x.st.lessons && (x.st.present + x.st.late) / x.st.lessons >= 0.75)
    .map((x) => ({ ...x, on: Math.min(x.b.end + r.int(3, 10), 0) }))
    .sort((x, y) => x.on - y.on || x.enrolment - y.enrolment);
  for (const x of certificates) {
    const attended = x.st.present + x.st.late;
    add('Certificates', { Number: certificateNumber(x.on), Enrolment: x.enrolment, 'Issued On': cal.day(x.on), 'Lessons Attended': attended, Hours: Math.round((attended * x.k.c.minutes) / 6) / 10 });
  }

  // Pay statements: one for each teacher for each month that has ended, the lessons taught its lines,
  // paid on the seventh of the month after (or the Monday after, at a weekend).
  const thisMonth = cal.day(0).slice(0, 7);
  const months = new Map();
  for (const l of taught) {
    const month = cal.day(l.o).slice(0, 7);
    if (month >= thisMonth) continue;
    const key = `${month}|${l.t.id}`;
    if (!months.has(key)) months.set(key, { month, t: l.t, lessons: [] });
    months.get(key).lessons.push(l);
  }
  const payNumber = numbered('PAY');
  for (const m of [...months.values()].sort((x, y) => (x.month < y.month ? -1 : x.month > y.month ? 1 : x.t.id - y.t.id))) {
    const [y, mo] = m.month.split('-').map(Number);
    const first = off(y, mo - 1, 1);
    const last = off(y, mo, 1) - 1;
    const paid = cal.workday(off(y, mo, 7));
    const hours = m.lessons.reduce((n, l) => n + l.minutes, 0) / 60;
    const gross = money(m.lessons.reduce((n, l) => n + money((l.minutes / 60) * m.t.rate), 0));
    const id = add('Pay Statements', {
      Number: payNumber(last), Teacher: m.t.id, Month: m.month, 'Period Start': cal.day(first), 'Period End': cal.day(last),
      'Paid On': paid <= 0 ? cal.day(paid) : null, 'Hourly Rate': m.t.rate, Lessons: m.lessons.length, Hours: Math.round(hours * 100) / 100, Gross: gross,
    });
    for (const l of m.lessons) base.update('Lessons', l.id, { 'Pay Statement': id });
  }

  // This year's school year for each child, and who is still coming.
  for (const s of students) {
    if (!s.id) continue;
    const year = s.adult ? null : now - s.cohort - 5;
    base.update('Students', s.id, { 'School Year': s.adult ? 'Adult' : `Year ${year}`, Active: !s.withdrew && s.last_on >= -60 });
  }
  return counts;
}
