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

import { FIRST, LAST, money, person, phoneFor } from './util.js';

const ref = (table) => ({ type: 'integer', points: table });
const FEE = 'Round([Fee per Lesson] * [Lessons Billed] * (1 - Nz([Discount])), 2)';
const STUDENT = '[Students].[First Name] & " " & [Students].[Last Name]';
const TEACHER = '[Teachers].[First Name] & " " & [Teachers].[Last Name]';
const OWED = 'Round([Amount] - Nz([Paid So Far]), 2)';
// Days from today to a birthday, counting each month as 31 days: near enough
// to find the birthdays in the coming month without a date built from parts.
const TO_BIRTHDAY = '(Month([Date of Birth]) * 31 + Day([Date of Birth]) - Month(Date()) * 31 - Day(Date()) + 372) Mod 372';

export const spec = {
  negative: {},
  tables: [
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
    ] },
    { name: 'Lessons', fields: [
      { name: 'Class', ...ref('Classes'), required: true }, { name: 'Teacher', ...ref('Teachers') }, { name: 'Lesson Date', type: 'date', required: true },
      { name: 'Minutes', type: 'integer' }, { name: 'Topic', type: 'text' }, { name: 'Status', type: 'text', default: 'Planned' },
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
      { name: 'Family', ...ref('Families'), required: true }, { name: 'Term', ...ref('Terms'), required: true },
      { name: 'Invoice Date', type: 'date' }, { name: 'Due Date', type: 'date' }, { name: 'Amount', type: 'currency' }, { name: 'Paid So Far', type: 'currency', default: 0 },
    ] },
    { name: 'Payments', fields: [
      { name: 'Invoice', ...ref('Fee Invoices'), required: true }, { name: 'Paid On', type: 'date' }, { name: 'Amount', type: 'currency' },
      { name: 'Method', type: 'text' },
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
      { table: 'Families', field: 'City' }, { table: 'Families', field: 'Postcode' }, { table: 'Fee Invoices', field: 'ID', as: 'Invoice' },
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
    { name: 'Fees due letters', title: 'Fees overdue', query: 'Overdue fees', group: 'Family', columns: ['Contact', 'Address', 'City', 'Postcode', { field: 'Invoice', type: 'integer' }, 'Term', { field: 'Due Date', type: 'date' }, { field: 'Days Overdue', type: 'integer' }, { field: 'Outstanding', type: 'currency' }], totals: [{ field: 'Outstanding', fn: 'sum' }] },
    { name: 'Progress report', query: 'Progress by student', group: 'Student', sort: 'Assessment Date', columns: ['Course', 'Assessment', { field: 'Assessment Date', type: 'date' }, { field: 'Score', type: 'integer' }, { field: 'Out Of', type: 'integer' }, { field: 'Percent', type: 'integer' }, 'Comment'], totals: [{ field: 'Percent', fn: 'avg' }] },
    { name: 'Teacher pay', query: 'Teacher hours last month', columns: ['Month', 'Teacher', { field: 'Hourly Rate', type: 'currency' }, { field: 'Lessons', type: 'integer' }, { field: 'Hours', type: 'number' }, { field: 'Pay', type: 'currency' }], totals: [{ field: 'Hours', fn: 'sum' }, { field: 'Pay', fn: 'sum' }] },
    { name: 'Attendance concerns', query: 'Attendance below 85 percent', columns: ['Student', 'School Year', 'Contact', 'Phone', { field: 'Lessons', type: 'integer' }, { field: 'Absences', type: 'integer' }, { field: 'Attendance %', type: 'number' }] },
    { name: 'Enrolments by course', query: 'Enrolments by course and term', group: 'Course', sort: 'Starts', columns: ['Term', { field: 'Starts', type: 'date' }, { field: 'Enrolments', type: 'integer' }, { field: 'Fees', type: 'currency' }], totals: [{ field: 'Enrolments', fn: 'sum' }, { field: 'Fees', fn: 'sum' }] },
    { name: 'Fees received', query: 'Fees received by month', columns: ['Month', { field: 'Payments', type: 'integer' }, { field: 'Received', type: 'currency' }], totals: [{ field: 'Payments', fn: 'sum' }, { field: 'Received', fn: 'sum' }] },
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
    } },
    'Lessons': { about: 'Every lesson of every class: those taught, those planned for the rest of the term and the next, and the few cancelled.', fields: {
      'Class': 'The class the lesson is for.',
      'Teacher': 'Who took the lesson, which is the class\'s teacher unless someone covered. Teacher pay is worked out from it.',
      'Lesson Date': 'The day of the lesson.',
      'Minutes': 'How long it lasted, which is what the teacher is paid for.',
      'Topic': 'What the lesson covered, or is planned to cover.',
      'Status': 'Taught, Planned or Cancelled. Only taught lessons are paid.',
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
    'Fee Invoices': { about: 'One invoice for each family for each term or holiday they book, covering every class their students take in it.', fields: {
      'Family': 'Who the invoice is for.',
      'Term': 'The term or holiday it is for.',
      'Invoice Date': 'The day it was sent, three weeks before the term starts or when a late place was booked.',
      'Due Date': 'When it should be paid: by the first day of the term.',
      'Amount': 'The total of the family\'s fees for the term, after discounts.',
      'Paid So Far': 'What has been paid against it. Keep it equal to the total of its payments; what is left is outstanding.',
    } },
    'Payments': { about: 'Each payment received against an invoice. A family paying in instalments has several.', fields: {
      'Invoice': 'The invoice the payment is for.',
      'Paid On': 'The day the money arrived.',
      'Amount': 'How much was paid.',
      'Method': 'Bank transfer, Card or Cash.',
    } },
  },
  queries: {
    'Class lists': 'Every student in every class that is running now or about to start, with the class\'s day, time, room and teacher, any medical note, photo consent and the family\'s phone number. Print it as the register before a lesson.',
    'Timetable, next 7 days': 'Every lesson from today to a week ahead, in date and time order, with the room, the teacher and the planned topic. Use it to plan rooms and cover.',
    'Attendance below 85 percent': 'Students who have missed more than fifteen in a hundred of their lessons in the last 90 days, with at least four lessons in that time, and the family to call.',
    'Fees outstanding by family': 'Each family that still owes something, with how many invoices are open, the oldest due date and the total left to pay, largest first.',
    'Overdue fees': 'Each invoice past its due date that is not fully paid, with the family\'s address, the term, how many days late it is and what is left to pay. The fees letters are written from it.',
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
  empty: 'Start with Terms (this year\'s terms and holidays), then Teachers and Courses (what you teach, how long and at what price). Then set up the Classes for the first term, each with its course, teacher, day, time and room. Enter Families and then their Students as they book, and an Enrolment for each place. Add the Lessons of each class (one for each week of the term), take the register in Attendance as the lessons happen, and record Assessments and Results when tests are marked. Raise a Fee Invoice for each family each term and enter Payments as they come in, adding each to the invoice\'s Paid So Far.',
  adapt: [
    'For one-to-one tutoring, set each course\'s Capacity to 1 and make a class for each student\'s regular slot; the timetable and teacher pay work as they are.',
    'For a music or dance school, use the subjects for instruments or dance styles and the levels for grades; add a Grade field to Results for exam grades.',
    'If you charge by the month instead of by the term, make each month a Term record and raise the invoices monthly.',
    'Add a Waiting status to Enrolments for full classes, and a query listing waiting students by class, oldest first.',
    'For a driving school, make each learner\'s block of lessons a class with Capacity 1 and the instructor as the teacher; add a Test Date field to Students.',
    'To keep a fuller register, add a Minutes Late field to Attendance, or a Homework Done yes/no field.',
    'If teachers are paid a fixed amount a lesson instead of by the hour, add a Rate per Lesson field to Teachers and change the Pay column of Teacher hours last month to use it.',
    'Add a Rooms table with each room\'s size if you have many rooms, and point Classes at it instead of typing the room.',
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
const STREETS = ['Harehills Lane', 'Mill Lane', 'Station Road', 'Orchard Way', 'Church Street', 'Canal Walk', 'Beech Grove', 'Riverside Drive', 'Quarry Road', 'Cedar Avenue', 'Linden Close', 'Moor View', 'Park Terrace'];

/** Fill a Lantern Learning Centre database made from `spec`. `r` is random(), `cal` calendar(). Answers counts. */
export function fill(base, { r, cal }) {
  const counts = {};
  const add = (table, values) => { counts[table] = (counts[table] || 0) + 1; return base.insert(table, values); };

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
    return { id, subjects: subjects.split(', '), load: 0 };
  });

  // Courses.
  const courses = COURSES.map(([code, name, subject, level, audience, minutes, fee, capacity, years, days]) => {
    const id = add('Courses', { Code: code, Name: name, Subject: subject, Level: level, For: audience, 'Holiday Course': Boolean(days), Minutes: minutes, 'Fee per Lesson': fee, Capacity: capacity });
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
        add('Enrolments', {
          Student: save(s, p.on), Class: classId, 'Enrolled On': cal.day(p.on), Status: gone ? 'Withdrawn' : ended ? 'Completed' : 'Enrolled',
          'Lessons Billed': lessons, Discount: discount, 'Left On': gone ? cal.day(left) : null,
        });
        s.last_on = Math.max(s.last_on, gone ? left : b.end);
        const bill = billed.get(s.family) || [];
        bill.push({ amount: money(k.c.fee * lessons * (1 - discount)), on: p.on });
        billed.set(s.family, bill);
        places.push({ s, on: p.on, left: gone ? left : null });
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
        const here = [];
        for (const p of places) {
          if (p.on > o || (p.left !== null && o >= p.left)) continue;
          const x = r.next();
          const m = x > p.s.attend ? 'Absent' : x > p.s.attend - 0.03 ? 'Late' : 'Present';
          add('Attendance', { Lesson: lesson, Student: p.s.id, Mark: m, Reason: m === 'Absent' ? absentWhy(p.s) : m === 'Late' && r.chance(0.3) ? r.pick(['Traffic', 'Bus late', 'Came from a school club']) : null });
          if (m !== 'Absent') here.push(p.s);
        }
        if (tests.includes(i)) {
          const outOf = OUT_OF[subject];
          const test = add('Assessments', { Class: classId, 'Assessment Date': cal.day(o), Title: topic, 'Out Of': outOf });
          for (const s of here) {
            s.ability[subject] ??= 0.45 + r.next() * 0.45;
            const score = Math.round(outOf * Math.min(1, Math.max(0.15, s.ability[subject] + (r.next() - 0.5) * 0.2)));
            add('Results', { Assessment: test, Student: s.id, Score: score, Comment: r.chance(0.25) ? r.pick(COMMENTS) : null });
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
      invoices.push({ f, termId, on, due: Math.max(b.start, on + 7), amount: money(bill.reduce((n, x) => n + x.amount, 0)) });
    }
  }

  // Invoices and what has come in against them: most on time, some in three parts, some late.
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
    const id = add('Fee Invoices', { Family: inv.f.id, Term: inv.termId, 'Invoice Date': cal.day(on), 'Due Date': cal.day(due), Amount: amount, 'Paid So Far': paid });
    for (const [o, x] of payments) add('Payments', { Invoice: id, 'Paid On': cal.day(Math.max(o, on)), Amount: x, Method: r.chance(0.9) ? inv.f.method : r.pick(['Card', 'Bank transfer']) });
  }

  // This year's school year for each child, and who is still coming.
  for (const s of students) {
    if (!s.id) continue;
    const year = s.adult ? null : now - s.cohort - 5;
    base.update('Students', s.id, { 'School Year': s.adult ? 'Adult' : `Year ${year}`, Active: !s.withdrew && s.last_on >= -60 });
  }
  return counts;
}
