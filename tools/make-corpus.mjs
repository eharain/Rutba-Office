// Make the starter corpus in tests/corpus-public: a few files of every format
// the corpus run reads, each one either written by this suite's own writers
// from words this project wrote, or copied from the repository's own test
// fixtures, which this project made (some by driving Microsoft Office over
// content written here, see tools/make-binary-fixtures.ps1 and
// tools/make-rich-fixtures.ps1). Nothing is downloaded, and nothing comes from
// anyone else, so every file is under the repository's own licence. Each file's
// source and licence are recorded in tests/corpus-public/SOURCES.json, which
// tests/corpus.test.js holds the folder to.
//
//   node tools/make-corpus.mjs
//
// The real-world corpus the release criteria speak of is gathered apart (its
// files belong to other people and never enter the repository, see .corpus/
// in .gitignore); this one is the start the run is proven on, and what CI can
// run without anyone's files.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { buildPptx, TEMPLATES } from '@rutba/presentation';
import pdf from '@rutba/pdf';
import { toRfc822 } from '@rutba/mailbox/rfc822';
import { writeCalendar, dateTimeAt } from '@rutba/calendar/ical';
import { writeVCards } from '@rutba/contacts/vcard';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { buildComplexDocument } from '../tests/fixtures/complex-document.js';
import { buildComplexWorkbook } from '../tests/fixtures/complex-workbook.js';
import { buildCoverPageDocument } from '../tests/fixtures/cover-page-document.js';
import { withEquations, QUADRATIC } from '../tests/fixtures/equations.js';
import { writePst } from '../tests/fixtures/pst-writer.js';
import { PST_MESSAGES } from '../tests/fixtures/mail-archives.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = path.join(root, 'tests', 'corpus-public');
const LICENCE = 'AGPL-3.0-or-later';
const OWNER = 'Rutba Office project (Tech Style Ltd)';

const sources = [];
const put = (rel, bytes, source) => {
  const file = path.join(out, rel);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, bytes);
  sources.push({ file: rel, source, licence: LICENCE, author: OWNER });
  return file;
};
const copy = (rel, from, how) => put(rel, fs.readFileSync(path.join(root, from)), `copied from ${from}; ${how}`);

// Starting again: a file left from an older list would be one with no source.
for (const entry of fs.existsSync(out) ? fs.readdirSync(out) : []) {
  if (entry !== 'README.md') fs.rmSync(path.join(out, entry), { recursive: true, force: true });
}

/* ── words ─────────────────────────────────────────────────────────────── */

const LETTER = [
  'Dear Ms Rahman,',
  'Thank you for your order of 14 March. The goods described in Schedule 2 will be delivered within thirty days, to the address on the order.',
  'The invoice follows by post. Payment is due within thirty days of the invoice date.',
  'Yours sincerely,',
  'Bo Lindqvist',
];
const MULTILINGUAL = [
  'Rutba Office reads and writes documents in many scripts.',
  'یہ اردو میں ایک جملہ ہے۔',
  'هذه جملة باللغة العربية.',
  'यह हिंदी में एक वाक्य है।',
  'Ceci est une phrase en français, avec des accents: é, è, ç.',
];

/* ── documents, workbooks, decks: the suite's own writers ─────────────── */

const docx = {
  'docx/letter.docx': [buildDocx({ paragraphs: LETTER, styles: true }), 'written by @rutba/ooxml buildDocx (tools/make-corpus.mjs)'],
  'docx/multilingual.docx': [buildDocx({ paragraphs: MULTILINGUAL, styles: true }), 'written by @rutba/ooxml buildDocx (tools/make-corpus.mjs)'],
  'docx/complex-parts.docx': [buildComplexDocument(), 'written by tests/fixtures/complex-document.js (styles, numbering, header, footer, footnotes, comments, a content control, a field)'],
  'docx/cover-page.docx': [buildCoverPageDocument(), 'written by tests/fixtures/cover-page-document.js (a cover page and a table of contents in content controls)'],
  'docx/equation.docx': [withEquations(buildDocx({ paragraphs: ['The roots of a quadratic:', 'EQUATION', 'as every school book gives them.'] }), [QUADRATIC]), 'written by @rutba/ooxml buildDocx with tests/fixtures/equations.js (an Office Math equation)'],
};
const xlsx = {
  'xlsx/budget.xlsx': [
    buildXlsx({
      sheets: [
        { name: 'Budget', rows: [['Item', 'Q1', 'Q2', 'Total'], ['Rent', 1200, 1200, '=SUM(B2:C2)'], ['Salaries', 8400, 8650.5, '=SUM(B3:C3)'], ['Travel', 310.25, 0, '=SUM(B4:C4)'], ['Total', '=SUM(B2:B4)', '=SUM(C2:C4)', '=SUM(D2:D4)']] },
        { name: 'Notes', rows: [['Figures in pounds.'], ['Q2 travel was cancelled.']] },
      ],
      definedNames: [{ name: 'GrandTotal', ref: 'Budget!$D$5' }],
    }),
    'written by @rutba/ooxml buildXlsx (tools/make-corpus.mjs)',
  ],
  'xlsx/multilingual.xlsx': [buildXlsx({ sheets: [{ name: 'Words', rows: MULTILINGUAL.map((t, i) => [i + 1, t]) }] }), 'written by @rutba/ooxml buildXlsx (tools/make-corpus.mjs)'],
  'xlsx/complex-parts.xlsx': [buildComplexWorkbook({ protect: false }), 'written by tests/fixtures/complex-workbook.js (a chart, a pivot cache, merged cells, conditional formats, validation)'],
};
const pptx = {
  'pptx/agenda.pptx': [
    buildPptx({ title: 'Quarterly review', slides: [{ layout: 'title', title: 'Quarterly review', body: 'Northwind, second quarter' }, { layout: 'obj', title: 'Agenda', body: ['Results', 'Hiring', 'Next quarter'] }, { layout: 'obj', title: 'Results', body: ['Revenue up eight per cent', 'Two new customers', MULTILINGUAL[1]] }] }),
    'written by @rutba/presentation buildPptx (tools/make-corpus.mjs)',
  ],
  'pptx/pitch-template.pptx': [TEMPLATES.pitch(), 'written by @rutba/presentation TEMPLATES.pitch, the suite\'s own pitch deck template'],
};
for (const [rel, [bytes, how]] of Object.entries({ ...docx, ...xlsx, ...pptx })) put(rel, Buffer.from(bytes), how);

/* ── made by this project with Microsoft Office, in the repository already ── */

const OFFICE_RICH = 'made by this project with Microsoft Office from content written here (tools/make-rich-fixtures.ps1)';
const OFFICE_BINARY = 'made by this project with Microsoft Office from content written here (tools/make-binary-fixtures.ps1)';
copy('docx/office-showcase.docx', 'tests/fixtures/rich/showcase.docx', OFFICE_RICH);
copy('docx/password-protected.docx', 'tests/fixtures/encrypted.docx', 'a document encrypted with a password by this project\'s test suite; the corpus run must ask for the password, not fail');
copy('xlsx/office-showcase.xlsx', 'tests/fixtures/rich/showcase.xlsx', OFFICE_RICH);
copy('pptx/office-showcase.pptx', 'tests/fixtures/rich/showcase.pptx', OFFICE_RICH);
copy('ods/office-showcase.ods', 'tests/fixtures/rich/showcase.ods', OFFICE_RICH);
copy('odp/office-showcase.odp', 'tests/fixtures/rich/showcase.odp', OFFICE_RICH);
for (const name of ['text', 'paragraphs', 'structure', 'floats', 'locked']) copy(`doc/${name}.doc`, `tests/fixtures/binary/${name}.doc`, OFFICE_BINARY);
for (const name of ['conditions', 'charts', 'showcase', 'charts-95']) copy(`xls/${name}.xls`, `tests/fixtures/binary/${name}.xls`, OFFICE_BINARY);
copy('ppt/tables.ppt', 'tests/fixtures/binary/tables.ppt', OFFICE_BINARY);

/* ── OpenDocument and PDF, written by the suite's own Save As and Export ── */

const service = createDocumentService({ holdBlob: () => ({ url: 'blob://held' }) });
const exportOf = (from, rel, format, how) => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-make-corpus-'));
  try {
    const opened = service.open({ path: path.join(out, from) });
    const target = path.join(tmp, path.basename(rel));
    service.export({ id: opened.id, format, path: target });
    service.close({ id: opened.id });
    put(rel, fs.readFileSync(target), `${how} from ${from} (tools/make-corpus.mjs)`);
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
};
exportOf('docx/letter.docx', 'odt/letter.odt', 'odt', 'written by the suite\'s OpenDocument export');
exportOf('docx/complex-parts.docx', 'odt/complex-parts.odt', 'odt', 'written by the suite\'s OpenDocument export');
exportOf('xlsx/budget.xlsx', 'ods/budget.ods', 'ods', 'written by the suite\'s OpenDocument export');
exportOf('pptx/agenda.pptx', 'odp/agenda.odp', 'odp', 'written by the suite\'s OpenDocument export');
exportOf('docx/letter.docx', 'pdf/letter.pdf', 'pdf', 'written by the suite\'s PDF export');
exportOf('docx/complex-parts.docx', 'pdf/complex-parts.pdf', 'pdf', 'written by the suite\'s PDF export');

/* ── the macro-enabled and template forms, written by the suite's own Save As ── */

const saveAs = (from, rel) => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-make-corpus-'));
  try {
    const opened = service.open({ path: path.join(out, from) });
    const target = path.join(tmp, path.basename(rel));
    service.save({ id: opened.id, path: target });
    service.close({ id: opened.id });
    put(rel, fs.readFileSync(target), `written by the suite's Save As from ${from} (tools/make-corpus.mjs)`);
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
};
saveAs('docx/letter.docx', 'docm/letter.docm');
saveAs('docx/letter.docx', 'dotx/letter.dotx');
saveAs('xlsx/budget.xlsx', 'xlsm/budget.xlsm');
saveAs('xlsx/budget.xlsx', 'xltx/budget.xltx');
saveAs('pptx/agenda.pptx', 'pptm/agenda.pptm');
saveAs('pptx/agenda.pptx', 'potx/agenda.potx');
{
  const doc = new pdf.PdfDocument({ title: 'Delivery note', author: 'Northwind' });
  for (let n = 1; n <= 3; n++) {
    const page = doc.addPage();
    page.text(`Delivery note, page ${n} of 3`, 56, 780, { font: 'Helvetica-Bold', size: 16 });
    page.text('Twelve boxes of A4 paper, delivered to the front desk.', 56, 750, { size: 11 });
    page.rect(56, 600, 480, 120, { stroke: '#333333' });
  }
  put('pdf/delivery-note.pdf', doc.toBuffer(), 'written by @rutba/pdf PdfDocument (tools/make-corpus.mjs)');
}

/* ── mail, calendar, contacts ─────────────────────────────────────────── */

const amina = { name: 'Amina Rahman', address: 'amina@northwind.example' };
const bo = { name: 'Bo Lindqvist', address: 'bo@northwind.example' };
const dana = { name: 'Dana Okafor', address: 'dana@contoso.example' };
const mails = {
  'eml/plain.eml': { subject: 'Delivery on Thursday', from: [bo], to: [amina], date: '2026-03-12T09:30:00Z', messageId: 'plain-1@northwind.example', text: 'The paper arrives on Thursday morning.\n\nBo' },
  'eml/html-with-attachment.eml': {
    subject: 'Signed order',
    from: [amina],
    to: [bo],
    cc: [dana],
    date: '2026-03-14T11:05:00Z',
    messageId: 'order-2@northwind.example',
    text: 'The signed order is attached.',
    html: '<p>The <b>signed order</b> is attached.</p>',
    attachments: [{ filename: 'order.txt', type: 'text/plain', bytes: Buffer.from('Order 1042: twelve boxes of A4 paper.\n') }],
  },
  'eml/non-ascii.eml': { subject: 'Réunion: اجلاس', from: [{ name: 'Zoë Ångström', address: 'zoe@contoso.example' }], to: [amina], date: '2026-03-15T08:00:00Z', messageId: 'intl-3@contoso.example', text: MULTILINGUAL.join('\n') },
};
for (const [rel, m] of Object.entries(mails)) put(rel, toRfc822(m), 'written by @rutba/mailbox toRfc822 (tools/make-corpus.mjs)');

{
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-make-corpus-'));
  try {
    const file = path.join(tmp, 'northwind.pst');
    // The archive fixtures' messages, their large made-up attachments left
    // out and a small one of words in their place: the corpus stays small.
    const light = (list) => list.map((m) => ({ ...m, attachments: (m.attachments || []).slice(0, 1).map((a) => ({ filename: a.filename.replace(/\.\w+$/, '.txt'), type: 'text/plain', bytes: Buffer.from(`${m.subject}\n`) })) }));
    writePst(
      file,
      [
        { name: 'Inbox', messages: light(PST_MESSAGES.Inbox), children: [{ name: 'Projects', messages: light(PST_MESSAGES['Inbox/Projects']) }] },
        { name: 'Sent Items', messages: light(PST_MESSAGES['Sent Items']) },
      ],
      { name: 'Northwind' }
    );
    put('pst/northwind.pst', fs.readFileSync(file), 'written by tests/fixtures/pst-writer.js with the messages of tests/fixtures/mail-archives.js (attachments replaced by a line of text)');
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

const day = (iso) => Date.parse(iso);
put(
  'ics/meetings.ics',
  writeCalendar({
    name: 'Northwind',
    events: [
      { uid: 'standup@northwind.example', summary: 'Stand-up', start: dateTimeAt(day('2026-03-02T09:00:00Z'), 'Europe/London'), end: dateTimeAt(day('2026-03-02T09:15:00Z'), 'Europe/London'), rrule: { freq: 'WEEKLY', interval: 1, byDay: [{ n: 0, day: 1 }, { n: 0, day: 3 }, { n: 0, day: 5 }], count: 30 } },
      { uid: 'review@northwind.example', summary: 'Quarterly review', location: 'Room 4', description: 'Results, hiring, next quarter.', start: dateTimeAt(day('2026-03-20T14:00:00Z')), end: dateTimeAt(day('2026-03-20T15:30:00Z')) },
      { uid: 'holiday@northwind.example', summary: 'Office closed', allDay: true, start: dateTimeAt(day('2026-04-03T00:00:00Z'), null, true), end: dateTimeAt(day('2026-04-04T00:00:00Z'), null, true) },
    ],
  }),
  'written by @rutba/calendar writeCalendar (tools/make-corpus.mjs)'
);

const people = [
  { name: { full: 'Amina Rahman', given: 'Amina', family: 'Rahman' }, org: 'Northwind', title: 'Buyer', emails: [{ value: 'amina@northwind.example', type: 'work', pref: true }], phones: [{ value: '+44 20 7946 0001', type: 'work' }] },
  { name: { full: 'Bo Lindqvist', given: 'Bo', family: 'Lindqvist' }, org: 'Northwind', emails: [{ value: 'bo@northwind.example', type: 'work' }], phones: [{ value: '+46 8 555 0102', type: 'cell' }], addresses: [{ street: '1 Harbour Road', city: 'Stockholm', postcode: '111 22', country: 'Sweden', type: 'work' }] },
  { name: { full: 'زینب علی', given: 'زینب', family: 'علی' }, org: 'Contoso', emails: [{ value: 'zainab@contoso.example', type: 'work' }] },
];
put('vcf/people.vcf', writeVCards(people), 'written by @rutba/contacts writeVCards, vCard 3.0 (tools/make-corpus.mjs)');
put('vcf/people-v4.vcf', writeVCards(people, { version: '4.0' }), 'written by @rutba/contacts writeVCards, vCard 4.0 (tools/make-corpus.mjs)');

/* ── the record ────────────────────────────────────────────────────────── */

for (const s of sources) s.sha256 = crypto.createHash('sha256').update(fs.readFileSync(path.join(out, s.file))).digest('hex');
sources.sort((a, b) => a.file.localeCompare(b.file));
fs.writeFileSync(
  path.join(out, 'SOURCES.json'),
  JSON.stringify({ about: 'Where every file in this folder came from, and its licence. Made by tools/make-corpus.mjs; tests/corpus.test.js holds the folder to it.', files: sources }, null, 2) + '\n'
);
console.log(`${sources.length} files in ${path.relative(root, out)}`);
