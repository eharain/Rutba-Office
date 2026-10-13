// Screenshots of the real application, for the website.
//
//   node build/make-screens.js [outDir]
//   npm run screens
//
// `npm run smoke` already boots every window and photographs it — that is what
// this borrows. What it adds is the two things that stop a smoke capture being
// usable as a picture of a product:
//
//   1. **Documents worth photographing.** A smoke run opens each app empty, so
//      the spreadsheet screenshot is an empty grid and the document screenshot
//      is a blank page. That proves the window paints and sells nothing. This
//      builds a small, realistic file for each of the three document apps
//      first, and hands it to that app's capture.
//
//   2. **A profile of its own.** The smoke run writes into the real user
//      profile, so its seeded mail account accumulates: after eleven runs the
//      mail screenshot showed eleven identical accounts in the sidebar. Each
//      capture here runs against a fresh `--user-data-dir` in a temp
//      directory, which is also the only honest way to photograph a mail
//      client — nobody's actual inbox belongs in a marketing asset.
//
// The content is invented and says so where it can. The numbers are made up,
// the correspondents are fictional, and the one thing that is real is the
// software drawing them.
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
// The same writer the smoke run draws its pictures with.
import { gradientPng } from '../main/sample-picture.js';
import electron from 'electron';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
// The document sample gets a chart the way a person adds one: the caret in
// the table, Insert → Chart. The engine draws it from the table's figures.
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { buildPptx, Deck } from '@rutba/presentation';
import { dateTimeAt, writeCalendar } from '@rutba/calendar/ical';
import { writeVCards } from '@rutba/contacts/vcard';
import { Sheet, INK } from '@rutba/pdf';
import { emptyProject, addMedia, setTransition, addTitle, addMarker } from '@rutba/studio/project';
import { writeRvid } from '@rutba/studio/rvid';

const appDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outDir = path.resolve(process.argv[2] || path.join(appDir, 'build', 'screens'));

const work = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-screens-'));
// The status bar shows this path, so it is named rather than random: a
// screenshot whose footer reads `rutba-screens-JAsyZI\docs` reads as a
// test artefact, which is what it is and not what it is for.
const docsDir = path.join(os.tmpdir(), 'Rutba Office Samples');
fs.rmSync(docsDir, { recursive: true, force: true });
fs.mkdirSync(docsDir, { recursive: true });
fs.mkdirSync(outDir, { recursive: true });

const at = (name) => path.join(docsDir, name);

/**
 * A workbook with the two things a spreadsheet screenshot has to show: money
 * formatted as money, and a total that is a formula rather than a typed
 * number. Everything else is restraint — a screenshot of forty columns reads
 * as noise at the size a web page renders it.
 */
{
  const regions = [
    ['North', 142000, 161000, 158500, 174200],
    ['South', 86000, 91000, 94750, 98300],
    ['East', 51200, 57400, 60100, 66800],
    ['West', 73900, 71100, 78650, 82400],
    ['Midlands', 64300, 66900, 70250, 73100],
    ['London', 188400, 196200, 201750, 214900],
    ['Scotland', 38700, 41200, 42900, 45600],
    ['Wales', 22100, 23400, 24850, 26300],
  ];
  const last = regions.length + 1; // the row number of the last region
  const total = last + 1;
  // A summary block of its own beside the table — the year by region, in
  // I1:J9 with a blank column between — which is what the chart is drawn
  // from: charted from the table itself, the Year, Share and Total figures
  // would dwarf every quarter. A blank cell is a hole in the row, not an
  // empty string, so the block and the table stay two regions.
  const rows = [
    ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Year', 'Share', null, 'Region', 'Year'],
    ...regions.map((r, i) => [...r, `=SUM(B${i + 2}:E${i + 2})`, `=F${i + 2}/F$${total}`, null, r[0], `=F${i + 2}`]),
    ['Total', ...['B', 'C', 'D', 'E', 'F'].map((c) => `=SUM(${c}2:${c}${last})`), `=SUM(G2:G${last})`],
    [],
    ['Growth, Q4 on Q1', `=E${total}/B${total}-1`],
    ['Best quarter', `=MAX(B${total}:E${total})`],
  ];
  const book = SheetView.open(
    buildXlsx({
      sheets: [
        {
          name: 'Sales',
          rows,
          styles: {
            'A1:G1': { bold: true, fill: '#1F3A5F', colour: '#FFFFFF', border: true },
            [`A2:A${last}`]: { bold: true },
            [`B2:F${last}`]: { numFmt: '#,##0' },
            [`G2:G${last}`]: { numFmt: '0.0%' },
            // Declared after the number format and repeated for the numeric
            // half, because the last matching declaration wins: a total row
            // styled as a block silently dropped the format from its own
            // figures and printed 353100 beside 142,000.
            [`A${total}:G${total}`]: { bold: true, fill: '#EEF2F7', border: true },
            [`B${total}:F${total}`]: { bold: true, fill: '#EEF2F7', border: true, numFmt: '#,##0' },
            [`G${total}:G${total}`]: { bold: true, fill: '#EEF2F7', border: true, numFmt: '0.0%' },
            [`A${total + 2}:A${total + 3}`]: { bold: true },
            [`B${total + 2}:B${total + 2}`]: { numFmt: '0.0%', bold: true, colour: '#0E7C66' },
            [`B${total + 3}:B${total + 3}`]: { numFmt: '#,##0', bold: true },
            'I1:J1': { bold: true, fill: '#EEF2F7', border: true },
            [`I2:I${last}`]: { bold: true },
            [`J2:J${last}`]: { numFmt: '#,##0' },
          },
        },
      ],
    }),
    { viewportWidth: 1100, viewportHeight: 620 },
  );
  // A first column wide enough for its own labels, and a chart from the
  // figures — the way a person adds one: the cursor on the table, Insert →
  // Chart. Saved from the view, so both are in the file as Excel writes them.
  book.setColWidth(0, 132);
  book.select(0, 8);
  book.insertChart({ kind: 'column', title: 'Revenue by region, full year' });
  book.select(0, 0);
  fs.writeFileSync(at('regional-sales.xlsx'), book.save());
}

/**
 * A document with a heading hierarchy, because a screenshot of one paragraph
 * of Lorem Ipsum tells a reader nothing about a word processor — and with
 * the things the page has learned to draw since: a table at the widths its
 * file gives it, a chart drawn from that table, and enough of a report that
 * it runs onto a second page, so the status bar can say so. The ruler shows
 * above the page by default.
 */
{
  const view = openDocx(
    buildDocx({
      styles: true,
      paragraphs: [
        { text: 'Quarterly Review', style: 'Title', bold: true, size: 30 },
        { text: 'Prepared for the board — figures are illustrative', size: 20 },
        { text: 'Summary', style: 'Heading1', bold: true, size: 26 },
        {
          text:
            'The northern region grew fastest in absolute terms, adding £32,200 between the first and fourth quarters. Growth was steadier in the east, which finished the year 30% ahead of where it started without a single down quarter.',
        },
        {
          text:
            'Churn fell to 2.1% across the year, from 3.4%. Most of the improvement came in the second half and is attributable to the change in onboarding rather than to pricing.',
        },
        { text: 'Revenue by region', style: 'Heading1', bold: true, size: 26 },
        {
          table: {
            rows: [
              ['Region', 'Q1', 'Q2', 'Q3', 'Q4'],
              ['North', '142,000', '161,000', '158,500', '174,200'],
              ['South', '86,000', '91,000', '94,750', '98,300'],
              ['East', '51,200', '57,400', '60,100', '66,800'],
              ['West', '73,900', '71,100', '78,650', '82,400'],
            ],
            header: true,
            align: [null, 'right', 'right', 'right', 'right'],
            // Twips; together they are the text width of an A4 page with
            // the default margins, so the table sits flush with the prose.
            columns: [2626, 1600, 1600, 1600, 1600],
          },
        },
        {
          text:
            'The figures are quarterly revenue in pounds, net of refunds. The chart below is drawn from the table above and updates with it.',
        },
        { text: 'What we are watching', style: 'Heading1', bold: true, size: 26 },
        {
          text:
            'Two things. The west was flat in the second quarter and has not fully recovered its trend, and the cost of servicing the largest three accounts has risen faster than the revenue they bring.',
        },
        {
          text:
            'Neither is urgent. Both would become urgent if the pattern repeats in the first quarter of next year, which is when we should look again.',
        },
        { text: 'Costs', style: 'Heading1', bold: true, size: 26 },
        {
          text:
            'Payroll rose 6% with two hires in the east, both of which were planned. Hosting fell 11% after the move to reserved capacity in the second quarter, which paid for itself in the third.',
        },
        {
          text:
            'Travel is the one line above plan: the two large renewals in the south were won in person, and the visits are what won them.',
        },
        { text: 'Next quarter', style: 'Heading1', bold: true, size: 26 },
        {
          text:
            'Pricing stays as it is. Onboarding gets the same treatment the south received, region by region, starting with the west, and the three largest accounts get a named owner each.',
        },
      ],
    }),
  );
  // The caret in the table's first cell; the chart goes in after the table.
  const cell = view.render({ pages: false }).blocks.find((b) => /^t\d+:r0:c0$/.test(b.container || ''));
  view.setSelection({ block: cell.index, offset: 0 });
  view.insertChart({ kind: 'column', title: 'Revenue by region' });
  fs.writeFileSync(at('quarterly-review.docx'), view.save());
}

/**
 * Three slides, so the thumbnail rail in the screenshot is not one item —
 * and on the first, three figure cards: rounded rectangles in the theme's
 * accent with white text, added the way the Shapes button adds one, so the
 * slide that is photographed looks designed rather than typed.
 */
{
  const deck = Deck.open(
    buildPptx({
      title: 'Quarterly Review',
      slides: [
        { layout: 'title', title: 'Quarterly Review', body: 'Illustrative figures, prepared for the board' },
        { layout: 'obj', title: 'Where growth came from', body: ['North, in absolute terms', 'East, in consistency', 'West is flat and worth watching'] },
        { layout: 'obj', title: 'What changed', body: ['Churn 3.4% to 2.1%', 'Onboarding rewritten in Q3', 'Pricing unchanged'] },
      ],
    }),
  );
  const card = (y, figure, label) => deck.addShape(0, {
    preset: 'roundRect',
    x: 800, y, w: 380, h: 120,
    line: 'none',
    text: [
      { align: 'center', runs: [{ text: figure, size: 30, bold: true, color: '#FFFFFF' }] },
      { align: 'center', runs: [{ text: label, size: 13, color: '#FFFFFF' }] },
    ],
    name: label,
  });
  card(150, '£2.9m', 'Revenue, full year');
  card(300, '17.3%', 'Growth, Q4 on Q1');
  card(450, '2.1%', 'Churn, down from 3.4%');
  fs.writeFileSync(at('product-review.pptx'), deck.save());
}

/**
 * Sample pictures, generated here.
 *
 * **This exists because of what happened without it.** Given no file, the
 * Pictures app opens the operating system's own pictures folder — so the
 * first run of this script photographed the machine owner's family
 * photographs and portraits, filenames included, and those frames were one
 * commit away from a public website. Nothing was published; the capture was
 * deleted the moment it was looked at.
 *
 * The fix is not to remember to check. It is that these three apps are never
 * again pointed at a real folder: they open a directory this script created,
 * containing images this script drew. A generated gradient is a poor
 * photograph and an excellent guarantee.
 */
const picturesDir = path.join(docsDir, 'Sample Pictures');
fs.mkdirSync(picturesDir, { recursive: true });
const SAMPLES = [
  ['gradient-teal.png', [16, 74, 92], [126, 214, 214]],
  ['gradient-navy.png', [14, 26, 56], [86, 110, 178]],
  ['gradient-sand.png', [122, 92, 52], [232, 214, 178]],
  ['gradient-rose.png', [104, 32, 62], [226, 154, 176]],
];
/**
 * Real photographs, when there are any to use.
 *
 * `RUTBA_SCREEN_PICTURES=<dir>` copies images from that directory into the
 * sample folder instead of drawing gradients — so a set of properly licensed
 * stock photographs can be dropped in without touching this file. The bar for
 * what goes in that directory is the same bar as any published marketing
 * asset: a licence that permits it, and a model release for anybody
 * recognisable in it. That is why the default is a gradient rather than
 * whatever happens to be on the machine.
 */
const supplied = process.env.RUTBA_SCREEN_PICTURES;
let usedSupplied = 0;
if (supplied && fs.existsSync(supplied)) {
  for (const name of fs.readdirSync(supplied)) {
    if (!/.(jpe?g|png|webp|avif|gif)$/i.test(name)) continue;
    fs.copyFileSync(path.join(supplied, name), path.join(picturesDir, name));
    usedSupplied += 1;
  }
}

if (!usedSupplied) {
  for (const [name, from, to] of SAMPLES) {
    fs.writeFileSync(path.join(picturesDir, name), gradientPng(1200, 800, from, to));
  }
}

const firstPicture = () => path.join(picturesDir, fs.readdirSync(picturesDir).sort()[0]);

/**
 * A working week, for the Calendar capture.
 *
 * **Anchored to the current week rather than to a fixed date**, and that is the
 * whole reason this is generated instead of committed as a fixture. A calendar
 * screenshot is the one marketing asset that states what day it is; a file
 * holding March 2026 photographs as a product nobody has touched since March
 * 2026. Re-running this in June gives a June week, with no edit.
 *
 * The week runs Monday-first, which is what the month grid draws, and it is
 * deliberately not full: eleven entries over five days reads as somebody's
 * diary, and a wall of blocks reads as a stress test. One of them repeats,
 * because the recurrence engine is a thing this app has and most free
 * calendars get wrong.
 *
 * Every name is invented.
 */
const monday = (() => {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  // getDay() is Sunday-first; this shifts it to Monday-first before subtracting.
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
  return d;
})();

const dayAt = (day, hour, minute = 0) =>
  new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + day, hour, minute).getTime();

const meeting = (day, hour, minutes, summary, extra = {}) => {
  const at = dayAt(day, hour);
  return {
    summary,
    start: dateTimeAt(at, null),
    end: dateTimeAt(at + minutes * 60000, null),
    ...extra,
  };
};

fs.writeFileSync(
  at('this-week.ics'),
  writeCalendar({
    name: 'Work',
    events: [
      meeting(0, 9, 30, 'Monday stand-up', {
        // Weekly, so the week view shows the same slot filled on the day the
        // rule lands rather than only where an event was written out.
        rrule: { freq: 'WEEKLY', byDay: ['MO'] },
        location: 'Room 2',
      }),
      meeting(0, 14, 60, 'Quarterly review — draft walkthrough', { location: 'Room 2' }),
      meeting(1, 10, 45, 'Onboarding rewrite: what shipped', { location: 'Room 4' }),
      meeting(1, 16, 30, 'Call with Halton & Rowe', { location: 'Dial in' }),
      meeting(2, 11, 90, 'Regional sales — Q4 numbers', { location: 'Room 2' }),
      meeting(3, 9, 60, 'Interview: Aoife Brennan', { location: 'Room 1' }),
      meeting(3, 15, 30, 'Release readiness', { location: 'Room 4' }),
      meeting(4, 12, 60, 'Team lunch', { location: 'The Waterman' }),
      {
        summary: 'Marta out — annual leave',
        allDay: true,
        start: dateTimeAt(dayAt(3, 0), null, true),
        end: dateTimeAt(dayAt(5, 0), null, true),
        transparency: 'TRANSPARENT',
      },
    ],
  }),
);

/**
 * An address book, for the Contacts capture.
 *
 * Enough fields on each that the detail pane has something to draw — a title,
 * an organisation, two ways to reach them — because a contacts screenshot
 * whose right-hand pane holds one email address is a screenshot of an empty
 * pane. Fictional people at fictional companies, on example.com, which exists
 * to be exactly this.
 */
fs.writeFileSync(
  at('address-book.vcf'),
  writeVCards([
    {
      name: { given: 'Aoife', family: 'Brennan' },
      org: 'Halton & Rowe',
      title: 'Operations Director',
      emails: [{ value: 'a.brennan@example.com', type: 'work' }],
      phones: [{ value: '+44 20 7946 0812', type: 'work' }],
      addresses: [{ street: '14 Prospect Row', city: 'Bristol', postcode: 'BS1 4QT', country: 'United Kingdom', type: 'work' }],
    },
    {
      name: { given: 'Marta', family: 'Kowalczyk' },
      org: 'Northbank Logistics',
      title: 'Head of Finance',
      emails: [{ value: 'marta.k@example.com', type: 'work' }],
      phones: [{ value: '+44 161 496 0177', type: 'work' }],
    },
    {
      name: { given: 'Yusuf', family: 'Demir' },
      org: 'Demir Studio',
      title: 'Principal',
      emails: [{ value: 'yusuf@example.com', type: 'work' }],
      phones: [{ value: '+44 7700 900412', type: 'mobile' }],
    },
    {
      name: { given: 'Priya', family: 'Raghavan' },
      org: 'Halton & Rowe',
      title: 'Contracts Manager',
      emails: [{ value: 'p.raghavan@example.com', type: 'work' }],
    },
    {
      name: { given: 'Tomas', family: 'Lindqvist' },
      org: 'Vasa Partners',
      title: 'Analyst',
      emails: [{ value: 't.lindqvist@example.com', type: 'work' }],
      phones: [{ value: '+46 8 505 12 00', type: 'work' }],
    },
    {
      name: { given: 'Nkechi', family: 'Obi' },
      org: 'Northbank Logistics',
      title: 'Fleet Coordinator',
      emails: [{ value: 'n.obi@example.com', type: 'work' }],
    },
  ]),
);

/**
 * A report as a PDF, for the PDF app's capture: a title, prose, a table and
 * a bar chart over four pages, so the thumbnail strip shows pages that
 * differ. Drawn by the suite's own writer; the figures are invented.
 */
function writeReportPdf(file) {
  const sheet = new Sheet({ title: 'Northbank Logistics: Annual Review 2026', author: 'Northbank Logistics', created: '2026-10-13T09:00:00Z', footer: 'Northbank Logistics · Annual Review 2026', fontSize: 10.5 });
  const regions = [['North', 412, 468], ['Midlands', 386, 401], ['South East', 521, 590], ['South West', 244, 262], ['Scotland', 198, 231], ['Wales', 156, 170]];
  sheet.heading('Annual Review 2026', { size: 26, gap: 4 });
  sheet.paragraph('Northbank Logistics', { size: 13, colour: INK.quiet, gap: 14 });
  sheet.paragraph('A year of steady growth: every region moved more freight than last year, the new South East depot opened on time, and the fleet ran its first full year on renewable diesel. This review sets out the figures region by region, what changed, and where we go next.', { gap: 10 });
  sheet.keyValues([['Revenue', '£48.6 million'], ['Deliveries', '1.2 million'], ['On time', '97.4%'], ['Depots', '14']], { width: 260, gap: 3 });
  sheet.spacer(10);
  sheet.heading('Freight by region', { size: 15 });
  sheet.table({
    columns: [{ key: 'r', label: 'Region', flex: 2 }, { key: 'a', label: '2025 (kt)', align: 'right' }, { key: 'b', label: '2026 (kt)', align: 'right' }, { key: 'c', label: 'Change', align: 'right' }],
    rows: regions.map(([r, a, b]) => ({ r, a: String(a), b: String(b), c: `+${(((b - a) / a) * 100).toFixed(1)}%` })),
    zebra: true,
  });
  sheet.spacer(6);
  sheet.space(170);
  const base = sheet.y + 150;
  const barW = 26;
  regions.forEach(([r, a, b], i) => {
    const x = sheet.left + 20 + i * 82;
    sheet.page.rect(x, base - a * 0.24, barW, a * 0.24, { fill: '#9db8e6' });
    sheet.page.rect(x + barW + 4, base - b * 0.24, barW, b * 0.24, { fill: '#2b5797' });
    sheet.page.text(r, x, base + 14, { size: 8.5, colour: INK.quiet });
  });
  sheet.page.rect(sheet.left, base, sheet.contentWidth, 0.6, { fill: INK.rule });
  sheet.y = base + 30;
  sheet.newPage();
  sheet.heading('What changed', { size: 15 });
  for (const para of [
    'The South East depot at Dartford opened in March and handled a fifth of the region\'s freight by the summer. Journeys into London fell by a quarter, and with them the hours lost to traffic.',
    'Scotland grew fastest, at more than sixteen per cent, after two supermarket contracts moved to us in the spring. The Glasgow cross-dock now runs a night shift.',
    'Every lorry ran on renewable diesel from January, cutting the fleet\'s emissions by about four fifths for the same miles. The cost was a little over two per cent of fuel spend.',
  ]) sheet.paragraph(para, { gap: 10 });
  sheet.heading('Next year', { size: 15 });
  for (const para of [
    'Two electric tractor units go into service on the Midlands trunk routes in January, charged at the Daventry depot overnight.',
    'A new booking system lets customers choose a delivery hour rather than a day, starting with the North in February.',
  ]) sheet.paragraph(para, { gap: 10 });
  sheet.newPage();
  sheet.heading('Our people', { size: 15 });
  sheet.paragraph('We employ 1,840 people across fourteen depots, 212 more than a year ago. Staff turnover fell to nine per cent, the lowest since 2019, and every driver completed the new safety course.', { gap: 10 });
  sheet.keyValues([['Employees', '1,840'], ['Drivers', '1,106'], ['Apprentices', '64'], ['Turnover', '9%']], { width: 260, gap: 3 });
  sheet.newPage();
  sheet.heading('Accounts in brief', { size: 15 });
  sheet.table({
    columns: [{ key: 'k', label: '', flex: 3 }, { key: 'a', label: '2025', align: 'right' }, { key: 'b', label: '2026', align: 'right' }],
    rows: [['Revenue', '£44.1m', '£48.6m'], ['Operating costs', '£39.8m', '£43.2m'], ['Operating profit', '£4.3m', '£5.4m'], ['Margin', '9.8%', '11.1%']].map(([k, a, b]) => ({ k, a, b })),
    zebra: true,
  });
  fs.writeFileSync(file, sheet.toBuffer());
}
writeReportPdf(at('Annual Review 2026.pdf'));

/**
 * A video project for the Video app's capture: the sample pictures as stills
 * on the timeline with fades between them, a title over the first and a
 * marker, so the timeline, the preview and the layers all have something
 * to show. Video sources are not generated (the suite writes no sample
 * film), but stills are what an introduction video is often made of.
 */
function writeVideoProject(file) {
  let doc = emptyProject();
  const stills = fs.readdirSync(picturesDir).filter((n) => /\.(png|jpe?g|webp)$/i.test(n)).sort();
  for (const name of stills) {
    const full = path.join(picturesDir, name);
    doc = addMedia(doc, { url: full, name, kind: 'image', width: 1200, height: 800, size: fs.statSync(full).size }).doc;
  }
  const ids = doc.patches.filter((p) => p.type === 'image').map((p) => p.id);
  for (const id of ids.slice(1)) doc = setTransition(doc, id, { kind: 'fade', seconds: 1 });
  doc = addTitle(doc, 'title', { at: 0, text: 'Our year at Northbank' }).doc;
  doc = addMarker(doc, 10, 'Depot opening');
  fs.writeFileSync(file, writeRvid(doc, { projectPath: file }));
}
writeVideoProject(at('Our year.rvid'));

/** app -> the file it should open, or null for its empty state. */
const PLAN = [
  ['home', null],
  ['word', at('quarterly-review.docx')],
  ['sheets', at('regional-sales.xlsx')],
  ['slides', at('product-review.pptx')],
  ['pdf', at('Annual Review 2026.pdf')],
  ['mail', null],
  // Added when the site was found publishing seven apps over a suite of nine:
  // Calendar and Contacts shipped in 1.8.0 and this list was never extended, so
  // they had no screenshot to be missing from anything. Both take a file for the
  // same reason the document apps do — an empty week and an empty address book
  // prove the window paints and sell nothing.
  ['calendar', at('this-week.ics')],
  ['contacts', at('address-book.vcf')],
  // NEVER null. Null means the app falls back to the operating system's
  // own folders, which is how somebody's family photographs ended up in a
  // build artefact. See the note above `gradientPng`.
  ['pictures', firstPicture()],
  ['image', firstPicture()],
  // Video opens a project of the sample stills (writeVideoProject above):
  // nothing real is pointed at, and the timeline has something to show.
  ['video', at('Our year.rvid')],
];

/**
 * A window size of its own, where the one size does not suit: the launcher's
 * tiles and templates fill a smaller window, and a launcher photographed at
 * a document's size is mostly its empty recent-files list.
 */
const SIZE = {
  home: '1180x780',
  // The timeline takes a fixed height, so the preview only has room to
  // show the film in a taller window.
  video: '1440x1080',
};

/**
 * How long to let each window settle before photographing it, in ms.
 *
 * Per-app because the work differs: a launcher has nothing to wait for, a
 * photo viewer is decoding a half-megabyte JPEG and rendering a thumbnail
 * for every file in the folder, and mail is reading a store off disk.
 */
const SETTLE = {
  home: 1500,
  pdf: 4000,
  video: 5000,
  pictures: 6000,
  image: 5000,
  mail: 5000,
  // Both parse a file and lay out a view over it — the calendar expands a
  // recurrence rule across the week before it can draw the grid.
  calendar: 3500,
  contacts: 3000,
};

function capture(app, file, profile) {
  return new Promise((resolve) => {
    const env = {
      ...process.env,
      RUTBA_OFFICE_SMOKE: '1',
      RUTBA_SMOKE_OUT: outDir,
      // A smoke run wants to be quick and a marketing capture wants to be
      // right. At the smoke default of 900ms the photo viewer photographed
      // blank thumbnails over a JPEG decoded to the halfway line, and mail
      // caught its folder before the messages landed - both of which look
      // like product defects in a screenshot and are neither.
      RUTBA_SMOKE_SETTLE: String(SETTLE[app] ?? 2500),
      ELECTRON_DISABLE_SECURITY_WARNINGS: '1',
      // One window size for every app, so the pictures sit in a row on the
      // site at one height; above the primary display, so a run never
      // covers whoever is working and the window is never clamped to a
      // smaller display at the end of the row; and at a device scale of one
      // and a half, so the capture is sharp at the site's largest size on
      // whatever display the run happened on — a scale of two shrinks a
      // 2560-wide primary to 1280 logical and the window with it. Each can
      // be overridden from the environment.
      RUTBA_WINDOW_SIZE: process.env.RUTBA_WINDOW_SIZE || SIZE[app] || '1440x900',
      RUTBA_WINDOW_DISPLAY: process.env.RUTBA_WINDOW_DISPLAY || 'above',
      RUTBA_SCREEN_SCALE: process.env.RUTBA_SCREEN_SCALE || '1.5',
    };
    if (file) env.RUTBA_SMOKE_FILE = file;
    else delete env.RUTBA_SMOKE_FILE;
    // Mail is the one app whose empty state is not worth a picture, so it gets
    // the smoke seeder's invented correspondents.
    if (app === 'mail') {
      env.RUTBA_SMOKE_SEED = '1';
      // The import takes the account and folder name from the mbox filename,
      // so this string is the one the sidebar, the window title, the search
      // box and the status bar all repeat. Left at its default the mail
      // screenshot was captioned "smoke-seed — smoke-seed", which is accurate
      // and unpublishable. `Archive` is also the honest word for what it is:
      // an imported mailbox, which is the feature the picture is selling.
      env.RUTBA_SMOKE_SEED_NAME = 'Archive';
    } else {
      delete env.RUTBA_SMOKE_SEED;
      delete env.RUTBA_SMOKE_SEED_NAME;
    }

    const child = spawn(electron, [appDir, `--user-data-dir=${profile}`, app], {
      stdio: 'ignore',
      env,
    });
    child.on('exit', (code) => resolve(code === 0));
  });
}

const results = [];
// `RUTBA_SCREENS_ONLY=pdf,video` captures just those, to look at one again.
const only = (process.env.RUTBA_SCREENS_ONLY || '').split(',').filter(Boolean);
for (const [app, file] of PLAN.filter(([name]) => !only.length || only.includes(name))) {
  // A profile per app, not per run: the mail seeder is idempotent-ish but the
  // recent-files list is not, and a "recently opened" panel full of temp paths
  // is exactly the detail that makes a screenshot look like a test artefact.
  const profile = path.join(work, `profile-${app}`);
  fs.mkdirSync(profile, { recursive: true });
  const ok = await capture(app, file, profile);
  const png = path.join(outDir, `${app}.png`);
  const size = fs.existsSync(png) ? fs.statSync(png).size : 0;
  results.push({ app, ok: ok && size > 5000, size });
  process.stdout.write(
    `${ok && size > 5000 ? 'ok  ' : 'FAIL'} ${app.padEnd(9)} ${String(Math.round(size / 1024)).padStart(5)} KB\n`,
  );
}

fs.rmSync(work, { recursive: true, force: true });

const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} captured into ${outDir}`);
process.exit(failed.length ? 1 : 0);
