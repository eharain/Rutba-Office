/**
 * Where the pages are cut.
 *
 * Printing is two halves: deciding what goes on which page, and putting ink
 * on it. The second half is Chromium's and is checked by looking at paper;
 * the first is arithmetic over the geometry of a sheet or the size of a
 * slide, and that is what is held here — a suite that prints six columns per
 * page when they would all fit is worse than one that cannot print at all,
 * because it wastes the paper before anybody notices.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { SheetView } from '@rutba/sheet-view';
import { printHtml, printSummary, printableArea, pageSetup, readPageSetup, writePageSetup, PAPER } from '@rutba/sheet-view/print';
import { Deck, buildPptx } from '@rutba/presentation';
import { deckPrintHtml, deckPrintSummary } from '@rutba/presentation/print';
import { buildXlsx } from '@rutba/ooxml/build';

const REPORT = [['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total']];
for (let r = 0; r < 120; r++) REPORT.push([`Row ${r + 1}`, r, r * 2, r * 3, r * 4, `=SUM(B${r + 2}:E${r + 2})`]);
const WIDE = Array.from({ length: 60 }, (_, r) => Array.from({ length: 40 }, (_, c) => `c${r}-${c}`));
const BOOK = buildXlsx({ sheets: [{ name: 'Report', rows: REPORT }, { name: 'Wide', rows: WIDE }] });

test('a page is the paper less its margins, in the direction it is turned', () => {
  const portrait = printableArea(pageSetup({ paper: 'A4' }));
  const landscape = printableArea(pageSetup({ paper: 'A4', orientation: 'landscape' }));
  assert.ok(portrait.height > portrait.width, 'portrait is taller than it is wide');
  assert.ok(landscape.width > landscape.height, 'and landscape is not');
  assert.equal(Math.round(portrait.width), Math.round(landscape.height), 'the paper is the same paper');
  // A4 is 210 × 297; a 12.7 mm margin each side leaves 184.6 mm, which is 698 px.
  assert.equal(Math.round(portrait.width), 698);
  assert.ok(PAPER.Letter.width > PAPER.A4.width, 'Letter is the wider paper');
});

test('columns and rows are cut at a page, never through one', () => {
  const view = new SheetView(BOOK);
  const report = printSummary(view, {});
  assert.equal(report.pages, 3, 'a 121-row report is three A4 pages');
  assert.equal(report.sheets[0].ref, 'A1:F121', 'and it prints what has anything in it');

  view.activeSheet = 'Wide';
  assert.equal(printSummary(view, {}).pages, 8, 'forty columns need four sheets across, twice down');
  // Fitting to width scales once, for everything, and then it is one page.
  const fitted = printSummary(view, { fit: 'width' });
  assert.equal(fitted.pages, 1);
  assert.ok(fitted.sheets[0].scale < 0.3 && fitted.sheets[0].scale > 0.2, `scaled to ${fitted.sheets[0].scale}`);
  // The scaled width is exactly the page width; a comparison without slack
  // put the last column on a page of its own.
  assert.equal(printSummary(view, { fit: 'page' }).pages, 1);
});

test('a print area, and a workbook printed whole', () => {
  const view = new SheetView(BOOK);
  assert.equal(printSummary(view, { area: 'A1:C10' }).pages, 1);
  assert.equal(printSummary(view, { area: '$A$1:$C$10' }).sheets[0].ref, 'A1:C10', 'dollars are addresses, not text');
  const all = printSummary(view, { sheets: 'all' });
  assert.equal(all.pages, 11);
  assert.deepEqual(all.sheets.map((s) => s.sheet), ['Report', 'Wide'], 'in tab order');
});

test('the page carries what the sheet says, and what the setup asks for', () => {
  const view = new SheetView(BOOK);
  const html = printHtml(view, { headings: true, gridlines: true, repeatRows: 1, footer: 'Page &P of &N', file: 'sales.xlsx' });
  const pages = html.split('<section class="page');
  assert.equal(pages.length - 1, 3);
  assert.match(html, /@page \{ size: A4 portrait/);
  assert.ok(pages[2].includes('Region'), 'a repeated row is on the second page');
  assert.ok(pages[2].includes('Page 2 of 3'), 'and the footer counts');
  assert.ok(html.includes('>10<'), 'a formula prints what it computed, not itself');
  assert.match(html, /<th>A<\/th>/, 'the column headings are drawn when asked for');
  // Nothing a cell holds can become markup.
  const nasty = new SheetView(buildXlsx({ sheets: [{ name: 'S', rows: [['<script>alert(1)</script>', '5 < 6 & 7']] }] }));
  const escaped = printHtml(nasty, {});
  assert.doesNotMatch(escaped, /<script>alert/);
  assert.match(escaped, /5 &lt; 6 &amp; 7/);
});

test('a deck prints as slides, as notes, or as a handout', () => {
  const deck = Deck.open(
    buildPptx({
      title: 'Quarterly',
      slides: [
        { layout: 'title', title: 'Quarterly review', subtitle: 'September 2026' },
        ...Array.from({ length: 12 }, (_, i) => ({ layout: 'titleAndContent', title: `Slide ${i + 2}`, bullets: ['One', 'Two'] })),
      ],
    })
  );
  deck.setNotes(0, 'Open with the tracker report.');

  const slides = deckPrintSummary(deck, { layout: 'slides' });
  assert.equal(slides.pages, 13, 'one slide to a page');
  assert.equal(slides.setup.orientation, 'landscape', 'because a slide is wider than it is tall');

  const notes = deckPrintSummary(deck, { layout: 'notes' });
  assert.equal(notes.pages, 13);
  assert.equal(notes.setup.orientation, 'portrait', 'notes need the room below the slide');
  assert.ok(deckPrintHtml(deck, { layout: 'notes' }).includes('Open with the tracker report'));

  assert.equal(deckPrintSummary(deck, { layout: 'handout', perPage: 6 }).pages, 3);
  assert.equal(deckPrintSummary(deck, { layout: 'handout', perPage: 4 }).pages, 4);
  assert.equal(deckPrintSummary(deck, { layout: 'handout', perPage: 7 }).setup.perPage, 6, 'seven to a page is not a handout PowerPoint offers');

  const chosen = deckPrintHtml(deck, { slides: [0, 2, 4] });
  assert.equal((chosen.match(/<section class="page/g) || []).length, 3);
  assert.match(chosen, /<svg /, 'and each one is drawn, not described');
});

test('the page setup belongs to the workbook, not to the dialog', () => {
  // Excel keeps paper, orientation, margins, scaling, gridlines, the print
  // area and the repeated rows in the file — in the sheet's own tail and in
  // two sheet-scoped defined names. A setup that lives only in a dialog is
  // one a person sets again every time, and one that never reaches whoever
  // opens the file next.
  const book = buildXlsx({ sheets: [{ name: 'Report', rows: [['Region', 'Total'], ['North', 12], ['South', 8]] }, { name: 'Second', rows: [['a', 'b']] }] });
  const view = new SheetView(book);
  assert.equal(readPageSetup(view, 'Report').paper, 'A4', 'a workbook with no setup reads as the default');

  writePageSetup(view, 'Report', {
    paper: 'Letter', orientation: 'landscape', fit: 'width',
    margins: { top: 6.4, right: 6.4, bottom: 6.4, left: 6.4 },
    gridlines: true, headings: true, repeatRows: 1, area: 'A1:B3', centre: { horizontal: true },
  });
  writePageSetup(view, 'Second', { orientation: 'portrait', area: 'A1:A2' });

  const back = new SheetView(view.save());
  const read = readPageSetup(back, 'Report');
  assert.equal(read.paper, 'Letter');
  assert.equal(read.orientation, 'landscape');
  assert.equal(read.fit, 'width');
  assert.equal(read.gridlines, true);
  assert.equal(read.headings, true);
  assert.equal(read.repeatRows, 1);
  assert.equal(read.area, 'A1:B3');
  assert.equal(read.centre.horizontal, true);
  assert.equal(Math.round(read.margins.left * 10) / 10, 6.4, 'margins survive the trip through inches');

  // Each sheet keeps its own: a name scoped to one sheet must not be the
  // other's, which is what matching a defined name by its name alone did.
  assert.equal(readPageSetup(back, 'Second').area, 'A1:A2');
  assert.equal(readPageSetup(back, 'Second').orientation, 'portrait');

  const sheetXml = back.pkg.text(back.workbook._sheetPart('Report').sheet.part);
  assert.match(sheetXml, /fitToPage="1"/, 'Excel ignores fitToWidth without it');
  assert.deepEqual(
    (sheetXml.match(/<(printOptions|pageMargins|pageSetup)\b/g) || []).map((t) => t.slice(1)),
    ['printOptions', 'pageMargins', 'pageSetup'],
    'and the schema wants them in that order'
  );

  // And the plan the printer follows is the file's, not the default's.
  back.activeSheet = 'Report';
  assert.equal(printSummary(back, readPageSetup(back, 'Report')).sheets[0].ref, 'A1:B3');
});
