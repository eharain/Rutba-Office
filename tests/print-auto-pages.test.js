// Two things the print dialog does for every app that prints through it (2.27.0).
//
// Which way round the paper goes is chosen for a person who has not chosen: a sheet too wide for
// upright paper, a slide and a picture wider than it is tall go on paper turned sideways, and the
// rest stay upright. What a person chose, and what a workbook's own file says, stands.
//
// And the Pages box takes page numbers as they are typed everywhere ("1,2,5,6", "1-5"): the job is
// made a PDF, the pages asked for are copied out of it, and only those reach the printer.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import pdfDocument from '@rutba/pdf/document';
import pdfRead from '@rutba/pdf/reader';
import { SheetView } from '@rutba/sheet-view';
import { chooseOrientation, printHtml, printSummary, readPageSetup, writePageSetup } from '@rutba/sheet-view/print';
import { Deck, buildPptx } from '@rutba/presentation';
import { deckPrintSetup, deckPrintSummary, deckPrintHtml } from '@rutba/presentation/print';
import { buildXlsx } from '@rutba/ooxml/build';
import { tempDir } from './temp-dirs.js';
import { createPrintService, pagesWanted, pdfOfPages, pdfPages } from '../apps/desktop/main/print.js';
import { setMemoryFiles } from '../apps/desktop/main/memory-files.js';
import { pagesTyped, wayChosen, wayOfShape } from '../apps/desktop/renderer/print-page.js';

const { PdfDocument } = pdfDocument;
const { openPdf } = pdfRead;

after(() => setMemoryFiles(null));

/* ── the pages typed ────────────────────────────────────────────────────── */

test('page numbers are read as they are typed: a list, a range, or both', () => {
  assert.deepEqual(pagesTyped('1,2,5,6', 9), { pages: [1, 2, 5, 6] });
  assert.deepEqual(pagesTyped('1-5', 9), { pages: [1, 2, 3, 4, 5] });
  assert.deepEqual(pagesTyped('2-4, 9', 9), { pages: [2, 3, 4, 9] });
  assert.deepEqual(pagesTyped(' 3 ', 9), { pages: [3] });
  assert.deepEqual(pagesTyped('1 2 5', 9), { pages: [1, 2, 5] }, 'spaces part numbers as commas do');
  assert.deepEqual(pagesTyped('1;3', 9), { pages: [1, 3] });
  assert.deepEqual(pagesTyped('2 - 4', 9), { pages: [2, 3, 4] }, 'a range may be typed with room around its dash');
  assert.deepEqual(pagesTyped('1,2,', 9), { pages: [1, 2] }, 'a comma left at the end is not a mistake');
});

test('a range may be left open, runs no further than the last page, and a page is printed once', () => {
  assert.deepEqual(pagesTyped('7-', 9), { pages: [7, 8, 9] }, 'seven to the end');
  assert.deepEqual(pagesTyped('-3', 9), { pages: [1, 2, 3] }, 'the start to three');
  assert.deepEqual(pagesTyped('8-20', 9), { pages: [8, 9] }, 'a range past the last page stops at it');
  assert.deepEqual(pagesTyped('5,1-3,2,5', 9), { pages: [5, 1, 2, 3] }, 'in the order typed, each once: more of a page is Copies');
});

test('the numbers may be Arabic or Persian digits, and the dash any dash', () => {
  assert.deepEqual(pagesTyped('١،٢،٥', 9), { pages: [1, 2, 5] });
  assert.deepEqual(pagesTyped('۱-۳', 9), { pages: [1, 2, 3] });
  assert.deepEqual(pagesTyped('2\u20134', 9), { pages: [2, 3, 4] }, 'an en dash');
  assert.deepEqual(pagesTyped('2\u20144', 9), { pages: [2, 3, 4] }, 'an em dash');
  assert.deepEqual(pagesTyped('2\u22124', 9), { pages: [2, 3, 4] }, 'a minus sign');
});

test('nothing typed is every page; what is not pages of this print is said to be wrong', () => {
  assert.deepEqual(pagesTyped('', 9), { all: true });
  assert.deepEqual(pagesTyped('   ', 9), { all: true });
  assert.deepEqual(pagesTyped(null, 9), { all: true });
  for (const wrong of ['0', '10', '5-2', '10-12', 'two', '1,x', '1.5', '-', '1--3', ',']) {
    assert.deepEqual(pagesTyped(wrong, 9), { wrong: true }, `"${wrong}" of nine pages`);
  }
  // Before the pages are counted there is nothing to read the numbers against.
  assert.deepEqual(pagesTyped('1-3', null), { waiting: true });
  assert.deepEqual(pagesTyped('1-3', 0), { waiting: true });
  assert.deepEqual(pagesTyped('', null), { all: true });
});

test('the main process takes the page numbers again, and refuses one that is not a page', () => {
  assert.equal(pagesWanted(null, 5), null);
  assert.equal(pagesWanted(undefined, 5), null);
  assert.equal(pagesWanted([], 5), null, 'none named is every page');
  assert.deepEqual(pagesWanted([1, 2, 5], 5), [0, 1, 4], 'counted from nought for the PDF');
  assert.deepEqual(pagesWanted([5, 1, 5], 5), [4, 0], 'in the order given, each once');
  assert.throws(() => pagesWanted([6], 5), /There is no page 6: this print has 5 pages\./);
  assert.throws(() => pagesWanted([0], 5), /There is no page 0/);
  assert.throws(() => pagesWanted([1.5], 5), /There is no page 1\.5/);
  assert.throws(() => pagesWanted([2], 1), /this print has one page\./);
  assert.throws(() => pagesWanted('1-3', 5), /given as page numbers/);
});

/** A PDF of so many pages, each saying which it is, as Chromium would hand one back. */
function pdfOf(pages, size = null) {
  const doc = new PdfDocument({ title: 'Printed', created: '2026-11-22T09:00:00Z' });
  for (let i = 0; i < pages; i++) doc.addPage(size).text(`Sheet ${i + 1}`, 72, 100, { font: 'Helvetica', size: 12 });
  return doc.toBuffer();
}
const wordsOf = (bytes) => {
  const file = openPdf(new Uint8Array(bytes));
  return Array.from({ length: file.pageCount }, (_, i) => file.pageText(i).trim());
};

test('the pages asked for are copied into a PDF of their own, in the order asked', () => {
  const whole = pdfOf(6);
  assert.deepEqual(wordsOf(pdfOfPages(whole, [0, 1, 4, 5])), ['Sheet 1', 'Sheet 2', 'Sheet 5', 'Sheet 6']);
  assert.deepEqual(wordsOf(pdfOfPages(whole, [3])), ['Sheet 4']);
  assert.deepEqual(wordsOf(pdfOfPages(whole, [5, 0])), ['Sheet 6', 'Sheet 1']);
  assert.equal(pdfPages(whole), 6, 'and the whole is as it was');
});

/**
 * A stand-in for Electron's BrowserWindow: each window says what it was handed to load (the bytes,
 * whatever they are), what it was asked to make a PDF with, and what it printed with. `size` is the
 * paper Chromium lays the pages out on (points), for a page whose own CSS turns it sideways.
 */
function fakeElectron({ pages = 6, size = null } = {}) {
  const seen = { loads: [], prints: [], pdfs: [], destroyed: 0 };
  class BrowserWindow {
    constructor(options) {
      this.options = options;
      this.webContents = {
        on: () => {},
        setWindowOpenHandler: () => {},
        printToPDF: async (options) => {
          seen.pdfs.push(options);
          return pdfOf(pages, size);
        },
        print: (options, done) => {
          seen.prints.push(options);
          done(true, '');
        },
        getURL: () => '',
      };
    }
    async loadFile(file) {
      seen.loads.push({ file, bytes: fs.readFileSync(file), plugins: this.options.webPreferences.plugins });
    }
    isDestroyed() { return false; }
    destroy() { seen.destroyed++; }
  }
  return { seen, electron: { BrowserWindow } };
}

const PAGE = '<!doctype html><html><body><h1>Quarterly figures</h1></body></html>';
/** A4 turned sideways, in points: what Chromium writes for a page laid out landscape. */
const SIDEWAYS = [841.89, 595.28];

test('a page an app laid out prints only the pages asked for: laid out whole, then those pages sent', async () => {
  setMemoryFiles(null);
  const { seen, electron } = fakeElectron({ pages: 6, size: SIDEWAYS });
  const print = createPrintService({ docs: {}, electron });
  assert.deepEqual(await print.page({ html: PAGE, name: 'A view', options: { orientation: 'landscape' }, copies: 2, pages: [1, 2, 5, 6] }), { ok: true, reason: null });
  assert.equal(seen.loads.length, 2, 'one window lays the page out, the next prints the pages kept');
  assert.equal(seen.loads[0].bytes.toString('utf8'), PAGE);
  assert.deepEqual(seen.pdfs[0], { printBackground: true, preferCSSPageSize: true, margins: { marginType: 'none' } }, 'laid out as the count lays it out');
  assert.equal(path.extname(seen.loads[1].file), '.pdf');
  assert.equal(seen.loads[1].plugins, true, 'the second window is one that shows a PDF');
  assert.deepEqual(wordsOf(seen.loads[1].bytes), ['Sheet 1', 'Sheet 2', 'Sheet 5', 'Sheet 6']);
  // The printer is told the paper and the way round of the pages it is handed: A4, sideways.
  assert.deepEqual(seen.prints, [{ silent: false, deviceName: undefined, printBackground: true, copies: 2, landscape: true, pageSize: 'A4', margins: { marginType: 'none' } }]);
  assert.equal(seen.destroyed, 2);
  assert.ok(seen.loads.every((l) => !fs.existsSync(path.dirname(l.file))), 'and neither job leaves its temp folder behind');
});

test('with no pages named a page prints as it always did, through one window', async () => {
  setMemoryFiles(null);
  for (const args of [{}, { pages: null }, { pages: [] }]) {
    const { seen, electron } = fakeElectron();
    await createPrintService({ docs: {}, electron }).page({ html: PAGE, name: 'A view', ...args });
    assert.equal(seen.loads.length, 1);
    assert.equal(seen.pdfs.length, 0, 'nothing is made a PDF on the way');
    assert.equal(seen.loads[0].bytes.toString('utf8'), PAGE);
  }
});

test('a page number that is not a page of the print stops it in words, and nothing is sent', async () => {
  setMemoryFiles(null);
  const { seen, electron } = fakeElectron({ pages: 3 });
  const print = createPrintService({ docs: {}, electron });
  await assert.rejects(print.page({ html: PAGE, name: 'A view', pages: [2, 4] }), /There is no page 4: this print has 3 pages\./);
  assert.equal(seen.prints.length, 0);
  assert.equal(seen.destroyed, 1, 'and the window that laid it out is gone');
});

test('Save as PDF of a page keeps only the pages asked for, and says how many', async () => {
  setMemoryFiles(null);
  const dir = tempDir('rl-print-pages-');
  const target = path.join(dir, 'Quarterly figures.pdf');
  const { electron } = fakeElectron({ pages: 6 });
  const print = createPrintService({ docs: {}, electron });
  assert.deepEqual(await print.pagePdf({ html: PAGE, name: 'Quarterly figures', path: target, pages: [2, 3] }), { path: target, format: 'pdf', pages: 2 });
  assert.deepEqual(wordsOf(fs.readFileSync(target)), ['Sheet 2', 'Sheet 3']);
  assert.deepEqual(await print.pagePdf({ html: PAGE, name: 'Quarterly figures', path: target }), { path: target, format: 'pdf', pages: 6 });
  assert.equal(pdfPages(fs.readFileSync(target)), 6);
});

test('a workbook or a deck prints the pages asked for, turned the way it was laid out', async () => {
  setMemoryFiles(null);
  const { seen, electron } = fakeElectron({ pages: 6, size: SIDEWAYS });
  const asked = [];
  const docs = {
    printSource: ({ options }) => {
      asked.push(options);
      // The engine chose: nobody had, and the sheet is wide.
      return { kind: 'sheet', name: 'Budget.xlsx', locked: false, orientation: 'landscape', html: '<p>Budget</p>' };
    },
  };
  const print = createPrintService({ docs, electron });
  assert.deepEqual(await print.document({ id: 'x', options: { orientation: 'auto', paper: 'A4' }, pages: [3, 4] }), { ok: true, reason: null });
  assert.deepEqual(asked[0], { orientation: 'auto', paper: 'A4' }, 'the engine is asked to lay out, not which pages');
  assert.equal(seen.pdfs[0].landscape, true, 'the whole is laid out the way the engine chose');
  assert.deepEqual(wordsOf(seen.loads[1].bytes), ['Sheet 3', 'Sheet 4']);
  assert.equal(seen.prints[0].landscape, true, 'and the printer is told the same way');

  // With nothing asked for, one window and the way the engine chose.
  const plain = fakeElectron();
  await createPrintService({ docs, electron: plain.electron }).document({ id: 'x', options: { orientation: 'auto' } });
  assert.equal(plain.seen.loads.length, 1);
  assert.equal(plain.seen.prints[0].landscape, true);

  // A way a person chose is the one an engine that says nothing is printed by, as before.
  const quiet = fakeElectron();
  await createPrintService({ docs: { printSource: () => ({ kind: 'sheet', name: 'B.xlsx', html: '<p>B</p>' }) }, electron: quiet.electron }).document({ id: 'x', options: { orientation: 'landscape' } });
  assert.equal(quiet.seen.prints[0].landscape, true);
});

test('a document, which is a PDF already, prints the pages asked for without being laid out twice', async () => {
  setMemoryFiles(null);
  const { seen, electron } = fakeElectron();
  const docs = { printSource: () => ({ kind: 'doc', name: 'Letter.docx', locked: false, pdf: new Uint8Array(pdfOf(8)) }) };
  await createPrintService({ docs, electron }).document({ id: 'x', pages: [1, 8] });
  assert.equal(seen.loads.length, 1);
  assert.equal(seen.pdfs.length, 0);
  assert.deepEqual(wordsOf(seen.loads[0].bytes), ['Sheet 1', 'Sheet 8']);
});

test('Save as PDF of a document keeps the pages asked for, named beside the options or among them', async () => {
  setMemoryFiles(null);
  const dir = tempDir('rl-print-pages-');
  const target = path.join(dir, 'Budget.pdf');
  const asked = [];
  const sheet = { printSource: ({ options }) => { asked.push(options); return { kind: 'sheet', name: 'Budget.xlsx', orientation: 'landscape', html: '<p>Budget</p>' }; } };
  const { seen, electron } = fakeElectron({ pages: 6 });
  const print = createPrintService({ docs: sheet, electron });
  // A window's Save as PDF hands on the dialog's options, the pages among them.
  await print.pdf({ id: 'x', path: target, options: { orientation: 'auto', pages: [2, 5] } });
  assert.deepEqual(asked[0], { orientation: 'auto' }, 'the page numbers are the print\'s, not the layout\'s');
  assert.equal(seen.pdfs[0].landscape, true);
  assert.deepEqual(wordsOf(fs.readFileSync(target)), ['Sheet 2', 'Sheet 5']);
  await print.pdf({ id: 'x', path: target, options: {}, pages: [6] });
  assert.deepEqual(wordsOf(fs.readFileSync(target)), ['Sheet 6']);
  await print.pdf({ id: 'x', path: target, options: {} });
  assert.equal(pdfPages(fs.readFileSync(target)), 6, 'and every page when none is named');

  const letter = { printSource: () => ({ kind: 'doc', name: 'Letter.docx', pdf: new Uint8Array(pdfOf(4)) }) };
  await createPrintService({ docs: letter, electron }).pdf({ id: 'x', path: target, options: { pages: [4, 1] } });
  assert.deepEqual(wordsOf(fs.readFileSync(target)), ['Sheet 4', 'Sheet 1']);
  await assert.rejects(createPrintService({ docs: letter, electron }).pdf({ id: 'x', path: target, options: { pages: [5] } }), /There is no page 5: this print has 4 pages\./);
});

/* ── which way round ────────────────────────────────────────────────────── */

test('a way is chosen only when it is one of the two, and a shape wider than it is tall is turned', () => {
  assert.equal(wayChosen('portrait'), true);
  assert.equal(wayChosen('landscape'), true);
  for (const not of ['auto', '', undefined, null, 'sideways']) assert.equal(wayChosen(not), false);
  assert.equal(wayOfShape(1600, 900), 'landscape');
  assert.equal(wayOfShape(900, 1600), 'portrait');
  assert.equal(wayOfShape(800, 800), 'portrait', 'a square stays upright');
  assert.equal(wayOfShape(0, 0), 'portrait');
  assert.equal(wayOfShape(undefined, undefined), 'portrait', 'and so does a picture whose size is not known');
});

const grid = (rows, cols) => Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => (r === 0 ? 'H' + c : r * 10 + c)));
const book = (rows, cols) => new SheetView(buildXlsx({ sheets: [{ name: 'Report', rows: grid(rows, cols) }] }));

test('asked to choose, a sheet too wide for upright paper is turned, and a narrow one is not', () => {
  // Fourteen columns: two sheets across upright, one across sideways.
  const wide = book(150, 14);
  const tall = printSummary(wide, { orientation: 'portrait' });
  const turned = printSummary(wide, { orientation: 'landscape' });
  assert.ok(turned.pages < tall.pages, `${turned.pages} sheets sideways against ${tall.pages} upright`);
  assert.equal(chooseOrientation(wide, { orientation: 'auto' }), 'landscape');
  const auto = printSummary(wide, { orientation: 'auto' });
  assert.equal(auto.setup.orientation, 'landscape', 'the summary says which way was chosen');
  assert.equal(auto.pages, turned.pages);
  assert.match(printHtml(wide, { orientation: 'auto' }), /@page \{ size: [^;]* landscape;/);

  // Five columns fit upright paper, and 150 rows take fewer sheets that way.
  const narrow = book(150, 5);
  assert.equal(chooseOrientation(narrow, { orientation: 'auto' }), 'portrait');
  assert.equal(printSummary(narrow, { orientation: 'auto' }).setup.orientation, 'portrait');
  assert.match(printHtml(narrow, { orientation: 'auto' }), /@page \{ size: [^;]* portrait;/);

  // A sheet that is one page either way stays upright.
  assert.equal(chooseOrientation(book(10, 4), { orientation: 'auto' }), 'portrait');
});

test('where the sheet is shrunk to fit, the way that shrinks it less is chosen', () => {
  const wide = book(30, 40);
  const tall = printSummary(wide, { orientation: 'portrait', fit: 'width' });
  const turned = printSummary(wide, { orientation: 'landscape', fit: 'width' });
  assert.ok(turned.sheets[0].scale > tall.sheets[0].scale);
  assert.equal(chooseOrientation(wide, { orientation: 'auto', fit: 'width' }), 'landscape');
  // A long narrow list fitted to one page is larger upright.
  const long = book(200, 3);
  assert.equal(chooseOrientation(long, { orientation: 'auto', fit: 'page' }), 'portrait');
  // A print area is what is weighed, not the whole sheet.
  assert.equal(chooseOrientation(book(150, 14), { orientation: 'auto', area: 'A1:C150' }), 'portrait');
});

test('what a person chose stands, and a setup that names no way is upright as it always was', () => {
  const wide = book(150, 14);
  assert.equal(chooseOrientation(wide, { orientation: 'portrait' }), 'portrait');
  assert.equal(chooseOrientation(wide, { orientation: 'landscape' }), 'landscape');
  assert.equal(chooseOrientation(wide, {}), 'portrait');
  assert.equal(printSummary(wide, {}).setup.orientation, 'portrait');
  assert.equal(printSummary(wide, { orientation: 'portrait' }).pages, printSummary(wide, {}).pages);
});

test('a workbook says whether its file names a way, and one nobody chose is not written into it', () => {
  const view = book(150, 14);
  const xml = () => {
    const part = view.workbook.partNameFor('Report');
    return view.workbook.snapshotParts([part])[part];
  };
  assert.equal(readPageSetup(view, 'Report').orientationStated, false);
  assert.equal(readPageSetup(view, 'Report').orientation, 'portrait', 'read as upright, as before');

  // The dialog on Automatic remembers the rest of the setup and leaves the way unsaid.
  writePageSetup(view, 'Report', { paper: 'Letter', orientation: 'auto' });
  assert.match(xml(), /<pageSetup paperSize="1"\/>/);
  assert.equal(readPageSetup(view, 'Report').orientationStated, false);
  assert.equal(readPageSetup(view, 'Report').paper, 'Letter');

  // A print area set from the ribbon sends the setup back as it was read: still nothing said.
  writePageSetup(view, 'Report', { ...readPageSetup(view, 'Report'), area: 'A1:N20' });
  assert.doesNotMatch(xml(), /orientation=/);
  assert.equal(readPageSetup(new SheetView(view.save()), 'Report').orientationStated, false, 'nor after a save');

  // A way a person chose is written, either way, and read back as said.
  writePageSetup(view, 'Report', { ...readPageSetup(view, 'Report'), orientation: 'landscape' });
  assert.match(xml(), /orientation="landscape"/);
  assert.equal(readPageSetup(view, 'Report').orientationStated, true);
  writePageSetup(view, 'Report', { ...readPageSetup(view, 'Report'), orientation: 'portrait' });
  assert.match(xml(), /orientation="portrait"/);
  assert.equal(readPageSetup(view, 'Report').orientationStated, true);
  // Portrait chosen in the ribbon over a file that said nothing: the ribbon says it was chosen.
  const fresh = book(150, 14);
  writePageSetup(fresh, 'Report', { ...readPageSetup(fresh, 'Report'), orientation: 'portrait', orientationStated: true });
  assert.equal(readPageSetup(fresh, 'Report').orientationStated, true);
  // And the dialog, which sends its own choices, writes the one made in it.
  const other = book(150, 14);
  writePageSetup(other, 'Report', { paper: 'A4', orientation: 'portrait' });
  assert.equal(readPageSetup(other, 'Report').orientationStated, true);
  // Back to Automatic takes the saying away again.
  writePageSetup(other, 'Report', { paper: 'A4', orientation: 'auto' });
  assert.equal(readPageSetup(other, 'Report').orientationStated, false);
});

test('slides go on paper turned the way they are, notes and handouts upright, and a choice stands', () => {
  const deck = Deck.open(buildPptx({ title: 'Quarterly', slides: [{ layout: 'title', title: 'Quarterly review' }, { layout: 'titleAndContent', title: 'Two', bullets: ['One'] }] }));
  assert.ok(deck.size.width > deck.size.height, 'a deck is wide as it is made');
  assert.equal(deckPrintSetup({ orientation: 'auto', layout: 'slides' }, deck).orientation, 'landscape');
  assert.equal(deckPrintSetup({ layout: 'slides' }, deck).orientation, 'landscape', 'no way named is no way chosen');
  assert.equal(deckPrintSetup({ orientation: 'auto', layout: 'notes' }, deck).orientation, 'portrait');
  assert.equal(deckPrintSetup({ orientation: 'auto', layout: 'handout' }, deck).orientation, 'portrait');
  assert.equal(deckPrintSetup({ orientation: 'portrait', layout: 'slides' }, deck).orientation, 'portrait');
  assert.equal(deckPrintSetup({ orientation: 'landscape', layout: 'notes' }, deck).orientation, 'landscape');
  assert.equal(deckPrintSummary(deck, { orientation: 'auto' }).setup.orientation, 'landscape');
  assert.match(deckPrintHtml(deck, { orientation: 'auto' }), /@page \{ size: [^;]* landscape;/);
  assert.match(deckPrintHtml(deck, { orientation: 'auto', layout: 'notes' }), /@page \{ size: [^;]* portrait;/);

  // A deck whose slides are taller than they are wide (a poster, a phone's screen) prints them upright.
  deck.setSlideSize(6858000, 12192000);
  assert.ok(deck.size.height > deck.size.width);
  assert.equal(deckPrintSetup({ orientation: 'auto', layout: 'slides' }, deck).orientation, 'portrait');
  assert.equal(deckPrintSummary(deck, { orientation: 'auto' }).setup.orientation, 'portrait');
  assert.match(deckPrintHtml(deck, { orientation: 'auto' }), /@page \{ size: [^;]* portrait;/);
  assert.equal(deckPrintSetup({ orientation: 'auto', layout: 'slides' }).orientation, 'landscape', 'with no deck to go by, a slide is taken to be wide');
});
