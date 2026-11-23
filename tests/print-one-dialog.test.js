// What the print dialog was found to do wrong on 22 Nov, and no longer does (2.27.0).
//
// It asked twice: Office's dialog, and then the system's, for the same printer, copies and paper.
// Now a print goes straight to the printer chosen in Office's dialog, and the system's dialog is
// shown only when it is asked for.
//
// The paper chosen never reached the printer: Electron was told no paper and told every printer
// A4. Now the printer is told the paper the pages were laid out for.
//
// It said "Sent 0 pages" when the pages were not counted yet, showed Chromium's own English when
// a print failed, and opened on whichever printer sorted first, because Electron 44 names no
// default printer and gives no printer's state. Now the default and the states are read from
// Windows, a failure is said in the suite's own words, and a number is said only when it is known.
//
// And it names what is about to print, with page one of it as a small picture: the PDF the count
// is made from, kept in place of being thrown away.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import pdfDocument from '@rutba/pdf/document';
import pdfRead from '@rutba/pdf/reader';
import { createPrintService, printerPaper, printAnswer, pdfPages } from '../apps/desktop/main/print.js';
import { createPrinterFacts, defaultPrinterOf, printerState, printerStatesOf, cupsState } from '../apps/desktop/main/printer-facts.js';
import { setMemoryFiles } from '../apps/desktop/main/memory-files.js';
import { printerToOpenOn, printOutcome } from '../apps/desktop/renderer/print-page.js';

const { PdfDocument } = pdfDocument;
const { openPdf } = pdfRead;

after(() => setMemoryFiles(null));

const MM = 72 / 25.4;

/** A PDF of so many pages on a paper of `size` points, each saying which it is. */
function pdfOf(pages, size = null) {
  const doc = new PdfDocument({ title: 'Printed', created: '2026-11-23T09:00:00Z' });
  for (let i = 0; i < pages; i++) doc.addPage(size).text(`Sheet ${i + 1}`, 72, 100, { font: 'Helvetica', size: 12 });
  return doc.toBuffer();
}
const wordsOf = (bytes) => {
  const file = openPdf(new Uint8Array(bytes));
  return Array.from({ length: file.pageCount }, (_, i) => file.pageText(i).trim());
};

/**
 * A stand-in for Electron: each hidden window says what it printed with and what it was asked to
 * make a PDF with. Nothing here reaches a printer. `printed` is what the print calls back with,
 * `throws` what it throws before asking the printer anything, `printers` what the system lists.
 */
function fakeElectron({ pages = 6, size = null, printed = [true, ''], throws = null } = {}) {
  const seen = { prints: [], pdfs: [], loads: 0, destroyed: 0 };
  class BrowserWindow {
    constructor() {
      this.webContents = {
        on: () => {},
        setWindowOpenHandler: () => {},
        printToPDF: async (options) => {
          seen.pdfs.push(options);
          return pdfOf(options.pageRanges === '1' ? 1 : pages, size);
        },
        print: (options, done) => {
          seen.prints.push(options);
          if (throws) throw new Error(throws);
          done(...printed);
        },
      };
    }
    async loadFile(file) {
      seen.loads += 1;
      assert.ok(fs.existsSync(file));
    }
    isDestroyed() { return false; }
    destroy() { seen.destroyed += 1; }
  }
  return { seen, electron: { BrowserWindow } };
}

const PAGE = '<!doctype html><html><body><p>A calendar, laid out by its app.</p></body></html>';

/* ── the paper the printer is told ──────────────────────────────────────── */

test('a page of HTML goes to the printer on the paper it was laid out with', () => {
  assert.deepEqual(printerPaper({ paper: 'Letter', landscape: false }), { pageSize: 'Letter', landscape: false });
  assert.deepEqual(printerPaper({ paper: 'A3', landscape: true }), { pageSize: 'A3', landscape: true });
  assert.deepEqual(printerPaper({ paper: 'Tabloid' }), { pageSize: 'Tabloid', landscape: false }, 'a workbook whose file says Tabloid');
  assert.deepEqual(printerPaper({ paper: 'Foolscap' }), { pageSize: 'A4', landscape: false }, 'a paper nobody lays out for is A4, as the page was laid out');
  assert.deepEqual(printerPaper({}), { pageSize: 'A4', landscape: false });
});

test('a PDF says its own paper: page one is measured, and the way it is turned', () => {
  assert.deepEqual(printerPaper({ pdf: pdfOf(2, 'A4'), paper: 'Letter' }), { pageSize: 'A4', landscape: false }, 'what the pages are, not what the options say');
  assert.deepEqual(printerPaper({ pdf: pdfOf(1, 'Legal') }), { pageSize: 'Legal', landscape: false });
  assert.deepEqual(printerPaper({ pdf: pdfOf(1, [841.89, 595.28]) }), { pageSize: 'A4', landscape: true }, 'A4 turned sideways is A4, landscape');
  assert.deepEqual(printerPaper({ pdf: pdfOf(1, [1190.55, 841.89]) }), { pageSize: 'A3', landscape: true });
  // A paper of the document's own: 100 by 150 millimetres, told to the printer in microns, upright.
  assert.deepEqual(printerPaper({ pdf: pdfOf(1, [100 * MM, 150 * MM]) }), { pageSize: { width: 100000, height: 150000 }, landscape: false });
  assert.deepEqual(printerPaper({ pdf: pdfOf(1, [150 * MM, 100 * MM]) }), { pageSize: { width: 100000, height: 150000 }, landscape: true });
});

test('a PDF that cannot be measured prints on the paper chosen, and is not refused', () => {
  assert.deepEqual(printerPaper({ pdf: Buffer.from('%PDF-not-one'), paper: 'Letter', landscape: true }), { pageSize: 'Letter', landscape: true });
});

/* ── one dialog, and the paper reaching the printer ─────────────────────── */

test('a page an app laid out goes straight to the printer chosen, on the paper chosen', async () => {
  const { seen, electron } = fakeElectron();
  const print = createPrintService({ docs: {}, electron });
  const answer = await print.page({ html: PAGE, name: 'November', options: { paper: 'Letter', orientation: 'landscape' }, printer: 'Office Laser', copies: 2, silent: true });
  assert.deepEqual(answer, { ok: true, reason: null });
  assert.deepEqual(seen.prints[0], { silent: true, deviceName: 'Office Laser', printBackground: true, copies: 2, landscape: true, pageSize: 'Letter', margins: { marginType: 'none' } });
});

test('the system dialog is shown only when asked for, opened on the same printer and paper', async () => {
  const { seen, electron } = fakeElectron();
  const print = createPrintService({ docs: {}, electron });
  await print.page({ html: PAGE, name: 'November', options: { paper: 'A5' }, printer: 'Office Laser', silent: false });
  assert.equal(seen.prints[0].silent, false);
  assert.equal(seen.prints[0].deviceName, 'Office Laser');
  assert.equal(seen.prints[0].pageSize, 'A5');
  // A print never starts unasked: with no printer named there is nothing to send it straight to.
  await print.page({ html: PAGE, name: 'November', silent: true });
  assert.equal(seen.prints[1].silent, false);
  assert.equal(seen.prints[1].pageSize, 'A4');
});

test('a workbook goes on the paper and the way it was laid out; a document on its own', async () => {
  const sheet = fakeElectron();
  const laidOut = [];
  const sheets = createPrintService({
    docs: { printSource: ({ options }) => { laidOut.push(options); return { kind: 'sheet', name: 'Budget.xlsx', orientation: 'landscape', html: PAGE }; } },
    electron: sheet.electron,
  });
  // Automatic, and the sheet's shape asked for paper turned sideways: the printer is told what the engine chose.
  assert.deepEqual(await sheets.document({ id: 's', options: { paper: 'A3', orientation: 'auto' }, printer: 'Office Laser', copies: 1, silent: true }), { ok: true, reason: null });
  assert.deepEqual(sheet.seen.prints[0], { silent: true, deviceName: 'Office Laser', printBackground: true, copies: 1, landscape: true, pageSize: 'A3', margins: { marginType: 'none' } });
  assert.deepEqual(laidOut[0], { paper: 'A3', orientation: 'auto' });

  // A document is a PDF, on Legal whatever the dialog's options still say.
  const doc = fakeElectron();
  const docs = createPrintService({ docs: { printSource: () => ({ kind: 'doc', name: 'Lease.docx', pdf: new Uint8Array(pdfOf(3, 'Legal')) }) }, electron: doc.electron });
  await docs.document({ id: 'd', options: { paper: 'A4', orientation: 'portrait' }, printer: 'Office Laser', silent: true });
  assert.equal(doc.seen.prints[0].pageSize, 'Legal');
  assert.equal(doc.seen.prints[0].landscape, false);
  assert.equal(doc.seen.prints[0].silent, true);
});

test('pages cut from a print go on the paper the whole was laid out on', async () => {
  const { seen, electron } = fakeElectron({ pages: 6, size: [792, 612] });
  const print = createPrintService({ docs: {}, electron });
  await print.page({ html: PAGE, name: 'November', options: { paper: 'Letter', orientation: 'landscape' }, printer: 'Office Laser', silent: true, pages: [2, 4] });
  assert.equal(seen.pdfs.length, 1, 'laid out once, as a PDF, to cut the pages from');
  assert.deepEqual(seen.prints[0].pageSize, 'Letter');
  assert.equal(seen.prints[0].landscape, true);
});

/* ── what a print came to ───────────────────────────────────────────────── */

test('what Chromium says of a print is turned into a reason, never shown as it is', () => {
  assert.deepEqual(printAnswer(true, ''), { ok: true, reason: null });
  assert.deepEqual(printAnswer(false, ''), { ok: false, reason: 'cancelled' });
  assert.deepEqual(printAnswer(false, 'Print job canceled'), { ok: false, reason: 'cancelled' });
  assert.deepEqual(printAnswer(false, 'Print job cancelled'), { ok: false, reason: 'cancelled' });
  assert.deepEqual(printAnswer(false, 'Print job failed'), { ok: false, reason: 'failed', said: 'Print job failed' });
  assert.deepEqual(printAnswer(false, 'Invalid printer settings'), { ok: false, reason: 'settings', said: 'Invalid printer settings' });
});

test('a print that fails, is cancelled or is refused its paper answers with the reason', async () => {
  const failed = fakeElectron({ printed: [false, 'Print job failed'] });
  assert.deepEqual(await createPrintService({ docs: {}, electron: failed.electron }).page({ html: PAGE, name: 'x', printer: 'Office Laser', silent: true }), { ok: false, reason: 'failed', said: 'Print job failed' });
  const stopped = fakeElectron({ printed: [false, 'Print job canceled'] });
  assert.deepEqual(await createPrintService({ docs: {}, electron: stopped.electron }).page({ html: PAGE, name: 'x' }), { ok: false, reason: 'cancelled' });
  // Electron throws over a paper it cannot make sense of, before the printer is asked anything.
  const refused = fakeElectron({ throws: 'Unsupported pageSize: Foolscap' });
  const answer = await createPrintService({ docs: {}, electron: refused.electron }).page({ html: PAGE, name: 'x', printer: 'Office Laser', silent: true });
  assert.equal(answer.ok, false);
  assert.equal(answer.reason, 'settings');
  assert.equal(refused.seen.destroyed, 1, 'and its window is gone all the same');
});

test('the dialog says a number only when it knows one, and never "0 pages"', () => {
  assert.deepEqual(printOutcome({ ok: true, reason: null }, { count: 3 }), { kind: 'sent-count', count: 3, done: true });
  assert.deepEqual(printOutcome({ ok: true, reason: null }, { count: null }), { kind: 'sent', done: true }, 'not counted yet: no number');
  assert.deepEqual(printOutcome({ ok: true, reason: null }, { count: 0 }), { kind: 'sent', done: true }, 'nought is not a number of pages sent');
  assert.deepEqual(printOutcome({ ok: true, reason: null }, {}), { kind: 'sent', done: true });
  // Through the system's dialog the printer and the pages could both be changed: neither is said.
  assert.deepEqual(printOutcome({ ok: true, reason: null }, { system: true, count: 3 }), { kind: 'sent-system', done: true });
});

test('a print that did not go leaves the dialog open, and a cancelled one says nothing', () => {
  assert.deepEqual(printOutcome({ ok: false, reason: 'cancelled' }, { system: true }), { kind: 'cancelled', done: false });
  assert.deepEqual(printOutcome({ ok: false, reason: 'failed', said: 'Print job failed' }, { count: 2 }), { kind: 'failed', done: false });
  assert.deepEqual(printOutcome({ ok: false, reason: 'settings' }, { count: 2 }), { kind: 'settings', done: false });
  assert.deepEqual(printOutcome({ ok: false, reason: 'Something new from Chromium' }, {}), { kind: 'failed', done: false }, 'a reason nobody knows is a failure in our words');
  assert.deepEqual(printOutcome(undefined, {}), { kind: 'cancelled', done: false });
});

/* ── the small picture of page one ──────────────────────────────────────── */

test('the count of a page an app laid out keeps page one of the PDF it made, when asked', async () => {
  const { seen, electron } = fakeElectron({ pages: 4 });
  const print = createPrintService({ docs: {}, electron });
  const plain = await print.pageSummary({ html: PAGE, name: 'November' });
  assert.deepEqual(plain, { kind: 'page', pages: 4 }, 'not asked for, nothing more crosses');
  const pictured = await print.pageSummary({ html: PAGE, name: 'November', preview: true });
  assert.equal(pictured.pages, 4);
  assert.deepEqual(wordsOf(pictured.first), ['Sheet 1']);
  assert.equal(seen.pdfs.length, 2, 'one layout each: the picture comes from the PDF the count made');
});

test('page one of a document is copied out of its PDF; of a workbook, laid out as the print is', async () => {
  const doc = fakeElectron();
  const docs = createPrintService({ docs: { printSource: () => ({ kind: 'doc', name: 'Lease.docx', pdf: new Uint8Array(pdfOf(14, 'Legal')) }) }, electron: doc.electron });
  const first = (await docs.preview({ id: 'd', options: { paper: 'A4' } })).first;
  assert.deepEqual(wordsOf(first), ['Sheet 1']);
  assert.deepEqual(printerPaper({ pdf: first }), { pageSize: 'Legal', landscape: false }, 'on the document’s own paper');
  assert.equal(doc.seen.pdfs.length, 0, 'no window: a document is a PDF already');

  const sheet = fakeElectron({ pages: 9 });
  const laidOut = [];
  const sheets = createPrintService({
    docs: { printSource: ({ options }) => { laidOut.push(options); return { kind: 'sheet', name: 'Budget.xlsx', orientation: 'landscape', html: PAGE }; } },
    electron: sheet.electron,
  });
  const one = (await sheets.preview({ id: 's', options: { paper: 'A4', orientation: 'auto', pages: [3] } })).first;
  assert.equal(pdfPages(one), 1);
  assert.deepEqual(sheet.seen.pdfs[0], { printBackground: true, preferCSSPageSize: true, landscape: true, margins: { marginType: 'none' }, pageRanges: '1' }, 'Chromium writes the first sheet only, turned as the engine chose');
  assert.deepEqual(laidOut[0], { paper: 'A4', orientation: 'auto' }, 'laid out by the choices, the page numbers aside');
  assert.equal(sheet.seen.destroyed, 1);

  const unread = createPrintService({ docs: { printSource: () => ({ kind: 'doc', name: 'Odd.docx', pdf: new Uint8Array(Buffer.from('%PDF-not-one')) }) }, electron: doc.electron });
  assert.deepEqual(await unread.preview({ id: 'e' }), { first: null }, 'a PDF with no page to copy has no picture, and that is not an error');
});

/* ── which printer, and whether it is ready ─────────────────────────────── */

const REG = '\r\nHKEY_CURRENT_USER\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows\r\n    Device    REG_SZ    Office Laser (HP LaserJet P4015),winspool,Ne05:\r\n\r\n';

test('the default printer is read from what the registry says', () => {
  assert.equal(defaultPrinterOf(REG), 'Office Laser (HP LaserJet P4015)');
  assert.equal(defaultPrinterOf('    Device    REG_SZ    طابعة المكتب,winspool,Ne01:\r\n'), 'طابعة المكتب', 'a name in any writing');
  assert.equal(defaultPrinterOf('    Device    REG_SZ    ,winspool,Ne01:'), null, 'no name is no default');
  assert.equal(defaultPrinterOf(''), null);
  assert.equal(defaultPrinterOf(null), null);
  assert.equal(defaultPrinterOf('ERROR: The system was unable to find the specified registry key or value.'), null);
});

test('a printer’s state is one word, from what Windows says of it', () => {
  assert.equal(printerState({}), '');
  assert.equal(printerState({ bits: 0, extended: 2, error: 0 }), '', 'idle, nothing wrong known');
  assert.equal(printerState({ bits: 0x400, extended: 4 }), '', 'printing is ready');
  assert.equal(printerState({ bits: 128, extended: 7, error: 9 }), 'offline');
  assert.equal(printerState({ offline: true, extended: 2 }), 'offline', 'set to be used offline');
  assert.equal(printerState({ bits: 0x1000 }), 'offline', 'not available');
  assert.equal(printerState({ bits: 0x1 }), 'paused');
  assert.equal(printerState({ extended: 8 }), 'paused');
  assert.equal(printerState({ bits: 0x10 }), 'paper');
  assert.equal(printerState({ error: 4 }), 'paper');
  assert.equal(printerState({ bits: 0x8 }), 'trouble', 'a jam');
  assert.equal(printerState({ bits: 0x400000 }), 'trouble', 'a door open');
  assert.equal(printerState({ error: 6 }), 'trouble', 'no toner');
  assert.equal(printerState({ bits: 0x80 | 0x10 }), 'offline', 'offline first: a printer that cannot be reached says nothing true of its paper');
  assert.equal(printerState({ bits: 0x20000 }), '', 'toner low still prints');
});

test('the states Windows gives are read by printer name', () => {
  const text = JSON.stringify([
    { n: 'Label printer', d: false, w: true, s: 0, x: 2, e: 0 },
    { n: 'Office Laser', d: true, w: false, s: 128, x: 7, e: 9 },
    { n: 'Print to PDF', d: false, w: false, s: 0, x: 2, e: 0 },
  ]);
  assert.deepEqual(printerStatesOf(text), {
    'Label printer': { default: false, state: 'offline' },
    'Office Laser': { default: true, state: 'offline' },
    'Print to PDF': { default: false, state: '' },
  });
  assert.deepEqual(printerStatesOf('{"n":"Only one","d":true,"w":false,"s":16,"x":2,"e":0}'), { 'Only one': { default: true, state: 'paper' } }, 'one printer, not in a list');
  assert.deepEqual(printerStatesOf('[]'), {});
  assert.equal(printerStatesOf(''), null);
  assert.equal(printerStatesOf('Get-CimInstance : Access denied'), null, 'an answer that is not one is nothing known');
  assert.equal(printerStatesOf(null), null);
});

test('Windows is asked through the two runners, and a read that fails is nothing known', async () => {
  const asked = [];
  const facts = createPrinterFacts({
    platform: 'win32',
    reg: async (args) => { asked.push(args); return REG; },
    spooler: async () => JSON.stringify([{ n: 'Office Laser', d: true, w: false, s: 1, x: 8, e: 0 }]),
  });
  assert.equal(await facts.defaultName(), 'Office Laser (HP LaserJet P4015)');
  assert.deepEqual(asked[0], ['query', 'HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows', '/v', 'Device']);
  assert.deepEqual(await facts.states(), { 'Office Laser': { default: true, state: 'paused' } });

  const silent = createPrinterFacts({ platform: 'win32', reg: async () => null, spooler: async () => null });
  assert.equal(await silent.defaultName(), null, 'a reg.exe that never answered');
  assert.equal(await silent.states(), null, 'a PowerShell that never answered');
  const broken = createPrinterFacts({ platform: 'win32', reg: async () => { throw new Error('no'); }, spooler: async () => { throw new Error('no'); } });
  assert.equal(await broken.defaultName(), null);
  assert.equal(await broken.states(), null);

  // On another system nothing is asked.
  let ran = 0;
  const other = createPrinterFacts({ platform: 'linux', reg: async () => { ran += 1; return REG; }, spooler: async () => { ran += 1; return '[]'; } });
  assert.equal(await other.defaultName(), null);
  assert.equal(await other.states(), null);
  assert.equal(ran, 0);
});

test('the system is asked once at a time: a second dialog shares the answer still on its way', async () => {
  let runs = 0;
  let answer;
  const facts = createPrinterFacts({
    platform: 'win32',
    reg: async () => REG,
    spooler: () => { runs += 1; return new Promise((resolve) => { answer = resolve; }); },
  });
  const one = facts.states();
  const two = facts.states();
  assert.equal(runs, 1, 'two questions together make one PowerShell');
  answer(JSON.stringify([{ n: 'Office Laser', d: true, w: true, s: 0, x: 0, e: 0 }]));
  assert.deepEqual(await one, { 'Office Laser': { default: true, state: 'offline' } });
  assert.deepEqual(await two, await one);

  // Asked again afterwards, Windows is asked again: a printer switched on meanwhile shows as ready.
  const again = facts.states();
  assert.equal(runs, 2);
  answer(JSON.stringify([{ n: 'Office Laser', d: true, w: false, s: 0, x: 2, e: 0 }]));
  assert.deepEqual(await again, { 'Office Laser': { default: true, state: '' } });
});

test('what CUPS says of a printer is the same one word', () => {
  assert.equal(cupsState({}), '');
  assert.equal(cupsState(undefined), '');
  assert.equal(cupsState({ 'printer-state': '3', 'printer-state-reasons': 'none' }), '');
  assert.equal(cupsState({ 'printer-state': '5', 'printer-state-reasons': 'paused' }), 'paused');
  assert.equal(cupsState({ 'printer-state': '3', 'printer-state-reasons': 'offline-report' }), 'offline');
  assert.equal(cupsState({ 'printer-state-reasons': 'media-empty-error' }), 'paper');
  assert.equal(cupsState({ 'printer-state-reasons': 'media-jam-error' }), 'trouble');
});

/** A window whose system lists these printers, in the order Electron 44 gives them: by nothing a person would choose. */
const windowWith = (names) => ({ webContents: { getPrintersAsync: async () => names.map((name) => ({ name, displayName: name, description: '', options: {} })) } });
const LISTED = ['Label printer', 'Fax', 'Office Laser', 'Print to PDF'];

test('the list of printers names the default one first, at once, without the slow read', async () => {
  let slow = 0;
  const facts = { defaultName: async () => 'Office Laser', states: async () => { slow += 1; return {}; } };
  const print = createPrintService({ docs: {}, electron: fakeElectron().electron, facts });
  const list = await print.printers(undefined, windowWith(LISTED));
  assert.deepEqual(list.map((p) => p.name), ['Office Laser', 'Label printer', 'Fax', 'Print to PDF']);
  assert.deepEqual(list.map((p) => p.default), [true, false, false, false]);
  assert.deepEqual(list.map((p) => p.state), ['', '', '', ''], 'unmarked until the states are asked for');
  assert.equal(slow, 0);
  assert.equal(printerToOpenOn(list), 'Office Laser', 'and that is the printer the dialog opens on');
  assert.deepEqual(await print.printers(undefined, null), [], 'no window, no printers');
});

test('asked for the states, each printer says whether it is ready', async () => {
  const facts = {
    defaultName: async () => { throw new Error('not asked when the slow read says which is the default'); },
    states: async () => ({ 'Label printer': { default: false, state: 'offline' }, 'Office Laser': { default: true, state: 'paper' }, 'Print to PDF': { default: false, state: '' } }),
  };
  const print = createPrintService({ docs: {}, electron: fakeElectron().electron, facts });
  const list = await print.printers({ states: true }, windowWith(LISTED));
  assert.deepEqual(list, [
    { name: 'Office Laser', description: 'Office Laser', default: true, state: 'paper' },
    { name: 'Label printer', description: 'Label printer', default: false, state: 'offline' },
    { name: 'Fax', description: 'Fax', default: false, state: '' },
    { name: 'Print to PDF', description: 'Print to PDF', default: false, state: '' },
  ]);
});

test('when Windows says nothing, the printers are listed unmarked and the dialog opens on the first', async () => {
  const facts = { defaultName: async () => null, states: async () => null };
  const print = createPrintService({ docs: {}, electron: fakeElectron().electron, facts });
  for (const payload of [undefined, { states: true }]) {
    const list = await print.printers(payload, windowWith(LISTED));
    assert.deepEqual(list.map((p) => p.name), LISTED);
    assert.ok(list.every((p) => !p.default && p.state === ''));
    assert.equal(printerToOpenOn(list), 'Label printer');
  }
  assert.equal(printerToOpenOn([]), '', 'no printer installed');
  assert.equal(printerToOpenOn(null), '');
});
