// The route a page an app laid out takes to paper (2.27.0: Mail, Calendar, Contacts, Image and
// Pictures print): counted by laying it out, written as a PDF, sent to a printer, all through the
// one hidden window that documents print through. A page that draws a picture from the Locker is
// a job of the Locker's: held in memory while it prints, and never written as a PDF outside it.
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import pdfDocument from '@rutba/pdf/document';
import { tempDir } from './temp-dirs.js';
import { createPrintService, filesDrawn, pdfPages } from '../apps/desktop/main/print.js';
import { createMemoryFiles, setMemoryFiles } from '../apps/desktop/main/memory-files.js';
import { messagePrintHtml } from '../apps/desktop/renderer/apps/mail/print.js';

const { PdfDocument } = pdfDocument;

after(() => setMemoryFiles(null));

/** A PDF of so many pages, as Chromium would hand one back. */
function pdfOf(pages) {
  const doc = new PdfDocument({ title: 'Printed', created: '2026-11-22T09:00:00Z' });
  for (let i = 0; i < pages; i++) doc.addPage().text(`Sheet ${i + 1}`, 72, 100, { font: 'Helvetica', size: 12 });
  return doc.toBuffer();
}

const addressOf = (file) => `rutba://file/${Buffer.from(file, 'utf8').toString('base64url')}`;

/**
 * A stand-in for Electron's BrowserWindow: what the hidden window was made with and asked to load
 * (the page as it stood then), what it was told about leaving the page, and what it printed with.
 */
function fakeElectron({ pages = 3, table = null, printed = [true, ''] } = {}) {
  const seen = { made: [], loads: [], listeners: [], openHandlers: [], prints: [], pdfs: [], destroyed: 0 };
  class BrowserWindow {
    constructor(options) {
      seen.made.push(options);
      this.webContents = {
        on: (name, fn) => seen.listeners.push({ name, fn }),
        setWindowOpenHandler: (fn) => seen.openHandlers.push(fn),
        printToPDF: async (options) => {
          seen.pdfs.push(options);
          return pdfOf(pages);
        },
        print: (options, done) => {
          seen.prints.push(options);
          done(...printed);
        },
        getURL: () => seen.loads.at(-1)?.url || '',
      };
    }
    async loadFile(file) {
      seen.loads.push({ file, html: fs.readFileSync(file, 'utf8') });
    }
    async loadURL(url) {
      const held = Buffer.from(url.replace('rutba://file/', ''), 'base64url').toString('utf8');
      seen.loads.push({ url, held, html: table?.has(held) ? (await table.read(held)).toString() : null });
    }
    isDestroyed() { return false; }
    destroy() { seen.destroyed++; }
  }
  return { seen, electron: { BrowserWindow } };
}

const PAGE = '<!doctype html><html><body><h1>Quarterly figures</h1><img src="rutba://blob/abc"></body></html>';

test('the files a page draws from are read out of its addresses, and a PDF\'s pages are counted', () => {
  const a = path.join('C:', 'Pictures', 'Holiday 1.png');
  const b = path.join('C:', 'Pictures', 'نقشہ.jpg');
  assert.deepEqual(filesDrawn(`<img src="${addressOf(a)}"><img src='${addressOf(b)}?page=2'><img src="rutba://blob/xyz">`), [a, b]);
  assert.deepEqual(filesDrawn(`<img src="${addressOf(a).replace('rutba://file', 'RUTBA://File')}">`), [a], 'in whatever case the address begins');
  assert.deepEqual(filesDrawn(''), []);
  assert.deepEqual(filesDrawn(null), []);
  assert.equal(pdfPages(pdfOf(1)), 1);
  assert.equal(pdfPages(pdfOf(7)), 7);
});

test('the dialog\'s count is the page laid out as the printer will have it', async () => {
  setMemoryFiles(null);
  const { seen, electron } = fakeElectron({ pages: 9 });
  const print = createPrintService({ docs: {}, electron });
  assert.deepEqual(await print.pageSummary({ html: PAGE, name: 'Quarterly figures' }), { kind: 'page', pages: 9 });
  assert.equal(seen.loads.length, 1);
  assert.equal(seen.loads[0].html, PAGE, 'the hidden window is handed the page the app wrote');
  assert.deepEqual(seen.pdfs[0], { printBackground: true, preferCSSPageSize: true, margins: { marginType: 'none' } }, 'the paper and margins are the page\'s own');
  const prefs = seen.made[0].webPreferences;
  assert.equal(seen.made[0].show, false);
  assert.deepEqual([prefs.javascript, prefs.sandbox, prefs.contextIsolation, prefs.nodeIntegration, prefs.plugins], [false, true, true, false, false], 'nothing in a printed page runs');
  assert.equal(seen.destroyed, 1);
  assert.equal(fs.existsSync(path.dirname(seen.loads[0].file)), false, 'the job\'s temp folder is gone when the job ends');
});

test('a page stays the page it was handed: it goes nowhere and opens no window', async () => {
  setMemoryFiles(null);
  const { seen, electron } = fakeElectron();
  await createPrintService({ docs: {}, electron }).pageSummary({ html: PAGE, name: 'A message' });
  assert.deepEqual(seen.listeners.map((l) => l.name).sort(), ['will-navigate', 'will-redirect']);
  for (const { fn } of seen.listeners) {
    let stopped = false;
    fn({ preventDefault: () => { stopped = true; } });
    assert.equal(stopped, true);
  }
  assert.equal(seen.openHandlers.length, 1);
  assert.deepEqual(seen.openHandlers[0]({ url: 'https://example.com' }), { action: 'deny' });
});

test('a page is written as a PDF where the person said, and says how many pages it is', async () => {
  setMemoryFiles(null);
  const dir = tempDir('rl-print-page-');
  const target = path.join(dir, 'Quarterly figures.pdf');
  const { electron } = fakeElectron({ pages: 4 });
  const answer = await createPrintService({ docs: {}, electron }).pagePdf({ html: PAGE, name: 'Quarterly figures', path: target });
  assert.deepEqual(answer, { path: target, format: 'pdf', pages: 4 });
  assert.equal(pdfPages(fs.readFileSync(target)), 4);
});

test('a page goes to a printer asked and answered as a document is', async () => {
  setMemoryFiles(null);
  const { seen, electron } = fakeElectron();
  const print = createPrintService({ docs: {}, electron });
  assert.deepEqual(await print.page({ html: PAGE, name: 'A view' }), { ok: true, reason: null });
  assert.deepEqual(seen.prints[0], { silent: false, deviceName: undefined, printBackground: true, copies: 1, landscape: false, pageSize: 'A4', margins: { marginType: 'none' } });
  await print.page({ html: PAGE, name: 'A view', options: { orientation: 'landscape' }, printer: 'Office Laser', copies: 3, silent: true });
  assert.deepEqual(seen.prints[1], { silent: true, deviceName: 'Office Laser', printBackground: true, copies: 3, landscape: true, pageSize: 'A4', margins: { marginType: 'none' } });
  await print.page({ html: PAGE, name: 'A view', copies: 0, silent: true });
  assert.equal(seen.prints[2].silent, false, 'a print with no printer named always asks');
  assert.equal(seen.prints[2].copies, 1);
  assert.equal(seen.destroyed, 3);

  const stopped = fakeElectron({ printed: [false, ''] });
  assert.deepEqual(await createPrintService({ docs: {}, electron: stopped.electron }).page({ html: PAGE, name: 'A view' }), { ok: false, reason: 'cancelled' });
  const failed = fakeElectron({ printed: [false, 'The printer is out of paper'] });
  assert.deepEqual(await createPrintService({ docs: {}, electron: failed.electron }).page({ html: PAGE, name: 'A view' }), { ok: false, reason: 'failed', said: 'The printer is out of paper' });
});

test('nothing to print is said in words, and no window is made for it', async () => {
  const { seen, electron } = fakeElectron();
  const print = createPrintService({ docs: {}, electron });
  for (const html of ['', '   \n', null, undefined, 42]) {
    await assert.rejects(print.pageSummary({ html, name: 'x' }), /nothing to print/);
    await assert.rejects(print.pagePdf({ html, name: 'x', path: 'nowhere.pdf' }), /nothing to print/);
    await assert.rejects(print.page({ html, name: 'x' }), /nothing to print/);
  }
  assert.equal(seen.made.length, 0);
});

test('a page that draws a picture from the Locker prints from memory, and is never written as a PDF outside it', async () => {
  const home = tempDir('rl-print-page-');
  const table = createMemoryFiles({ root: path.join(home, 'locker', 'mem') });
  setMemoryFiles(table);
  const held = table.add({ name: 'Passport scan.png', bytes: Buffer.from('PICTURE-BYTES') });
  const page = `<!doctype html><html><body><img src="${addressOf(held)}"><p>LOCKED-PICTURE-PAGE</p></body></html>`;
  const { seen, electron } = fakeElectron({ pages: 2, table });
  const print = createPrintService({ docs: {}, electron });

  assert.deepEqual(await print.pageSummary({ html: page, name: 'Passport scan.png' }), { kind: 'page', pages: 2 });
  assert.deepEqual(await print.page({ html: page, name: 'Passport scan.png' }), { ok: true, reason: null });
  assert.equal(seen.loads.length, 2);
  assert.ok(seen.loads.every((l) => !l.file && l.url.startsWith('rutba://file/')), 'loaded over rutba://file, never from a file in the temp folder');
  assert.ok(seen.loads.every((l) => l.html === page), 'the job was in the table while it printed');
  assert.ok(seen.loads.every((l) => path.basename(l.held) === 'job.html'));
  assert.equal(table.entries().some((e) => e.info?.job), false, 'and gone from it after');
  assert.ok(table.has(held), 'the picture itself stays open');

  const target = path.join(home, 'Passport scan.pdf');
  await assert.rejects(print.pagePdf({ html: page, name: 'Passport scan.png', path: target }), (err) => err.code === 'LOCKER_ONLY' && /Passport scan\.png is in the Locker/.test(err.message));
  assert.equal(fs.existsSync(target), false);
  assert.equal(seen.loads.length, 2, 'refused before anything is laid out');

  // A page that draws nothing from the Locker is a plain job, even while the Locker is open.
  const plain = fakeElectron({ table });
  await createPrintService({ docs: {}, electron: plain.electron }).pageSummary({ html: PAGE, name: 'Plain' });
  assert.ok(plain.seen.loads[0].file);
  setMemoryFiles(null);
});

test('a message that names a file in the Locker is a plain job: a sender cannot have it drawn, nor a PDF of the message refused', async () => {
  const home = tempDir('rl-print-page-');
  const table = createMemoryFiles({ root: path.join(home, 'locker', 'mem') });
  setMemoryFiles(table);
  const held = table.add({ name: 'Passport scan.png', bytes: Buffer.from('PICTURE-BYTES') });
  // The sender's own HTML, naming the file as Office itself would, in a picture and in its words.
  const page = messagePrintHtml({ subject: 'Hello', html: `<p><img src="${addressOf(held)}"><img src=${addressOf(held).toUpperCase()}> ${addressOf(held)}</p>` });
  assert.deepEqual(filesDrawn(page), [], 'no file is named in the page the main process is handed');
  assert.ok(!page.toLowerCase().includes('rutba://file/'));
  assert.match(page, /img-src data: rutba:\/\/blob;/, 'and the page may draw held pictures only');
  const { seen, electron } = fakeElectron({ pages: 1, table });
  const print = createPrintService({ docs: {}, electron });
  const target = path.join(home, 'Hello.pdf');
  assert.deepEqual(await print.pagePdf({ html: page, name: 'Hello', path: target }), { path: target, format: 'pdf', pages: 1 });
  assert.ok(seen.loads[0].file, 'a plain job, laid out from a file as any message is');
  setMemoryFiles(null);
});

test('a picture a window made is held for the page that prints it, the last one replaced by the next', async () => {
  const held = [];
  const holdBlob = (bytes, type, name, tags) => {
    held.push({ size: bytes.length, type, name, tags });
    return { url: `rutba://blob/${held.length}`, size: bytes.length };
  };
  const print = createPrintService({ docs: {}, holdBlob });
  const bytes = new Uint8Array([137, 80, 78, 71]);
  assert.deepEqual(await print.picture({ bytes, type: 'image/png', name: 'Poster.png' }), { url: 'rutba://blob/1', size: 4 });
  assert.deepEqual(await print.picture({ bytes, type: 'image/svg+xml' }), { url: 'rutba://blob/2', size: 4 });
  await print.picture({ bytes, type: 'image/jpeg', name: 7 });
  assert.deepEqual(held, [
    { size: 4, type: 'image/png', name: 'Poster.png', tags: { group: 'print-picture', generation: '1' } },
    { size: 4, type: 'image/png', name: '', tags: { group: 'print-picture', generation: '2' } },
    { size: 4, type: 'image/jpeg', name: '7', tags: { group: 'print-picture', generation: '3' } },
  ], 'a type that could carry script is held as a plain picture');
  await assert.rejects(print.picture({ bytes: new Uint8Array(0) }), /no picture to print/);
  await assert.rejects(print.picture({}), /no picture to print/);
  await assert.rejects(createPrintService({ docs: {} }).picture({ bytes }), /holds no pictures/);
});
