// Printing, and writing a PDF of anything the suite can open.
//
// A suite that cannot print is not an office suite, and until now nothing
// here could: the Print button in Word called a command that was never
// defined, the one in Worksheets asked for a PDF that was refused, and the
// print namespace on the bridge had no caller at all. This is that half.
//
// The division of labour is deliberate. WHERE the pages are cut — which
// columns fit across, which rows go over, what a slide looks like on paper —
// is a decision about the document, and it is made by the engines
// (sheet-view/print.js, presentation/print.js, doc-view's PDF writer). What
// happens to a laid-out page after that — fonts, page breaks, colour, and the
// conversation with a printer driver — is Chromium's, and nobody should write
// that twice. So each job becomes either a page of HTML or a finished PDF,
// and then goes through one hidden window.
//
// The window is hidden, never shown, and destroyed when the job ends. It is
// not the document's own window: printing what is on screen would print the
// scrollbars and the ribbon, and would print nothing at all of the ninety
// rows below the fold.

import fs from 'node:fs';
import { removeTree } from '@rutba/office-shell/electron/remove-tree';
import os from 'node:os';
import path from 'node:path';
import pdfRead from '@rutba/pdf/reader';
import pdfAssemble from '@rutba/pdf/assemble';
import { memoryFiles } from './memory-files.js';
import { createPrinterFacts, cupsState } from './printer-facts.js';

const { openPdf } = pdfRead;
const { assemble } = pdfAssemble;

/**
 * The files a page an app laid out draws from: every rutba://file address in it, as paths. Read as
 * Chromium reads an address, in whatever case its first part is written.
 */
export function filesDrawn(html) {
  const out = [];
  for (const m of String(html || '').matchAll(/rutba:\/\/file\/([A-Za-z0-9_-]+)/gi)) {
    try {
      out.push(Buffer.from(m[1], 'base64url').toString('utf8'));
    } catch {
      /* not a path: nothing is drawn from it */
    }
  }
  return out;
}

/** How many pages a PDF Chromium wrote has. */
export function pdfPages(bytes) {
  return openPdf(new Uint8Array(bytes)).pageCount;
}

/**
 * The pages a person asked for, of a print of `count` pages: page numbers as they are typed, counted
 * from 1, in the order typed, each once (more of a page is the Copies box). Null when every page is
 * wanted. A number that is not a page of this print is refused in words: the dialog says so before it
 * gets here, so this is the last line and not the first.
 */
export function pagesWanted(pages, count) {
  if (pages == null) return null;
  if (!Array.isArray(pages)) throw new Error('The pages to print are given as page numbers.');
  if (!pages.length) return null;
  const out = [];
  for (const p of pages) {
    const n = Number(p);
    if (!Number.isInteger(n) || n < 1 || n > count) throw new Error(count === 1 ? `There is no page ${p}: this print has one page.` : `There is no page ${p}: this print has ${count} pages.`);
    if (!out.includes(n - 1)) out.push(n - 1);
  }
  return out;
}

/** A PDF of only the pages named (0-based), in that order, each copied as it is. */
export function pdfOfPages(bytes, pages) {
  const file = openPdf(new Uint8Array(bytes));
  return assemble(pages.map((page) => ({ file, page })));
}

/** Papers by name, upright, in millimetres: the ones a print is laid out on. Electron knows each by this name. */
const PAPERS_MM = { A3: [297, 420], A4: [210, 297], A5: [148, 210], Letter: [215.9, 279.4], Legal: [215.9, 355.6], Tabloid: [279.4, 431.8] };

/**
 * The paper the printer is told to load, and which way the pages go on it: the paper the pages were
 * laid out for. Left out, Electron tells every printer A4, whatever was chosen in the dialog, and a
 * page laid out for Letter or A3 comes out on the wrong sheet.
 *
 * A job that is a PDF (a document, or pages cut from any print) says its own paper: page one is
 * measured, so a document on a paper of its own goes to the printer as that size, in microns, and a
 * paper with a name goes by its name. A page of HTML goes by the paper it was laid out with, and so
 * does a PDF whose first page cannot be measured: a print is not refused over its paper.
 */
export function printerPaper({ pdf = null, paper = null, landscape = false } = {}) {
  const chosen = { pageSize: PAPERS_MM[paper] ? paper : 'A4', landscape: Boolean(landscape) };
  if (!pdf) return chosen;
  let page;
  try {
    page = openPdf(new Uint8Array(pdf)).pageInfo(0);
  } catch {
    return chosen;
  }
  const turned = page.rotate === 90 || page.rotate === 270;
  const w = ((turned ? page.height : page.width) * 25.4) / 72;
  const h = ((turned ? page.width : page.height) * 25.4) / 72;
  if (!(w > 0) || !(h > 0)) return chosen;
  const short = Math.min(w, h);
  const long = Math.max(w, h);
  const named = Object.entries(PAPERS_MM).find(([, [pw, ph]]) => Math.abs(pw - short) < 1.5 && Math.abs(ph - long) < 1.5)?.[0];
  return { pageSize: named || { width: Math.round(short * 1000), height: Math.round(long * 1000) }, landscape: w > h + 0.5 };
}

/**
 * What a print came to, for the dialog to say in the person's own language: `reason` is null when it
 * went, 'cancelled' when the person closed the system's dialog, 'settings' when the printer would not
 * have the paper or the settings, and 'failed' for everything else, with what Chromium `said` kept for
 * the log. Chromium's own English ("Print job failed") is never what a person is shown.
 */
export function printAnswer(ok, said) {
  if (ok) return { ok: true, reason: null };
  const text = String(said || '').trim();
  if (!text || /cancel/i.test(text)) return { ok: false, reason: 'cancelled' };
  return { ok: false, reason: /invalid|settings|page ?size/i.test(text) ? 'settings' : 'failed', said: text };
}

/**
 * A job for a document from the Locker never touches the disk (owner, 13 Nov): it is put in the
 * table of files opened in memory for as long as the print takes, and the hidden window loads it
 * over rutba://file, which answers such a file from the table.
 */
function memoryJob(table, name, bytes) {
  const held = table.add({ name, bytes, info: { job: true } });
  return { url: `rutba://file/${Buffer.from(held, 'utf8').toString('base64url')}`, drop: () => table.drop(held) };
}

/** Chromium wants its own copy of the page; a temp file is the simplest one. */
function tempFile(suffix) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-print-'));
  return { dir, file: path.join(dir, `job${suffix}`) };
}

const cleanup = (dir) => {
  try {
    removeTree(dir);
  } catch {
    /* a temp folder that will not go is the operating system's business */
  }
};

export function createPrintService({ docs, holdBlob = null, electron: given = null, facts: knownFacts = null }) {
  // `electron` is for the tests: the hidden window, without a window. `facts` too: Windows, without asking it.
  let electron = given;
  let facts = knownFacts;
  // A document's equations are measured by Chromium before its pages are
  // laid out; a failure there prints the linear form rather than nothing.
  const prepareMath = async (id) => {
    try {
      if (typeof docs.prepareMath === 'function') await docs.prepareMath({ id });
    } catch {
      /* the print path falls back to the linear form */
    }
  };
  const load = async () => (electron ||= await import('electron'));
  let pictureCount = 0;

  /**
   * A window holding the laid-out job, for as long as the job takes.
   *
   * `plugins` because a PDF is shown by Chromium's own viewer, which is how a
   * document — already written as a PDF by the engine that paginates it —
   * reaches a printer without being laid out a second time by something that
   * knows less about it.
   */
  async function withJob(source, fn) {
    const { BrowserWindow } = await load();
    const isPdf = Boolean(source.pdf);
    const bytes = isPdf ? Buffer.from(source.pdf) : Buffer.from(source.html, 'utf8');
    const table = source.locked ? memoryFiles() : null;
    if (source.locked && !table) throw new Error(`${source.name} is in the Locker, and could not be printed without writing it out.`);
    const job = table ? memoryJob(table, isPdf ? 'job.pdf' : 'job.html', bytes) : null;
    const temp = job ? null : tempFile(isPdf ? '.pdf' : '.html');
    if (temp) fs.writeFileSync(temp.file, bytes);
    bytes.fill(0);
    const open = () => (job ? win.loadURL(job.url) : win.loadFile(temp.file));

    const letGo = async () => {
      if (temp) cleanup(temp.dir);
      await job?.drop();
    };
    let win;
    try {
      win = new BrowserWindow({
        show: false,
        webPreferences: {
          // Nothing in a printed page runs: the HTML is ours, and a document
          // that could run something on its way to the printer is a document
          // that can run something.
          javascript: false,
          plugins: isPdf,
          contextIsolation: true,
          sandbox: true,
          nodeIntegration: false,
        },
      });
    } catch (err) {
      await letGo();
      throw err;
    }
    // A page an app laid out stays the page it was handed: nothing in it (a message's own refresh, a
    // link) takes the window anywhere else, and it opens no window of its own.
    if (source.stay) {
      win.webContents.on('will-navigate', (e) => e.preventDefault());
      win.webContents.on('will-redirect', (e) => e.preventDefault());
      win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
    }
    try {
      // Loading a local file straight after another print window was torn
      // down comes back ERR_FAILED now and then — Chromium, not the page: the
      // same file loads on the next try. One retry rather than a failed print
      // job, and a sentence if it still will not.
      try {
        await open();
      } catch (first) {
        await new Promise((r) => setTimeout(r, 200));
        try {
          await open();
        } catch {
          throw new Error(`${source.name} could not be laid out for printing. (${first.message})`);
        }
      }
      // The PDF viewer reports the page loaded before it has drawn; a print
      // that beats it produces a blank sheet.
      if (isPdf) await new Promise((r) => setTimeout(r, 400));
      return await fn(win, { temp: temp ? temp.dir : null });
    } finally {
      try {
        if (!win.isDestroyed()) win.destroy();
      } catch {
        /* already gone */
      }
      await letGo();
    }
  }

  /**
   * A page an app laid out itself (Mail's message, Calendar's view, Contacts' cards, a picture), as
   * a job for the hidden window. The page is the app's own HTML, and the window runs nothing in it.
   * A picture drawn from the Locker makes it a job of the Locker's: held in memory, never written out.
   */
  const pageSource = ({ html, name }) => {
    if (typeof html !== 'string' || !html.trim()) throw new Error('There is nothing to print.');
    const table = memoryFiles();
    const locked = Boolean(table && filesDrawn(html).some((p) => table.isUnder(p) || table.isLocked?.(p)));
    return { html, name: String(name || 'The page'), locked, stay: true };
  };
  const PAGE_PDF = { printBackground: true, preferCSSPageSize: true, margins: { marginType: 'none' } };

  /**
   * The job with only the pages asked for (`pages`: page numbers counted from 1; none means every page,
   * and the job goes as it is). The whole job is made a PDF first: a document is one already, and a page
   * of HTML is laid out by Chromium exactly as the printer would have it. The pages asked for are then
   * copied into a new PDF, and that is what prints. So the printer is handed those pages and no others,
   * whatever the system's own dialog goes on to say about a range.
   */
  const onlyPages = async (source, pages, pdfOptions) => {
    if (pages == null || (Array.isArray(pages) && !pages.length)) return source;
    const whole = source.pdf ? Buffer.from(source.pdf) : Buffer.from(await withJob(source, (win) => win.webContents.printToPDF(pdfOptions)));
    const wanted = pagesWanted(pages, pdfPages(whole));
    return { name: source.name, kind: source.kind, locked: source.locked, stay: source.stay, orientation: source.orientation, pdf: pdfOfPages(whole, wanted) };
  };
  /** The pages asked for: named beside the options, or among them (a window's Save as PDF hands only options on). */
  const pagesOf = (pages, options) => pages ?? options?.pages ?? null;
  /** The options an engine lays out by: the page numbers are the print's, not the layout's. */
  const layoutOf = (options) => {
    const { pages: _pages, ...rest } = options || {};
    return rest;
  };

  /**
   * The job handed to a printer.
   *
   * Office's own dialog is the asking: it names the printer, the copies, the paper, the way round and
   * the pages, and says how many sheets that makes. So a job that comes with a printer named and
   * `silent` goes straight to that printer, and the system's dialog is not shown after ours to ask the
   * same things again. Without `silent` the system's dialog is shown, opened on the same printer and
   * paper: the dialog offers that as its own button, for what only a printer's driver knows (two sides,
   * a tray). A print never starts unasked: `silent` with no printer named shows the system's dialog.
   *
   * The printer is told the paper the pages were laid out for (`printerPaper`).
   */
  const send = (source, { printer, copies, silent, paper, landscape }) => {
    const media = printerPaper({ pdf: source.pdf, paper, landscape });
    return withJob(
      source,
      (win) =>
        new Promise((resolve) => {
          try {
            win.webContents.print(
              {
                silent: Boolean(silent && printer),
                deviceName: printer || undefined,
                printBackground: true,
                copies: Math.max(1, Number(copies) || 1),
                landscape: media.landscape,
                pageSize: media.pageSize,
                margins: { marginType: 'none' },
              },
              (ok, reason) => resolve(printAnswer(ok, reason))
            );
          } catch (err) {
            // Electron refuses a paper it cannot make sense of before it asks the printer anything.
            resolve(printAnswer(false, `Invalid printer settings (${err?.message || err})`));
          }
        })
    );
  };

  return {
    /**
     * How many sheets a page an app laid out takes. Chromium lays it out to say, as it will for the
     * printer: a message or a list has no geometry to count from before it is drawn.
     */
    pageSummary: async ({ html, name, preview = false }) => {
      const source = pageSource({ html, name });
      const bytes = await withJob(source, (win) => win.webContents.printToPDF(PAGE_PDF));
      const pages = pdfPages(bytes);
      // The count is made by laying the page out as a PDF; page one of it is the dialog's small picture
      // of what is about to print, when it asks for one, in place of being thrown away.
      return { kind: 'page', pages, ...(preview && pages ? { first: pdfOfPages(bytes, [0]) } : {}) };
    },

    /** A page an app laid out, written as a PDF: all of it, or the pages asked for. */
    pagePdf: async ({ html, name, path: target, pages = null }) => {
      const source = pageSource({ html, name });
      // What is in the Locker is saved only back into it: a PDF of it somewhere else is the Locker's Send.
      if (source.locked) throw Object.assign(new Error(`${source.name} is in the Locker, so a PDF of it is not written outside it. To have a copy somewhere else, use Unlock or Send in the Locker.`), { code: 'LOCKER_ONLY' });
      const whole = await withJob(source, (win) => win.webContents.printToPDF(PAGE_PDF));
      const wanted = pagesWanted(pages, pdfPages(whole));
      const bytes = wanted ? pdfOfPages(whole, wanted) : whole;
      fs.writeFileSync(target, bytes);
      return { path: target, format: 'pdf', pages: wanted ? wanted.length : pdfPages(bytes) };
    },

    /** A page an app laid out, sent to a printer: asked and answered as `document` is. */
    page: async ({ html, name, options = {}, printer = null, copies = 1, silent = false, pages = null }) => {
      const source = await onlyPages(pageSource({ html, name }), pagesOf(pages, options), PAGE_PDF);
      return send(source, { printer, copies, silent, paper: options.paper, landscape: options.orientation === 'landscape' });
    },

    /**
     * A picture a window made (Image's, as edited), held for the page that prints it: the page names
     * it by address, so its bytes cross once however many times the dialog counts the pages. The
     * window's last one goes when the next arrives, and all of them when the window closes.
     */
    picture: async ({ bytes, type = 'image/png', name = '' }) => {
      if (!holdBlob) throw new Error('This build holds no pictures for the printer.');
      if (!bytes?.length) throw new Error('There is no picture to print.');
      const held = holdBlob(bytes, /^image\/(png|jpeg|webp|gif|bmp)$/.test(type) ? type : 'image/png', String(name || ''), { group: 'print-picture', generation: String(++pictureCount) });
      return { url: held.url, size: held.size };
    },

    /**
     * Pages, scale and paper — what a print dialog says before it prints.
     * Equations are laid out for paper first (documents.js prepareMath):
     * their sizes decide where the pages fall.
     */
    summary: async ({ id, options }) => {
      await prepareMath(id);
      return await docs.printSummary({ id, options });
    },

    /**
     * For the window checks only (verify-locker.js): a window is never offered it, as the contract lists
     * what a window may call. The job laid out as a print lays it out, where the hidden window loaded it
     * from, whether a temp file was made, whether the table held it, and, for a page of HTML, a PDF of it.
     */
    jobProbe: async ({ id, options = {} }) => {
      await prepareMath(id);
      const source = await docs.printSource({ id, options });
      return withJob(source, async (win, { temp }) => ({
        kind: source.pdf ? 'pdf' : 'html',
        locked: Boolean(source.locked),
        url: win.webContents.getURL(),
        temp,
        held: Boolean(memoryFiles()?.entries().some((e) => e.info?.job)),
        printed: source.pdf ? null : (await win.webContents.printToPDF({ printBackground: true, preferCSSPageSize: true, margins: { marginType: 'none' } })).length,
      }));
    },

    /**
     * Page one of a document's print, as a PDF, for the dialog's small picture of what is about to print.
     * A document is a PDF already and its first page is copied out; a workbook and a deck are laid out as
     * a print lays them out and Chromium writes the first sheet only.
     */
    preview: async ({ id, options = {} }) => {
      await prepareMath(id);
      const source = await docs.printSource({ id, options: layoutOf(options) });
      if (source.pdf) {
        // A PDF with no first page to copy has no picture: the print itself is not held up by that.
        try {
          return { first: pdfPages(source.pdf) ? pdfOfPages(source.pdf, [0]) : null };
        } catch {
          return { first: null };
        }
      }
      const landscape = (source.orientation || options.orientation) === 'landscape';
      const one = await withJob(source, (win) => win.webContents.printToPDF({ printBackground: true, preferCSSPageSize: true, landscape, margins: { marginType: 'none' }, pageRanges: '1' }));
      return { first: Buffer.from(one) };
    },

    /**
     * The printers this machine can reach, the default one first, each with its `state`: '' when it is
     * ready or nothing is known, 'offline', 'paused', 'paper' or 'trouble' (printer-facts.js).
     *
     * Electron names the printers and nothing more, so which is the default is read from Windows, at
     * once. The states take Windows a second or two: they are read only when asked for (`states`), which
     * the dialog does after it has drawn its list.
     */
    printers: async (payload, win) => {
      if (!win) return [];
      facts ||= createPrinterFacts();
      const list = await win.webContents.getPrintersAsync();
      const known = payload?.states ? await facts.states() : null;
      const first = known ? Object.keys(known).find((name) => known[name].default) || null : await facts.defaultName();
      return list
        .map((p) => ({
          name: p.name,
          description: p.displayName || p.description || '',
          default: p.name === first,
          state: known ? known[p.name]?.state || '' : cupsState(p.options),
        }))
        .sort((a, b) => Number(b.default) - Number(a.default));
    },

    /**
     * Write the document as a PDF.
     *
     * A document is already one. A workbook and a deck are laid out into
     * pages here and printed to PDF by Chromium, with the page size taken
     * from the CSS the layout wrote rather than from a guess.
     */
    pdf: async ({ id, path: target, options = {}, pages = null }) => {
      // A document from the Locker is saved only back into it: a PDF of it somewhere else is the Locker's Send.
      // Awaited: the document service on its own thread answers every call with a promise.
      await docs.staysInLocker?.({ id, to: target });
      await prepareMath(id);
      const source = await docs.printSource({ id, options: layoutOf(options) });
      const whole = source.pdf
        ? Buffer.from(source.pdf)
        : await withJob(source, (win) =>
          win.webContents.printToPDF({
            printBackground: true,
            preferCSSPageSize: true,
            // The way the pages were laid out: the one chosen, or the one the engine chose when nobody did.
            landscape: (source.orientation || options.orientation) === 'landscape',
            margins: { marginType: 'none' },
          })
        );
      // Only the pages asked for, when some were.
      const asked = pagesOf(pages, options);
      const wanted = asked == null ? null : pagesWanted(asked, pdfPages(whole));
      fs.writeFileSync(target, wanted ? pdfOfPages(whole, wanted) : whole);
      return { path: target, format: 'pdf' };
    },

    /**
     * Send it to a printer: straight to the one named when the dialog says so (`silent`), and through
     * the system's dialog otherwise (`send`, above). The paper, the copies and the pages are the
     * person's to choose, and they were chosen in Office's dialog before this is called.
     */
    document: async ({ id, options = {}, printer = null, copies = 1, silent = false, pages = null }) => {
      await prepareMath(id);
      const laid = await docs.printSource({ id, options: layoutOf(options) });
      // The way the pages were laid out: the one chosen, or the one the engine chose when nobody did.
      const landscape = (laid.orientation || options.orientation) === 'landscape';
      const source = await onlyPages(laid, pagesOf(pages, options), { printBackground: true, preferCSSPageSize: true, landscape, margins: { marginType: 'none' } });
      return send(source, { printer, copies, silent, paper: options.paper, landscape });
    },
  };
}
