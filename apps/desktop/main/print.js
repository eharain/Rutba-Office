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

export function createPrintService({ docs, holdBlob = null, electron: given = null }) {
  // `electron` is for the tests: the hidden window, without a window.
  let electron = given;
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

  return {
    /**
     * How many sheets a page an app laid out takes. Chromium lays it out to say, as it will for the
     * printer: a message or a list has no geometry to count from before it is drawn.
     */
    pageSummary: async ({ html, name }) => {
      const source = pageSource({ html, name });
      const bytes = await withJob(source, (win) => win.webContents.printToPDF(PAGE_PDF));
      return { kind: 'page', pages: pdfPages(bytes) };
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
      return withJob(
        source,
        (win) =>
          new Promise((resolve) => {
            win.webContents.print(
              {
                silent: Boolean(silent && printer),
                deviceName: printer || undefined,
                printBackground: true,
                copies: Math.max(1, Number(copies) || 1),
                landscape: options.orientation === 'landscape',
                margins: { marginType: 'none' },
              },
              (ok, reason) => resolve({ ok, reason: ok ? null : reason || 'cancelled' })
            );
          })
      );
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

    /** The printers this machine can reach, the default one first. */
    printers: async (_payload, win) => {
      if (!win) return [];
      const list = await win.webContents.getPrintersAsync();
      return list
        .map((p) => ({ name: p.name, description: p.displayName || p.description || '', status: p.status, default: Boolean(p.isDefault) }))
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
     * Send it to a printer.
     *
     * The system dialog is shown unless a printer is named outright, because
     * a print that starts without asking is a print somebody did not want:
     * the paper, the copies and the range are the person's to choose.
     */
    document: async ({ id, options = {}, printer = null, copies = 1, silent = false, pages = null }) => {
      await prepareMath(id);
      const laid = await docs.printSource({ id, options: layoutOf(options) });
      // The way the pages were laid out: the one chosen, or the one the engine chose when nobody did.
      const landscape = (laid.orientation || options.orientation) === 'landscape';
      const source = await onlyPages(laid, pagesOf(pages, options), { printBackground: true, preferCSSPageSize: true, landscape, margins: { marginType: 'none' } });
      return withJob(
        source,
        (win) =>
          new Promise((resolve) => {
            win.webContents.print(
              {
                silent: Boolean(silent && printer),
                deviceName: printer || undefined,
                printBackground: true,
                copies: Math.max(1, Number(copies) || 1),
                landscape,
                margins: { marginType: 'none' },
              },
              (ok, reason) => resolve({ ok, reason: ok ? null : reason || 'cancelled' })
            );
          })
      );
    },
  };
}
