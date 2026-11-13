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
import { memoryFiles } from './memory-files.js';

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

export function createPrintService({ docs, electron: given = null }) {
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

  return {
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
    pdf: async ({ id, path: target, options = {} }) => {
      // A document from the Locker is saved only back into it: a PDF of it somewhere else is the Locker's Send.
      // Awaited: the document service on its own thread answers every call with a promise.
      await docs.staysInLocker?.({ id, to: target });
      await prepareMath(id);
      const source = await docs.printSource({ id, options });
      if (source.pdf) {
        fs.writeFileSync(target, Buffer.from(source.pdf));
        return { path: target, format: 'pdf' };
      }
      const bytes = await withJob(source, (win) =>
        win.webContents.printToPDF({
          printBackground: true,
          preferCSSPageSize: true,
          landscape: options.orientation === 'landscape',
          margins: { marginType: 'none' },
        })
      );
      fs.writeFileSync(target, bytes);
      return { path: target, format: 'pdf' };
    },

    /**
     * Send it to a printer.
     *
     * The system dialog is shown unless a printer is named outright, because
     * a print that starts without asking is a print somebody did not want:
     * the paper, the copies and the range are the person's to choose.
     */
    document: async ({ id, options = {}, printer = null, copies = 1, silent = false }) => {
      await prepareMath(id);
      const source = await docs.printSource({ id, options });
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
  };
}
