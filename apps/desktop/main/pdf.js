// The PDF app's engine, on the main process: the files a window has open,
// read by @rutba/pdf's reader, and new files made from their pages.
//
// A window draws pages itself (pdf.js, in the window) and keeps the list of
// pages it is building, so undo and redo are the window's. What happens
// here is the part that must not redraw anything: the structure of each
// file read once, and on Save, Extract, Split or a recovery copy, a new PDF
// assembled from the pages named, each copied byte for byte.
//
// Every source is held whole in memory while its window has it open, so a
// file can be saved over itself: the new bytes come from the copy held
// here, and are written beside the file and renamed over it.

import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import pdfRead from '@rutba/pdf/reader';
import pdfAssemble from '@rutba/pdf/assemble';

const { openPdf, PdfReadError } = pdfRead;
const { assemble } = pdfAssemble;

/** A sentence for a file that cannot be read, by what went wrong. */
function readFailure(err, name) {
  if (err instanceof PdfReadError && err.code === 'not-pdf') return new Error(`${name} is not a PDF.`);
  if (err instanceof PdfReadError) return new Error(`${name} could not be read as a PDF: ${err.message}`);
  return err;
}

async function writeWhole(target, data) {
  const temp = path.join(path.dirname(target), `.${path.basename(target)}.${process.pid}.saving`);
  const old = await fsp.stat(target).catch(() => null);
  try {
    await fsp.writeFile(temp, data);
    try {
      await fsp.rename(temp, target);
    } catch (err) {
      if (!old || !['EPERM', 'EACCES', 'EBUSY'].includes(err.code)) throw err;
      await fsp.copyFile(temp, target);
    }
  } finally {
    await fsp.rm(temp, { force: true }).catch(() => {});
  }
}

export function createPdfService() {
  let next = 1;
  /** id -> { file, path, name, owner } */
  const open = new Map();

  const forget = (owner) => {
    for (const [id, entry] of open) if (entry.owner === owner) open.delete(id);
  };
  const watched = new Set();
  const watch = (win) => {
    if (!win || watched.has(win.id)) return;
    const id = win.id;
    watched.add(id);
    win.once('closed', () => { watched.delete(id); forget(id); });
  };

  const get = (id) => {
    const entry = open.get(id);
    if (!entry) throw new Error('That PDF is no longer open.');
    return entry;
  };

  const describe = (id, entry) => {
    const { file } = entry;
    const pages = [];
    for (let i = 0; i < file.pageCount; i++) {
      const p = file.pageInfo(i);
      pages.push({ width: p.width, height: p.height, rotate: p.rotate });
    }
    return {
      id,
      name: entry.name,
      path: entry.path,
      size: file.buf.length,
      version: file.version,
      encrypted: file.encrypted,
      repaired: file.repaired,
      title: file.info().title || '',
      pages,
    };
  };

  return {
    /** Read a PDF (from disk, or bytes a window holds) and keep it for its window. */
    open: async ({ path: file = null, bytes = null, name = '' }, win) => {
      const label = name || (file ? path.basename(file) : 'This file');
      const data = bytes ? Buffer.from(bytes) : await fsp.readFile(file);
      let pdf;
      try {
        pdf = openPdf(data);
        pdf.pages();
      } catch (err) {
        throw readFailure(err, label);
      }
      const id = `p${next++}`;
      const entry = { file: pdf, path: file || null, name: label, owner: win?.id ?? null };
      open.set(id, entry);
      watch(win);
      return describe(id, entry);
    },

    close: ({ id }) => ({ closed: open.delete(id) }),

    /**
     * A new PDF of the pages named, in order: [{ source, page, rotate }],
     * `page` 0-based and `rotate` the page's rotation in degrees. Written to
     * `path` when one is given (whole or not at all), else handed back.
     */
    assemble: async ({ pages, path: target = null, title }) => {
      const items = (pages || []).map((p) => ({ file: get(p.source).file, page: Number(p.page), rotate: p.rotate }));
      const bytes = assemble(items, { title });
      if (!target) return { bytes: new Uint8Array(bytes), size: bytes.length };
      await fsp.mkdir(path.dirname(target), { recursive: true });
      await writeWhole(target, bytes);
      return { path: target, size: bytes.length };
    },

    /** The words on one page of an open PDF, as its content draws them. */
    text: ({ id, page }) => get(id).file.pageText(Number(page) || 0),

    /** A file on disk read and described without keeping it: for the checks, and for a page count before inserting. */
    inspect: async ({ path: file }) => {
      const pdf = openPdf(fs.readFileSync(file));
      return {
        pages: Array.from({ length: pdf.pageCount }, (_, i) => ({ ...pdf.pageInfo(i), text: pdf.pageText(i) })),
        encrypted: pdf.encrypted,
        title: pdf.info().title || '',
      };
    },
  };
}
