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
//
// Recognise Text adds an invisible text layer to scanned pages as an
// incremental update (@rutba/pdf/textlayer): the file's bytes with the
// update after them become a new source, the old one kept for Undo. A Save
// of a source's own pages, all of them, in order and unturned, writes those
// bytes as they are, so the original stays the exact start of the file.
//
// Comments are written the same way (@rutba/pdf/annotations): each change
// an incremental update, the file with it a new source, the old one kept
// for Undo. So are a form's values (@rutba/pdf/forms) and signatures placed
// as stamps.
//
// Redaction is the one change written whole (@rutba/pdf/redact): an update
// would keep the old bytes, the words under the marks among them. Applying
// writes a new file of the window's pages with what the marks covered taken
// out of each page's content, and the window opens it.

import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import pdfRead from '@rutba/pdf/reader';
import pdfAssemble from '@rutba/pdf/assemble';
import pdfTextLayer from '@rutba/pdf/textlayer';
import pdfAnnotations from '@rutba/pdf/annotations';
import pdfForms from '@rutba/pdf/forms';
import pdfRedact from '@rutba/pdf/redact';
import pdfNavigation from '@rutba/pdf/navigation';

const { openPdf, PdfReadError } = pdfRead;
const { assemble } = pdfAssemble;
const { addTextLayer } = pdfTextLayer;
const { readAnnotations, annotate } = pdfAnnotations;
const { readFields, fillFields } = pdfForms;
const { applyRedactions, redactionPlan } = pdfRedact;
const { outlinePages, readLinks, pageLabels, labelText } = pdfNavigation;

/** Who a comment is signed by: the account at the keyboard, as the other apps sign theirs. */
function userName() {
  try {
    return os.userInfo().username || 'Rutba Office user';
  } catch {
    return 'Rutba Office user';
  }
}

/** Whether the pages named are one source's own, every one, in order, each as the file turns it. */
function wholeSource(items) {
  const file = items[0]?.file;
  if (!file || items.some((it) => it.file !== file) || items.length !== file.pageCount) return false;
  return items.every((it, i) => it.page === i && (it.rotate === undefined || it.rotate === null || Number(it.rotate) === file.pageInfo(i).rotate));
}

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
      // A recognised file saved as it is: its bytes, the text layer appended to the original.
      const own = (pages || []).length && open.get(pages[0].source)?.updated && wholeSource(items)
        && (title === undefined || title === null || title === items[0].file.info().title);
      const bytes = own ? items[0].file.buf : assemble(items, { title });
      if (!target) return { bytes: new Uint8Array(bytes), size: bytes.length };
      await fsp.mkdir(path.dirname(target), { recursive: true });
      await writeWhole(target, bytes);
      return { path: target, size: bytes.length };
    },

    /**
     * Recognise Text's answer written into the file: an invisible text layer
     * on the pages named ([{ page, lines: [{ words: [{ text, x, y, width,
     * size, angle }] }] }], in each page's own coordinates) as an
     * incremental update. The result is a new source beside the old one,
     * which stays open for Undo.
     */
    addText: ({ id, pages }, win) => {
      const entry = get(id);
      const { bytes, pages: written } = addTextLayer(entry.file.buf, pages || []);
      const file = openPdf(bytes);
      file.pages();
      const newId = `p${next++}`;
      const added = { file, path: entry.path, name: entry.name, owner: entry.owner ?? win?.id ?? null, updated: true };
      open.set(newId, added);
      return { ...describe(newId, added), written, bytes: new Uint8Array(bytes) };
    },

    /** The comments on every page of an open PDF, whoever wrote them. */
    annotations: ({ id }) => readAnnotations(get(id).file),

    /**
     * Comments added, answered, resolved, changed or deleted, written as an
     * incremental update signed by the account at the keyboard (or the
     * author given). A new source, the old one kept for Undo; its comments
     * come back with it.
     */
    annotate: ({ id, changes, author }, win) => {
      const entry = get(id);
      const { bytes, added: made } = annotate(entry.file.buf, changes || {}, { author: author || userName() });
      const file = openPdf(bytes);
      file.pages();
      const newId = `p${next++}`;
      const added = { file, path: entry.path, name: entry.name, owner: entry.owner ?? win?.id ?? null, updated: true };
      open.set(newId, added);
      return { ...describe(newId, added), added: made, annotations: readAnnotations(file), fields: readFields(file), bytes: new Uint8Array(bytes) };
    },

    /**
     * The pictures under the redaction marks of the sources named that only
     * the window can decode (JPEG): each with its bytes and the pixels the
     * marks cover, for the window to hand back decoded.
     */
    redactPlan: ({ sources }) => (sources || []).flatMap((id) => {
      const entry = get(id);
      return redactionPlan(entry.file.buf).map((p) => ({ ...p, source: id, data: new Uint8Array(p.data) }));
    }),

    /**
     * Apply the redaction marks: each source's marked pages rewritten with
     * what the marks covered removed, and the window's pages written whole
     * to `path` as a new file. `images`: { [source]: { [key]: { width, height, data } } }.
     */
    redact: async ({ pages, path: target, images = {}, title }) => {
      const removed = { glyphs: 0, pictures: 0, images: 0, annotations: 0, pages: 0 };
      const redacted = new Map();
      for (const id of new Set((pages || []).map((p) => p.source))) {
        const entry = get(id);
        try {
          const out = applyRedactions(entry.file.buf, { images: images[id] || null });
          const file = openPdf(out.bytes);
          file.pages();
          redacted.set(id, file);
          for (const k of Object.keys(removed)) removed[k] += out.removed[k] || 0;
        } catch (err) {
          if (err?.code !== 'nothing') throw err;
          redacted.set(id, entry.file);
        }
      }
      if (!removed.pages) throw new Error('Nothing is marked for redaction.');
      const items = pages.map((p) => ({ file: redacted.get(p.source), page: Number(p.page), rotate: p.rotate }));
      const bytes = assemble(items, { title });
      await fsp.mkdir(path.dirname(target), { recursive: true });
      await writeWhole(target, bytes);
      return { path: target, size: bytes.length, removed };
    },

    /** An open PDF's bookmarks (each with its page), its pages' labels, and its links. */
    navigation: ({ id }) => {
      const { file } = get(id);
      const labels = pageLabels(file);
      return { outline: outlinePages(file), labels: labels ? labels.map(labelText) : null, links: readLinks(file) };
    },

    /** The fields of an open PDF's form, with their values and where their widgets are. */
    fields: ({ id }) => readFields(get(id).file),

    /**
     * Form values written, as an incremental update with the appearances
     * other readers show: { [fieldId or name]: value }. A new source, the old
     * one kept for Undo.
     */
    fill: ({ id, values }, win) => {
      const entry = get(id);
      const { bytes, changed } = fillFields(entry.file.buf, values || {});
      const file = openPdf(bytes);
      file.pages();
      const newId = `p${next++}`;
      const added = { file, path: entry.path, name: entry.name, owner: entry.owner ?? win?.id ?? null, updated: true };
      open.set(newId, added);
      return { ...describe(newId, added), changed, annotations: readAnnotations(file), fields: readFields(file), bytes: new Uint8Array(bytes) };
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
