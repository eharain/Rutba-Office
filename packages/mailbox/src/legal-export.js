// Export for legal hold and handover.
//
// A folder, several folders, or the results of a search, from any account or
// archive, written to a new folder of their own:
//
//   messages.mbox       every message, as RFC 5322 in an mbox anything opens
//   messages.pdf        every message on paper, after a summary page, or
//   pdf/0001 ….pdf      one PDF a message
//   attachments/0001/…  each message's files beside it, when asked for
//   manifest.csv        a line a message: date, from, to, subject,
//   manifest.json       Message-ID, where it came from, and the SHA-256 of
//                       what was written for it
//   summary.pdf         what was exported, from where, and every file's hash
//
// The hash of a message is of its bytes as written into the mbox: read the
// mbox back, take a message less the one line feed that separates it from the
// next (and undo the ">From " quoting mboxrd does), and its SHA-256 is the one
// in the manifest. A message from an mbox is exported exactly as that archive
// holds it; any other is written by rfc822.js, the same bytes every time.
//
// The sources are only ever read. The export runs a message at a time and
// gives the event loop back between them, reports progress, and stops when
// asked; a stopped export keeps what it wrote and says in its manifest and
// summary that it was stopped.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { toRfc822 } from './rfc822.js';
import { attachmentFileName } from './mime.js';
import { messagePdf, combinedPdf, summaryPdf, dateText, PDF_LABELS } from './message-pdf.js';

const sha256 = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex');

/** A file name made from a subject: safe on every system, and short. */
function nameFrom(n, subject) {
  const base = String(subject || '')
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/, '')
    .slice(0, 60);
  return `${String(n).padStart(4, '0')}${base ? ` ${base}` : ''}`;
}

/** A folder for the export that does not exist yet: `name`, else `name (2)`, … */
export function freshFolder(parent, name) {
  let dir = path.join(parent, name);
  for (let i = 2; fs.existsSync(dir); i++) dir = path.join(parent, `${name} (${i})`);
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

/** mboxrd quoting, on the bytes as one character a byte. */
function quoteFromLines(bytes) {
  const text = Buffer.from(bytes).toString('latin1').replace(/^(>*From )/gm, '>$1');
  return Buffer.from(text, 'latin1');
}

/** A CSV field. A leading = + - @ is quoted with ' so a spreadsheet never runs it. */
function csv(value) {
  let s = String(value ?? '');
  if (/^[=+\-@\t\r]/.test(s)) s = `'${s}`;
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

const person = (p) => (p ? (p.name && p.address && p.name !== p.address ? `${p.name} <${p.address}>` : p.address || p.name || '') : '');
const people = (list) => (list || []).map(person).filter(Boolean).join('; ');

/** Loads a message from the local store, its attachments' bytes read from disk. */
export function storeLoader(store) {
  return ({ sourceId, folder, id }) => {
    const m = store.get(sourceId, folder, id);
    if (!m) return null;
    const attachments = (m.attachments || []).map((a) => {
      if (!a.stored) return { ...a, bytes: null };
      try {
        return { ...a, bytes: fs.readFileSync(store.attachmentPath(sourceId, folder, id, a.stored)) };
      } catch {
        return { ...a, bytes: null };
      }
    });
    return { message: { ...m, attachments }, raw: null };
  };
}

/** Loads a message from an archive opened in place, as the archive holds it where it can. */
export function archiveLoader(source) {
  return ({ folder, id }) => {
    const message = source.message(folder, id);
    return message ? { message, raw: source.raw(folder, id) } : null;
  };
}

/**
 * Export messages.
 * @param {{
 *   items: Array<{ sourceId: string, sourceName: string, sourceKind?: string, sourceFile?: string, folder: string, id: string, date?: string }>,
 *   load: (item) => ({ message: object, raw: Uint8Array|null } | null | Promise<any>),
 *   dir: string,
 *   mbox?: boolean, pdf?: 'combined'|'each'|'both'|null, attachments?: boolean,
 *   selection?: string, labels?: object, unicodeFont?: Uint8Array|null,
 *   onProgress?: (p: { done: number, total: number }) => void, cancelled?: () => boolean, now?: () => number,
 * }} opts
 */
export async function exportMessages(opts) {
  const { items, load, dir, mbox = true, pdf = 'combined', attachments = false, selection = '', labels = {}, unicodeFont = null, onProgress = null, cancelled = () => false, now = () => Date.now() } = opts;
  const L = { ...PDF_LABELS, ...labels };
  const exportedAt = new Date(now()).toISOString();
  fs.mkdirSync(dir, { recursive: true });
  const total = items.length;

  // What is in it, by source and folder, for the summary page.
  const bySource = new Map();
  let first = null;
  let last = null;
  for (const item of items) {
    const s = bySource.get(item.sourceId) || { name: item.sourceName || item.sourceId, folders: new Set(), count: 0 };
    s.folders.add(item.folder);
    s.count++;
    bySource.set(item.sourceId, s);
    if (item.date) {
      if (!first || item.date < first) first = item.date;
      if (!last || item.date > last) last = item.date;
    }
  }
  const sourcesTable = () => [...bySource.values()].map((s) => ({ name: s.name, folders: [...s.folders].join(', '), count: String(s.count) }));
  const datesText = () => (first ? `${dateText(first).slice(0, 10)} to ${dateText(last).slice(0, 10)}` : '');

  const wantEach = pdf === 'each' || pdf === 'both';
  const wantCombined = pdf === 'combined' || pdf === 'both';
  const mboxPath = path.join(dir, 'messages.mbox');
  const mboxFd = mbox ? fs.openSync(mboxPath, 'w') : null;
  const mboxHash = mbox ? crypto.createHash('sha256') : null;
  let mboxBytes = 0;
  const combined = wantCombined ? combinedPdf({ labels: L, unicodeFont, created: exportedAt, title: selection || L.summaryTitle }) : null;
  combined?.summary({ exportedAt, selection, messages: total, dates: datesText(), sources: sourcesTable() });
  if (wantEach) fs.mkdirSync(path.join(dir, 'pdf'), { recursive: true });

  const rows = [];
  let stopped = false;
  let attachmentFiles = 0;
  let skipped = 0;
  try {
    for (let n = 0; n < items.length; n++) {
      if (cancelled()) {
        stopped = true;
        break;
      }
      const item = items[n];
      const loaded = await load(item);
      if (!loaded?.message) {
        skipped++;
        continue;
      }
      const m = loaded.message;
      const number = rows.length + 1;
      const files = (m.attachments || []).map((a) => ({
        ...a,
        size: a.size ?? a.bytes?.length ?? 0,
        sha256: a.bytes ? sha256(a.bytes) : null,
      }));
      const bytes = loaded.raw ? Buffer.from(loaded.raw) : toRfc822({ ...m, attachments: files });
      const record = {
        n: number,
        date: m.date || item.date || null,
        from: person(Array.isArray(m.from) ? m.from[0] : m.from),
        to: people(m.to),
        cc: people(m.cc),
        subject: m.subject || '',
        messageId: m.messageId ? `<${String(m.messageId).replace(/^<|>$/g, '')}>` : null,
        source: item.sourceName || item.sourceId,
        sourceKind: item.sourceKind || null,
        sourceFile: item.sourceFile || null,
        folder: item.folder,
        sha256: sha256(bytes),
        bytes: bytes.length,
        pdf: null,
        attachments: [],
      };

      if (mboxFd !== null) {
        const from = (Array.isArray(m.from) ? m.from[0] : m.from)?.address || 'unknown@localhost';
        const when = m.date && !Number.isNaN(new Date(m.date).getTime()) ? new Date(m.date) : new Date(0);
        const chunk = Buffer.concat([Buffer.from(`From ${from} ${when.toUTCString()}\n`, 'latin1'), quoteFromLines(bytes), Buffer.from('\n')]);
        fs.writeSync(mboxFd, chunk);
        mboxHash.update(chunk);
        record.mboxOffset = mboxBytes;
        mboxBytes += chunk.length;
      }

      const where = { source: item.sourceFile ? `${item.sourceName} (${path.basename(item.sourceFile)})` : item.sourceName, folder: item.folder };
      if (wantEach) {
        const out = messagePdf({ ...m, attachments: files }, { ...where, footer: `${L.footer} · ${number} / ${total}` }, { labels: L, unicodeFont, created: exportedAt });
        const name = `${nameFrom(number, m.subject)}.pdf`;
        fs.writeFileSync(path.join(dir, 'pdf', name), out.bytes);
        record.pdf = { file: `pdf/${name}`, pages: out.pages, sha256: sha256(out.bytes) };
      }
      if (combined) {
        const pages = combined.add({ ...m, attachments: files }, where);
        record.pages = pages.from === pages.to ? String(pages.from) : `${pages.from}-${pages.to}`;
      }
      if (files.length) {
        const taken = new Set();
        const folder = path.join(dir, 'attachments', String(number).padStart(4, '0'));
        for (const a of files) {
          const entry = { name: a.filename || '', size: a.size, sha256: a.sha256, file: null };
          if (attachments && a.bytes) {
            let name = attachmentFileName(a.filename);
            const ext = path.extname(name);
            for (let i = 2; taken.has(name.toLowerCase()); i++) name = `${path.basename(attachmentFileName(a.filename), ext)} (${i})${ext}`;
            taken.add(name.toLowerCase());
            fs.mkdirSync(folder, { recursive: true });
            fs.writeFileSync(path.join(folder, name), a.bytes);
            entry.file = `attachments/${String(number).padStart(4, '0')}/${name}`;
            attachmentFiles++;
          }
          record.attachments.push(entry);
        }
      }
      rows.push(record);
      onProgress?.({ done: n + 1, total });
      await new Promise((resolve) => setImmediate(resolve));
    }
  } finally {
    if (mboxFd !== null) fs.closeSync(mboxFd);
  }

  // The files, each with its hash.
  const files = [];
  const listFile = (name, bytes, hash = null) => files.push({ name, size: bytes, sha256: hash });
  if (mbox) listFile('messages.mbox', mboxBytes, mboxHash.digest('hex'));
  let combinedPages = 0;
  if (combined) {
    const out = combined.finish();
    combinedPages = out.pages;
    fs.writeFileSync(path.join(dir, 'messages.pdf'), out.bytes);
    listFile('messages.pdf', out.bytes.length, sha256(out.bytes));
  }

  const header = ['n', 'date', 'from', 'to', 'cc', 'subject', 'message_id', 'source', 'source_kind', 'source_file', 'folder', 'message_sha256', 'pdf_file', 'pdf_pages', 'pdf_sha256', 'attachments', 'attachment_sha256'];
  const lines = rows.map((r) =>
    [
      r.n, r.date, r.from, r.to, r.cc, r.subject, r.messageId, r.source, r.sourceKind, r.sourceFile, r.folder, r.sha256,
      r.pdf?.file || (r.pages ? 'messages.pdf' : ''), r.pdf ? String(r.pdf.pages) : r.pages || '', r.pdf?.sha256 || '',
      r.attachments.map((a) => a.file || a.name).join('; '), r.attachments.map((a) => a.sha256 || '').join('; '),
    ].map(csv).join(',')
  );
  // A byte-order mark, so a spreadsheet reads the names in UTF-8.
  const csvBytes = Buffer.from(`\ufeff${[header.join(','), ...lines].join('\r\n')}\r\n`, 'utf8');
  fs.writeFileSync(path.join(dir, 'manifest.csv'), csvBytes);
  listFile('manifest.csv', csvBytes.length, sha256(csvBytes));

  const summaryInfo = { exportedAt, selection, messages: rows.length, dates: datesText(), sources: sourcesTable(), files: files.slice(), cancelled: stopped };
  const summary = summaryPdf(summaryInfo, { labels: L, unicodeFont, created: exportedAt });
  fs.writeFileSync(path.join(dir, 'summary.pdf'), summary.bytes);
  listFile('summary.pdf', summary.bytes.length, sha256(summary.bytes));

  const manifest = {
    exportedAt,
    selection,
    complete: !stopped,
    requested: total,
    exported: rows.length,
    skipped,
    hashRule: 'message_sha256 is the SHA-256 of the message as written into messages.mbox: read back, less the line feed that ends it and with mboxrd ">From " quoting undone.',
    sources: [...bySource.entries()].map(([id, s]) => ({ id, name: s.name, folders: [...s.folders], messages: s.count })),
    files,
    messages: rows,
  };
  const json = Buffer.from(JSON.stringify(manifest, null, 2), 'utf8');
  fs.writeFileSync(path.join(dir, 'manifest.json'), json);

  return {
    dir,
    exported: rows.length,
    requested: total,
    skipped,
    cancelled: stopped,
    sources: bySource.size,
    attachments: attachmentFiles,
    pdfs: wantEach ? rows.length : 0,
    combinedPages,
    mbox: mbox ? mboxPath : null,
    files: [...files.map((f) => f.name), 'manifest.json'],
  };
}
