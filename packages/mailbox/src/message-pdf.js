// Messages on paper: a PDF of one message, of many, and of an export's summary.
//
// For legal hold and handover a message has to be readable without a mail
// program, by someone who will print it: its headers set out plainly, its
// body, and the files it carried named with their sizes and SHA-256, so a
// file handed over beside it can be matched to it. Built on @rutba/pdf's
// Sheet, which pages, wraps and writes "Page 2 of 5" at the foot of every
// page. Words the base-14 fonts cannot draw (Urdu, Arabic, Hindi) need the
// TrueType font the caller hands over; without one they print as '?'.

import { createRequire } from 'node:module';
import { bodyText } from './archive.js';

const require = createRequire(import.meta.url);
const { Sheet, INK } = require('@rutba/pdf');

/** The words on the page, in English unless the caller translates them. */
export const PDF_LABELS = {
  from: 'From',
  to: 'To',
  cc: 'Cc',
  bcc: 'Bcc',
  date: 'Date',
  messageId: 'Message-ID',
  source: 'Source',
  folder: 'Folder',
  attachments: 'Attachments',
  name: 'Name',
  type: 'Type',
  size: 'Size',
  sha256: 'SHA-256',
  noSubject: '(no subject)',
  noBody: '(This message has no text.)',
  summaryTitle: 'Export summary',
  exported: 'Exported',
  selection: 'What was exported',
  messages: 'Messages',
  dates: 'Dates',
  sources: 'Sources',
  files: 'Files',
  untouched: 'The sources were opened for reading only; nothing in them was changed.',
  cancelled: 'Cancelled before the end: only the messages listed in the manifest were exported.',
  footer: 'Exported by Rutba Office',
};

const person = (p) => (p ? (p.name && p.address && p.name !== p.address ? `${p.name} <${p.address}>` : p.name || p.address || '') : '');
const people = (list) => (list || []).map(person).filter(Boolean).join(', ');

/** A size as people read it. */
export function bytesText(n) {
  const v = Number(n) || 0;
  if (v < 1024) return `${v} B`;
  if (v < 1024 * 1024) return `${(v / 1024).toFixed(v < 10240 ? 1 : 0)} KB`;
  if (v < 1024 ** 3) return `${(v / 1048576).toFixed(1)} MB`;
  return `${(v / 1024 ** 3).toFixed(2)} GB`;
}

/** A date as a person reads it, in UTC so two readers agree. */
export function dateText(iso) {
  const d = iso ? new Date(iso) : null;
  if (!d || Number.isNaN(d.getTime())) return '';
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())} ${p(d.getUTCHours())}:${p(d.getUTCMinutes())} UTC`;
}

function sheetFor({ title, footer, unicodeFont, created }) {
  const sheet = new Sheet({ title, footer, created, fontSize: 9.5 });
  if (unicodeFont) sheet.doc.useUnicodeFont(unicodeFont);
  return sheet;
}

/**
 * Draw one message onto `sheet` at its cursor.
 * @param {object} m the message, attachments with `sha256` and `size`
 * @param {{ source?: string, folder?: string }} where
 */
function drawMessage(sheet, m, where, L) {
  sheet.heading(m.subject || L.noSubject, { size: 14, gap: 6 });
  const rows = [
    [L.from, person(Array.isArray(m.from) ? m.from[0] : m.from)],
    [L.to, people(m.to)],
    [L.cc, people(m.cc)],
    [L.bcc, people(m.bcc)],
    [L.date, dateText(m.date)],
    [L.messageId, m.messageId ? `<${String(m.messageId).replace(/^<|>$/g, '')}>` : ''],
    [L.source, where.source || ''],
    [L.folder, where.folder || ''],
  ];
  sheet.keyValues(rows, { labelWidth: 78, align: 'left' });
  sheet.rule({ gap: 6 });
  const text = bodyText(m).replace(/\r\n?/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  sheet.paragraph(text || L.noBody, { size: 10, colour: text ? INK.text : INK.quiet });
  const files = m.attachments || [];
  if (files.length) {
    sheet.spacer(6);
    sheet.heading(`${L.attachments} (${files.length})`, { size: 11, gap: 4 });
    sheet.table({
      size: 8.5,
      columns: [
        { key: 'name', label: L.name, flex: 3 },
        { key: 'size', label: L.size, width: 58, align: 'right' },
        { key: 'sha', label: L.sha256, flex: 4, font: 'Courier', size: 7 },
      ],
      rows: files.map((a) => ({ name: a.filename || '', size: bytesText(a.size ?? a.bytes?.length ?? 0), sha: a.sha256 || '' })),
    });
  }
}

/** One message as a PDF. */
export function messagePdf(m, where = {}, { labels = {}, unicodeFont = null, created = null } = {}) {
  const L = { ...PDF_LABELS, ...labels };
  const sheet = sheetFor({ title: m.subject || L.noSubject, footer: where.footer || L.footer, unicodeFont, created });
  drawMessage(sheet, m, where, L);
  const doc = sheet.end();
  return { bytes: doc.toBuffer(), pages: doc.pages.length };
}

/**
 * Many messages in one PDF, each from a new page, after the summary.
 * `add` takes a message and returns the pages it took (first and last, from 1).
 */
export function combinedPdf({ labels = {}, unicodeFont = null, created = null, footer = null, title = '' } = {}) {
  const L = { ...PDF_LABELS, ...labels };
  const sheet = sheetFor({ title: title || L.summaryTitle, footer: footer || L.footer, unicodeFont, created });
  let first = true;
  return {
    summary(info) {
      drawSummary(sheet, info, L);
      first = false;
    },
    add(m, where) {
      if (!first) sheet.newPage();
      first = false;
      const from = sheet.doc.pages.length;
      drawMessage(sheet, m, where, L);
      return { from, to: sheet.doc.pages.length };
    },
    finish() {
      const doc = sheet.end();
      return { bytes: doc.toBuffer(), pages: doc.pages.length };
    },
  };
}

function drawSummary(sheet, info, L) {
  sheet.heading(L.summaryTitle, { size: 16, gap: 10 });
  sheet.keyValues(
    [
      [L.exported, info.exportedAt ? dateText(info.exportedAt) : ''],
      [L.selection, info.selection || ''],
      [L.messages, String(info.messages ?? 0)],
      [L.dates, info.dates || ''],
    ],
    { labelWidth: 120, align: 'left' }
  );
  if (info.sources?.length) {
    sheet.spacer(6);
    sheet.heading(L.sources, { size: 11, gap: 4 });
    sheet.table({
      size: 8.5,
      columns: [
        { key: 'name', label: L.source, flex: 3 },
        { key: 'folders', label: L.folder, flex: 3 },
        { key: 'count', label: L.messages, width: 70, align: 'right' },
      ],
      rows: info.sources,
    });
  }
  if (info.files?.length) {
    sheet.spacer(6);
    sheet.heading(L.files, { size: 11, gap: 4 });
    sheet.table({
      size: 8.5,
      columns: [
        { key: 'name', label: L.name, flex: 2 },
        { key: 'size', label: L.size, width: 58, align: 'right' },
        { key: 'sha256', label: L.sha256, flex: 4, font: 'Courier', size: 7 },
      ],
      rows: info.files.map((f) => ({ ...f, size: bytesText(f.size) })),
    });
  }
  sheet.spacer(6);
  if (info.cancelled) sheet.paragraph(L.cancelled, { colour: '#a33a2a' });
  sheet.paragraph(L.untouched, { colour: INK.quiet });
}

/** The summary alone, as a PDF. */
export function summaryPdf(info, { labels = {}, unicodeFont = null, created = null } = {}) {
  const L = { ...PDF_LABELS, ...labels };
  const sheet = sheetFor({ title: L.summaryTitle, footer: L.footer, unicodeFont, created });
  drawSummary(sheet, info, L);
  const doc = sheet.end();
  return { bytes: doc.toBuffer(), pages: doc.pages.length };
}
