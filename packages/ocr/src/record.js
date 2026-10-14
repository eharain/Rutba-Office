// What the windows make of a document read with recognize(image, { document:
// true }): one row of an expense sheet from a receipt or an invoice, the note
// Mail shows over an attached invoice, names for the tick boxes of a form
// made into check box fields, and the limits of what can be read in the
// languages chosen. Pure, so the windows and their tests share it.
//
//   import { documentRecord, recordCells, invoiceNote } from '@rutba/ocr/record';
//   const row = documentRecord(read.document);
//   row.values.total   // 190.13
//   row.check.total    // false: the page's sums agree, so it goes in as read
//
// A figure (subtotal, tax, total) is as read only on a page whose sums agree
// (docs/OCR.md: on every such page all three were right); on any other page
// it is marked to check. A date, a number or a supplier is marked when the
// reading was unsure of a word of it.

import { parseAmount, parseDate } from './fields.js';

/** The row's columns, in order. */
export const RECORD_FIELDS = Object.freeze(['date', 'supplier', 'number', 'subtotal', 'tax', 'total', 'currency']);

/** The columns' headings, in English: the window translates them. */
export const RECORD_HEADINGS = Object.freeze({
  date: 'Date', supplier: 'Supplier', number: 'Number', subtotal: 'Subtotal', tax: 'Tax', total: 'Total', currency: 'Currency',
});

const FIGURES = new Set(['subtotal', 'tax', 'total']);

/** A word the reading was surer of than this is taken as read. */
export const SURE = 75;

/** Currencies written with no decimals. */
const WHOLE = new Set(['JPY', 'KRW']);

/** How many decimals a currency's amounts are written with. */
export const decimalsOf = (currency) => (WHOLE.has(currency) ? 0 : 2);

/**
 * One row from a document read.
 * @param {object|null} doc the `document` recognize answered
 * @returns {{ kind: string, values: object, check: object, boxes: object, consistent: boolean, found: boolean }}
 *   values by RECORD_FIELDS (a date as yyyy-mm-dd, amounts as numbers, the rest as text, null when not read);
 *   check: which to look at before they go in; boxes: where each was read, in the picture's pixels
 */
export function documentRecord(doc) {
  const fields = doc?.fields || {};
  const consistent = Boolean(doc?.checks?.consistent);
  const values = {};
  const check = {};
  const boxes = {};
  for (const name of RECORD_FIELDS) {
    if (name === 'currency') {
      values.currency = doc?.currency || null;
      check.currency = false;
      boxes.currency = null;
      continue;
    }
    const field = fields[name];
    const value = field?.value ?? null;
    values[name] = value;
    boxes[name] = field?.box || null;
    if (FIGURES.has(name)) check[name] = value === null ? name === 'total' : !consistent;
    else if (value === null) check[name] = name === 'date';
    else check[name] = typeof field.confidence === 'number' && field.confidence < SURE;
  }
  const found = RECORD_FIELDS.some((name) => name !== 'currency' && values[name] !== null);
  return { kind: doc?.kind || 'other', values, check, boxes, consistent, found };
}

/** A value as the preview shows it, to be corrected as text: a date yyyy-mm-dd, an amount with its decimals. */
export function shownValue(name, value, currency = null) {
  if (value === null || value === undefined) return '';
  if (FIGURES.has(name) && typeof value === 'number') return value.toFixed(decimalsOf(currency));
  return String(value);
}

/** A date (yyyy-mm-dd) as a sheet's serial number: days since 30 December 1899. */
export function dateSerial(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
  if (!m) return null;
  return Math.round(Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000) + 25569;
}

/**
 * A row as it goes into a sheet, from the preview's text (as read, or as the
 * person corrected it): the date a date, the amounts numbers, the number and
 * the supplier text (an invoice number keeps its leading zeros). Text that is
 * not a date or an amount where one is wanted goes in as the text it is.
 * @param {{ [field]: string }} shown the preview's text by RECORD_FIELDS
 * @returns {{ value: string|number|null, text?: boolean, format?: string }[]} one a column
 */
export function recordCells(shown) {
  const currency = String(shown.currency || '').trim().toUpperCase() || null;
  return RECORD_FIELDS.map((name) => {
    const text = String(shown[name] ?? '').trim();
    if (!text) return { value: null };
    if (name === 'date') {
      const serial = dateSerial(parseDate(text));
      return serial === null ? { value: text, text: true } : { value: serial, format: 'yyyy-mm-dd' };
    }
    if (FIGURES.has(name)) {
      const amount = /\d/.test(text) ? parseAmount(text) : null;
      return amount ? { value: amount.value, format: decimalsOf(currency) ? '#,##0.00' : '#,##0' } : { value: text, text: true };
    }
    return { value: text, text: true };
  });
}

/**
 * What Mail says of an attachment read as a document, or null when it is not
 * an invoice or a receipt: who it is from, the total, the due date and the
 * number, each with the box it was read from.
 */
export function invoiceNote(doc) {
  if (!doc || (doc.kind !== 'invoice' && doc.kind !== 'receipt')) return null;
  const f = doc.fields || {};
  const part = (name) => (f[name] && f[name].value !== null && f[name].value !== undefined
    ? { value: f[name].value, box: f[name].box || null, sure: !(typeof f[name].confidence === 'number' && f[name].confidence < SURE) }
    : null);
  const note = { kind: doc.kind, from: part('supplier'), total: part('total'), dueDate: part('dueDate'), number: part('number'), currency: doc.currency || null, consistent: Boolean(doc.checks?.consistent) };
  // A figure is sure only when the sums agree.
  if (note.total) note.total.sure = note.consistent;
  // Who it is from alone is a guess (the largest writing near the top): a letter that names an invoice is not one.
  if (!note.total && !note.dueDate && !note.number) return null;
  return note;
}

/**
 * Names for the tick boxes of a form: the words beside each box on its row,
 * after it in the page's direction (to its right, or to its left right to
 * left), up to a gap or the next box; "Check Box n" where nothing is beside
 * it. Names are made unique as a form's fields must be.
 * @param {{ box: { x, y, width, height } }[]} checkboxes as findCheckboxes answers them
 * @param {{ text: string, box: object }[]} words the words read from the page
 * @param {{ rtl?: boolean }} [options]
 * @returns {string[]} one a box
 */
export function checkboxLabels(checkboxes, words = [], { rtl = false } = {}) {
  const used = new Map();
  return checkboxes.map((cb, i) => {
    const b = cb.box;
    const mid = b.y + b.height / 2;
    const others = checkboxes.filter((o) => o !== cb).map((o) => o.box);
    const onRow = words
      .filter((w) => w.text && Math.abs(w.box.y + w.box.height / 2 - mid) < Math.max(b.height, w.box.height) * 0.6)
      .filter((w) => (rtl ? w.box.x + w.box.width <= b.x + 2 : w.box.x >= b.x + b.width - 2))
      .filter((w) => !(w.box.width <= b.width * 1.6 && Math.abs(w.box.x - b.x) < b.width && Math.abs(w.box.y - b.y) < b.height))
      .sort((p, q) => (rtl ? q.box.x - p.box.x : p.box.x - q.box.x));
    const taken = [];
    let edge = rtl ? b.x : b.x + b.width;
    for (const w of onRow) {
      const gap = rtl ? edge - (w.box.x + w.box.width) : w.box.x - edge;
      if (gap > Math.max(b.width, w.box.height) * (taken.length ? 1.6 : 3)) break;
      // Another box between, or under the word (a box read as marks): its words are its own.
      if (others.some((o) => Math.abs(o.y + o.height / 2 - mid) < b.height && (rtl ? o.x < edge && o.x + o.width > w.box.x : o.x > edge - 1 && o.x < w.box.x + w.box.width))) break;
      // Marks with no letter or digit (a box's sides read as brackets) are not words of a name.
      if (!/[\p{L}\p{N}]/u.test(w.text)) continue;
      taken.push(w.text);
      edge = rtl ? w.box.x : w.box.x + w.box.width;
    }
    // A box's side read with the word beside it (Phone’) is no part of the name: marks at either end go.
    let name = taken.join(' ').replace(/^[^\p{L}\p{N}(]+|[^\p{L}\p{N})]+$/gu, '').trim() || `Check Box ${i + 1}`;
    const n = (used.get(name) || 0) + 1;
    used.set(name, n);
    if (n > 1) name = `${name} ${n}`;
    return name;
  });
}

/**
 * What cannot be read yet in the languages chosen, to be said where it applies:
 * 'digits' when Arabic or Urdu is read (the models do not read Arabic-Indic or
 * Urdu digits, so a page printed in them gives no figures or dates), and
 * 'nastaliq' when Urdu is (Naskh type reads well; Nastaliq print, the style
 * most Urdu is printed in, is beta).
 */
export function ocrLimits(languages = []) {
  const out = [];
  if (languages.includes('ara') || languages.includes('urd')) out.push('digits');
  if (languages.includes('urd')) out.push('nastaliq');
  return out;
}
