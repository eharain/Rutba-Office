/**
 * A table of authorities, as Word builds one from TA fields.
 *
 * Mark Citation writes a TA field where a citation is marked — ` TA \l "Long
 * citation" \s "Short" \c 1 ` the first time an authority is marked, and
 * ` TA \s "Short" ` each time after — and Insert Table of Authorities writes
 * a TOA field for a category: its heading, then every authority in it,
 * sorted by its long citation, with the pages it is cited on, or "passim"
 * where that is five pages or more. Nothing here touches a package: this
 * reads and writes the two field codes and lays the table out, so the
 * window and the engine agree about what a table of authorities says.
 */
import { fieldWords } from './wordindex.js';

/** Word's own categories: 1–7 named, 8–16 numbers until someone renames them. */
export const TOA_CATEGORIES = ['Cases', 'Statutes', 'Other Authorities', 'Rules', 'Treatises', 'Regulations', 'Constitutional Provisions', '8', '9', '10', '11', '12', '13', '14', '15', '16'];

const clamp = (n) => Math.max(1, Math.min(16, Math.round(Number(n)) || 1));

/** A category's name, as its heading reads. */
export const categoryName = (n) => TOA_CATEGORIES[clamp(n) - 1];

/** A citation's words as a field code holds them: one line, its quotes made single. */
const quoted = (t) => String(t || '').replace(/[\r\n\t]+/g, ' ').replace(/\s{2,}/g, ' ').replace(/"/g, "'").trim();

/**
 * A TA field's code, as Mark Citation writes it: the long citation, the
 * short one and the category the first time; the short one alone after.
 * `bookmark` makes the page a range, `bold` and `italic` its number.
 */
export function taInstr({ long = '', short = '', category = 1, bookmark = '', bold = false, italic = false } = {}) {
  const l = quoted(long);
  const s = quoted(short) || l;
  let instr = ' TA ';
  if (l) instr += '\\l "' + l + '" ';
  if (s) instr += '\\s "' + s + '" ';
  if (l) instr += '\\c ' + clamp(category) + ' ';
  if (bookmark) instr += '\\r "' + bookmark + '" ';
  if (bold) instr += '\\b ';
  if (italic) instr += '\\i ';
  return instr;
}

/** A TA field's code, read back: `{ long, short, category, bookmark, bold, italic }`, or null. */
export function parseTaInstr(instr) {
  const w = fieldWords(instr);
  if (!w.length || w[0].text.toUpperCase() !== 'TA') return null;
  const out = { long: '', short: '', category: 1, bookmark: '', bold: false, italic: false };
  for (let i = 1; i < w.length; i++) {
    const s = w[i].quoted ? '' : w[i].text.toLowerCase();
    if (s === '\\l') out.long = w[++i]?.text || '';
    else if (s === '\\s') out.short = w[++i]?.text || '';
    else if (s === '\\c') out.category = clamp(w[++i]?.text);
    else if (s === '\\r') out.bookmark = w[++i]?.text || '';
    else if (s === '\\b') out.bold = true;
    else if (s === '\\i') out.italic = true;
  }
  return out.long || out.short ? out : null;
}

/**
 * A TOA field's code from Insert Table of Authorities' choices: `\c` the
 * category, `\h` its heading over it, `\p` passim, `\f` the entries without
 * the formatting they have where they were marked.
 */
export function toaInstr({ category = 1, heading = true, passim = true, keepFormatting = true } = {}) {
  let instr = ' TOA ';
  if (heading) instr += '\\h ';
  instr += '\\c "' + clamp(category) + '" ';
  if (passim) instr += '\\p ';
  if (!keepFormatting) instr += '\\f ';
  return instr;
}

/** A TOA field's code, read back. */
export function parseToaInstr(instr) {
  const w = fieldWords(instr);
  if (!w.length || w[0].text.toUpperCase() !== 'TOA') return null;
  const out = { category: 1, heading: false, passim: false, keepFormatting: true };
  for (let i = 1; i < w.length; i++) {
    const s = w[i].quoted ? '' : w[i].text.toLowerCase();
    if (s === '\\c') out.category = clamp(w[++i]?.text);
    else if (s === '\\h') out.heading = true;
    else if (s === '\\p') out.passim = true;
    else if (s === '\\f') out.keepFormatting = false;
    else if (s === '\\e' || s === '\\l' || s === '\\g' || s === '\\s' || s === '\\d' || s === '\\b') i++;
  }
  return out;
}

const keyOf = (t) => String(t || '').trim().toLocaleLowerCase();
const cmp = (a, b) => {
  const c = String(a).localeCompare(String(b), undefined, { sensitivity: 'base', numeric: true });
  return c || String(a).localeCompare(String(b));
};

/**
 * The authorities, laid out. `entries` are the marked TA fields, each with
 * the page it landed on (`page`) and, for a range, the page its bookmark
 * ends on (`pageEnd`). A field with a long citation names an authority; a
 * field with only a short one counts for the authority of that short
 * citation. Answers every authority `{ long, short, category, pages:
 * [{ text, bold, italic }], passim }`, sorted by its long citation — in one
 * category when `category` is given. `passim` is set where `usePassim` is
 * and the authority is cited on five pages or more.
 */
export function buildAuthorities(entries, { category = null, usePassim = true } = {}) {
  const list = [];
  const byShort = new Map();
  const byLong = new Map();
  for (const e of entries || []) {
    if (!e.long) continue;
    let a = byLong.get(keyOf(e.long));
    if (!a) {
      a = { long: e.long, short: e.short || e.long, category: clamp(e.category), refs: [] };
      list.push(a);
      byLong.set(keyOf(e.long), a);
    }
    if (!byShort.has(keyOf(a.short))) byShort.set(keyOf(a.short), a);
  }
  for (const e of entries || []) {
    const a = e.long ? byLong.get(keyOf(e.long)) : byShort.get(keyOf(e.short)) || byLong.get(keyOf(e.short));
    if (!a || e.page === null || e.page === undefined || e.page === '') continue;
    a.refs.push({ page: Number(e.page), pageEnd: e.pageEnd != null && e.pageEnd !== '' ? Number(e.pageEnd) : null, bold: Boolean(e.bold), italic: Boolean(e.italic) });
  }
  const pagesOf = (a) => {
    const seen = new Map();
    for (const r of a.refs) {
      const text = r.pageEnd && r.pageEnd !== r.page ? `${Math.min(r.page, r.pageEnd)}–${Math.max(r.page, r.pageEnd)}` : String(r.page);
      const was = seen.get(text);
      if (was) { was.bold ||= r.bold; was.italic ||= r.italic; } else seen.set(text, { text, bold: r.bold, italic: r.italic, at: Math.min(r.page, r.pageEnd ?? r.page) });
    }
    return [...seen.values()].sort((x, y) => x.at - y.at || cmp(x.text, y.text)).map(({ text, bold, italic }) => ({ text, bold, italic }));
  };
  return list
    .filter((a) => category == null || a.category === clamp(category))
    .sort((x, y) => cmp(x.long, y.long))
    .map((a) => {
      const pages = pagesOf(a);
      // Word counts the pages an authority is cited on, not the times.
      const distinct = new Set(a.refs.map((r) => r.page)).size;
      return { long: a.long, short: a.short, category: a.category, pages, passim: Boolean(usePassim) && distinct >= 5 };
    });
}

/** One authority's words after its citation: a tab, then its pages ("2, 5") or "passim", as segments `{ text, bold, italic }`. */
export function authorityTail(a) {
  const segs = [{ text: '\t' }];
  if (a.passim) segs.push({ text: 'passim' });
  else a.pages.forEach((p, i) => {
    if (i) segs.push({ text: ', ' });
    segs.push({ text: p.text, ...(p.bold ? { bold: true } : {}), ...(p.italic ? { italic: true } : {}) });
  });
  return segs;
}

/**
 * Words Next Citation stops at, as Word's does: "v." between parties, "In
 * re", Id., supra and infra, a section sign, and the reporters and codes a
 * citation names.
 */
export const CITATION_HINTS = /\bv\.\s|\bin re\b|\bId\.|\bsupra\b|\binfra\b|\bibid\b|§|\bU\.S\.C\.|\bU\.S\.\s|\bF\.\s?(?:2d|3d|4th|Supp)\b|\bS\.\s?Ct\.|\bCir\.|\bCong\.|\bSess\.|\bStat\.|\bC\.F\.R\./i;
