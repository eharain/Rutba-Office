// A paragraph's tracked change of formatting — Word's `w:pPrChange`, last in
// the paragraph's `w:pPr`, holding the properties before in a `w:pPr` of its
// own (which never carries the mark's run properties, a section break or
// another change). The run-level counterpart, `w:rPrChange`, is in runs.js.

import { esc, attrs } from './package.js';
import { unesc } from './workbook.js';

const PPR_CHANGE_RE = /<w:pPrChange\b[^>]*>[\s\S]*?<\/w:pPrChange>/;

/**
 * A paragraph's own properties and the tracked change of them it carries,
 * apart: `own` the `<w:pPr>` without its `<w:pPrChange>` (null when nothing
 * else is left), `change` that element (null when there is none).
 */
export function splitParagraphChange(pPr) {
  if (!pPr || !pPr.includes('<w:pPrChange')) return { own: pPr ?? null, change: null };
  const m = PPR_CHANGE_RE.exec(pPr);
  if (!m) return { own: pPr, change: null };
  const own = pPr.slice(0, m.index) + pPr.slice(m.index + m[0].length);
  return { own: /^<w:pPr\b[^>]*>\s*<\/w:pPr>$/.test(own) ? null : own, change: m[0] };
}

/** `own` with the change put back, last in the `<w:pPr>` as the schema has it. */
export function joinParagraphChange(own, change) {
  if (!change) return own ?? null;
  if (!own || /^<w:pPr\b[^>]*\/>$/.test(own)) return '<w:pPr>' + change + '</w:pPr>';
  return own.replace(/<\/w:pPr>$/, () => change + '</w:pPr>');
}

/** What a change can record of a paragraph's properties: everything but the mark's run properties and a section break. */
function recordable(pPr) {
  const own = splitParagraphChange(pPr).own;
  if (!own || /^<w:pPr\b[^>]*\/>$/.test(own)) return '';
  return own.replace(/^<w:pPr\b[^>]*>/, '').replace(/<\/w:pPr>$/, '')
    .replace(/<w:rPr\b[^>]*>(?:<w:rPrChange\b[\s\S]*?<\/w:rPrChange>|(?!<\/w:rPr>)[\s\S])*?<\/w:rPr>|<w:rPr\b[^>]*\/>/, '')
    .replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/, '');
}

/**
 * Review → Track Changes, recording a change to a paragraph's formatting as
 * Word does: `after` carrying a `<w:pPrChange>` (an id, an author, a date)
 * that holds what `before` had. A paragraph that already carried one keeps
 * it — the change is from how it first was — and one formatted back to how
 * it first was carries none.
 */
export function withParagraphChange(before, after, meta) {
  const was = splitParagraphChange(before);
  const now = splitParagraphChange(after);
  const first = was.change ? (/<w:pPr\b[^>]*>([\s\S]*)<\/w:pPr>/.exec(was.change)?.[1] ?? '') : recordable(was.own);
  if (recordable(now.own) === first) return now.own;
  const change = was.change
    ?? '<w:pPrChange w:id="' + esc(String(meta?.id ?? '0')) + '" w:author="' + esc(String(meta?.author ?? '')) + '"'
      + (meta?.date ? ' w:date="' + esc(String(meta.date)) + '"' : '') + '><w:pPr>' + first + '</w:pPr></w:pPrChange>';
  return joinParagraphChange(now.own, change);
}

/** The tracked change of formatting a paragraph carries — who and when — or null. */
export function paragraphChangeOf(pPr) {
  const { change } = splitParagraphChange(pPr);
  if (!change) return null;
  const a = attrs(/^<w:pPrChange\b([^>]*)>/.exec(change)[1]);
  return { id: a['w:id'] ?? '0', author: a['w:author'] ? unesc(a['w:author']) : '', date: a['w:date'] ?? null };
}
