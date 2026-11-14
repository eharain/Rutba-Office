// Redaction in the window: words found marked for redaction on every page,
// the pictures under the marks that only the browser can decode (JPEG)
// decoded for the main process to paint out, and the dialog that applies
// the marks, which says what applying does and writes a new file.
//
// Marks are Redact annotations, written like any comment (annots.js), so
// they are reviewed and deleted in the Comments pane and taken away by Undo.
// Applying is not undone: it writes a new file, whole, and leaves the
// original as it was.

import React from 'react';
import { Button, Dialog, Icon, language, t, tn } from '@rutba/office-ui';
import { findAll } from '@rutba/editing/find';
import { drawText, getPage } from '../../pdfjs.js';
import { joinLines } from './annots.js';
import { exactQuads } from './quads.js';

/** How far past a measured match a mark reaches, in the line's height. */
const WIDEN = 0.5;

/**
 * Where `query` is found on a page, as QuadPoints in the page's own
 * coordinates: the page's words laid out off screen as pdf.js lays them over
 * a page, the matches measured there, line by line. That layer is drawn in a
 * stand-in font, so within a line it is only near: when the page's own glyphs
 * (`exact`) find the same number of matches, their boxes are the ones used.
 * Where they cannot, each mark is widened, since a mark too short leaves letters behind.
 */
export async function searchQuads(doc, index, query, exact = null) {
  const page = await getPage(doc, index);
  const vp = page.getViewport({ scale: 1, rotation: 0 });
  const host = document.createElement('div');
  host.className = 'textLayer pdf-measure';
  host.style.cssText = `position:fixed;left:-20000px;top:0;width:${vp.width}px;height:${vp.height}px;visibility:hidden;`;
  document.body.append(host);
  try {
    await drawText(doc, index, host, { scale: 1, rotate: 0 });
    const nodes = [];
    let text = '';
    const walker = document.createTreeWalker(host, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (n.nodeType === 3) { nodes.push({ node: n, at: text.length }); text += n.data; } else if (n.nodeName === 'BR') text += '\n';
    }
    const locate = (offset) => {
      for (let k = nodes.length - 1; k >= 0; k--) if (nodes[k].at <= offset) return { node: nodes[k].node, offset: Math.min(offset - nodes[k].at, nodes[k].node.data.length) };
      return null;
    };
    const box = host.getBoundingClientRect();
    const out = [];
    const matches = findAll(text, query);
    const measured = matches.length ? await exact?.().catch(() => null) : null;
    const sure = exactQuads(matches.length, measured);
    if (sure) return sure;
    for (const m of matches) {
      const a = locate(m.start);
      const b = locate(m.end);
      if (!a || !b) continue;
      const range = document.createRange();
      try { range.setStart(a.node, a.offset); range.setEnd(b.node, b.offset); } catch { continue; }
      const rects = [...range.getClientRects()].filter((r) => r.width > 0.5 && r.height > 0.5)
        .map((r) => ({ left: r.left - box.left, right: r.right - box.left, top: r.top - box.top, bottom: r.bottom - box.top }));
      for (const line of joinLines(rects)) {
        const reach = (line.bottom - line.top) * WIDEN;
        const tl = vp.convertToPdfPoint(line.left - reach, line.top);
        const tr = vp.convertToPdfPoint(line.right + reach, line.top);
        const bl = vp.convertToPdfPoint(line.left - reach, line.bottom);
        const br = vp.convertToPdfPoint(line.right + reach, line.bottom);
        out.push([...tl, ...tr, ...bl, ...br].map((v) => Math.round(v * 100) / 100));
      }
    }
    return out;
  } finally {
    host.remove();
  }
}

/**
 * The pictures of a redaction plan decoded here: { [source]: { [key]: { width, height, data } } }
 * as RGB samples. A picture the browser cannot decode is left out; the main
 * process then leaves it off its page, and the count says so.
 */
export async function decodePlan(plan) {
  const out = {};
  for (const item of plan) {
    if (!/DCT/.test(item.filter || '')) continue;
    try {
      const bitmap = await createImageBitmap(new Blob([item.data], { type: 'image/jpeg' }));
      const canvas = document.createElement('canvas');
      canvas.width = bitmap.width;
      canvas.height = bitmap.height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(bitmap, 0, 0);
      const { data } = ctx.getImageData(0, 0, bitmap.width, bitmap.height);
      const rgb = new Uint8Array(bitmap.width * bitmap.height * 3);
      for (let i = 0, o = 0; i < data.length; i += 4, o += 3) { rgb[o] = data[i]; rgb[o + 1] = data[i + 1]; rgb[o + 2] = data[i + 2]; }
      (out[item.source] ||= {})[item.key] = { width: bitmap.width, height: bitmap.height, data: rgb };
      bitmap.close?.();
    } catch {
      /* left out of its page instead */
    }
  }
  return out;
}

/** What applying takes out besides the pages' words, by kind, in a sentence's list: "2 bookmarks, 1 comment and the title". */
const KINDS = [
  ['title', () => t('pdf.redact.the-title')],
  ['subject', () => t('pdf.redact.the-subject')],
  ['keywords', () => t('pdf.redact.the-keywords')],
  ['author', () => t('pdf.redact.the-author')],
  ['creator', () => t('pdf.redact.name-program-made')],
  ['document metadata', () => t('pdf.redact.the-documents')],
  ['metadata', (n) => tn(n, 'pdf.redact.count-piece-metadata')],
  ['bookmark', (n) => tn(n, 'pdf.redact.count-bookmark')],
  ['comment', (n) => tn(n, 'common.count-comment')],
  ['link', (n) => tn(n, 'pdf.redact.count-link')],
  ['attached file', (n) => tn(n, 'pdf.redact.count-attached-file')],
  ['form field', (n) => tn(n, 'pdf.redact.count-form-field')],
  ['form value', (n) => tn(n, 'pdf.redact.count-form-value')],
  ['tooltip', (n) => tn(n, 'pdf.redact.count-field-tooltip')],
  ['field name', (n) => tn(n, 'pdf.redact.count-field-name')],
  ['page label', (n) => tn(n, 'pdf.redact.count-page-label')],
];
export function elsewhereList(counts) {
  const parts = KINDS.filter(([k]) => counts?.[k]).map(([k, say]) => say(counts[k]));
  if (!parts.length) return '';
  try { return new Intl.ListFormat(language(), { type: 'conjunction' }).format(parts); } catch { return parts.join(', '); }
}

/** Apply Redactions: what applying does, and what it found besides the marks, said before it is done. */
export function ApplyRedactionsDialog({ marks, pages, preview, onApply, onClose }) {
  const also = elsewhereList(preview?.elsewhere);
  const unpainted = preview?.unpainted || [];
  return (
    <Dialog
      title={t('common.apply-redactions')}
      onClose={onClose}
      width={500}
      actions={(
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('pdf.redact.apply-save')} data-role="redact-go" disabled={!marks} onClick={onApply} />
        </>
      )}
    >
      <div className="pdf-redact-form" data-role="redact-dialog">
        <div className="pdf-redact-lead">
          <span className="pdf-redact-icon"><Icon name="redact" size={22} /></span>
          <span>{pages > 1 ? t('pdf.redact.count-marks-pages', { count: marks, pages }) : tn(marks, 'pdf.redact.count-mark-one')}</span>
        </div>
        <p>{t('pdf.redact.applying-removes-words')}</p>
        {preview?.words?.length ? (
          <p className="pdf-redact-words" data-role="redact-words">{t('pdf.redact.words-marked-words', { words: preview.words.map((w) => `"${w}"`).join(', ') })}</p>
        ) : null}
        {also ? (
          <p className="pdf-redact-also" data-role="redact-also">{t('pdf.redact.found-elsewhere-file', { list: also })}</p>
        ) : (
          <p className="pdf-redact-also" data-role="redact-also">{t('pdf.redact.marked-words-found')}</p>
        )}
        {unpainted.length ? (
          <div className="pdf-redact-pictures" data-role="redact-pictures">
            <Icon name="info" size={16} />
            <span>
              {t('pdf.redact.picture-under-mark')}
              {' '}
              {unpainted.map((u) => tn(u.count, 'pdf.redact.page-page-count', { page: u.page })).join('; ')}
            </span>
          </div>
        ) : null}
        <p className="pdf-redact-warn" data-role="redact-warning">
          {t('pdf.redact.redacted-pdf-saved')}
        </p>
        <p className="pdf-redact-small">{t('pdf.redact.lines-shapes-drawn')}</p>
      </div>
    </Dialog>
  );
}
