// Recognise Text: the words in scanned pages found and written into the PDF
// as an invisible text layer, so Find, selecting and copying work on scans.
//
// A page that already carries text is left as it is. Each scanned page is
// drawn by pdf.js at up to 300 dots per inch (its long side at most 4,200
// pixels), its pixels read on the main process's threads, and each word's
// box turned from the picture's pixels into the page's own coordinates by
// the same viewport that drew it, so a rotated or cropped page comes out
// right. The main process then appends the layer to the file as an
// incremental update (main/pdf.js addText): nothing of a page is redrawn.
//
// With Find fields chosen, each scanned page is read as a document as well
// (one reading gives both): an invoice's or a receipt's fields and a form's
// tick boxes come back for the Fields pane (found.js).

import React, { useState } from 'react';
import { Button, Dialog, Progress, t, tn } from '@rutba/office-ui';
import { OcrLanguagePicker, chosenOcrLanguages, rememberOcrLanguages, recognisePixels, readDocument } from '../../ocr-ui.js';
import { foundOnPage } from './found.js';

const FIELDS_KEY = 'rutba-ocr-find-fields';
/** Whether Recognise Text also finds the fields, as last chosen (yes the first time). */
export function chosenFindFields() {
  try { return localStorage.getItem(FIELDS_KEY) !== 'no'; } catch { return true; }
}
function rememberFindFields(on) {
  try { localStorage.setItem(FIELDS_KEY, on ? 'yes' : 'no'); } catch { /* asked again next time */ }
}
import { getPage, pageText } from './pdfjs.js';

const MAX_DPI = 300;
const MAX_SIDE = 4200;

/** Whether pdf.js finds text on a page already. */
export async function hasText(doc, index) {
  try {
    return (await pageText(doc, index)).replace(/\s+/g, '').length >= 3;
  } catch {
    return false;
  }
}

/**
 * A page drawn to pixels for reading, and the viewport that drew it; at up to
 * `maxDpi` (a bill is read as well at 200 dots per inch, and sooner).
 */
export async function pagePixels(doc, index, { maxDpi = MAX_DPI } = {}) {
  const page = await getPage(doc, index);
  const unit = page.getViewport({ scale: 1 });
  const dpi = Math.min(maxDpi, (MAX_SIDE / Math.max(unit.width, unit.height)) * 72);
  const viewport = page.getViewport({ scale: dpi / 72 });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const ctx = canvas.getContext('2d', { alpha: false, willReadFrequently: true });
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, canvas, viewport, background: '#ffffff' }).promise;
  const image = ctx.getImageData(0, 0, canvas.width, canvas.height);
  return { pixels: { width: canvas.width, height: canvas.height, data: image.data }, viewport, dpi };
}

/**
 * A recognised word in the page's own coordinates: where its baseline
 * starts, how wide it runs and at what angle, and its height. The layer's
 * font rises 0.8 of its size above the baseline and falls 0.2 below, so the
 * baseline is put a fifth of the way up the word's box.
 */
export function wordOnPage(word, viewport) {
  const { x, y, width, height } = word.box;
  const [blx, bly] = viewport.convertToPdfPoint(x, y + height);
  const [brx, bry] = viewport.convertToPdfPoint(x + width, y + height);
  const [tlx, tly] = viewport.convertToPdfPoint(x, y);
  return {
    text: word.text,
    x: blx + (tlx - blx) * 0.2,
    y: bly + (tly - bly) * 0.2,
    width: Math.hypot(brx - blx, bry - bly),
    size: Math.hypot(tlx - blx, tly - bly),
    angle: (Math.atan2(bry - bly, brx - blx) * 180) / Math.PI,
  };
}

/**
 * Recognise the scanned pages among `targets` ([{ source, page }], each
 * once). `onProgress({ done, total, page })` after each; `cancelled()`
 * stops between pages. Returns, per source, the pages to write and counts,
 * and with `fields`, what each page offered as a document
 * (found: [{ source, page, found }]).
 */
export async function recognisePages({ shell, sources, targets, languages, onProgress, cancelled, job, fields = false }) {
  const scanned = [];
  let skipped = 0;
  for (const target of targets) {
    const source = sources[target.source];
    if (!source?.doc) continue;
    if (await hasText(source.doc, target.page)) skipped += 1;
    else scanned.push(target);
  }
  const bySource = new Map();
  let words = 0;
  let done = 0;
  const found = [];
  onProgress?.({ done, total: scanned.length });
  for (const target of scanned) {
    if (cancelled?.()) break;
    const source = sources[target.source];
    const { pixels, viewport, dpi } = await pagePixels(source.doc, target.page);
    if (cancelled?.()) break;
    const result = fields ? await readDocument(shell, pixels, { languages, dpi, job }) : await recognisePixels(shell, pixels, { languages, dpi, job });
    const offered = fields ? foundOnPage(result.document, viewport, { languages }) : null;
    if (offered) found.push({ source: target.source, page: target.page, found: offered });
    const lines = result.lines.map((line) => ({
      words: result.words.slice(line.words[0], line.words[1]).map((w) => wordOnPage(w, viewport)),
    }));
    words += result.words.length;
    if (!bySource.has(target.source)) bySource.set(target.source, []);
    if (lines.length) bySource.get(target.source).push({ page: target.page, lines });
    done += 1;
    onProgress?.({ done, total: scanned.length, page: target.page });
  }
  return { bySource, scanned: scanned.length, skipped, done, words, found };
}

/** The Recognise Text dialog: which pages, which languages; then its progress, with Cancel. */
export function RecogniseDialog({ count, picked, current, progress, onStart, onCancel, onClose }) {
  const [scope, setScope] = useState(picked > 1 ? 'picked' : 'all');
  const [languages, setLanguages] = useState(chosenOcrLanguages);
  const [fields, setFields] = useState(chosenFindFields);
  if (progress) {
    const { done = 0, total = 0 } = progress;
    return (
      <Dialog
        title={t('Recognise Text')}
        onClose={onCancel}
        actions={<Button label={t('Cancel')} data-role="ocr-cancel" onClick={onCancel} />}
      >
        <div className="ocr-progress" data-role="ocr-progress" data-done={done} data-total={total}>
          <div className="ocr-progress-line">
            {total ? t('Reading page {page} of {count}…', { page: Math.min(done + 1, total), count: total }) : t('Finding the scanned pages…')}
          </div>
          <Progress value={total ? done / total : 0} />
          <div className="ocr-foot">{t('The words are read on this computer. Nothing is sent anywhere.')}</div>
        </div>
      </Dialog>
    );
  }
  return (
    <Dialog
      title={t('Recognise Text')}
      onClose={onClose}
      actions={(
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Recognise')} data-role="ocr-go" onClick={() => { rememberOcrLanguages(languages); rememberFindFields(fields); onStart({ scope, languages, fields }); }} />
        </>
      )}
    >
      <div className="ocr-form" data-role="ocr-dialog">
        <div>
          <div className="ocr-form-label">{t('Pages')}</div>
          <label className="ocr-choice">
            <input type="radio" name="ocr-scope" checked={scope === 'all'} onChange={() => setScope('all')} />
            <span>{tn(count, 'All {count} page', 'All {count} pages')}</span>
          </label>
          <label className="ocr-choice">
            <input type="radio" name="ocr-scope" checked={scope === 'picked'} disabled={picked < 1} onChange={() => setScope('picked')} />
            <span>{picked ? tn(picked, 'The {count} selected page', 'The {count} selected pages') : t('The selected pages')}</span>
          </label>
          <label className="ocr-choice">
            <input type="radio" name="ocr-scope" checked={scope === 'current'} disabled={current < 1} onChange={() => setScope('current')} />
            <span>{t('This page ({page})', { page: current })}</span>
          </label>
        </div>
        <div>
          <div className="ocr-form-label">{t('Languages in the pages')}</div>
          <OcrLanguagePicker value={languages} onChange={setLanguages} />
        </div>
        <label className="ocr-choice">
          <input type="checkbox" checked={fields} onChange={(e) => setFields(e.target.checked)} data-role="ocr-find-fields" />
          <span>{t("Find fields: an invoice's or a receipt's details and a form's tick boxes, offered in the Fields pane")}</span>
        </label>
        <div className="ocr-foot">{t('Pages that already have text are left as they are. The words are added as an invisible layer, so the pages look exactly as they did, and can then be searched, selected and copied. Nothing leaves this computer.')}</div>
      </div>
    </Dialog>
  );
}
