// Find Fields: what Recognise Text finds on a scanned bill or form beyond its
// words. A page read as a document (@rutba/ocr's document pass, asked for
// in the Recognise Text dialog) gives an invoice's or a receipt's fields and
// a form's tick boxes; they are offered in the Fields pane beside the pages.
// A field pressed is marked on its page; an invoice's fields copy as one row
// a sheet takes cell by cell; a form's tick boxes, ticked or not as they were
// read, become real check box fields of the PDF's form (@rutba/pdf/forms
// addCheckboxes, an incremental update, so Undo takes them away again).
//
// Everything here is in the page's own coordinates (points, from its
// bottom-left), turned from the picture's pixels by the viewport that drew
// the page for reading, so a rotated or cropped page comes out right.

import React from 'react';
import { Button, Icon, t, tn } from '@rutba/office-ui';
import { documentRecord, checkboxLabels, shownValue } from '@rutba/ocr/record';
import { OcrLimits } from '../../ocr-ui.js';
import { usePageViewport, boxOf } from './fields.js';

/** A box in the picture's pixels as a rectangle on the page: [x0, y0, x1, y1] in points. */
function pageRect(box, viewport) {
  const [ax, ay] = viewport.convertToPdfPoint(box.x, box.y);
  const [bx, by] = viewport.convertToPdfPoint(box.x + box.width, box.y + box.height);
  return [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)].map((v) => Math.round(v * 100) / 100);
}

const FIELD_ORDER = ['supplier', 'number', 'date', 'dueDate', 'subtotal', 'tax', 'total', 'currency'];

/**
 * What a page read as a document offers, or null when it offers nothing:
 * { kind, consistent, languages, fields: [{ field, text, sure, rect }], checkboxes: [{ rect, checked, label }] }.
 */
export function foundOnPage(document, viewport, { languages = [] } = {}) {
  if (!document) return null;
  const bill = document.kind === 'invoice' || document.kind === 'receipt';
  const fields = [];
  // A bill's fields where it has more than a guess at who it is from (a letter that names an invoice is not one).
  const real = ['total', 'subtotal', 'tax', 'number', 'dueDate'].some((k) => document.fields?.[k]);
  if (bill && real) {
    const row = documentRecord(document);
    for (const name of FIELD_ORDER) {
      const f = name === 'dueDate' ? document.fields?.dueDate : null;
      const value = name === 'dueDate' ? f?.value ?? null : row.values[name];
      if (value === null || value === undefined) continue;
      const box = name === 'dueDate' ? f?.box : row.boxes[name];
      const sure = name === 'dueDate' ? !(typeof f?.confidence === 'number' && f.confidence < 75) : !row.check[name];
      fields.push({ field: name, text: shownValue(name, value, row.values.currency), sure, rect: box ? pageRect(box, viewport) : null });
    }
  }
  const rtl = (document.words || []).filter((w) => /[֐-ࣿ]/u.test(w.text)).length > (document.words || []).length / 2;
  const labels = checkboxLabels(document.checkboxes || [], document.words || [], { rtl });
  const checkboxes = (document.checkboxes || []).map((c, i) => ({ rect: pageRect(c.box, viewport), checked: c.checked, label: labels[i] }));
  if (!fields.length && !checkboxes.length) return null;
  return { kind: document.kind, consistent: Boolean(document.checks?.consistent), languages, fields, checkboxes };
}

export const fieldNames = () => ({
  supplier: t('Supplier'), number: t('Number'), date: t('Date'), dueDate: t('Due'), subtotal: t('Subtotal'), tax: t('Tax'), total: t('Total'), currency: t('Currency'),
});

/** The rectangle marked on its page, over the page as drawn. */
export function FoundMark({ entry, source, scale, mark }) {
  const vp = usePageViewport(source, entry, scale);
  if (!vp || !mark || mark.key !== entry.key || !mark.rect) return null;
  const b = boxOf(vp, mark.rect);
  const pad = 3;
  return <div className="pdf-found-mark" data-role="found-mark" style={{ left: b.left - pad, top: b.top - pad, width: b.width + pad * 2, height: b.height + pad * 2 }} />;
}

/** Whether a form's tick box already has a check box field over it. */
function made(box, fields, page) {
  return (fields || []).some((f) => f.type === 'checkbox' && f.widgets.some((w) => w.page === page && w.rect
    && Math.min(w.rect[2], box.rect[2]) - Math.max(w.rect[0], box.rect[0]) > 0 && Math.min(w.rect[3], box.rect[3]) - Math.max(w.rect[1], box.rect[1]) > 0));
}

const KINDS = () => ({ invoice: t('Invoice'), receipt: t('Receipt'), form: t('Form'), letter: t('Letter'), other: t('Page') });

/**
 * The Fields pane: each page that offered something, in page order.
 * `items`: [{ key, number, source, page, found }]; `mark`: { key, field | box } shown on the page.
 */
export function FieldsPane({ items, sources, mark, onMark, onCopyRow, onMake, onClose }) {
  const names = fieldNames();
  const kinds = KINDS();
  const languages = [...new Set(items.flatMap((it) => it.found.languages || []))];
  return (
    <aside className="pdf-comments pdf-found" data-role="fields-pane" aria-label={t('Fields')}>
      <div className="pdf-comments-head">
        <Icon name="scanReceipt" size={15} />
        <span className="pdf-comments-title">{t('Fields')}</span>
        <span className="pdf-comments-count" data-count={items.length}>{items.length}</span>
        <span className="spacer" />
        <Button ghost icon="close" title={t('Close the Fields pane')} onClick={onClose} />
      </div>
      <div className="pdf-comments-list">
        {!items.length ? (
          <div className="pdf-comments-empty">{t('No fields were found. Recognise Text with Find fields ticked reads an invoice\'s or a receipt\'s fields and a form\'s tick boxes.')}</div>
        ) : items.map((it) => {
          const f = it.found;
          const fieldsOfSource = sources[it.source]?.fields || [];
          const waiting = f.checkboxes.filter((c) => !made(c, fieldsOfSource, it.page));
          const ticked = f.checkboxes.filter((c) => c.checked).length;
          return (
            <section key={it.key} className="pdf-found-card" data-role="found-card" data-kind={f.kind} data-key={it.key}>
              <div className="pdf-found-head">
                <strong>{kinds[f.kind] || kinds.other}</strong>
                <span className="pdf-found-page">{t('Page {page}', { page: it.number })}</span>
              </div>
              {f.fields.length ? (
                <>
                  <div className={`pdf-found-sums${f.consistent ? ' good' : ''}`} data-role="found-sums">
                    <Icon name={f.consistent ? 'check' : 'info'} size={13} />
                    <span>{f.consistent ? t('Its sums agree.') : t('Its sums do not agree, or were not all found: check the figures against the page.')}</span>
                  </div>
                  <div className="pdf-found-fields">
                    {f.fields.map((x) => (
                      <button
                        key={x.field}
                        type="button"
                        className={`pdf-found-field${x.sure ? '' : ' unsure'}${mark?.key === it.key && mark.field === x.field ? ' on' : ''}`}
                        data-field={x.field}
                        data-sure={x.sure ? 'yes' : 'no'}
                        disabled={!x.rect}
                        title={t('{field}: show where it was read', { field: names[x.field] })}
                        onClick={() => onMark({ key: it.key, field: x.field, rect: x.rect })}
                      >
                        <span className="k">{names[x.field]}</span>
                        <span className="v" dir="auto">{x.text}</span>
                      </button>
                    ))}
                  </div>
                  <Button icon="copy" label={t('Copy as a Row')} title={t('Copy as a Row: the date, supplier, number, subtotal, tax, total and currency, to paste into a sheet as one row')} data-role="found-copy-row" onClick={() => onCopyRow(it)} />
                </>
              ) : null}
              {f.checkboxes.length ? (
                <>
                  <div className="pdf-found-sub">{tn(f.checkboxes.length, '{count} tick box, {ticked} ticked', '{count} tick boxes, {ticked} ticked', { ticked })}</div>
                  <div className="pdf-found-boxes">
                    {f.checkboxes.map((c, i) => (
                      <button
                        key={i}
                        type="button"
                        className={`pdf-found-box${mark?.key === it.key && mark.box === i ? ' on' : ''}`}
                        data-box={i}
                        data-checked={c.checked ? 'yes' : 'no'}
                        title={t('{field}: show where it was read', { field: c.label })}
                        onClick={() => onMark({ key: it.key, box: i, rect: c.rect })}
                      >
                        <span className="pdf-found-tick">{c.checked ? <Icon name="check" size={11} /> : null}</span>
                        <span dir="auto">{c.label}</span>
                      </button>
                    ))}
                  </div>
                  {waiting.length ? (
                    <Button primary icon="check" label={t('Make Check Box Fields')} title={t('Make Check Box Fields: the tick boxes become check boxes of the form, ticked as they were read, to fill and save')} data-role="found-make" onClick={() => onMake(it, waiting)} />
                  ) : (
                    <div className="pdf-found-done" data-role="found-made"><Icon name="check" size={13} /> {t('These tick boxes are check box fields of the form.')}</div>
                  )}
                </>
              ) : null}
            </section>
          );
        })}
        {languages.length ? <div className="pdf-found-limits"><OcrLimits languages={languages} compact /></div> : null}
      </div>
    </aside>
  );
}

export const FOUND_CSS = `
.pdf-found-mark { position: absolute; z-index: 6; border: 2px solid var(--accent); border-radius: 3px; background: color-mix(in srgb, var(--accent) 14%, transparent); box-shadow: 0 0 0 4px color-mix(in srgb, var(--accent) 18%, transparent); pointer-events: none; animation: rw-fade var(--fast); }
.pdf-found .pdf-comments-list { padding: 10px; display: flex; flex-direction: column; gap: 10px; }
.pdf-found-card { border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); padding: 10px 10px 11px; display: flex; flex-direction: column; gap: 8px; }
.pdf-found-head { display: flex; align-items: baseline; gap: 8px; font-size: 12.5px; }
.pdf-found-head strong { font-weight: 600; color: var(--ink); }
.pdf-found-page { color: var(--ink-3); font-size: 11.5px; }
.pdf-found-sums { display: flex; align-items: flex-start; gap: 6px; font-size: 11.5px; line-height: 1.4; color: var(--ink-2); padding: 5px 8px; border-radius: var(--r-1); background: color-mix(in srgb, var(--warn) 10%, var(--surface)); }
.pdf-found-sums.good { background: color-mix(in srgb, var(--good) 9%, var(--surface)); }
.pdf-found-sums svg { flex: none; margin-top: 1px; }
.pdf-found-sums.good svg { color: var(--good); }
.pdf-found-fields { display: flex; flex-direction: column; gap: 2px; }
.pdf-found-field { display: grid; grid-template-columns: 76px minmax(0, 1fr); align-items: baseline; gap: 8px; padding: 4px 6px; border: 0; border-radius: var(--r-1); background: transparent; font: inherit; font-size: 12px; color: var(--ink); text-align: start; cursor: pointer; }
.pdf-found-field .k { color: var(--ink-3); font-size: 11.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.pdf-found-field .v { font-weight: 600; font-variant-numeric: tabular-nums; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.pdf-found-field:hover:not(:disabled), .pdf-found-box:hover { background: var(--hover); }
.pdf-found-field.on, .pdf-found-box.on { background: var(--selected); }
.pdf-found-field.unsure .v { text-decoration: underline dotted var(--warn); text-underline-offset: 3px; }
.pdf-found-field:disabled { cursor: default; }
.pdf-found-sub { font-size: 11.5px; color: var(--ink-2); }
.pdf-found-boxes { display: flex; flex-direction: column; gap: 1px; }
.pdf-found-box { display: flex; align-items: center; gap: 6px; padding: 4px 6px; border: 0; border-radius: var(--r-1); background: transparent; font: inherit; font-size: 12px; color: var(--ink); text-align: start; cursor: pointer; min-width: 0; }
.pdf-found-box span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.pdf-found-tick { width: 14px; height: 14px; flex: none; display: grid; place-items: center; border: 1.5px solid var(--ink-3); border-radius: 3px; background: var(--surface); }
.pdf-found-box[data-checked='yes'] .pdf-found-tick { border-color: var(--accent-fill); background: var(--accent-fill); color: #fff; }
.pdf-found-done { display: flex; align-items: center; gap: 6px; font-size: 11.5px; color: var(--ink-2); }
.pdf-found-done svg { color: var(--good); }
.pdf-found-card > .rw-btn { align-self: flex-start; }
.pdf-found-limits { margin-top: 2px; }
`;
