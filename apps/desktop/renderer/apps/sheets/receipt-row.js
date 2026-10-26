// Data → From Picture → Receipt or Invoice: a receipt's or an invoice's
// details read into one row of the sheet (date, supplier, number, subtotal,
// tax, total, currency), on this computer.
//
// Each picture, and each page of a PDF, is read as a document on the main
// process's threads (@rutba/ocr's document pass: its labels, its figures read
// again, its sums checked), and shown as a row of a preview that can be
// corrected cell by cell, beside the picture with the box each value was
// read from marked. The figures of a page whose sums agree are as read; on
// any other page they are marked to check, as Data from Picture marks the
// cells it was unsure of. Insert puts the rows in at the selected cell, a
// date as a date and the amounts as numbers, as one step Undo takes back.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Button, Dialog, Icon, Spinner, t, tn } from '@rutba/office-ui';
import { RECORD_FIELDS, documentRecord, recordCells, shownValue } from '@rutba/ocr/record';
import { OcrLanguagePicker, OcrLimits, BoxPreview, BoxLoupe, chosenOcrLanguages, rememberOcrLanguages, ocrLanguageNames, pixelsOf, pixelsUrl, loadPicture, readDocument } from '../../ocr-ui.js';
import { fileUrl } from '../../file-urls.js';
import { openDocument, pagePixels } from '../../pdfjs.js';

/** The pages of one PDF read at most: a bundle of receipts scanned together. */
const MAX_PAGES = 20;
/** Pages read at once: each holds its pixels until the pool has it. */
const AT_ONCE = 2;

export const recordHeadings = () => ({
  date: t('Date'), supplier: t('Supplier'), number: t('Number'), subtotal: t('Subtotal'), tax: t('Tax'), total: t('Total'), currency: t('Currency'),
});

const kindWords = () => ({ invoice: t('Read as an invoice.'), receipt: t('Read as a receipt.') });

const baseName = (p) => String(p).split(/[\\/]/).pop();

/**
 * @param {{ shell, files: string[], at: string, onInsert: (rows: object[][]) => Promise<void>|void, onClose: () => void }} props
 *   files: pictures and PDFs, each picture and each page of a PDF a row
 */
export function ReceiptDialog({ shell, files, at, onInsert, onClose }) {
  const [rows, setRows] = useState([]);
  const [selected, setSelected] = useState(0);
  const [focus, setFocus] = useState('total');
  const [languages, setLanguages] = useState(chosenOcrLanguages);
  const [choosing, setChoosing] = useState(false);
  const [headings, setHeadings] = useState(() => /^[A-Z]+1$/.test(at));
  const [error, setError] = useState(null);
  const sources = useRef([]);
  const run = useRef(0);
  const jobs = useRef(new Set());
  const urls = useRef([]);

  const update = (key, patch) => setRows((all) => all.map((r) => (r.key === key ? { ...r, ...(typeof patch === 'function' ? patch(r) : patch) } : r)));

  // What there is to read: each picture, and each page of each PDF.
  useEffect(() => {
    let live = true;
    (async () => {
      const list = [];
      try {
        for (const file of files) {
          if (/\.pdf$/i.test(file)) {
            const { bytes } = await shell.fs.read({ path: file });
            const doc = await openDocument(bytes);
            const count = Math.min(doc.numPages, MAX_PAGES);
            for (let i = 0; i < count; i++) list.push({ key: `${file}#${i}`, name: baseName(file), page: i + 1, pages: doc.numPages, pdf: doc, index: i });
          } else {
            list.push({ key: file, name: baseName(file), url: fileUrl(file) });
          }
        }
      } catch (err) {
        if (live) setError(t('The file could not be read: {reason}', { reason: err.message }));
        return;
      }
      if (!live) return;
      sources.current = list;
      setRows(list.map((s) => ({ key: s.key, name: s.name, page: s.page || null, pages: s.pages || null, url: s.url || null, state: 'waiting' })));
    })();
    return () => { live = false; };
  }, [shell, files]);

  const readAll = useCallback(async (langs) => {
    const n = ++run.current;
    for (const job of jobs.current) shell.ocr.cancel({ job }).catch(() => {});
    jobs.current.clear();
    const list = sources.current;
    setRows((all) => all.map((r) => ({ ...r, state: 'waiting', record: null, shown: null, failed: null })));
    let next = 0;
    const one = async () => {
      while (next < list.length && n === run.current) {
        const s = list[next++];
        const job = `receipt-${n}-${next}`;
        jobs.current.add(job);
        update(s.key, { state: 'reading' });
        try {
          let pixels;
          let url = s.url;
          if (s.pdf) {
            ({ pixels } = await pagePixels(s.pdf, s.index, { maxDpi: 200 }));
            if (!s.preview) {
              s.preview = await pixelsUrl(pixels);
              urls.current.push(s.preview);
            }
            url = s.preview;
          } else {
            pixels = pixelsOf(await loadPicture(s.url), { max: 3200 });
          }
          if (n !== run.current) return;
          const result = await readDocument(shell, pixels, { languages: langs, job });
          if (n !== run.current) return;
          const record = documentRecord(result.document);
          const shown = Object.fromEntries(RECORD_FIELDS.map((f) => [f, shownValue(f, record.values[f], record.values.currency)]));
          update(s.key, { state: 'done', url, width: pixels.width, height: pixels.height, record, shown, check: { ...record.check } });
        } catch (err) {
          if (n === run.current && err?.code !== 'OCR_CANCELLED') update(s.key, { state: 'failed', failed: err.message });
        } finally {
          jobs.current.delete(job);
        }
      }
    };
    await Promise.all(Array.from({ length: AT_ONCE }, one));
  }, [shell]);

  const count = rows.length;
  useEffect(() => { if (count) readAll(languages); }, [count]); // eslint-disable-line react-hooks/exhaustive-deps

  // Whatever is still being read stops when the dialog goes; the previews made here are let go.
  useEffect(() => () => {
    run.current += 1;
    for (const job of jobs.current) shell.ocr.cancel({ job }).catch(() => {});
    for (const u of urls.current) URL.revokeObjectURL(u);
  }, [shell]);

  const edit = (key, field, value) => update(key, (r) => ({ shown: { ...r.shown, [field]: value }, check: { ...r.check, [field]: false } }));
  const remove = (key) => {
    setRows((all) => all.filter((r) => r.key !== key));
    setSelected((s) => Math.max(0, Math.min(s, rows.length - 2)));
  };

  const reading = rows.some((r) => r.state === 'waiting' || r.state === 'reading');
  const ready = rows.filter((r) => r.state === 'done');
  const toCheck = ready.reduce((n, r) => n + RECORD_FIELDS.filter((f) => r.check?.[f]).length, 0);
  const current = rows[Math.min(selected, rows.length - 1)] || null;
  const names = ocrLanguageNames();
  const heads = recordHeadings();

  const insert = async () => {
    const out = ready.map((r) => recordCells(r.shown));
    if (!out.length) return;
    if (headings) out.unshift(RECORD_FIELDS.map((f) => ({ value: heads[f], text: true })));
    onClose();
    await onInsert(out, ready.length);
  };

  const cancel = () => onClose();

  let note = null;
  if (current?.state === 'done') {
    const kind = kindWords()[current.record.kind] || '';
    if (!current.record.found) note = { tone: 'warn', text: t('Nothing was read on this page. Type its details, or remove its row.') };
    else if (current.record.consistent) note = { tone: 'good', text: `${kind} ${t('Its sums agree, so its figures go in as read.')}`.trim() };
    else note = { tone: 'warn', text: `${kind} ${t('Its sums do not agree, or were not all found: check the marked figures against the picture.')}`.trim() };
  } else if (current?.state === 'failed') {
    note = { tone: 'bad', text: t('This page could not be read: {reason}', { reason: current.failed }) };
  }

  return (
    <Dialog
      title={t('Receipt or Invoice')}
      onClose={cancel}
      width={1060}
      className="rr-dialog"
      actions={(
        <>
          <Button label={t('Cancel')} onClick={cancel} />
          <Button primary label={ready.length > 1 ? tn(ready.length, 'Insert {count} Row', 'Insert {count} Rows') : t('Insert Row')} data-role="rr-insert" disabled={reading || !ready.length} onClick={insert} />
        </>
      )}
    >
      <div className="rr-body" data-role="rr-dialog">
        <div className="rr-picture" data-role="rr-picture">
          {current?.url && current.record?.boxes?.[focus] ? (
            <div className="rr-where" data-role="rr-where" data-field={focus}>
              <div className="rr-where-label">{t('{field}, where it was read', { field: heads[focus] })}</div>
              <BoxLoupe url={current.url} width={current.width} height={current.height} box={current.record.boxes[focus]} />
            </div>
          ) : null}
          {current?.url ? (
            <BoxPreview url={current.url} width={current.width} height={current.height} box={current.record?.boxes?.[focus] || null} label={heads[focus]} />
          ) : <div className="rr-picture-wait"><Spinner /></div>}
          {current ? (
            <div className="rr-caption">{current.page ? t('{name}, page {page} of {count}', { name: current.name, page: current.page, count: current.pages }) : current.name}</div>
          ) : null}
        </div>
        <div className="rr-side">
          {error ? <div className="pd-error" data-role="rr-error">{error}</div> : null}
          <div className="rr-note" data-role="rr-note" data-tone={note?.tone || 'none'} data-check={toCheck}>
            {current?.state === 'reading' || current?.state === 'waiting'
              ? <><Spinner /> <span>{t('Reading {name}…', { name: current.page ? `${current.name} (${current.page})` : current.name })}</span></>
              : note ? <><Icon name={note.tone === 'good' ? 'check' : 'info'} size={15} /> <span>{note.text}</span></> : <span>&nbsp;</span>}
          </div>
          <div className="rr-grid-wrap">
            <table className="rr-grid" data-role="rr-grid" data-rows={rows.length}>
              <thead>
                <tr>
                  <th className="rr-n" aria-label={t('Row')} />
                  {RECORD_FIELDS.map((f) => <th key={f} className={`rr-h-${f}`}>{heads[f]}</th>)}
                  <th className="rr-x" aria-label={t('Remove')} />
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => (
                  <tr key={r.key} data-state={r.state} data-consistent={r.record?.consistent ? 'yes' : 'no'} className={i === selected ? 'sel' : undefined} onMouseDown={() => setSelected(i)}>
                    <td className="rr-n">
                      {r.state === 'reading' || r.state === 'waiting' ? <Spinner /> : r.state === 'failed' ? <Icon name="info" size={14} /> : i + 1}
                    </td>
                    {RECORD_FIELDS.map((f) => (
                      <td key={f} className={`rr-c-${f}`} data-field={f} data-check={r.state === 'done' && r.check?.[f] ? 'yes' : undefined}>
                        <input
                          value={r.shown?.[f] ?? ''}
                          disabled={r.state !== 'done'}
                          dir="auto"
                          placeholder={r.state === 'done' ? '' : '…'}
                          aria-label={t('{field} of row {row}', { field: heads[f], row: i + 1 })}
                          onFocus={() => { setSelected(i); setFocus(f); }}
                          onChange={(e) => edit(r.key, f, e.target.value)}
                        />
                      </td>
                    ))}
                    <td className="rr-x">
                      <Button ghost icon="close" title={t('Remove this row')} data-role="rr-remove" onClick={() => remove(r.key)} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="rr-foot">
            {toCheck ? (
              <span className="rr-count" data-role="rr-count">{tn(toCheck, '{count} value is marked to check.', '{count} values are marked to check.')}</span>
            ) : null}
            <label className="rr-headings">
              <input type="checkbox" checked={headings} onChange={(e) => setHeadings(e.target.checked)} data-role="rr-headings" />
              <span>{t('Put a row of headings above them')}</span>
            </label>
            <span className="pd-where">{t('The rows go in from the selected cell, {cell}, a date as a date and the amounts as numbers.', { cell: at })}</span>
          </div>
          <div className="pd-langs">
            {choosing ? (
              <>
                <OcrLanguagePicker value={languages} onChange={setLanguages} />
                <div className="pd-langs-row">
                  <Button label={t('Read Again')} data-role="rr-read-again" disabled={reading} onClick={() => { rememberOcrLanguages(languages); setChoosing(false); readAll(languages); }} />
                </div>
              </>
            ) : (
              <>
                <div className="pd-langs-row">
                  <span className="pd-langs-label">{t('Languages: {list}', { list: languages.map((c) => names[c]).join(', ') })}</span>
                  <Button ghost label={t('Change…')} data-role="rr-languages" onClick={() => setChoosing(true)} />
                </div>
                <OcrLimits languages={languages} compact />
              </>
            )}
          </div>
        </div>
      </div>
    </Dialog>
  );
}

export const RECEIPT_ROW_CSS = `
.rw-dialog.rr-dialog { max-width: min(94vw, 1080px); }
.rr-where { display: flex; flex-direction: column; gap: 5px; flex: none; }
.rr-where-label { font-size: 11.5px; font-weight: 600; color: var(--ink-2); }
.rr-body { display: grid; grid-template-columns: minmax(220px, 290px) minmax(560px, 1fr); gap: 18px; min-height: 360px; max-height: 66vh; }
.rr-picture { display: flex; flex-direction: column; gap: 6px; min-height: 0; }
.rr-picture .ocr-boxview { flex: none; }
.rr-picture { overflow: auto; background: var(--sunken); border: 1px solid var(--line-soft); border-radius: var(--r-2); padding: 8px; }
.rr-picture-wait { display: grid; place-items: center; min-height: 200px; }
.rr-caption { font-size: 11.5px; color: var(--ink-3); line-height: 1.35; overflow-wrap: anywhere; }
.rr-side { display: flex; flex-direction: column; gap: 10px; min-width: 0; min-height: 0; }
.rr-note { display: flex; align-items: center; gap: 8px; min-height: 34px; padding: 6px 10px; border-radius: var(--r-2); font-size: 12.5px; line-height: 1.4; color: var(--ink-2); background: var(--sunken); border: 1px solid var(--line-soft); }
.rr-note .rw-spinner { width: 15px; height: 15px; flex: none; }
.rr-note[data-tone='good'] { color: var(--ink); background: color-mix(in srgb, var(--good) 9%, var(--surface)); border-color: color-mix(in srgb, var(--good) 28%, var(--line)); }
.rr-note[data-tone='good'] svg { color: var(--good); flex: none; }
.rr-note[data-tone='warn'] { color: var(--ink); background: color-mix(in srgb, var(--warn) 10%, var(--surface)); border-color: color-mix(in srgb, var(--warn) 30%, var(--line)); }
.rr-note[data-tone='warn'] svg, .rr-note[data-tone='bad'] svg { flex: none; }
.rr-note[data-tone='bad'] { color: var(--bad); }
.rr-grid-wrap { flex: 1; min-height: 96px; overflow: auto; border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); }
.rr-grid { border-collapse: collapse; font-size: 12.5px; width: 100%; table-layout: fixed; }
.rr-grid th { position: sticky; top: 0; z-index: 1; text-align: start; font-weight: 600; font-size: 11.5px; color: var(--ink-2); padding: 6px 7px; background: color-mix(in srgb, var(--ink) 4%, var(--surface)); border-bottom: 1px solid var(--line); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.rr-grid td { border-bottom: 1px solid var(--line-soft); border-inline-end: 1px solid var(--line-soft); padding: 0; height: 30px; }
.rr-grid tr.sel td { background: color-mix(in srgb, var(--accent) 6%, var(--surface)); }
.rr-grid td[data-check='yes'], .rr-grid tr.sel td[data-check='yes'] { background: color-mix(in srgb, var(--warn) 22%, var(--surface)); }
.rr-grid input { width: 100%; box-sizing: border-box; border: 0; background: transparent; padding: 5px 7px; font: inherit; color: var(--ink); outline: none; text-overflow: ellipsis; }
.rr-grid input:disabled { color: var(--ink-3); }
.rr-grid input:focus { box-shadow: inset 0 0 0 2px var(--accent); background: var(--surface); }
.rr-grid .rr-n { width: 30px; text-align: center; color: var(--ink-3); font-size: 11.5px; font-variant-numeric: tabular-nums; }
.rr-grid td.rr-n .rw-spinner { width: 13px; height: 13px; display: inline-block; vertical-align: middle; }
.rr-grid .rr-x { width: 34px; text-align: center; border-inline-end: 0; }
.rr-grid .rr-x .rw-btn { opacity: 0.55; }
.rr-grid tr:hover .rr-x .rw-btn { opacity: 1; }
.rr-h-date, .rr-c-date { width: 98px; }
.rr-h-number, .rr-c-number { width: 92px; }
.rr-h-subtotal, .rr-c-subtotal, .rr-h-tax, .rr-c-tax, .rr-h-total, .rr-c-total { width: 80px; }
.rr-h-currency, .rr-c-currency { width: 68px; }
.rr-c-subtotal input, .rr-c-tax input, .rr-c-total input { text-align: end; font-variant-numeric: tabular-nums; }
.rr-foot { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 16px; font-size: 12.5px; color: var(--ink-2); }
.rr-count { color: var(--ink); font-weight: 500; }
.rr-headings { display: flex; align-items: center; gap: 7px; }
.rr-headings input { margin: 0; }
.rr-foot .pd-where { flex-basis: 100%; }
`;
