// An attachment's details: a PDF or a picture attached to a message, read
// as a document when the person opens the message, and when it reads as an
// invoice or a receipt, a small note under the attachments says who it is
// from, the total, the due date and the number. Each detail, pressed, shows
// the part of the attachment it was read from.
//
// The reading is done on the main process's text recognition threads
// (@rutba/ocr's document pass), one attachment at a time, only for the
// message open in the reading pane: never as mail arrives or is imported,
// never in the background. Moving to another message stops it. What was read
// is kept with the message (mail.keepReading), so opening it again reads
// nothing; a page that is not a bill is kept too, as nothing to show.

import React, { useEffect, useRef, useState } from 'react';
import { Button, Icon, Spinner, t, language } from '@rutba/office-ui';
import { invoiceNote } from '@rutba/ocr/record';
import { BoxLoupe, BoxPreview, OcrLimits, OCR_CSS, chosenOcrLanguages, loadPicture, pixelsOf, pixelsUrl, readDocument } from '../../ocr-ui.js';
import { openDocument } from '../pdf/pdfjs.js';
import { pagePixels } from '../pdf/ocr.js';

/** The reading kept with a message is of this shape; another is read again. */
const VERSION = 1;
/** Larger attachments are not read: a bill is a page or two. */
const MAX_BYTES = 20 * 1024 * 1024;

const PICTURE = /\.(png|jpe?g|webp|bmp|gif)$/i;

/** Whether an attachment is a PDF or a picture the window can read. */
export function readableAttachment(a) {
  if (!a || a.inline || !a.stored || (a.size && a.size > MAX_BYTES)) return false;
  const type = String(a.type || '').toLowerCase();
  const name = String(a.filename || '');
  if (type.startsWith('application/pdf') || /\.pdf$/i.test(name)) return 'pdf';
  if (/^image\/(png|jpe?g|webp|bmp|gif)/.test(type) || PICTURE.test(name)) return 'picture';
  return false;
}

/**
 * An attachment's first page as pixels to read, and a picture of it to show.
 * The pixels are those the reading measured its boxes in (a PDF's page at 200
 * dots per inch, a picture at most 3,200 pixels on its long side), so a kept
 * reading's boxes fall on the picture drawn again.
 */
async function attachmentPage(shell, args, kind, { show = false } = {}) {
  const held = await shell.mail.attachment(args);
  if (!held) throw new Error(t('That attachment is not stored.'));
  const bytes = new Uint8Array(await (await fetch(held.url)).arrayBuffer());
  if (kind === 'pdf') {
    const doc = await openDocument(bytes);
    try {
      const { pixels } = await pagePixels(doc, 0, { maxDpi: 200 });
      return { pixels, url: show ? await pixelsUrl(pixels) : null };
    } finally {
      doc.destroy?.();
    }
  }
  const url = URL.createObjectURL(new Blob([bytes], { type: held.type || 'image/png' }));
  try {
    const pixels = pixelsOf(await loadPicture(url), { max: 3200 });
    if (show) return { pixels, url };
    URL.revokeObjectURL(url);
    return { pixels, url: null };
  } catch (err) {
    URL.revokeObjectURL(url);
    throw err;
  }
}

/**
 * The notes of the message open: { [attachment index]: { state, name, kind, reading, error } },
 * state 'reading', 'done' (an invoice or a receipt), 'none' (not one) or 'failed'.
 */
export function useAttachmentNotes({ shell, where, message }) {
  const [notes, setNotes] = useState({});
  const current = useRef(message);
  current.current = message;
  // The message must be the one asked for: just after a move to another, the old one is still shown.
  const key = where && message && (message.id === undefined || message.id === where.id) ? `${where.accountId}\u0001${where.folder}\u0001${where.id}\u0001${(message.attachments || []).length}` : '';
  useEffect(() => {
    setNotes({});
    const msg = current.current;
    if (!key || !msg) return undefined;
    const list = (msg.attachments || []).map((a, index) => ({ a, index, kind: readableAttachment(a) })).filter((x) => x.kind);
    if (!list.length) return undefined;
    let live = true;
    const jobs = [];
    const set = (index, value) => { if (live) setNotes((all) => ({ ...all, [index]: value })); };
    (async () => {
      for (const { a, index, kind } of list) {
        if (!live) return;
        const args = { accountId: where.accountId, folder: where.folder, id: where.id, index };
        const kept = await shell.mail.attachmentReading(args).catch(() => null);
        if (!live) return;
        if (kept?.version === VERSION) {
          set(index, { state: kept.note ? 'done' : 'none', name: a.filename, kind, reading: kept, kept: true });
          continue;
        }
        set(index, { state: 'reading', name: a.filename, kind });
        const job = `mail-doc-${where.id}-${index}`;
        jobs.push(job);
        try {
          const { pixels } = await attachmentPage(shell, args, kind);
          if (!live) return;
          const languages = chosenOcrLanguages();
          const result = await readDocument(shell, pixels, { languages, job });
          if (!live) return;
          const reading = { version: VERSION, languages, kind: result.document?.kind || 'other', note: invoiceNote(result.document), width: pixels.width, height: pixels.height };
          set(index, { state: reading.note ? 'done' : 'none', name: a.filename, kind, reading });
          await shell.mail.keepReading({ ...args, reading }).catch(() => {});
        } catch (err) {
          if (err?.code !== 'OCR_CANCELLED') set(index, { state: 'failed', name: a.filename, kind, error: err?.message || String(err) });
        }
      }
    })();
    return () => {
      live = false;
      for (const job of jobs) shell.ocr.cancel({ job }).catch(() => {});
    };
  }, [shell, key]); // eslint-disable-line react-hooks/exhaustive-deps
  return notes;
}

/** An amount in its currency, as the window's language writes it. */
function money(value, currency) {
  if (typeof value !== 'number') return String(value ?? '');
  try {
    if (currency && /^[A-Z]{3}$/.test(currency)) return new Intl.NumberFormat(language(), { style: 'currency', currency }).format(value);
  } catch { /* a code Intl does not know: the figure alone */ }
  return value.toLocaleString(language(), { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/** A date (yyyy-mm-dd) as the window's language writes it. */
function day(iso) {
  const at = Date.parse(`${iso}T00:00:00Z`);
  if (!Number.isFinite(at)) return String(iso || '');
  try {
    return new Intl.DateTimeFormat(language(), { dateStyle: 'medium', timeZone: 'UTC' }).format(at);
  } catch {
    return String(iso);
  }
}

/** The facts of a note, in the order shown: { field, label, text, sure }. */
export function noteFacts(note) {
  const out = [];
  if (note.from) out.push({ field: 'from', label: t('From'), text: String(note.from.value), sure: note.from.sure });
  if (note.total) out.push({ field: 'total', label: t('Total'), text: money(note.total.value, note.currency), sure: note.total.sure });
  if (note.dueDate) out.push({ field: 'dueDate', label: t('Due'), text: day(note.dueDate.value), sure: note.dueDate.sure });
  if (note.number) out.push({ field: 'number', label: t('Number'), text: String(note.number.value), sure: note.number.sure });
  return out;
}

/** The notes under a message's attachments, and the part of an attachment a detail was read from. */
export function AttachmentNotes({ shell, where, notes }) {
  const [shown, setShown] = useState(null); // { index, field, url }
  const made = useRef(null);
  useEffect(() => { setShown(null); }, [where?.accountId, where?.folder, where?.id]);
  useEffect(() => () => { if (made.current) URL.revokeObjectURL(made.current); }, []);
  const entries = Object.entries(notes).filter(([, n]) => n.state === 'reading' || n.state === 'done');
  if (!entries.length) return null;

  const show = async (index, field, entry) => {
    if (shown && shown.index === index && shown.field === field) { setShown(null); return; }
    if (shown && shown.index === index && shown.url) { setShown({ ...shown, field }); return; }
    setShown({ index, field, url: null });
    try {
      const { url } = await attachmentPage(shell, { accountId: where.accountId, folder: where.folder, id: where.id, index }, entry.kind, { show: true });
      if (made.current) URL.revokeObjectURL(made.current);
      made.current = url;
      setShown((s) => (s && s.index === index ? { ...s, url } : s));
    } catch {
      setShown(null);
    }
  };

  return (
    <div className="ml-docnotes" data-role="ml-docnotes">
      <style>{OCR_CSS + DOCNOTE_CSS}</style>
      {entries.map(([index, n]) => {
        if (n.state === 'reading') {
          return (
            <div key={index} className="ml-docnote reading" data-role="ml-docnote" data-state="reading">
              <Spinner />
              <span>{t('Looking for an invoice or a receipt in {name}…', { name: n.name })}</span>
            </div>
          );
        }
        const note = n.reading.note;
        const facts = noteFacts(note);
        const open = shown && String(shown.index) === String(index) ? shown : null;
        const box = open ? note[open.field]?.box : null;
        const fact = open ? facts.find((f) => f.field === open.field) : null;
        return (
          <div key={index} className="ml-docnote" data-role="ml-docnote" data-state="done" data-kind={note.kind} data-index={index}>
            <div className="ml-docnote-row">
              <span className="ml-docnote-icon"><Icon name="scanReceipt" size={17} /></span>
              <div className="ml-docnote-main">
                <div className="ml-docnote-title">
                  <strong>{note.kind === 'invoice' ? t('Invoice') : t('Receipt')}</strong>
                  <span className="ml-docnote-file">{n.name}</span>
                </div>
                <div className="ml-docnote-facts">
                  {facts.map((f) => (
                    <button
                      key={f.field}
                      type="button"
                      className={`ml-docnote-fact${f.sure ? '' : ' unsure'}${open?.field === f.field ? ' on' : ''}`}
                      data-field={f.field}
                      data-sure={f.sure ? 'yes' : 'no'}
                      title={f.sure ? t('{field}: show where it was read', { field: f.label }) : t('{field}: show where it was read. Check it against the attachment: the reading was not sure of it.', { field: f.label })}
                      onClick={() => show(index, f.field, n)}
                    >
                      <span className="k">{f.label}</span>
                      <span className="v" dir="auto">{f.text}</span>
                    </button>
                  ))}
                </div>
              </div>
            </div>
            {open ? (
              <div className="ml-docnote-peek" data-role="ml-docnote-peek" data-field={open.field}>
                <div className="ml-docnote-peek-head">
                  <span>{t('{field}, where it was read in {name}', { field: fact?.label || '', name: n.name })}</span>
                  <Button ghost icon="close" title={t('Close')} onClick={() => setShown(null)} />
                </div>
                {open.url ? (
                  <div className="ml-docnote-peek-body">
                    <BoxLoupe url={open.url} width={n.reading.width} height={n.reading.height} box={box} className="ml-docnote-loupe" />
                    <div className="ml-docnote-page">
                      <BoxPreview url={open.url} width={n.reading.width} height={n.reading.height} box={box} label={fact?.label} />
                    </div>
                  </div>
                ) : <div className="ml-docnote-wait"><Spinner /></div>}
                <OcrLimits languages={n.reading.languages || []} compact />
              </div>
            ) : null}
          </div>
        );
      })}
    </div>
  );
}

export const DOCNOTE_CSS = `
.ml-docnotes { display: flex; flex-direction: column; gap: 8px; margin-top: 10px; }
.ml-docnote { border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); padding: 9px 12px; font-size: 12.5px; }
.ml-docnote.reading { display: flex; align-items: center; gap: 9px; color: var(--ink-3); background: transparent; border-style: dashed; padding: 7px 12px; }
.ml-docnote.reading .rw-spinner { width: 14px; height: 14px; flex: none; }
.ml-docnote-row { display: flex; align-items: flex-start; gap: 11px; }
.ml-docnote-icon { width: 30px; height: 30px; border-radius: 8px; display: grid; place-items: center; flex: none; color: var(--accent-text); background: color-mix(in srgb, var(--accent) 11%, var(--surface)); }
.ml-docnote-main { flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 5px; }
.ml-docnote-title { display: flex; align-items: baseline; gap: 8px; min-width: 0; }
.ml-docnote-title strong { font-weight: 600; color: var(--ink); }
.ml-docnote-file { color: var(--ink-3); font-size: 11.5px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-docnote-facts { display: flex; flex-wrap: wrap; gap: 6px; }
.ml-docnote-fact { display: inline-flex; align-items: baseline; gap: 6px; max-width: 100%; padding: 3px 9px; border-radius: 999px; border: 1px solid var(--line); background: var(--chrome); color: var(--ink); font: inherit; font-size: 12px; cursor: pointer; }
.ml-docnote-fact .k { color: var(--ink-3); font-size: 11px; }
.ml-docnote-fact .v { font-weight: 600; font-variant-numeric: tabular-nums; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-docnote-fact:hover { border-color: var(--accent-line); background: var(--selected); }
.ml-docnote-fact.on { border-color: var(--accent); background: color-mix(in srgb, var(--accent) 10%, var(--surface)); }
.ml-docnote-fact.unsure .v { text-decoration: underline dotted color-mix(in srgb, var(--warn) 80%, transparent); text-underline-offset: 3px; }
.ml-docnote-fact.unsure::after { content: ''; width: 6px; height: 6px; border-radius: 50%; background: var(--warn); align-self: center; }
.ml-docnote-peek { margin-top: 10px; padding-top: 10px; border-top: 1px solid var(--line-soft); display: flex; flex-direction: column; gap: 8px; }
.ml-docnote-peek-head { display: flex; align-items: center; gap: 8px; font-size: 11.5px; font-weight: 600; color: var(--ink-2); }
.ml-docnote-peek-head span { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ml-docnote-peek-body { display: grid; grid-template-columns: minmax(0, 560px) 120px; justify-content: start; gap: 14px; align-items: start; }

.ml-docnote-page { width: 120px; }
.ml-docnote-wait { display: grid; place-items: center; min-height: 80px; }
`;
