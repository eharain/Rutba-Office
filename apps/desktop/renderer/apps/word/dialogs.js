// The dialogs behind the Word ribbon.
//
// Each one reaches an engine capability that existed with no way in: a
// hyperlink, a comment, a header or footer, a table of a chosen size, find and
// replace, and the word count that every writer checks and no free suite puts
// within one press.

import React, { useMemo, useState } from 'react';
import { Button, Dialog, Field, Input, Select, Icon, Chip, Empty, formatWhen, t, tn, language } from '@rutba/office-ui';

/** The languages whose windows offer Match diacritics: those written in Arabic's letters. */
const RIGHT_TO_LEFT = new Set(['ar', 'ur', 'fa']);

/* ── links ───────────────────────────────────────────────────────────────── */

export function LinkDialog({ current, selectedText, onClose, onApply, onRemove }) {
  const [url, setUrl] = useState(current || '');

  return (
    <Dialog
      title={t('common.link')}
      width={480}
      onClose={onClose}
      actions={
        <>
          {current ? <Button label={t('common.remove-link')} onClick={onRemove} /> : null}
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={current ? t('common.update') : t('common.add-link')} disabled={!url.trim()} onClick={() => onApply(normalise(url))} />
        </>
      }
    >
      <div className="ml-form">
        {selectedText ? (
          <Field label={t('common.text')}>
            <Input value={selectedText} disabled />
          </Field>
        ) : (
          <p style={{ marginTop: 0, fontSize: 12.5 }}>{t('word.dialogs.select-some-text')}</p>
        )}
        <Field label={t('common.address')} hint={t('word.dialogs.web-address-mailto')}>
          <Input value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://office.rutba.io" autoFocus />
        </Field>
      </div>
    </Dialog>
  );
}

/** `office.rutba.io` is an address; `office.rutba.io` with no scheme is not. */
function normalise(url) {
  const text = url.trim();
  if (/^[a-z][a-z0-9+.-]*:/i.test(text)) return text;
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(text)) return `mailto:${text}`;
  return `https://${text}`;
}

/* ── tables ──────────────────────────────────────────────────────────────── */

export function TableDialog({ onClose, onInsert }) {
  const [rows, setRows] = useState(3);
  const [cols, setCols] = useState(3);
  const [header, setHeader] = useState(true);

  return (
    <Dialog
      title={t('word.dialogs.insert-table')}
      width={420}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.insert')} onClick={() => onInsert({ rows: Number(rows), cols: Number(cols), header })} />
        </>
      }
    >
      <div className="ml-form">
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr' }}>
          <Field label={t('common.rows')}>
            <Input type="number" min="1" max="200" value={rows} onChange={(e) => setRows(e.target.value)} autoFocus />
          </Field>
          <Field label={t('common.columns')}>
            <Input type="number" min="1" max="30" value={cols} onChange={(e) => setCols(e.target.value)} />
          </Field>
        </div>
        <label className="about-auto">
          <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
          <span>{t('word.dialogs.repeat-first-row')}</span>
        </label>
      </div>
    </Dialog>
  );
}

/**
 * Table Layout → Properties, as Word's: the table's preferred width and where
 * it sits, the caret's row (its height, whether it may break across pages,
 * whether it repeats as a header) and column (its width), and the table's
 * Alt Text. Only what is changed is sent.
 */
export function TablePropertiesDialog({ current, onClose, onApply }) {
  const [width, setWidth] = useState(current.width || { type: 'auto', value: 100 });
  const [align, setAlign] = useState(current.align || 'left');
  const [rowHeight, setRowHeight] = useState(current.rowHeightCm ?? '');
  const [cantSplit, setCantSplit] = useState(Boolean(current.cantSplit));
  const [header, setHeader] = useState(Boolean(current.header));
  const [columnWidth, setColumnWidth] = useState(current.columnWidthCm ?? '');
  const [title, setTitle] = useState(current.alt?.title || '');
  const [description, setDescription] = useState(current.alt?.description || '');
  const apply = () => {
    const out = {};
    if (width.type !== (current.width?.type || 'auto') || Number(width.value) !== Number(current.width?.value ?? 100)) out.width = { type: width.type, value: width.type === 'dxa' ? Math.round(Number(width.value) * 567) : Number(width.value) };
    if (align !== (current.align || 'left')) out.align = align;
    if (String(rowHeight) !== String(current.rowHeightCm ?? '')) out.rowHeightTwips = rowHeight === '' ? null : Math.round(Number(rowHeight) * 567);
    if (cantSplit !== Boolean(current.cantSplit)) out.cantSplit = cantSplit;
    if (header !== Boolean(current.header)) out.header = header;
    if (String(columnWidth) !== String(current.columnWidthCm ?? '') && columnWidth !== '') out.columnWidthCm = Number(columnWidth);
    if (title !== (current.alt?.title || '') || description !== (current.alt?.description || '')) out.alt = { title, description };
    onApply(out);
  };
  return (
    <Dialog
      title={t('word.dialogs.table-properties')}
      width={480}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.ok')} onClick={apply} />
        </>
      }
    >
      <div className="ml-form">
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr 1fr' }}>
          <Field label={t('word.dialogs.preferred-width')}>
            <Select value={width.type} onChange={(e) => setWidth({ type: e.target.value, value: e.target.value === 'pct' ? 100 : e.target.value === 'dxa' ? 15 : 0 })}>
              <option value="auto">{t('common.none')}</option>
              <option value="pct">{t('word.dialogs.percent-text')}</option>
              <option value="dxa">{t('word.dialogs.centimetres')}</option>
            </Select>
          </Field>
          <Field label={width.type === 'dxa' ? t('common.width-cm') : t('word.dialogs.width')}>
            <Input type="number" min="1" max={width.type === 'dxa' ? 55 : 100} step="0.1" disabled={width.type === 'auto'} value={width.type === 'auto' ? '' : width.value} onChange={(e) => setWidth({ ...width, value: e.target.value })} />
          </Field>
          <Field label={t('common.alignment')}>
            <Select value={align} onChange={(e) => setAlign(e.target.value)}>
              <option value="left">{t('common.left')}</option>
              <option value="center">{t('common.centre')}</option>
              <option value="right">{t('common.right')}</option>
            </Select>
          </Field>
        </div>
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr' }}>
          <Field label={t('word.dialogs.rows-height-cm')}>
            <Input type="number" min="0.1" max="55" step="0.1" placeholder={t('word.dialogs.words-need')} value={rowHeight} onChange={(e) => setRowHeight(e.target.value)} />
          </Field>
          <Field label={t('word.dialogs.columns-width-cm')}>
            <Input type="number" min="0.3" max="50" step="0.1" value={columnWidth} onChange={(e) => setColumnWidth(e.target.value)} />
          </Field>
        </div>
        <label className="about-auto">
          <input type="checkbox" checked={cantSplit} onChange={(e) => setCantSplit(e.target.checked)} />
          <span>{t('word.dialogs.keep-row-one')}</span>
        </label>
        <label className="about-auto">
          <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
          <span>{t('word.dialogs.repeat-header-row')}</span>
        </label>
        <Field label={t('word.dialogs.alt-text-title')}>
          <Input value={title} placeholder={t('word.dialogs.table')} onChange={(e) => setTitle(e.target.value)} />
        </Field>
        <Field label={t('word.dialogs.alt-text-description')}>
          <Input value={description} placeholder={t('word.dialogs.shows-someone-cannot')} onChange={(e) => setDescription(e.target.value)} />
        </Field>
      </div>
    </Dialog>
  );
}

/* ── headers and footers ─────────────────────────────────────────────────── */

/** What Of total puts in: the page number, "of" and the count of pages. */
const OF_TOTAL = '{PAGE} of {PAGES}';

/**
 * A header or footer is lines of text, with the fields a page needs.
 *
 * Word's own editor puts you inside the page and then has to explain what
 * "different first page" means. This asks for the lines and offers the two
 * fields anybody actually uses.
 */
export function BandDialog({ band, current, onClose, onApply }) {
  const [lines, setLines] = useState(() => (current || []).join('\n'));
  const name = band === 'header' ? t('common.header') : t('common.footer');

  // A field goes in at the end, a space before it only when the line does
  // not end in one already ("Page " then Of total is "Page {PAGE} of {PAGES}").
  // Of total after a page number finishes it rather than adding a second.
  const insert = (token) => setLines((t) => {
    if (token === OF_TOTAL && /\{PAGE\}\s*$/.test(t)) return t.replace(/\s*$/, '') + ' of {PAGES}';
    return t ? `${t}${/\s$/.test(t) ? '' : ' '}${token}` : token;
  });
  // Tab types a tab, as the hint says; Shift+Tab still moves on.
  const onKeyDown = (e) => {
    if (e.key !== 'Tab' || e.shiftKey || e.ctrlKey || e.altKey || e.metaKey) return;
    e.preventDefault();
    const el = e.currentTarget;
    const { selectionStart: from, selectionEnd: to, value } = el;
    setLines(value.slice(0, from) + '\t' + value.slice(to));
    requestAnimationFrame(() => { el.selectionStart = el.selectionEnd = from + 1; });
  };

  return (
    <Dialog
      title={name}
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.remove')} onClick={() => onApply([])} />
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.apply')} onClick={() => onApply(lines.split('\n'))} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={band === 'header' ? t('word.dialogs.header-text') : t('word.dialogs.footer-text')} hint={t('word.dialogs.one-line-per')}>
          <textarea
            className="rw-input ml-compose-body"
            style={{ minHeight: 90 }}
            rows={4}
            value={lines}
            onChange={(e) => setLines(e.target.value)}
            onKeyDown={onKeyDown}
            data-role="band-text"
            autoFocus
          />
        </Field>
        <div className="ml-toolbar">
          <Button icon="file" label={t('word.dialogs.page-number')} onClick={() => insert('{PAGE}')} />
          <Button icon="file" label={t('word.dialogs.of-total')} onClick={() => insert(OF_TOTAL)} />
          <Button icon="clock" label={t('common.date')} onClick={() => insert('{DATE}')} />
          <Button icon="word" label={t('common.file-name')} onClick={() => insert('{FILENAME}')} />
        </div>
        <p className="rw-hint" style={{ margin: 0 }}>
          {t('word.dialogs.fields-braces-resolved')}
        </p>
      </div>
    </Dialog>
  );
}

/* ── comments ────────────────────────────────────────────────────────────── */

/* ── watermark ─────────────────────────────────────────────────────────── */

/** Custom watermark: the words, faint and rising across every page. */
export function WatermarkDialog({ current, onClose, onApply }) {
  const [text, setText] = useState(current || '');
  const ok = text.trim().length > 0;
  return (
    <Dialog
      title={t('word.watermark')}
      width={420}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.apply')} className="wd-watermark-ok" disabled={!ok} onClick={() => onApply(text.trim())} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('common.words')} hint={t('word.dialogs.faint-rising-across')}>
          <Input className="wd-watermark-text" value={text} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && ok) onApply(text.trim()); }} autoFocus />
        </Field>
      </div>
    </Dialog>
  );
}

export function CommentDialog({ onClose, onAdd }) {
  const [text, setText] = useState('');
  return (
    <Dialog
      title={t('common.new-comment-2')}
      width={460}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.add')} disabled={!text.trim()} onClick={() => onAdd(text.trim())} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('common.comment')} hint={t('word.dialogs.anchored-selected-saved')}>
          <textarea
            className="rw-input ml-compose-body"
            style={{ minHeight: 100 }}
            rows={4}
            value={text}
            onChange={(e) => setText(e.target.value)}
            autoFocus
          />
        </Field>
      </div>
    </Dialog>
  );
}

/** A footnote or endnote: its words, on the way in or on the way to a change. */
export function NoteDialog({ kind = 'footnote', initial = '', onClose, onSave }) {
  const [text, setText] = useState(initial);
  const label = kind === 'endnote' ? t('word.endnote') : t('word.footnote');
  return (
    <Dialog
      title={initial ? (kind === 'endnote' ? t('word.dialogs.edit-endnote') : t('word.dialogs.edit-footnote')) : (kind === 'endnote' ? t('word.dialogs.insert-endnote') : t('word.dialogs.insert-footnote'))}
      width={460}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={initial ? t('common.save') : t('common.insert')} disabled={!text.trim()} onClick={() => onSave(text.trim())} />
        </>
      }
    >
      <div className="ml-form">
        <Field
          label={label}
          hint={initial ? t('word.dialogs.words-change-number') : kind === 'endnote' ? t('word.dialogs.raised-number-caret-words') : t('word.dialogs.raised-number-caret')}
        >
          <textarea
            className="rw-input ml-compose-body"
            style={{ minHeight: 90 }}
            rows={3}
            value={text}
            onChange={(e) => setText(e.target.value)}
            autoFocus
          />
        </Field>
      </div>
    </Dialog>
  );
}

export function CommentsDialog({ comments, onClose, onGoto, onDelete = null }) {
  return (
    <Dialog title={t('common.comments')} width={520} onClose={onClose} actions={<Button primary label={t('common.close')} onClick={onClose} />}>
      {comments?.length ? (
        <div className="ml-import-folders" style={{ maxHeight: 340 }}>
          {comments.map((c) => (
            <div key={c.id} className="wd-comment-row" style={{ display: 'flex', alignItems: 'center', borderBottom: '1px solid var(--line-soft)' }}>
              <button
                type="button"
                className="ml-found-item grow"
                style={{ border: 0, borderRadius: 0 }}
                onClick={() => onGoto(c)}
              >
                <span className="ml-found-logo"><Icon name="reply" size={14} /></span>
                <span className="grow">
                  <div className="who">{c.author || t('common.someone')}{c.date ? ` · ${formatWhen(c.date)}` : ''}</div>
                  <div className="what">{c.text}</div>
                </span>
              </button>
              {onDelete ? (
                <Button icon="trash" className="wd-comment-delete" title={t('word.dialogs.delete-comment-any')} aria-label={t('word.dialogs.delete-comment-author', { author: c.author || t('common.someone-2') })} onClick={() => onDelete(c)} />
              ) : null}
            </div>
          ))}
        </div>
      ) : (
        <Empty icon="reply" title={t('word.dialogs.no-comments')}>{t('word.dialogs.nothing-marked')}</Empty>
      )}
    </Dialog>
  );
}

/* ── bookmarks ───────────────────────────────────────────────────────────── */

const BOOKMARK_NAME_RE = /^[A-Za-z][A-Za-z0-9_]{0,39}$/;

function bookmarkSpan(b) {
  return b.from === b.to ? t('common.paragraph-number', { number: b.from + 1 }) : t('word.dialogs.paragraphs-from-to', { from: b.from + 1, to: b.to + 1 });
}

/**
 * Insert > Bookmark — Word's older, position-based anchor: a name on a span
 * of paragraphs, so Go To (and, once the engine writes them, a cross-
 * reference) can find the spot again. Adding clears the name field and
 * leaves the dialog open, the way Word's own does, so a document gets several
 * bookmarks in one visit.
 */
export function BookmarkDialog({ bookmarks, onClose, onAdd, onDelete, onGoto }) {
  const [name, setName] = useState('');
  const [selected, setSelected] = useState(null);
  const valid = BOOKMARK_NAME_RE.test(name);
  const picked = (bookmarks || []).some((b) => b.name === selected);

  const add = () => {
    if (!valid) return;
    onAdd(name);
    setName('');
  };

  return (
    <Dialog
      title={t('common.bookmark')}
      width={460}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.close')} onClick={onClose} />
          <Button className="wd-bookmark-delete" label={t('common.delete')} disabled={!picked} onClick={() => { onDelete(selected); setSelected(null); }} />
          <Button className="wd-bookmark-goto" label={t('common.go-to')} disabled={!picked} onClick={() => onGoto(selected)} />
          <Button primary className="wd-bookmark-add" label={t('common.add')} disabled={!valid} onClick={add} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('word.dialogs.bookmark-name')} hint={t('word.dialogs.letters-digits')}>
          <Input
            className="wd-bookmark-name"
            value={name}
            onChange={(e) => setName(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter' && valid) add(); }}
            autoFocus
          />
        </Field>
        {bookmarks?.length ? (
          <div className="ml-import-folders" style={{ maxHeight: 240 }}>
            {bookmarks.map((b) => (
              <button
                key={b.name}
                type="button"
                className={'ml-found-item wd-bookmark-row' + (selected === b.name ? ' picked' : '')}
                style={{ border: 0, borderBottom: '1px solid var(--line-soft)', borderRadius: 0 }}
                onClick={() => setSelected(b.name)}
              >
                <span className="ml-found-logo"><Icon name="flag" size={14} /></span>
                <span className="grow">
                  <div className="who">{b.name}</div>
                  <div className="what">{bookmarkSpan(b)}</div>
                </span>
              </button>
            ))}
          </div>
        ) : (
          <Empty icon="flag" title={t('word.no-bookmarks')}>{t('word.dialogs.name-selected-paragraphs')}</Empty>
        )}
      </div>
    </Dialog>
  );
}

/**
 * Insert → Cross-reference: a REF field to a bookmark, so its words follow
 * the bookmark wherever the document goes — Update Fields refreshes it, the
 * way Word's own F9 does. There is nothing to type here; picking a bookmark
 * is the whole dialog, so a double-click or Enter on the picked row inserts
 * it too, the way Word's own list does.
 */
export function CrossReferenceDialog({ bookmarks, onClose, onInsert }) {
  const [selected, setSelected] = useState(null);
  const picked = (bookmarks || []).some((b) => b.name === selected);

  const insert = (name) => { if (name) onInsert(name); };

  return (
    <Dialog
      title={t('word.cross')}
      width={460}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.close')} onClick={onClose} />
          <Button primary className="wd-xref-insert" label={t('common.insert')} disabled={!picked} onClick={() => insert(selected)} />
        </>
      }
    >
      <div className="ml-form" onKeyDown={(e) => { if (e.key === 'Enter' && picked) { e.preventDefault(); insert(selected); } }}>
        {bookmarks?.length ? (
          <div className="ml-import-folders" style={{ maxHeight: 240 }}>
            {bookmarks.map((b) => (
              <button
                key={b.name}
                type="button"
                className={'ml-found-item wd-xref-row' + (selected === b.name ? ' picked' : '')}
                style={{ border: 0, borderBottom: '1px solid var(--line-soft)', borderRadius: 0 }}
                onClick={() => setSelected(b.name)}
                onDoubleClick={() => insert(b.name)}
              >
                <span className="ml-found-logo"><Icon name="flag" size={14} /></span>
                <span className="grow">
                  <div className="who">{b.name}</div>
                  <div className="what">{bookmarkSpan(b)}</div>
                </span>
              </button>
            ))}
          </div>
        ) : (
          <Empty icon="flag" title={t('word.no-bookmarks')}>{t('word.dialogs.add-bookmark-first')}</Empty>
        )}
      </div>
    </Dialog>
  );
}

const CAPTION_LABELS = ['Figure', 'Table', 'Equation'];
const CAPTION_NAMES = { Figure: t('word.dialogs.figure'), Table: t('common.table'), Equation: t('common.equation') };

/**
 * Insert → Captions → Insert Caption: a label, the caption's own words, and
 * a running SEQ number Update Fields keeps live — `Figure 3: a diagram of
 * it`, previewed here the way Word's own dialog previews it before OK is
 * pressed. The preview counts `fields` (every SEQ this label already has,
 * from the model), the same count the engine itself makes when the caption
 * actually lands — it only differs when the caption is about to be inserted
 * ahead of others with the same label, and Update Fields (or the next
 * caption) settles that, the same as it would after Word's own dialog.
 */
export function CaptionDialog({ fields, onClose, onInsert }) {
  const [label, setLabel] = useState('Figure');
  const [text, setText] = useState('');
  const n = (fields || []).filter((f) => f.kind === 'seq' && f.name === label).length + 1;
  const preview = `${label} ${n}: ${text}`;

  const insert = () => onInsert(label, text);

  return (
    <Dialog
      title={t('common.caption')}
      width={440}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary className="wd-caption-insert" label={t('common.ok')} onClick={insert} />
        </>
      }
    >
      <div className="ml-form" onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); insert(); } }}>
        <Field label={t('common.label')}>
          <Select className="wd-caption-label" value={label} onChange={(e) => setLabel(e.target.value)}>
            {CAPTION_LABELS.map((l) => <option key={l} value={l}>{CAPTION_NAMES[l]}</option>)}
          </Select>
        </Field>
        <Field label={t('word.dialogs.caption-text')}>
          <Input className="wd-caption-text" value={text} onChange={(e) => setText(e.target.value)} autoFocus />
        </Field>
        <p className="wd-caption-preview" style={{ marginTop: 0, fontSize: 12.5, color: 'var(--text-soft)' }}>{preview}</p>
      </div>
    </Dialog>
  );
}

/* ── find and replace ────────────────────────────────────────────────────── */

export function FindDialog({ onClose, onReplaceAll }) {
  const [find, setFind] = useState('');
  const [replace, setReplace] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  // Off, as Word in Arabic has it: a word is found whatever marks it is written with.
  const [matchDiacritics, setMatchDiacritics] = useState(false);
  const [note, setNote] = useState(null);

  return (
    <Dialog
      title={t('common.find-and-replace')}
      width={480}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.close')} onClick={onClose} />
          <Button primary label={t('common.replace-all')} disabled={!find} onClick={async () => setNote(await onReplaceAll(find, replace, matchCase, matchDiacritics))} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('common.find')}>
          <Input value={find} onChange={(e) => setFind(e.target.value)} autoFocus />
        </Field>
        <Field label={t('common.replace-with')}>
          <Input value={replace} onChange={(e) => setReplace(e.target.value)} />
        </Field>
        <label className="about-auto">
          <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
          <span>{t('common.match-case')}</span>
        </label>
        {RIGHT_TO_LEFT.has(language()) ? (
          <label className="about-auto" title={t('common.find-word-only')}>
            <input type="checkbox" className="wd-find-diacritics" checked={matchDiacritics} onChange={(e) => setMatchDiacritics(e.target.checked)} />
            <span>{t('common.match-diacritics')}</span>
          </label>
        ) : null}
        {note ? <div className="ml-note"><Icon name="info" size={14} />{note}</div> : null}
      </div>
    </Dialog>
  );
}

/* ── date and time ───────────────────────────────────────────────────────── */

/** The formats Word offers, produced by the system's own locale. */
const DATE_FORMATS = [
  (d) => d.toLocaleDateString(undefined, { day: 'numeric', month: 'long', year: 'numeric' }),
  (d) => d.toLocaleDateString(undefined, { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }),
  (d) => d.toLocaleDateString(),
  (d) => d.toLocaleDateString(undefined, { day: '2-digit', month: '2-digit', year: 'numeric' }),
  (d) => d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' }),
  (d) => d.toISOString().slice(0, 10),
  (d) => d.toLocaleString(undefined, { day: 'numeric', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' }),
  (d) => d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' }),
  (d) => d.toLocaleTimeString(),
];

export function DateTimeDialog({ onClose, onInsert }) {
  const now = new Date();
  return (
    <Dialog title={t('common.date-and-time')} width={420} onClose={onClose} actions={<Button label={t('common.cancel')} onClick={onClose} />}>
      <p style={{ marginTop: 0, fontSize: 12.5 }}>{t('word.dialogs.inserted-text-own')}</p>
      <div className="ml-found">
        {DATE_FORMATS.map((fmt, i) => {
          const text = fmt(now);
          return (
            <button key={i} type="button" className="ml-found-item" style={{ padding: '7px 12px' }} onClick={() => onInsert(text)}>
              <span className="grow"><div className="who" style={{ fontWeight: 500 }}>{text}</div></span>
              <Icon name="chevronRight" size={13} />
            </button>
          );
        })}
      </div>
    </Dialog>
  );
}

// The dialog is the core's (symbol-dialog.js): Worksheets and Presentations open it too.
export { SymbolDialog } from '../../symbol-dialog.js';

/* ── properties ──────────────────────────────────────────────────────────── */

export function PropertiesDialog({ doc, model, onClose }) {
  const blocks = model?.blocks || [];
  const words = blocks.reduce((n, b) => n + ((b.text || (b.runs || []).map((r) => r.text).join('')).match(/[^\s]+/g) || []).length, 0);
  const section = model?.section;
  return (
    <Dialog title={t('common.properties')} width={460} onClose={onClose} actions={<Button primary label={t('common.close')} onClick={onClose} />}>
      <dl className="about-list">
        <dt>{t('common.name')}</dt><dd>{doc?.name || '\u2014'}</dd>
        <dt>{t('common.location')}</dt><dd style={{ wordBreak: 'break-all' }}>{doc?.path || t('common.not-saved-yet')}</dd>
        <dt>{t('common.format')}</dt><dd>{doc?.source ? doc.source.toUpperCase() : 'DOCX'}{doc?.converted?.from ? ` ${t('word.dialogs.opened-from-format', { format: doc.converted.from.toUpperCase() })}` : ''}</dd>
        <dt>{t('word.dialogs.paragraphs')}</dt><dd>{blocks.filter((b) => !b.container).length.toLocaleString()}</dd>
        <dt>{t('common.words')}</dt><dd>{words.toLocaleString()}</dd>
        <dt>{t('common.tables')}</dt><dd>{new Set(blocks.map((b) => (b.container || '').split(':')[0]).filter(Boolean)).size}</dd>
        <dt>{t('common.pictures')}</dt><dd>{blocks.reduce((n, b) => n + (b.images?.length || 0), 0)}</dd>
        <dt>{t('common.comments')}</dt><dd>{(model?.comments || []).length}</dd>
        <dt>{t('common.page')}</dt><dd>{section ? t('word.dialogs.width-height-mm', { width: Math.round((section.widthPx / 96) * 25.4), height: Math.round((section.heightPx / 96) * 25.4), orientation: section.orientation === 'landscape' ? t('common.landscape-2') : section.orientation === 'portrait' ? t('common.portrait-2') : section.orientation }) : '\u2014'}</dd>
        <dt>{t('common.styles')}</dt><dd>{(model?.styles || []).length}</dd>
        <dt>{t('common.unsaved-changes')}</dt><dd>{doc?.dirty ? t('common.yes') : t('common.no')}</dd>
      </dl>
    </Dialog>
  );
}

/* ── shortcuts ───────────────────────────────────────────────────────────── */

const SHORTCUTS = [
  ['Ctrl+N', t('word.dialogs.new-document')], ['Ctrl+O', t('common.open')], ['Ctrl+S', t('common.save')], ['Ctrl+Shift+S', t('common.save-as-2')], ['Ctrl+P', t('common.print')],
  ['Ctrl+Z / Ctrl+Y', t('common.undo-redo')], ['Ctrl+X / C / V', t('word.dialogs.cut-copy-paste')], ['Ctrl+A', t('common.select-all')], ['Ctrl+F', t('common.find-and-replace')],
  ['Ctrl+B / I / U', t('common.bold-italic-underline')], ['Ctrl+L / E / R / J', t('word.dialogs.align-left-centre')], ['Ctrl+K', t('common.link')], ['Ctrl+Enter', t('word.page-break')],
  ['F11', t('common.full-screen-2')], ['Esc', t('common.leave-full-screen')],
];

export function ShortcutsDialog({ onClose }) {
  return (
    <Dialog title={t('common.keyboard-shortcuts')} width={440} onClose={onClose} actions={<Button primary label={t('common.close')} onClick={onClose} />}>
      <dl className="about-list">
        {SHORTCUTS.map(([keys, what]) => (
          <React.Fragment key={keys}>
            <dt><kbd>{keys}</kbd></dt>
            <dd>{what}</dd>
          </React.Fragment>
        ))}
      </dl>
    </Dialog>
  );
}

/* ── the reviewing pane ──────────────────────────────────────────────────── */

/** What a paragraph's changes are, in words: "2 insertions, formatting, a new paragraph". */
export function trackedKinds(tracked) {
  return [
    tracked.inserted ? tn(tracked.inserted, 'word.dialogs.an-insertion') : null,
    tracked.deleted ? tn(tracked.deleted, 'word.dialogs.a-deletion') : null,
    tracked.formatted ? t('word.dialogs.formatting') : null,
    tracked.moved ? t('word.dialogs.words-moved') : null,
    tracked.mark === 'inserted' ? t('word.dialogs.a-new-paragraph') : tracked.mark === 'deleted' ? t('word.dialogs.paragraphs-joined') : null,
  ].filter(Boolean).join(', ') || t('word.dialogs.changed');
}

/** Every tracked change the document carries, read from the file. */
export function TrackedDialog({ blocks, onClose, onGoto }) {
  const changed = (blocks || []).filter((b) => b.tracked);
  return (
    <Dialog title={t('word.tracked-changes')} width={560} onClose={onClose} actions={<Button primary label={t('common.close')} onClick={onClose} />}>
      {changed.length ? (
        <div className="ml-import-folders" style={{ maxHeight: 340 }}>
          {changed.map((b) => (
            <button key={b.index} type="button" className="ml-found-item" style={{ border: 0, borderBottom: '1px solid var(--line-soft)', borderRadius: 0 }} onClick={() => onGoto(b.index)}>
              <span className="ml-found-logo"><Icon name="eye" size={14} /></span>
              <span className="grow">
                <div className="who">{typeof b.tracked === 'object' ? ((b.tracked.authors || []).join(', ') || t('common.someone')) : t('common.changed')}{typeof b.tracked === 'object' ? <span className="wd-tracked-kinds"> {'\u2014'} {trackedKinds(b.tracked)}</span> : null}</div>
                <div className="what">{(b.text || (b.runs || []).map((r) => r.text).join('')).slice(0, 120) || t('word.dialogs.empty-paragraph')}</div>
              </span>
            </button>
          ))}
        </div>
      ) : (
        <Empty icon="eye" title={t('word.dialogs.no-tracked-changes')}>{t('word.dialogs.document-none')}</Empty>
      )}
      <p className="rw-hint">{t('word.dialogs.click-row-go')}</p>
    </Dialog>
  );
}

/* ── word count ──────────────────────────────────────────────────────────── */

/**
 * Counted from the document, not from a cached number.
 *
 * The reading time is at 238 words a minute, which is the figure the research
 * actually supports for adult silent reading of prose — not the 200 that word
 * processors round to.
 */
export function WordCountDialog({ blocks, onClose }) {
  const counts = useMemo(() => {
    let words = 0;
    let characters = 0;
    let withoutSpaces = 0;
    let paragraphs = 0;

    for (const block of blocks || []) {
      const text = (block.runs || []).map((r) => r.text ?? '').join('') || block.text || '';
      if (text.trim()) paragraphs++;
      characters += text.length;
      withoutSpaces += text.replace(/\s/g, '').length;
      words += (text.match(/[^\s]+/g) || []).length;
    }

    return { words, characters, withoutSpaces, paragraphs, minutes: Math.max(1, Math.round(words / 238)) };
  }, [blocks]);

  return (
    <Dialog title={t('word.dialogs.word-count')} width={380} onClose={onClose} actions={<Button primary label={t('common.close')} onClick={onClose} />}>
      <dl className="about-list">
        <dt>{t('common.words')}</dt>
        <dd>{counts.words.toLocaleString()}</dd>
        <dt>{t('word.dialogs.characters')}</dt>
        <dd>{counts.characters.toLocaleString()}</dd>
        <dt>{t('word.dialogs.without-spaces')}</dt>
        <dd>{counts.withoutSpaces.toLocaleString()}</dd>
        <dt>{t('word.dialogs.paragraphs')}</dt>
        <dd>{counts.paragraphs.toLocaleString()}</dd>
        <dt>{t('word.dialogs.reading-time')}</dt>
        <dd>{tn(counts.minutes, 'word.dialogs.about-count-minute')}</dd>
      </dl>
    </Dialog>
  );
}

/* ── Layout → Hyphenation ────────────────────────────────────────────────── */

const CM = 96 / 2.54;

/**
 * Hyphenation Options, as Word lays the dialog out: hyphenate automatically,
 * hyphenate words in capitals, the hyphenation zone, and a limit on the
 * lines in a row that may end in a hyphen. The zone is shown in centimetres
 * and written in twips.
 */
export function HyphenationDialog({ current, onClose, onApply, onManual }) {
  const [auto, setAuto] = useState(Boolean(current?.auto));
  const [caps, setCaps] = useState(current?.caps !== false);
  const [zone, setZone] = useState(String(Math.round(((current?.zoneTwips ?? 360) / 15 / CM) * 100) / 100));
  const [limit, setLimit] = useState(current?.limit ? String(current.limit) : '');
  const zoneCm = Number(zone);
  const ok = zoneCm >= 0 && zoneCm <= 20 && (limit === '' || (Number(limit) >= 0 && Number.isInteger(Number(limit))));
  const apply = () => onApply({ auto, caps, zoneTwips: Math.round(zoneCm * CM * 15), limit: limit === '' ? 0 : Number(limit) });
  return (
    <Dialog
      title={t('word.hyphenation')}
      width={420}
      onClose={onClose}
      actions={
        <>
          <Button label={t('word.dialogs.manual')} className="wd-hyph-manual" onClick={onManual} />
          <span style={{ flex: 1 }} />
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button primary label={t('common.ok')} className="wd-hyph-ok" disabled={!ok} onClick={apply} />
        </>
      }
    >
      <div className="ml-form">
        <label className="about-auto">
          <input type="checkbox" className="wd-hyph-auto" checked={auto} onChange={(e) => setAuto(e.target.checked)} />
          <span>{t('word.dialogs.automatically')}</span>
        </label>
        <label className="about-auto">
          <input type="checkbox" className="wd-hyph-caps" checked={caps} onChange={(e) => setCaps(e.target.checked)} />
          <span>{t('word.dialogs.hyphenate-words-caps')}</span>
        </label>
        <Field label={t('word.dialogs.hyphenation-zone-cm')} hint={t('word.dialogs.word-broken-only')}>
          <Input type="number" min="0" max="20" step="0.05" className="wd-hyph-zone" value={zone} onChange={(e) => setZone(e.target.value)} style={{ width: 96 }} />
        </Field>
        <Field label={t('word.dialogs.limit-consecutive')} hint={t('word.dialogs.how-many-lines')}>
          <Input type="number" min="0" max="99" className="wd-hyph-limit" placeholder={t('common.no-limit')} value={limit} onChange={(e) => setLimit(e.target.value)} style={{ width: 96 }} />
        </Field>
      </div>
    </Dialog>
  );
}

/**
 * Manual Hyphenation: the words that could fill the line before them, one
 * at a time — "Hyphenate at:", the word with every place it may break, the
 * proposed one marked and any other a click away. Yes puts an optional
 * hyphen there, No leaves the word whole, Cancel stops.
 */
export function ManualHyphenationDialog({ candidate, onYes, onNo, onClose }) {
  const [at, setAt] = useState(null);
  const points = candidate?.points || [];
  const chosen = at ?? candidate?.proposed ?? points[points.length - 1] ?? null;
  React.useEffect(() => { setAt(null); }, [candidate?.key]);
  return (
    <Dialog
      title={t('word.dialogs.manual-hyphenation')}
      width={440}
      onClose={onClose}
      actions={
        <>
          <Button label={t('common.cancel')} onClick={onClose} />
          <Button label={t('common.no')} className="wd-hyph-no" disabled={!candidate} onClick={onNo} />
          <Button primary label={t('common.yes')} className="wd-hyph-yes" disabled={!candidate || chosen == null} onClick={() => onYes(chosen)} />
        </>
      }
    >
      {candidate ? (
        <div className="ml-form">
          <Field label={t('word.dialogs.hyphenate-at')} hint={t('word.dialogs.click-another-place')}>
            <div className="wd-hyph-word" data-word={candidate.word}>
              {[...candidate.word].map((ch, i) => (
                <React.Fragment key={i}>
                  {i > 0 && points.includes(i) ? (
                    <button type="button" className={`wd-hyph-point${chosen === i ? ' on' : ''}`} data-point={i} title={t('word.dialogs.break-after-text', { text: candidate.word.slice(0, i) })} onClick={() => setAt(i)}>-</button>
                  ) : null}
                  <span>{ch}</span>
                </React.Fragment>
              ))}
            </div>
          </Field>
        </div>
      ) : (
        <p className="wd-hyph-done">{t('word.dialogs.no-more-words')}</p>
      )}
    </Dialog>
  );
}
