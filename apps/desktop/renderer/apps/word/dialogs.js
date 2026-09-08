// The dialogs behind the Word ribbon.
//
// Each one reaches an engine capability that existed with no way in: a
// hyperlink, a comment, a header or footer, a table of a chosen size, find and
// replace, and the word count that every writer checks and no free suite puts
// within one press.

import React, { useMemo, useState } from 'react';
import { Button, Dialog, Field, Input, Select, Icon, Chip, Empty, formatWhen } from '@rutba/office-ui';

/* ── links ───────────────────────────────────────────────────────────────── */

export function LinkDialog({ current, selectedText, onClose, onApply, onRemove }) {
  const [url, setUrl] = useState(current || '');

  return (
    <Dialog
      title="Link"
      width={480}
      onClose={onClose}
      actions={
        <>
          {current ? <Button label="Remove link" onClick={onRemove} /> : null}
          <Button label="Cancel" onClick={onClose} />
          <Button primary label={current ? 'Update' : 'Add link'} disabled={!url.trim()} onClick={() => onApply(normalise(url))} />
        </>
      }
    >
      <div className="ml-form">
        {selectedText ? (
          <Field label="Text">
            <Input value={selectedText} disabled />
          </Field>
        ) : (
          <p style={{ marginTop: 0, fontSize: 12.5 }}>Select some text first, and it becomes the link.</p>
        )}
        <Field label="Address" hint="A web address, or mailto: for an email link.">
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
      title="Insert table"
      width={420}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Insert" onClick={() => onInsert({ rows: Number(rows), cols: Number(cols), header })} />
        </>
      }
    >
      <div className="ml-form">
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr' }}>
          <Field label="Rows">
            <Input type="number" min="1" max="200" value={rows} onChange={(e) => setRows(e.target.value)} autoFocus />
          </Field>
          <Field label="Columns">
            <Input type="number" min="1" max="30" value={cols} onChange={(e) => setCols(e.target.value)} />
          </Field>
        </div>
        <label className="about-auto">
          <input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} />
          <span>Repeat the first row as a header on every page</span>
        </label>
      </div>
    </Dialog>
  );
}

/* ── headers and footers ─────────────────────────────────────────────────── */

/**
 * A header or footer is lines of text, with the fields a page needs.
 *
 * Word's own editor puts you inside the page and then has to explain what
 * "different first page" means. This asks for the lines and offers the two
 * fields anybody actually uses.
 */
export function BandDialog({ band, current, onClose, onApply }) {
  const [lines, setLines] = useState(() => (current || []).join('\n'));
  const name = band === 'header' ? 'Header' : 'Footer';

  const insert = (token) => setLines((t) => (t ? `${t}${t.endsWith('\n') ? '' : ' '}${token}` : token));

  return (
    <Dialog
      title={name}
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label="Remove" onClick={() => onApply([])} />
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Apply" onClick={() => onApply(lines.split('\n'))} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={`${name} text`} hint="One line per line. A tab moves to the centre, a second to the right.">
          <textarea
            className="rw-input ml-compose-body"
            style={{ minHeight: 90 }}
            rows={4}
            value={lines}
            onChange={(e) => setLines(e.target.value)}
            autoFocus
          />
        </Field>
        <div className="ml-toolbar">
          <Button icon="file" label="Page number" onClick={() => insert('{PAGE}')} />
          <Button icon="file" label="Of total" onClick={() => insert('{PAGE} of {PAGES}')} />
          <Button icon="clock" label="Date" onClick={() => insert('{DATE}')} />
          <Button icon="word" label="File name" onClick={() => insert('{FILENAME}')} />
        </div>
        <p className="rw-hint" style={{ margin: 0 }}>
          Fields in braces are resolved when the page is drawn and when it is printed.
        </p>
      </div>
    </Dialog>
  );
}

/* ── comments ────────────────────────────────────────────────────────────── */

export function CommentDialog({ onClose, onAdd }) {
  const [text, setText] = useState('');
  return (
    <Dialog
      title="New comment"
      width={460}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Add" disabled={!text.trim()} onClick={() => onAdd(text.trim())} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Comment" hint="Anchored to what is selected, and saved into the document.">
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
  const label = kind === 'endnote' ? 'Endnote' : 'Footnote';
  return (
    <Dialog
      title={initial ? `Edit ${label.toLowerCase()}` : `Insert ${label.toLowerCase()}`}
      width={460}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label={initial ? 'Save' : 'Insert'} disabled={!text.trim()} onClick={() => onSave(text.trim())} />
        </>
      }
    >
      <div className="ml-form">
        <Field
          label={label}
          hint={initial ? 'The words change; the number and its reference stay where they are.' : `A raised number at the caret, and these words ${kind === 'endnote' ? 'at the end of the document' : 'under the body'}.`}
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

export function CommentsDialog({ comments, onClose, onGoto }) {
  return (
    <Dialog title="Comments" width={520} onClose={onClose} actions={<Button primary label="Close" onClick={onClose} />}>
      {comments?.length ? (
        <div className="ml-import-folders" style={{ maxHeight: 340 }}>
          {comments.map((c) => (
            <button
              key={c.id}
              type="button"
              className="ml-found-item"
              style={{ border: 0, borderBottom: '1px solid var(--line-soft)', borderRadius: 0 }}
              onClick={() => onGoto(c)}
            >
              <span className="ml-found-logo"><Icon name="reply" size={14} /></span>
              <span className="grow">
                <div className="who">{c.author || 'Someone'}{c.date ? ` · ${formatWhen(c.date)}` : ''}</div>
                <div className="what">{c.text}</div>
              </span>
            </button>
          ))}
        </div>
      ) : (
        <Empty icon="reply" title="No comments">Nothing has been marked up in this document.</Empty>
      )}
    </Dialog>
  );
}

/* ── find and replace ────────────────────────────────────────────────────── */

export function FindDialog({ onClose, onReplaceAll }) {
  const [find, setFind] = useState('');
  const [replace, setReplace] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [note, setNote] = useState(null);

  return (
    <Dialog
      title="Find and replace"
      width={480}
      onClose={onClose}
      actions={
        <>
          <Button label="Close" onClick={onClose} />
          <Button primary label="Replace all" disabled={!find} onClick={async () => setNote(await onReplaceAll(find, replace, matchCase))} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Find">
          <Input value={find} onChange={(e) => setFind(e.target.value)} autoFocus />
        </Field>
        <Field label="Replace with">
          <Input value={replace} onChange={(e) => setReplace(e.target.value)} />
        </Field>
        <label className="about-auto">
          <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
          <span>Match case</span>
        </label>
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
    <Dialog title="Date and time" width={420} onClose={onClose} actions={<Button label="Cancel" onClick={onClose} />}>
      <p style={{ marginTop: 0, fontSize: 12.5 }}>Inserted as text, in your own language and region.</p>
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

/* ── symbols ─────────────────────────────────────────────────────────────── */

const SYMBOLS = [
  ['Common', '— – … • · © ® ™ § ¶ † ‡ ° ± × ÷ ≠ ≤ ≥ ≈ ∞ √ ∑ ∏ ∆ µ « » “ ” ‘ ’ ‹ › ¡ ¿'],
  ['Currency', '€ £ $ ¥ ₹ ₽ ₩ ₺ ₪ ¢ ₫ ₴ ₦ ₱'],
  ['Arrows', '← → ↑ ↓ ↔ ↕ ⇐ ⇒ ⇑ ⇓ ⇔ ↩ ↪ ➜ ➔ ➤'],
  ['Marks', '✓ ✔ ✗ ✘ ☐ ☑ ☒ ★ ☆ ♥ ♦ ♣ ♠ ♪ ♫ ☎ ✉ ⚠ ⚡ ☀ ☁ ☂'],
  ['Greek', 'α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ τ υ φ χ ψ ω Γ Δ Θ Λ Ξ Π Σ Φ Ψ Ω'],
  ['Fractions and numbers', '½ ⅓ ⅔ ¼ ¾ ⅛ ⅜ ⅝ ⅞ ¹ ² ³ ⁴ ⁿ ₀ ₁ ₂ ₃ ①  ② ③ ④ ⑤'],
  ['Letters', 'à á â ä å æ ç è é ê ë ì í î ï ñ ò ó ô ö ø ù ú û ü ý ÿ ß À Á Â Ä Å Æ Ç È É Ê Ë Ñ Ö Ø Ü'],
];

export function SymbolDialog({ onClose, onInsert }) {
  const [set, setSet] = useState(0);
  const chars = SYMBOLS[set][1].split(/\s+/).filter(Boolean);
  return (
    <Dialog title="Symbol" width={520} onClose={onClose} actions={<Button label="Close" onClick={onClose} />}>
      <div className="ml-filters" style={{ flexWrap: 'wrap', marginBottom: 10 }}>
        {SYMBOLS.map(([label], i) => (
          <button key={label} type="button" className={`ml-filter${set === i ? ' on' : ''}`} onClick={() => setSet(i)}>{label}</button>
        ))}
      </div>
      <div className="wd-symbols">
        {chars.map((c, i) => (
          <button key={`${c}${i}`} type="button" className="wd-symbol" title={`U+${c.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`} onClick={() => onInsert(c)}>
            {c}
          </button>
        ))}
      </div>
      <p className="rw-hint">Click a symbol to insert it at the caret. The dialog stays open for the next one.</p>
    </Dialog>
  );
}

/* ── properties ──────────────────────────────────────────────────────────── */

export function PropertiesDialog({ doc, model, onClose }) {
  const blocks = model?.blocks || [];
  const words = blocks.reduce((n, b) => n + ((b.text || (b.runs || []).map((r) => r.text).join('')).match(/[^\s]+/g) || []).length, 0);
  const section = model?.section;
  return (
    <Dialog title="Properties" width={460} onClose={onClose} actions={<Button primary label="Close" onClick={onClose} />}>
      <dl className="about-list">
        <dt>Name</dt><dd>{doc?.name || '—'}</dd>
        <dt>Location</dt><dd style={{ wordBreak: 'break-all' }}>{doc?.path || 'Not saved yet'}</dd>
        <dt>Format</dt><dd>{doc?.source ? doc.source.toUpperCase() : 'DOCX'}{doc?.converted?.from ? ` (opened from ${doc.converted.from.toUpperCase()})` : ''}</dd>
        <dt>Paragraphs</dt><dd>{blocks.filter((b) => !b.container).length.toLocaleString()}</dd>
        <dt>Words</dt><dd>{words.toLocaleString()}</dd>
        <dt>Tables</dt><dd>{new Set(blocks.map((b) => (b.container || '').split(':')[0]).filter(Boolean)).size}</dd>
        <dt>Pictures</dt><dd>{blocks.reduce((n, b) => n + (b.images?.length || 0), 0)}</dd>
        <dt>Comments</dt><dd>{(model?.comments || []).length}</dd>
        <dt>Page</dt><dd>{section ? `${Math.round((section.widthPx / 96) * 25.4)} × ${Math.round((section.heightPx / 96) * 25.4)} mm, ${section.orientation}` : '—'}</dd>
        <dt>Styles</dt><dd>{(model?.styles || []).length}</dd>
        <dt>Unsaved changes</dt><dd>{doc?.dirty ? 'Yes' : 'No'}</dd>
      </dl>
    </Dialog>
  );
}

/* ── shortcuts ───────────────────────────────────────────────────────────── */

const SHORTCUTS = [
  ['Ctrl+N', 'New document'], ['Ctrl+O', 'Open'], ['Ctrl+S', 'Save'], ['Ctrl+Shift+S', 'Save as'], ['Ctrl+P', 'Print'],
  ['Ctrl+Z / Ctrl+Y', 'Undo / Redo'], ['Ctrl+X / C / V', 'Cut / Copy / Paste'], ['Ctrl+A', 'Select all'], ['Ctrl+F', 'Find and replace'],
  ['Ctrl+B / I / U', 'Bold / Italic / Underline'], ['Ctrl+K', 'Link'], ['Ctrl+Enter', 'Page break'],
  ['F11', 'Full screen'], ['Esc', 'Leave full screen'],
];

export function ShortcutsDialog({ onClose }) {
  return (
    <Dialog title="Keyboard shortcuts" width={440} onClose={onClose} actions={<Button primary label="Close" onClick={onClose} />}>
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

/** Every tracked change the document carries, read from the file. */
export function TrackedDialog({ blocks, onClose, onGoto }) {
  const changed = (blocks || []).filter((b) => b.tracked);
  return (
    <Dialog title="Tracked changes" width={560} onClose={onClose} actions={<Button primary label="Close" onClick={onClose} />}>
      {changed.length ? (
        <div className="ml-import-folders" style={{ maxHeight: 340 }}>
          {changed.map((b) => (
            <button key={b.index} type="button" className="ml-found-item" style={{ border: 0, borderBottom: '1px solid var(--line-soft)', borderRadius: 0 }} onClick={() => onGoto(b.index)}>
              <span className="ml-found-logo"><Icon name="eye" size={14} /></span>
              <span className="grow">
                <div className="who">{typeof b.tracked === 'object' ? (b.tracked.author || 'Someone') : 'Changed'}{typeof b.tracked === 'object' && b.tracked.date ? ` · ${formatWhen(b.tracked.date)}` : ''}</div>
                <div className="what">{(b.text || (b.runs || []).map((r) => r.text).join('')).slice(0, 120) || '(empty paragraph)'}</div>
              </span>
            </button>
          ))}
        </div>
      ) : (
        <Empty icon="eye" title="No tracked changes">This document has none recorded.</Empty>
      )}
      <p className="rw-hint">Shown as the file records them. Recording new ones, and accepting or rejecting, is not built yet.</p>
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
    <Dialog title="Word count" width={380} onClose={onClose} actions={<Button primary label="Close" onClick={onClose} />}>
      <dl className="about-list">
        <dt>Words</dt>
        <dd>{counts.words.toLocaleString()}</dd>
        <dt>Characters</dt>
        <dd>{counts.characters.toLocaleString()}</dd>
        <dt>Without spaces</dt>
        <dd>{counts.withoutSpaces.toLocaleString()}</dd>
        <dt>Paragraphs</dt>
        <dd>{counts.paragraphs.toLocaleString()}</dd>
        <dt>Reading time</dt>
        <dd>about {counts.minutes} minute{counts.minutes === 1 ? '' : 's'}</dd>
      </dl>
    </Dialog>
  );
}
