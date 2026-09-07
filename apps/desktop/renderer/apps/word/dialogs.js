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
