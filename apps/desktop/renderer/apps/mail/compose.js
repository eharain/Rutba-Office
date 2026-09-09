// Writing a message.
//
// A compose window is where a mail client is most often disappointing: plain
// text only, no way to attach without a file dialog dance, no second thought
// once Send is pressed. This one writes rich text, takes attachments by drop or
// by dialog, keeps a signature, and — because everybody has sent the wrong mail
// to the wrong person — never sends immediately. Send puts the message in a
// queue with a few seconds on it, and those seconds belong to the sender.
//
// The editor is a contentEditable with execCommand behind it. That interface is
// deprecated on paper and is still what every web mail client in the world uses
// for compose, because the alternative — a full document model — is the Word
// editor next door, and a reply does not need one.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Button, Dialog, Field, Input, Icon, Select, Separator, formatBytes } from '@rutba/office-ui';
import { stripTags } from './parts.js';

const exec = (command, value) => {
  try {
    document.execCommand(command, false, value);
  } catch {
    /* a command this engine does not know; the text is unharmed */
  }
};

/** Local midnight-relative choices, the way a person thinks about "later". */
function scheduleChoices() {
  const now = new Date();
  const at = (days, hour) => {
    const d = new Date(now);
    d.setDate(d.getDate() + days);
    d.setHours(hour, 0, 0, 0);
    return d;
  };
  const tonight = at(0, 21);
  const options = [];
  if (tonight > now) options.push({ label: 'Later today, 9pm', at: tonight });
  options.push({ label: 'Tomorrow morning, 8am', at: at(1, 8) });
  options.push({ label: 'Tomorrow afternoon, 1pm', at: at(1, 13) });
  const monday = new Date(now);
  monday.setDate(monday.getDate() + ((8 - monday.getDay()) % 7 || 7));
  monday.setHours(8, 0, 0, 0);
  options.push({ label: `Monday morning, 8am`, at: monday });
  return options;
}

/**
 * An address line that completes as you type: the cards in Contacts first,
 * then the people mail has seen, once each. The token under the caret is
 * what is matched; picking one replaces it with "Name <address>" and a comma.
 */
function AddressInput({ value, onChange, shell, placeholder, autoFocus }) {
  const [hits, setHits] = useState([]);
  const [on, setOn] = useState(-1);
  const [open, setOpen] = useState(false);
  const timer = useRef(null);
  const tokenOf = (text) => {
    const parts = String(text || '').split(',');
    return { head: parts.slice(0, -1).map((s) => s.trim()).filter(Boolean), tail: (parts[parts.length - 1] || '').trim() };
  };
  const lookup = (text) => {
    clearTimeout(timer.current);
    const { tail } = tokenOf(text);
    if (tail.length < 2 || !shell?.contacts) {
      setHits([]);
      setOpen(false);
      return;
    }
    timer.current = setTimeout(async () => {
      try {
        const found = await shell.contacts.suggest({ query: tail, limit: 8 });
        setHits(found);
        setOn(found.length ? 0 : -1);
        setOpen(found.length > 0);
      } catch {
        setHits([]);
      }
    }, 120);
  };
  const pick = (hit) => {
    const { head } = tokenOf(value);
    const piece = hit.name && hit.name !== hit.email ? `${hit.name} <${hit.email}>` : hit.email;
    onChange([...head, piece].join(', ') + ', ');
    setHits([]);
    setOpen(false);
  };
  const onKeyDown = (e) => {
    if (!open || !hits.length) return;
    if (e.key === 'ArrowDown') { e.preventDefault(); setOn((i) => (i + 1) % hits.length); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); setOn((i) => (i - 1 + hits.length) % hits.length); }
    else if (e.key === 'Enter' || e.key === 'Tab') { if (on >= 0) { e.preventDefault(); pick(hits[on]); } }
    else if (e.key === 'Escape') { e.preventDefault(); setOpen(false); }
  };
  return (
    <div className="ml-address">
      <Input value={value || ''} placeholder={placeholder} autoFocus={autoFocus} onChange={(e) => { onChange(e.target.value); lookup(e.target.value); }} onKeyDown={onKeyDown} onBlur={() => setTimeout(() => setOpen(false), 150)} />
      {open ? (
        <div className="ml-suggest" role="listbox">
          {hits.map((h, i) => (
            <button type="button" key={h.email} className={i === on ? 'on' : ''} onMouseDown={(e) => { e.preventDefault(); pick(h); }}>
              <span className="n">{h.name || h.email}</span>
              {h.name ? <span className="e">{h.email}</span> : null}
              <span className="s">{h.source === 'contacts' ? 'contact' : 'mail'}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

export default function Compose({ draft, accounts, accountId, onAccount, onChange, onClose, onSend, onSaveDraft, shell, toast }) {
  const bodyRef = useRef(null);
  const [rich, setRich] = useState(draft.rich !== false);
  const [showCc, setShowCc] = useState(Boolean(draft.cc || draft.bcc));
  const [scheduling, setScheduling] = useState(false);
  const account = accounts.find((a) => a.id === accountId);

  // The editor is uncontrolled on purpose: writing innerHTML back on every
  // keystroke destroys the caret, and a compose box that loses your place after
  // every letter is worse than no formatting at all. It is seeded once, and read
  // when the message is sent.
  useEffect(() => {
    const el = bodyRef.current;
    if (!el || !rich) return;
    if (el.innerHTML.trim() === '') el.innerHTML = draft.html || (draft.text ? textToHtml(draft.text) : '');
    // Put the caret at the top, above any quoted reply, which is where the
    // person is about to type.
    const range = document.createRange();
    range.setStart(el, 0);
    range.collapse(true);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    el.focus();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rich]);

  const collect = useCallback(() => {
    const el = bodyRef.current;
    const html = rich && el ? el.innerHTML : null;
    return {
      ...draft,
      rich,
      html: rich ? html : null,
      text: rich ? stripTags(html) : draft.text || '',
      accountId,
    };
  }, [draft, rich, accountId]);

  const attach = useCallback(async () => {
    const paths = await shell.dialog.open({ title: 'Attach files', multiple: true });
    if (!paths?.length) return;
    const added = [];
    for (const p of paths) {
      const name = String(p).split(/[\\/]/).pop();
      let size = 0;
      try {
        size = (await shell.fs.stat({ path: p }))?.size ?? 0;
      } catch {
        /* a file we cannot stat is still worth attaching */
      }
      added.push({ filename: name, path: p, size });
    }
    onChange({ ...draft, attachments: [...(draft.attachments || []), ...added] });
  }, [shell, draft, onChange]);

  const total = (draft.attachments || []).reduce((n, a) => n + (a.size || 0), 0);

  return (
    <Dialog
      title={draft.subject ? `Message — ${draft.subject}` : 'New message'}
      width={760}
      onClose={onClose}
      actions={
        <>
          <Button label="Discard" onClick={onClose} />
          <Button label="Save draft" onClick={() => onSaveDraft(collect())} />
          <Button icon="clock" label="Send later" disabled={!draft.to} onClick={() => setScheduling((s) => !s)} />
          <Button primary icon="send" label="Send" disabled={!draft.to} onClick={() => onSend(collect(), null)} />
        </>
      }
    >
      <div className="ml-form">
        {accounts.length > 1 ? (
          <Field label="From">
            <Select value={accountId || ''} onChange={(e) => onAccount(e.target.value)}>
              {accounts.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.name ? `${a.name} <${a.email}>` : a.email}
                </option>
              ))}
            </Select>
          </Field>
        ) : (
          <Field label="From">
            <Input value={account ? `${account.name || ''} <${account.email}>` : ''} disabled />
          </Field>
        )}

        <Field label="To">
          <div className="ml-to-row">
            <span className="grow">
              <AddressInput
                autoFocus
                shell={shell}
                value={draft.to || ''}
                onChange={(to) => onChange({ ...draft, to })}
                placeholder="someone@example.com, another@example.com"
              />
            </span>
            <button type="button" className={`ml-linkchip${showCc ? ' on' : ''}`} onClick={() => setShowCc((s) => !s)}>
              Cc / Bcc
            </button>
          </div>
        </Field>

        {showCc ? (
          <>
            <Field label="Cc">
              <AddressInput shell={shell} value={draft.cc || ''} onChange={(cc) => onChange({ ...draft, cc })} />
            </Field>
            <Field label="Bcc" hint="Nobody on this line is visible to the other recipients.">
              <AddressInput shell={shell} value={draft.bcc || ''} onChange={(bcc) => onChange({ ...draft, bcc })} />
            </Field>
          </>
        ) : null}

        <Field label="Subject">
          <Input value={draft.subject || ''} onChange={(e) => onChange({ ...draft, subject: e.target.value })} />
        </Field>

        {rich ? (
          <div>
            <div className="ml-toolbar">
              <Button icon="bold" title="Bold" onClick={() => exec('bold')} />
              <Button icon="italic" title="Italic" onClick={() => exec('italic')} />
              <Button icon="underline" title="Underline" onClick={() => exec('underline')} />
              <Button icon="strike" title="Strikethrough" onClick={() => exec('strikeThrough')} />
              <Separator />
              <Button icon="listBullet" title="Bulleted list" onClick={() => exec('insertUnorderedList')} />
              <Button icon="listNumber" title="Numbered list" onClick={() => exec('insertOrderedList')} />
              <Button icon="alignLeft" title="Align left" onClick={() => exec('justifyLeft')} />
              <Button icon="alignCenter" title="Centre" onClick={() => exec('justifyCenter')} />
              <Separator />
              <Button
                icon="link"
                title="Insert link"
                onClick={() => {
                  const url = window.prompt('Link address');
                  if (url) exec('createLink', url);
                }}
              />
              <Button icon="formula" title="Quote" onClick={() => exec('formatBlock', 'blockquote')} />
              <Button icon="undo" title="Clear formatting" onClick={() => exec('removeFormat')} />
              <Separator />
              <Button icon="attach" label="Attach" onClick={attach} />
              <Button icon="file" title="Plain text" onClick={() => setRich(false)} />
            </div>
            <div
              ref={bodyRef}
              className="ml-rich"
              contentEditable
              suppressContentEditableWarning
              spellCheck
              onDrop={(e) => {
                const paths = [...(e.dataTransfer?.files || [])].map((f) => f.path).filter(Boolean);
                if (!paths.length) return;
                e.preventDefault();
                onChange({
                  ...draft,
                  attachments: [
                    ...(draft.attachments || []),
                    ...[...e.dataTransfer.files].map((f) => ({ filename: f.name, path: f.path, size: f.size })),
                  ],
                });
              }}
            />
          </div>
        ) : (
          <Field label="Message">
            <div className="ml-toolbar">
              <Button icon="attach" label="Attach" onClick={attach} />
              <Button icon="word" label="Rich text" onClick={() => setRich(true)} />
            </div>
            <textarea
              className="rw-input ml-compose-body"
              value={draft.text || ''}
              onChange={(e) => onChange({ ...draft, text: e.target.value })}
              rows={14}
            />
          </Field>
        )}

        {draft.attachments?.length ? (
          <div className="ml-files">
            {draft.attachments.map((a, i) => (
              <span key={`${a.filename}-${i}`} className="ml-attachment">
                <Icon name="attach" size={13} />
                <span className="name">{a.filename}</span>
                <span className="size">{formatBytes(a.size || 0)}</span>
                <button
                  type="button"
                  className="ml-linkchip"
                  title="Remove"
                  onClick={() => onChange({ ...draft, attachments: draft.attachments.filter((_, n) => n !== i) })}
                >
                  <Icon name="close" size={11} />
                </button>
              </span>
            ))}
            <span className="ml-badge">{formatBytes(total)} in total</span>
          </div>
        ) : null}

        {scheduling ? (
          <div className="ml-found">
            {scheduleChoices().map((choice) => (
              <button key={choice.label} type="button" className="ml-found-item" onClick={() => onSend(collect(), choice.at.toISOString())}>
                <span className="ml-found-logo">
                  <Icon name="clock" size={15} />
                </span>
                <span className="grow">
                  <div className="who">{choice.label}</div>
                  <div className="what">{choice.at.toLocaleString()}</div>
                </span>
                <Icon name="chevronRight" size={14} />
              </button>
            ))}
            <p className="rw-hint" style={{ margin: 0 }}>
              A scheduled message waits on this computer, so it goes out when Rutba Office is running.
            </p>
          </div>
        ) : null}
      </div>
    </Dialog>
  );
}

/** Plain text into the markup the editor expects, quoting preserved. */
function textToHtml(text) {
  return String(text)
    .split('\n')
    .map((line) => {
      const safe = line.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
      return line.startsWith('>') ? `<blockquote>${safe.replace(/^&gt;\s?/, '')}</blockquote>` : `<div>${safe || '<br>'}</div>`;
    })
    .join('');
}
