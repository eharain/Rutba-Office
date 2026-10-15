// Review → Comments: the pane of a slide's comment threads, the composer
// for a new one, replies, Resolve and Delete — and the markers on the stage
// that say where each thread is anchored, as PowerPoint 365 draws them.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, t, tn } from '@rutba/office-ui';
import { unreadThreads, markRead } from '@rutba/presentation/changes';

/** A person's colour: the same name, the same colour, every time. */
export function personColour(name) {
  const palette = ['#2b5fd9', '#0f9d58', '#c2408f', '#e08b2b', '#7b5cd6', '#1b998b', '#d7263d', '#3b8ea5'];
  let h = 0;
  for (const ch of String(name || '')) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
  return palette[h % palette.length];
}

/** When, as the pane says it: minutes and hours ago today, else the date. */
export function whenOf(created) {
  if (!created) return '';
  const d = new Date(/Z|[+-]\d\d:?\d\d$/.test(created) ? created : `${created}Z`);
  if (Number.isNaN(d.getTime())) return '';
  const mins = Math.round((Date.now() - d.getTime()) / 60000);
  if (mins < 1) return t('Just now');
  if (mins < 60) return tn(mins, '{count} min ago', '{count} min ago');
  if (mins < 60 * 12) return tn(Math.round(mins / 60), '{count} h ago', '{count} h ago');
  return d.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: d.getFullYear() === new Date().getFullYear() ? undefined : 'numeric' });
}

const Avatar = ({ name, initials, size = 26 }) => (
  <span className="sl-cm-avatar" style={{ width: size, height: size, background: personColour(name), fontSize: size * 0.4 }} aria-hidden="true">{initials || '?'}</span>
);

/**
 * A comment's words, its mentions shown as PowerPoint shows them: @ and
 * the name, in the accent colour. A mention is the name of someone in the
 * deck's comments after an @, or else the one word after it.
 */
export function withMentions(text, people = []) {
  const names = [...people].filter(Boolean).sort((a, b) => b.length - a.length).map((n) => n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  const re = new RegExp('@(?:' + (names.length ? names.join('|') + '|' : '') + "[\\p{L}][\\p{L}\\p{M}'-]*)", 'gu');
  const out = [];
  let at = 0;
  for (const m of String(text || '').matchAll(re)) {
    if (m.index > at) out.push(text.slice(at, m.index));
    out.push(<span key={m.index} className="sl-cm-mention">{m[0]}</span>);
    at = m.index + m[0].length;
  }
  if (at < String(text || '').length) out.push(String(text).slice(at));
  return out;
}

/**
 * A box that posts on Ctrl+Enter, the pane's composer and every reply box.
 * Typing @ offers the people in the deck's comments, as PowerPoint offers
 * the people it knows; the name goes in as words.
 */
function Composer({ placeholder, onPost, onCancel = null, autoFocus = false, className = '', label = t('Post'), initial = '', people = [] }) {
  const [text, setText] = useState(initial);
  const [caret, setCaret] = useState(initial.length);
  const ref = useRef(null);
  useEffect(() => { if (autoFocus) ref.current?.focus(); }, [autoFocus]);
  const post = () => {
    const t = text.trim();
    if (!t) return;
    onPost(t);
    setText('');
  };
  const typed = /(?:^|\s)@([\p{L}\p{M} '-]{0,30})$/u.exec(text.slice(0, caret));
  const offers = typed ? people.filter((p) => p.toLowerCase().startsWith(typed[1].toLowerCase()) && p.toLowerCase() !== typed[1].toLowerCase()).slice(0, 5) : [];
  const mention = (name) => {
    const start = caret - typed[1].length;
    const next = text.slice(0, start) + name + ' ' + text.slice(caret);
    setText(next);
    const at = start + name.length + 1;
    setCaret(at);
    requestAnimationFrame(() => { ref.current?.focus(); ref.current?.setSelectionRange(at, at); });
  };
  return (
    <div className={`sl-cm-compose ${className}`}>
      <textarea
        ref={ref}
        className="rw-input sl-cm-input"
        rows={2}
        placeholder={placeholder}
        value={text}
        onChange={(e) => { setText(e.target.value); setCaret(e.target.selectionStart ?? e.target.value.length); }}
        onSelect={(e) => setCaret(e.target.selectionStart ?? 0)}
        onKeyDown={(e) => {
          if (offers.length && (e.key === 'Tab' || (e.key === 'Enter' && !e.ctrlKey && !e.metaKey))) { e.preventDefault(); mention(offers[0]); return; }
          if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); post(); }
          if (e.key === 'Escape' && onCancel) { e.preventDefault(); onCancel(); }
        }}
      />
      {offers.length ? (
        <div className="sl-cm-people" role="listbox" aria-label={t('Mention someone')}>
          {offers.map((p) => (
            <button key={p} type="button" role="option" className="sl-cm-person" data-person={p} onMouseDown={(e) => e.preventDefault()} onClick={() => mention(p)}>
              <span className="sl-cm-avatar" style={{ width: 20, height: 20, background: personColour(p), fontSize: 8 }} aria-hidden="true">{initialsFrom(p)}</span>
              <span>{p}</span>
            </button>
          ))}
        </div>
      ) : null}
      <div className="sl-cm-compose-actions">
        {onCancel ? <Button label={t('Cancel')} className="sl-cm-cancel" onClick={onCancel} /> : null}
        <Button primary label={label} className="sl-cm-post" disabled={!text.trim()} title={`${label} (Ctrl+Enter)`} onClick={post} />
      </div>
    </div>
  );
}

/**
 * The Comments pane: this slide's threads, the draft of a new one at the
 * top while it is being written, each thread with its replies, a reply
 * box, Resolve (or Reopen) and Delete. An older-format thread says so and
 * offers Delete only, as PowerPoint 365 does.
 */
export function CommentsPane({ threads, draft, selected, shapes, me, unread = null, onSelect, onPost, onCancelDraft, onReply, onResolve, onDelete, onDeleteReply, onNew, onEdit = null, people = [] }) {
  const shapeName = (id) => shapes.find((s) => String(s.id) === String(id))?.name || t('a shape');
  // Edit comment: the thread's own words or a reply's, one's own only, as PowerPoint 365 allows.
  const [editing, setEditing] = useState(null);
  const mine = (name) => Boolean(onEdit) && String(name || '') === String(me || '');
  const editBox = (th, reply, text) => (
    <Composer autoFocus initial={text} people={people} label={t('Save')} className="sl-cm-editbox" onCancel={() => setEditing(null)} onPost={(words) => { setEditing(null); onEdit(th, reply, words); }} />
  );
  const listRef = useRef(null);
  const fresh = unread ? threads.filter((th) => unread.has(String(th.id))).length : 0;
  useEffect(() => {
    if (!selected || !listRef.current) return;
    listRef.current.querySelector(`[data-comment="${CSS.escape(selected)}"]`)?.scrollIntoView({ block: 'nearest' });
  }, [selected]);
  return (
    <div className="sl-cm-pane">
      <div className="sl-cm-tools">
        <Button icon="plus" label={t('New')} title={t('New comment — on the selected shape, or on the slide')} className="sl-cm-new" onClick={onNew} />
        <span className="sl-cm-count">{threads.length ? tn(threads.length, '{count} on this slide', '{count} on this slide') : ''}{threads.length && fresh ? ` · ${tn(fresh, '{count} unread', '{count} unread')}` : ''}</span>
      </div>
      <div className="sl-cm-list" ref={listRef}>
        {draft ? (
          <div className="sl-cm-card sl-cm-draft">
            <div className="sl-cm-head">
              <Avatar name={me} initials={initialsFrom(me)} />
              <div className="sl-cm-who"><span className="sl-cm-name">{me}</span><span className="sl-cm-anchor">{draft.shape != null ? t('On {shape}', { shape: shapeName(draft.shape) }) : t('On this slide')}</span></div>
            </div>
            <Composer autoFocus placeholder={t('Type your comment… @ mentions someone')} onPost={onPost} onCancel={onCancelDraft} className="sl-cm-draftbox" people={people} />
          </div>
        ) : null}
        {!threads.length && !draft ? (
          <div className="sl-cm-empty">
            <Icon name="reply" size={22} />
            <div>{t('No comments on this slide.')}</div>
            <Button label={t('New Comment')} className="sl-cm-new-empty" onClick={onNew} />
          </div>
        ) : null}
        {threads.map((th) => (
          <div
            key={th.id}
            className={`sl-cm-card${selected === th.id ? ' selected' : ''}${th.status === 'resolved' ? ' resolved' : ''}${th.legacy ? ' legacy' : ''}${unread?.has(String(th.id)) ? ' unread' : ''}`}
            data-comment={th.id}
            onMouseDown={() => onSelect(th.id)}
          >
            <div className="sl-cm-head">
              <Avatar name={th.author} initials={th.initials} />
              <div className="sl-cm-who">
                <span className="sl-cm-name">{unread?.has(String(th.id)) ? <span className="sl-cm-dot" role="img" aria-label={t('Unread')} data-tip={t('Unread: opened, it is marked read')} /> : null}{th.author}</span>
                <span className="sl-cm-anchor">{whenOf(th.created)}{th.shapeId != null ? ` · ${t('on {shape}', { shape: shapeName(th.shapeId) })}` : ''}</span>
              </div>
              <div className="sl-cm-actions">
                {!th.legacy && mine(th.author) && th.status !== 'resolved' ? (
                  <button type="button" className="sl-cm-act sl-cm-edit" data-tip={t('Edit comment')} aria-label={t('Edit comment')} onClick={(e) => { e.stopPropagation(); setEditing({ id: th.id, reply: null }); }}>
                    <Icon name="pen" size={14} />
                  </button>
                ) : null}
                {!th.legacy ? (
                  <button type="button" className="sl-cm-act sl-cm-resolve" data-tip={th.status === 'resolved' ? t('Reopen this thread') : t('Resolve this thread')} aria-label={th.status === 'resolved' ? t('Reopen this thread') : t('Resolve this thread')} onClick={(e) => { e.stopPropagation(); onResolve(th, th.status !== 'resolved'); }}>
                    <Icon name={th.status === 'resolved' ? 'refresh' : 'check'} size={14} />
                  </button>
                ) : null}
                <button type="button" className="sl-cm-act sl-cm-delete" data-tip={t('Delete this thread')} aria-label={t('Delete this thread')} onClick={(e) => { e.stopPropagation(); onDelete(th); }}>
                  <Icon name="trash" size={14} />
                </button>
              </div>
            </div>
            {th.status === 'resolved' ? <div className="sl-cm-badge">{t('Resolved')}</div> : null}
            {editing?.id === th.id && !editing.reply ? editBox(th, null, th.text) : <div className="sl-cm-text">{withMentions(th.text, people)}</div>}
            {th.replies.length ? (
              <div className="sl-cm-replies">
                {th.replies.map((r) => (
                  <div key={r.id} className="sl-cm-reply" data-reply={r.id}>
                    <div className="sl-cm-head">
                      <Avatar name={r.author} initials={r.initials} size={22} />
                      <div className="sl-cm-who"><span className="sl-cm-name">{r.author}</span><span className="sl-cm-anchor">{whenOf(r.created)}</span></div>
                      {!th.legacy ? (
                        <div className="sl-cm-actions">
                          {mine(r.author) && th.status !== 'resolved' ? <button type="button" className="sl-cm-act sl-cm-edit-reply" data-tip={t('Edit reply')} aria-label={t('Edit reply')} onClick={(e) => { e.stopPropagation(); setEditing({ id: th.id, reply: r.id }); }}><Icon name="pen" size={13} /></button> : null}
                          <button type="button" className="sl-cm-act sl-cm-delete-reply" data-tip={t('Delete this reply')} aria-label={t('Delete this reply')} onClick={(e) => { e.stopPropagation(); onDeleteReply(th, r); }}><Icon name="trash" size={13} /></button>
                        </div>
                      ) : null}
                    </div>
                    {editing?.id === th.id && editing.reply === r.id ? editBox(th, r, r.text) : <div className="sl-cm-text">{withMentions(r.text, people)}</div>}
                  </div>
                ))}
              </div>
            ) : null}
            {th.legacy ? (
              <div className="sl-cm-note">{t('Made in an older version of PowerPoint: it can be read and deleted, not replied to.')}</div>
            ) : th.status === 'resolved' ? null : selected === th.id ? (
              <Composer placeholder={t('Reply… @ mentions someone')} label={t('Reply')} className="sl-cm-replybox" onPost={(text) => onReply(th, text)} people={people} />
            ) : null}
          </div>
        ))}
      </div>
    </div>
  );
}

/**
 * Review → Mark All as Read: which threads hold words not yet read on this
 * computer. What was read is kept in the store by the deck's path (the
 * hundred decks read most lately), so it is still known next time; a deck
 * never saved keeps it while it is open. Nothing counts as unread until
 * the list is loaded, so a thread is never marked read over a list not
 * yet read back.
 * @returns {{ unread: Set<string>, mark: (threads) => number }} mark returns how many threads it read
 */
export function useCommentsRead(shell, doc, comments, me) {
  const [read, setRead] = useState(null);
  const at = useRef({ id: null, path: null, read: null });
  const keep = useCallback(async (path, list) => {
    if (!path || !shell?.store) return;
    const all = (await shell.store.get({ key: 'slides.commentsRead', fallback: {} }).catch(() => ({}))) || {};
    all[String(path).toLowerCase()] = { ids: list, at: Date.now() };
    const kept = Object.entries(all).sort((a, b) => (b[1]?.at || 0) - (a[1]?.at || 0)).slice(0, 100);
    await shell.store.set({ key: 'slides.commentsRead', value: Object.fromEntries(kept) }).catch(() => {});
  }, [shell]);
  useEffect(() => {
    const id = doc?.id || null;
    const path = doc?.path || null;
    const was = at.current;
    at.current = { ...was, id, path };
    // The same deck saved under a new name keeps what was read in it.
    if (id && was.id === id && was.read) {
      if (path && path !== was.path) keep(path, was.read);
      return;
    }
    at.current.read = null;
    setRead(null);
    if (!id) return;
    if (!path || !shell?.store) { at.current.read = []; setRead([]); return; }
    let live = true;
    shell.store.get({ key: 'slides.commentsRead', fallback: {} }).catch(() => ({})).then((all) => {
      if (!live) return;
      const list = Array.isArray(all?.[String(path).toLowerCase()]?.ids) ? all[String(path).toLowerCase()].ids : [];
      at.current.read = list;
      setRead(list);
    });
    return () => { live = false; };
  }, [shell, doc?.id, doc?.path, keep]);
  const unread = useMemo(() => (read ? unreadThreads(comments || [], read, me) : new Set()), [comments, read, me]);
  const mark = useCallback((threads) => {
    const cur = at.current.read;
    if (!cur) return 0;
    const before = unreadThreads(comments || [], cur, me);
    const next = markRead(cur, threads, comments || []);
    const after = unreadThreads(comments || [], next, me);
    at.current.read = next;
    setRead(next);
    keep(at.current.path, next);
    return before.size - after.size;
  }, [comments, me, keep]);
  return { unread, mark };
}

export function initialsFrom(name) {
  const words = String(name || '').trim().split(/\s+/).filter(Boolean);
  if (!words.length) return '?';
  return (words.length === 1 ? words[0].slice(0, 2) : words[0][0] + words[words.length - 1][0]).toUpperCase();
}

/**
 * Where a thread's marker sits on the slide: at the top right of the
 * shape it is anchored to, at its own point, or down the slide's top left
 * corner for one on the slide as a whole.
 */
export function markerSpots(threads, shapes) {
  let loose = 0;
  return threads.map((t) => {
    const g = t.shapeId != null ? shapes.find((s) => String(s.id) === String(t.shapeId))?.geometry : null;
    if (g) return { thread: t, x: g.x + g.w - 6, y: g.y - 14 };
    if (t.pos) return { thread: t, x: t.pos.x, y: t.pos.y };
    const at = loose++;
    return { thread: t, x: 10, y: 10 + at * 34 };
  });
}

export const COMMENTS_CSS = `
.sl-cm-mention { color: var(--accent); font-weight: 600; }
.sl-cm-compose { position: relative; }
.sl-cm-people { position: absolute; left: 0; right: 0; top: calc(100% - 34px); z-index: 5; background: var(--surface); border: 1px solid var(--line); border-radius: 8px; box-shadow: 0 6px 18px rgba(15,23,42,.16); padding: 4px; }
.sl-cm-person { display: flex; align-items: center; gap: 8px; width: 100%; border: 0; background: transparent; padding: 5px 8px; border-radius: 5px; font: inherit; font-size: 12.5px; color: var(--ink); cursor: pointer; text-align: start; }
.sl-cm-person:hover, .sl-cm-person:first-child { background: var(--hover); }
/* Review → Comments: the pane. */
.sl-cm-pane { display: flex; flex-direction: column; min-height: 0; height: 100%; }
.sl-cm-tools { display: flex; align-items: center; gap: 8px; padding: 6px 10px; border-bottom: 1px solid var(--line-soft); }
.sl-cm-count { font-size: 11.5px; color: var(--ink-3); margin-left: auto; }
.sl-cm-list { flex: 1; min-height: 0; overflow: auto; padding: 10px; display: flex; flex-direction: column; gap: 10px; }
.sl-cm-card { border: 1px solid var(--line); border-radius: 10px; background: var(--surface); padding: 10px 11px; box-shadow: 0 1px 2px rgba(15,20,30,.05); transition: border-color var(--fast), box-shadow var(--fast); cursor: default; }
.sl-cm-card:hover { border-color: var(--line-strong); }
.sl-cm-card.selected { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }
.sl-cm-card.resolved { background: var(--surface-2); }
.sl-cm-card.resolved .sl-cm-text, .sl-cm-card.resolved .sl-cm-name { color: var(--ink-3); }
.sl-cm-draft { border-color: var(--accent); }
.sl-cm-head { display: flex; align-items: center; gap: 8px; }
.sl-cm-avatar { flex: none; border-radius: 50%; color: #fff; display: grid; place-items: center; font-weight: 600; letter-spacing: .02em; }
.sl-cm-who { display: flex; flex-direction: column; min-width: 0; flex: 1; line-height: 1.25; }
.sl-cm-name { font-size: 12.5px; font-weight: 600; color: var(--ink); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-cm-anchor { font-size: 11px; color: var(--ink-3); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sl-cm-actions { display: flex; gap: 2px; flex: none; opacity: .55; transition: opacity var(--fast); }
.sl-cm-card:hover .sl-cm-actions, .sl-cm-card.selected .sl-cm-actions { opacity: 1; }
.sl-cm-act { border: 0; background: transparent; color: var(--ink-2); width: 24px; height: 24px; border-radius: 6px; display: grid; place-items: center; cursor: pointer; }
.sl-cm-act:hover { background: var(--hover); color: var(--accent); }
.sl-cm-text { font-size: 12.5px; color: var(--ink); margin: 7px 0 0 34px; white-space: pre-wrap; word-break: break-word; line-height: 1.45; }
.sl-cm-badge { display: inline-block; margin: 6px 0 0 34px; padding: 1px 7px; border-radius: 999px; background: color-mix(in srgb, #0f9d58 14%, transparent); color: #0b7a44; font-size: 10.5px; font-weight: 600; }
.sl-cm-replies { margin: 8px 0 0 34px; display: flex; flex-direction: column; gap: 8px; border-left: 2px solid var(--line-soft); padding-left: 10px; }
.sl-cm-reply .sl-cm-text { margin-left: 30px; margin-top: 4px; }
.sl-cm-note { margin: 8px 0 0 34px; font-size: 11px; color: var(--ink-3); font-style: italic; }
.sl-cm-compose { margin-top: 8px; display: flex; flex-direction: column; gap: 6px; }
.sl-cm-card .sl-cm-compose { margin-left: 34px; }
.sl-cm-input { width: 100%; box-sizing: border-box; resize: vertical; min-height: 44px; font: inherit; font-size: 12.5px; padding: 6px 8px; line-height: 1.4; }
.sl-cm-compose-actions { display: flex; justify-content: flex-end; gap: 6px; }
.sl-cm-compose-actions .rw-btn { height: 26px; padding: 0 10px; font-size: 12px; }
.sl-cm-empty { display: flex; flex-direction: column; align-items: center; gap: 8px; padding: 28px 10px; color: var(--ink-3); font-size: 12.5px; text-align: center; }
/* The markers on the stage: a speech bubble with the author's initials, at the thread's anchor. */
.sl-cm-marker { position: absolute; z-index: 8; display: grid; place-items: center; padding: 0; border: 2px solid #fff; border-radius: 50% 50% 50% 4px; color: #fff; font-weight: 700; cursor: pointer; box-shadow: 0 2px 6px rgba(15,20,30,.28); line-height: 1; }
.sl-cm-marker.resolved { opacity: .45; }
.sl-cm-marker.current { outline: 2px solid var(--accent); outline-offset: 1px; }
/* Unread: a dot on the marker's free corner, and on the card a stripe and a dot before the name. */
.sl-cm-marker.unread::after { content: ''; position: absolute; left: -.35em; top: -.35em; width: .9em; height: .9em; border-radius: 50%; background: var(--accent); border: .2em solid #fff; }
.sl-cm-card.unread { box-shadow: inset 3px 0 0 var(--accent); }
.sl-cm-card.unread.selected { box-shadow: inset 3px 0 0 var(--accent), 0 0 0 2px var(--accent-soft); }
.sl-cm-dot { display: inline-block; width: 7px; height: 7px; border-radius: 50%; background: var(--accent); margin-inline-end: 6px; vertical-align: 1px; }
.sl-cm-marker .sl-cm-more { position: absolute; right: -6px; top: -6px; min-width: 14px; height: 14px; border-radius: 7px; background: var(--ink); color: #fff; font-size: 9px; display: grid; place-items: center; padding: 0 3px; }
/* The strip: how many threads a slide has. */
.sl-thumb-cm { display: grid; place-items: center; color: var(--accent); }
.sl-thumb.active .sl-thumb-cm { color: var(--accent); }
`;
