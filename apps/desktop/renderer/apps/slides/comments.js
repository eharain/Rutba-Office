// Review → Comments: the pane of a slide's comment threads, the composer
// for a new one, replies, Resolve and Delete — and the markers on the stage
// that say where each thread is anchored, as PowerPoint 365 draws them.

import React, { useEffect, useRef, useState } from 'react';
import { Button, Icon } from '@rutba/office-ui';

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
  if (mins < 1) return 'Just now';
  if (mins < 60) return `${mins} min ago`;
  if (mins < 60 * 12) return `${Math.round(mins / 60)} h ago`;
  return d.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: d.getFullYear() === new Date().getFullYear() ? undefined : 'numeric' });
}

const Avatar = ({ name, initials, size = 26 }) => (
  <span className="sl-cm-avatar" style={{ width: size, height: size, background: personColour(name), fontSize: size * 0.4 }} aria-hidden="true">{initials || '?'}</span>
);

/** A box that posts on Ctrl+Enter, the pane's composer and every reply box. */
function Composer({ placeholder, onPost, onCancel = null, autoFocus = false, className = '', label = 'Post' }) {
  const [text, setText] = useState('');
  const ref = useRef(null);
  useEffect(() => { if (autoFocus) ref.current?.focus(); }, [autoFocus]);
  const post = () => {
    const t = text.trim();
    if (!t) return;
    onPost(t);
    setText('');
  };
  return (
    <div className={`sl-cm-compose ${className}`}>
      <textarea
        ref={ref}
        className="rw-input sl-cm-input"
        rows={2}
        placeholder={placeholder}
        value={text}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); post(); }
          if (e.key === 'Escape' && onCancel) { e.preventDefault(); onCancel(); }
        }}
      />
      <div className="sl-cm-compose-actions">
        {onCancel ? <Button label="Cancel" className="sl-cm-cancel" onClick={onCancel} /> : null}
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
export function CommentsPane({ threads, draft, selected, shapes, me, onSelect, onPost, onCancelDraft, onReply, onResolve, onDelete, onDeleteReply, onNew }) {
  const shapeName = (id) => shapes.find((s) => String(s.id) === String(id))?.name || 'a shape';
  const listRef = useRef(null);
  useEffect(() => {
    if (!selected || !listRef.current) return;
    listRef.current.querySelector(`[data-comment="${CSS.escape(selected)}"]`)?.scrollIntoView({ block: 'nearest' });
  }, [selected]);
  return (
    <div className="sl-cm-pane">
      <div className="sl-cm-tools">
        <Button icon="plus" label="New" title="New comment — on the selected shape, or on the slide" className="sl-cm-new" onClick={onNew} />
        <span className="sl-cm-count">{threads.length ? `${threads.length} on this slide` : ''}</span>
      </div>
      <div className="sl-cm-list" ref={listRef}>
        {draft ? (
          <div className="sl-cm-card sl-cm-draft">
            <div className="sl-cm-head">
              <Avatar name={me} initials={initialsFrom(me)} />
              <div className="sl-cm-who"><span className="sl-cm-name">{me}</span><span className="sl-cm-anchor">{draft.shape != null ? `On ${shapeName(draft.shape)}` : 'On this slide'}</span></div>
            </div>
            <Composer autoFocus placeholder="Type your comment…" onPost={onPost} onCancel={onCancelDraft} className="sl-cm-draftbox" />
          </div>
        ) : null}
        {!threads.length && !draft ? (
          <div className="sl-cm-empty">
            <Icon name="reply" size={22} />
            <div>No comments on this slide.</div>
            <Button label="New Comment" className="sl-cm-new-empty" onClick={onNew} />
          </div>
        ) : null}
        {threads.map((t) => (
          <div
            key={t.id}
            className={`sl-cm-card${selected === t.id ? ' selected' : ''}${t.status === 'resolved' ? ' resolved' : ''}${t.legacy ? ' legacy' : ''}`}
            data-comment={t.id}
            onMouseDown={() => onSelect(t.id)}
          >
            <div className="sl-cm-head">
              <Avatar name={t.author} initials={t.initials} />
              <div className="sl-cm-who">
                <span className="sl-cm-name">{t.author}</span>
                <span className="sl-cm-anchor">{whenOf(t.created)}{t.shapeId != null ? ` · on ${shapeName(t.shapeId)}` : ''}</span>
              </div>
              <div className="sl-cm-actions">
                {!t.legacy ? (
                  <button type="button" className="sl-cm-act sl-cm-resolve" data-tip={t.status === 'resolved' ? 'Reopen this thread' : 'Resolve this thread'} onClick={(e) => { e.stopPropagation(); onResolve(t, t.status !== 'resolved'); }}>
                    <Icon name={t.status === 'resolved' ? 'refresh' : 'check'} size={14} />
                  </button>
                ) : null}
                <button type="button" className="sl-cm-act sl-cm-delete" data-tip="Delete this thread" onClick={(e) => { e.stopPropagation(); onDelete(t); }}>
                  <Icon name="trash" size={14} />
                </button>
              </div>
            </div>
            {t.status === 'resolved' ? <div className="sl-cm-badge">Resolved</div> : null}
            <div className="sl-cm-text">{t.text}</div>
            {t.replies.length ? (
              <div className="sl-cm-replies">
                {t.replies.map((r) => (
                  <div key={r.id} className="sl-cm-reply" data-reply={r.id}>
                    <div className="sl-cm-head">
                      <Avatar name={r.author} initials={r.initials} size={22} />
                      <div className="sl-cm-who"><span className="sl-cm-name">{r.author}</span><span className="sl-cm-anchor">{whenOf(r.created)}</span></div>
                      {!t.legacy ? (
                        <div className="sl-cm-actions">
                          <button type="button" className="sl-cm-act sl-cm-delete-reply" data-tip="Delete this reply" onClick={(e) => { e.stopPropagation(); onDeleteReply(t, r); }}><Icon name="trash" size={13} /></button>
                        </div>
                      ) : null}
                    </div>
                    <div className="sl-cm-text">{r.text}</div>
                  </div>
                ))}
              </div>
            ) : null}
            {t.legacy ? (
              <div className="sl-cm-note">Made in an older version of PowerPoint: it can be read and deleted, not replied to.</div>
            ) : t.status === 'resolved' ? null : selected === t.id ? (
              <Composer placeholder="Reply…" label="Reply" className="sl-cm-replybox" onPost={(text) => onReply(t, text)} />
            ) : null}
          </div>
        ))}
      </div>
    </div>
  );
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
.sl-cm-marker .sl-cm-more { position: absolute; right: -6px; top: -6px; min-width: 14px; height: 14px; border-radius: 7px; background: var(--ink); color: #fff; font-size: 9px; display: grid; place-items: center; padding: 0 3px; }
/* The strip: how many threads a slide has. */
.sl-thumb-cm { display: grid; place-items: center; color: var(--accent); }
.sl-thumb.active .sl-thumb-cm { color: var(--accent); }
`;
