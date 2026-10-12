// Draw: the marks made on the picture with the mouse, and the one being made.
//
// A mark is drawn in the result's pixels and stays editable (moved, its
// corners or ends dragged, its colour and width changed) until it is
// committed (Done, Enter, another mark, another tool); then it becomes an
// operation kept in the source's pixels (pipeline.js markToOperation) and
// takes its place in the edit list, where Undo takes it back.
import React, { useEffect, useRef } from 'react';
import { t } from '@rutba/office-ui';
import { MARK_FONT } from './draw.js';

export const TOOLS = ['pen', 'rect', 'ellipse', 'arrow', 'text'];

/** The box round a mark, in the result's pixels. */
export function markBox(m) {
  if (m.shape === 'pen') {
    const xs = m.points.map((p) => p.x);
    const ys = m.points.map((p) => p.y);
    return { x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) };
  }
  if (m.shape === 'text') {
    const lines = String(m.text || ' ').split('\n');
    const longest = Math.max(1, ...lines.map((l) => l.length));
    return { x: m.x1, y: m.y1, width: Math.max(m.size * 2, longest * m.size * 0.58), height: lines.length * m.size * 1.25 };
  }
  return { x: Math.min(m.x1, m.x2), y: Math.min(m.y1, m.y2), width: Math.abs(m.x2 - m.x1), height: Math.abs(m.y2 - m.y1) };
}

/** Is the mark big enough to keep? A click with no drag makes nothing. */
export function markIsSomething(m) {
  if (!m) return false;
  if (m.shape === 'text') return String(m.text || '').trim().length > 0;
  if (m.shape === 'pen') return m.points.length > 1;
  return Math.hypot(m.x2 - m.x1, m.y2 - m.y1) >= 3;
}

const moved = (m, dx, dy) => (m.shape === 'pen'
  ? { ...m, points: m.points.map((p) => ({ x: p.x + dx, y: p.y + dy })) }
  : { ...m, x1: m.x1 + dx, y1: m.y1 + dy, x2: (m.x2 ?? m.x1) + dx, y2: (m.y2 ?? m.y1) + dy });

/**
 * The layer over the picture while a Draw tool or Red eye is chosen: it
 * takes the mouse, makes the mark, and draws the frame and handles of the
 * one being edited.
 */
export function DrawLayer({ tool, draft, onDraft, onCommit, scale, style, onSpot, spotRadius }) {
  const layer = useRef(null);
  const ring = useRef(null);
  const text = useRef(null);
  const at = (e) => {
    const r = layer.current.getBoundingClientRect();
    return { x: (e.clientX - r.left) / scale, y: (e.clientY - r.top) / scale };
  };

  useEffect(() => {
    if (draft?.shape === 'text') text.current?.focus({ preventScroll: true });
  }, [draft?.shape, draft?.id]);

  const follow = (step, end) => {
    const move = (ev) => step(ev);
    const up = (ev) => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      end?.(ev);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const down = (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    const p = at(e);
    if (tool === 'redeye') {
      onSpot({ x: p.x, y: p.y, radius: spotRadius / scale });
      return;
    }
    if (draft) onCommit();
    const id = Date.now();
    if (tool === 'text') {
      onDraft({ id, shape: 'text', x1: p.x, y1: p.y, text: '', colour: style.colour, size: style.size, width: style.width });
      return;
    }
    let mark = tool === 'pen'
      ? { id, shape: 'pen', points: [p], colour: style.colour, width: style.width }
      : { id, shape: tool, x1: p.x, y1: p.y, x2: p.x, y2: p.y, colour: style.colour, width: style.width };
    onDraft(mark);
    follow((ev) => {
      const q = at(ev);
      if (mark.shape === 'pen') {
        const last = mark.points[mark.points.length - 1];
        if (Math.hypot(q.x - last.x, q.y - last.y) * scale < 1.5) return;
        mark = { ...mark, points: [...mark.points, q] };
      } else if (ev.shiftKey && (mark.shape === 'rect' || mark.shape === 'ellipse')) {
        // Shift keeps it square or round, as in Paint.
        const side = Math.max(Math.abs(q.x - mark.x1), Math.abs(q.y - mark.y1));
        mark = { ...mark, x2: mark.x1 + Math.sign(q.x - mark.x1 || 1) * side, y2: mark.y1 + Math.sign(q.y - mark.y1 || 1) * side };
      } else {
        mark = { ...mark, x2: q.x, y2: q.y };
      }
      onDraft(mark);
    }, () => {
      if (!markIsSomething(mark)) onDraft(null);
    });
  };

  const hover = (e) => {
    if (tool !== 'redeye' || !ring.current) return;
    const r = layer.current.getBoundingClientRect();
    ring.current.style.transform = `translate(${e.clientX - r.left - spotRadius}px, ${e.clientY - r.top - spotRadius}px)`;
    ring.current.style.opacity = '1';
  };

  // The frame of the mark being edited: drag it to move, its handles to reshape.
  const grab = (e, handle) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const start = at(e);
    const from = draft;
    follow((ev) => {
      const q = at(ev);
      const dx = q.x - start.x;
      const dy = q.y - start.y;
      if (handle === 'move') onDraft(moved(from, dx, dy));
      else if (handle === 'a') onDraft({ ...from, x1: from.x1 + dx, y1: from.y1 + dy });
      else if (handle === 'b') onDraft({ ...from, x2: from.x2 + dx, y2: from.y2 + dy });
      else {
        const box = { x1: Math.min(from.x1, from.x2), y1: Math.min(from.y1, from.y2), x2: Math.max(from.x1, from.x2), y2: Math.max(from.y1, from.y2) };
        if (handle.includes('w')) box.x1 += dx;
        if (handle.includes('e')) box.x2 += dx;
        if (handle.includes('n')) box.y1 += dy;
        if (handle.includes('s')) box.y2 += dy;
        onDraft({ ...from, ...box });
      }
    });
  };

  const px = (v) => `${v * scale}px`;
  let frame = null;
  if (draft && draft.shape !== 'text') {
    const b = markBox(draft);
    const pad = 6 / scale + (draft.width || 0) / 2;
    const handles = draft.shape === 'arrow'
      ? [['a', draft.x1, draft.y1], ['b', draft.x2, draft.y2]]
      : draft.shape === 'pen'
        ? []
        : [['nw', b.x, b.y], ['ne', b.x + b.width, b.y], ['sw', b.x, b.y + b.height], ['se', b.x + b.width, b.y + b.height]];
    frame = (
      <>
        <div
          className="im-draft"
          style={{ left: px(b.x - pad), top: px(b.y - pad), width: px(b.width + pad * 2), height: px(b.height + pad * 2) }}
          onMouseDown={(e) => grab(e, 'move')}
          data-shape={draft.shape}
        />
        {handles.map(([h, x, y]) => (
          <span key={h} className={`im-handle h-${h}`} style={{ left: px(x), top: px(y) }} onMouseDown={(e) => grab(e, h)} data-handle={h} />
        ))}
      </>
    );
  } else if (draft?.shape === 'text') {
    frame = (
      <div className="im-draft text" style={{ left: px(draft.x1), top: px(draft.y1) }} data-shape="text">
        <span className="im-text-grip" onMouseDown={(e) => grab(e, 'move')} data-tip={t('Move the text')} />
        <textarea
          ref={text}
          className="im-text-input"
          value={draft.text}
          rows={Math.max(1, String(draft.text).split('\n').length)}
          placeholder={t('Type here')}
          spellCheck={false}
          style={{
            color: draft.colour,
            font: `600 ${draft.size * scale}px ${MARK_FONT}`,
            lineHeight: 1.25,
            width: `${Math.max(4, ...String(draft.text || '').split('\n').map((l) => l.length + 1)) * draft.size * scale * 0.6}px`,
          }}
          onChange={(e) => onDraft({ ...draft, text: e.target.value })}
          onKeyDown={(e) => {
            if (e.key === 'Escape') { e.preventDefault(); onDraft(null); }
            else if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); onCommit(); }
          }}
          onMouseDown={(e) => e.stopPropagation()}
        />
      </div>
    );
  }

  return (
    <div
      ref={layer}
      className={`im-draw tool-${tool}`}
      onMouseDown={down}
      onMouseMove={hover}
      onMouseLeave={() => { if (ring.current) ring.current.style.opacity = '0'; }}
    >
      {tool === 'redeye' ? <span ref={ring} className="im-ring" style={{ width: spotRadius * 2, height: spotRadius * 2 }} /> : null}
      {frame}
    </div>
  );
}
