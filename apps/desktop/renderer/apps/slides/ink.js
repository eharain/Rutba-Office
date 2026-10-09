// Draw: a pen surface over the slide — Pen, Pencil and Highlighter strokes
// drawn as the pointer moves and kept as ink shapes ("Ink N", see
// deck.addInk); the Eraser takes away the strokes it passes over; the Lasso
// selects the strokes inside a loop; the Ruler lays a straight edge on the
// slide that a stroke started along it follows; Ink to Shape turns a drawn
// rectangle, oval or triangle into that shape; Ink Replay draws the
// strokes again in the order they were made.

import React, { useEffect, useRef, useState } from 'react';
import { t } from '@rutba/office-ui';

import { DEFAULT_PENS, PEN_COLOURS, PEN_WIDTHS, strokeLook, isInk, alongRuler, recognise, thin, polygonHas } from './ink-geometry.js';

export { DEFAULT_PENS, PEN_COLOURS, PEN_WIDTHS, strokeLook, isInk, alongRuler, recognise };

const PT = 96 / 72;

/**
 * The surface: over the slide while a pen, the Eraser or the Lasso is on.
 * `size` is the slide's in its own pixels; strokes come back in them.
 * `mirrored`: the surface is drawn mirrored (a sheet right to left), so a
 * point is measured from its right edge.
 */
export function InkSurface({ size, tool, pen, ruler, onRuler, touch, shapes, onStroke, onErase, onLasso, hit = null, mirrored = false }) {
  const host = useRef(null);
  const [live, setLive] = useState(null);
  const erased = useRef(new Set());
  const drag = useRef(null);
  const at = (e) => {
    const r = host.current.getBoundingClientRect();
    return [(((mirrored ? r.right - e.clientX : e.clientX - r.left)) * size.width) / r.width, ((e.clientY - r.top) * size.height) / r.height];
  };
  const inkIds = new Set((shapes || []).filter(isInk).map((s) => String(s.id)));

  // Which stroke an element under the pointer belongs to: a slide's ink by its data-shape, or the app's own `hit`.
  const hitOf = hit || ((el) => {
    const g = el.closest?.('[data-shape]');
    const id = g?.getAttribute('data-shape');
    return id && inkIds.has(id) ? { id, node: g } : null;
  });
  const eraseAt = (e) => {
    for (const el of document.elementsFromPoint(e.clientX, e.clientY)) {
      const h = hitOf(el);
      if (h && !erased.current.has(h.id)) {
        erased.current.add(h.id);
        h.node.style.visibility = 'hidden';
      }
    }
  };

  const down = (e) => {
    if (e.button !== 0 && e.pointerType === 'mouse') return;
    if (e.pointerType === 'touch' && !touch) return;
    e.preventDefault();
    e.stopPropagation();
    host.current.setPointerCapture?.(e.pointerId);
    if (tool === 'eraser') { erased.current = new Set(); eraseAt(e); drag.current = { kind: 'erase' }; return; }
    drag.current = { kind: tool === 'lasso' ? 'lasso' : 'pen', points: [at(e)] };
    setLive(drag.current.points.slice());
  };
  const move = (e) => {
    const d = drag.current;
    if (!d) return;
    e.preventDefault();
    if (d.kind === 'erase') { eraseAt(e); return; }
    const events = e.nativeEvent?.getCoalescedEvents?.() || [e];
    for (const ev of events) d.points.push(at(ev));
    setLive(d.points.slice());
  };
  const up = () => {
    const d = drag.current;
    drag.current = null;
    setLive(null);
    if (!d) return;
    if (d.kind === 'erase') { if (erased.current.size) onErase([...erased.current]); return; }
    if (d.kind === 'lasso') {
      const ids = (shapes || []).filter((s) => isInk(s) && s.geometry && polygonHas(d.points, [s.geometry.x + s.geometry.w / 2, s.geometry.y + s.geometry.h / 2])).map((s) => s.id);
      onLasso(ids);
      return;
    }
    const straight = alongRuler(d.points, ruler);
    onStroke(straight || thin(d.points), Boolean(straight));
  };

  const look = pen ? strokeLook(pen) : null;
  return (
    <div
      ref={host}
      className={`sl-ink-surface sl-ink-${tool}`}
      style={{ width: size.width, height: size.height }}
      onPointerDown={down}
      // A press here draws; it does not also select what is under it.
      onMouseDown={(e) => e.stopPropagation()}
      onPointerMove={move}
      onPointerUp={up}
      onPointerCancel={up}
    >
      {ruler ? <Ruler ruler={ruler} size={size} onChange={onRuler} /> : null}
      {live && live.length > 0 ? (
        <svg className="sl-ink-live" width={size.width} height={size.height} viewBox={`0 0 ${size.width} ${size.height}`}>
          {tool === 'lasso'
            ? <polygon points={live.map((p) => p.join(',')).join(' ')} fill="rgba(0,120,212,0.08)" stroke="#0078d4" strokeWidth="1.5" strokeDasharray="5 4" />
            : <polyline points={live.map((p) => p.join(',')).join(' ')} fill="none" stroke={look.color} strokeOpacity={look.alpha} strokeWidth={look.width * PT} strokeLinecap={look.cap === 'sq' ? 'square' : 'round'} strokeLinejoin="round" />}
        </svg>
      ) : null}
    </div>
  );
}

/** The Ruler on its own, when no pen is in hand: still on the slide, still moved and turned. */
export function RulerOverlay({ ruler, size, onChange }) {
  return (
    <div className="sl-ink-rulerhost" style={{ width: size.width, height: size.height }}>
      <Ruler ruler={ruler} size={size} onChange={onChange} />
    </div>
  );
}

/** Draw → Ruler: a straight edge on the slide — dragged to move it, turned with the wheel (Shift for 15° at a time). */
function Ruler({ ruler, size, onChange }) {
  const grab = useRef(null);
  const len = size.width * 0.72;
  return (
    <div
      className="sl-ink-ruler"
      style={{ left: ruler.x - len / 2, top: ruler.y - ruler.depth / 2, width: len, height: ruler.depth, transform: `rotate(${ruler.angle}deg)` }}
      onPointerDown={(e) => { e.stopPropagation(); e.currentTarget.setPointerCapture?.(e.pointerId); grab.current = { x: e.clientX, y: e.clientY, from: { ...ruler } }; }}
      onPointerMove={(e) => {
        const g = grab.current;
        if (!g) return;
        e.stopPropagation();
        const host = e.currentTarget.parentElement.getBoundingClientRect();
        const k = size.width / host.width;
        onChange({ ...g.from, x: g.from.x + (e.clientX - g.x) * k, y: g.from.y + (e.clientY - g.y) * k });
      }}
      onPointerUp={(e) => { e.stopPropagation(); grab.current = null; }}
      onWheel={(e) => { e.stopPropagation(); onChange({ ...ruler, angle: Math.round(((ruler.angle + Math.sign(e.deltaY) * (e.shiftKey ? 15 : 1)) % 360) * 10) / 10 }); }}
      title={t('Ruler — drag to move, the wheel to turn it; a stroke begun along its top edge follows it')}
    >
      <span className="sl-ink-ruler-angle">{Math.round(((ruler.angle % 360) + 360) % 360)}°</span>
    </div>
  );
}

/**
 * Ink Replay: the strokes drawn again, one after another, in the order
 * they were made, over the stage. Answers a promise for when it is done.
 */
export async function replayInk(root, ids, { each = 450 } = {}) {
  const paths = ids.map((id) => root.querySelector(`[data-shape="${CSS.escape(String(id))}"] path`)).filter(Boolean);
  const kept = paths.map((p) => ({ p, dash: p.style.strokeDasharray, offset: p.style.strokeDashoffset, transition: p.style.transition }));
  for (const p of paths) {
    const len = p.getTotalLength?.() || 0;
    p.style.transition = 'none';
    p.style.strokeDasharray = `${len} ${len}`;
    p.style.strokeDashoffset = `${len}`;
  }
  for (const p of paths) {
    // eslint-disable-next-line no-unused-expressions
    p.getBoundingClientRect();
    p.style.transition = `stroke-dashoffset ${each}ms linear`;
    p.style.strokeDashoffset = '0';
    await new Promise((r) => setTimeout(r, each + 40));
  }
  for (const k of kept) Object.assign(k.p.style, { strokeDasharray: k.dash, strokeDashoffset: k.offset, transition: k.transition });
}

export const INK_CSS = `
.sl-ink-surface { position: absolute; left: 0; top: 0; z-index: 20; touch-action: none; cursor: crosshair; }
.sl-ink-surface.sl-ink-eraser { cursor: cell; }
.sl-ink-surface.sl-ink-lasso { cursor: crosshair; }
.sl-ink-live { position: absolute; left: 0; top: 0; pointer-events: none; overflow: visible; }
.wd-ink-host { position: absolute; left: 0; top: 0; width: 0; height: 0; z-index: 30; }
.rw-btn.sl-pen svg { color: var(--pen); filter: drop-shadow(0 0 1px rgba(0,0,0,0.45)); }
.sl-ink-rulerhost { position: absolute; left: 0; top: 0; z-index: 19; pointer-events: none; }
.sl-ink-rulerhost .sl-ink-ruler { pointer-events: auto; }
.sl-ink-ruler { position: absolute; transform-origin: 50% 50%; background: repeating-linear-gradient(90deg, rgba(40,44,52,0.9) 0 1px, transparent 1px 12px), rgba(230,232,236,0.92); border: 1px solid rgba(40,44,52,0.6); border-radius: 4px; cursor: move; box-shadow: 0 2px 8px rgba(0,0,0,0.18); background-size: 12px 30%, auto; background-repeat: repeat-x; display: grid; place-items: center; }
.sl-ink-ruler-angle { font: 600 12px/1 var(--font-ui, sans-serif); color: #222; background: rgba(255,255,255,0.85); padding: 3px 6px; border-radius: 4px; pointer-events: none; }
`;
