// The tools on the picture (docs/STUDIO.md 7.2, I10 to I14): what a press,
// a drag and a release over the canvas mean for the tool in hand. A
// rectangle or an oval dragged out, a lasso drawn, a click for the wand, the
// colour, the bucket; a stroke for the brush, the eraser, the clone stamp and
// the healing brush (Alt+click sets where they copy from); a drag for the
// gradient; four corner handles for free transform and four pins for
// perspective. Everything is handed to the window in the canvas's own
// pixels; this file only follows the pointer and draws the guides.

import React, { useEffect, useRef, useState } from 'react';

/** Pointer events, with the pen's pressure where it gives one. */
const pressureOf = (e) => (e.pointerType === 'pen' ? (e.pressure || 0.5) : 1);

export function ToolStage({ tool, size, onSelect, onPick, onStroke, onSource, onGradient, quad, onQuad, onQuadEnd, pins, onPins, brush, source }) {
  const ref = useRef(null);
  const [drag, setDrag] = useState(null);
  const [hover, setHover] = useState(null);
  const live = useRef(null);
  const [box, setBox] = useState({ width: 0, height: 0 });

  // The overlay's size on screen, for drawing the guides at the picture's scale.
  useEffect(() => {
    const el = ref.current;
    if (!el) return undefined;
    const measure = () => setBox({ width: el.clientWidth, height: el.clientHeight });
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const k = box.width ? box.width / size.width : 1;
  const toCanvas = (e) => {
    const r = ref.current.getBoundingClientRect();
    const s = r.width / size.width;
    return { x: (e.clientX - r.left) / s, y: (e.clientY - r.top) / s, pressure: pressureOf(e) };
  };

  const down = (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    const el = ref.current;
    try { el.setPointerCapture(e.pointerId); } catch { /* a synthetic event has no pointer to hold */ }
    const p = toCanvas(e);
    const mods = { shift: e.shiftKey, alt: e.altKey, ctrl: e.ctrlKey || e.metaKey };
    if (tool === 'wand' || tool === 'colour' || tool === 'bucket') {
      onPick(tool, p, mods);
      return;
    }
    if ((tool === 'clone' || tool === 'heal') && e.altKey) {
      onSource(p);
      return;
    }
    if (tool === 'transform' && quad) {
      const tol = 10 / (el.getBoundingClientRect().width / size.width);
      const index = quad.findIndex((q) => tol >= Math.hypot(q[0] - p.x, q[1] - p.y));
      live.current = { kind: 'quad', index: index >= 0 ? index : 'move', start: quad.map((q) => [...q]), from: p };
      return;
    }
    if (tool === 'perspective' && pins) {
      const tol = 12 / (el.getBoundingClientRect().width / size.width);
      const index = pins.findIndex((q) => tol >= Math.hypot(q[0] - p.x, q[1] - p.y));
      if (index >= 0) live.current = { kind: 'pin', index };
      return;
    }
    if (tool === 'brush' || tool === 'eraser' || tool === 'clone' || tool === 'heal') {
      const stroke = onStroke(p);
      live.current = stroke ? { kind: 'stroke', stroke } : null;
      return;
    }
    if (tool === 'marquee' || tool === 'ellipse' || tool === 'gradient') {
      live.current = { kind: tool, from: p, mods };
      setDrag({ kind: tool, from: p, to: p });
      return;
    }
    if (tool === 'lasso') {
      live.current = { kind: 'lasso', points: [[p.x, p.y]], mods };
      setDrag({ kind: 'lasso', points: [[p.x, p.y]] });
    }
  };

  const move = (e) => {
    const p = toCanvas(e);
    setHover(p);
    const l = live.current;
    if (!l) return;
    if (l.kind === 'stroke') {
      const events = typeof e.getCoalescedEvents === 'function' ? e.getCoalescedEvents() : [];
      for (const ev of events.length ? events : [e]) l.stroke.move(toCanvas(ev));
      return;
    }
    if (l.kind === 'quad') {
      onQuad(l.index, l.start, p, { shift: e.shiftKey, from: l.from });
      return;
    }
    if (l.kind === 'pin') {
      const next = pins.map((q, i) => (i === l.index ? [Math.max(0, Math.min(size.width, p.x)), Math.max(0, Math.min(size.height, p.y))] : q));
      onPins(next);
      return;
    }
    if (l.kind === 'lasso') {
      const last = l.points[l.points.length - 1];
      if (Math.hypot(last[0] - p.x, last[1] - p.y) * k < 2) return;
      l.points.push([p.x, p.y]);
      setDrag({ kind: 'lasso', points: [...l.points] });
      return;
    }
    let to = p;
    if (e.shiftKey && (l.kind === 'marquee' || l.kind === 'ellipse') && !l.mods.shift) {
      // Shift while dragging keeps it square or round.
      const s = Math.max(Math.abs(p.x - l.from.x), Math.abs(p.y - l.from.y));
      to = { x: l.from.x + Math.sign(p.x - l.from.x || 1) * s, y: l.from.y + Math.sign(p.y - l.from.y || 1) * s };
    }
    l.to = to;
    setDrag({ kind: l.kind, from: l.from, to });
  };

  const up = (e) => {
    const l = live.current;
    live.current = null;
    setDrag(null);
    try { ref.current?.releasePointerCapture(e.pointerId); } catch { /* not held */ }
    if (!l) return;
    if (l.kind === 'stroke') { l.stroke.end(); return; }
    if (l.kind === 'quad') { onQuadEnd?.(); return; }
    if (l.kind === 'pin') return;
    if (l.kind === 'lasso') {
      if (l.points.length > 2) onSelect('lasso', l.points, l.mods);
      return;
    }
    const to = l.to || l.from;
    if (l.kind === 'gradient') {
      if (Math.hypot(to.x - l.from.x, to.y - l.from.y) * k >= 3) onGradient(l.from, to);
      return;
    }
    const rect = { x: Math.min(l.from.x, to.x), y: Math.min(l.from.y, to.y), width: Math.abs(to.x - l.from.x), height: Math.abs(to.y - l.from.y) };
    // A click without a drag is no selection: it lets go of the one there is.
    onSelect(l.kind, rect.width * k < 3 || rect.height * k < 3 ? null : rect, l.mods);
  };

  const ring = brush && hover && (tool === 'brush' || tool === 'eraser' || tool === 'clone' || tool === 'heal') ? brush * k : 0;
  const pt = (p) => `${p[0] * k},${p[1] * k}`;
  return (
    <div
      ref={ref}
      className={`im-tool-stage t-${tool}`}
      data-stage-tool={tool}
      onPointerDown={down}
      onPointerMove={move}
      onPointerUp={up}
      onPointerCancel={up}
      onPointerLeave={() => setHover(null)}
    >
      <svg className="im-tool-guides" width={box.width} height={box.height}>
        {drag && (drag.kind === 'marquee' || drag.kind === 'gradient' || drag.kind === 'ellipse') && drag.to ? (
          drag.kind === 'gradient' ? (
            <g className="im-guide-line">
              <line x1={drag.from.x * k} y1={drag.from.y * k} x2={drag.to.x * k} y2={drag.to.y * k} />
              <circle cx={drag.from.x * k} cy={drag.from.y * k} r={4} />
              <circle cx={drag.to.x * k} cy={drag.to.y * k} r={4} />
            </g>
          ) : drag.kind === 'ellipse' ? (
            <ellipse className="im-guide-dash" cx={((drag.from.x + drag.to.x) / 2) * k} cy={((drag.from.y + drag.to.y) / 2) * k} rx={(Math.abs(drag.to.x - drag.from.x) / 2) * k} ry={(Math.abs(drag.to.y - drag.from.y) / 2) * k} />
          ) : (
            <rect className="im-guide-dash" x={Math.min(drag.from.x, drag.to.x) * k} y={Math.min(drag.from.y, drag.to.y) * k} width={Math.abs(drag.to.x - drag.from.x) * k} height={Math.abs(drag.to.y - drag.from.y) * k} />
          )
        ) : null}
        {drag && drag.kind === 'lasso' ? <polyline className="im-guide-dash" points={drag.points.map(pt).join(' ')} /> : null}
        {tool === 'transform' && quad ? (
          <g className="im-guide-quad" data-role="transform-quad">
            <polygon points={quad.map(pt).join(' ')} />
            {quad.map((q, i) => <rect key={i} className="im-quad-handle" data-corner={i} x={q[0] * k - 5} y={q[1] * k - 5} width={10} height={10} />)}
          </g>
        ) : null}
        {tool === 'perspective' && pins ? (
          <g className="im-guide-pins" data-role="perspective-pins">
            <polygon points={pins.map(pt).join(' ')} />
            {[0, 1, 2, 3].map((i) => <line key={`l${i}`} className="im-pin-grid" x1={(pins[0][0] + (pins[1][0] - pins[0][0]) * (i / 3)) * k} y1={(pins[0][1] + (pins[1][1] - pins[0][1]) * (i / 3)) * k} x2={(pins[3][0] + (pins[2][0] - pins[3][0]) * (i / 3)) * k} y2={(pins[3][1] + (pins[2][1] - pins[3][1]) * (i / 3)) * k} />)}
            {[1, 2].map((i) => <line key={`h${i}`} className="im-pin-grid" x1={(pins[0][0] + (pins[3][0] - pins[0][0]) * (i / 3)) * k} y1={(pins[0][1] + (pins[3][1] - pins[0][1]) * (i / 3)) * k} x2={(pins[1][0] + (pins[2][0] - pins[1][0]) * (i / 3)) * k} y2={(pins[1][1] + (pins[2][1] - pins[1][1]) * (i / 3)) * k} />)}
            {pins.map((q, i) => <circle key={i} className="im-pin" data-pin={i} cx={q[0] * k} cy={q[1] * k} r={6} />)}
          </g>
        ) : null}
        {source && (tool === 'clone' || tool === 'heal') ? (
          <g className="im-guide-source" data-role="clone-mark">
            <line x1={source.x * k - 8} y1={source.y * k} x2={source.x * k + 8} y2={source.y * k} />
            <line x1={source.x * k} y1={source.y * k - 8} x2={source.x * k} y2={source.y * k + 8} />
            <circle cx={source.x * k} cy={source.y * k} r={5} />
          </g>
        ) : null}
        {ring ? <circle className="im-brush-ring" cx={hover.x * k} cy={hover.y * k} r={Math.max(2, ring / 2)} /> : null}
      </svg>
    </div>
  );
}

export const TOOL_STAGE_CSS = `
.im-tool-stage { position: absolute; inset: 0; line-height: normal; z-index: 4; touch-action: none; cursor: crosshair; }
.im-tool-stage.t-brush, .im-tool-stage.t-eraser, .im-tool-stage.t-clone, .im-tool-stage.t-heal { cursor: none; }
.im-tool-stage.t-transform { cursor: move; }
.im-tool-stage.t-perspective { cursor: default; }
.im-tool-guides { position: absolute; left: 0; top: 0; overflow: visible; pointer-events: none; }
.im-guide-dash { fill: none; stroke: #fff; stroke-width: 1.5; stroke-dasharray: 4 4; filter: drop-shadow(0 0 1px rgba(0,0,0,0.9)); }
.im-guide-line line { stroke: #fff; stroke-width: 1.5; filter: drop-shadow(0 0 1px rgba(0,0,0,0.9)); }
.im-guide-line circle { fill: #fff; stroke: var(--accent); stroke-width: 1.5; }
.im-guide-quad polygon { fill: none; stroke: var(--accent); stroke-width: 1.5; }
.im-quad-handle { fill: #fff; stroke: var(--accent); stroke-width: 1.5; }
.im-guide-pins polygon { fill: rgba(47,125,225,0.12); stroke: var(--accent); stroke-width: 1.5; }
.im-pin-grid { stroke: rgba(255,255,255,0.7); stroke-width: 1; stroke-dasharray: 3 3; }
.im-pin { fill: #fff; stroke: var(--accent); stroke-width: 2; }
.im-guide-source line { stroke: #fff; stroke-width: 1.5; filter: drop-shadow(0 0 1px rgba(0,0,0,0.9)); }
.im-guide-source circle { fill: none; stroke: #fff; stroke-width: 1.5; filter: drop-shadow(0 0 1px rgba(0,0,0,0.9)); }
.im-brush-ring { fill: none; stroke: #fff; stroke-width: 1; filter: drop-shadow(0 0 1px rgba(0,0,0,0.95)); }
`;
