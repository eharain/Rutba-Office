// The tone panel's two drawings: the histogram behind Levels and the curve
// you drag in Curves, as Photoshop and Paint.NET draw them.
import React, { useEffect, useRef } from 'react';
import { curveTable } from '@rutba/imaging/pixels';
import { t } from '@rutba/office-ui';

/** The luminance histogram, drawn into a small canvas, the tallest bar at the top. */
export function Histogram({ hist, height = 56, marks = null }) {
  const ref = useRef(null);
  useEffect(() => {
    const canvas = ref.current;
    if (!canvas) return;
    const w = 256;
    canvas.width = w;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, w, height);
    if (!hist) return;
    // The tallest bar would flatten the rest when one tone dominates: the
    // scale is set by the second highest percentile, as editors do.
    const counts = [...hist.l].sort((a, b) => a - b);
    const top = Math.max(1, counts[Math.floor(counts.length * 0.985)]);
    const ink = getComputedStyle(canvas).color || '#888';
    ctx.fillStyle = ink;
    ctx.globalAlpha = 0.55;
    for (let i = 0; i < 256; i++) {
      const h = Math.min(height, (hist.l[i] / top) * height);
      ctx.fillRect(i, height - h, 1, h);
    }
    ctx.globalAlpha = 1;
    if (marks) {
      ctx.fillStyle = 'rgba(0,0,0,0.18)';
      ctx.fillRect(0, 0, marks.inBlack, height);
      ctx.fillRect(marks.inWhite + 1, 0, 255 - marks.inWhite, height);
    }
  }, [hist, height, marks?.inBlack, marks?.inWhite]);
  return <canvas ref={ref} className="im-hist" aria-label={t('Histogram')} />;
}

const SIZE = 256;

/**
 * The curve: a straight line to begin, points added by clicking on it and
 * moved by dragging; a point dragged off the top or bottom, or double-clicked,
 * goes. The two ends stay at the left and right edges.
 */
export function CurveEditor({ points, onChange, hist, tone = 'var(--ink-2)' }) {
  const ref = useRef(null);
  const pts = points && points.length >= 2 ? points : [[0, 0], [255, 255]];
  const table = curveTable(pts);
  let path = '';
  for (let i = 0; i < 256; i++) path += `${i ? 'L' : 'M'}${i} ${255 - table[i]}`;

  const toCurve = (e) => {
    const r = ref.current.getBoundingClientRect();
    // The view box runs from -8 to 264, so a point at the edge is whole and can be taken hold of.
    return [((e.clientX - r.left) / r.width) * (SIZE + 16) - 8, 255 - (((e.clientY - r.top) / r.height) * (SIZE + 16) - 8)];
  };

  const down = (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    const [x, y] = toCurve(e);
    let list = pts.map((p) => [...p]);
    let at = list.findIndex(([px, py]) => 12 > Math.hypot(px - x, py - y));
    if (at < 0) {
      // A new point, on the curve where it was clicked rather than where the pointer is.
      const nx = Math.round(Math.max(1, Math.min(254, x)));
      list.push([nx, table[nx]]);
      list.sort((a, b) => a[0] - b[0]);
      at = list.findIndex((p) => p[0] === nx);
      onChange(list);
    }
    const svg = ref.current;
    svg.setPointerCapture?.(e.pointerId);
    const move = (ev) => {
      const [mx, my] = toCurve(ev);
      const next = list.map((p) => [...p]);
      const first = at === 0;
      const last = at === next.length - 1;
      const lo = first ? 0 : next[at - 1][0] + 1;
      const hi = last ? 255 : next[at + 1][0] - 1;
      const x2 = first ? 0 : last ? 255 : Math.round(Math.max(lo, Math.min(hi, mx)));
      const gone = !first && !last && (my > 290 || my < -35);
      next[at] = [x2, Math.round(Math.max(0, Math.min(255, my)))];
      list = next;
      onChange(gone ? next.filter((_, i) => i !== at) : next);
    };
    const up = (ev) => {
      const [, my] = toCurve(ev);
      if (at > 0 && at < list.length - 1 && (my > 290 || my < -35)) onChange(list.filter((_, i) => i !== at));
      svg.removeEventListener('pointermove', move);
      svg.removeEventListener('pointerup', up);
    };
    svg.addEventListener('pointermove', move);
    svg.addEventListener('pointerup', up);
  };

  const remove = (e) => {
    const [x, y] = toCurve(e);
    const at = pts.findIndex(([px, py]) => 12 > Math.hypot(px - x, py - y));
    if (at > 0 && at < pts.length - 1) onChange(pts.filter((_, i) => i !== at));
  };

  const bars = [];
  if (hist) {
    const counts = [...hist.l].sort((a, b) => a - b);
    const top = Math.max(1, counts[Math.floor(counts.length * 0.985)]);
    for (let i = 0; i < 256; i += 2) {
      const h = Math.min(255, ((hist.l[i] + hist.l[i + 1]) / 2 / top) * 120);
      if (h >= 0.5) {
        bars.push(
          <rect key={i} x={i} y={255 - h} width={2} height={h} />
        );
      }
    }
  }

  return (
    <svg
      ref={ref}
      className="im-curve"
      viewBox="-8 -8 272 272"
      role="img"
      aria-label={t('Tone curve')}
      onPointerDown={down}
      onDoubleClick={remove}
    >
      <rect x="0" y="0" width="256" height="256" className="im-curve-bg" />
      <g className="im-curve-hist">{bars}</g>
      {[64, 128, 192].map((v) => (
        <g key={v} className="im-curve-grid">
          <line x1={v} y1="0" x2={v} y2="256" />
          <line x1="0" y1={v} x2="256" y2={v} />
        </g>
      ))}
      <line x1="0" y1="255" x2="255" y2="0" className="im-curve-ref" />
      <path d={path} className="im-curve-line" style={{ stroke: tone }} />
      {pts.map(([x, y], i) => (
        <circle key={i} cx={x} cy={255 - y} r="5.5" className="im-curve-point" data-point={i} />
      ))}
    </svg>
  );
}
