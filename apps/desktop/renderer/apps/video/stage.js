// The stage: the frame at the playhead, painted by the engine, on the dark
// ground an editor's monitor sits on.
//
// The canvas is the project's own size (1920 by 1080 for landscape) and the
// page scales it to fit. It is painted by `paintFrame` whenever the playhead,
// the project or a decoder changes; a paused clip's decoder seeks to the
// playhead in the painter, and its 'seeked' paints the frame it lands on.
//
// Everything that shows at the playhead can be picked on the stage: a clip,
// a still, a title or a shape, the top one under the pointer. A picked one
// wears a frame turned with it, with a grip at each corner and a round grip
// above it: a drag moves it, a corner grip resizes it from the opposite
// corner, and the round grip turns it to any angle (Shift in steps of 15
// degrees). A clip or a still that fills the frame has corner grips only: the
// first drag makes it a picture in picture, which can then be moved and
// turned. The grips are measured with the engine's own tools (layerBounds,
// layerHandles, resizePatch), as in Image, so they sit where the engine paints
// and the export draws what the stage shows.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { paintFrame, layerBounds, layerHandles, withLayerStateAt, scaleFromDrag, resizePatch } from '@rutba/studio/host';
import { resizeCentred, rotationAt } from '@rutba/studio/layers';
import { t } from '@rutba/office-ui';
import { keepTurnedCorner } from '../image/stage.js';

/** The kinds of clip that are drawn on the frame. */
const PICTURE = new Set(['video', 'image', 'text', 'shape']);
const measure = () => document.createElement('canvas').getContext('2d');

/** A clip or still that fills the frame rather than standing in a box of its own. */
export const fillsFrame = (l) => (l.type === 'video' && !l.fw) || (l.type === 'image' && Boolean(l.fit));

/** What can be picked at `time`: drawn, showing, on screen at the playhead, and not locked (an introduction's background and screen, or a locked lane). */
export function liveLayers(plan, time, isLocked = () => false) {
  return plan.layers.filter((l) => PICTURE.has(l.type) && !l.locked && !isLocked(l.id) && l.visible !== false && (!l.timing || (time >= l.timing.start && time <= l.timing.end)));
}

/** The topmost live layer under (x, y) in frame pixels, by its turned outline. */
export function pickAt(ctx, plan, layers, time, x, y) {
  const ordered = [...layers].sort((a, b) => (b.z || 0) - (a.z || 0));
  for (const layer of ordered) {
    const b = withLayerStateAt(plan, layer, time, () => layerBounds(ctx, plan, layer));
    if (!b) continue;
    let px = x;
    let py = y;
    if (layer.rot) {
      const cx = b.x + b.w / 2;
      const cy = b.y + b.h / 2;
      const a = (-layer.rot * Math.PI) / 180;
      px = cx + (x - cx) * Math.cos(a) - (y - cy) * Math.sin(a);
      py = cy + (x - cx) * Math.sin(a) + (y - cy) * Math.cos(a);
    }
    if (px >= b.x && px <= b.x + b.w && py >= b.y && py <= b.y + b.h) return layer;
  }
  return null;
}

/** A still's height over its width, as it is drawn. */
function stillAspect(layer) {
  const img = layer.src?.img || layer.src;
  const w = img?.naturalWidth || img?.width;
  const h = img?.naturalHeight || img?.height;
  return w && h ? h / w : 9 / 16;
}

/**
 * The change a corner drag by `k` makes, upright, from the grip `kind`. A
 * still that fills the frame is made a picture of its own size, its first
 * box the frame scaled about the corner opposite the grip.
 */
function resizedUpright(ctx, plan, layer, start, kind, k) {
  if (layer.type === 'image' && layer.fit) {
    const kk = Math.max(0.1, Math.min(1, k));
    const w = plan.W * kk;
    const h = w * stillAspect(layer);
    const x = kind.includes('w') ? start.x + start.w - w : start.x;
    const y = kind.includes('n') ? start.y + start.h - h : start.y;
    return { fit: undefined, fx: +(x / plan.W).toFixed(4), fy: +(y / plan.H).toFixed(4), fw: +(w / plan.W).toFixed(4) };
  }
  if (layer.type === 'shape') return resizeCentred({ width: plan.W, height: plan.H }, layer, start, kind, k);
  return resizePatch(ctx, plan, layer, kind, k);
}

/** The change a corner drag makes, with the opposite corner held where it stands turned. */
export function resizedLayer(ctx, plan, layer, start, kind, k) {
  const upright = resizedUpright(ctx, plan, layer, start, kind, k);
  if (!upright) return null;
  const { id, ...rest } = keepTurnedCorner(plan, layer, start, kind, upright) || upright;
  if ('fit' in upright) rest.fit = undefined;
  return rest;
}

export default function Stage({ plan, time, selected, onSelect, onBegin, onMove, onLive, isLocked = () => false, paintTick, canvasRef }) {
  const wrapRef = useRef(null);
  const [box, setBox] = useState({ w: 0, h: 0 });
  const [frame, setFrame] = useState(null);
  const drag = useRef(null);

  // Fit the frame in the space there is, keeping its shape.
  useEffect(() => {
    const el = wrapRef.current;
    if (!el || !plan) return undefined;
    const measureBox = () => {
      const r = el.getBoundingClientRect();
      const pad = 24;
      const k = Math.min((r.width - pad * 2) / plan.W, (r.height - pad * 2) / plan.H);
      setBox({ w: Math.max(1, Math.floor(plan.W * k)), h: Math.max(1, Math.floor(plan.H * k)) });
    };
    measureBox();
    const ro = new ResizeObserver(measureBox);
    ro.observe(el);
    return () => ro.disconnect();
  }, [plan?.W, plan?.H]);

  // Paint the frame, and measure the frame and grips of what is picked.
  useEffect(() => {
    const c = canvasRef.current;
    if (!c || !plan) return;
    if (c.width !== plan.W) c.width = plan.W;
    if (c.height !== plan.H) c.height = plan.H;
    const g = c.getContext('2d');
    paintFrame(g, plan, time);
    const picked = liveLayers(plan, time, isLocked).find((l) => l.id === selected[selected.length - 1]) || null;
    if (!picked) { setFrame(null); return; }
    const shown = withLayerStateAt(plan, picked, time, () => {
      const b = layerBounds(g, plan, picked);
      const h = b ? layerHandles(g, plan, picked) : null;
      return b ? { b, h, rot: picked.rot || 0 } : null;
    });
    if (!shown) { setFrame(null); return; }
    const grips = (shown.h?.handles || []).filter((p) => !(p.kind === 'rotate' && fillsFrame(picked)));
    setFrame({ id: picked.id, type: picked.type, ...shown, grips });
    // isLocked is read through the plan and the lanes it was made from.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [plan, time, selected, paintTick, canvasRef]);

  const scale = plan ? box.w / plan.W : 1;

  const toFrame = useCallback((e) => {
    const r = canvasRef.current.getBoundingClientRect();
    return { x: ((e.clientX - r.left) / r.width) * plan.W, y: ((e.clientY - r.top) / r.height) * plan.H };
  }, [plan, canvasRef]);

  const down = (e) => {
    if (!plan || e.button !== 0) return;
    e.preventDefault();
    const ctx = measure();
    const at = toFrame(e);
    const live = liveLayers(plan, time, isLocked);
    let mode = null;
    let layer = null;
    let grip = null;
    if (frame) {
      const tol = 10 / scale;
      grip = frame.grips.find((p) => tol >= Math.hypot(at.x - p.x, at.y - p.y)) || null;
      if (grip) {
        layer = live.find((l) => l.id === frame.id) || null;
        mode = grip.kind === 'rotate' ? 'rotate' : 'resize';
      }
    }
    if (!layer) {
      layer = pickAt(ctx, plan, live, time, at.x, at.y);
      if (!layer) {
        onSelect([]);
        return;
      }
      if (e.shiftKey || e.ctrlKey) {
        onSelect(selected.includes(layer.id) ? selected.filter((x) => x !== layer.id) : [...selected, layer.id]);
        return;
      }
      if (!selected.includes(layer.id)) onSelect([layer.id]);
      // A clip that fills the frame stays where it is: only its grips change it.
      mode = fillsFrame(layer) ? null : 'move';
    }
    if (!mode) return;
    const start = withLayerStateAt(plan, layer, time, () => layerBounds(ctx, plan, layer));
    const centre = start ? { x: start.x + start.w / 2, y: start.y + start.h / 2 } : null;
    const hit = grip && frame?.h ? { ...grip, center: frame.h.center, bounds: frame.h.bounds } : null;
    drag.current = { id: layer.id, moved: false };
    const move = (ev) => {
      const d = drag.current;
      if (!d) return;
      const p = toFrame(ev);
      if (!d.moved) {
        if (Math.hypot(p.x - at.x, p.y - at.y) * scale < 3) return;
        d.moved = true;
        onBegin(mode === 'move' ? (layer.type === 'text' ? t('Move title') : t('Move')) : mode === 'rotate' ? t('Rotate') : t('Resize'));
      }
      if (mode === 'move') onMove(d.id, (p.x - at.x) / plan.W, (p.y - at.y) / plan.H);
      else if (mode === 'rotate' && centre) onLive(d.id, { rot: rotationAt(centre, p, { snap: ev.shiftKey }) });
      else if (mode === 'resize' && hit && start) {
        const change = resizedLayer(ctx, plan, layer, start, grip.kind, scaleFromDrag(hit, at, p));
        if (change) onLive(d.id, change);
      }
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      drag.current = null;
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const rotGrip = frame?.grips.find((p) => p.kind === 'rotate');
  return (
    <div className="vd-stage" ref={wrapRef}>
      <div className="vd-frame" style={{ width: box.w, height: box.h }} data-picked={frame?.id || ''}>
        <canvas ref={canvasRef} className="vd-canvas" onPointerDown={down} />
        {frame ? (
          <div className="vd-picked" aria-hidden="true">
            <div
              className="vd-outline"
              data-rot={frame.rot}
              style={{
                left: frame.b.x * scale,
                top: frame.b.y * scale,
                width: frame.b.w * scale,
                height: frame.b.h * scale,
                transform: frame.rot ? `rotate(${frame.rot}deg)` : undefined,
              }}
            />
            {rotGrip ? (
              <svg className="vd-stalk" width="100%" height="100%">
                {(() => {
                  const cx = frame.h.center.x;
                  const cy = frame.h.center.y;
                  const a = (frame.rot * Math.PI) / 180;
                  const hx = cx + (frame.h.bounds.h / 2) * Math.sin(a);
                  const hy = cy - (frame.h.bounds.h / 2) * Math.cos(a);
                  return <line x1={hx * scale} y1={hy * scale} x2={rotGrip.x * scale} y2={rotGrip.y * scale} />;
                })()}
              </svg>
            ) : null}
            {frame.grips.map((p) => (
              <span key={p.kind} className={`vd-grip g-${p.kind}`} data-grip={p.kind} style={{ left: p.x * scale, top: p.y * scale }} />
            ))}
          </div>
        ) : null}
      </div>
    </div>
  );
}
