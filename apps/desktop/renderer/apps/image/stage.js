// The layers on the canvas (I8): a click picks the layer under the pointer
// (the top one that shows and is not locked), a drag moves it, a corner
// handle resizes it from the opposite corner, and the round handle above it
// turns it (Shift in steps of 15 degrees). Measured on the preview's own plan
// with the engine's tools (layerBounds, layerHandles, resizePatch), so the
// handles sit where the engine paints.

import React, { useRef } from 'react';
import { layerBounds, layerHandles, scaleFromDrag, resizePatch, wordArtResize } from '@rutba/studio/host';
import { resizeCentred, rotationAt } from '@rutba/studio/layers';

const measure = () => document.createElement('canvas').getContext('2d');

/** The topmost layer under (x, y) in plan pixels that shows and can be picked. */
export function layerAt(plan, x, y, isLocked) {
  const ctx = measure();
  const ordered = [...plan.layers].sort((a, b) => (b.z || 0) - (a.z || 0));
  for (const layer of ordered) {
    if (layer.visible === false || isLocked(layer.id)) continue;
    if (layer.type === 'adjust' && layer.fw == null && layer.fh == null) continue;
    const b = layerBounds(ctx, plan, layer);
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

/** The patch a corner drag by `k` writes for a layer of any kind. */
function resized(plan, layer, start, kind, k) {
  if (layer.type === 'wordart') return { ...wordArtResize(layer, k) };
  if (layer.type === 'text' || layer.type === 'image' || layer.type === 'qr') return resizePatch(measure(), plan, layer, kind, k);
  return resizeCentred({ width: plan.W, height: plan.H }, layer, start, kind, k);
}

export function LayerStage({ plan, cssPerPx, doc, selected, onSelect, isLocked, onBegin, onLive, onOpenText }) {
  const ref = useRef(null);
  const ctx = measure();
  const primary = selected[selected.length - 1] || null;
  const primaryLayer = primary ? plan?.layers.find((l) => l.id === primary) : null;
  // An adjustment over the whole canvas has nowhere to be moved to: no handles, no box.
  const whole = (l) => l && l.type === 'adjust' && l.fw == null && l.fh == null;
  const handles = primaryLayer && !isLocked(primary) && primaryLayer.visible !== false && !whole(primaryLayer) ? layerHandles(ctx, plan, primaryLayer) : null;

  const toPlan = (e) => {
    const r = ref.current.getBoundingClientRect();
    return { x: (e.clientX - r.left) / cssPerPx, y: (e.clientY - r.top) / cssPerPx };
  };

  const begin = (e) => {
    if (e.button !== 0 || !plan) return;
    e.preventDefault();
    const at = toPlan(e);
    let mode = null;
    let layer = null;
    let hit = null;
    if (handles) {
      const tol = 9 / cssPerPx;
      hit = handles.handles.find((h) => tol >= Math.hypot(at.x - h.x, at.y - h.y)) || null;
      if (hit) {
        layer = primaryLayer;
        mode = hit.kind === 'rotate' ? 'rotate' : 'resize';
        if (mode === 'rotate' && layer.type === 'adjust') mode = null;
        hit = { ...hit, center: handles.center, bounds: handles.bounds };
      }
    }
    if (!mode) {
      layer = layerAt(plan, at.x, at.y, isLocked);
      if (!layer) {
        if (!e.shiftKey) onSelect([]);
        return;
      }
      if (e.shiftKey || e.ctrlKey) {
        onSelect(selected.includes(layer.id) ? selected.filter((x) => x !== layer.id) : [...selected, layer.id]);
        return;
      }
      if (!selected.includes(layer.id)) onSelect([layer.id]);
      mode = 'move';
    }
    if (e.detail === 2 && mode === 'move' && /^(text|textblock|wordart)$/.test(layer.type)) {
      onOpenText?.(layer.id);
      return;
    }
    const patch = doc.patches.find((p) => p.id === layer.id);
    if (!patch) return;
    const start = { at, patch, layer, plan, bounds: layerBounds(ctx, plan, layer) };
    let begun = false;
    const move = (ev) => {
      const now = toPlan(ev);
      if (!begun) {
        if (Math.hypot(now.x - at.x, now.y - at.y) * cssPerPx < 3) return;
        begun = true;
        onBegin(mode === 'move' ? 'move' : mode === 'rotate' ? 'rotate' : 'resize');
      }
      let change;
      if (mode === 'move') {
        const dx = (now.x - at.x) / plan.W;
        const dy = (now.y - at.y) / plan.H;
        change = { fx: +((start.patch.fx ?? start.layer.fx ?? 0.5) + dx).toFixed(4), fy: +((start.patch.fy ?? start.layer.fy ?? 0.5) + dy).toFixed(4) };
      } else if (mode === 'resize') {
        const k = scaleFromDrag(hit, at, now);
        const { id, ...rest } = resized(start.plan, start.layer, start.bounds, hit.kind, k) || {};
        change = rest;
      } else {
        change = { rot: rotationAt(hit.center, now, { snap: ev.shiftKey }) };
      }
      onLive(layer.id, change);
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const boxes = plan ? selected.map((id) => {
    const layer = plan.layers.find((l) => l.id === id);
    const b = layer && layer.visible !== false && !whole(layer) ? layerBounds(ctx, plan, layer) : null;
    return b ? { id, b, rot: layer.rot || 0, locked: isLocked(id) } : null;
  }).filter(Boolean) : [];

  return (
    <div className="im-stage-layers" ref={ref} onMouseDown={begin} data-selected={selected.join(',')}>
      {boxes.map(({ id, b, rot, locked }) => (
        <div
          key={id}
          className={`im-sel${locked ? ' locked' : ''}`}
          data-sel={id}
          style={{
            left: b.x * cssPerPx,
            top: b.y * cssPerPx,
            width: b.w * cssPerPx,
            height: b.h * cssPerPx,
            transform: rot ? `rotate(${rot}deg)` : undefined,
          }}
        />
      ))}
      {handles ? (
        <>
          {handles.handles.some((h) => h.kind === 'rotate') ? (
            <svg className="im-sel-stalk" style={{ left: 0, top: 0 }} width="100%" height="100%">
              {(() => {
                const r = handles.handles.find((h) => h.kind === 'rotate');
                const cx = handles.center.x;
                const cy = handles.center.y;
                const rot = ((primaryLayer.rot || 0) * Math.PI) / 180;
                // The top edge's middle, turned with the layer.
                const hx = cx + (handles.bounds.h / 2) * Math.sin(rot);
                const hy = cy - (handles.bounds.h / 2) * Math.cos(rot);
                return <line x1={hx * cssPerPx} y1={hy * cssPerPx} x2={r.x * cssPerPx} y2={r.y * cssPerPx} />;
              })()}
            </svg>
          ) : null}
          {handles.handles.map((h) => (
            <span key={h.kind} className={`im-grip g-${h.kind}`} data-grip={h.kind} style={{ left: h.x * cssPerPx, top: h.y * cssPerPx }} />
          ))}
        </>
      ) : null}
    </div>
  );
}

export const STAGE_CSS = `
.im-stage-layers { position: absolute; inset: 0; cursor: default; line-height: normal; }
.im-sel { position: absolute; box-sizing: border-box; border: 1.5px solid var(--accent); outline: 1px solid rgba(255,255,255,0.75); outline-offset: -3px; pointer-events: none; transform-origin: 50% 50%; }
.im-sel.locked { border-style: dashed; border-color: var(--ink-3); }
.im-sel-stalk { position: absolute; pointer-events: none; overflow: visible; }
.im-sel-stalk line { stroke: var(--accent); stroke-width: 1.5; }
.im-grip { position: absolute; width: 11px; height: 11px; margin: -6px 0 0 -6px; background: #fff; border: 1.5px solid var(--accent); border-radius: 2px; box-shadow: var(--shadow-1); pointer-events: none; box-sizing: border-box; }
.im-grip.g-rotate { border-radius: 50%; width: 13px; height: 13px; margin: -7px 0 0 -7px; background: var(--accent); border-color: #fff; }
`;
