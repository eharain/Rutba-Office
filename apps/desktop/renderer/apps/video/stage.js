// The stage: the frame at the playhead, painted by the engine, on the dark
// ground an editor's monitor sits on.
//
// The canvas is the project's own size (1920 by 1080 for landscape) and the
// page scales it to fit. It is painted by `paintFrame` whenever the playhead,
// the project or a decoder changes; a paused clip's decoder seeks to the
// playhead in the painter, and its 'seeked' paints the frame it lands on.
// A title or a shape on the stage can be picked and dragged into place; a
// picked one wears the selection's outline.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { paintFrame, layerBounds, withLayerStateAt, hitTestLayers } from '@rutba/studio/host';
import { t } from '@rutba/office-ui';

const MOVABLE = new Set(['text', 'shape']);
/** What can be picked and dragged on the stage: a title, its shape, and a clip drawn as an inset (a recording's camera). */
const movable = (l) => MOVABLE.has(l.type) || (l.type === 'video' && Boolean(l.fw));

export default function Stage({ plan, time, selected, onSelect, onBegin, onMove, paintTick, canvasRef }) {
  const wrapRef = useRef(null);
  const [box, setBox] = useState({ w: 0, h: 0 });
  const [outline, setOutline] = useState(null);
  const drag = useRef(null);

  // Fit the frame in the space there is, keeping its shape.
  useEffect(() => {
    const el = wrapRef.current;
    if (!el || !plan) return undefined;
    const measure = () => {
      const r = el.getBoundingClientRect();
      const pad = 24;
      const k = Math.min((r.width - pad * 2) / plan.W, (r.height - pad * 2) / plan.H);
      setBox({ w: Math.max(1, Math.floor(plan.W * k)), h: Math.max(1, Math.floor(plan.H * k)) });
    };
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [plan?.W, plan?.H]);

  // Paint the frame, and the outline of what is picked.
  useEffect(() => {
    const c = canvasRef.current;
    if (!c || !plan) return;
    if (c.width !== plan.W) c.width = plan.W;
    if (c.height !== plan.H) c.height = plan.H;
    const g = c.getContext('2d');
    paintFrame(g, plan, time);
    const picked = plan.layers.find((l) => selected.includes(l.id) && movable(l));
    if (picked && picked.timing && time >= picked.timing.start && time <= picked.timing.end) {
      const b = withLayerStateAt(plan, picked, time, () => layerBounds(g, plan, picked));
      setOutline(b ? { x: b.x / plan.W, y: b.y / plan.H, w: b.w / plan.W, h: b.h / plan.H } : null);
    } else setOutline(null);
  }, [plan, time, selected, paintTick, canvasRef]);

  const toFrame = useCallback((e) => {
    const r = canvasRef.current.getBoundingClientRect();
    return { x: ((e.clientX - r.left) / r.width) * plan.W, y: ((e.clientY - r.top) / r.height) * plan.H };
  }, [plan, canvasRef]);

  const down = (e) => {
    if (!plan || e.button !== 0) return;
    const g = canvasRef.current.getContext('2d');
    const at = toFrame(e);
    // Only what is on screen at the playhead can be picked.
    const live = { ...plan, layers: plan.layers.filter((l) => movable(l) && l.visible !== false && (!l.timing || (time >= l.timing.start && time <= l.timing.end))) };
    const hit = hitTestLayers(g, live, at.x, at.y);
    if (!hit) {
      onSelect([]);
      return;
    }
    onSelect([hit.id]);
    drag.current = { id: hit.id, from: at, moved: false, inset: hit.layer?.type === 'video' };
    const move = (ev) => {
      const d = drag.current;
      const p = toFrame(ev);
      const dx = (p.x - d.from.x) / plan.W;
      const dy = (p.y - d.from.y) / plan.H;
      if (!d.moved && Math.abs(dx) + Math.abs(dy) < 0.004) return;
      if (!d.moved) {
        d.moved = true;
        onBegin(d.inset ? t('Move the inset') : t('Move title'));
      }
      onMove(d.id, dx, dy);
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      drag.current = null;
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  return (
    <div className="vd-stage" ref={wrapRef}>
      <div className="vd-frame" style={{ width: box.w, height: box.h }}>
        <canvas ref={canvasRef} className="vd-canvas" onPointerDown={down} />
        {outline ? (
          <div
            className="vd-outline"
            style={{ left: `${outline.x * 100}%`, top: `${outline.y * 100}%`, width: `${outline.w * 100}%`, height: `${outline.h * 100}%` }}
          />
        ) : null}
      </div>
    </div>
  );
}
