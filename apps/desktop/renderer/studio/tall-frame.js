// Set up's frame when it is on Tall (plan A11: "Set up gets Wide | Tall over the
// frame"). The scene's tall layout, drawn as it will go out: drag a picture to
// move it and a corner to size it. What is changed is kept for that scene, and
// until something is, the layout is the one made from the wide one.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, t } from '@rutba/office-ui';
import { clampBox } from '@rutba/studio/stack';
import { setTallLayers, tallKey, tallLayersFor, tallSize, tallStackOf } from '@rutba/studio/tall';
import { CanvasSlot } from './parts.js';
import { Program } from './program.js';
import { itemName } from './list.js';

const MIN = 0.06;

/** The box after a drag: the whole of it moved, or one corner. Fractions of the frame. */
export function draggedBox(box, handle, dx, dy) {
  if (!handle) return clampBox({ ...box, x: Math.min(1 - box.w, Math.max(0, box.x + dx)), y: Math.min(1 - box.h, Math.max(0, box.y + dy)) });
  let { x, y, w, h } = box;
  if (handle.includes('w')) { const nx = Math.min(x + w - MIN, Math.max(0, x + dx)); w += x - nx; x = nx; }
  if (handle.includes('e')) w = Math.max(MIN, Math.min(1 - x, w + dx));
  if (handle[0] === 'n') { const ny = Math.min(y + h - MIN, Math.max(0, y + dy)); h += y - ny; y = ny; }
  if (handle[0] === 's') h = Math.max(MIN, Math.min(1 - y, h + dy));
  return clampBox({ x, y, w, h });
}

/** @param {{ s: object }} p  the studio (window.js) */
export function TallFrame({ s }) {
  const { stack, media, settings, changeStack } = s;
  const frame = useRef(null);
  const [drag, setDrag] = useState(null); // { id, handle, layers } while something is being moved
  const canvas = useMemo(() => {
    const c = document.createElement('canvas');
    c.className = 'st-canvas st-tall-preview';
    return c;
  }, []);
  const key = tallKey(stack);
  const aspectOf = (it) => { const z = media.sizeOf(it); return z ? z.w / z.h : 16 / 9; };
  const layers = drag ? drag.layers : tallLayersFor(stack, key, { aspectOf });
  const kept = Boolean(stack.tall?.[key]);

  const program = useRef(null);
  useEffect(() => {
    const p = new Program({ media, canvas });
    const size = tallSize(settings.size);
    canvas.width = size.width;
    canvas.height = size.height;
    program.current = p;
    p.start();
    media.web?.setTall(true);
    return () => { p.stopDrawing(); media.web?.setTall(false); program.current = null; };
  }, [media, canvas, settings.size]);

  const shownStack = drag ? setTallLayers(stack, key, drag.layers) : stack;
  const shown = useRef(shownStack);
  shown.current = shownStack;
  useEffect(() => {
    const put = () => program.current?.setStack(tallStackOf(shown.current, key, { aspectOf }));
    put();
    const timer = setInterval(put, 1000);
    return () => clearInterval(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shownStack, key, media]);

  const begin = (e, l, handle) => {
    e.preventDefault();
    e.stopPropagation();
    s.setSelected(l.id);
    const r = frame.current.getBoundingClientRect();
    const start = { x: e.clientX, y: e.clientY };
    const from = layers;
    let now = null;
    const move = (ev) => {
      const box = draggedBox(l.box, handle, (ev.clientX - start.x) / r.width, (ev.clientY - start.y) / r.height);
      now = from.map((x) => (x.id === l.id ? { ...x, box } : x));
      setDrag({ id: l.id, handle, layers: now });
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      setDrag(null);
      if (now) changeStack((st) => setTallLayers(st, key, now));
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const named = (id) => { const it = s.named.items.find((x) => x.id === id); return it ? itemName(it) : ''; };
  return (
    <div className="st-tall-wrap">
      <div ref={frame} className="st-frame st-tall-frame" aria-label={t('The tall picture')}>
        <CanvasSlot canvas={canvas} />
        <div className="st-boxes">
          {layers.map((l) => (
            <div key={l.id} className={`st-box st-tall-box${s.selected === l.id ? ' on' : ''}`} data-id={l.id} title={named(l.id)}
              style={{ left: `${l.box.x * 100}%`, top: `${l.box.y * 100}%`, width: `${l.box.w * 100}%`, height: `${l.box.h * 100}%` }}
              onPointerDown={(e) => begin(e, l, null)}>
              {s.selected === l.id ? ['nw', 'ne', 'sw', 'se'].map((c) => <span key={c} className="st-handle" data-c={c} onPointerDown={(e) => begin(e, l, c)} />) : null}
            </div>
          ))}
        </div>
      </div>
      <div className="st-tall-side">
        <span className="st-tall-note">{kept ? t('This tall layout is kept for this scene.') : t('Made from the wide layout. Drag a picture to change it.')}</span>
        {kept ? <Button label={t('Make it again from the wide')} className="st-tall-reset" onClick={() => changeStack((st) => setTallLayers(st, key, null))} /> : null}
      </div>
    </div>
  );
}

export const TALL_FRAME_CSS = `
.st-frame-shape-row { flex: none; display: flex; justify-content: center; }
.st-tall-wrap { display: flex; align-items: center; justify-content: center; gap: 16px; height: 100%; width: 100%; }
.st-tall-frame { width: auto; height: 100%; aspect-ratio: 9 / 16; }
.st-tall-frame .st-canvas { width: 100%; height: 100%; display: block; }
.st-tall-frame .st-canvas-slot { width: 100%; height: 100%; }
.st-tall-side { display: flex; flex-direction: column; gap: 8px; max-width: 200px; }
.st-tall-note { font-size: 12.5px; color: var(--ink-2); }
.st-copy-shape { flex: none; }
`;
