// Insert (I8, I9): text, word art, shapes, icons, charts, a picture and
// adjustment layers, each put on the picture as a layer of its own. A
// gallery's tiles are drawn by the engine itself from the layer they insert,
// so a tile is what the click makes, never a picture of something like it.

import React, { useEffect, useRef, useState } from 'react';
import { Group, Button, t } from '@rutba/office-ui';
import { buildImagePlan, paintFrame, WORDART_PRESETS, SHAPE_GROUPS, ICON_GROUPS, CHART_KINDS, ADJUST_PRESETS } from '@rutba/studio/host';
import { SHAPE_NAMES, SHAPE_GROUP_NAMES, WORDART_NAMES, ICON_NAMES_SHOWN, ICON_GROUP_NAMES, CHART_NAMES, ADJUST_NAMES } from './words.js';

/** What each insert puts on the picture, before it is placed. */
export const NEW_LAYER = {
  text: () => ({ type: 'textblock', text: t('Your words here'), fx: 0.5, fy: 0.5, fw: 0.6, sizeFrac: 0.06, weight: 600, align: 'center', fill: { kind: 'solid', color: '#ffffff' } }),
  wordart: (preset) => ({ type: 'wordart', text: t('Headline'), preset, fx: 0.5, fy: 0.3, sizeFrac: 0.1 }),
  shape: (geometry) => ({ type: 'shape', geometry, fx: 0.5, fy: 0.5, fw: 0.28, fill: { kind: 'solid', color: '#2f7de1' }, ...(geometry === 'rect' ? { radius: 0.08 } : {}) }),
  icon: (name) => ({ type: 'icon', name, fx: 0.5, fy: 0.5, fw: 0.16, fill: { kind: 'solid', color: '#ffffff' } }),
  chart: (chart) => ({
    type: 'chart', chart, fx: 0.5, fy: 0.5, fw: 0.6, fh: 0.4,
    ...(chart === 'pie' || chart === 'donut'
      ? { labels: [t('Ours'), t('Theirs'), t('Other')], values: [58, 27, 15] }
      : { labels: [t('Mon'), t('Tue'), t('Wed'), t('Thu')], values: [12, 19, 9, 24] }),
  }),
  adjust: (preset) => ({ type: 'adjust', preset }),
};

/** A sample for an adjustment's tile: a band of colour the look is seen on. */
const SAMPLE = { id: 'sample', type: 'shape', geometry: 'rect', fx: 0.5, fy: 0.5, fw: 1, fh: 1, fill: { kind: 'linear', angle: 0, stops: [[0, '#e0393e'], [0.35, '#f5d90a'], [0.65, '#2e9e4f'], [1, '#2f7de1']] } };

/** One tile: the layer drawn by the engine on a small canvas. */
function Tile({ patch, width = 64, height = 44, ground = '#3b3f47' }) {
  const ref = useRef(null);
  useEffect(() => {
    const c = ref.current;
    if (!c) return;
    const dpr = window.devicePixelRatio || 1;
    const size = { width: Math.round(width * dpr), height: Math.round(height * dpr) };
    try {
      const patches = patch.type === 'adjust' ? [SAMPLE, { ...patch, id: 'tile' }] : [{ ...patch, id: 'tile' }];
      const plan = buildImagePlan({ canvas: document.createElement('canvas'), size, options: { bare: true, background: ground }, layerPatches: patches });
      c.width = plan.W;
      c.height = plan.H;
      paintFrame(c.getContext('2d'), plan, 0);
    } catch {
      /* a tile that cannot draw stays blank */
    }
  }, [patch, width, height, ground]);
  return <canvas ref={ref} className="im-tile-canvas" style={{ width, height }} />;
}

/** A gallery under a ribbon button: sections of tiles, closed by a click elsewhere or Escape. */
function Gallery({ anchor, sections, onPick, onClose, wide = false, columns = 5 }) {
  const ref = useRef(null);
  const [pos, setPos] = useState(() => {
    const r = anchor.getBoundingClientRect();
    return { x: r.left, y: r.bottom + 4 };
  });
  useEffect(() => {
    const el = ref.current;
    if (!el) return undefined;
    const r = el.getBoundingClientRect();
    const rtl = document.documentElement.dir === 'rtl';
    const a = anchor.getBoundingClientRect();
    const x = rtl ? a.right - r.width : a.left;
    setPos({ x: Math.max(8, Math.min(x, window.innerWidth - r.width - 8)), y: Math.min(a.bottom + 4, window.innerHeight - r.height - 8) });
    const off = (e) => { if (!el.contains(e.target) && !anchor.contains(e.target)) onClose(); };
    const key = (e) => e.key === 'Escape' && onClose();
    window.addEventListener('mousedown', off);
    window.addEventListener('keydown', key);
    return () => { window.removeEventListener('mousedown', off); window.removeEventListener('keydown', key); };
  }, [anchor, onClose]);
  return (
    <div ref={ref} className={`im-gallery${wide ? ' wide' : ''}`} style={{ left: pos.x, top: pos.y }} role="menu">
      {sections.map((s) => (
        <section key={s.label || 'all'} className="im-gallery-sec">
          {s.label ? <div className="im-gallery-head">{s.label}</div> : null}
          <div className="im-gallery-grid" style={{ gridTemplateColumns: `repeat(${s.columns || columns}, auto)` }}>
            {s.items.map((it) => (
              <button key={it.key} type="button" className="im-gallery-item" data-tip={it.label} aria-label={it.label} data-item={it.key} onClick={() => { onPick(it); onClose(); }}>
                <Tile patch={it.patch} width={it.w || 52} height={it.h || 40} ground={it.ground} />
                {it.caption ? <span>{it.label}</span> : null}
              </button>
            ))}
          </div>
        </section>
      ))}
    </div>
  );
}

const SECTIONS = {
  wordart: () => [{ label: null, columns: 4, items: Object.keys(WORDART_PRESETS).map((k) => ({ key: k, label: WORDART_NAMES[k] || k, patch: { ...NEW_LAYER.wordart(k), text: 'Aa', fy: 0.5, sizeFrac: 0.42 }, w: 72, h: 48, caption: true })) }],
  shape: () => SHAPE_GROUPS.map((g) => ({ label: SHAPE_GROUP_NAMES[g.label] || g.label, columns: 7, items: g.geometries.map((geo) => ({ key: geo, label: SHAPE_NAMES[geo] || geo, patch: { ...NEW_LAYER.shape(geo), fw: 0.62, fh: 0.8 }, w: 34, h: 30, ground: 'transparent' })) })),
  icon: () => ICON_GROUPS.map((g) => ({ label: ICON_GROUP_NAMES[g.label] || g.label, columns: 11, items: g.names.map((n) => ({ key: n, label: ICON_NAMES_SHOWN[n] || n, patch: { ...NEW_LAYER.icon(n), fw: 0.7, fill: { kind: 'solid', color: '#454b55' } }, w: 26, h: 26, ground: 'transparent' })) })),
  chart: () => [{ label: null, columns: 3, items: CHART_KINDS.map((k) => ({ key: k, label: CHART_NAMES[k] || k, patch: { ...NEW_LAYER.chart(k), fw: 0.92, fh: 0.86, showLegend: false, showLabels: false }, w: 72, h: 52, ground: '#ffffff', caption: true })) }],
  adjust: () => [{ label: null, columns: 4, items: Object.keys(ADJUST_PRESETS).filter((k) => k !== 'none').map((k) => ({ key: k, label: ADJUST_NAMES[k] || k, patch: NEW_LAYER.adjust(k), w: 60, h: 38, caption: true })) }],
};

/** The Insert tab's groups. `onAdd(patch)` puts a layer on the picture; `onPicture()` asks for a file. */
export function InsertGroups({ onAdd, onPicture, disabled }) {
  const [open, setOpen] = useState(null);
  const gallery = (kind, label, icon, tip) => (
    <Button
      tall
      icon={icon}
      label={label}
      title={tip}
      disabled={disabled}
      pressed={open?.kind === kind}
      data-insert={kind}
      onClick={(e) => { const a = e.currentTarget; setOpen((o) => (o?.kind === kind ? null : { kind, anchor: a })); }}
    />
  );
  const make = { wordart: NEW_LAYER.wordart, shape: NEW_LAYER.shape, icon: NEW_LAYER.icon, chart: NEW_LAYER.chart, adjust: NEW_LAYER.adjust };
  return (
    <>
      <Group label={t('Text')}>
        <Button tall icon="mkText" label={t('Text')} title={t('Text: words in a box of their own, on a layer')} disabled={disabled} data-insert="text" onClick={() => onAdd(NEW_LAYER.text())} />
        {gallery('wordart', t('WordArt'), 'wordArt', t('WordArt: a headline in a ready-made style'))}
      </Group>
      <Group label={t('Illustrations')}>
        {gallery('shape', t('Shapes'), 'shape', t('Shapes: rectangles, rounds, stars, arrows and labels'))}
        {gallery('icon', t('Icons'), 'star', t('Icons: a small drawing that stays sharp at any size'))}
        {gallery('chart', t('Chart'), 'chart', t('Chart: bars, columns, a line or a pie from a few numbers'))}
        <Button tall icon="picture" label={t('Picture')} title={t('Picture: another picture, on a layer of its own')} disabled={disabled} data-insert="picture" onClick={onPicture} />
      </Group>
      <Group label={t('Adjust')}>
        {gallery('adjust', t('Adjustment'), 'sliders', t('Adjustment: a layer that changes the look of every layer under it'))}
      </Group>
      {open ? (
        <Gallery
          anchor={open.anchor}
          sections={SECTIONS[open.kind]()}
          wide={open.kind === 'icon' || open.kind === 'shape'}
          onPick={(it) => onAdd(make[open.kind](it.key))}
          onClose={() => setOpen(null)}
        />
      ) : null}
    </>
  );
}

export const INSERT_CSS = `
.im-gallery { position: fixed; z-index: 60; background: var(--surface); border: 1px solid var(--line); border-radius: var(--r-3); box-shadow: var(--shadow-3); padding: 8px 10px 10px; max-height: 70vh; overflow: auto; }
.im-gallery-sec + .im-gallery-sec { margin-top: 6px; }
.im-gallery-head { font-size: 11px; font-weight: 600; color: var(--ink-3); padding: 4px 2px 5px; }
.im-gallery-grid { display: grid; gap: 4px; }
.im-gallery-item { display: flex; flex-direction: column; align-items: center; gap: 4px; padding: 4px; border: 1px solid transparent; border-radius: var(--r-2); background: none; cursor: pointer; font: inherit; font-size: 11px; color: var(--ink-2); }
.im-gallery-item:hover { background: var(--hover); border-color: var(--line); color: var(--ink); }
.im-gallery-item:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.im-tile-canvas { display: block; border-radius: 4px; }
`;
