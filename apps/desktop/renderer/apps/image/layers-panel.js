// The Layers panel (I7): the stack top first, a group's layers under it; an
// eye to show or hide, a padlock to lock, a name renamed by a double-click,
// a row dragged up or down the list (or into a group) to reorder; under the
// list, the chosen layer's opacity, blend mode and what its kind has to set
// (its words, its look, its colour, a chart's numbers, an adjustment's
// strength). Everything it changes goes to the window as a change to the
// project, so undo and save see it the same way as anything else.

import React, { useEffect, useRef, useState } from 'react';
import { Button, Icon, Select, t } from '@rutba/office-ui';
import { layerRows, isGroup, BLEND_MODES } from '@rutba/studio/layers';
import { ADJUST_PRESETS, WORDART_PRESETS, GEOMETRIES, ICON_NAMES, CHART_KINDS } from '@rutba/studio/host';
import { dataToText, textToData } from '@rutba/studio/chart-data';
import { nameOf, BLEND_NAMES, ADJUST_NAMES, WORDART_NAMES, SHAPE_NAMES, ICON_NAMES_SHOWN, CHART_NAMES } from './words.js';

const TYPE_ICON = { image: 'picture', textblock: 'mkText', text: 'mkText', wordart: 'wordArt', shape: 'shape', icon: 'star', chart: 'chart', adjust: 'sliders', group: 'folder' };
const COLOURS = ['#ffffff', '#111111', '#e0393e', '#f5d90a', '#2e9e4f', '#2f7de1', '#7b49d6'];

/** The adjustment sliders an adjustment layer shows: [key, label, min, max, step]. */
const ADJUST_SLIDERS = [
  ['exposure', t('Exposure'), -2, 2, 0.05],
  ['brightness', t('Brightness'), -1, 1, 0.02],
  ['contrast', t('Contrast'), -1, 1, 0.02],
  ['saturation', t('Saturation'), -1, 1, 0.02],
  ['vibrance', t('Vibrance'), -1, 1, 0.02],
  ['temperature', t('Temperature'), -1, 1, 0.02],
  ['tint', t('Tint'), -1, 1, 0.02],
  ['vignette', t('Vignette'), -1, 1, 0.02],
  ['sharpen', t('Sharpen'), 0, 1, 0.02],
];

/** The colour a layer is painted in, where it has one to choose. */
function colourOf(p) {
  if (p.type === 'wordart') return p.style?.fill?.color || WORDART_PRESETS[p.preset]?.style?.fill?.color || '#ffffff';
  return p.fill?.color || p.color || '#ffffff';
}
function colourChange(p, c) {
  if (p.type === 'wordart') return { style: { fill: { kind: 'solid', color: c } } };
  return { fill: { kind: 'solid', color: c } };
}

export function LayersPanel({ doc, selected, onSelect, onChange, onDrop, onStep, onGroup, onUngroup, onDuplicate, onRemove, assets, target = null, onTarget = null, onMask = null, hasSelection = false, canApplyMask = null }) {
  const rows = layerRows(doc);
  const [collapsed, setCollapsed] = useState(() => new Set());
  const [renaming, setRenaming] = useState(null);
  const [drag, setDrag] = useState(null);
  const listRef = useRef(null);
  const primary = selected[selected.length - 1] || null;
  const layer = primary ? doc.patches.find((p) => p.id === primary) : null;
  const shown = rows.filter((r) => !(r.parent && collapsed.has(r.parent)));

  // A row held and moved more than a few pixels is being dragged: where it
  // would land is worked out from the pointer over the rows.
  const startDrag = (e, id) => {
    if (e.button !== 0 || e.target.closest('button, input')) return;
    const y0 = e.clientY;
    let moving = false;
    const where = (ev) => {
      const items = [...(listRef.current?.querySelectorAll('.im-layer') || [])];
      for (const el of items) {
        const r = el.getBoundingClientRect();
        if (ev.clientY < r.top || ev.clientY > r.bottom) continue;
        const f = (ev.clientY - r.top) / r.height;
        const target = el.dataset.layer;
        const group = el.dataset.group === '1';
        return { target, where: group && f > 0.3 && f < 0.7 ? 'into' : f < 0.5 ? 'above' : 'below' };
      }
      const last = items[items.length - 1];
      if (last && ev.clientY > last.getBoundingClientRect().bottom) return { target: last.dataset.layer, where: 'below' };
      return null;
    };
    const move = (ev) => {
      if (!moving && Math.abs(ev.clientY - y0) < 4) return;
      moving = true;
      const at = where(ev);
      setDrag(at && at.target !== id ? { id, ...at } : { id });
    };
    const up = (ev) => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      if (moving) {
        const at = where(ev);
        if (at && at.target !== id) onDrop(id, at.target, at.where);
      }
      setDrag(null);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const pick = (e, id) => {
    if (e.ctrlKey || e.metaKey) onSelect(selected.includes(id) ? selected.filter((x) => x !== id) : [...selected, id]);
    else if (e.shiftKey && primary) {
      const ids = shown.map((r) => r.id);
      const a = ids.indexOf(primary);
      const b = ids.indexOf(id);
      onSelect(ids.slice(Math.min(a, b), Math.max(a, b) + 1));
    } else onSelect([id]);
  };

  return (
    <div className="im-layers" data-role="layers">
      <div className="im-layer-list" ref={listRef} role="listbox" aria-label={t('Layers')} aria-multiselectable="true">
        {shown.map((r) => {
          const p = r.patch;
          const hidden = p.visible === false;
          const on = selected.includes(r.id);
          const mark = drag && drag.target === r.id ? drag.where : null;
          return (
            <div
              key={r.id}
              className={`im-layer${on ? ' on' : ''}${hidden ? ' hidden' : ''}${r.depth ? ' deep' : ''}${mark ? ` drop-${mark}` : ''}${drag?.id === r.id ? ' lifted' : ''}`}
              data-layer={r.id}
              data-type={p.type}
              data-group={r.group ? '1' : '0'}
              role="option"
              aria-selected={on}
              onMouseDown={(e) => { startDrag(e, r.id); }}
              onClick={(e) => pick(e, r.id)}
            >
              {r.group ? (
                <button
                  type="button"
                  className="im-layer-fold"
                  aria-label={collapsed.has(r.id) ? t('Show its layers') : t('Hide its layers')}
                  onClick={(e) => { e.stopPropagation(); setCollapsed((s) => { const n = new Set(s); if (n.has(r.id)) n.delete(r.id); else n.add(r.id); return n; }); }}
                >
                  <Icon name={collapsed.has(r.id) ? 'chevronRight' : 'chevronDown'} size={12} />
                </button>
              ) : null}
              <button
                type="button"
                className="im-layer-eye"
                data-role="eye"
                aria-pressed={!hidden}
                data-tip={hidden ? t('Show the layer') : t('Hide the layer')}
                onClick={(e) => { e.stopPropagation(); onChange(r.id, { visible: hidden }); }}
              >
                <Icon name={hidden ? 'eyeOff' : 'eye'} size={14} />
              </button>
              <span className={`im-layer-thumbs${target?.id === r.id ? ` target-${target.part}` : ''}`}>
                <span className="im-thumb-hit pixels" data-role="target-pixels" onClick={(e) => { if (onTarget && !r.group) { e.stopPropagation(); onSelect([r.id]); onTarget(r.id, 'pixels'); } }}>
                  <Thumb patch={p} assets={assets} />
                </span>
                {p.paintMask ? (
                  <span className="im-thumb-hit mask" data-role="target-mask" data-tip={t('The layer\'s mask: white shows, black hides. Click to paint on it.')} onClick={(e) => { e.stopPropagation(); onSelect([r.id]); onTarget?.(r.id, 'mask'); }}>
                    <MaskThumb mask={p.paintMask} asset={assets?.[p.paintMask.url]} />
                  </span>
                ) : null}
              </span>
              {renaming === r.id ? (
                <input
                  className="im-layer-rename"
                  defaultValue={nameOf(p)}
                  autoFocus
                  onClick={(e) => e.stopPropagation()}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') e.currentTarget.blur();
                    if (e.key === 'Escape') { e.currentTarget.value = nameOf(p); e.currentTarget.blur(); }
                    e.stopPropagation();
                  }}
                  onBlur={(e) => {
                    const v = e.currentTarget.value.trim();
                    if (v && v !== nameOf(p)) onChange(r.id, { label: v });
                    setRenaming(null);
                  }}
                />
              ) : (
                <span className="im-layer-name" onDoubleClick={(e) => { e.stopPropagation(); setRenaming(r.id); }}>{nameOf(p)}</span>
              )}
              {p.opacity != null && p.opacity < 1 ? <span className="im-layer-meta">{`${Math.round(p.opacity * 100)}%`}</span> : null}
              <button
                type="button"
                className={`im-layer-lock${p.locked ? ' on' : ''}`}
                data-role="lock"
                aria-pressed={Boolean(p.locked)}
                data-tip={p.locked ? t('Unlock the layer') : t('Lock the layer')}
                onClick={(e) => { e.stopPropagation(); onChange(r.id, { locked: !p.locked }); }}
              >
                <Icon name={p.locked ? 'lock' : 'unlock'} size={13} />
              </button>
            </div>
          );
        })}
      </div>
      <div className="im-layer-tools">
        <Button icon="layerUp" title={t('Bring forward: one place up the stack')} disabled={!primary} onClick={() => onStep(primary, +1)} data-role="layer-up" />
        <Button icon="layerDown" title={t('Send backward: one place down the stack')} disabled={!primary} onClick={() => onStep(primary, -1)} data-role="layer-down" />
        <Button icon="groupObj" title={t('Group: the chosen layers move, show and fade as one')} disabled={!selected.length} onClick={() => onGroup(selected)} data-role="layer-group" />
        <Button icon="ungroupObj" title={t('Ungroup: the layers come out of the group')} disabled={!layer || !isGroup(layer)} onClick={() => onUngroup(primary)} data-role="layer-ungroup" />
        <Button icon="duplicate" title={t('Duplicate the layer')} disabled={!primary} onClick={() => onDuplicate(primary)} data-role="layer-duplicate" />
        <Button icon="trash" title={t('Delete the layer')} disabled={!selected.length} onClick={() => onRemove(selected)} data-role="layer-delete" />
      </div>
      {layer ? <LayerProperties layer={layer} onChange={onChange} onMask={onMask} canvasWidth={doc.size?.width || 1} hasSelection={hasSelection} canApplyMask={canApplyMask ? canApplyMask(layer) : false} /> : <div className="rw-hint im-layer-hint">{t('Choose a layer to set its opacity, blend and look.')}</div>}
    </div>
  );
}

function Thumb({ patch, assets }) {
  const ref = useRef(null);
  const asset = patch.type === 'image' ? assets?.[patch.url] : null;
  const version = asset?.version || 0;
  useEffect(() => {
    const c = ref.current;
    if (!c || !asset?.img) return;
    const ctx = c.getContext('2d');
    ctx.clearRect(0, 0, c.width, c.height);
    const k = Math.min(c.width / asset.width, c.height / asset.height);
    const w = asset.width * k;
    const h = asset.height * k;
    ctx.drawImage(asset.img, (c.width - w) / 2, (c.height - h) / 2, w, h);
  }, [asset, version, patch.rev]);
  if (asset?.img) return <canvas ref={ref} className="im-layer-thumb pic" width={36} height={26} />;
  if (patch.type === 'shape') return <span className="im-layer-thumb" data-chip="1"><span className="im-layer-block" style={{ background: colourOf(patch) }} /></span>;
  const swatch = patch.type === 'icon' || patch.type === 'wordart' || patch.type === 'textblock' ? colourOf(patch) : null;
  return (
    <span className="im-layer-thumb" style={swatch ? { '--chip': swatch } : undefined} data-chip={swatch ? '1' : '0'}>
      <Icon name={TYPE_ICON[patch.type] || 'layers'} size={14} />
    </span>
  );
}

/** A layer's painted mask, small: white where it shows the layer, black where it hides it. */
function MaskThumb({ mask, asset }) {
  const ref = useRef(null);
  const version = asset?.version || 0;
  useEffect(() => {
    const c = ref.current;
    if (!c) return;
    const g = c.getContext('2d');
    g.setTransform(1, 0, 0, 1, 0, 0);
    g.globalCompositeOperation = 'source-over';
    g.fillStyle = mask.outside === 0 ? '#000000' : '#ffffff';
    g.fillRect(0, 0, c.width, c.height);
    if (asset?.img) {
      const x = (mask.fx ?? 0) * c.width;
      const y = (mask.fy ?? 0) * c.height;
      const w = (mask.fw ?? 1) * c.width;
      const h = (mask.fh ?? 1) * c.height;
      g.fillStyle = '#000000';
      g.fillRect(x, y, w, h);
      g.drawImage(asset.img, x, y, w, h);
    }
    if (mask.invert) {
      g.globalCompositeOperation = 'difference';
      g.fillStyle = '#ffffff';
      g.fillRect(0, 0, c.width, c.height);
    }
  }, [asset, version, mask.fx, mask.fy, mask.fw, mask.fh, mask.invert, mask.outside, mask.rev]);
  return <canvas ref={ref} className={`im-layer-thumb mask${mask.off ? ' off' : ''}`} width={36} height={26} data-mask-off={mask.off ? '1' : '0'} />;
}

/** The chosen layer's mask: add one, or set its feather, turn it over, turn it off, apply it, delete it. */
function MaskProperties({ layer, onMask, canvasWidth, hasSelection, canApply }) {
  const m = layer.paintMask;
  if (!m) {
    return (
      <div className="im-mask-add" data-role="mask-add">
        <div className="im-props-sub"><Icon name="mask" size={13} /><span>{t('Add a mask')}</span></div>
        <span className="im-mask-kinds">
          <Button label={t('Reveal All')} title={t('Reveal All: a white mask, to paint black where the layer should not show')} onClick={() => onMask(layer.id, 'add', 'reveal')} data-role="mask-reveal" />
          <Button label={t('Hide All')} title={t('Hide All: a black mask, to paint white where the layer should show')} onClick={() => onMask(layer.id, 'add', 'hide')} data-role="mask-hide" />
          <Button label={t('From Selection')} title={t('From Selection: the layer shows only where the selection is')} disabled={!hasSelection} onClick={() => onMask(layer.id, 'add', 'selection')} data-role="mask-selection" />
        </span>
      </div>
    );
  }
  const px = Math.round((m.feather || 0) * canvasWidth);
  return (
    <div className="im-mask-props" data-role="mask-props">
      <div className="im-props-sub"><Icon name="mask" size={13} /><span>{t('Mask')}</span></div>
      <label className="im-prop" data-prop="feather">
        <span className="im-prop-label">{t('Feather')}</span>
        <span className="im-prop-range">
          <input type="range" min={0} max={120} value={px} data-role="mask-feather" onChange={(e) => onMask(layer.id, 'feather', Number(e.target.value) / canvasWidth)} />
          <span className="im-ps-value">{`${px} ${t('px')}`}</span>
        </span>
      </label>
      <div className="im-mask-checks">
        <label className="im-check"><input type="checkbox" checked={Boolean(m.invert)} data-role="mask-invert" onChange={(e) => onMask(layer.id, 'invert', e.target.checked)} /><span>{t('Invert')}</span></label>
        <label className="im-check"><input type="checkbox" checked={Boolean(m.off)} data-role="mask-off" onChange={(e) => onMask(layer.id, 'off', e.target.checked)} /><span>{t('Turn off')}</span></label>
      </div>
      <div className="im-mask-actions">
        <Button label={t('Apply')} title={t('Apply: the mask made part of the layer\'s pixels')} disabled={!canApply} onClick={() => onMask(layer.id, 'apply')} data-role="mask-apply" />
        <Button icon="trash" label={t('Delete')} title={t('Delete the mask: the whole layer shows again')} onClick={() => onMask(layer.id, 'remove')} data-role="mask-delete" />
      </div>
    </div>
  );
}

/** The chosen layer's own settings. */
function LayerProperties({ layer, onChange, onMask = null, canvasWidth = 1, hasSelection = false, canApplyMask = false }) {
  const id = layer.id;
  const group = isGroup(layer);
  const opacity = Math.round((layer.opacity == null ? 1 : layer.opacity) * 100);
  const [chartText, setChartText] = useState(() => (layer.type === 'chart' ? dataToText(layer) : ''));
  useEffect(() => { if (layer.type === 'chart') setChartText(dataToText(layer)); }, [id]); // eslint-disable-line react-hooks/exhaustive-deps

  const row = (label, control, role) => (
    <label className="im-prop" data-prop={role}>
      <span className="im-prop-label">{label}</span>
      {control}
    </label>
  );
  const colours = (
    <div className="im-prop-colours" data-prop="colour">
      {COLOURS.map((c) => (
        <button key={c} type="button" className="im-swatch small" style={{ background: c }} aria-pressed={colourOf(layer).toLowerCase() === c} aria-label={c} data-swatch={c} onClick={() => onChange(id, colourChange(layer, c), t('Colour'))} />
      ))}
      <label className="im-swatch small more" data-tip={t('More colours')}>
        <input type="color" value={/^#[0-9a-f]{6}$/i.test(colourOf(layer)) ? colourOf(layer) : '#ffffff'} onChange={(e) => onChange(id, colourChange(layer, e.target.value), t('Colour'))} aria-label={t('More colours')} />
      </label>
    </div>
  );

  return (
    <div className="im-props" data-props={layer.type}>
      <div className="im-props-head">{nameOf(layer)}</div>
      {layer.quick ? <div className="rw-hint" data-role="adjusted-note">{t('The Adjust and Filters tabs change this layer while its selection stays.')}</div> : null}
      {row(t('Opacity'), (
        <span className="im-prop-range">
          <input type="range" min={0} max={100} value={opacity} data-role="opacity" onChange={(e) => onChange(id, { opacity: Number(e.target.value) / 100 }, t('Opacity'), 'opacity')} />
          <span className="im-ps-value">{`${opacity}%`}</span>
        </span>
      ), 'opacity')}
      {!group ? row(t('Blend'), (
        <Select value={layer.blend || 'normal'} data-role="blend" onChange={(e) => onChange(id, { blend: e.target.value }, t('Blend'))}>
          {BLEND_MODES.map((m) => <option key={m} value={m}>{BLEND_NAMES[m] || m}</option>)}
        </Select>
      ), 'blend') : null}
      {layer.type === 'textblock' || layer.type === 'text' || layer.type === 'wordart' ? (
        <>
          {row(t('Words'), (
            <textarea className="rw-input im-prop-text" rows={3} value={layer.text || ''} data-role="layer-text" onChange={(e) => onChange(id, { text: e.target.value }, t('Words'), 'text')} />
          ), 'text')}
          {layer.type === 'wordart' ? row(t('Style'), (
            <Select value={layer.preset || 'plain'} data-role="wordart-preset" onChange={(e) => onChange(id, { preset: e.target.value, style: undefined }, t('Style'))}>
              {Object.keys(WORDART_PRESETS).map((k) => <option key={k} value={k}>{WORDART_NAMES[k] || k}</option>)}
            </Select>
          ), 'preset') : null}
          {layer.type === 'textblock' ? row(t('Align'), (
            <span className="im-prop-seg">
              {[['left', 'alignLeft', t('Align left')], ['center', 'alignCenter', t('Centre')], ['right', 'alignRight', t('Align right')]].map(([v, icon, tip]) => (
                <Button key={v} icon={icon} title={tip} pressed={(layer.align || 'left') === v} onClick={() => onChange(id, { align: v }, t('Align'))} />
              ))}
            </span>
          ), 'align') : null}
          {row(t('Colour'), colours, 'colour')}
        </>
      ) : null}
      {layer.type === 'shape' ? (
        <>
          {row(t('Shape'), (
            <Select value={layer.geometry || 'rect'} data-role="geometry" onChange={(e) => onChange(id, { geometry: e.target.value }, t('Shape'))}>
              {GEOMETRIES.map((g) => <option key={g} value={g}>{SHAPE_NAMES[g] || g}</option>)}
            </Select>
          ), 'geometry')}
          {row(t('Fill'), colours, 'colour')}
        </>
      ) : null}
      {layer.type === 'icon' ? (
        <>
          {row(t('Icon'), (
            <Select value={layer.name || 'star'} data-role="icon-name" onChange={(e) => onChange(id, { name: e.target.value }, t('Icon'))}>
              {ICON_NAMES.map((n) => <option key={n} value={n}>{ICON_NAMES_SHOWN[n] || n}</option>)}
            </Select>
          ), 'icon')}
          {row(t('Colour'), colours, 'colour')}
        </>
      ) : null}
      {layer.type === 'chart' ? (
        <>
          {row(t('Chart'), (
            <Select value={layer.chart || 'bar'} data-role="chart-kind" onChange={(e) => onChange(id, { chart: e.target.value }, t('Chart'))}>
              {CHART_KINDS.map((k) => <option key={k} value={k}>{CHART_NAMES[k] || k}</option>)}
            </Select>
          ), 'chart')}
          {row(t('Numbers'), (
            <textarea
              className="rw-input im-prop-text"
              rows={5}
              value={chartText}
              data-role="chart-data"
              onChange={(e) => { setChartText(e.target.value); onChange(id, textToData(e.target.value), t('Numbers'), 'chart-data'); }}
            />
          ), 'data')}
          <div className="rw-hint">{t('One name and one number on each line, such as: Mon, 12')}</div>
        </>
      ) : null}
      {layer.type === 'adjust' ? (
        <>
          {row(t('Look'), (
            <Select value={layer.preset || 'none'} data-role="adjust-preset" onChange={(e) => onChange(id, { preset: e.target.value, adjust: undefined }, t('Look'))}>
              {Object.keys(ADJUST_PRESETS).map((k) => <option key={k} value={k}>{ADJUST_NAMES[k] || k}</option>)}
            </Select>
          ), 'preset')}
          {ADJUST_SLIDERS.map(([key, label, min, max, step]) => {
            const base = ADJUST_PRESETS[layer.preset]?.adjust?.[key] ?? 0;
            const v = layer.adjust?.[key] ?? base;
            return (
              <label key={key} className="im-ps" data-adjust={key}>
                <span className="im-ps-label">{label}</span>
                <input type="range" min={min} max={max} step={step} value={v} onChange={(e) => onChange(id, { adjust: { ...(layer.adjust || {}), [key]: Number(e.target.value) } }, label, `adjust-${key}`)} onDoubleClick={() => onChange(id, { adjust: { ...(layer.adjust || {}), [key]: base } }, label)} />
                <span className="im-ps-value">{Number(v).toFixed(2)}</span>
              </label>
            );
          })}
          <div className="rw-hint">{t('An adjustment changes every layer under it, and none above it.')}</div>
        </>
      ) : null}
      {!group && onMask ? <MaskProperties layer={layer} onMask={onMask} canvasWidth={canvasWidth} hasSelection={hasSelection} canApply={canApplyMask} /> : null}
    </div>
  );
}

export const LAYERS_CSS = `
.im-side-tabs { display: flex; gap: 2px; padding: 8px 10px 6px; border-bottom: 1px solid var(--line-soft); }
.im-side-tabs button { flex: 1; border: 0; background: none; font: inherit; font-size: 12px; font-weight: 600; color: var(--ink-2); padding: 5px 8px; border-radius: var(--r-1); cursor: pointer; }
.im-side-tabs button:hover { background: var(--hover); color: var(--ink); }
.im-side-tabs button[aria-selected='true'] { background: var(--selected); color: var(--accent-text); }
.im-layers { display: flex; flex-direction: column; padding: 8px 10px 16px; gap: 8px; }
.im-layer-list { display: flex; flex-direction: column; border: 1px solid var(--line); border-radius: var(--r-2); background: var(--surface); overflow: hidden; min-height: 40px; }
.im-layer { position: relative; display: flex; align-items: center; gap: 6px; height: 36px; padding: 0 6px; border-bottom: 1px solid var(--line-soft); cursor: default; user-select: none; font-size: 12px; color: var(--ink); }
.im-layer:last-child { border-bottom: 0; }
.im-layer:hover { background: var(--hover); }
.im-layer.on { background: var(--selected); }
.im-layer.deep { padding-inline-start: 22px; }
.im-layer.hidden .im-layer-name, .im-layer.hidden .im-layer-thumb { opacity: 0.45; }
.im-layer.lifted { opacity: 0.55; }
.im-layer.drop-above::before, .im-layer.drop-below::after { content: ''; position: absolute; inset-inline: 4px; height: 2px; background: var(--accent); border-radius: 1px; z-index: 2; }
.im-layer.drop-above::before { top: -1px; }
.im-layer.drop-below::after { bottom: -1px; }
.im-layer.drop-into { box-shadow: inset 0 0 0 2px var(--accent); }
.im-layer-eye, .im-layer-lock, .im-layer-fold { flex: none; width: 22px; height: 22px; display: grid; place-items: center; border: 0; background: none; color: var(--ink-2); border-radius: var(--r-1); cursor: pointer; padding: 0; }
.im-layer-eye:hover, .im-layer-lock:hover, .im-layer-fold:hover { background: var(--hover); color: var(--ink); }
.im-layer-fold { width: 16px; margin-inline-end: -4px; }
.im-layer-lock { color: var(--ink-3); opacity: 0; }
.im-layer:hover .im-layer-lock, .im-layer-lock.on { opacity: 1; }
.im-layer-lock.on { color: var(--ink-2); }
.im-layer-thumb { flex: none; width: 36px; height: 26px; border-radius: 3px; display: grid; place-items: center; color: var(--ink-2);
  background: repeating-conic-gradient(#e9e9ec 0% 25%, #ffffff 0% 50%) 50% / 8px 8px; border: 1px solid var(--line); }
.im-layer-thumb[data-chip='1'] { color: var(--chip); background: #3b3f47; }
.im-layer-thumb.pic { padding: 0; }
.im-layer-block { width: 18px; height: 13px; border-radius: 2px; box-shadow: inset 0 0 0 1px rgba(255,255,255,0.35); }
.im-layer-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.im-layer[data-type='group'] .im-layer-name { font-weight: 600; }
.im-layer-meta { color: var(--ink-3); font-size: 11px; font-variant-numeric: tabular-nums; }
.im-layer-rename { flex: 1; min-width: 0; font: inherit; font-size: 12px; padding: 2px 4px; border: 1px solid var(--accent); border-radius: 3px; background: var(--surface); color: var(--ink); }
.im-layer-tools { display: flex; gap: 2px; justify-content: space-between; }
.im-layer-tools .rw-btn.icon { width: 30px; height: 26px; }
.im-layer-hint { padding: 4px 2px; }
.im-props { display: flex; flex-direction: column; gap: 9px; padding-top: 10px; border-top: 1px solid var(--line-soft); }
.im-props-head { font-size: 12px; font-weight: 600; color: var(--ink); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.im-prop { display: grid; grid-template-columns: 64px 1fr; align-items: center; gap: 8px; font-size: 12px; }
.im-prop-label { color: var(--ink-2); }
.im-prop .rw-select { width: 100%; min-width: 0; }
.im-prop-range { display: grid; grid-template-columns: 1fr 36px; align-items: center; gap: 6px; }
.im-prop-range input[type=range] { -webkit-appearance: none; appearance: none; width: 100%; height: 16px; margin: 0; background: transparent; }
.im-prop-range input[type=range]::-webkit-slider-runnable-track { height: 3px; border-radius: 2px; background: var(--line); }
.im-prop-range input[type=range]::-webkit-slider-thumb { -webkit-appearance: none; width: 13px; height: 13px; margin-top: -5px; border-radius: 50%; background: var(--bg, #fff); border: 2px solid var(--accent); }
.im-prop-text { width: 100%; box-sizing: border-box; resize: vertical; font: inherit; font-size: 12px; min-height: 52px; line-height: 1.4; }
.im-prop-colours { display: flex; flex-wrap: wrap; gap: 4px; }
.im-swatch.small { width: 18px; height: 18px; }
.im-prop-seg { display: flex; gap: 2px; }
.im-prop-seg .rw-btn.icon { width: 28px; height: 24px; }
.im-layer-thumbs { display: inline-flex; gap: 3px; flex: none; margin-inline-end: 2px; }
.im-thumb-hit { display: inline-flex; border-radius: 4px; outline: 2px solid transparent; outline-offset: 1px; cursor: pointer; }
.im-layer-thumbs.target-pixels .im-thumb-hit.pixels, .im-layer-thumbs.target-mask .im-thumb-hit.mask { outline-color: var(--accent); }
.im-layer-thumb.mask { background: #000; }
.im-layer-thumb.mask.off { opacity: 0.45; box-shadow: inset 0 0 0 1px #d23c3c; }
.im-mask-add { display: flex; flex-direction: column; gap: 7px; padding-top: 8px; border-top: 1px solid var(--line-soft); font-size: 12px; }
.im-mask-kinds { display: grid; grid-template-columns: 1fr 1fr 1.4fr; gap: 4px; }
.im-mask-kinds .rw-btn, .im-mask-actions .rw-btn { height: 26px; padding: 0 8px; font-size: 11.5px; border: 1px solid var(--line); background: var(--surface); justify-content: center; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.im-mask-kinds .rw-btn:hover:not(:disabled), .im-mask-actions .rw-btn:hover:not(:disabled) { border-color: var(--accent); }
.im-mask-kinds .rw-btn { padding: 0 4px; }
.im-mask-props { display: flex; flex-direction: column; gap: 8px; padding-top: 8px; border-top: 1px solid var(--line-soft); }
.im-props-sub { display: flex; align-items: center; gap: 6px; font-size: 12px; font-weight: 600; color: var(--ink); }
.im-mask-checks { display: flex; gap: 16px; padding-inline-start: 72px; }
.im-mask-actions { display: flex; gap: 6px; padding-inline-start: 72px; }
`;
