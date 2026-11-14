// The advanced editor's tools (docs/STUDIO.md 7.2, I10 to I14), placed as
// photo editors place them: a column of tools down the left of the picture
// (move; the selections; the brush, eraser, fill and gradient; the clone
// stamp and the healing brush; free transform and perspective) with the two
// colours at its foot, and a bar above the picture with the chosen tool's
// options and what it can do next.

import React, { useState } from 'react';
import { Button, Icon, Menu, Select, t } from '@rutba/office-ui';

/** The tools, in the column's order; '-' is a gap between kinds. */
export const TOOLS = [
  { id: 'move', icon: 'move', key: 'v', label: t('common.move'), tip: t('image.tools.move-v-choose') },
  '-',
  { id: 'marquee', icon: 'marquee', key: 'm', label: t('image.tools.rectangle-select'), tip: t('image.tools.rectangle-select-m') },
  { id: 'ellipse', icon: 'marqueeEllipse', label: t('image.tools.ellipse-select'), tip: t('image.tools.ellipse-select-drag') },
  { id: 'lasso', icon: 'lasso', key: 'l', label: t('common.lasso'), tip: t('image.tools.lasso-l-draw') },
  { id: 'wand', icon: 'wand', key: 'w', label: t('image.tools.magic-wand'), tip: t('image.tools.magic-wand-w') },
  { id: 'colour', icon: 'eyedropper', label: t('image.tools.select-by-colour'), tip: t('image.tools.select-colour-click') },
  '-',
  { id: 'brush', icon: 'brush', key: 'b', label: t('image.tools.brush'), tip: t('image.tools.brush-b-paint') },
  { id: 'eraser', icon: 'eraser', key: 'e', label: t('common.eraser'), tip: t('image.tools.eraser-e-rub') },
  { id: 'bucket', icon: 'bucket', key: 'g', label: t('common.fill'), tip: t('image.tools.fill-g-click') },
  { id: 'gradient', icon: 'gradient', label: t('common.gradient'), tip: t('image.tools.gradient-drag-blend') },
  '-',
  { id: 'clone', icon: 'cloneStamp', key: 's', label: t('image.tools.clone-stamp'), tip: t('image.tools.clone-stamp-s') },
  { id: 'heal', icon: 'heal', key: 'j', label: t('image.tools.healing-brush'), tip: t('image.tools.healing-brush-j') },
  '-',
  { id: 'transform', icon: 'transform', key: 't', label: t('image.tools.free-transform'), tip: t('image.tools.free-transform-t') },
  { id: 'perspective', icon: 'perspective', label: t('image.tools.perspective'), tip: t('image.tools.perspective-set-four') },
];

export const TOOL_BY_ID = Object.fromEntries(TOOLS.filter((x) => x !== '-').map((x) => [x.id, x]));
export const SELECT_TOOLS = new Set(['marquee', 'ellipse', 'lasso', 'wand', 'colour']);
export const PAINT_TOOLS = new Set(['brush', 'eraser', 'clone', 'heal']);

export const DEFAULT_OPTIONS = {
  mode: 'replace',
  edge: 4,
  tolerance: 32,
  contiguous: true,
  size: 30,
  hardness: 80,
  opacity: 100,
  pressure: true,
  colour: '#e0393e',
  colour2: '#ffffff',
  gradient: 'linear',
  gradientTo: 'clear',
  aligned: true,
  transform: 'scale',
};

/** The column of tools. */
export function ToolBar({ tool, onTool, options, onOptions, disabled }) {
  return (
    <div className="im-tools" role="toolbar" aria-orientation="vertical" aria-label={t('common.tools')}>
      {TOOLS.map((x, i) => (x === '-' ? <span key={`gap-${i}`} className="im-tools-gap" /> : (
        <button
          key={x.id}
          type="button"
          className="im-tool"
          data-tool-id={x.id}
          aria-pressed={tool === x.id}
          aria-label={x.label}
          data-tip={x.tip}
          disabled={disabled}
          onClick={() => onTool(x.id)}
        >
          <Icon name={x.icon} size={18} />
        </button>
      )))}
      <span className="im-tools-fill" />
      <div className="im-colours" data-role="colours">
        <label className="im-colour fore" data-tip={t('image.tools.colour-brush-fill')}>
          <span style={{ background: options.colour }} />
          <input type="color" value={options.colour} onChange={(e) => onOptions({ colour: e.target.value })} aria-label={t('common.colour')} data-role="colour" />
        </label>
        <label className="im-colour back" data-tip={t('image.tools.second-colour-gradient')}>
          <span style={{ background: options.colour2 }} />
          <input type="color" value={options.colour2} onChange={(e) => onOptions({ colour2: e.target.value })} aria-label={t('image.tools.second-colour')} data-role="colour2" />
        </label>
        <button type="button" className="im-colour-swap" data-tip={t('image.tools.swap-colours-x')} aria-label={t('image.tools.swap-colours-x')} onClick={() => onOptions({ colour: options.colour2, colour2: options.colour })}>
          <Icon name="swap" size={12} />
        </button>
      </div>
    </div>
  );
}

/** A labelled slider of the options bar. */
function Range({ label, value, min, max, step = 1, unit = '', role, onChange, width = 84 }) {
  return (
    <label className="im-opt im-opt-range" data-opt={role}>
      <span className="im-opt-label">{label}</span>
      <input type="range" min={min} max={max} step={step} value={value} data-role={`opt-${role}`} style={{ width }} onChange={(e) => onChange(Number(e.target.value))} />
      <span className="im-opt-value">{`${value}${unit}`}</span>
    </label>
  );
}

/** Buttons of which one is chosen. */
function Segments({ value, items, role, onChange }) {
  return (
    <span className="im-opt-seg" role="radiogroup" data-opt={role}>
      {items.map(([v, label, tip]) => (
        <button key={v} type="button" role="radio" aria-checked={value === v} data-value={v} data-tip={tip || label} onClick={() => onChange(v)}>{label}</button>
      ))}
    </span>
  );
}

const Check = ({ label, checked, role, onChange, tip }) => (
  <label className="im-opt im-opt-check" data-opt={role} data-tip={tip}>
    <input type="checkbox" checked={checked} data-role={`opt-${role}`} onChange={(e) => onChange(e.target.checked)} />
    <span>{label}</span>
  </label>
);

const Sep = () => <span className="im-opt-sep" />;

/**
 * The bar above the picture: what the chosen tool does and how. `act` holds
 * the actions it offers (select all, deselect, fill, straighten, …), `state`
 * what the window knows (a selection, where paint is going, the clone's
 * source, whether a layer can be transformed).
 */
export function OptionsBar({ tool, options: o, onOptions: set, act, state }) {
  const info = TOOL_BY_ID[tool] || TOOL_BY_ID.move;
  // Feather, Grow and Shrink under one Modify menu, as photo editors keep them: the bar stays one short row.
  const [modify, setModify] = useState(null);
  const openModify = (e) => {
    const r = e.currentTarget.getBoundingClientRect();
    const rtl = document.documentElement.dir === 'rtl';
    setModify({ x: rtl ? Math.max(8, r.right - 220) : r.left, y: r.bottom + 4 });
  };
  const modifyItems = [
    { label: t('image.feather'), title: t('image.tools.feather-soften'), run: () => act.reshape('feather') },
    { label: t('image.tools.grow'), title: t('image.tools.grow-widen-selection'), run: () => act.reshape('grow') },
    { label: t('image.tools.shrink'), title: t('image.tools.shrink-narrow-selection'), run: () => act.reshape('shrink') },
  ];
  const selectionActions = (
    <>
      <Sep />
      <Button icon="selectAll" title={t('image.tools.select-all-ctrl')} onClick={act.selectAll} data-role="select-all" />
      <Button icon="deselect" title={t('image.tools.deselect-ctrl-d')} disabled={!state.selection} onClick={act.deselect} data-role="deselect" />
      <Button icon="invertSelection" title={t('image.tools.invert-selection-ctrl')} disabled={!state.selection} onClick={act.invert} data-role="invert-selection" />
      <Sep />
      <Button label={t('image.tools.modify')} title={t('image.tools.modify-feather-grow')} disabled={!state.selection} onMouseDown={(e) => e.stopPropagation()} onClick={(e) => (modify ? setModify(null) : openModify(e))} aria-haspopup="menu" aria-expanded={Boolean(modify)} data-role="modify-selection">
        <Icon name="chevronDown" size={12} />
      </Button>
      <label className="im-opt" data-opt="edge">
        <span className="im-opt-label">{t('image.tools.by')}</span>
        <input className="rw-input im-opt-num" type="number" min={1} max={200} value={o.edge} data-role="opt-edge" onChange={(e) => set({ edge: Math.max(1, Math.min(200, Number(e.target.value) || 1)) })} />
        <span className="im-opt-value">{t('common.px')}</span>
      </label>
      {modify && state.selection ? <Menu x={modify.x} y={modify.y} items={modifyItems} onClose={() => setModify(null)} /> : null}
      <Sep />
      <Button icon="duplicate" title={t('image.tools.copy-layer-ctrl')} disabled={!state.selection} onClick={act.copyToLayer} data-role="copy-to-layer" />
      <Button icon="bucket" title={t('image.tools.fill-selection-painted')} disabled={!state.selection} onClick={act.fillSelection} data-role="fill-selection" />
      <Button icon="trash" title={t('image.tools.delete-del-selected')} disabled={!state.selection} onClick={act.clearSelection} data-role="clear-selection" />
    </>
  );
  const brushOptions = (
    <>
      <Range label={t('common.size')} value={o.size} min={1} max={400} unit={` ${t('common.px')}`} role="size" onChange={(v) => set({ size: v })} />
      <Range label={t('image.tools.hardness')} value={o.hardness} min={0} max={100} unit="%" role="hardness" onChange={(v) => set({ hardness: v })} width={64} />
      <Range label={t('common.opacity')} value={o.opacity} min={1} max={100} unit="%" role="opacity" onChange={(v) => set({ opacity: v })} width={64} />
      <Check label={t('image.tools.pen-pressure')} checked={o.pressure} role="pressure" tip={t('image.tools.pen-pressure-pen')} onChange={(v) => set({ pressure: v })} />
    </>
  );
  const where = state.target ? <><Sep /><span className="im-opt-note" data-role="paint-target">{state.target}</span></> : null;
  let body;
  if (tool === 'move' || !tool) {
    body = <span className="im-opt-note">{t('image.tools.click-layer-picture')}</span>;
  } else if (tool === 'marquee' || tool === 'ellipse' || tool === 'lasso' || tool === 'wand' || tool === 'colour') {
    body = (
      <>
        <Segments
          value={o.mode}
          role="mode"
          onChange={(v) => set({ mode: v })}
          items={[
            ['replace', t('common.new'), t('image.tools.new-each-selection')],
            ['add', t('common.add'), t('image.tools.add-each-selection')],
            ['subtract', t('common.subtract'), t('image.tools.subtract-each-selection')],
            ['intersect', t('image.tools.intersect'), t('image.tools.intersect-only-both')],
          ]}
        />
        {tool === 'wand' || tool === 'colour' ? (
          <>
            <Sep />
            <Range label={t('image.tools.tolerance')} value={o.tolerance} min={0} max={255} role="tolerance" onChange={(v) => set({ tolerance: v })} width={64} />
            {tool === 'wand' ? <Check label={t('image.tools.contiguous')} checked={o.contiguous} role="contiguous" tip={t('image.tools.contiguous-only-area')} onChange={(v) => set({ contiguous: v })} /> : null}
          </>
        ) : null}
        {selectionActions}
      </>
    );
  } else if (tool === 'brush' || tool === 'eraser') {
    body = <>{brushOptions}{where}</>;
  } else if (tool === 'clone' || tool === 'heal') {
    body = (
      <>
        {brushOptions}
        <Check label={t('image.tools.aligned')} checked={o.aligned} role="aligned" tip={t('image.tools.aligned-every-stroke')} onChange={(v) => set({ aligned: v })} />
        <Sep />
        <span className="im-opt-note" data-role="clone-source">{state.source ? t('image.tools.copying-x-y', { x: Math.round(state.source.x), y: Math.round(state.source.y) }) : t('image.tools.alt-click-copy')}</span>
        {where}
      </>
    );
  } else if (tool === 'bucket') {
    body = (
      <>
        <Range label={t('image.tools.tolerance')} value={o.tolerance} min={0} max={255} role="tolerance" onChange={(v) => set({ tolerance: v })} />
        <Check label={t('image.tools.contiguous')} checked={o.contiguous} role="contiguous" tip={t('image.tools.contiguous-only-area')} onChange={(v) => set({ contiguous: v })} />
        <Range label={t('common.opacity')} value={o.opacity} min={1} max={100} unit="%" role="opacity" onChange={(v) => set({ opacity: v })} width={64} />
        {where}
      </>
    );
  } else if (tool === 'gradient') {
    body = (
      <>
        <Segments value={o.gradient} role="gradient" onChange={(v) => set({ gradient: v })} items={[['linear', t('common.linear')], ['radial', t('image.tools.radial')], ['reflected', t('image.tools.reflected')]]} />
        <Select value={o.gradientTo} data-role="opt-gradient-to" data-opt="gradient-to" onChange={(e) => set({ gradientTo: e.target.value })} style={{ width: 168 }}>
          <option value="clear">{t('image.tools.colour-to-clear')}</option>
          <option value="second">{t('image.tools.colour-second-colour')}</option>
        </Select>
        <Range label={t('common.opacity')} value={o.opacity} min={1} max={100} unit="%" role="opacity" onChange={(v) => set({ opacity: v })} width={64} />
        {where}
      </>
    );
  } else if (tool === 'transform') {
    body = state.canTransform ? (
      <>
        <Segments
          value={o.transform}
          role="transform"
          onChange={(v) => set({ transform: v })}
          items={[
            ['scale', t('common.scale'), t('image.tools.scale-corner-grows')],
            ['skew', t('image.tools.skew'), t('image.tools.skew-corner-slides')],
            ['distort', t('image.tools.distort'), t('image.tools.distort-each-corner')],
            ['perspective', t('image.tools.perspective'), t('image.tools.perspective-corner')],
          ]}
        />
        <Sep />
        <Button icon="refresh" label={t('common.reset')} title={t('image.tools.reset-picture-back')} disabled={!state.transformed} onClick={act.resetTransform} data-role="reset-transform" />
        <Button icon="check" label={t('common.done')} title={t('image.tools.done-put-tool')} onClick={act.done} data-role="transform-done" />
      </>
    ) : <span className="im-opt-note">{t('image.tools.choose-picture-paint')}</span>;
  } else if (tool === 'perspective') {
    body = state.canTransform ? (
      <>
        <span className="im-opt-note">{t('image.tools.drag-four-pins')}</span>
        <Sep />
        <Button icon="perspective" label={t('image.straighten')} title={t('image.tools.straighten-four-pins')} onClick={() => act.straighten(false)} data-role="straighten-perspective" />
        <Button icon="crop" label={t('image.tools.straighten')} title={t('image.tools.straighten-crop')} onClick={() => act.straighten(true)} data-role="straighten-crop" />
        <Button icon="refresh" title={t('image.tools.put-pins-back')} onClick={act.resetPins} data-role="reset-pins" />
      </>
    ) : <span className="im-opt-note">{t('image.tools.choose-picture-paint-layer')}</span>;
  }
  return (
    <div className="im-options" role="toolbar" aria-label={t('image.tools.tool-options')} data-tool={tool || 'move'}>
      <span className="im-options-tool"><Icon name={info.icon} size={16} /><span>{info.label}</span></span>
      <Sep />
      {body}
    </div>
  );
}

export const TOOLS_CSS = `
.im-work { flex: 1; min-height: 0; min-width: 0; display: flex; }
.im-main { flex: 1; min-width: 0; min-height: 0; display: flex; flex-direction: column; container-type: inline-size; container-name: im-main; }
.im-tools { flex: none; width: 44px; display: flex; flex-direction: column; align-items: center; gap: 2px; padding: 8px 0 10px;
  background: var(--surface); border-inline-end: 1px solid var(--line-soft); overflow-y: auto; overflow-x: hidden; }
.im-tool { width: 34px; height: 32px; flex: none; display: grid; place-items: center; border: 1px solid transparent; border-radius: var(--r-2);
  background: none; color: var(--ink-2); cursor: pointer; padding: 0; }
.im-tool:hover:not(:disabled) { background: var(--hover); color: var(--ink); }
.im-tool[aria-pressed='true'] { background: var(--selected); color: var(--accent-text); border-color: var(--accent-line, var(--line)); }
.im-tool:disabled { opacity: 0.4; cursor: default; }
.im-tool:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.im-tools-gap { flex: none; width: 22px; height: 1px; background: var(--line-soft); margin: 4px 0; }
.im-tools-fill { flex: 1; min-height: 8px; }
/* A short window: the tools a little shorter, so the whole rail and its colours fit without a scroll bar. */
@media (max-height: 820px) {
  .im-tools { gap: 1px; padding: 6px 0 8px; }
  .im-tool { height: 26px; }
  .im-tools-gap { margin: 2px 0; }
}
.im-colours { position: relative; width: 36px; height: 36px; flex: none; }
.im-colour { position: absolute; width: 22px; height: 22px; border-radius: 4px; border: 1px solid rgba(0,0,0,0.35); box-shadow: 0 0 0 1.5px var(--surface); cursor: pointer; overflow: hidden; }
.im-colour span { position: absolute; inset: 0; }
.im-colour input { position: absolute; inset: 0; opacity: 0; width: 100%; height: 100%; cursor: pointer; border: 0; padding: 0; }
.im-colour.fore { left: 0; top: 0; z-index: 2; }
.im-colour.back { right: 0; bottom: 0; z-index: 1; }
.im-colour-swap { position: absolute; right: -2px; top: -3px; width: 14px; height: 14px; padding: 0; border: 0; background: none; color: var(--ink-3); cursor: pointer; display: grid; place-items: center; }
.im-colour-swap:hover { color: var(--accent-text); }
[dir='rtl'] .im-colour.fore { left: auto; right: 0; }
[dir='rtl'] .im-colour.back { right: auto; left: 0; }
[dir='rtl'] .im-colour-swap { right: auto; left: -2px; }
.im-options { flex: none; display: flex; align-items: center; gap: 6px; min-height: 38px; padding: 4px 12px; background: var(--surface);
  border-bottom: 1px solid var(--line-soft); font-size: 12px; color: var(--ink-2); overflow-x: auto; overflow-y: hidden; white-space: nowrap; }
.im-options::-webkit-scrollbar { height: 3px; }
.im-options::-webkit-scrollbar-thumb { background: var(--line); border-radius: 2px; }
.im-options .rw-btn { height: 26px; flex: none; padding: 0 7px; }
.im-options .rw-btn.icon { width: 28px; }
.im-options-tool { display: flex; align-items: center; gap: 6px; font-weight: 600; color: var(--ink); flex: none; }
.im-opt { display: inline-flex; align-items: center; gap: 6px; flex: none; }
.im-opt-label { color: var(--ink-2); }
.im-opt-value { color: var(--ink-3); font-variant-numeric: tabular-nums; min-width: 34px; }
.im-opt-num { width: 52px; height: 24px; padding: 0 6px; font-size: 12px; }
.im-opt-note { color: var(--ink-3); flex: none; }
.im-opt-sep { width: 1px; height: 20px; background: var(--line-soft); flex: none; }
.im-opt-check input { margin: 0; }
.im-opt-seg { display: inline-flex; border: 1px solid var(--line); border-radius: var(--r-2); overflow: hidden; flex: none; }
.im-opt-seg button { border: 0; background: var(--surface); color: var(--ink-2); font: inherit; font-size: 12px; padding: 3px 9px; cursor: pointer; }
.im-opt-seg button + button { border-inline-start: 1px solid var(--line); }
.im-opt-seg button:hover { background: var(--hover); color: var(--ink); }
.im-opt-seg button[aria-checked='true'] { background: var(--selected); color: var(--accent-text); font-weight: 600; }
.im-opt-range input[type=range] { -webkit-appearance: none; appearance: none; height: 16px; margin: 0; background: transparent; cursor: pointer; }
.im-opt-range input[type=range]::-webkit-slider-runnable-track { height: 3px; border-radius: 2px; background: var(--line); }
.im-opt-range input[type=range]::-webkit-slider-thumb { -webkit-appearance: none; width: 13px; height: 13px; margin-top: -5px; border-radius: 50%; background: var(--bg, #fff); border: 2px solid var(--accent); }
/* A narrow window: the tool's name goes (its icon stays, its tip names it), and the bar tightens. */
@container im-main (max-width: 1000px) {
  .im-options-tool > span { display: none; }
  .im-options { gap: 4px; padding: 4px 8px; }
  .im-opt-value { min-width: 0; }
}
`;
