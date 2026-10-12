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
  { id: 'move', icon: 'move', key: 'v', label: t('Move'), tip: t('Move (V): choose a layer on the picture, drag it, or resize it by its corners') },
  '-',
  { id: 'marquee', icon: 'marquee', key: 'm', label: t('Rectangle Select'), tip: t('Rectangle Select (M): drag a box to select; Shift adds to the selection, Alt takes away') },
  { id: 'ellipse', icon: 'marqueeEllipse', label: t('Ellipse Select'), tip: t('Ellipse Select: drag an oval to select; Shift adds to the selection, Alt takes away') },
  { id: 'lasso', icon: 'lasso', key: 'l', label: t('Lasso'), tip: t('Lasso (L): draw round what to select') },
  { id: 'wand', icon: 'wand', key: 'w', label: t('Magic Wand'), tip: t('Magic Wand (W): click a colour to select the area of it') },
  { id: 'colour', icon: 'eyedropper', label: t('Select by Colour'), tip: t('Select by Colour: click a colour to select it everywhere in the picture') },
  '-',
  { id: 'brush', icon: 'brush', key: 'b', label: t('Brush'), tip: t('Brush (B): paint on a layer, or on its mask in black and white') },
  { id: 'eraser', icon: 'eraser', key: 'e', label: t('Eraser'), tip: t('Eraser (E): rub out a layer\'s pixels, or hide them with a mask') },
  { id: 'bucket', icon: 'bucket', key: 'g', label: t('Fill'), tip: t('Fill (G): click to fill an area of one colour') },
  { id: 'gradient', icon: 'gradient', label: t('Gradient'), tip: t('Gradient: drag to blend one colour into another') },
  '-',
  { id: 'clone', icon: 'cloneStamp', key: 's', label: t('Clone Stamp'), tip: t('Clone Stamp (S): Alt+click where to copy from, then paint the copy') },
  { id: 'heal', icon: 'heal', key: 'j', label: t('Healing Brush'), tip: t('Healing Brush (J): Alt+click clean skin, then paint over a blemish') },
  '-',
  { id: 'transform', icon: 'transform', key: 't', label: t('Free Transform'), tip: t('Free Transform (T): drag a picture\'s corners to scale, skew or distort it') },
  { id: 'perspective', icon: 'perspective', label: t('Perspective'), tip: t('Perspective: set four pins on something that should be square, and straighten it') },
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
    <div className="im-tools" role="toolbar" aria-orientation="vertical" aria-label={t('Tools')}>
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
        <label className="im-colour fore" data-tip={t('Colour: what the brush, the fill and the gradient paint with')}>
          <span style={{ background: options.colour }} />
          <input type="color" value={options.colour} onChange={(e) => onOptions({ colour: e.target.value })} aria-label={t('Colour')} data-role="colour" />
        </label>
        <label className="im-colour back" data-tip={t('Second colour: where a gradient ends')}>
          <span style={{ background: options.colour2 }} />
          <input type="color" value={options.colour2} onChange={(e) => onOptions({ colour2: e.target.value })} aria-label={t('Second colour')} data-role="colour2" />
        </label>
        <button type="button" className="im-colour-swap" data-tip={t('Swap the colours (X)')} aria-label={t('Swap the colours (X)')} onClick={() => onOptions({ colour: options.colour2, colour2: options.colour })}>
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
    { label: t('Feather'), title: t('Feather: soften the selection\'s edge by that many pixels'), run: () => act.reshape('feather') },
    { label: t('Grow'), title: t('Grow: widen the selection by that many pixels'), run: () => act.reshape('grow') },
    { label: t('Shrink'), title: t('Shrink: narrow the selection by that many pixels'), run: () => act.reshape('shrink') },
  ];
  const selectionActions = (
    <>
      <Sep />
      <Button icon="selectAll" title={t('Select All (Ctrl+A)')} onClick={act.selectAll} data-role="select-all" />
      <Button icon="deselect" title={t('Deselect (Ctrl+D)')} disabled={!state.selection} onClick={act.deselect} data-role="deselect" />
      <Button icon="invertSelection" title={t('Invert the selection (Ctrl+Shift+I)')} disabled={!state.selection} onClick={act.invert} data-role="invert-selection" />
      <Sep />
      <Button label={t('Modify')} title={t('Modify: feather, grow or shrink the selection by that many pixels')} disabled={!state.selection} onMouseDown={(e) => e.stopPropagation()} onClick={(e) => (modify ? setModify(null) : openModify(e))} aria-haspopup="menu" aria-expanded={Boolean(modify)} data-role="modify-selection">
        <Icon name="chevronDown" size={12} />
      </Button>
      <label className="im-opt" data-opt="edge">
        <span className="im-opt-label">{t('By')}</span>
        <input className="rw-input im-opt-num" type="number" min={1} max={200} value={o.edge} data-role="opt-edge" onChange={(e) => set({ edge: Math.max(1, Math.min(200, Number(e.target.value) || 1)) })} />
        <span className="im-opt-value">{t('px')}</span>
      </label>
      {modify && state.selection ? <Menu x={modify.x} y={modify.y} items={modifyItems} onClose={() => setModify(null)} /> : null}
      <Sep />
      <Button icon="duplicate" title={t('Copy to Layer (Ctrl+J): what is selected, as a layer of its own')} disabled={!state.selection} onClick={act.copyToLayer} data-role="copy-to-layer" />
      <Button icon="bucket" title={t('Fill: the selection painted in the colour')} disabled={!state.selection} onClick={act.fillSelection} data-role="fill-selection" />
      <Button icon="trash" title={t('Delete (Del): what is selected cleared from the layer')} disabled={!state.selection} onClick={act.clearSelection} data-role="clear-selection" />
    </>
  );
  const brushOptions = (
    <>
      <Range label={t('Size')} value={o.size} min={1} max={400} unit={` ${t('px')}`} role="size" onChange={(v) => set({ size: v })} />
      <Range label={t('Hardness')} value={o.hardness} min={0} max={100} unit="%" role="hardness" onChange={(v) => set({ hardness: v })} width={64} />
      <Range label={t('Opacity')} value={o.opacity} min={1} max={100} unit="%" role="opacity" onChange={(v) => set({ opacity: v })} width={64} />
      <Check label={t('Pen pressure')} checked={o.pressure} role="pressure" tip={t('Pen pressure: a pen pressed lightly paints a thinner line')} onChange={(v) => set({ pressure: v })} />
    </>
  );
  const where = state.target ? <><Sep /><span className="im-opt-note" data-role="paint-target">{state.target}</span></> : null;
  let body;
  if (tool === 'move' || !tool) {
    body = <span className="im-opt-note">{t('Click a layer on the picture to choose it; drag it to move it, or its corners to resize it.')}</span>;
  } else if (tool === 'marquee' || tool === 'ellipse' || tool === 'lasso' || tool === 'wand' || tool === 'colour') {
    body = (
      <>
        <Segments
          value={o.mode}
          role="mode"
          onChange={(v) => set({ mode: v })}
          items={[
            ['replace', t('New'), t('New: each selection replaces the last')],
            ['add', t('Add'), t('Add: each selection joins the last (Shift)')],
            ['subtract', t('Subtract'), t('Subtract: each selection is taken out of the last (Alt)')],
            ['intersect', t('Intersect'), t('Intersect: only what both selections hold')],
          ]}
        />
        {tool === 'wand' || tool === 'colour' ? (
          <>
            <Sep />
            <Range label={t('Tolerance')} value={o.tolerance} min={0} max={255} role="tolerance" onChange={(v) => set({ tolerance: v })} width={64} />
            {tool === 'wand' ? <Check label={t('Contiguous')} checked={o.contiguous} role="contiguous" tip={t('Contiguous: only the area joined to where you click')} onChange={(v) => set({ contiguous: v })} /> : null}
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
        <Check label={t('Aligned')} checked={o.aligned} role="aligned" tip={t('Aligned: every stroke keeps the same distance from where it copies')} onChange={(v) => set({ aligned: v })} />
        <Sep />
        <span className="im-opt-note" data-role="clone-source">{state.source ? t('Copying from {x}, {y}', { x: Math.round(state.source.x), y: Math.round(state.source.y) }) : t('Alt+click where to copy from')}</span>
        {where}
      </>
    );
  } else if (tool === 'bucket') {
    body = (
      <>
        <Range label={t('Tolerance')} value={o.tolerance} min={0} max={255} role="tolerance" onChange={(v) => set({ tolerance: v })} />
        <Check label={t('Contiguous')} checked={o.contiguous} role="contiguous" tip={t('Contiguous: only the area joined to where you click')} onChange={(v) => set({ contiguous: v })} />
        <Range label={t('Opacity')} value={o.opacity} min={1} max={100} unit="%" role="opacity" onChange={(v) => set({ opacity: v })} width={64} />
        {where}
      </>
    );
  } else if (tool === 'gradient') {
    body = (
      <>
        <Segments value={o.gradient} role="gradient" onChange={(v) => set({ gradient: v })} items={[['linear', t('Linear')], ['radial', t('Radial')], ['reflected', t('Reflected')]]} />
        <Select value={o.gradientTo} data-role="opt-gradient-to" data-opt="gradient-to" onChange={(e) => set({ gradientTo: e.target.value })} style={{ width: 168 }}>
          <option value="clear">{t('Colour to clear')}</option>
          <option value="second">{t('Colour to second colour')}</option>
        </Select>
        <Range label={t('Opacity')} value={o.opacity} min={1} max={100} unit="%" role="opacity" onChange={(v) => set({ opacity: v })} width={64} />
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
            ['scale', t('Scale'), t('Scale: a corner grows or shrinks the picture from the opposite corner; Shift keeps its shape')],
            ['skew', t('Skew'), t('Skew: a corner slides its edge along, slanting the picture')],
            ['distort', t('Distort'), t('Distort: each corner goes where it is dragged')],
            ['perspective', t('Perspective'), t('Perspective: a corner and its neighbour move apart, as if seen at an angle')],
          ]}
        />
        <Sep />
        <Button icon="refresh" label={t('Reset')} title={t('Reset: the picture back to its own rectangle')} disabled={!state.transformed} onClick={act.resetTransform} data-role="reset-transform" />
        <Button icon="check" label={t('Done')} title={t('Done: put the tool down (Enter)')} onClick={act.done} data-role="transform-done" />
      </>
    ) : <span className="im-opt-note">{t('Choose a picture or a paint layer to transform it.')}</span>;
  } else if (tool === 'perspective') {
    body = state.canTransform ? (
      <>
        <span className="im-opt-note">{t('Drag the four pins onto the corners of something that should be a rectangle.')}</span>
        <Sep />
        <Button icon="perspective" label={t('Straighten')} title={t('Straighten: the four pins made a rectangle, the picture with them')} onClick={() => act.straighten(false)} data-role="straighten-perspective" />
        <Button icon="crop" label={t('Straighten and Crop')} title={t('Straighten and Crop: straightened, and the canvas cut to the rectangle')} onClick={() => act.straighten(true)} data-role="straighten-crop" />
        <Button icon="refresh" title={t('Put the pins back')} onClick={act.resetPins} data-role="reset-pins" />
      </>
    ) : <span className="im-opt-note">{t('Choose a picture or a paint layer to straighten it.')}</span>;
  }
  return (
    <div className="im-options" role="toolbar" aria-label={t('Tool options')} data-tool={tool || 'move'}>
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
.im-tool[aria-pressed='true'] { background: var(--selected); color: var(--accent); border-color: var(--accent-line, var(--line)); }
.im-tool:disabled { opacity: 0.4; cursor: default; }
.im-tool:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.im-tools-gap { flex: none; width: 22px; height: 1px; background: var(--line-soft); margin: 4px 0; }
.im-tools-fill { flex: 1; min-height: 8px; }
.im-colours { position: relative; width: 36px; height: 36px; flex: none; }
.im-colour { position: absolute; width: 22px; height: 22px; border-radius: 4px; border: 1px solid rgba(0,0,0,0.35); box-shadow: 0 0 0 1.5px var(--surface); cursor: pointer; overflow: hidden; }
.im-colour span { position: absolute; inset: 0; }
.im-colour input { position: absolute; inset: 0; opacity: 0; width: 100%; height: 100%; cursor: pointer; border: 0; padding: 0; }
.im-colour.fore { left: 0; top: 0; z-index: 2; }
.im-colour.back { right: 0; bottom: 0; z-index: 1; }
.im-colour-swap { position: absolute; right: -2px; top: -3px; width: 14px; height: 14px; padding: 0; border: 0; background: none; color: var(--ink-3); cursor: pointer; display: grid; place-items: center; }
.im-colour-swap:hover { color: var(--accent); }
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
.im-opt-seg button[aria-checked='true'] { background: var(--selected); color: var(--accent); font-weight: 600; }
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
