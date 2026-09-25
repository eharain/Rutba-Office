// Worksheets: what floats over the cells — slicer panels, and the handles a
// picked drawing is moved and resized by.
//
// A slicer is drawn here as the window's own panel rather than a picture:
// its buttons are pressed, its header dragged, its corner pulled, and every
// press is one operation on the engine, which filters the table or pivot and
// writes the choice where Excel keeps it.

import React, { useState } from 'react';
import { Button, Dialog, Icon } from '@rutba/office-ui';

/** Excel's eight resize handles: where on the box, and the cursor each wears. */
export const HANDLES = [
  ['nw', 0, 0, 'nwse-resize'], ['n', 0.5, 0, 'ns-resize'], ['ne', 1, 0, 'nesw-resize'],
  ['e', 1, 0.5, 'ew-resize'], ['se', 1, 1, 'nwse-resize'], ['s', 0.5, 1, 'ns-resize'],
  ['sw', 0, 1, 'nesw-resize'], ['w', 0, 0.5, 'ew-resize'],
];

/** A box moved by (dx, dy), or resized from one of the handles; never smaller than 16 px. */
export function draggedBox(box, mode, handle, dx, dy) {
  if (mode === 'move') return { ...box, x: Math.max(0, Math.round(box.x + dx)), y: Math.max(0, Math.round(box.y + dy)) };
  let { x, y, width: w, height: h } = box;
  if (handle.includes('e')) w += dx;
  if (handle.includes('s')) h += dy;
  if (handle.includes('w')) { x += dx; w -= dx; }
  if (handle.includes('n')) { y += dy; h -= dy; }
  const min = 16;
  if (w < min) { if (handle.includes('w')) x -= min - w; w = min; }
  if (h < min) { if (handle.includes('n')) y -= min - h; h = min; }
  return { x: Math.max(0, Math.round(x)), y: Math.max(0, Math.round(y)), width: Math.round(w), height: Math.round(h) };
}

/**
 * Follow the pointer from a press on a drawing (move) or on one of its
 * handles (resize), in the grid's own pixels — the grid is zoomed, the
 * pointer is not. `onMove` is called with the box as it would be, `onDone`
 * once with the last one when the button comes up after a real drag.
 */
export function followPointer(e, { box, mode, handle = null, zoom = 1, onMove, onDone }) {
  const start = { x: e.clientX, y: e.clientY };
  let last = null;
  let moved = false;
  const move = (ev) => {
    const dx = (ev.clientX - start.x) / zoom;
    const dy = (ev.clientY - start.y) / zoom;
    if (!moved && Math.abs(dx) < 3 && Math.abs(dy) < 3) return;
    moved = true;
    last = draggedBox(box, mode, handle, dx, dy);
    onMove(last);
  };
  const up = () => {
    window.removeEventListener('mousemove', move, true);
    window.removeEventListener('mouseup', up, true);
    onDone(moved ? last : null);
  };
  window.addEventListener('mousemove', move, true);
  window.addEventListener('mouseup', up, true);
}

/** The ring and handles round a picked drawing. */
export function ObjectHandles({ box, onHandle }) {
  return (
    <>
      <div className="sh-obj-ring" style={{ left: box.x, top: box.y, width: box.width, height: box.height }} />
      {HANDLES.map(([name, fx, fy, cursor]) => (
        <div
          key={name}
          className="sh-obj-handle"
          data-handle={name}
          style={{ left: box.x + box.width * fx - 4.5, top: box.y + box.height * fy - 4.5, cursor }}
          onMouseDown={(e) => { if (e.button === 0) { e.preventDefault(); e.stopPropagation(); onHandle(e, name); } }}
        />
      ))}
    </>
  );
}

/**
 * A slicer panel: its caption with Multi-Select and Clear Filter, and a
 * button per item — pressed when selected, faded when it has no data under
 * the other filters. A plain click shows that item alone; Ctrl+click, or
 * any click with Multi-Select on, adds or takes away one item.
 */
export function SlicerPanel({ d, box, picked, multi, onPick, onStartMove, onToggleMulti, onChoose, onClear }) {
  const s = d.slicer || { items: [] };
  const selected = s.items.filter((i) => i.selected).map((i) => i.label);
  const press = (e, item) => {
    e.preventDefault();
    e.stopPropagation();
    if (s.broken) return;
    const adding = multi || e.ctrlKey || e.metaKey;
    let next;
    if (!adding) next = [item.label];
    else if (!s.filtered) next = s.items.filter((i) => i.label !== item.label).map((i) => i.label);
    else next = selected.includes(item.label) ? selected.filter((l) => l !== item.label) : [...selected, item.label];
    if (!next.length) return;
    if (next.length === s.items.length) onClear();
    else onChoose(next);
  };
  const rows = Math.max(18, s.rowHeightPx || 25);
  return (
    <div
      className={`sh-drawing sh-slicer${picked ? ' picked' : ''}${s.filtered ? ' filtered' : ''}`}
      data-slicer={s.name}
      style={{ left: box.x, top: box.y, width: box.width, height: box.height }}
      onMouseDown={(e) => { if (e.button === 0) { e.stopPropagation(); onPick(e); } }}
    >
      {s.showCaption !== false ? (
        <div
          className="sh-slicer-head"
          onMouseDown={(e) => { if (e.button === 0 && !e.target.closest('button')) { e.preventDefault(); e.stopPropagation(); onPick(e); onStartMove(e); } }}
          data-tip={`${s.caption || s.name} — drag to move; the corner handles resize`}
        >
          <span className="sh-slicer-caption">{s.caption || s.name}</span>
          <button
            type="button"
            className={`sh-slicer-tool sh-slicer-multi${multi ? ' on' : ''}`}
            data-tip={`Multi-Select — ${multi ? 'on: each click adds or takes away an item' : 'off: a click shows that item alone (Ctrl+click adds one)'}`}
            aria-pressed={multi}
            onMouseDown={(e) => e.stopPropagation()}
            onClick={(e) => { e.stopPropagation(); onToggleMulti(); }}
          >
            <Icon name="multiSelect" size={14} />
          </button>
          <button
            type="button"
            className="sh-slicer-tool sh-slicer-clear"
            data-tip={s.filtered ? 'Clear Filter — show every item' : 'Clear Filter — nothing is filtered'}
            disabled={!s.filtered}
            onMouseDown={(e) => e.stopPropagation()}
            onClick={(e) => { e.stopPropagation(); onClear(); }}
          >
            <Icon name="filterClear" size={14} />
          </button>
        </div>
      ) : null}
      {s.broken ? (
        <div className="sh-slicer-broken">This slicer cannot filter: {s.broken}.</div>
      ) : (
        <div className="sh-slicer-items" style={{ gridTemplateColumns: `repeat(${Math.max(1, s.columns || 1)}, minmax(0, 1fr))`, gridAutoRows: rows }}>
          {s.items.map((item) => (
            <button
              key={item.label}
              type="button"
              className={`sh-slicer-item${item.selected ? ' on' : ''}${item.hasData ? '' : ' nodata'}`}
              data-item={item.label}
              data-tip={`${item.label}${item.hasData ? '' : ' — no data under the other filters'}`}
              onMouseDown={(e) => press(e, item)}
            >
              <span>{item.label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

/**
 * Insert → Slicer: Excel's Insert Slicers dialog — a tick per field of the
 * table or pivot the cursor is in, one slicer made for each ticked.
 */
export function InsertSlicersDialog({ source, onClose, onInsert }) {
  const [ticked, setTicked] = useState(() => new Set());
  const toggle = (name) => setTicked((s) => {
    const next = new Set(s);
    if (next.has(name)) next.delete(name);
    else next.add(name);
    return next;
  });
  return (
    <Dialog
      title="Insert Slicers"
      width={340}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="OK" className="sh-slicers-ok" disabled={!ticked.size} onClick={() => onInsert([...ticked])} />
        </>
      }
    >
      <div className="sh-slicers-from">
        <Icon name={source.kind === 'pivot' ? 'table' : 'grid'} size={14} />
        <span>{source.kind === 'pivot' ? 'Pivot table' : 'Table'} <b>{source.name}</b></span>
      </div>
      <div className="sh-slicers-list">
        {source.fields.map((f) => (
          <label key={f.name} className="sh-slicers-field" data-field={f.name}>
            <input type="checkbox" checked={ticked.has(f.name)} onChange={() => toggle(f.name)} />
            <span className="grow">{f.name}</span>
            {f.has ? <span className="sh-slicers-has">has a slicer</span> : null}
          </label>
        ))}
      </div>
      <p className="rw-hint" style={{ margin: '8px 0 0' }}>Each ticked field gets a panel of buttons, one per item, that filters the {source.kind === 'pivot' ? 'pivot table' : 'table'}.</p>
    </Dialog>
  );
}

export const OBJECTS_CSS = `
/* A picked drawing: a ring and Excel's eight handles. */
.sh-obj-ring { position: absolute; z-index: 6; pointer-events: none; border: 1.5px solid var(--accent); box-sizing: border-box; border-radius: 1px; }
.sh-obj-handle { position: absolute; z-index: 7; width: 9px; height: 9px; background: #fff; border: 1.5px solid var(--accent); border-radius: 50%; box-sizing: border-box; box-shadow: 0 1px 2px rgba(15,20,30,.18); }

/* Slicers: a card with a caption, two tools and a grid of buttons. */
.sh-slicer {
  display: flex; flex-direction: column; box-sizing: border-box; overflow: hidden; cursor: default;
  background: var(--surface); border: 1px solid var(--line); border-radius: 8px;
  box-shadow: 0 1px 2px rgba(15,20,30,.06), 0 6px 18px rgba(15,20,30,.10);
  font-family: var(--font); color: var(--ink); z-index: 3;
}
.sh-slicer.picked { border-color: var(--accent-line); }
.sh-slicer-head { display: flex; align-items: center; gap: 4px; padding: 7px 6px 6px 11px; border-bottom: 1px solid var(--line-soft); cursor: move; flex: none; user-select: none; }
.sh-slicer-caption { flex: 1; min-width: 0; font-size: 12.5px; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sh-slicer-tool { width: 24px; height: 24px; display: grid; place-items: center; padding: 0; border: 1px solid transparent; border-radius: 6px; background: transparent; color: var(--ink-2); cursor: pointer; flex: none; }
.sh-slicer-tool:hover:not(:disabled) { background: var(--hover); color: var(--ink); }
.sh-slicer-tool.on { background: var(--selected); border-color: var(--accent-line); color: var(--accent); }
.sh-slicer-tool:disabled { opacity: .38; cursor: default; }
.sh-slicer.filtered .sh-slicer-clear { color: var(--accent); }
.sh-slicer-items { flex: 1; min-height: 0; overflow: auto; display: grid; gap: 5px; padding: 8px 9px 9px; align-content: start; }
.sh-slicer-item {
  display: flex; align-items: center; min-width: 0; padding: 0 9px; box-sizing: border-box;
  border: 1px solid var(--line); border-radius: 6px; background: var(--surface); color: var(--ink-2);
  font: inherit; font-size: 12px; text-align: left; cursor: pointer; transition: background var(--fast), border-color var(--fast), color var(--fast);
}
.sh-slicer-item > span { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sh-slicer-item:hover { border-color: var(--line-strong); background: var(--hover); }
.sh-slicer-item.on { background: var(--accent-soft); border-color: var(--accent-line); color: var(--ink); font-weight: 600; }
.sh-slicer-item.on:hover { border-color: var(--accent); }
.sh-slicer-item.nodata { color: var(--ink-3); border-style: dashed; background: var(--surface-2); font-weight: 400; }
.sh-slicer-item.nodata.on { background: color-mix(in srgb, var(--accent-soft) 45%, var(--surface-2)); }
.sh-slicer-broken { padding: 12px; font-size: 12px; color: var(--ink-3); }

/* PivotChart & PivotTable: the chart kinds as a row of tiles. */
.sh-pivot-kinds { display: grid; grid-template-columns: repeat(6, 1fr); gap: 6px; }
.sh-pivot-kind { display: flex; flex-direction: column; align-items: center; gap: 4px; padding: 8px 4px 6px; border: 1px solid var(--line); border-radius: 8px; background: var(--surface); color: var(--ink-2); font: inherit; font-size: 11.5px; cursor: pointer; }
.sh-pivot-kind:hover { background: var(--hover); }
.sh-pivot-kind.on { border-color: var(--accent); background: var(--selected); color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }

/* Insert → Slicer's dialog. */
.sh-slicers-from { display: flex; align-items: center; gap: 7px; font-size: 12.5px; color: var(--ink-2); margin-bottom: 8px; }
.sh-slicers-list { display: flex; flex-direction: column; gap: 1px; max-height: 280px; overflow: auto; border: 1px solid var(--line); border-radius: 8px; padding: 4px; background: var(--surface); }
.sh-slicers-field { display: flex; align-items: center; gap: 9px; padding: 6px 8px; border-radius: 6px; font-size: 12.5px; cursor: pointer; }
.sh-slicers-field:hover { background: var(--hover); }
.sh-slicers-field .grow { flex: 1; min-width: 0; }
.sh-slicers-has { font-size: 11px; color: var(--ink-3); }
`;
