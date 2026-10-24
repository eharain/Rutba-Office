// A Data Files form's Design: its fields moved and resized on the form's
// grid (dragged, or the one picked moved with the arrow keys and sized with
// Shift and them), taken off and put back, their labels renamed, and the
// layout saved with the form. The record's own values are not touched; the
// layout's arithmetic is the engine's (database/src/form-design.js).

import React, { useEffect, useRef, useState } from 'react';
import { Button, Select, t } from '@rutba/office-ui';
import { GRID, LABEL_WIDTH, fieldsOff, labelLeft, placeControl, putBack, relabel, takeOff } from '@rutba/database/form-design';

const KIND_NAMES = () => ({ text: t('Text'), longtext: t('Long text'), number: t('Number'), currency: t('Currency'), date: t('Date'), datetime: t('Date and time'), check: t('Yes/No'), list: t('List'), auto: t('AutoNumber') });

export function FormDesigner({ shell, db, layout: start, onCancel, onSaved, toast }) {
  const [layout, setLayout] = useState(start);
  const [picked, setPicked] = useState(start.controls[0]?.field || null);
  const [fields, setFields] = useState([]);
  const [busy, setBusy] = useState(false);
  const drag = useRef(null);
  useEffect(() => {
    shell.base.fields({ id: db.id, table: start.table }).then(setFields).catch(() => {});
  }, [shell, db.id, start.table]);
  const changed = JSON.stringify(layout) !== JSON.stringify(start);
  const off = fieldsOff(layout, fields);

  const apply = (make) => {
    // Worked out here, not inside React's update, so a refusal is said and not thrown while drawing.
    try { setLayout(make(layout)); } catch (e) { toast(e.message, { tone: 'bad' }); }
  };
  const control = (field) => layout.controls.find((c) => c.field === field);
  const move = (field, to) => apply((l) => placeControl(l, field, to));

  /** A drag of a field (to move it) or of its corner (to size it), snapped to the grid as it goes. */
  const startDrag = (e, field, how) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const c = control(field);
    setPicked(field);
    drag.current = { field, how, x0: e.clientX, y0: e.clientY, c };
    try { e.currentTarget.setPointerCapture(e.pointerId); } catch { /* a pointer the window did not make */ }
  };
  const dragging = (e) => {
    const d = drag.current;
    if (!d) return;
    const dx = e.clientX - d.x0;
    const dy = e.clientY - d.y0;
    move(d.field, d.how === 'size' ? { w: d.c.w + dx, h: d.c.h + dy } : { x: d.c.x + dx, y: d.c.y + dy });
  };
  const endDrag = () => { drag.current = null; };

  const keys = (e, field) => {
    const c = control(field);
    const step = { ArrowLeft: [-GRID, 0], ArrowRight: [GRID, 0], ArrowUp: [0, -GRID], ArrowDown: [0, GRID] }[e.key];
    if (step) {
      e.preventDefault();
      move(field, e.shiftKey ? { w: c.w + step[0], h: c.h + step[1] } : { x: c.x + step[0], y: c.y + step[1] });
    } else if (e.key === 'Delete') {
      e.preventDefault();
      takeOffPicked(field);
    }
  };
  const takeOffPicked = (field = picked) => {
    if (!field) return;
    try {
      const next = takeOff(layout, field);
      setLayout(next);
      setPicked(next.controls[0]?.field || null);
    } catch (e) {
      toast(e.message, { tone: 'bad' });
    }
  };
  // A field put back that points at another table is a list of that table's records by their first words.
  const putOnLookup = async (name) => {
    const f = fields.find((x) => x.name === name);
    if (!f) return;
    const rel = (db.relationships || []).find((r) => r.from.table === layout.table && r.from.field === name);
    let theirs = [];
    if (rel) theirs = await shell.base.fields({ id: db.id, table: rel.to.table }).catch(() => []);
    apply((l) => putBack(l, f, () => theirs, db.relationships || []));
    setPicked(name);
  };

  const save = async () => {
    setBusy(true);
    try {
      const s = await shell.base.saveForm({ id: db.id, layout });
      toast(t('The form {name} is kept as it is laid out now.', { name: layout.name }), { tone: 'good' });
      onSaved?.(s, layout);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    } finally {
      setBusy(false);
    }
  };

  const at = picked ? control(picked) : null;
  return (
    <div className="bs-form-run bs-fd">
      <div className="bs-fd-bar" role="toolbar" aria-label={t('Form design')}>
        <Select className="rw-select bs-fd-add" aria-label={t('Put a field on the form')} value="" disabled={!off.length} onChange={(e) => { if (e.target.value) putOnLookup(e.target.value); }}>
          <option value="">{off.length ? t('Put a field on…') : t('Every field is on the form')}</option>
          {off.map((f) => <option key={f.name} value={f.name}>{f.name}</option>)}
        </Select>
        <Button icon="trash" label={t('Take Off')} className="bs-fd-remove" disabled={!at || layout.controls.length < 2} title={t('Take the field picked off the form. It stays in the table.')} onClick={() => takeOffPicked()} />
        {at ? <span className="bs-fd-where" data-field={at.field}>{t('{field}: {w} by {h}, at {x}, {y}', { field: at.field, w: at.w, h: at.h, x: at.x, y: at.y })}</span> : null}
        <span className="grow" />
        <Button label={t('Cancel')} className="bs-fd-cancel" onClick={onCancel} />
        <Button primary icon="save" label={t('Save')} className="bs-fd-save" disabled={busy || !changed} onClick={save} />
      </div>
      <p className="bs-hint bs-fd-hint">{t('Drag a field to move it and its corner to size it. The arrow keys move the field picked; with Shift they size it. Click a label to rename it.')}</p>
      <div className="bs-form-page bs-fd-page" style={{ width: layout.width, minHeight: layout.height + 48, backgroundSize: `${GRID}px ${GRID}px` }} data-form={layout.name}>
        <h2 className="bs-form-title">{layout.title}</h2>
        {layout.controls.map((c) => (
          <React.Fragment key={c.field}>
            <input className="rw-input bs-fd-label" data-field={c.field} aria-label={t('Label of {field}', { field: c.field })} style={{ left: labelLeft(c), top: c.y, width: Math.min(LABEL_WIDTH, c.x - labelLeft(c) - 4) }}
              value={c.label} onFocus={() => setPicked(c.field)} onChange={(e) => apply((l) => ({ ...l, controls: l.controls.map((x) => (x.field === c.field ? { ...x, label: e.target.value } : x)) }))}
              onBlur={(e) => apply((l) => relabel(l, c.field, e.target.value))} onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }} />
            <div className={`bs-fd-box${picked === c.field ? ' on' : ''}`} data-field={c.field} role="button" tabIndex={0} aria-label={t('{field}, a {kind} field', { field: c.field, kind: KIND_NAMES()[c.kind] || c.kind })}
              style={{ left: c.x, top: c.y, width: c.w, height: c.h }}
              onPointerDown={(e) => startDrag(e, c.field, 'move')} onPointerMove={dragging} onPointerUp={endDrag} onPointerCancel={endDrag}
              onFocus={() => setPicked(c.field)} onKeyDown={(e) => keys(e, c.field)}>
              <span className="bs-fd-name">{c.field}</span>
              {c.kind === 'check' ? null : <span className="bs-fd-size" aria-hidden="true" onPointerDown={(e) => startDrag(e, c.field, 'size')} onPointerMove={dragging} onPointerUp={endDrag} onPointerCancel={endDrag} />}
            </div>
          </React.Fragment>
        ))}
        {(layout.subforms || []).map((s) => (
          <div key={s.table} className="bs-fd-sub" style={{ left: s.x, top: s.y, width: s.w, height: s.h }} data-subform={s.table}>{s.label}</div>
        ))}
      </div>
    </div>
  );
}

export const FORM_DESIGN_CSS = `
.bs-fd-bar { display: flex; align-items: center; gap: 6px; padding: 8px 16px; border-bottom: 1px solid var(--line-soft); background: var(--window); position: sticky; top: 0; z-index: 2; }
.bs-fd-where { font-size: 12px; color: var(--ink-3); font-variant-numeric: tabular-nums; }
.bs-fd-hint { margin: 8px 16px 0; }
.bs-fd-page { background-image: radial-gradient(circle, var(--line-soft) 1px, transparent 1px); background-position: 0 48px; }
.bs-fd-label { position: absolute; height: 24px; box-sizing: border-box; transform: translateY(48px); font-size: 13px; color: var(--ink-2); background: transparent; border-color: transparent; padding: 0 4px; }
.bs-fd-label:hover, .bs-fd-label:focus { border-color: var(--line-strong); background: var(--surface); }
.bs-fd-box { position: absolute; box-sizing: border-box; transform: translateY(48px); border: 1px dashed var(--line-strong); border-radius: 3px; background: var(--sunken); cursor: move; display: flex; align-items: center; padding: 0 6px; overflow: hidden; touch-action: none; user-select: none; }
.bs-fd-box.on { border: 1.5px solid var(--accent); background: var(--accent-soft); outline: none; }
.bs-fd-box:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.bs-fd-name { font-size: 12px; color: var(--ink-2); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; pointer-events: none; }
.bs-fd-size { position: absolute; right: 0; bottom: 0; width: 10px; height: 10px; cursor: nwse-resize; background: linear-gradient(135deg, transparent 50%, var(--accent) 50%); }
.bs-fd-sub { position: absolute; transform: translateY(48px); box-sizing: border-box; border: 1px dashed var(--line-soft); border-radius: var(--r-2); padding: 6px 10px; font-size: 12.5px; color: var(--ink-3); }
`;
