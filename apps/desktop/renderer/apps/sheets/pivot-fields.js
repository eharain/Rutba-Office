// PivotTable Fields: the pane Excel shows beside a pivot when the cell is in
// it — the source's fields, each ticked when the pivot uses it, and the
// three areas a field can stand in: Rows, Columns and Values. A tick puts
// a field in Rows (or, a field of numbers, in Values); each field in an
// area moves up or down, to another area, or out; a value is summed,
// counted, averaged, or its largest, smallest or product taken. Every
// change lays the pivot out again at once, as Excel's does.

import React from 'react';
import { Button } from '@rutba/office-ui';

export const SUMMARIES = [['sum', 'Sum'], ['count', 'Count'], ['average', 'Average'], ['max', 'Max'], ['min', 'Min'], ['product', 'Product']];
const AREAS = [['rows', 'Rows'], ['cols', 'Columns'], ['values', 'Values']];

/** The layout as the pane holds it: names in rows and columns, `{ field, subtotal }` in values. */
export function layoutOf(pivot) {
  return {
    rows: [...(pivot.rows || [])],
    cols: [...(pivot.cols || [])],
    values: (pivot.valueFields || []).map((v) => ({ field: v.field, subtotal: v.subtotal || 'sum' })),
  };
}

/** The layout with `field` taken out of every area. */
const without = (layout, field) => ({
  rows: layout.rows.filter((f) => f !== field),
  cols: layout.cols.filter((f) => f !== field),
  values: layout.values.filter((v) => v.field !== field),
});

export function PivotFieldsPane({ pivot, numeric = () => false, onChange }) {
  const layout = layoutOf(pivot);
  const used = new Set([...layout.rows, ...layout.cols, ...layout.values.map((v) => v.field)]);
  const tick = (field, on) => {
    if (!on) return onChange(without(layout, field));
    // As Excel puts a ticked field: numbers summed in Values, anything else in Rows.
    if (numeric(field)) return onChange({ ...layout, values: [...layout.values, { field, subtotal: 'sum' }] });
    return onChange({ ...layout, rows: [...layout.rows, field] });
  };
  const move = (area, i, step) => {
    const list = [...layout[area]];
    const j = i + step;
    if (j < 0 || j >= list.length) return;
    [list[i], list[j]] = [list[j], list[i]];
    onChange({ ...layout, [area]: list });
  };
  const moveTo = (area, i, to) => {
    const item = layout[area][i];
    const field = area === 'values' ? item.field : item;
    const rest = without(layout, field);
    if (to === 'values') onChange({ ...rest, values: [...rest.values, { field, subtotal: area === 'values' ? item.subtotal : 'sum' }] });
    else onChange({ ...rest, [to]: [...rest[to], field] });
  };
  const summarise = (i, subtotal) => onChange({ ...layout, values: layout.values.map((v, k) => (k === i ? { ...v, subtotal } : v)) });
  const remove = (area, i) => {
    const item = layout[area][i];
    onChange(without(layout, area === 'values' ? item.field : item));
  };
  return (
    <div className="sh-pf" data-pivot={pivot.name}>
      <div className="sh-pf-lead">Choose fields to add to the report:</div>
      <div className="sh-pf-fields" role="list">
        {(pivot.fields || []).map((field) => (
          <label key={field} className="sh-pf-field" role="listitem">
            <input type="checkbox" checked={used.has(field)} data-field={field} onChange={(e) => tick(field, e.target.checked)} />
            <span>{field}</span>
          </label>
        ))}
      </div>
      {AREAS.map(([area, title]) => (
        <div key={area} className="sh-pf-area" data-area={area}>
          <div className="sh-pf-area-title">{title}</div>
          {layout[area].length ? layout[area].map((item, i) => {
            const field = area === 'values' ? item.field : item;
            return (
              <div key={field + ':' + i} className="sh-pf-item" data-field={field}>
                <span className="sh-pf-name">{area === 'values' ? `${SUMMARIES.find(([k]) => k === item.subtotal)?.[1] || 'Sum'} of ${field}` : field}</span>
                {area === 'values' ? (
                  <select className="sh-pf-sum" value={item.subtotal} title="Summarise values by" onChange={(e) => summarise(i, e.target.value)}>
                    {SUMMARIES.map(([k, label]) => <option key={k} value={k}>{label}</option>)}
                  </select>
                ) : null}
                <Button icon="chevronUp" title="Move up" disabled={i === 0} onClick={() => move(area, i, -1)} />
                <Button icon="chevronDown" title="Move down" disabled={i === layout[area].length - 1} onClick={() => move(area, i, 1)} />
                <select className="sh-pf-to" value="" title="Move to another area" onChange={(e) => { if (e.target.value) moveTo(area, i, e.target.value); }}>
                  <option value="">Move to…</option>
                  {AREAS.filter(([a]) => a !== area).map(([a, label]) => <option key={a} value={a}>{label}</option>)}
                </select>
                <Button icon="close" title="Remove field" disabled={area === 'values' && layout.values.length === 1} onClick={() => remove(area, i)} />
              </div>
            );
          }) : <div className="sh-pf-empty">None — tick a field above, or move one here.</div>}
        </div>
      ))}
    </div>
  );
}

export const PIVOT_FIELDS_CSS = `
.sh-pf { display: flex; flex-direction: column; gap: 10px; font-size: 12.5px; }
.sh-pf-lead { color: var(--ink-2); }
.sh-pf-fields { display: flex; flex-direction: column; border: 1px solid var(--line-soft); border-radius: var(--r-2); max-height: 220px; overflow: auto; }
.sh-pf-field { display: flex; align-items: center; gap: 6px; padding: 4px 8px; border-bottom: 1px solid var(--line-soft); cursor: pointer; }
.sh-pf-area { border: 1px solid var(--line-soft); border-radius: var(--r-2); padding: 6px; display: flex; flex-direction: column; gap: 4px; }
.sh-pf-area-title { font-weight: 600; font-size: 11.5px; color: var(--ink-2); text-transform: uppercase; letter-spacing: .03em; }
.sh-pf-item { display: flex; align-items: center; gap: 2px; background: var(--window); border: 1px solid var(--line-soft); border-radius: var(--r-2); padding: 2px 2px 2px 6px; }
.sh-pf-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sh-pf-item { flex-wrap: wrap; }
.sh-pf-area[data-area="values"] .sh-pf-name { flex-basis: 100%; padding: 2px 0; }
.sh-pf-sum, .sh-pf-to { font: inherit; font-size: 11.5px; max-width: 86px; }
.sh-pf-empty { color: var(--ink-3); font-size: 11.5px; padding: 2px; }
`;
