// PivotTable Fields: the pane Excel shows beside a pivot when the cell is in
// it — the source's fields, each ticked when the pivot uses it, and the
// four areas a field can stand in: Filters, Columns, Rows and Values. A tick
// puts a field in Rows (or, a field of numbers, in Values); each field in an
// area moves up or down, to another area, or out; a value is summed,
// counted, averaged, or its largest, smallest or product taken; a filter
// shows all its items or one; a date field can be shown by month. Every
// change lays the pivot out again at once, as Excel's does.

import React from 'react';
import { Button, t } from '@rutba/office-ui';

export const SUMMARIES = [['sum', t('Sum')], ['count', t('Count')], ['average', t('Average')], ['max', t('Max')], ['min', t('Min')], ['product', t('Product')]];
// In Excel's order: the filters above the columns, the rows, then the values.
const AREAS = [['filters', t('Filters')], ['cols', t('Columns')], ['rows', t('Rows')], ['values', t('Values')]];
// The areas whose entries are objects, not bare names.
const KEYED = new Set(['filters', 'values']);

/** The layout as the pane holds it: names in rows and columns, `{ field, subtotal }` in values, `{ field, item }` in filters. */
export function layoutOf(pivot) {
  return {
    filters: (pivot.filters || []).map((f) => ({ field: f.field, item: f.item ?? null })),
    rows: [...(pivot.rows || [])],
    cols: [...(pivot.cols || [])],
    values: (pivot.valueFields || []).map((v) => ({ field: v.field, subtotal: v.subtotal || 'sum' })),
  };
}

/** The layout with `field` taken out of every area. */
const without = (layout, field) => ({
  filters: layout.filters.filter((f) => f.field !== field),
  rows: layout.rows.filter((f) => f !== field),
  cols: layout.cols.filter((f) => f !== field),
  values: layout.values.filter((v) => v.field !== field),
});

export function PivotFieldsPane({ pivot, numeric = () => false, onChange, onGroup = () => {} }) {
  const layout = layoutOf(pivot);
  const used = new Set([...layout.filters.map((f) => f.field), ...layout.rows, ...layout.cols, ...layout.values.map((v) => v.field)]);
  const fieldAt = (field) => (pivot.fields || []).indexOf(field);
  const isDate = (field) => Boolean(pivot.dates?.[fieldAt(field)]);
  const byMonth = (field) => Boolean(pivot.months?.[fieldAt(field)]);
  const tick = (field, on) => {
    if (!on) return onChange(without(layout, field));
    // As Excel puts a ticked field: numbers summed in Values, anything else in Rows.
    if (numeric(field) && !isDate(field)) return onChange({ ...layout, values: [...layout.values, { field, subtotal: 'sum' }] });
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
    const field = KEYED.has(area) ? item.field : item;
    const rest = without(layout, field);
    if (to === 'values') onChange({ ...rest, values: [...rest.values, { field, subtotal: area === 'values' ? item.subtotal : 'sum' }] });
    else if (to === 'filters') onChange({ ...rest, filters: [...rest.filters, { field, item: null }] });
    else onChange({ ...rest, [to]: [...rest[to], field] });
  };
  const summarise = (i, subtotal) => onChange({ ...layout, values: layout.values.map((v, k) => (k === i ? { ...v, subtotal } : v)) });
  const choose = (i, item) => onChange({ ...layout, filters: layout.filters.map((f, k) => (k === i ? { ...f, item: item === '' ? null : item } : f)) });
  const remove = (area, i) => {
    const item = layout[area][i];
    onChange(without(layout, KEYED.has(area) ? item.field : item));
  };
  return (
    <div className="sh-pf" data-pivot={pivot.name}>
      <div className="sh-pf-lead">{t('Choose fields to add to the report:')}</div>
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
            const field = KEYED.has(area) ? item.field : item;
            const filter = area === 'filters' ? (pivot.filters || []).find((f) => f.field === field) : null;
            return (
              <div key={field + ':' + i} className="sh-pf-item" data-field={field}>
                <span className="sh-pf-name">{area === 'values' ? t('{summary} of {field}', { summary: SUMMARIES.find(([k]) => k === item.subtotal)?.[1] || t('Sum'), field }) : field}</span>
                {area === 'values' ? (
                  <select className="sh-pf-sum" value={item.subtotal} title={t('Summarise values by')} onChange={(e) => summarise(i, e.target.value)}>
                    {SUMMARIES.map(([k, label]) => <option key={k} value={k}>{label}</option>)}
                  </select>
                ) : null}
                {area === 'filters' ? (
                  <select className="sh-pf-pick" value={item.item ?? ''} title={t('Show the items')} onChange={(e) => choose(i, e.target.value)}>
                    <option value="">{t('(All)')}</option>
                    {(filter?.items || []).map((label) => <option key={label} value={label}>{label}</option>)}
                  </select>
                ) : null}
                {area !== 'values' && isDate(field) ? (
                  <label className="sh-pf-month" title={t('Show the dates as months')}>
                    <input type="checkbox" checked={byMonth(field)} data-field={field} onChange={(e) => onGroup(field, e.target.checked ? 'months' : null)} />
                    <span>{t('By month')}</span>
                  </label>
                ) : null}
                <Button icon="chevronUp" title={t('Move up')} disabled={i === 0} onClick={() => move(area, i, -1)} />
                <Button icon="chevronDown" title={t('Move down')} disabled={i === layout[area].length - 1} onClick={() => move(area, i, 1)} />
                <select className="sh-pf-to" value="" title={t('Move to another area')} onChange={(e) => { if (e.target.value) moveTo(area, i, e.target.value); }}>
                  <option value="">{t('Move to…')}</option>
                  {AREAS.filter(([a]) => a !== area).map(([a, label]) => <option key={a} value={a}>{label}</option>)}
                </select>
                <Button icon="close" title={t('Remove field')} disabled={area === 'values' && layout.values.length === 1} onClick={() => remove(area, i)} />
              </div>
            );
          }) : <div className="sh-pf-empty">{area === 'filters' ? t('None. Move a field here to show one of its items at a time.') : t('None. Tick a field above, or move one here.')}</div>}
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
.sh-pf-area[data-area="values"] .sh-pf-name, .sh-pf-area[data-area="filters"] .sh-pf-name { flex-basis: 100%; padding: 2px 0; }
.sh-pf-sum, .sh-pf-to { font: inherit; font-size: 11.5px; max-width: 86px; }
.sh-pf-pick { font: inherit; font-size: 11.5px; max-width: 120px; }
.sh-pf-month { display: inline-flex; align-items: center; gap: 3px; font-size: 11.5px; color: var(--ink-2); margin-right: 2px; cursor: pointer; }
.sh-pf-empty { color: var(--ink-3); font-size: 11.5px; padding: 2px; }
`;
