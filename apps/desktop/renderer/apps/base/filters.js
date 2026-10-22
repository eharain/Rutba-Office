// Searching and filtering a grid, the same above a table's records and a
// query's results (owner, 2026-10-22, D-2026-10-22-07): a search box for
// words in any column, a filter for a column by its type (dates and numbers
// by a range, words by matching, yes/no), filter by selection from a cell,
// and each filter a chip with its cross. The database does the work
// (packages/database/src/filter.js); this only says what to look for.

import React, { useCallback, useState } from 'react';
import { Button, Dialog, Icon, Input, Select, Search, t } from '@rutba/office-ui';
import { filterKindOf } from '@rutba/database/filter';

/** A grid's search and filters, and how to change them. */
export function useGridFilter() {
  const [search, setSearch] = useState('');
  const [filters, setFilters] = useState([]);
  // One filter a column: a new one for a column takes the old one's place.
  const put = useCallback((f) => setFilters((all) => [...all.filter((x) => x.column !== f.column || (f.kind === 'notValue' && x.kind === 'notValue' && x.value !== f.value)), f]), []);
  const drop = useCallback((i) => setFilters((all) => all.filter((_, k) => k !== i)), []);
  const clear = useCallback(() => { setFilters([]); setSearch(''); }, []);
  return { search, setSearch, filters, put, drop, clear, active: Boolean(search.trim() || filters.length) };
}

/** A filter in words, for its chip. */
export function filterWords(f) {
  switch (f.kind) {
    case 'text': return f.op === 'is' ? t('{column} is {value}', f) : f.op === 'starts' ? t('{column} starts with {value}', f) : t('{column} contains {value}', f);
    case 'range':
      if (f.from !== '' && f.from != null && f.to !== '' && f.to != null) return t('{column} from {from} to {to}', f);
      return f.from !== '' && f.from != null ? t('{column} from {from}', f) : t('{column} up to {to}', f);
    case 'yesno': return f.value ? t('{column}: Yes', f) : t('{column}: No', f);
    case 'value': return t('{column} is {value}', { column: f.column, value: f.value ?? t('(empty)') });
    case 'notValue': return t('{column} is not {value}', { column: f.column, value: f.value ?? t('(empty)') });
    case 'empty': return t('{column} is empty', f);
    case 'notEmpty': return t('{column} is not empty', f);
    default: return f.column;
  }
}

/** The search box and the filters' chips, above a grid. */
export function FilterBar({ grid }) {
  return (
    <div className="bs-filterbar">
      <span className="bs-search"><Search value={grid.search} onChange={grid.setSearch} placeholder={t('Search these records')} style={{ width: 260 }} /></span>
      {grid.filters.map((f, i) => (
        <span key={`${f.column}-${f.kind}-${i}`} className="bs-chip" data-filter={f.column}>
          <Icon name="filter" size={12} />
          <span>{filterWords(f)}</span>
          <button type="button" className="bs-chip-x" aria-label={t('Take this filter off')} title={t('Take this filter off')} onClick={() => grid.drop(i)}><Icon name="close" size={12} /></button>
        </span>
      ))}
      {grid.active ? <Button ghost label={t('Clear all')} className="bs-filter-clear" onClick={grid.clear} /> : null}
    </div>
  );
}

/** A column's filter, by its type: a range for dates and numbers, matching for words, Yes or No. */
export function ColumnFilter({ column, onApply, onClose }) {
  const kind = filterKindOf(column.type);
  const dates = column.type === 'date' || column.type === 'datetime';
  const [op, setOp] = useState('contains');
  const [value, setValue] = useState('');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const apply = (f) => { onApply({ column: column.name, ...f }); onClose(); };
  const ready = kind === 'text' ? value.trim() !== '' : kind === 'range' ? from !== '' || to !== '' : true;
  return (
    <Dialog
      title={t('Filter by {column}', { column: column.name })}
      width={420}
      onClose={onClose}
      actions={(
        <>
          <Button ghost label={t('Empty')} className="bs-filter-empty" onClick={() => apply({ kind: 'empty' })} />
          <Button ghost label={t('Not empty')} className="bs-filter-notempty" onClick={() => apply({ kind: 'notEmpty' })} />
          <span style={{ flex: 1 }} />
          <Button label={t('Cancel')} onClick={onClose} />
          {kind === 'yesno' ? null : <Button primary label={t('Filter')} className="bs-filter-apply" disabled={!ready} onClick={() => apply(kind === 'text' ? { kind, op, value: value.trim() } : { kind, from, to })} />}
        </>
      )}
    >
      {kind === 'text' ? (
        <div className="bs-filter-row">
          <Select className="bs-filter-op" value={op} aria-label={t('How it matches')} onChange={(e) => setOp(e.target.value)}>
            <option value="contains">{t('Contains')}</option>
            <option value="starts">{t('Starts with')}</option>
            <option value="is">{t('Is exactly')}</option>
          </Select>
          <Input className="bs-filter-value" autoFocus value={value} aria-label={t('Words')} onChange={(e) => setValue(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && ready) apply({ kind, op, value: value.trim() }); }} />
        </div>
      ) : kind === 'range' ? (
        <div className="bs-filter-row">
          <label className="bs-field"><span>{t('From')}</span><Input className="bs-filter-from" type={dates ? 'date' : 'number'} autoFocus value={from} onChange={(e) => setFrom(e.target.value)} /></label>
          <label className="bs-field"><span>{t('To')}</span><Input className="bs-filter-to" type={dates ? 'date' : 'number'} value={to} onChange={(e) => setTo(e.target.value)} /></label>
        </div>
      ) : (
        <div className="bs-filter-row">
          <Button label={t('Yes')} className="bs-filter-yes" onClick={() => apply({ kind: 'yesno', value: true })} />
          <Button label={t('No')} className="bs-filter-no" onClick={() => apply({ kind: 'yesno', value: false })} />
        </div>
      )}
      {kind === 'range' ? <p className="bs-hint">{t('Leave either end empty for no limit on that side.')}</p> : null}
    </Dialog>
  );
}

/** Filter by selection: a cell's value, only it or all but it, from the cell's menu. */
export function selectionItems(grid, column, value) {
  const v = value === undefined ? null : value;
  return [
    { label: t('Only this value'), run: () => grid.put({ column, kind: 'value', value: v }) },
    { label: t('All but this value'), run: () => grid.put({ column, kind: 'notValue', value: v }) },
  ];
}

export const FILTER_CSS = `
.bs-filterbar { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; padding: 6px 16px; border-bottom: 1px solid var(--line-soft); }
.bs-chip { display: inline-flex; align-items: center; gap: 4px; padding: 2px 2px 2px 8px; border: 1px solid var(--line); border-radius: 12px; background: var(--sunken); font-size: 12.5px; }
.bs-chip-x { display: grid; place-items: center; width: 20px; height: 20px; border: 0; border-radius: 50%; background: none; color: var(--ink-3); cursor: pointer; }
.bs-chip-x:hover { background: var(--wash-strong); color: var(--ink); }
.bs-filter-row { display: flex; gap: 8px; align-items: end; }
.bs-filter-row .bs-field { flex: 1; }
.bs-filter-value { flex: 1; }
.bs-th-filter { display: grid; place-items: center; width: 22px; height: 22px; margin-inline-start: auto; border: 0; border-radius: var(--r-1, 4px); background: none; color: var(--ink-3); cursor: pointer; opacity: 0.55; }
.bs-th:hover .bs-th-filter, .bs-th-filter.on { opacity: 1; }
.bs-th-filter.on { color: var(--accent-text, var(--accent)); }
.bs-th-filter:hover { background: var(--wash-strong); }
`;
