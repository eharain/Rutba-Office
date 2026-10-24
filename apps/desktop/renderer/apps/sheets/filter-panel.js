// The filter arrows in a list's heading row, and the menu each opens.
//
// Data → Filter (or Format as Table) puts an arrow in every heading cell, as
// Excel does; the arrow of a column that filters shows a funnel. Its menu
// sorts the list by the column, clears the column's filter, and lists the
// column's values with a tick each, a search over them and Select All; OK
// keeps the ticked values and hides the rows of the rest. The engine works
// the list out (filterOptions) and writes the filter as Excel does; this is
// only the menu.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Search, t } from '@rutba/office-ui';

/** The heading cells of every filtered list on the sheet: "row:col" → the list and its column. */
export function filterHeads(tables) {
  const heads = new Map();
  for (const tb of tables || []) {
    const headerRows = tb.headerRows ?? tb.headerRowCount ?? 1;
    if (!headerRows) continue;
    (tb.columns || []).forEach((column, i) => {
      heads.set(tb.top + ':' + (tb.left + i), { table: tb.name, column, filtered: (tb.filtered || []).includes(i) });
    });
  }
  return heads;
}

/** The arrow drawn at a heading cell's right. */
export function FilterArrow({ head, onOpen }) {
  return (
    <button
      type="button"
      className={`sh-filter-arrow${head.filtered ? ' on' : ''}`}
      data-tip={head.filtered ? t('{column}: filtered. Click to change the filter', { column: head.column }) : t('Filter or sort by {column}', { column: head.column })}
      aria-label={head.filtered ? t('{column}: filtered', { column: head.column }) : t('Filter or sort by {column}', { column: head.column })}
      onMouseDown={(e) => {
        e.stopPropagation();
        e.preventDefault();
        const box = e.currentTarget.getBoundingClientRect();
        onOpen(head, { x: box.left, y: box.bottom + 2, right: box.right });
      }}
    >
      <Icon name={head.filtered ? 'filter' : 'chevronDown'} size={10} />
    </button>
  );
}

/**
 * The menu: `panel` is the engine's filterOptions for one column
 * ({ table, column, values, selected }), `at` where its arrow is on screen.
 */
export function FilterMenu({ panel, at, onApply, onSort, onClose }) {
  const [query, setQuery] = useState('');
  const [ticked, setTicked] = useState(() => new Set(panel.selected === null ? panel.values : panel.selected.filter((v) => panel.values.includes(v))));
  const ref = useRef(null);

  // Escape, or a press anywhere outside, closes it as a menu closes.
  useEffect(() => {
    const away = (e) => { if (ref.current && !ref.current.contains(e.target)) onClose(); };
    const key = (e) => { if (e.key === 'Escape') { e.stopPropagation(); onClose(); } };
    document.addEventListener('mousedown', away, true);
    document.addEventListener('keydown', key, true);
    return () => { document.removeEventListener('mousedown', away, true); document.removeEventListener('keydown', key, true); };
  }, [onClose]);

  const shown = useMemo(() => {
    const q = query.trim().toLowerCase();
    return q ? panel.values.filter((v) => (v || t('(Blanks)')).toLowerCase().includes(q)) : panel.values;
  }, [panel.values, query]);
  const allShownTicked = shown.length > 0 && shown.every((v) => ticked.has(v));
  const toggle = (v) => setTicked((s) => { const n = new Set(s); if (n.has(v)) n.delete(v); else n.add(v); return n; });
  const toggleAll = () => setTicked((s) => {
    const n = new Set(s);
    for (const v of shown) { if (allShownTicked) n.delete(v); else n.add(v); }
    return n;
  });
  const apply = () => {
    // A search keeps only what it found and ticked, as Excel's does.
    const keep = query.trim() ? shown.filter((v) => ticked.has(v)) : panel.values.filter((v) => ticked.has(v));
    onApply(keep.length === panel.values.length ? null : keep);
  };

  // Kept on screen: under its arrow, its right edge at the arrow's when it would run off the window.
  const width = 240;
  const left = Math.max(8, Math.min(at.x, window.innerWidth - width - 8));
  const top = Math.min(at.y, Math.max(8, window.innerHeight - 380));

  return (
    <div ref={ref} className="sh-filter-menu" role="dialog" aria-label={t('Filter {column}', { column: panel.column })} style={{ left, top, width }}>
      <button type="button" className="sh-filter-item sh-filter-asc" onClick={() => onSort(true)}>
        <Icon name="sort" size={13} />{t('Sort A to Z')}
      </button>
      <button type="button" className="sh-filter-item sh-filter-desc" onClick={() => onSort(false)}>
        <Icon name="sort" size={13} />{t('Sort Z to A')}
      </button>
      <div className="sh-filter-sep" />
      <button type="button" className="sh-filter-item sh-filter-clear" disabled={panel.selected === null} onClick={() => onApply(null)}>
        <Icon name="filterClear" size={13} />{t('Clear filter from "{column}"', { column: panel.column })}
      </button>
      <div className="sh-filter-sep" />
      <Search value={query} onChange={setQuery} placeholder={t('Search')} autoFocus />
      <div className="sh-filter-values">
        <label className="sh-filter-value sh-filter-all">
          <input type="checkbox" checked={allShownTicked} onChange={toggleAll} />
          {query.trim() ? t('(Select all search results)') : t('(Select all)')}
        </label>
        {shown.map((v) => (
          <label key={v} className="sh-filter-value" data-value={v}>
            <input type="checkbox" checked={ticked.has(v)} onChange={() => toggle(v)} />
            {v === '' ? t('(Blanks)') : v}
          </label>
        ))}
        {shown.length ? null : <div className="sh-filter-none">{t('Nothing matches')}</div>}
      </div>
      <div className="sh-filter-actions">
        <Button primary label={t('OK')} disabled={!ticked.size} onClick={apply} />
        <Button label={t('Cancel')} onClick={onClose} />
      </div>
    </div>
  );
}

export const FILTER_CSS = `
.sh-filter-arrow { position: absolute; right: 1px; top: 50%; transform: translateY(-50%); z-index: 3; width: 16px; height: 16px; display: inline-flex; align-items: center; justify-content: center; padding: 0; border: 1px solid var(--line); border-radius: 2px; background: var(--surface); color: var(--ink-2); cursor: pointer; }
.sh-filter-arrow:hover { border-color: var(--accent); color: var(--text); }
.sh-filter-arrow.on { color: var(--accent-text); border-color: var(--accent); }
.sh-cell.filter-head { padding-right: 20px; }
.sh-filter-menu { position: fixed; z-index: 40; display: flex; flex-direction: column; gap: 2px; padding: 6px; border: 1px solid var(--line); border-radius: 6px; background: var(--surface); color: var(--text); box-shadow: 0 8px 24px rgba(0, 0, 0, 0.18); font-size: 12.5px; }
.sh-filter-item { display: flex; align-items: center; gap: 8px; padding: 5px 8px; border: 0; border-radius: 4px; background: transparent; color: inherit; text-align: left; cursor: pointer; font: inherit; }
.sh-filter-item:hover:not(:disabled) { background: var(--hover); }
.sh-filter-item:disabled { color: var(--ink-3); cursor: default; }
.sh-filter-sep { height: 1px; margin: 3px 0; background: var(--line-soft); }
.sh-filter-values { max-height: 190px; overflow: auto; margin-top: 4px; padding: 4px; border: 1px solid var(--line-soft); border-radius: 4px; }
.sh-filter-value { display: flex; align-items: center; gap: 6px; padding: 2px 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; cursor: pointer; }
.sh-filter-none { padding: 4px; color: var(--ink-3); }
.sh-filter-actions { display: flex; justify-content: flex-end; gap: 6px; margin-top: 6px; }
`;
