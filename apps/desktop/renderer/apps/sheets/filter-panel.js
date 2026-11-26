// The filter arrows in a list's heading row, and the menu each opens.
//
// Data → Filter (or Format as Table) puts an arrow in every heading cell, as
// Excel does; the arrow of a column that filters shows a funnel. Its menu
// sorts the list by the column, clears the column's filter, and lists the
// column's values with a tick each, a search over them and Select All; OK
// keeps the ticked values and hides the rows of the rest. Number Filters,
// Text Filters or Date Filters (whichever the column holds) opens a rule in
// place of the list: greater than a figure, between two, the top ten, a
// text that contains a word, a date after a day. The engine works the list
// out (filterOptions) and writes the filter as Excel does; this is only the
// menu.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Input, Search, Select, t } from '@rutba/office-ui';
import { FILTER_OPS, ruleFromForm } from '@rutba/sheet-view/filter-rules';

/** The words for a comparison, as Excel's filter menus have them. */
function opLabel(op) {
  switch (op) {
    case 'equals': return t('sheets.filter-panel.equals');
    case 'notEquals': return t('sheets.filter-panel.does-not-equal');
    case 'greater': return t('common.greater-than');
    case 'greaterOrEqual': return t('sheets.filter-panel.greater-equal');
    case 'less': return t('common.less-than');
    case 'lessOrEqual': return t('sheets.filter-panel.less-equal');
    case 'between': return t('sheets.dialogs.between');
    case 'top': return t('common.top');
    case 'bottom': return t('common.bottom');
    case 'begins': return t('sheets.filter-panel.begins-with');
    case 'ends': return t('sheets.filter-panel.ends-with');
    case 'contains': return t('base.filters.contains');
    case 'notContains': return t('sheets.filter-panel.does-not-contain');
    case 'before': return t('sheets.filter-panel.before');
    case 'after': return t('sheets.filter-panel.after');
    default: return op;
  }
}

/** What the rule's entry in the menu is called, by what the column holds. */
export function ruleMenuLabel(type) {
  if (type === 'number') return t('sheets.filter-panel.number-filters');
  if (type === 'date') return t('sheets.filter-panel.date-filters');
  return t('sheets.filter-panel.text-filters');
}

/** The rule form: a comparison and what to compare with. */
function RuleForm({ panel, onRule, onClose }) {
  const type = panel.type || 'text';
  const was = panel.rule;
  const ops = type === 'number' ? [...FILTER_OPS.number, 'top', 'bottom'] : FILTER_OPS[type];
  const [op, setOp] = useState(() => {
    if (was?.kind === 'top') return was.bottom ? 'bottom' : 'top';
    if (was?.op && ops.includes(was.op)) return was.op;
    return type === 'text' ? 'contains' : type === 'date' ? 'after' : 'greater';
  });
  const [a, setA] = useState(() => (was && was.kind !== 'top' && was.value != null ? String(was.value) : ''));
  const [b, setB] = useState(() => (was?.value2 != null ? String(was.value2) : ''));
  const [count, setCount] = useState(() => String(was?.kind === 'top' ? was.count : 10));
  const [percent, setPercent] = useState(() => Boolean(was?.kind === 'top' && was.percent));
  const top = op === 'top' || op === 'bottom';
  const rule = ruleFromForm(type, { op, a, b, count, percent });
  const apply = () => { if (rule) onRule(rule); };
  const onKey = (e) => { if (e.key === 'Enter') { e.preventDefault(); apply(); } };
  const kind = type === 'date' ? 'date' : 'text';
  return (
    <div className="sh-filter-rule">
      <Select className="sh-filter-op" value={op} onChange={(e) => setOp(e.target.value)} aria-label={t('sheets.filter-panel.comparison')}>
        {ops.map((o) => <option key={o} value={o}>{opLabel(o)}</option>)}
      </Select>
      {top ? (
        <div className="sh-filter-rule-row">
          <Input className="rw-input sh-filter-count" type="number" min={1} max={percent ? 100 : undefined} value={count} onChange={(e) => setCount(e.target.value)} onKeyDown={onKey} aria-label={t('sheets.queries.how-many')} autoFocus />
          <Select className="sh-filter-unit" value={percent ? 'percent' : 'items'} onChange={(e) => setPercent(e.target.value === 'percent')} aria-label={t('sheets.filter-panel.items-or-percent')}>
            <option value="items">{t('sheets.filter-panel.items')}</option>
            <option value="percent">{t('common.percent')}</option>
          </Select>
        </div>
      ) : (
        <>
          <Input className="rw-input sh-filter-a" type={kind} value={a} onChange={(e) => setA(e.target.value)} onKeyDown={onKey} aria-label={t('common.value')} autoFocus />
          {op === 'between' ? (
            <>
              <span className="sh-filter-and">{t('mail.and')}</span>
              <Input className="rw-input sh-filter-b" type={kind} value={b} onChange={(e) => setB(e.target.value)} onKeyDown={onKey} aria-label={t('sheets.filter-panel.second-value')} />
            </>
          ) : null}
        </>
      )}
      <div className="sh-filter-actions">
        <Button primary label={t('common.ok')} className="sh-filter-rule-ok" disabled={!rule} onClick={apply} />
        <Button label={t('common.cancel')} onClick={onClose} />
      </div>
    </div>
  );
}

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
      data-tip={head.filtered ? t('sheets.filter-panel.column-filtered-click', { column: head.column }) : t('sheets.filter-panel.filter-sort-column', { column: head.column })}
      aria-label={head.filtered ? t('sheets.filter-panel.column-filtered', { column: head.column }) : t('sheets.filter-panel.filter-sort-column', { column: head.column })}
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
export function FilterMenu({ panel, at, onApply, onRule, onSort, onClose }) {
  const [query, setQuery] = useState('');
  // The rule form takes the list's place; it opens by itself on a column a rule already filters.
  const [ruleOpen, setRuleOpen] = useState(() => Boolean(panel.rule));
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
    return q ? panel.values.filter((v) => (v || t('sheets.filter-panel.blanks')).toLowerCase().includes(q)) : panel.values;
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
    <div ref={ref} className="sh-filter-menu" role="dialog" aria-label={t('sheets.filter-panel.filter-column', { column: panel.column })} style={{ left, top, width }}>
      <button type="button" className="sh-filter-item sh-filter-asc" onClick={() => onSort(true)}>
        <Icon name="sort" size={13} />{t('sheets.sort-z')}
      </button>
      <button type="button" className="sh-filter-item sh-filter-desc" onClick={() => onSort(false)}>
        <Icon name="sort" size={13} />{t('sheets.sort-z-to-a')}
      </button>
      <div className="sh-filter-sep" />
      <button type="button" className="sh-filter-item sh-filter-clear" disabled={!(panel.filtered ?? panel.selected !== null)} onClick={() => onApply(null)}>
        <Icon name="filterClear" size={13} />{t('sheets.filter-panel.clear-filter-column', { column: panel.column })}
      </button>
      {onRule ? (
        <button type="button" className={`sh-filter-item sh-filter-rules${ruleOpen ? ' on' : ''}`} aria-expanded={ruleOpen} onClick={() => setRuleOpen((o) => !o)}>
          <Icon name="filter" size={13} />{ruleMenuLabel(panel.type)}
          <Icon name={ruleOpen ? 'chevronUp' : 'chevronDown'} size={10} style={{ marginLeft: 'auto' }} />
        </button>
      ) : null}
      <div className="sh-filter-sep" />
      {ruleOpen && onRule ? <RuleForm panel={panel} onRule={onRule} onClose={onClose} /> : (
      <>
      <Search value={query} onChange={setQuery} placeholder={t('common.search')} autoFocus />
      <div className="sh-filter-values">
        <label className="sh-filter-value sh-filter-all">
          <input type="checkbox" checked={allShownTicked} onChange={toggleAll} />
          {query.trim() ? t('sheets.filter-panel.select-all-search') : t('sheets.filter-panel.select-all')}
        </label>
        {shown.map((v) => (
          <label key={v} className="sh-filter-value" data-value={v}>
            <input type="checkbox" checked={ticked.has(v)} onChange={() => toggle(v)} />
            {v === '' ? t('sheets.filter-panel.blanks') : v}
          </label>
        ))}
        {shown.length ? null : <div className="sh-filter-none">{t('common.nothing-matches')}</div>}
      </div>
      <div className="sh-filter-actions">
        <Button primary label={t('common.ok')} disabled={!ticked.size} onClick={apply} />
        <Button label={t('common.cancel')} onClick={onClose} />
      </div>
      </>
      )}
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
.sh-filter-item.sh-filter-rules.on { background: var(--selected); }
.sh-filter-rule { display: flex; flex-direction: column; gap: 6px; margin-top: 4px; }
.sh-filter-rule select, .sh-filter-rule input { width: 100%; min-width: 0; }
.sh-filter-rule-row { display: flex; align-items: center; gap: 6px; }
.sh-filter-rule-row > * { flex: 1; }
.sh-filter-and { color: var(--ink-3); font-size: 12px; }
`;
