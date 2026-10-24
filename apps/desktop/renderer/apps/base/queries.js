// Data Files' query designer: tables put in, the fields to show, criteria in
// the words Access uses, the sort and totals, and the answer as a table. The
// design is kept in the file under a name; the SQL it stands for can be seen
// (the SQL view takes a query written by hand, and keeps it under a name as
// it is written: SqlView). The tables are joined
// by their relationships, so a person never writes a join. A column can be
// calculated instead, typed as Access takes it in a Field cell:
// "Line Total: [Quantity] * [Unit Price]".

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Input, Select, Dialog, Spinner, useMenu, t, tn } from '@rutba/office-ui';
import { useGridFilter, FilterBar, ColumnFilter, selectionItems } from './filters.js';
import { splitNamedExpression } from '@rutba/database/query-design';
import { canDesign } from './people.js';

/** The Field list's own choice that turns its cell into an expression box. */
const EXPRESSION = '\u0002expression';

const TOTALS = () => [['group', t('Group By')], ['sum', t('Sum')], ['avg', t('Average')], ['count', t('Count')], ['min', t('Min')], ['max', t('Max')], ['where', t('Where')]];
const blank = () => ({ table: '', field: '', exprText: null, show: true, sort: '', total: 'group', decimals: '', criteria: ['', ''] });
/** The places a column's answer is rounded to, beside its Total: as it comes, or 0 to 6. */
const DECIMALS = () => [['', t('As it comes')], ...[0, 1, 2, 3, 4, 5, 6].map((n) => [String(n), tn(n, '{count} place', '{count} places')])];
/** A column in use: a field chosen, or an expression typed. */
const used = (c) => Boolean((c.table && c.field) || (c.exprText !== null && c.exprText !== undefined && c.exprText.trim()));

/** A design from what the designer holds, the empty columns left out. */
function designOf(tables, columns, totals) {
  return {
    tables,
    columns: columns.filter(used).map((c) => ({
      ...(c.exprText !== null && c.exprText !== undefined ? splitNamedExpression(c.exprText) : { table: c.table, field: c.field }),
      show: c.show,
      sort: c.sort || null,
      total: totals ? c.total || 'group' : null,
      // Rounded where the person said so: an average spend to 2 places.
      ...(c.decimals !== '' && c.decimals !== null && c.decimals !== undefined ? { decimals: Number(c.decimals) } : {}),
      criteria: c.criteria.filter((x, i, all) => x.trim() || all.slice(i + 1).some((y) => y.trim())),
    })),
  };
}

export function QueryDesigner({ shell, db, name, onSaved, toast }) {
  const [tables, setTables] = useState([]);
  const [columns, setColumns] = useState([blank()]);
  const [totals, setTotals] = useState(false);
  const [fieldsOf, setFieldsOf] = useState({});
  const [answer, setAnswer] = useState(null);
  const [sql, setSql] = useState('');
  const [error, setError] = useState('');
  const [naming, setNaming] = useState(null);
  const [adding, setAdding] = useState('');
  // A saved query opens run, its records shown, its design behind Edit query (D-2026-10-22-06); a new one opens in design.
  const [editing, setEditing] = useState(!name);
  const [loaded, setLoaded] = useState(null);
  const [sqlEdit, setSqlEdit] = useState(null);

  // A saved query opens with its design; a new one with nothing in it.
  useEffect(() => {
    let live = true;
    setAnswer(null);
    setSql('');
    setError('');
    setEditing(!name);
    setLoaded(null);
    if (!name) { setTables([]); setColumns([blank()]); setTotals(false); return undefined; }
    shell.base.queries({ id: db.id }).then((all) => {
      const found = all.find((x) => x.name === name);
      if (!live || !found) return;
      const d = found.design;
      setTables(d.tables || []);
      setTotals((d.columns || []).some((c) => c.total));
      setColumns([...(d.columns || []).map((c) => ({ ...blank(), ...c, exprText: c.expr ? (c.as ? `${c.as}: ${c.expr}` : c.expr) : null, sort: c.sort || '', total: c.total || 'group', decimals: Number.isInteger(c.decimals) ? String(c.decimals) : '', criteria: [...(c.criteria || []), '', ''].slice(0, Math.max(2, (c.criteria || []).length)) })), blank()]);
      setLoaded(d);
    }).catch((e) => toast(e.message, { tone: 'bad' }));
    return () => { live = false; };
  }, [shell, db.id, name, toast]);

  useEffect(() => {
    let live = true;
    Promise.all(tables.filter((x) => !fieldsOf[x]).map((x) => shell.base.fields({ id: db.id, table: x }).then((f) => [x, f])))
      .then((pairs) => { if (live && pairs.length) setFieldsOf((m) => ({ ...m, ...Object.fromEntries(pairs) })); })
      .catch(() => {});
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, db.id, tables]);

  const design = useMemo(() => designOf(tables, columns, totals), [tables, columns, totals]);
  const setColumn = (i, patch) => setColumns((cs) => {
    const next = cs.map((c, k) => (k === i ? { ...c, ...patch } : c));
    // A field chosen (or an expression begun) on the last line opens another below it.
    const last = next[next.length - 1];
    return last.field || last.exprText !== null ? [...next, blank()] : next;
  });
  const addTable = (table) => {
    if (!table || tables.includes(table)) return;
    setTables([...tables, table]);
    setAdding('');
  };
  const removeTable = (table) => {
    setTables(tables.filter((x) => x !== table));
    setColumns((cs) => { const kept = cs.filter((c) => c.table !== table); return kept.length ? kept : [blank()]; });
  };

  const run = async () => {
    setError('');
    try {
      setAnswer(await shell.base.runQuery({ id: db.id, design }));
      setSql('');
    } catch (e) {
      setError(e.message);
      setAnswer(null);
    }
  };
  const showSql = async () => {
    setError('');
    try { setSql(await shell.base.sqlOf({ id: db.id, design })); } catch (e) { setError(e.message); }
  };
  const save = async (as) => {
    try {
      const summary = await shell.base.saveQuery({ id: db.id, name: as, design });
      setNaming(null);
      setLoaded(design);
      onSaved(summary, as);
      toast(t('Kept as {name}', { name: as }), { tone: 'good' });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };

  // Opened run: the saved design's records in the grid, its count as the grid reads it.
  const [found, setFound] = useState(null);
  const [runs, setRuns] = useState(0);
  const counted = useCallback((n, filtered) => setFound({ n, filtered }), []);

  const results = answer ? (
    <div className="bs-result">
      <table>
        <thead><tr>{answer.columns.map((c) => <th key={c}>{c}</th>)}</tr></thead>
        <tbody>{answer.rows.map((r, i) => <tr key={i}>{answer.columns.map((c) => <td key={c}>{r[c] === null || r[c] === undefined ? '' : String(r[c])}</td>)}</tr>)}</tbody>
      </table>
    </div>
  ) : null;

  if (!editing) {
    return (
      <div className="bs-query bs-query-run">
        <div className="bs-tools bs-q-runbar">
          <Icon name="filter" size={16} />
          <strong className="bs-q-title">{name}</strong>
          {found ? <span className="bs-count">{found.filtered ? tn(found.n, '{count} row found', '{count} rows found') : tn(found.n, '{count} row', '{count} rows')}</span> : null}
          <span className="bs-q-grow" />
          <Button icon="refresh" label={t('Run again')} className="bs-q-rerun" onClick={() => setRuns((n) => n + 1)} />
          {canDesign(db) ? <Button primary icon="pen" label={t('Edit query')} className="bs-q-edit" onClick={() => { setEditing(true); setAnswer(null); }} /> : null}
        </div>
        {loaded ? <QueryResults key={runs} shell={shell} db={db} design={loaded} onCount={counted} toast={toast} /> : <div className="bs-q-wait"><Spinner /></div>}
      </div>
    );
  }

  // A query written in SQL is changed as it was written, in the SQL view's box.
  if (typeof loaded?.sql === 'string') {
    return <SqlView shell={shell} db={db} name={name} text={sqlEdit ?? loaded.sql} onText={setSqlEdit} toast={toast}
      onSaved={(summary, as) => { setLoaded({ sql: sqlEdit ?? loaded.sql }); onSaved(summary, as); }} onResults={() => setEditing(false)} />;
  }

  const others = db.tables.map((x) => x.name).filter((x) => !tables.includes(x));
  return (
    <div className="bs-query">
      <div className="bs-q-tables">
        <span className="bs-q-caption">{t('Tables')}</span>
        {tables.map((x) => (
          <span key={x} className="bs-q-chip" data-table={x}>
            {x}
            <Button icon="close" label="" title={t('Take {table} out of the query', { table: x })} className="bs-q-chip-remove" onClick={() => removeTable(x)} />
          </span>
        ))}
        {others.length ? (
          <Select className="rw-select bs-q-add-table" aria-label={t('Add a table')} value={adding} onChange={(e) => addTable(e.target.value)}>
            <option value="">{t('Add a table')}</option>
            {others.map((x) => <option key={x} value={x}>{x}</option>)}
          </Select>
        ) : null}
        <label className="bs-check bs-q-totals"><input type="checkbox" checked={totals} onChange={(e) => setTotals(e.target.checked)} /><span>{t('Totals')}</span></label>
      </div>

      <div className="bs-q-grid-wrap">
        <table className="bs-q-grid">
          <thead>
            <tr>
              <th>{t('Field')}</th><th>{t('Show')}</th><th>{t('Sort')}</th>{totals ? <><th>{t('Total')}</th><th>{t('Decimals')}</th></> : null}<th>{t('Criteria')}</th><th>{t('Or')}</th><th />
            </tr>
          </thead>
          <tbody>
            {columns.map((c, i) => (
              <tr key={i} data-column={i}>
                <td>
                  {c.exprText !== null && c.exprText !== undefined ? (
                    <Input className="rw-input bs-q-expr" aria-label={t('Expression')} autoFocus={!c.exprText} placeholder={t('Line Total: [Quantity] * [Unit Price]')} title={t('A name, a colon, then the expression: [Field] names, + - * /, & to join text, and Round, Nz, IIf, Year, Month, Format, DateDiff and more')} value={c.exprText} onChange={(e) => setColumn(i, { exprText: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter') run(); }} />
                  ) : (
                    <Select className="rw-select bs-q-field" aria-label={t('Field')} value={c.table && c.field ? `${c.table}\u0001${c.field}` : ''} onChange={(e) => {
                      if (e.target.value === EXPRESSION) { setColumn(i, { table: '', field: '', exprText: '' }); return; }
                      const [tb, f] = e.target.value.split('\u0001');
                      setColumn(i, { table: tb || '', field: f || '' });
                    }}>
                      <option value="">{tables.length ? t('Choose a field') : t('Add a table first')}</option>
                      {tables.map((tb) => (
                        <optgroup key={tb} label={tb}>
                          {(fieldsOf[tb] || []).map((f) => <option key={f.name} value={`${tb}\u0001${f.name}`}>{f.name}</option>)}
                        </optgroup>
                      ))}
                      {tables.length ? <option value={EXPRESSION}>{t('Expression…')}</option> : null}
                    </Select>
                  )}
                </td>
                <td className="bs-q-center"><input type="checkbox" aria-label={t('Show')} checked={c.show} onChange={(e) => setColumn(i, { show: e.target.checked })} /></td>
                <td>
                  <Select className="rw-select bs-q-sort" aria-label={t('Sort')} value={c.sort} onChange={(e) => setColumn(i, { sort: e.target.value })}>
                    <option value="">{t('(not sorted)')}</option>
                    <option value="asc">{t('Ascending')}</option>
                    <option value="desc">{t('Descending')}</option>
                  </Select>
                </td>
                {totals ? (
                  <td>
                    <Select className="rw-select bs-q-total" aria-label={t('Total')} value={c.total} onChange={(e) => setColumn(i, { total: e.target.value })}>
                      {TOTALS().map(([k, label]) => <option key={k} value={k}>{label}</option>)}
                    </Select>
                  </td>
                ) : null}
                {totals ? (
                  <td>
                    <Select className="rw-select bs-q-decimals" aria-label={t('Decimals')} title={t('The places the answer is rounded to: an average to 2, for 15.08')} value={c.decimals ?? ''} onChange={(e) => setColumn(i, { decimals: e.target.value })}>
                      {DECIMALS().map(([k, label]) => <option key={k} value={k}>{label}</option>)}
                    </Select>
                  </td>
                ) : null}
                {[0, 1].map((r) => (
                  <td key={r}>
                    <Input className={`rw-input bs-q-criteria bs-q-row-${r}`} aria-label={r ? t('Or') : t('Criteria')} placeholder={r ? '' : t('e.g. > 10, Like "A*"')} value={c.criteria[r] || ''} onChange={(e) => setColumn(i, { criteria: c.criteria.map((x, k) => (k === r ? e.target.value : x)) })} onKeyDown={(e) => { if (e.key === 'Enter') run(); }} />
                  </td>
                ))}
                <td>{used(c) || c.exprText !== null ? <Button icon="close" label="" title={t('Take this field out')} className="bs-q-remove" onClick={() => setColumns((cs) => cs.filter((_, k) => k !== i))} /> : null}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="bs-tools">
        <Button primary icon="play" label={t('Run')} className="bs-q-run" disabled={!design.columns.length} onClick={run} />
        <Button icon="formula" label={t('SQL')} className="bs-q-sql" disabled={!design.columns.length} onClick={showSql} />
        <Button icon="save" label={name ? t('Save') : t('Save As')} className="bs-q-save" disabled={!design.columns.length} onClick={() => (name ? save(name) : setNaming(''))} />
        {name ? <Button label={t('Save As')} className="bs-q-save-as" disabled={!design.columns.length} onClick={() => setNaming('')} /> : null}
        {answer ? <span className="bs-count">{tn(answer.count, '{count} row', '{count} rows')}</span> : null}
        <span className="bs-q-grow" />
        {name ? <Button icon="table" label={t('Show results')} className="bs-q-results" onClick={() => { setEditing(false); setAnswer(null); }} /> : null}
      </div>
      {error ? <p className="bs-error" role="alert">{error}</p> : null}
      {sql ? <pre className="bs-q-sql-text selectable">{sql}</pre> : null}
      {results}

      {naming !== null ? (
        <Dialog title={t('Save Query')} onClose={() => setNaming(null)} actions={<><Button label={t('Cancel')} onClick={() => setNaming(null)} /><Button primary label={t('Save')} className="bs-q-save-named" disabled={!String(naming).trim()} onClick={() => save(String(naming).trim())} /></>}>
          <label className="bs-field"><span>{t('Name')}</span><Input className="rw-input bs-q-name" aria-label={t('Name')} autoFocus value={naming} onChange={(e) => setNaming(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && String(naming).trim()) save(String(naming).trim()); }} /></label>
        </Dialog>
      ) : null}
    </div>
  );
}

/* ── a query written in SQL ───────────────────────────────────────── */

/**
 * The SQL view: a query written by hand, run with the database read-only,
 * and kept under a name as it is written (Save As), then opened from the
 * Queries list as any query is. `text` is held by the window, so it is
 * still there when the view is left and come back to.
 */
export function SqlView({ shell, db, name = null, text, onText, onSaved, onResults = null, toast }) {
  const [out, setOut] = useState(null);
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [naming, setNaming] = useState(null);
  const runIt = async () => {
    setBusy(true);
    setError('');
    try { setOut(await shell.base.query({ id: db.id, sql: text })); } catch (e) { setError(e.message); setOut(null); } finally { setBusy(false); }
  };
  const save = async (as) => {
    try {
      const summary = await shell.base.saveQuery({ id: db.id, name: as, design: { sql: text } });
      setNaming(null);
      onSaved?.(summary, as);
      toast(t('Kept as {name}', { name: as }), { tone: 'good' });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  return (
    <div className="bs-sql">
      {name ? <div className="bs-tools bs-q-runbar"><Icon name="formula" size={16} /><strong className="bs-q-title">{name}</strong></div> : null}
      <textarea className="rw-input bs-sql-text" spellCheck={false} value={text} onChange={(e) => onText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); runIt(); } }} aria-label={t('Query')} />
      <div className="bs-tools">
        <Button primary icon="play" label={t('Run')} className="bs-sql-run" disabled={busy} title={t('Run the query (Ctrl+Enter). It only reads: nothing in the database changes.')} onClick={runIt} />
        <Button icon="save" label={name ? t('Save') : t('Save As')} className="bs-sql-save" disabled={!String(text).trim()} title={t('Keep the query under a name, as it is written, with the other queries')} onClick={() => (name ? save(name) : setNaming(''))} />
        {name ? <Button label={t('Save As')} className="bs-sql-save-as" disabled={!String(text).trim()} onClick={() => setNaming('')} /> : null}
        {out ? <span className="bs-count">{out.count > out.rows.length ? t('{shown} of {count} rows', { shown: out.rows.length, count: out.count }) : tn(out.count, '{count} row', '{count} rows')}</span> : null}
        <span className="bs-q-grow" />
        {onResults ? <Button icon="table" label={t('Show results')} className="bs-q-results" onClick={onResults} /> : null}
      </div>
      {error ? <p className="bs-error" role="alert">{error}</p> : null}
      {out ? (
        <div className="bs-result">
          <table>
            <thead><tr>{out.columns.map((c) => <th key={c}>{c}</th>)}</tr></thead>
            <tbody>{out.rows.map((r, i) => <tr key={i}>{out.columns.map((c) => <td key={c}>{r[c] === null ? '' : String(r[c])}</td>)}</tr>)}</tbody>
          </table>
        </div>
      ) : null}
      {naming !== null ? (
        <Dialog title={t('Save Query')} onClose={() => setNaming(null)} actions={<><Button label={t('Cancel')} onClick={() => setNaming(null)} /><Button primary label={t('Save')} className="bs-sql-save-named" disabled={!String(naming).trim()} onClick={() => save(String(naming).trim())} /></>}>
          <label className="bs-field"><span>{t('Name')}</span><Input className="rw-input bs-sql-name" aria-label={t('Name')} autoFocus value={naming} onChange={(e) => setNaming(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && String(naming).trim()) save(String(naming).trim()); }} /></label>
        </Dialog>
      ) : null}
    </div>
  );
}

/* ── a query's results, read as a table's records are ─────────────── */

const PAGE = 100;
const ROW_H = 30;
const cellText = (type, v) => {
  if (v === null || v === undefined) return '';
  if (type === 'currency' && typeof v === 'number') return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 4 });
  if (type === 'yesno') return v ? t('Yes') : t('No');
  if (type === 'datetime') return String(v).replace('T', ' ');
  return String(v);
};

/**
 * A saved query's results as the datasheet shows a table's records
 * (D-2026-10-22-07): a page at a time as it is scrolled, sorted by a column,
 * searched and filtered in the database, filter by selection from a cell.
 */
function QueryResults({ shell, db, design, onCount, toast }) {
  const [columns, setColumns] = useState([]);
  const [count, setCount] = useState(0);
  const [order, setOrder] = useState(null);
  const [error, setError] = useState('');
  const [scrollTop, setScrollTop] = useState(0);
  const [height, setHeight] = useState(400);
  const [filtering, setFiltering] = useState(null);
  const [, redraw] = useState(0);
  const pages = useRef(new Map());
  const box = useRef(null);
  const grid = useGridFilter();
  const menu = useMenu();

  const load = useCallback(async (page) => {
    if (pages.current.has(page)) return;
    pages.current.set(page, null);
    try {
      const r = await shell.base.queryRows({ id: db.id, design, offset: page * PAGE, limit: PAGE, order, search: grid.search, filters: grid.filters });
      pages.current.set(page, r.rows);
      setColumns(r.columns);
      setCount(r.count);
      setError('');
      onCount?.(r.count, grid.active);
      redraw((n) => n + 1);
    } catch (e) {
      pages.current.delete(page);
      setError(e.message);
    }
  }, [shell, db.id, design, order, grid.search, grid.filters, grid.active, onCount]);
  useEffect(() => { pages.current = new Map(); load(0); }, [load]);

  useEffect(() => {
    const el = box.current;
    if (!el) return undefined;
    const ro = new ResizeObserver(() => setHeight(el.clientHeight));
    ro.observe(el);
    setHeight(el.clientHeight);
    return () => ro.disconnect();
  }, []);
  const first = Math.max(0, Math.floor(scrollTop / ROW_H) - 5);
  const last = Math.min(count, Math.ceil((scrollTop + height) / ROW_H) + 5);
  useEffect(() => {
    for (let p = Math.floor(first / PAGE); p <= Math.floor(Math.max(first, last - 1) / PAGE); p++) load(p);
  }, [first, last, load]);
  const rowAt = (i) => pages.current.get(Math.floor(i / PAGE))?.[i % PAGE] || null;

  const sortBy = (name) => setOrder((o) => (o === name ? `-${name}` : o === `-${name}` ? null : name));
  const widths = columns.map((c) => (c.type === 'yesno' ? 90 : ['number', 'integer', 'currency'].includes(c.type) ? 130 : 180));
  const template = widths.map((w) => `${w}px`).join(' ');
  const wide = widths.reduce((a, b) => a + b, 0);
  return (
    <>
      <FilterBar grid={grid} />
      {error ? <p className="bs-error" role="alert">{error}</p> : null}
      <div className="bs-grid bs-q-grid-results" ref={box} onScroll={(e) => setScrollTop(e.currentTarget.scrollTop)} role="grid" aria-rowcount={count + 1}>
        <div className="bs-row bs-head" style={{ gridTemplateColumns: template }} role="row">
          {columns.map((c) => (
            <div key={c.name} className="bs-th-cell" role="columnheader" aria-sort={order === c.name ? 'ascending' : order === `-${c.name}` ? 'descending' : 'none'}>
              <button type="button" className="bs-th" data-field={c.name} onClick={() => sortBy(c.name)}>
                <span>{c.name}</span>
                {order === c.name ? <Icon name="chevronUp" size={12} /> : order === `-${c.name}` ? <Icon name="chevronDown" size={12} /> : null}
              </button>
              <button type="button" className={`bs-th-filter${grid.filters.some((x) => x.column === c.name) ? ' on' : ''}`} data-filter-field={c.name} title={t('Filter by {column}', { column: c.name })} aria-label={t('Filter by {column}', { column: c.name })} onClick={() => setFiltering(c)}>
                <Icon name="filter" size={12} />
              </button>
            </div>
          ))}
        </div>
        <div className="bs-body" style={{ height: count * ROW_H, width: wide }}>
          {Array.from({ length: Math.max(0, last - first) }, (_, k) => {
            const i = first + k;
            const row = rowAt(i);
            return (
              <div key={i} className="bs-row" role="row" data-row={i} style={{ top: i * ROW_H, gridTemplateColumns: template }}>
                {columns.map((c) => (
                  <span key={c.name} className={`bs-td t-${c.type}`} role="gridcell" data-field={c.name}
                    onContextMenu={(e) => { if (!row) return; e.preventDefault(); menu.open(e, selectionItems(grid, c.name, row[c.name])); }}>
                    {row ? <span className="bs-text">{cellText(c.type, row[c.name])}</span> : null}
                  </span>
                ))}
              </div>
            );
          })}
        </div>
      </div>
      {filtering ? <ColumnFilter column={filtering} onApply={grid.put} onClose={() => setFiltering(null)} /> : null}
      {menu.node}
    </>
  );
}

export const QUERY_CSS = `
.bs-query { flex: 1; min-height: 0; display: flex; flex-direction: column; overflow: auto; }
.bs-q-tables { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; padding: 12px 16px; border-bottom: 1px solid var(--line-soft); }
.bs-q-caption { font-size: 12px; font-weight: 600; color: var(--ink-3); text-transform: uppercase; letter-spacing: .04em; }
.bs-q-chip { display: inline-flex; align-items: center; gap: 2px; padding: 2px 2px 2px 10px; border: 1px solid var(--line); border-radius: 14px; background: var(--sunken); font-size: 13px; }
.bs-q-add-table { width: auto; min-width: 140px; }
.bs-q-totals { margin-inline-start: auto; }
.bs-q-grid-wrap { padding: 8px 16px; overflow-x: auto; }
.bs-q-grid { border-collapse: collapse; min-width: 720px; }
.bs-q-grid th { text-align: start; font-size: 12px; font-weight: 600; color: var(--ink-3); padding: 4px 6px; border-bottom: 1px solid var(--line); }
.bs-q-grid td { padding: 4px 6px; border-bottom: 1px solid var(--line-soft); }
.bs-q-grid .rw-select, .bs-q-grid .rw-input { width: 100%; }
.bs-q-field { min-width: 180px; }
.bs-q-expr { min-width: 260px; font-family: var(--mono, ui-monospace, monospace); font-size: 12.5px; }
.bs-q-center { text-align: center; }
.bs-q-runbar { align-items: center; }
.bs-q-title { font-size: 14px; font-weight: 600; }
.bs-q-grow { flex: 1; }
.bs-q-wait { display: grid; place-items: center; padding: 40px; }
.bs-query-run { overflow: hidden; }
.bs-q-sql-text { margin: 0 16px 8px; padding: 8px 10px; border-radius: var(--r-2); background: var(--sunken); font: 12.5px/1.5 ui-monospace, Consolas, monospace; white-space: pre-wrap; }
`;
