// Worksheets: the Power Query editor and the Queries & Connections pane.
//
// The editor shows what a query would load — its source with the applied
// steps run, the first hundred rows — and builds the steps from the
// column picked: remove it or keep only it, rename it, give it a type,
// filter, sort, replace, split, tidy its text, group by it; the steps are
// listed beside, one clicked shows the table as it stood after it, and a
// step taken away is gone. Close & Load writes the result as a table on a
// sheet of its own and keeps the query in the workbook.

import React, { useCallback, useEffect, useState } from 'react';
import { Button, Dialog, Input, Select, useMenu } from '@rutba/office-ui';
import { JOIN_KINDS } from '@rutba/sheet-view/queries';

const TYPES = [['text', 'Text'], ['integer', 'Whole number'], ['number', 'Decimal number'], ['date', 'Date'], ['boolean', 'True/False']];
const FILTERS = [['equals', 'equals'], ['notEquals', 'does not equal'], ['contains', 'contains'], ['notContains', 'does not contain'], ['beginsWith', 'begins with'], ['endsWith', 'ends with'], ['greater', 'is more than'], ['greaterOrEqual', 'is at least'], ['less', 'is less than'], ['lessOrEqual', 'is at most'], ['blank', 'is blank'], ['notBlank', 'is not blank']];
const AGGREGATES = [['count', 'Count rows'], ['sum', 'Sum'], ['average', 'Average'], ['min', 'Least'], ['max', 'Most']];

/** A refusal from the main process, without Electron's "Error invoking remote method" in front of it. */
export const cleanError = (err) => String(err?.message || err).replace(/^Error invoking remote method '[^']+': (Error: )?/, '');

// Recent Sources: the files queries read lately, kept on this computer, newest first.
const RECENT = 'sheets.recentSources';
export function recentSources() {
  try {
    const list = JSON.parse(localStorage.getItem(RECENT) || '[]');
    return Array.isArray(list) ? list.filter((r) => r?.source?.kind === 'csv' && r.source.path) : [];
  } catch {
    return [];
  }
}
export function rememberSource(entry) {
  try {
    const rest = recentSources().filter((r) => r.source.path !== entry.source.path);
    localStorage.setItem(RECENT, JSON.stringify([entry, ...rest].slice(0, 8)));
  } catch { /* no storage, no memory */ }
}

const shown = (v) => (v === null || v === undefined ? '' : typeof v === 'number' ? String(Math.round(v * 1e10) / 1e10) : String(v));

/**
 * The editor. `query` is `{ id?, name, source, steps }`: an id when it is a
 * query already loaded, whose Close & Load changes it.
 */
export function QueryEditor({ query, preview, listSources = null, onClose, onLoad }) {
  const menu = useMenu();
  // Append and Merge: the workbook's tables and queries, but this query and the table it loads.
  const [sources, setSources] = useState([]);
  useEffect(() => {
    let live = true;
    Promise.resolve(listSources?.()).then((got) => {
      if (!live || !got) return;
      const own = (got.queries || []).find((q) => q.id === query.id);
      setSources([
        ...(got.queries || []).filter((q) => q.id !== query.id).map((q) => ({ label: `Query: ${q.name}`, source: { kind: 'query', id: q.id, name: q.name } })),
        ...(got.tables || []).filter((t) => t.name !== own?.table && !(query.source?.kind === 'table' && query.source.table === t.name)).map((t) => ({ label: `Table: ${t.name} (${t.sheet})`, source: { kind: 'table', table: t.name } })),
      ]);
    }).catch(() => {});
    return () => { live = false; };
  }, [listSources, query.id, query.source]);
  const sourceOptions = sources.map((s, i) => [String(i), s.label]);
  const [name, setName] = useState(query.name || 'Query');
  const [steps, setSteps] = useState(query.steps || []);
  const [upTo, setUpTo] = useState(null);
  const [column, setColumn] = useState(null);
  const [table, setTable] = useState(null);
  const [error, setError] = useState('');
  const [ask, setAsk] = useState(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    let live = true;
    setError('');
    Promise.resolve(preview({ source: query.source, steps, upTo })).then((t) => { if (live) setTable(t); }).catch((err) => { if (live) { setTable(null); setError(cleanError(err)); } });
    return () => { live = false; };
  }, [query.source, steps, upTo, preview]);

  // A step taken while an earlier one is picked goes in after it, as Power Query inserts one.
  const add = useCallback((step) => {
    setSteps((s) => (upTo === null ? [...s, step] : [...s.slice(0, upTo), step, ...s.slice(upTo)]));
    setUpTo((u) => (u === null ? null : u + 1));
  }, [upTo]);
  const need = () => { if (!column) { setError('Click a column\'s heading first.'); return false; } return true; };
  const askFor = (title, fields, run) => setAsk({ title, fields, values: Object.fromEntries(fields.map((f) => [f.key, f.value ?? ''])), run });

  const actions = [
    ['Remove Column', () => need() && add({ kind: 'removeColumns', columns: [column] })],
    ['Remove Other Columns', () => need() && add({ kind: 'keepColumns', columns: [column] })],
    ['Rename', () => need() && askFor(`Rename "${column}"`, [{ key: 'to', label: 'New name', value: column }], (v) => add({ kind: 'renameColumn', from: column, to: v.to }))],
    ['Data Type', (e) => need() && menu.open(e, TYPES.map(([type, label]) => ({ label, run: () => add({ kind: 'changeType', column, type }) })))],
    ['Filter', () => need() && askFor(`Keep the rows where "${column}"…`, [{ key: 'op', label: 'Condition', options: FILTERS, value: 'equals' }, { key: 'value', label: 'Value' }], (v) => add({ kind: 'filterRows', column, op: v.op, value: v.value !== '' && !Number.isNaN(Number(v.value)) ? Number(v.value) : v.value }))],
    ['Sort A → Z', () => need() && add({ kind: 'sort', column, descending: false })],
    ['Sort Z → A', () => need() && add({ kind: 'sort', column, descending: true })],
    ['Replace Values', () => need() && askFor(`Replace in "${column}"`, [{ key: 'find', label: 'Value to find' }, { key: 'replace', label: 'Replace with' }], (v) => add({ kind: 'replaceValues', column, find: v.find, replace: v.replace }))],
    ['Split Column', (e) => need() && menu.open(e, [['Comma', 'comma'], ['Tab', 'tab'], ['Space', 'space'], ['Semicolon', 'semicolon']].map(([label, delimiter]) => ({ label: `At each ${label.toLowerCase()}`, run: () => add({ kind: 'splitColumn', column, delimiter }) })).concat([{ label: 'At another character…', run: () => askFor(`Split "${column}" at`, [{ key: 'delimiter', label: 'Character' }], (v) => add({ kind: 'splitColumn', column, delimiter: v.delimiter })) }]))],
    ['Format', (e) => need() && menu.open(e, [['Trim', 'trim'], ['UPPERCASE', 'upper'], ['lowercase', 'lower'], ['Capitalise Each Word', 'proper']].map(([label, how]) => ({ label, run: () => add({ kind: 'transformText', column, how }) })))],
    // Transform → Extract, as Power Query's: part of each cell's text kept.
    ['Extract', (e) => need() && menu.open(e, [
      { label: 'First characters…', run: () => askFor(`The first characters of "${column}"`, [{ key: 'count', label: 'How many', value: '3' }], (v) => add({ kind: 'extractText', column, how: 'first', count: Number(v.count) || 0 })) },
      { label: 'Last characters…', run: () => askFor(`The last characters of "${column}"`, [{ key: 'count', label: 'How many', value: '3' }], (v) => add({ kind: 'extractText', column, how: 'last', count: Number(v.count) || 0 })) },
      { label: 'Text before delimiter…', run: () => askFor(`The text of "${column}" before`, [{ key: 'delimiter', label: 'Delimiter', value: '-' }], (v) => add({ kind: 'extractText', column, how: 'before', delimiter: v.delimiter })) },
      { label: 'Text after delimiter…', run: () => askFor(`The text of "${column}" after`, [{ key: 'delimiter', label: 'Delimiter', value: '-' }], (v) => add({ kind: 'extractText', column, how: 'after', delimiter: v.delimiter })) },
    ])],
    // Transform → Fill: a blank cell takes the value above it, or below.
    ['Fill', (e) => need() && menu.open(e, [
      { label: 'Down', run: () => add({ kind: 'fillDown', columns: [column] }) },
      { label: 'Up', run: () => add({ kind: 'fillUp', columns: [column] }) },
    ])],
    // Transform → Unpivot Other Columns: this column kept; every other column's cells become rows.
    ['Unpivot Other Columns', () => need() && add({ kind: 'unpivotOthers', columns: [column] })],
    // Transform → Merge Columns: this column and another joined, with a separator.
    ['Merge Columns', () => need() && askFor(`Merge "${column}" with`, [
      { key: 'other', label: 'Column', options: (table?.columns || []).filter((c) => c !== column).map((c) => [c, c]), value: (table?.columns || []).find((c) => c !== column) || '' },
      { key: 'separator', label: 'Separator', options: [[' ', 'Space'], [', ', 'Comma'], ['-', 'Dash'], ['', 'None']], value: ' ' },
      { key: 'name', label: 'New column name', value: 'Merged' },
    ], (v) => add({ kind: 'mergeColumns', columns: [column, v.other], separator: v.separator, name: v.name || 'Merged' }))],
    ['Group By', () => need() && askFor(`Group by "${column}"`, [{ key: 'fn', label: 'Operation', options: AGGREGATES, value: 'count' }, { key: 'of', label: 'Of column', options: (table?.columns || []).map((c) => [c, c]), value: (table?.columns || []).find((c) => c !== column) || column }, { key: 'as', label: 'New column name', value: 'Count' }], (v) => add({ kind: 'groupBy', columns: [column], aggregations: [{ fn: v.fn, column: v.fn === 'count' ? null : v.of, name: v.as || undefined }] }))],
    ['Append Queries', () => (sources.length
      ? askFor('Append the rows of', [{ key: 'with', label: 'Table or query', options: sourceOptions, value: '0' }], (v) => add({ kind: 'appendQuery', with: sources[Number(v.with)].source }))
      : setError('There is no other table or query in this workbook to append.'))],
    ['Merge Queries', () => {
      if (!need()) return;
      if (!sources.length) { setError('There is no other table or query in this workbook to merge with.'); return; }
      askFor(`Merge on "${column}" with`, [{ key: 'with', label: 'Table or query', options: sourceOptions, value: '0' }, { key: 'how', label: 'Join kind', options: JOIN_KINDS, value: 'left' }], async (v) => {
        const other = sources[Number(v.with)].source;
        try {
          const cols = (await preview({ source: other, steps: [] }))?.columns || [];
          askFor(`Match "${column}" with a column of ${sources[Number(v.with)].label.replace(/^\w+: /, '')}`, [{ key: 'withOn', label: 'Its column', options: cols.map((c) => [c, c]), value: cols.find((c) => c.toLowerCase() === String(column).toLowerCase()) || cols[0] }], (w) => add({ kind: 'mergeQueries', with: other, on: column, withOn: w.withOn, how: v.how }));
        } catch (err) { setError(cleanError(err)); }
      });
    }],
    ['Keep Top Rows', () => askFor('Keep the first rows', [{ key: 'count', label: 'Number of rows', value: '10' }], (v) => add({ kind: 'keepTopRows', count: Number(v.count) || 0 }))],
    ['Remove Duplicates', () => add({ kind: 'removeDuplicates', ...(column ? { columns: [column] } : {}) })],
    ['Remove Blank Rows', () => add({ kind: 'removeBlankRows' })],
    ['Use First Row as Headers', () => add({ kind: 'promoteHeaders' })],
    ['Index Column', () => add({ kind: 'addIndex', name: 'Index', start: 1 })],
  ];

  const load = async () => {
    setBusy(true);
    try { await onLoad({ id: query.id, name: name.trim() || 'Query', source: query.source, sourceText: query.sourceText, steps }); } catch (err) { setError(cleanError(err)); setBusy(false); }
  };

  return (
    <>
      {menu.node}
      <Dialog
        title={`Power Query Editor — ${query.sourceText || ''}`}
        width={980}
        onClose={onClose}
        actions={(
          <>
            <Input className="pq-name" value={name} onChange={(e) => setName(e.target.value)} style={{ width: 220 }} title="The query's name, and its sheet's" />
            <Button label="Cancel" onClick={onClose} />
            <Button primary className="pq-load" label={busy ? 'Loading…' : 'Close & Load'} disabled={busy || Boolean(error && !table)} onClick={load} />
          </>
        )}
      >
        <style>{QUERIES_CSS}</style>
        <div className="pq">
          <div className="pq-tools">
            {actions.map(([label, run]) => <Button key={label} className="pq-tool" label={label} onClick={(e) => { setError(''); run(e); }} />)}
          </div>
          {ask ? (
            <div className="pq-ask">
              <b>{ask.title}</b>
              {ask.fields.map((f) => (
                <label key={f.key}>
                  <span>{f.label}</span>
                  {f.options ? (
                    <Select className={`pq-ask-${f.key}`} value={ask.values[f.key]} onChange={(e) => setAsk((a) => ({ ...a, values: { ...a.values, [f.key]: e.target.value } }))}>
                      {f.options.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                    </Select>
                  ) : (
                    <Input className={`pq-ask-${f.key}`} value={ask.values[f.key]} autoFocus={f === ask.fields[0]} onChange={(e) => setAsk((a) => ({ ...a, values: { ...a.values, [f.key]: e.target.value } }))} onKeyDown={(e) => { if (e.key === 'Enter') { ask.run(ask.values); setAsk(null); } }} />
                  )}
                </label>
              ))}
              <Button primary className="pq-ask-ok" label="OK" onClick={() => { ask.run(ask.values); setAsk(null); }} />
              <Button label="Cancel" onClick={() => setAsk(null)} />
            </div>
          ) : null}
          <div className="pq-body">
            <div className="pq-grid">
              {error ? <div className="pq-error">{error}</div> : null}
              {table ? (
                <table>
                  <thead>
                    <tr>
                      <th className="pq-rownum" />
                      {table.columns.map((c) => <th key={c} className={c === column ? 'sel' : ''} data-col={c} onClick={() => setColumn(c)} title="Pick this column for the next step">{c}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {table.rows.map((r, i) => (
                      <tr key={i}>
                        <td className="pq-rownum">{i + 1}</td>
                        {table.columns.map((c, j) => <td key={c} className={c === column ? 'sel' : ''}>{shown(r[j])}</td>)}
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : !error ? <div className="pq-note">Reading the source…</div> : null}
            </div>
            <div className="pq-steps">
              <div className="pq-steps-head">Applied steps</div>
              <button type="button" className={`pq-step${upTo === 0 ? ' at' : ''}`} onClick={() => setUpTo(0)}>Source</button>
              {(table?.steps || steps.map((s) => s.kind)).map((label, i) => (
                <div key={i} className={`pq-step-row${upTo === i + 1 || (upTo === null && i === steps.length - 1) ? ' at' : ''}`}>
                  <button type="button" className="pq-step" onClick={() => setUpTo(i + 1 === steps.length ? null : i + 1)}>{label}</button>
                  <button type="button" className="pq-step-x" title="Take this step away" onClick={() => { setSteps((s) => s.filter((_, k) => k !== i)); setUpTo(null); }}>×</button>
                </div>
              ))}
              <div className="pq-count">{table ? `${table.total} row${table.total === 1 ? '' : 's'}${table.total > table.rows.length ? `, the first ${table.rows.length} shown` : ''}` : ''}</div>
            </div>
          </div>
        </div>
      </Dialog>
    </>
  );
}

/** Data → Queries & Connections: the workbook's queries, each refreshed, edited or deleted from here. */
export function QueriesPane({ queries, onRefresh, onEdit, onDelete, onRefreshAll, onGoTo }) {
  if (!queries?.length) return <div className="pq-note"><style>{QUERIES_CSS}</style>No queries in this workbook yet. Data → From Table/Range or From Text/CSV makes one.</div>;
  return (
    <div className="pq-pane">
      <style>{QUERIES_CSS}</style>
      <Button icon="refresh" className="pq-refresh-all" label="Refresh All" onClick={onRefreshAll} />
      {queries.map((q) => (
        <div key={q.id} className="pq-item" data-query={q.id}>
          <button type="button" className="pq-item-name" onClick={() => onGoTo(q)} title="Go to its sheet">{q.name}</button>
          <div className="pq-item-meta">{q.loaded != null ? `${q.loaded} row${q.loaded === 1 ? '' : 's'} loaded` : 'Not loaded'} · {q.sourceText}</div>
          <div className="pq-item-tools">
            <Button icon="refresh" className="pq-refresh" title="Refresh — run it again on its source as it is now" onClick={() => onRefresh(q)} />
            <Button icon="textbox" className="pq-edit" title="Edit — open it in the Power Query Editor" onClick={() => onEdit(q)} />
            <Button icon="trash" className="pq-delete" title="Delete the query; its sheet stays" onClick={() => onDelete(q)} />
          </div>
        </div>
      ))}
    </div>
  );
}

export const QUERIES_CSS = `
.pq { display: flex; flex-direction: column; gap: 8px; min-height: 480px; }
.pq-tools { display: flex; flex-wrap: wrap; gap: 4px; }
.pq-tool { font-size: 12px; }
.pq-ask { display: flex; flex-wrap: wrap; align-items: end; gap: 8px; padding: 8px; border: 1px solid var(--accent); border-radius: 6px; }
.pq-ask label { display: flex; flex-direction: column; gap: 2px; font-size: 12px; }
.pq-body { display: flex; gap: 8px; min-height: 0; flex: 1; }
.pq-grid { flex: 1; overflow: auto; max-height: 460px; border: 1px solid var(--line); border-radius: 4px; }
.pq-grid table { border-collapse: collapse; font-size: 12px; width: max-content; min-width: 100%; }
.pq-grid th, .pq-grid td { border: 1px solid var(--line); padding: 3px 8px; text-align: start; white-space: nowrap; }
.pq-grid th { position: sticky; top: 0; background: var(--surface-2, #f3f4f6); cursor: pointer; font-weight: 600; }
.pq-grid th.sel, .pq-grid td.sel { background: color-mix(in srgb, var(--accent) 14%, transparent); }
.pq-rownum { color: var(--muted, #6b7280); width: 1%; }
.pq-steps { width: 230px; display: flex; flex-direction: column; gap: 2px; border-left: 1px solid var(--line); padding-left: 8px; }
.pq-steps-head { font-weight: 600; margin-bottom: 4px; }
.pq-step-row { display: flex; align-items: center; }
.pq-step { flex: 1; text-align: start; background: none; border: 0; padding: 3px 6px; border-radius: 4px; color: inherit; font: inherit; font-size: 12px; cursor: pointer; }
.pq-step-row.at .pq-step, .pq-step.at { background: color-mix(in srgb, var(--accent) 16%, transparent); }
.pq-step-x { background: none; border: 0; color: inherit; cursor: pointer; opacity: .6; }
.pq-step-x:hover { opacity: 1; }
.pq-count { margin-top: auto; font-size: 12px; opacity: .75; }
.pq-error { color: var(--bad, #c62828); padding: 8px; }
.pq-note { padding: 12px; opacity: .75; font-size: 13px; }
.pq-pane { display: flex; flex-direction: column; gap: 8px; padding: 8px; }
.pq-item { border: 1px solid var(--line); border-radius: 6px; padding: 6px 8px; }
.pq-item-name { background: none; border: 0; padding: 0; font-weight: 600; color: inherit; cursor: pointer; }
.pq-item-meta { font-size: 12px; opacity: .75; margin: 2px 0 4px; }
.pq-item-tools { display: flex; gap: 4px; }
`;
