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
import { Button, Dialog, Input, Select, useMenu, t, tn } from '@rutba/office-ui';
import { JOIN_KINDS } from '@rutba/sheet-view/queries';

const TYPES = [['text', t('common.text')], ['integer', t('common.whole-number')], ['number', t('sheets.queries.decimal-number')], ['date', t('common.date')], ['boolean', t('sheets.queries.true-false')]];
const FILTERS = [['equals', t('sheets.queries.equals')], ['notEquals', t('sheets.queries.does-not-equal')], ['contains', t('common.contains')], ['notContains', t('common.does-not-contain')], ['beginsWith', t('sheets.queries.begins-with')], ['endsWith', t('common.ends-with')], ['greater', t('sheets.queries.is-more-than')], ['greaterOrEqual', t('sheets.queries.is-at-least')], ['less', t('sheets.queries.is-less-than')], ['lessOrEqual', t('sheets.queries.is-at-most')], ['blank', t('sheets.queries.is-blank')], ['notBlank', t('sheets.queries.is-not-blank')]];
const AGGREGATES = [['count', t('sheets.queries.count-rows')], ['sum', t('common.sum')], ['average', t('common.average')], ['min', t('sheets.queries.least')], ['max', t('sheets.queries.most')]];

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

/** Another table a step reads, in words: the engine's sourceWords, in the window's language. */
function sourceWords(src) {
  if (!src) return t('sheets.queries.a-table');
  if (src.kind === 'query') return t('sheets.queries.query-name', { name: src.name || src.id });
  if (src.kind === 'table') return t('sheets.queries.table-name', { name: src.table });
  if (src.kind === 'range') return `${src.sheet}!${src.ref}`;
  if (src.kind === 'csv') return String(src.path || '').split(/[\\/]/).pop();
  return src.kind;
}

/**
 * What a step says it does, for Applied Steps: the engine's describeStep
 * (@rutba/sheet-view/queries), each in a message of its own so a language
 * can put the column, the value and the verb in its own order.
 */
export function stepLabel(s) {
  const cols = (list) => (list || []).map((c) => `"${c}"`).join(', ');
  const column = s.column;
  switch (s.kind) {
    case 'removeColumns': return t('sheets.queries.removed-columns', { columns: cols(s.columns) });
    case 'keepColumns': return t('sheets.queries.kept-columns', { columns: cols(s.columns) });
    case 'renameColumn': return t('sheets.queries.renamed', { from: s.from, to: s.to });
    case 'changeType':
      return s.type === 'integer' ? t('sheets.queries.column-made-whole', { column })
        : s.type === 'number' ? t('sheets.queries.column-made-numbers', { column })
          : s.type === 'boolean' ? t('sheets.queries.column-made-true', { column })
            : s.type === 'date' ? t('sheets.queries.column-made-dates', { column })
              : t('sheets.queries.column-made-text', { column });
    case 'filterRows': {
      const value = JSON.stringify(s.value);
      switch (s.op) {
        case 'equals': return t('sheets.queries.rows-column-value', { column, value });
        case 'notEquals': return t('sheets.queries.rows-column-not-value', { column, value });
        case 'contains': return t('sheets.queries.rows-column-contains', { column, value });
        case 'notContains': return t('sheets.queries.rows-column-does', { column, value });
        case 'beginsWith': return t('sheets.queries.rows-column-begins', { column, value });
        case 'endsWith': return t('sheets.queries.rows-column-ends', { column, value });
        case 'greater': return t('sheets.queries.rows-column-more', { column, value });
        case 'greaterOrEqual': return t('sheets.queries.rows-column-least', { column, value });
        case 'less': return t('sheets.queries.rows-column-less', { column, value });
        case 'lessOrEqual': return t('sheets.queries.rows-column-most', { column, value });
        case 'blank': return t('sheets.queries.rows-column-blank', { column });
        case 'notBlank': return t('sheets.queries.rows-column-not', { column });
        default: return `Rows where "${column}" ${s.op} ${value}`; // words-ok: an operation no version of the editor makes
      }
    }
    case 'sort': {
      const by = s.by?.[0] || s;
      return by.descending ? t('sheets.queries.sorted-column-z', { column: by.column }) : t('sheets.queries.sorted-by-column', { column: by.column });
    }
    case 'removeDuplicates': return s.columns?.length ? t('sheets.queries.duplicates-columns', { columns: cols(s.columns) }) : t('sheets.queries.duplicate-rows');
    case 'removeBlankRows': return t('sheets.queries.blank-rows-removed');
    case 'keepTopRows': return t('sheets.queries.first-count-rows', { count: s.count });
    case 'replaceValues': return t('sheets.queries.column-find-replaced', { column, find: JSON.stringify(s.find), replace: JSON.stringify(s.replace) });
    case 'splitColumn':
      return s.delimiter === 'tab' ? t('sheets.queries.column-split-each', { column })
        : s.delimiter === 'space' ? t('sheets.queries.column-split-each-space', { column })
          : s.delimiter === 'comma' ? t('sheets.queries.column-split-each-comma', { column })
            : s.delimiter === 'semicolon' ? t('sheets.queries.column-split-each-semicolon', { column })
              : t('sheets.queries.column-split-delimiter', { column, delimiter: JSON.stringify(s.delimiter) });
    case 'groupBy': return t('sheets.queries.grouped-by-columns', { columns: cols(s.columns || [s.column]) });
    case 'addIndex': return t('sheets.queries.index-column-name', { name: s.name || t('common.index') });
    case 'transformText':
      return s.how === 'trim' ? t('sheets.queries.column-trimmed', { column })
        : s.how === 'upper' ? t('sheets.queries.column-in-capitals', { column })
          : s.how === 'lower' ? t('sheets.queries.column-small-letters', { column })
            : t('sheets.queries.column-title-case', { column });
    case 'promoteHeaders': return t('sheets.queries.first-row-used');
    case 'conditionalColumn': return t('sheets.queries.conditional-column-name', { name: s.name || t('common.custom') });
    case 'pivotColumn': {
      const values = s.values;
      if (!s.fn || s.fn === 'sum') return t('sheets.queries.column-pivoted-values', { column, values });
      if (s.fn === 'none') return t('sheets.queries.column-pivoted-values-values-not', { column, values });
      const fn = { count: t('sheets.queries.count'), average: t('sheets.queries.average'), min: t('sheets.queries.min'), max: t('sheets.queries.max') }[s.fn] || s.fn;
      return t('sheets.queries.column-pivoted-values-values', { column, values, fn });
    }
    case 'fillDown': return t('sheets.queries.columns-filled-down', { columns: cols(s.columns || [s.column]) });
    case 'fillUp': return t('sheets.queries.columns-filled-up', { columns: cols(s.columns || [s.column]) });
    case 'unpivotOthers': return t('sheets.queries.columns-other-columns', { columns: cols(s.columns || [s.column]) });
    case 'mergeColumns': return t('sheets.queries.columns-merged-name', { columns: cols(s.columns), name: s.name || t('sheets.queries.merged') });
    case 'extractText': {
      const delimiter = JSON.stringify(s.delimiter);
      return s.how === 'first' ? t('sheets.queries.column-first-count', { column, count: s.count })
        : s.how === 'last' ? t('sheets.queries.column-last-count', { column, count: s.count })
          : s.how === 'before' ? t('sheets.queries.column-text-before', { column, delimiter })
            : s.how === 'after' ? t('sheets.queries.column-text-after', { column, delimiter })
              : t('sheets.queries.column-text-kept', { column });
    }
    case 'appendQuery': return t('sheets.queries.appended-source', { source: sourceWords(s.with) });
    case 'mergeQueries': {
      const values = { source: sourceWords(s.with), on: s.on, withOn: s.withOn };
      return s.how === 'inner' ? t('sheets.queries.merged-source-inner', values)
        : s.how === 'leftAnti' ? t('sheets.queries.merged-source-left', values)
          : s.how === 'full' ? t('sheets.queries.merged-source-full', values)
            : t('sheets.queries.merged-source', values);
    }
    default: return s.kind;
  }
}

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
        ...(got.queries || []).filter((q) => q.id !== query.id).map((q) => ({ label: t('common.query-name', { name: q.name }), name: q.name, source: { kind: 'query', id: q.id, name: q.name } })),
        ...(got.tables || []).filter((t) => t.name !== own?.table && !(query.source?.kind === 'table' && query.source.table === t.name)).map((tb) => ({ label: t('sheets.queries.table-name-sheet', { name: tb.name, sheet: tb.sheet }), name: `${tb.name} (${tb.sheet})`, source: { kind: 'table', table: tb.name } })),
      ]);
    }).catch(() => {});
    return () => { live = false; };
  }, [listSources, query.id, query.source]);
  const sourceOptions = sources.map((s, i) => [String(i), s.label]);
  const [name, setName] = useState(query.name || t('common.query'));
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
  const need = () => { if (!column) { setError(t('sheets.queries.click-columns-heading')); return false; } return true; };
  const askFor = (title, fields, run) => setAsk({ title, fields, values: Object.fromEntries(fields.map((f) => [f.key, f.value ?? ''])), run });

  const actions = [
    [t('sheets.queries.remove-column'), () => need() && add({ kind: 'removeColumns', columns: [column] })],
    [t('sheets.queries.remove-other-columns'), () => need() && add({ kind: 'keepColumns', columns: [column] })],
    [t('common.rename'), () => need() && askFor(t('sheets.queries.rename-column', { column }), [{ key: 'to', label: t('sheets.queries.new-name'), value: column }], (v) => add({ kind: 'renameColumn', from: column, to: v.to }))],
    [t('sheets.queries.data-type'), (e) => need() && menu.open(e, TYPES.map(([type, label]) => ({ label, run: () => add({ kind: 'changeType', column, type }) })))],
    [t('common.filter'), () => need() && askFor(t('sheets.queries.keep-rows-column', { column }), [{ key: 'op', label: t('sheets.queries.condition'), options: FILTERS, value: 'equals' }, { key: 'value', label: t('common.value') }], (v) => add({ kind: 'filterRows', column, op: v.op, value: v.value !== '' && !Number.isNaN(Number(v.value)) ? Number(v.value) : v.value }))],
    [t('sheets.queries.sort-a-z'), () => need() && add({ kind: 'sort', column, descending: false })],
    [t('sheets.queries.sort-z-a'), () => need() && add({ kind: 'sort', column, descending: true })],
    [t('sheets.queries.replace-values'), () => need() && askFor(t('sheets.queries.replace-in-column', { column }), [{ key: 'find', label: t('sheets.queries.value-to-find') }, { key: 'replace', label: t('common.replace-with') }], (v) => add({ kind: 'replaceValues', column, find: v.find, replace: v.replace }))],
    [t('sheets.queries.split-column'), (e) => need() && menu.open(e, [[t('sheets.queries.at-each-comma'), 'comma'], [t('sheets.queries.at-each-tab'), 'tab'], [t('sheets.queries.at-each-space'), 'space'], [t('sheets.queries.at-each-semicolon'), 'semicolon']].map(([label, delimiter]) => ({ label, run: () => add({ kind: 'splitColumn', column, delimiter }) })).concat([{ label: t('sheets.queries.at-another'), run: () => askFor(t('sheets.queries.split-column-at', { column }), [{ key: 'delimiter', label: t('sheets.queries.character') }], (v) => add({ kind: 'splitColumn', column, delimiter: v.delimiter })) }]))],
    [t('common.format'), (e) => need() && menu.open(e, [[t('common.trim'), 'trim'], [t('common.uppercase'), 'upper'], [t('common.lowercase-2'), 'lower'], [t('common.capitalise-each-word'), 'proper']].map(([label, how]) => ({ label, run: () => add({ kind: 'transformText', column, how }) })))],
    // Transform → Extract, as Power Query's: part of each cell's text kept.
    [t('common.extract'), (e) => need() && menu.open(e, [
      { label: t('sheets.queries.first-characters'), run: () => askFor(t('sheets.queries.first-characters-column', { column }), [{ key: 'count', label: t('sheets.queries.how-many'), value: '3' }], (v) => add({ kind: 'extractText', column, how: 'first', count: Number(v.count) || 0 })) },
      { label: t('sheets.queries.last-characters'), run: () => askFor(t('sheets.queries.last-characters-column', { column }), [{ key: 'count', label: t('sheets.queries.how-many'), value: '3' }], (v) => add({ kind: 'extractText', column, how: 'last', count: Number(v.count) || 0 })) },
      { label: t('sheets.queries.text-before'), run: () => askFor(t('sheets.queries.text-column-before', { column }), [{ key: 'delimiter', label: t('sheets.queries.delimiter'), value: '-' }], (v) => add({ kind: 'extractText', column, how: 'before', delimiter: v.delimiter })) },
      { label: t('sheets.queries.text-after'), run: () => askFor(t('sheets.queries.text-column-after', { column }), [{ key: 'delimiter', label: t('sheets.queries.delimiter'), value: '-' }], (v) => add({ kind: 'extractText', column, how: 'after', delimiter: v.delimiter })) },
    ])],
    // Transform → Fill: a blank cell takes the value above it, or below.
    [t('common.fill'), (e) => need() && menu.open(e, [
      { label: t('common.down'), run: () => add({ kind: 'fillDown', columns: [column] }) },
      { label: t('common.up'), run: () => add({ kind: 'fillUp', columns: [column] }) },
    ])],
    // Transform → Pivot Column: this column's values become columns, filled from another, added up.
    [t('sheets.queries.pivot-column'), () => need() && askFor(t('sheets.queries.pivot-column-2', { column }), [
      { key: 'values', label: t('sheets.queries.values-from'), options: (table?.columns || []).filter((c) => c !== column).map((c) => [c, c]), value: [...(table?.columns || [])].reverse().find((c) => c !== column) || '' },
      { key: 'fn', label: t('sheets.queries.add-up-as'), options: [['sum', t('common.sum')], ['count', t('common.count')], ['average', t('common.average')], ['min', t('sheets.minimum')], ['max', t('sheets.maximum')], ['none', t('sheets.queries.do-not-add')]], value: 'sum' },
    ], (v) => add({ kind: 'pivotColumn', column, values: v.values, fn: v.fn }))],
    // Add Column → Conditional Column: if a column meets a condition then one value, otherwise another.
    [t('sheets.queries.conditional'), () => askFor(t('sheets.queries.add-conditional-column'), [
      { key: 'name', label: t('sheets.queries.new-column-name'), value: t('common.custom') },
      { key: 'column', label: t('sheets.queries.if-column'), options: (table?.columns || []).map((c) => [c, c]), value: column || table?.columns?.[0] || '' },
      { key: 'op', label: t('sheets.queries.is'), options: FILTERS, value: 'equals' },
      { key: 'value', label: t('common.value') },
      { key: 'output', label: t('sheets.queries.then') },
      { key: 'otherwise', label: t('sheets.queries.otherwise') },
    ], (v) => {
      const typedOf = (x) => (x !== '' && !Number.isNaN(Number(x)) ? Number(x) : x);
      add({ kind: 'conditionalColumn', name: v.name || t('common.custom'), rules: [{ column: v.column, op: v.op, value: typedOf(v.value), output: v.output === '' ? null : typedOf(v.output) }], otherwise: v.otherwise === '' ? null : typedOf(v.otherwise) });
    })],
    // Transform → Unpivot Other Columns: this column kept; every other column's cells become rows.
    [t('sheets.queries.unpivot-other'), () => need() && add({ kind: 'unpivotOthers', columns: [column] })],
    // Transform → Merge Columns: this column and another joined, with a separator.
    [t('sheets.queries.merge-columns'), () => need() && askFor(t('sheets.queries.merge-column-with', { column }), [
      { key: 'other', label: t('common.column'), options: (table?.columns || []).filter((c) => c !== column).map((c) => [c, c]), value: (table?.columns || []).find((c) => c !== column) || '' },
      { key: 'separator', label: t('sheets.queries.separator'), options: [[' ', t('common.space')], [', ', t('sheets.comma')], ['-', t('common.dash')], ['', t('common.none')]], value: ' ' },
      { key: 'name', label: t('sheets.queries.new-column-name'), value: t('sheets.queries.merged') },
    ], (v) => add({ kind: 'mergeColumns', columns: [column, v.other], separator: v.separator, name: v.name || t('sheets.queries.merged') }))],
    [t('common.group-by'), () => need() && askFor(t('sheets.queries.group-by-column', { column }), [{ key: 'fn', label: t('sheets.operation'), options: AGGREGATES, value: 'count' }, { key: 'of', label: t('sheets.queries.of-column'), options: (table?.columns || []).map((c) => [c, c]), value: (table?.columns || []).find((c) => c !== column) || column }, { key: 'as', label: t('sheets.queries.new-column-name'), value: t('common.count') }], (v) => add({ kind: 'groupBy', columns: [column], aggregations: [{ fn: v.fn, column: v.fn === 'count' ? null : v.of, name: v.as || undefined }] }))],
    [t('sheets.queries.append-queries'), () => (sources.length
      ? askFor(t('sheets.queries.append-rows'), [{ key: 'with', label: t('common.table-or-query'), options: sourceOptions, value: '0' }], (v) => add({ kind: 'appendQuery', with: sources[Number(v.with)].source }))
      : setError(t('sheets.queries.no-other-table')))],
    [t('sheets.queries.merge-queries'), () => {
      if (!need()) return;
      if (!sources.length) { setError(t('sheets.queries.no-other-table-query')); return; }
      askFor(t('sheets.queries.merge-column', { column }), [{ key: 'with', label: t('common.table-or-query'), options: sourceOptions, value: '0' }, { key: 'how', label: t('sheets.queries.join-kind'), options: JOIN_KINDS.map(([kind, label]) => [kind, t(label)]), value: 'left' }], async (v) => {
        const other = sources[Number(v.with)].source;
        try {
          const cols = (await preview({ source: other, steps: [] }))?.columns || [];
          askFor(t('sheets.queries.match-column-column', { column, source: sources[Number(v.with)].name }), [{ key: 'withOn', label: t('sheets.queries.its-column'), options: cols.map((c) => [c, c]), value: cols.find((c) => c.toLowerCase() === String(column).toLowerCase()) || cols[0] }], (w) => add({ kind: 'mergeQueries', with: other, on: column, withOn: w.withOn, how: v.how }));
        } catch (err) { setError(cleanError(err)); }
      });
    }],
    [t('sheets.queries.keep-top-rows'), () => askFor(t('sheets.queries.keep-first-rows'), [{ key: 'count', label: t('sheets.queries.number-of-rows'), value: '10' }], (v) => add({ kind: 'keepTopRows', count: Number(v.count) || 0 }))],
    [t('sheets.remove'), () => add({ kind: 'removeDuplicates', ...(column ? { columns: [column] } : {}) })],
    [t('sheets.queries.remove-blank-rows'), () => add({ kind: 'removeBlankRows' })],
    [t('sheets.queries.use-first-row'), () => add({ kind: 'promoteHeaders' })],
    [t('sheets.queries.index-column'), () => add({ kind: 'addIndex', name: t('common.index'), start: 1 })],
  ];

  const load = async () => {
    setBusy(true);
    try { await onLoad({ id: query.id, name: name.trim() || t('common.query'), source: query.source, sourceText: query.sourceText, steps }); } catch (err) { setError(cleanError(err)); setBusy(false); }
  };

  return (
    <>
      {menu.node}
      <Dialog
        title={t('sheets.queries.power-query-editor', { source: query.sourceText || '' })}
        width={980}
        onClose={onClose}
        actions={(
          <>
            <Input className="pq-name" value={name} onChange={(e) => setName(e.target.value)} style={{ width: 220 }} title={t('sheets.queries.querys-name-sheets')} />
            <Button label={t('common.cancel')} onClick={onClose} />
            <Button primary className="pq-load" label={busy ? t('sheets.queries.loading') : t('sheets.queries.close-and')} disabled={busy || Boolean(error && !table)} onClick={load} />
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
              <Button primary className="pq-ask-ok" label={t('common.ok')} onClick={() => { ask.run(ask.values); setAsk(null); }} />
              <Button label={t('common.cancel')} onClick={() => setAsk(null)} />
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
                      {table.columns.map((c) => <th key={c} className={c === column ? 'sel' : ''} data-col={c} onClick={() => setColumn(c)} title={t('sheets.queries.pick-column-next')}>{c}</th>)}
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
              ) : !error ? <div className="pq-note">{t('sheets.queries.reading-the-source')}</div> : null}
            </div>
            <div className="pq-steps">
              <div className="pq-steps-head">{t('sheets.queries.applied-steps')}</div>
              <button type="button" className={`pq-step${upTo === 0 ? ' at' : ''}`} onClick={() => setUpTo(0)}>{t('common.source')}</button>
              {steps.map(stepLabel).map((label, i) => (
                <div key={i} className={`pq-step-row${upTo === i + 1 || (upTo === null && i === steps.length - 1) ? ' at' : ''}`}>
                  <button type="button" className="pq-step" onClick={() => setUpTo(i + 1 === steps.length ? null : i + 1)}>{label}</button>
                  <button type="button" className="pq-step-x" title={t('sheets.queries.take-step-away')} onClick={() => { setSteps((s) => s.filter((_, k) => k !== i)); setUpTo(null); }}>×</button>
                </div>
              ))}
              <div className="pq-count">{table ? (table.total > table.rows.length ? t('sheets.queries.rows-first-shown', { rows: tn(table.total, 'common.count-row'), shown: table.rows.length }) : tn(table.total, 'common.count-row')) : ''}</div>
            </div>
          </div>
        </div>
      </Dialog>
    </>
  );
}

/** Data → Queries & Connections: the workbook's queries, each refreshed, edited or deleted from here. */
export function QueriesPane({ queries, onRefresh, onEdit, onDelete, onRefreshAll, onGoTo }) {
  if (!queries?.length) return <div className="pq-note"><style>{QUERIES_CSS}</style>{t('sheets.queries.no-queries-workbook')}</div>;
  return (
    <div className="pq-pane">
      <style>{QUERIES_CSS}</style>
      <Button icon="refresh" className="pq-refresh-all" label={t('sheets.refresh-all')} onClick={onRefreshAll} />
      {queries.map((q) => (
        <div key={q.id} className="pq-item" data-query={q.id}>
          <button type="button" className="pq-item-name" onClick={() => onGoTo(q)} title={t('sheets.queries.go-sheet')}>{q.name}</button>
          <div className="pq-item-meta">{q.loaded != null ? tn(q.loaded, 'sheets.queries.count-row-loaded') : t('sheets.queries.not-loaded')} · {q.sourceText}</div>
          <div className="pq-item-tools">
            <Button icon="refresh" className="pq-refresh" title={t('sheets.queries.refresh-run-again')} onClick={() => onRefresh(q)} />
            <Button icon="textbox" className="pq-edit" title={t('sheets.queries.edit-open-power')} onClick={() => onEdit(q)} />
            <Button icon="trash" className="pq-delete" title={t('sheets.queries.delete-query-sheet')} onClick={() => onDelete(q)} />
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
