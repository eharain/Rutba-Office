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

const TYPES = [['text', t('Text')], ['integer', t('Whole number')], ['number', t('Decimal number')], ['date', t('Date')], ['boolean', t('True/False')]];
const FILTERS = [['equals', t('equals')], ['notEquals', t('does not equal')], ['contains', t('contains')], ['notContains', t('does not contain')], ['beginsWith', t('begins with')], ['endsWith', t('ends with')], ['greater', t('is more than')], ['greaterOrEqual', t('is at least')], ['less', t('is less than')], ['lessOrEqual', t('is at most')], ['blank', t('is blank')], ['notBlank', t('is not blank')]];
const AGGREGATES = [['count', t('Count rows')], ['sum', t('Sum')], ['average', t('Average')], ['min', t('Least')], ['max', t('Most')]];

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
  if (!src) return t('a table');
  if (src.kind === 'query') return t('query {name}', { name: src.name || src.id });
  if (src.kind === 'table') return t('table {name}', { name: src.table });
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
    case 'removeColumns': return t('Removed {columns}', { columns: cols(s.columns) });
    case 'keepColumns': return t('Kept {columns}', { columns: cols(s.columns) });
    case 'renameColumn': return t('Renamed "{from}" to "{to}"', { from: s.from, to: s.to });
    case 'changeType':
      return s.type === 'integer' ? t('"{column}" made whole numbers', { column })
        : s.type === 'number' ? t('"{column}" made numbers', { column })
          : s.type === 'boolean' ? t('"{column}" made true or false', { column })
            : s.type === 'date' ? t('"{column}" made dates', { column })
              : t('"{column}" made text', { column });
    case 'filterRows': {
      const value = JSON.stringify(s.value);
      switch (s.op) {
        case 'equals': return t('Rows where "{column}" is {value}', { column, value });
        case 'notEquals': return t('Rows where "{column}" is not {value}', { column, value });
        case 'contains': return t('Rows where "{column}" contains {value}', { column, value });
        case 'notContains': return t('Rows where "{column}" does not contain {value}', { column, value });
        case 'beginsWith': return t('Rows where "{column}" begins with {value}', { column, value });
        case 'endsWith': return t('Rows where "{column}" ends with {value}', { column, value });
        case 'greater': return t('Rows where "{column}" is more than {value}', { column, value });
        case 'greaterOrEqual': return t('Rows where "{column}" is at least {value}', { column, value });
        case 'less': return t('Rows where "{column}" is less than {value}', { column, value });
        case 'lessOrEqual': return t('Rows where "{column}" is at most {value}', { column, value });
        case 'blank': return t('Rows where "{column}" is blank', { column });
        case 'notBlank': return t('Rows where "{column}" is not blank', { column });
        default: return `Rows where "${column}" ${s.op} ${value}`; // words-ok: an operation no version of the editor makes
      }
    }
    case 'sort': {
      const by = s.by?.[0] || s;
      return by.descending ? t('Sorted by "{column}", Z to A', { column: by.column }) : t('Sorted by "{column}"', { column: by.column });
    }
    case 'removeDuplicates': return s.columns?.length ? t('Duplicates of {columns} removed', { columns: cols(s.columns) }) : t('Duplicate rows removed');
    case 'removeBlankRows': return t('Blank rows removed');
    case 'keepTopRows': return t('The first {count} rows kept', { count: s.count });
    case 'replaceValues': return t('In "{column}", {find} replaced with {replace}', { column, find: JSON.stringify(s.find), replace: JSON.stringify(s.replace) });
    case 'splitColumn':
      return s.delimiter === 'tab' ? t('"{column}" split at each tab', { column })
        : s.delimiter === 'space' ? t('"{column}" split at each space', { column })
          : s.delimiter === 'comma' ? t('"{column}" split at each comma', { column })
            : s.delimiter === 'semicolon' ? t('"{column}" split at each semicolon', { column })
              : t('"{column}" split at {delimiter}', { column, delimiter: JSON.stringify(s.delimiter) });
    case 'groupBy': return t('Grouped by {columns}', { columns: cols(s.columns || [s.column]) });
    case 'addIndex': return t('Index column "{name}" added', { name: s.name || t('Index') });
    case 'transformText':
      return s.how === 'trim' ? t('"{column}" trimmed', { column })
        : s.how === 'upper' ? t('"{column}" in capitals', { column })
          : s.how === 'lower' ? t('"{column}" in small letters', { column })
            : t('"{column}" in title case', { column });
    case 'promoteHeaders': return t('First row used as headers');
    case 'conditionalColumn': return t('Conditional column "{name}" added', { name: s.name || t('Custom') });
    case 'pivotColumn': {
      const values = s.values;
      if (!s.fn || s.fn === 'sum') return t('"{column}" pivoted, its values from "{values}"', { column, values });
      if (s.fn === 'none') return t('"{column}" pivoted, its values from "{values}" (not added up)', { column, values });
      const fn = { count: t('count'), average: t('average'), min: t('min'), max: t('max') }[s.fn] || s.fn;
      return t('"{column}" pivoted, its values from "{values}" ({fn})', { column, values, fn });
    }
    case 'fillDown': return t('{columns} filled down', { columns: cols(s.columns || [s.column]) });
    case 'fillUp': return t('{columns} filled up', { columns: cols(s.columns || [s.column]) });
    case 'unpivotOthers': return t('Columns other than {columns} unpivoted', { columns: cols(s.columns || [s.column]) });
    case 'mergeColumns': return t('{columns} merged into "{name}"', { columns: cols(s.columns), name: s.name || t('Merged') });
    case 'extractText': {
      const delimiter = JSON.stringify(s.delimiter);
      return s.how === 'first' ? t('From "{column}", the first {count} characters kept', { column, count: s.count })
        : s.how === 'last' ? t('From "{column}", the last {count} characters kept', { column, count: s.count })
          : s.how === 'before' ? t('From "{column}", the text before {delimiter} kept', { column, delimiter })
            : s.how === 'after' ? t('From "{column}", the text after {delimiter} kept', { column, delimiter })
              : t('From "{column}", text kept', { column });
    }
    case 'appendQuery': return t('Appended {source}', { source: sourceWords(s.with) });
    case 'mergeQueries': {
      const values = { source: sourceWords(s.with), on: s.on, withOn: s.withOn };
      return s.how === 'inner' ? t('Merged with {source} on "{on}" = "{withOn}" (inner)', values)
        : s.how === 'leftAnti' ? t('Merged with {source} on "{on}" = "{withOn}" (left anti)', values)
          : s.how === 'full' ? t('Merged with {source} on "{on}" = "{withOn}" (full outer)', values)
            : t('Merged with {source} on "{on}" = "{withOn}"', values);
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
        ...(got.queries || []).filter((q) => q.id !== query.id).map((q) => ({ label: t('Query: {name}', { name: q.name }), name: q.name, source: { kind: 'query', id: q.id, name: q.name } })),
        ...(got.tables || []).filter((t) => t.name !== own?.table && !(query.source?.kind === 'table' && query.source.table === t.name)).map((tb) => ({ label: t('Table: {name} ({sheet})', { name: tb.name, sheet: tb.sheet }), name: `${tb.name} (${tb.sheet})`, source: { kind: 'table', table: tb.name } })),
      ]);
    }).catch(() => {});
    return () => { live = false; };
  }, [listSources, query.id, query.source]);
  const sourceOptions = sources.map((s, i) => [String(i), s.label]);
  const [name, setName] = useState(query.name || t('Query'));
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
  const need = () => { if (!column) { setError(t('Click a column\'s heading first.')); return false; } return true; };
  const askFor = (title, fields, run) => setAsk({ title, fields, values: Object.fromEntries(fields.map((f) => [f.key, f.value ?? ''])), run });

  const actions = [
    [t('Remove Column'), () => need() && add({ kind: 'removeColumns', columns: [column] })],
    [t('Remove Other Columns'), () => need() && add({ kind: 'keepColumns', columns: [column] })],
    [t('Rename'), () => need() && askFor(t('Rename "{column}"', { column }), [{ key: 'to', label: t('New name'), value: column }], (v) => add({ kind: 'renameColumn', from: column, to: v.to }))],
    [t('Data Type'), (e) => need() && menu.open(e, TYPES.map(([type, label]) => ({ label, run: () => add({ kind: 'changeType', column, type }) })))],
    [t('Filter'), () => need() && askFor(t('Keep the rows where "{column}"…', { column }), [{ key: 'op', label: t('Condition'), options: FILTERS, value: 'equals' }, { key: 'value', label: t('Value') }], (v) => add({ kind: 'filterRows', column, op: v.op, value: v.value !== '' && !Number.isNaN(Number(v.value)) ? Number(v.value) : v.value }))],
    [t('Sort A → Z'), () => need() && add({ kind: 'sort', column, descending: false })],
    [t('Sort Z → A'), () => need() && add({ kind: 'sort', column, descending: true })],
    [t('Replace Values'), () => need() && askFor(t('Replace in "{column}"', { column }), [{ key: 'find', label: t('Value to find') }, { key: 'replace', label: t('Replace with') }], (v) => add({ kind: 'replaceValues', column, find: v.find, replace: v.replace }))],
    [t('Split Column'), (e) => need() && menu.open(e, [[t('At each comma'), 'comma'], [t('At each tab'), 'tab'], [t('At each space'), 'space'], [t('At each semicolon'), 'semicolon']].map(([label, delimiter]) => ({ label, run: () => add({ kind: 'splitColumn', column, delimiter }) })).concat([{ label: t('At another character…'), run: () => askFor(t('Split "{column}" at', { column }), [{ key: 'delimiter', label: t('Character') }], (v) => add({ kind: 'splitColumn', column, delimiter: v.delimiter })) }]))],
    [t('Format'), (e) => need() && menu.open(e, [[t('Trim'), 'trim'], [t('UPPERCASE'), 'upper'], [t('lowercase'), 'lower'], [t('Capitalise Each Word'), 'proper']].map(([label, how]) => ({ label, run: () => add({ kind: 'transformText', column, how }) })))],
    // Transform → Extract, as Power Query's: part of each cell's text kept.
    [t('Extract'), (e) => need() && menu.open(e, [
      { label: t('First characters…'), run: () => askFor(t('The first characters of "{column}"', { column }), [{ key: 'count', label: t('How many'), value: '3' }], (v) => add({ kind: 'extractText', column, how: 'first', count: Number(v.count) || 0 })) },
      { label: t('Last characters…'), run: () => askFor(t('The last characters of "{column}"', { column }), [{ key: 'count', label: t('How many'), value: '3' }], (v) => add({ kind: 'extractText', column, how: 'last', count: Number(v.count) || 0 })) },
      { label: t('Text before delimiter…'), run: () => askFor(t('The text of "{column}" before', { column }), [{ key: 'delimiter', label: t('Delimiter'), value: '-' }], (v) => add({ kind: 'extractText', column, how: 'before', delimiter: v.delimiter })) },
      { label: t('Text after delimiter…'), run: () => askFor(t('The text of "{column}" after', { column }), [{ key: 'delimiter', label: t('Delimiter'), value: '-' }], (v) => add({ kind: 'extractText', column, how: 'after', delimiter: v.delimiter })) },
    ])],
    // Transform → Fill: a blank cell takes the value above it, or below.
    [t('Fill'), (e) => need() && menu.open(e, [
      { label: t('Down'), run: () => add({ kind: 'fillDown', columns: [column] }) },
      { label: t('Up'), run: () => add({ kind: 'fillUp', columns: [column] }) },
    ])],
    // Transform → Pivot Column: this column's values become columns, filled from another, added up.
    [t('Pivot Column'), () => need() && askFor(t('Pivot "{column}"', { column }), [
      { key: 'values', label: t('Values from'), options: (table?.columns || []).filter((c) => c !== column).map((c) => [c, c]), value: [...(table?.columns || [])].reverse().find((c) => c !== column) || '' },
      { key: 'fn', label: t('Add up as'), options: [['sum', t('Sum')], ['count', t('Count')], ['average', t('Average')], ['min', t('Minimum')], ['max', t('Maximum')], ['none', t('Do not add up')]], value: 'sum' },
    ], (v) => add({ kind: 'pivotColumn', column, values: v.values, fn: v.fn }))],
    // Add Column → Conditional Column: if a column meets a condition then one value, otherwise another.
    [t('Conditional Column'), () => askFor(t('Add a conditional column'), [
      { key: 'name', label: t('New column name'), value: t('Custom') },
      { key: 'column', label: t('If column'), options: (table?.columns || []).map((c) => [c, c]), value: column || table?.columns?.[0] || '' },
      { key: 'op', label: t('Is'), options: FILTERS, value: 'equals' },
      { key: 'value', label: t('Value') },
      { key: 'output', label: t('Then') },
      { key: 'otherwise', label: t('Otherwise') },
    ], (v) => {
      const typedOf = (x) => (x !== '' && !Number.isNaN(Number(x)) ? Number(x) : x);
      add({ kind: 'conditionalColumn', name: v.name || t('Custom'), rules: [{ column: v.column, op: v.op, value: typedOf(v.value), output: v.output === '' ? null : typedOf(v.output) }], otherwise: v.otherwise === '' ? null : typedOf(v.otherwise) });
    })],
    // Transform → Unpivot Other Columns: this column kept; every other column's cells become rows.
    [t('Unpivot Other Columns'), () => need() && add({ kind: 'unpivotOthers', columns: [column] })],
    // Transform → Merge Columns: this column and another joined, with a separator.
    [t('Merge Columns'), () => need() && askFor(t('Merge "{column}" with', { column }), [
      { key: 'other', label: t('Column'), options: (table?.columns || []).filter((c) => c !== column).map((c) => [c, c]), value: (table?.columns || []).find((c) => c !== column) || '' },
      { key: 'separator', label: t('Separator'), options: [[' ', t('Space')], [', ', t('Comma')], ['-', t('Dash')], ['', t('None')]], value: ' ' },
      { key: 'name', label: t('New column name'), value: t('Merged') },
    ], (v) => add({ kind: 'mergeColumns', columns: [column, v.other], separator: v.separator, name: v.name || t('Merged') }))],
    [t('Group By'), () => need() && askFor(t('Group by "{column}"', { column }), [{ key: 'fn', label: t('Operation'), options: AGGREGATES, value: 'count' }, { key: 'of', label: t('Of column'), options: (table?.columns || []).map((c) => [c, c]), value: (table?.columns || []).find((c) => c !== column) || column }, { key: 'as', label: t('New column name'), value: t('Count') }], (v) => add({ kind: 'groupBy', columns: [column], aggregations: [{ fn: v.fn, column: v.fn === 'count' ? null : v.of, name: v.as || undefined }] }))],
    [t('Append Queries'), () => (sources.length
      ? askFor(t('Append the rows of'), [{ key: 'with', label: t('Table or query'), options: sourceOptions, value: '0' }], (v) => add({ kind: 'appendQuery', with: sources[Number(v.with)].source }))
      : setError(t('There is no other table or query in this workbook to append.')))],
    [t('Merge Queries'), () => {
      if (!need()) return;
      if (!sources.length) { setError(t('There is no other table or query in this workbook to merge with.')); return; }
      askFor(t('Merge on "{column}" with', { column }), [{ key: 'with', label: t('Table or query'), options: sourceOptions, value: '0' }, { key: 'how', label: t('Join kind'), options: JOIN_KINDS.map(([kind, label]) => [kind, t(label)]), value: 'left' }], async (v) => {
        const other = sources[Number(v.with)].source;
        try {
          const cols = (await preview({ source: other, steps: [] }))?.columns || [];
          askFor(t('Match "{column}" with a column of {source}', { column, source: sources[Number(v.with)].name }), [{ key: 'withOn', label: t('Its column'), options: cols.map((c) => [c, c]), value: cols.find((c) => c.toLowerCase() === String(column).toLowerCase()) || cols[0] }], (w) => add({ kind: 'mergeQueries', with: other, on: column, withOn: w.withOn, how: v.how }));
        } catch (err) { setError(cleanError(err)); }
      });
    }],
    [t('Keep Top Rows'), () => askFor(t('Keep the first rows'), [{ key: 'count', label: t('Number of rows'), value: '10' }], (v) => add({ kind: 'keepTopRows', count: Number(v.count) || 0 }))],
    [t('Remove Duplicates'), () => add({ kind: 'removeDuplicates', ...(column ? { columns: [column] } : {}) })],
    [t('Remove Blank Rows'), () => add({ kind: 'removeBlankRows' })],
    [t('Use First Row as Headers'), () => add({ kind: 'promoteHeaders' })],
    [t('Index Column'), () => add({ kind: 'addIndex', name: t('Index'), start: 1 })],
  ];

  const load = async () => {
    setBusy(true);
    try { await onLoad({ id: query.id, name: name.trim() || t('Query'), source: query.source, sourceText: query.sourceText, steps }); } catch (err) { setError(cleanError(err)); setBusy(false); }
  };

  return (
    <>
      {menu.node}
      <Dialog
        title={t('Power Query Editor — {source}', { source: query.sourceText || '' })}
        width={980}
        onClose={onClose}
        actions={(
          <>
            <Input className="pq-name" value={name} onChange={(e) => setName(e.target.value)} style={{ width: 220 }} title={t('The query\'s name, and its sheet\'s')} />
            <Button label={t('Cancel')} onClick={onClose} />
            <Button primary className="pq-load" label={busy ? t('Loading…') : t('Close & Load')} disabled={busy || Boolean(error && !table)} onClick={load} />
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
              <Button primary className="pq-ask-ok" label={t('OK')} onClick={() => { ask.run(ask.values); setAsk(null); }} />
              <Button label={t('Cancel')} onClick={() => setAsk(null)} />
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
                      {table.columns.map((c) => <th key={c} className={c === column ? 'sel' : ''} data-col={c} onClick={() => setColumn(c)} title={t('Pick this column for the next step')}>{c}</th>)}
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
              ) : !error ? <div className="pq-note">{t('Reading the source…')}</div> : null}
            </div>
            <div className="pq-steps">
              <div className="pq-steps-head">{t('Applied steps')}</div>
              <button type="button" className={`pq-step${upTo === 0 ? ' at' : ''}`} onClick={() => setUpTo(0)}>{t('Source')}</button>
              {steps.map(stepLabel).map((label, i) => (
                <div key={i} className={`pq-step-row${upTo === i + 1 || (upTo === null && i === steps.length - 1) ? ' at' : ''}`}>
                  <button type="button" className="pq-step" onClick={() => setUpTo(i + 1 === steps.length ? null : i + 1)}>{label}</button>
                  <button type="button" className="pq-step-x" title={t('Take this step away')} onClick={() => { setSteps((s) => s.filter((_, k) => k !== i)); setUpTo(null); }}>×</button>
                </div>
              ))}
              <div className="pq-count">{table ? (table.total > table.rows.length ? t('{rows}, the first {shown} shown', { rows: tn(table.total, '{count} row', '{count} rows'), shown: table.rows.length }) : tn(table.total, '{count} row', '{count} rows')) : ''}</div>
            </div>
          </div>
        </div>
      </Dialog>
    </>
  );
}

/** Data → Queries & Connections: the workbook's queries, each refreshed, edited or deleted from here. */
export function QueriesPane({ queries, onRefresh, onEdit, onDelete, onRefreshAll, onGoTo }) {
  if (!queries?.length) return <div className="pq-note"><style>{QUERIES_CSS}</style>{t('No queries in this workbook yet. Data → From Table/Range or From Text/CSV makes one.')}</div>;
  return (
    <div className="pq-pane">
      <style>{QUERIES_CSS}</style>
      <Button icon="refresh" className="pq-refresh-all" label={t('Refresh All')} onClick={onRefreshAll} />
      {queries.map((q) => (
        <div key={q.id} className="pq-item" data-query={q.id}>
          <button type="button" className="pq-item-name" onClick={() => onGoTo(q)} title={t('Go to its sheet')}>{q.name}</button>
          <div className="pq-item-meta">{q.loaded != null ? tn(q.loaded, '{count} row loaded', '{count} rows loaded') : t('Not loaded')} · {q.sourceText}</div>
          <div className="pq-item-tools">
            <Button icon="refresh" className="pq-refresh" title={t('Refresh — run it again on its source as it is now')} onClick={() => onRefresh(q)} />
            <Button icon="textbox" className="pq-edit" title={t('Edit — open it in the Power Query Editor')} onClick={() => onEdit(q)} />
            <Button icon="trash" className="pq-delete" title={t('Delete the query; its sheet stays')} onClick={() => onDelete(q)} />
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
