// Data Files' query designer: tables put in, the fields to show, criteria in
// the words Access uses, the sort and totals, and the answer as a table. The
// design is kept in the file under a name; the SQL it stands for can be seen
// (the SQL view still takes a query written by hand). The tables are joined
// by their relationships, so a person never writes a join.

import React, { useEffect, useMemo, useState } from 'react';
import { Button, Input, Select, Dialog, t, tn } from '@rutba/office-ui';

const TOTALS = () => [['group', t('Group By')], ['sum', t('Sum')], ['avg', t('Average')], ['count', t('Count')], ['min', t('Min')], ['max', t('Max')], ['where', t('Where')]];
const blank = () => ({ table: '', field: '', show: true, sort: '', total: 'group', criteria: ['', ''] });

/** A design from what the designer holds, the empty columns left out. */
function designOf(tables, columns, totals) {
  return {
    tables,
    columns: columns.filter((c) => c.table && c.field).map((c) => ({
      table: c.table,
      field: c.field,
      show: c.show,
      sort: c.sort || null,
      total: totals ? c.total || 'group' : null,
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

  // A saved query opens with its design; a new one with nothing in it.
  useEffect(() => {
    let live = true;
    setAnswer(null);
    setSql('');
    setError('');
    if (!name) { setTables([]); setColumns([blank()]); setTotals(false); return undefined; }
    shell.base.queries({ id: db.id }).then((all) => {
      const found = all.find((x) => x.name === name);
      if (!live || !found) return;
      const d = found.design;
      setTables(d.tables || []);
      setTotals((d.columns || []).some((c) => c.total));
      setColumns([...(d.columns || []).map((c) => ({ ...blank(), ...c, sort: c.sort || '', total: c.total || 'group', criteria: [...(c.criteria || []), '', ''].slice(0, Math.max(2, (c.criteria || []).length)) })), blank()]);
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
    // A field chosen on the last line opens another below it.
    return next[next.length - 1].field ? [...next, blank()] : next;
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
      onSaved(summary, as);
      toast(t('Kept as {name}', { name: as }), { tone: 'good' });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };

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
              <th>{t('Field')}</th><th>{t('Show')}</th><th>{t('Sort')}</th>{totals ? <th>{t('Total')}</th> : null}<th>{t('Criteria')}</th><th>{t('Or')}</th><th />
            </tr>
          </thead>
          <tbody>
            {columns.map((c, i) => (
              <tr key={i} data-column={i}>
                <td>
                  <Select className="rw-select bs-q-field" aria-label={t('Field')} value={c.table && c.field ? `${c.table}\u0001${c.field}` : ''} onChange={(e) => { const [tb, f] = e.target.value.split('\u0001'); setColumn(i, { table: tb || '', field: f || '' }); }}>
                    <option value="">{tables.length ? t('Choose a field') : t('Add a table first')}</option>
                    {tables.map((tb) => (
                      <optgroup key={tb} label={tb}>
                        {(fieldsOf[tb] || []).map((f) => <option key={f.name} value={`${tb}\u0001${f.name}`}>{f.name}</option>)}
                      </optgroup>
                    ))}
                  </Select>
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
                {[0, 1].map((r) => (
                  <td key={r}>
                    <Input className={`rw-input bs-q-criteria bs-q-row-${r}`} aria-label={r ? t('Or') : t('Criteria')} placeholder={r ? '' : t('e.g. > 10, Like "A*"')} value={c.criteria[r] || ''} onChange={(e) => setColumn(i, { criteria: c.criteria.map((x, k) => (k === r ? e.target.value : x)) })} onKeyDown={(e) => { if (e.key === 'Enter') run(); }} />
                  </td>
                ))}
                <td>{c.field ? <Button icon="close" label="" title={t('Take this field out')} className="bs-q-remove" onClick={() => setColumns((cs) => cs.filter((_, k) => k !== i))} /> : null}</td>
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
      </div>
      {error ? <p className="bs-error" role="alert">{error}</p> : null}
      {sql ? <pre className="bs-q-sql-text selectable">{sql}</pre> : null}
      {answer ? (
        <div className="bs-result">
          <table>
            <thead><tr>{answer.columns.map((c) => <th key={c}>{c}</th>)}</tr></thead>
            <tbody>{answer.rows.map((r, i) => <tr key={i}>{answer.columns.map((c) => <td key={c}>{r[c] === null || r[c] === undefined ? '' : String(r[c])}</td>)}</tr>)}</tbody>
          </table>
        </div>
      ) : null}

      {naming !== null ? (
        <Dialog title={t('Save Query')} onClose={() => setNaming(null)} actions={<><Button label={t('Cancel')} onClick={() => setNaming(null)} /><Button primary label={t('Save')} className="bs-q-save-named" disabled={!String(naming).trim()} onClick={() => save(String(naming).trim())} /></>}>
          <label className="bs-field"><span>{t('Name')}</span><Input className="rw-input bs-q-name" aria-label={t('Name')} autoFocus value={naming} onChange={(e) => setNaming(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && String(naming).trim()) save(String(naming).trim()); }} /></label>
        </Dialog>
      ) : null}
    </div>
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
.bs-q-center { text-align: center; }
.bs-q-sql-text { margin: 0 16px 8px; padding: 8px 10px; border-radius: var(--r-2); background: var(--sunken); font: 12.5px/1.5 ui-monospace, Consolas, monospace; white-space: pre-wrap; }
`;
