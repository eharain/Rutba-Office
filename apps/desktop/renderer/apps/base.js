// Databases (2.9, D-2026-10-20-17): tables, their design, relationships
// and SQL over one SQLite file. Kept plain on purpose: a list of tables on
// the left and, for the one chosen, its records or its design; the
// relationships and a SQL query each a view of their own. The file is
// held by the main process (main/base.js); this window asks it for a page
// of records at a time, so a table of any size opens at once.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Input, Select, Content, Empty, Dialog, Spinner, useToast, t, tn } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, useFileDrop } from '../shell.js';
import { BASE_CSS } from './base/css.js';

export const TYPE_NAMES = () => ({
  autonumber: t('AutoNumber'),
  text: t('Short text'),
  longtext: t('Long text'),
  number: t('Number'),
  integer: t('Whole number'),
  currency: t('Currency'),
  date: t('Date'),
  datetime: t('Date and time'),
  yesno: t('Yes/No'),
});
const TYPES = ['text', 'longtext', 'number', 'integer', 'currency', 'date', 'datetime', 'yesno'];
const PAGE = 100;
const ROW_H = 30;
const OPENS = /\.(rbase|sqlite3?|db)$/i;
const IMPORTS = ['csv', 'tsv', 'txt', 'dbf', 'mdb', 'accdb', 'wk1', 'wks', 'wk3', 'wk4', 'slk', 'dif'];

/** A value as a cell shows it. */
function shown(field, value) {
  if (value === null || value === undefined) return '';
  if (field?.type === 'currency' && typeof value === 'number') return value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 4 });
  if (field?.type === 'datetime') return String(value).replace('T', ' ');
  return String(value);
}

/* ── the datasheet: a page of records at a time ───────────────────── */

function Datasheet({ shell, db, table, onChanged, toast }) {
  const [fields, setFields] = useState([]);
  const [count, setCount] = useState(0);
  const [order, setOrder] = useState(null);
  const [version, setVersion] = useState(0);
  const pages = useRef(new Map());
  const [, redraw] = useState(0);
  const [scrollTop, setScrollTop] = useState(0);
  const [height, setHeight] = useState(400);
  const [picked, setPicked] = useState(() => new Set());
  const [editing, setEditing] = useState(null); // { rowid, field, value }
  const [fresh, setFresh] = useState({});
  const box = useRef(null);

  useEffect(() => {
    let live = true;
    shell.base.fields({ id: db.id, table }).then((f) => { if (live) setFields(f); }).catch((e) => toast(e.message, { tone: 'bad' }));
    pages.current = new Map();
    setPicked(new Set());
    setEditing(null);
    return () => { live = false; };
  }, [shell, db.id, table, version, toast]);

  const load = useCallback(async (page) => {
    if (pages.current.has(page)) return;
    pages.current.set(page, null);
    try {
      const r = await shell.base.rows({ id: db.id, table, offset: page * PAGE, limit: PAGE, order });
      pages.current.set(page, r.rows);
      setCount(r.count);
      redraw((n) => n + 1);
    } catch (e) {
      pages.current.delete(page);
      toast(e.message, { tone: 'bad' });
    }
  }, [shell, db.id, table, order, toast]);

  useEffect(() => { pages.current = new Map(); load(0); }, [load, version]);

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

  const refresh = () => { setVersion((v) => v + 1); onChanged?.(); };

  const save = async (rowid, field, value) => {
    setEditing(null);
    try {
      await shell.base.update({ id: db.id, table, rowid, values: { [field.name]: value } });
      pages.current = new Map();
      load(Math.floor(first / PAGE));
      redraw((n) => n + 1);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 6000 });
    }
  };

  const addRecord = async () => {
    try {
      await shell.base.insert({ id: db.id, table, values: fresh });
      setFresh({});
      refresh();
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 6000 });
    }
  };

  const removePicked = async () => {
    if (!picked.size) return;
    try {
      await shell.base.remove({ id: db.id, table, rowids: [...picked] });
      setPicked(new Set());
      refresh();
      toast(tn(picked.size, '{count} record deleted', '{count} records deleted'));
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 6000 });
    }
  };

  const editable = fields.filter((f) => f.type !== 'autonumber');
  const cols = fields.map((f) => (f.type === 'longtext' ? 280 : f.type === 'yesno' ? 90 : f.type === 'autonumber' ? 80 : 160));
  const template = `36px ${cols.map((w) => `${w}px`).join(' ')}`;
  const wide = 36 + cols.reduce((a, b) => a + b, 0);
  const sortBy = (name) => { setOrder((o) => (o === name ? `-${name}` : o === `-${name}` ? null : name)); pages.current = new Map(); };

  const cell = (row, f) => {
    if (f.type === 'yesno') {
      return <input type="checkbox" checked={Boolean(row[f.name])} aria-label={f.name} onChange={(e) => save(row._rowid, f, e.target.checked)} />;
    }
    if (editing && editing.rowid === row._rowid && editing.field === f.name) {
      return (
        <input
          className="bs-edit"
          aria-label={f.name}
          autoFocus
          value={editing.value}
          onChange={(e) => setEditing({ ...editing, value: e.target.value })}
          onKeyDown={(e) => { if (e.key === 'Enter') save(row._rowid, f, editing.value); if (e.key === 'Escape') setEditing(null); }}
          onBlur={() => save(row._rowid, f, editing.value)}
        />
      );
    }
    return <span className="bs-text">{shown(f, row[f.name])}</span>;
  };

  return (
    <div className="bs-sheet">
      <div className="bs-tools">
        <Button icon="trash" label={t('Delete')} disabled={!picked.size} className="bs-delete" onClick={removePicked} />
        <span className="bs-count">{tn(count, '{count} record', '{count} records')}</span>
      </div>
      <div className="bs-grid" ref={box} onScroll={(e) => setScrollTop(e.currentTarget.scrollTop)} role="grid" aria-rowcount={count + 1}>
        <div className="bs-row bs-head" style={{ gridTemplateColumns: template }} role="row">
          <span />
          {fields.map((f) => (
            <button key={f.name} type="button" className="bs-th" role="columnheader" data-field={f.name} aria-sort={order === f.name ? 'ascending' : order === `-${f.name}` ? 'descending' : 'none'} onClick={() => sortBy(f.name)}>
              <span>{f.name}</span>
              {order === f.name ? <Icon name="chevronUp" size={12} /> : order === `-${f.name}` ? <Icon name="chevronDown" size={12} /> : null}
            </button>
          ))}
        </div>
        <div className="bs-body" style={{ height: count * ROW_H, width: wide }}>
          {Array.from({ length: Math.max(0, last - first) }, (_, k) => {
            const i = first + k;
            const row = rowAt(i);
            return (
              <div key={i} className={`bs-row${row && picked.has(row._rowid) ? ' picked' : ''}`} role="row" data-row={i} data-rowid={row?._rowid} style={{ top: i * ROW_H, gridTemplateColumns: template }}>
                <span className="bs-pick">
                  {row ? <input type="checkbox" aria-label={t('Pick')} checked={picked.has(row._rowid)} onChange={(e) => setPicked((p) => { const n = new Set(p); if (e.target.checked) n.add(row._rowid); else n.delete(row._rowid); return n; })} /> : null}
                </span>
                {fields.map((f) => (
                  <span
                    key={f.name}
                    className={`bs-td t-${f.type}`}
                    role="gridcell"
                    data-field={f.name}
                    onDoubleClick={() => row && f.type !== 'autonumber' && f.type !== 'yesno' && setEditing({ rowid: row._rowid, field: f.name, value: row[f.name] ?? '' })}
                  >
                    {row ? cell(row, f) : null}
                  </span>
                ))}
              </div>
            );
          })}
        </div>
        <div className="bs-row bs-new" style={{ gridTemplateColumns: template }} role="row">
          <span className="bs-pick"><Icon name="plus" size={14} /></span>
          {fields.map((f) => (
            <span key={f.name} className="bs-td">
              {f.type === 'autonumber' ? <span className="bs-auto">{t('(new)')}</span>
                : f.type === 'yesno' ? <input type="checkbox" aria-label={f.name} checked={Boolean(fresh[f.name])} onChange={(e) => setFresh({ ...fresh, [f.name]: e.target.checked })} />
                  : <input className="bs-new-input" data-field={f.name} aria-label={f.name} placeholder={TYPE_NAMES()[f.type]} value={fresh[f.name] ?? ''} onChange={(e) => setFresh({ ...fresh, [f.name]: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter') addRecord(); }} />}
            </span>
          ))}
        </div>
      </div>
      {editable.length ? null : <p className="bs-hint">{t('This table has no fields to fill in yet. Add some in its design.')}</p>}
    </div>
  );
}

/* ── the design: a table's fields ─────────────────────────────────── */

function Design({ shell, db, table, onChanged, onRemoved, toast }) {
  const [fields, setFields] = useState([]);
  const [version, setVersion] = useState(0);
  const [adding, setAdding] = useState({ name: '', type: 'text' });
  useEffect(() => {
    shell.base.fields({ id: db.id, table }).then(setFields).catch((e) => toast(e.message, { tone: 'bad' }));
  }, [shell, db.id, table, version, toast]);
  const names = TYPE_NAMES();
  const run = async (fn, done = null) => {
    try {
      const s = await fn();
      if (s) onChanged(s);
      setVersion((v) => v + 1);
      done?.();
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
      setVersion((v) => v + 1);
    }
  };
  const change = (f, change) => run(() => shell.base.changeField({ id: db.id, table, name: f.name, change }));
  return (
    <div className="bs-design">
      <table className="bs-fields">
        <thead>
          <tr><th>{t('Field')}</th><th>{t('Type')}</th><th>{t('Size')}</th><th>{t('Required')}</th><th>{t('Unique')}</th><th>{t('Default')}</th><th /></tr>
        </thead>
        <tbody>
          {fields.map((f) => (
            <tr key={f.name} data-field={f.name}>
              <td><Input className="rw-input bs-f-name" aria-label={t('Field name')} defaultValue={f.name} key={`n-${f.name}-${version}`} onBlur={(e) => { const v = e.target.value.trim(); if (v && v !== f.name) change(f, { name: v }); }} onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }} /></td>
              <td>
                {f.type === 'autonumber' ? <span className="bs-fixed">{names.autonumber}</span> : (
                  <Select className="rw-select bs-f-type" aria-label={t('Type')} value={f.type} onChange={(e) => change(f, { type: e.target.value })}>
                    {TYPES.map((k) => <option key={k} value={k}>{names[k]}</option>)}
                  </Select>
                )}
              </td>
              <td>{f.type === 'text' ? <Input className="rw-input bs-f-size" aria-label={t('Size')} type="number" defaultValue={f.size ?? 255} key={`s-${f.name}-${version}`} onBlur={(e) => { const v = Number(e.target.value); if (v > 0 && v !== (f.size ?? 255)) change(f, { size: v }); }} /> : null}</td>
              <td>{f.type === 'autonumber' || f.type === 'yesno' ? null : <input type="checkbox" aria-label={t('Required')} className="bs-f-required" checked={f.required} onChange={(e) => change(f, { required: e.target.checked })} />}</td>
              <td>{f.type === 'autonumber' || f.type === 'yesno' ? null : <input type="checkbox" aria-label={t('Unique')} className="bs-f-unique" checked={f.unique} onChange={(e) => change(f, { unique: e.target.checked })} />}</td>
              <td>{f.type === 'autonumber' ? null : <Input className="rw-input bs-f-default" aria-label={t('Default')} defaultValue={f.default ?? ''} key={`d-${f.name}-${version}`} onBlur={(e) => { const v = e.target.value; if (v !== String(f.default ?? '')) change(f, { default: v === '' ? undefined : v }); }} />}</td>
              <td>{f.type === 'autonumber' ? null : <Button icon="close" label="" title={t('Remove the field and its data')} className="bs-f-remove" onClick={() => run(() => shell.base.removeField({ id: db.id, table, name: f.name }))} />}</td>
            </tr>
          ))}
          <tr className="bs-f-new">
            <td><Input className="rw-input bs-add-name" aria-label={t('New field')} placeholder={t('New field')} value={adding.name} onChange={(e) => setAdding({ ...adding, name: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter' && adding.name.trim()) run(() => shell.base.addField({ id: db.id, table, field: { ...adding, name: adding.name.trim() } }), () => setAdding({ name: '', type: adding.type })); }} /></td>
            <td>
              <Select className="rw-select bs-add-type" aria-label={t('Type')} value={adding.type} onChange={(e) => setAdding({ ...adding, type: e.target.value })}>
                {TYPES.map((k) => <option key={k} value={k}>{names[k]}</option>)}
              </Select>
            </td>
            <td colSpan={5}>
              <Button icon="plus" label={t('Add Field')} className="bs-add-field" disabled={!adding.name.trim()} onClick={() => run(() => shell.base.addField({ id: db.id, table, field: { ...adding, name: adding.name.trim() } }), () => setAdding({ name: '', type: adding.type }))} />
            </td>
          </tr>
        </tbody>
      </table>
      <div className="bs-design-foot">
        <Button icon="trash" label={t('Delete Table')} className="bs-drop" onClick={() => run(() => shell.base.dropTable({ id: db.id, table }), onRemoved)} />
        <span className="bs-hint">{t('Each change is applied to the records at once, and kept in the database\'s history.')}</span>
      </div>
    </div>
  );
}

/* ── relationships ────────────────────────────────────────────────── */

function Relationships({ shell, db, onChanged, toast }) {
  const tables = db.tables.map((x) => x.name);
  const [fieldsOf, setFieldsOf] = useState({});
  const [form, setForm] = useState({ fromTable: '', fromField: '', toTable: '', toField: '', cascadeDelete: false });
  useEffect(() => {
    let live = true;
    Promise.all(tables.map((n) => shell.base.fields({ id: db.id, table: n }).then((f) => [n, f]))).then((all) => { if (live) setFieldsOf(Object.fromEntries(all)); }).catch(() => {});
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shell, db]);
  const keysOf = (n) => (fieldsOf[n] || []).filter((f) => f.type === 'autonumber' || f.unique);
  const relate = async () => {
    try {
      const s = await shell.base.addRelationship({ id: db.id, from: { table: form.fromTable, field: form.fromField }, to: { table: form.toTable, field: form.toField }, cascadeDelete: form.cascadeDelete });
      onChanged(s);
      setForm({ fromTable: '', fromField: '', toTable: '', toField: '', cascadeDelete: false });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  const ready = form.fromTable && form.fromField && form.toTable && form.toField;
  return (
    <div className="bs-relations">
      <h2>{t('Relationships')}</h2>
      <p className="bs-hint">{t('A relationship lets many records in one table point at one record in another, and keeps them from pointing at nothing.')}</p>
      <ul className="bs-rel-list">
        {db.relationships.map((r) => (
          <li key={r.id} className="bs-rel" data-relation={r.id}>
            <span><strong>{r.from.table}</strong>.{r.from.field}</span>
            <Icon name="chevronRight" size={14} />
            <span><strong>{r.to.table}</strong>.{r.to.field}</span>
            {r.cascadeDelete ? <span className="bs-chip">{t('Deleting one deletes its records here')}</span> : null}
            <Button icon="close" label="" title={t('Remove the relationship')} className="bs-rel-remove" onClick={async () => { try { onChanged(await shell.base.removeRelationship({ id: db.id, relation: r.id })); } catch (e) { toast(e.message, { tone: 'bad' }); } }} />
          </li>
        ))}
        {db.relationships.length ? null : <li className="bs-hint">{t('No relationships yet.')}</li>}
      </ul>
      <div className="bs-rel-form">
        <label><span>{t('Records in')}</span>
          <Select className="rw-select bs-rel-from-table" aria-label={t('Table')} value={form.fromTable} onChange={(e) => setForm({ ...form, fromTable: e.target.value, fromField: '' })}>
            <option value="">{t('Table')}</option>
            {tables.map((n) => <option key={n} value={n}>{n}</option>)}
          </Select>
          <Select className="rw-select bs-rel-from-field" aria-label={t('Field')} value={form.fromField} onChange={(e) => setForm({ ...form, fromField: e.target.value })}>
            <option value="">{t('Field')}</option>
            {(fieldsOf[form.fromTable] || []).filter((f) => f.type !== 'autonumber').map((f) => <option key={f.name} value={f.name}>{f.name}</option>)}
          </Select>
        </label>
        <label><span>{t('point at one in')}</span>
          <Select className="rw-select bs-rel-to-table" aria-label={t('Table')} value={form.toTable} onChange={(e) => setForm({ ...form, toTable: e.target.value, toField: keysOf(e.target.value)[0]?.name || '' })}>
            <option value="">{t('Table')}</option>
            {tables.map((n) => <option key={n} value={n}>{n}</option>)}
          </Select>
          <Select className="rw-select bs-rel-to-field" aria-label={t('Field')} value={form.toField} onChange={(e) => setForm({ ...form, toField: e.target.value })}>
            <option value="">{t('Field')}</option>
            {keysOf(form.toTable).map((f) => <option key={f.name} value={f.name}>{f.name}</option>)}
          </Select>
        </label>
        <label className="bs-check"><input type="checkbox" checked={form.cascadeDelete} onChange={(e) => setForm({ ...form, cascadeDelete: e.target.checked })} /><span>{t('Deleting one deletes its records here')}</span></label>
        <Button primary icon="link" label={t('Relate')} className="bs-relate" disabled={!ready} onClick={relate} />
      </div>
    </div>
  );
}

/* ── SQL ──────────────────────────────────────────────────────────── */

function Sql({ shell, db }) {
  const first = db.tables[0]?.name;
  const [text, setText] = useState(first ? `SELECT * FROM "${first}"` : 'SELECT 1');
  const [out, setOut] = useState(null);
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const runIt = async () => {
    setBusy(true);
    setError('');
    try { setOut(await shell.base.query({ id: db.id, sql: text })); } catch (e) { setError(e.message); setOut(null); } finally { setBusy(false); }
  };
  return (
    <div className="bs-sql">
      <textarea className="rw-input bs-sql-text" spellCheck={false} value={text} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); runIt(); } }} aria-label={t('Query')} />
      <div className="bs-tools">
        <Button primary icon="play" label={t('Run')} className="bs-sql-run" disabled={busy} title={t('Run the query (Ctrl+Enter). It only reads: nothing in the database changes.')} onClick={runIt} />
        {out ? <span className="bs-count">{out.count > out.rows.length ? t('{shown} of {count} rows', { shown: out.rows.length, count: out.count }) : tn(out.count, '{count} row', '{count} rows')}</span> : null}
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
    </div>
  );
}

/* ── the window ───────────────────────────────────────────────────── */

export default function Databases({ app, shell, boot }) {
  const toast = useToast();
  const [db, setDb] = useState(null);
  const [table, setTable] = useState(null);
  const [view, setView] = useState('data'); // data | design | relations | sql
  const [busy, setBusy] = useState(false);
  const [naming, setNaming] = useState(null);

  const show = useCallback((summary) => {
    setDb(summary);
    setTable((cur) => (summary.tables.some((x) => x.name === cur) ? cur : summary.tables[0]?.name || null));
  }, []);

  const openFile = useCallback(async (file) => {
    setBusy(true);
    try {
      const s = await shell.base.open({ file });
      show(s);
      setView('data');
      shell.app.addRecent({ path: file, app: 'base' }).catch(() => {});
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    } finally {
      setBusy(false);
    }
  }, [shell, show, toast]);

  const newDatabase = useCallback(async () => {
    const file = await shell.dialog.save({ title: t('New Database'), defaultPath: `${t('Database')}.rbase`, filters: [{ name: t('Rutba Database (.rbase)'), extensions: ['rbase'] }] });
    if (!file) return;
    try {
      const s = await shell.base.create({ file: /\.rbase$/i.test(file) ? file : `${file}.rbase` });
      show(s);
      setView('data');
      shell.app.addRecent({ path: s.file, app: 'base' }).catch(() => {});
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  }, [shell, show, toast]);

  const openDialog = useCallback(async () => {
    const file = await pickOpen(shell, 'base');
    if (file) openFile(file);
  }, [shell, openFile]);

  const importDialog = useCallback(async () => {
    if (!db) return;
    const files = await shell.dialog.open({ title: t('Import'), filters: [{ name: t('Tables and old files'), extensions: IMPORTS }] });
    if (!files?.[0]) return;
    setBusy(true);
    try {
      const r = await shell.base.importFile({ id: db.id, file: files[0] });
      show(r.summary);
      setTable(r.tables[0]?.table || null);
      setView('data');
      const made = r.tables.map((x) => t('{table} ({count})', { table: x.table, count: x.count })).join(', ');
      toast(t('Imported: {tables}', { tables: made }), { tone: 'good', ms: 7000 });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(false);
    }
  }, [shell, db, show, toast]);

  const saveAs = useCallback(async () => {
    if (!db) return;
    const file = await shell.dialog.save({ title: t('Save As'), defaultPath: `${db.name}.rbase`, filters: [{ name: t('Rutba Database (.rbase)'), extensions: ['rbase'] }] });
    if (!file) return;
    try { show(await shell.base.saveAs({ id: db.id, file })); } catch (e) { toast(e.message, { tone: 'bad', ms: 7000 }); }
  }, [shell, db, show, toast]);

  const appMenuOpen = useRef(null);
  appMenuOpen.current = openDialog;
  const appMenu = useAppMenu({
    shell,
    appKey: 'base',
    onNew: () => shell.win.create({ app: 'base' }),
    onOpen: () => appMenuOpen.current?.(),
    extra: [{ label: t('Save As'), icon: 'save', run: () => saveAs() }],
  });

  useEffect(() => { if (boot?.file) openFile(boot.file); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  useFileDrop(useCallback((files) => { const f = files.find((x) => OPENS.test(x)); if (f) openFile(f); }, [openFile]));
  useEffect(() => () => { if (db) shell.base.close({ id: db.id }).catch(() => {}); }, [db?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const refresh = useCallback(async () => { if (db) show(await shell.base.summary({ id: db.id })); }, [shell, db, show]);

  const makeTable = async () => {
    const name = String(naming || '').trim();
    if (!name) return;
    try {
      const s = await shell.base.createTable({ id: db.id, name, fields: [] });
      show(s);
      setTable(s.tables.find((x) => x.name === name)?.name || name);
      setView('design');
      setNaming(null);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };

  const title = db ? db.name : t('Databases');
  if (!db) {
    return (
      <AppFrame app={app} shell={shell} title={title} menu={appMenu}>
        <Content>
          <div className="bs-start">
            <Empty
              icon="base"
              title={t('Databases')}
              action={(
                <span className="bs-start-actions">
                  <Button primary icon="plus" label={t('New Database')} className="bs-new-db" onClick={newDatabase} />
                  <Button icon="open" label={t('Open')} className="bs-open-db" onClick={openDialog} />
                  {busy ? <Spinner /> : null}
                </span>
              )}
            >
              {t('Tables of records you keep, related to each other, in one file on this computer.')}
            </Empty>
          </div>
        </Content>
        <style>{BASE_CSS}</style>
      </AppFrame>
    );
  }

  const tab = (key, label) => <button type="button" role="tab" aria-selected={view === key} className={`bs-tab${view === key ? ' on' : ''}`} data-view={key} onClick={() => setView(key)}>{label}</button>;
  return (
    <AppFrame app={app} shell={shell} title={title} menu={appMenu}>
      <div className="bs-app">
        <nav className="bs-side" aria-label={t('Tables')}>
          <div className="bs-side-head">
            <h2>{t('Tables')}</h2>
            <Button icon="plus" label="" title={t('New Table')} className="bs-new-table" onClick={() => setNaming('')} />
          </div>
          <ul className="bs-tables">
            {db.tables.map((x) => (
              <li key={x.name}>
                <button type="button" className={`bs-table${table === x.name && (view === 'data' || view === 'design') ? ' on' : ''}`} data-table={x.name} onClick={() => { setTable(x.name); if (view !== 'design') setView('data'); }}>
                  <Icon name="sheets" size={14} />
                  <span>{x.name}</span>
                  <small>{x.count}</small>
                </button>
              </li>
            ))}
            {db.tables.length ? null : <li className="bs-hint">{t('No tables yet: make one, or import one.')}</li>}
          </ul>
          <div className="bs-side-foot">
            <Button icon="import" label={t('Import')} className="bs-import" onClick={importDialog} />
            <button type="button" className={`bs-table${view === 'relations' ? ' on' : ''}`} data-view="relations" onClick={() => setView('relations')}><Icon name="link" size={14} /><span>{t('Relationships')}</span></button>
            <button type="button" className={`bs-table${view === 'sql' ? ' on' : ''}`} data-view="sql" onClick={() => setView('sql')}><Icon name="formula" size={14} /><span>{t('SQL')}</span></button>
          </div>
        </nav>
        <main className="bs-main">
          {(view === 'data' || view === 'design') && table ? (
            <>
              <div className="bs-tabs" role="tablist">
                <h1>{table}</h1>
                {tab('data', t('Records'))}
                {tab('design', t('Design'))}
              </div>
              {view === 'data'
                ? <Datasheet key={`${table}-${db.steps}`} shell={shell} db={db} table={table} onChanged={refresh} toast={toast} />
                : <Design shell={shell} db={db} table={table} onChanged={show} onRemoved={() => { setTable(null); refresh(); }} toast={toast} />}
            </>
          ) : view === 'relations' ? (
            <Relationships shell={shell} db={db} onChanged={show} toast={toast} />
          ) : view === 'sql' ? (
            <Sql shell={shell} db={db} />
          ) : (
            <Empty icon="sheets" title={t('No table chosen')}>{t('Make a table, or import one from a file.')}</Empty>
          )}
          {busy ? <div className="bs-busy"><Spinner /></div> : null}
        </main>
      </div>
      {naming !== null ? (
        <Dialog title={t('New Table')} onClose={() => setNaming(null)} actions={<><Button label={t('Cancel')} onClick={() => setNaming(null)} /><Button primary label={t('Make')} className="bs-make-table" disabled={!String(naming).trim()} onClick={makeTable} /></>}>
          <label className="bs-field"><span>{t('Name')}</span><Input className="rw-input bs-table-name" aria-label={t('Name')} autoFocus value={naming} onChange={(e) => setNaming(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') makeTable(); }} /></label>
          <p className="bs-hint">{t('It starts with an ID that numbers each record; add its fields in its design.')}</p>
        </Dialog>
      ) : null}
      <style>{BASE_CSS}</style>
    </AppFrame>
  );
}
