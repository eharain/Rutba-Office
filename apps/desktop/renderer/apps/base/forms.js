// A Data Files form: one record at a time, laid out as its layout says (the
// form wizard's, kept in the file). Each control shows its field and keeps
// a change when it is left; a lookup is a list of the related table's
// records by their first words; the records elsewhere that point at this
// one are listed below; and the bar at the foot moves between records,
// starts a new one and deletes the one shown.

import React, { useCallback, useEffect, useState } from 'react';
import { Button, Select, t, tn } from '@rutba/office-ui';
import { valueFromControl } from '@rutba/database/form-design';
import { usePrinter } from './documents.js';
import { may } from './people.js';

const LABEL_X = 16;

/** A stored value as a control shows it. */
function shownValue(kind, v) {
  if (v === null || v === undefined) return kind === 'check' ? false : '';
  if (kind === 'check') return Boolean(v);
  if (kind === 'datetime') return String(v).slice(0, 16);
  return String(v);
}

export function FormView({ shell, db, name, onChanged, toast }) {
  const [layout, setLayout] = useState(null);
  const [count, setCount] = useState(0);
  const [at, setAt] = useState(0); // the record's place, count for a new one
  const [record, setRecord] = useState(null);
  const [draft, setDraft] = useState({});
  const [lists, setLists] = useState({});
  const [related, setRelated] = useState({});
  // The print templates for this form's table: the record shown printed, or saved as a PDF.
  const templates = (db.documents || []).filter((d) => d.table === layout?.table);
  const [template, setTemplate] = useState('');
  const printer = usePrinter(shell, toast);
  const printRecord = async (how) => {
    const chosen = template || templates[0]?.name;
    if (!chosen || !record) return;
    const key = record[layout.controls.find((c) => c.kind === 'auto')?.field || 'ID'] ?? record._rowid;
    try {
      const data = await shell.base.documentData({ id: db.id, name: chosen, keys: [key] });
      if (how === 'pdf') await printer.savePdf(data, `${chosen} ${key}`);
      else await printer.print(data);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };

  useEffect(() => {
    let live = true;
    shell.base.forms({ id: db.id }).then((all) => {
      const found = all.find((f) => f.name === name)?.layout || null;
      if (!live) return;
      setLayout(found);
      setAt(0);
    }).catch((e) => toast(e.message, { tone: 'bad' }));
    return () => { live = false; };
  }, [shell, db.id, name, toast]);

  // The lookup lists, once for the form.
  useEffect(() => {
    if (!layout) return undefined;
    let live = true;
    const lookups = layout.controls.filter((c) => c.kind === 'list' && c.lookup);
    Promise.all(lookups.map((c) => shell.base.lookup({ id: db.id, ...c.lookup }).then((items) => [c.field, items])))
      .then((pairs) => { if (live) setLists(Object.fromEntries(pairs)); })
      .catch(() => {});
    return () => { live = false; };
  }, [shell, db.id, layout]);

  const load = useCallback(async (place) => {
    if (!layout) return;
    const page = await shell.base.rows({ id: db.id, table: layout.table, offset: Math.max(0, place), limit: 1 });
    setCount(page.count);
    const r = place < page.count ? page.rows[0] || null : null;
    setRecord(r);
    setDraft(r ? { ...r } : {});
    const sub = {};
    for (const s of layout.subforms || []) {
      if (!r) { sub[s.table] = { rows: [], count: 0 }; continue; }
      sub[s.table] = await shell.base.rows({ id: db.id, table: s.table, where: { [s.field]: r[s.to] }, limit: 100 });
    }
    setRelated(sub);
  }, [shell, db.id, layout]);
  useEffect(() => { load(at).catch((e) => toast(e.message, { tone: 'bad' })); }, [at, load, toast]);

  const isNew = !record;
  /** A control left: its value kept, in a record shown; a new one waits for Add Record. */
  const keep = async (c, raw) => {
    const value = valueFromControl(c.kind, raw);
    setDraft((d) => ({ ...d, [c.field]: value }));
    if (isNew || c.readOnly) return;
    if ((record[c.field] ?? null) === value) return;
    try {
      await shell.base.update({ id: db.id, table: layout.table, rowid: record._rowid, values: { [c.field]: value } });
      setRecord((r) => ({ ...r, [c.field]: value }));
      onChanged?.();
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
      setDraft((d) => ({ ...d, [c.field]: record[c.field] }));
    }
  };
  const addRecord = async () => {
    const values = Object.fromEntries(layout.controls.filter((c) => !c.readOnly && draft[c.field] !== undefined).map((c) => [c.field, draft[c.field]]));
    try {
      await shell.base.insert({ id: db.id, table: layout.table, values });
      onChanged?.();
      setAt(count); // the record just added, the last
      await load(count);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  const remove = async () => {
    if (!record) return;
    try {
      await shell.base.remove({ id: db.id, table: layout.table, rowids: [record._rowid] });
      onChanged?.();
      const next = Math.min(at, Math.max(0, count - 2));
      setAt(next);
      await load(next);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };

  if (!layout) return <div className="bs-form-empty">{t('This form is not in the database.')}</div>;
  const field = (c) => {
    const v = shownValue(c.kind, draft[c.field]);
    const common = { className: `bs-fc bs-fc-${c.kind}`, 'data-field': c.field, 'aria-label': c.label, style: { left: c.x, top: c.y, width: c.w, height: c.h } };
    if (c.kind === 'check') return <input key={c.field} type="checkbox" {...common} checked={v} onChange={(e) => keep(c, e.target.checked)} />;
    if (c.kind === 'auto') return <span key={c.field} {...common}>{isNew ? t('(new)') : v}</span>;
    if (c.kind === 'list') {
      return (
        <select key={`${c.field}-${record?._rowid ?? 'new'}`} {...common} className="rw-select bs-fc bs-fc-list" defaultValue={v} onChange={(e) => keep(c, e.target.value)}>
          <option value="">{t('(none)')}</option>
          {(lists[c.field] || []).map((o) => <option key={String(o.value)} value={String(o.value)}>{o.text}</option>)}
        </select>
      );
    }
    if (c.kind === 'longtext') return <textarea key={`${c.field}-${record?._rowid ?? 'new'}`} {...common} className="rw-input bs-fc bs-fc-longtext" defaultValue={v} onBlur={(e) => keep(c, e.target.value)} />;
    const type = { date: 'date', datetime: 'datetime-local', number: 'text', currency: 'text' }[c.kind] || 'text';
    return <input key={`${c.field}-${record?._rowid ?? 'new'}`} type={type} {...common} className={`rw-input bs-fc bs-fc-${c.kind}`} defaultValue={v} onBlur={(e) => keep(c, e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }} />;
  };

  return (
    <div className="bs-form-run">
      {/* Its controls stand below the title (48 px), the page tall enough for them. */}
      <div className="bs-form-page" style={{ width: layout.width, minHeight: layout.height + 48 }} data-form={layout.name}>
        <h2 className="bs-form-title">{layout.title}</h2>
        {layout.controls.map((c) => (
          <React.Fragment key={c.field}>
            <span className="bs-fl" style={{ left: LABEL_X, top: c.y + 4 }}>{c.label}</span>
            {field(c)}
          </React.Fragment>
        ))}
        {(layout.subforms || []).map((s) => (
          <div key={s.table} className="bs-subform" style={{ left: s.x, top: s.y, width: s.w, height: s.h }} data-subform={s.table}>
            <div className="bs-subform-head">{s.label} <small>{tn(related[s.table]?.count || 0, '{count} record', '{count} records')}</small></div>
            <div className="bs-subform-body">
              <table>
                <thead><tr>{s.columns.map((col) => <th key={col}>{col}</th>)}</tr></thead>
                <tbody>{(related[s.table]?.rows || []).map((r) => <tr key={r._rowid}>{s.columns.map((col) => <td key={col}>{r[col] === null || r[col] === undefined ? '' : String(r[col])}</td>)}</tr>)}</tbody>
              </table>
            </div>
          </div>
        ))}
      </div>
      <div className="bs-form-bar" role="toolbar" aria-label={t('Records')}>
        <Button icon="skipBack" label="" title={t('First record')} className="bs-nav-first" disabled={at === 0} onClick={() => setAt(0)} />
        <Button icon="chevronLeft" label="" title={t('Previous record')} className="bs-nav-prev" disabled={at === 0} onClick={() => setAt(at - 1)} />
        <span className="bs-nav-place">{isNew ? t('New record') : t('Record {at} of {count}', { at: at + 1, count })}</span>
        <Button icon="chevronRight" label="" title={t('Next record')} className="bs-nav-next" disabled={at >= count - 1} onClick={() => setAt(at + 1)} />
        <Button icon="skipForward" label="" title={t('Last record')} className="bs-nav-last" disabled={at >= count - 1} onClick={() => setAt(Math.max(0, count - 1))} />
        <span className="grow" />
        {templates.length && !isNew ? (
          <span className="bs-form-print">
            {templates.length > 1 ? (
              <Select className="rw-select bs-form-template" aria-label={t('Print template')} value={template || templates[0].name} onChange={(e) => setTemplate(e.target.value)}>
                {templates.map((d) => <option key={d.name} value={d.name}>{d.name}</option>)}
              </Select>
            ) : null}
            <Button icon="print" label={templates.length > 1 ? t('Print') : t('Print {name}', { name: templates[0].name })} className="bs-form-print-now" onClick={() => printRecord('print')} />
            <Button icon="pdf" label="" title={t('Save as PDF')} className="bs-form-pdf" onClick={() => printRecord('pdf')} />
          </span>
        ) : null}
        {!may(db, layout?.table, 'add') ? null : isNew
          ? <Button primary icon="plus" label={t('Add Record')} className="bs-nav-add" onClick={addRecord} />
          : <Button icon="plus" label={t('New Record')} className="bs-nav-new" onClick={() => setAt(count)} />}
        {isNew || !may(db, layout?.table, 'delete') ? null : <Button icon="trash" label={t('Delete')} className="bs-nav-delete" onClick={remove} />}
      </div>
      {printer.portal}
    </div>
  );
}

export const FORM_CSS = `
.bs-form-run { flex: 1; min-height: 0; display: flex; flex-direction: column; }
.bs-form-run > .bs-form-page { margin: 16px; }
.bs-form-page { position: relative; flex: none; background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--r-3, 8px); padding-top: 48px; box-shadow: var(--shadow-1); }
.bs-form-run { overflow: auto; }
.bs-form-title { position: absolute; left: 16px; top: 12px; margin: 0; font-size: 16px; font-weight: 650; }
.bs-fl { position: absolute; width: 136px; font-size: 13px; color: var(--ink-2); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; transform: translateY(48px); }
.bs-fc { position: absolute; box-sizing: border-box; transform: translateY(48px); }
.bs-fc-auto { display: flex; align-items: center; color: var(--ink-3); font-variant-numeric: tabular-nums; }
.bs-fc-check { width: 18px !important; height: 18px !important; accent-color: var(--accent); margin: 4px 0 0; }
.bs-fc-number, .bs-fc-currency { text-align: end; }
.bs-fc-longtext { resize: none; }
.bs-subform { position: absolute; transform: translateY(48px); display: flex; flex-direction: column; border: 1px solid var(--line-soft); border-radius: var(--r-2); }
.bs-subform-head { padding: 6px 10px; font-size: 13px; font-weight: 600; border-bottom: 1px solid var(--line-soft); background: var(--sunken); }
.bs-subform-head small { color: var(--ink-3); font-weight: 400; margin-inline-start: 6px; }
.bs-subform-body { flex: 1; min-height: 0; overflow: auto; }
.bs-subform table { border-collapse: collapse; width: 100%; font-size: 12.5px; }
.bs-subform th, .bs-subform td { text-align: start; padding: 4px 8px; border-bottom: 1px solid var(--line-soft); }
.bs-form-bar { display: flex; align-items: center; gap: 4px; padding: 8px 16px; border-top: 1px solid var(--line-soft); background: var(--window); position: sticky; bottom: 0; }
.bs-form-print { display: inline-flex; gap: 4px; align-items: center; margin-inline-end: 8px; }
.bs-nav-place { min-width: 140px; text-align: center; font-size: 12.5px; color: var(--ink-2); font-variant-numeric: tabular-nums; }
.bs-form-empty { padding: 24px; color: var(--ink-3); }
`;
