// A Data Files form: one record at a time, laid out as its layout says (the
// form wizard's, kept in the file). Each control shows its field and keeps
// a change when it is left; a lookup is a list of the related table's
// records by their first words; the records elsewhere that point at this
// one are listed below, their lines changed in place and added at the foot
// as a datasheet's are; and the bar at the foot moves between records,
// starts a new one and deletes the one shown. Design moves, sizes, takes
// off and puts back its fields, and renames their labels (FormDesigner).

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Button, Dialog, Select, t, tn } from '@rutba/office-ui';
import { valueFromControl, labelLeft } from '@rutba/database/form-design';
import { usePrinter } from './documents.js';
import { FormDesigner, FORM_DESIGN_CSS } from './form-designer.js';
import { may, canDesign } from './people.js';

/** The control a subform's column is typed in, by its field's type. */
const SUB_KINDS = { integer: 'number', number: 'number', currency: 'currency', yesno: 'check', date: 'date', datetime: 'datetime', autonumber: 'auto' };

/** A stored value as a control shows it. */
function shownValue(kind, v) {
  if (v === null || v === undefined) return kind === 'check' ? false : '';
  if (kind === 'check') return Boolean(v);
  if (kind === 'datetime') return String(v).slice(0, 16);
  return String(v);
}

/** How many choices a form's list holds; a table with more is searched as it is typed in. */
const LIST_MAX = 500;

/**
 * A list too long to hold every choice (a sale's lines point at one of
 * 20,000 sales): a box showing the record it points at, searched as it is
 * typed in, the matches below to pick from.
 */
function SearchList({ shell, db, control: c, common, value, onPick }) {
  const [words, setWords] = useState('');
  const [typed, setTyped] = useState(null);
  const [found, setFound] = useState([]);
  const [open, setOpen] = useState(false);
  const timer = useRef(null);
  useEffect(() => {
    let live = true;
    if (value === '' || value === null || value === undefined) { setWords(''); return undefined; }
    shell.base.lookupPage({ id: db.id, table: c.lookup.table, field: c.lookup.field, values: [Number.isFinite(Number(value)) ? Number(value) : value] })
      .then((page) => { if (live) setWords(page.items[0]?.text ?? String(value)); }).catch(() => {});
    return () => { live = false; };
  }, [shell, db.id, c.lookup.table, c.lookup.field, value]);
  const search = (text) => {
    setTyped(text);
    setOpen(true);
    clearTimeout(timer.current);
    timer.current = setTimeout(() => {
      shell.base.lookupPage({ id: db.id, table: c.lookup.table, field: c.lookup.field, search: text, limit: 30 }).then((page) => setFound(page.items)).catch(() => {});
    }, 150);
  };
  const pick = (o) => { setOpen(false); setTyped(null); setWords(o.text); onPick(String(o.value)); };
  return (
    <span className="bs-fc bs-fc-search" style={common.style} data-field={c.field}>
      <input className="rw-input bs-fc-search-input" aria-label={c.label} role="combobox" aria-expanded={open} placeholder={t('Type to find')}
        value={typed ?? words} onFocus={() => search(typed ?? '')} onChange={(e) => search(e.target.value)}
        onKeyDown={(e) => { if (e.key === 'Escape') { setOpen(false); setTyped(null); } if (e.key === 'Enter' && found[0]) pick(found[0]); }}
        onBlur={() => setTimeout(() => { setOpen(false); setTyped(null); }, 150)} />
      {open && found.length ? (
        <ul className="bs-fc-search-list" role="listbox">
          {found.map((o) => <li key={String(o.value)} role="option" aria-selected={String(o.value) === String(value)} onMouseDown={(e) => { e.preventDefault(); pick(o); }}>{o.text}</li>)}
        </ul>
      ) : null}
    </span>
  );
}

/**
 * A subform: the lines of another table that point at the record shown,
 * each value changed where it stands (clicked, typed, kept when its box is
 * left), a new line typed in the row at the foot and added with Enter or
 * Add, pointing at the record shown; a line deleted after asking. What the person may do in
 * that table decides which of these they have.
 */
function SubformLines({ shell, db, sub: s, record, rows, count, words, onChanged, toast }) {
  const [fields, setFields] = useState([]);
  const [choices, setChoices] = useState({});
  const [fresh, setFresh] = useState({});
  const [removing, setRemoving] = useState(null);
  const [editing, setEditing] = useState(null); // `${rowid}.${col}`
  const mayChange = may(db, s.table, 'change');
  const mayAdd = may(db, s.table, 'add') && Boolean(record);
  const mayDelete = may(db, s.table, 'delete');
  const pointing = (col) => (db.relationships || []).find((r) => r.from.table === s.table && r.from.field === col);
  useEffect(() => {
    let live = true;
    shell.base.fields({ id: db.id, table: s.table }).then((f) => { if (live) setFields(f); }).catch(() => {});
    // A column that points at another table is a list of its records, as the form's own lists are.
    Promise.all(s.columns.filter((col) => pointing(col)).map((col) => {
      const rel = pointing(col);
      return shell.base.lookupPage({ id: db.id, table: rel.to.table, field: rel.to.field, limit: LIST_MAX }).then((page) => [col, page]);
    })).then((pairs) => { if (live) setChoices(Object.fromEntries(pairs)); }).catch(() => {});
    return () => { live = false; };
  }, [shell, db.id, s.table, db.steps]);
  const kindOf = (col) => (pointing(col) ? 'list' : SUB_KINDS[fields.find((f) => f.name === col)?.type] || 'text');
  const fixed = (col) => kindOf(col) === 'auto' || Boolean(fields.find((f) => f.name === col)?.readOnly);

  const keep = async (r, col, raw) => {
    setEditing(null);
    const value = valueFromControl(kindOf(col), raw);
    if ((r[col] ?? null) === value || (kindOf(col) === 'check' && Boolean(r[col]) === value)) return;
    try {
      await shell.base.update({ id: db.id, table: s.table, rowid: r._rowid, values: { [col]: value } });
      onChanged();
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
      onChanged();
    }
  };
  const add = async () => {
    const values = Object.fromEntries(Object.entries(fresh).filter(([, v]) => v !== '' && v !== undefined).map(([col, v]) => [col, valueFromControl(kindOf(col), v)]));
    if (!Object.keys(values).length) return;
    try {
      await shell.base.insert({ id: db.id, table: s.table, values: { ...values, [s.field]: record[s.to] } });
      setFresh({});
      onChanged();
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  const remove = async () => {
    const r = removing;
    setRemoving(null);
    try {
      await shell.base.remove({ id: db.id, table: s.table, rowids: [r._rowid] });
      onChanged();
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  const options = (col) => (choices[col]?.items || []).map((o) => <option key={String(o.value)} value={String(o.value)}>{o.text}</option>);

  const cell = (r, col) => {
    const kind = kindOf(col);
    const v = r[col];
    if (!mayChange || fixed(col)) return words(col, v);
    const key = `${r._rowid}-${col}-${String(v)}`;
    if (kind === 'check') return <input key={key} type="checkbox" className="bs-sub-cell" data-col={col} aria-label={col} checked={Boolean(v)} onChange={(e) => keep(r, col, e.target.checked)} />;
    // The value as words until it is clicked, as the record's lines read; then the box to change it in.
    if (editing !== `${r._rowid}.${col}`) {
      return <span className="bs-sub-text" role="button" tabIndex={0} title={t('Click to change')} onClick={() => setEditing(`${r._rowid}.${col}`)} onKeyDown={(e) => { if (e.key === 'Enter' || e.key === 'F2') setEditing(`${r._rowid}.${col}`); }}>{words(col, v)}</span>;
    }
    const leave = (e) => { if (e.key === 'Escape') setEditing(null); };
    if (kind === 'list') {
      return (
        <select key={key} autoFocus className="rw-select bs-sub-cell" data-col={col} aria-label={col} defaultValue={v === null || v === undefined ? '' : String(v)} onChange={(e) => keep(r, col, e.target.value)} onBlur={() => setEditing(null)} onKeyDown={leave}>
          <option value="">{t('(none)')}</option>
          {options(col)}
        </select>
      );
    }
    const type = { date: 'date', datetime: 'datetime-local' }[kind] || 'text';
    return <input key={key} autoFocus type={type} className={`rw-input bs-sub-cell bs-sub-${kind}`} data-col={col} aria-label={col} defaultValue={shownValue(kind, v)} onBlur={(e) => keep(r, col, e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); leave(e); }} />;
  };
  const newCell = (col) => {
    const kind = kindOf(col);
    if (fixed(col)) return <span className="bs-auto">{t('(new)')}</span>;
    if (kind === 'check') return <input type="checkbox" className="bs-sub-new" data-col={col} aria-label={col} checked={Boolean(fresh[col])} onChange={(e) => setFresh({ ...fresh, [col]: e.target.checked })} />;
    if (kind === 'list') {
      return (
        <select className="rw-select bs-sub-new" data-col={col} aria-label={col} value={fresh[col] ?? ''} onChange={(e) => setFresh({ ...fresh, [col]: e.target.value })}>
          <option value="">{t('(none)')}</option>
          {options(col)}
        </select>
      );
    }
    const type = { date: 'date', datetime: 'datetime-local' }[kind] || 'text';
    return <input type={type} className={`rw-input bs-sub-new bs-sub-${kind}`} data-col={col} aria-label={col} placeholder={col} value={fresh[col] ?? ''} onChange={(e) => setFresh({ ...fresh, [col]: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter') add(); }} />;
  };

  return (
    <div className="bs-subform" style={{ left: s.x, top: s.y, width: s.w, height: s.h }} data-subform={s.table}>
      <div className="bs-subform-head">{s.label} <small>{tn(count || 0, '{count} record', '{count} records')}</small></div>
      <div className="bs-subform-body">
        <table>
          <thead><tr>{s.columns.map((col) => <th key={col}>{col}</th>)}{mayDelete || mayAdd ? <th aria-label={t('Lines')} /> : null}</tr></thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r._rowid} data-rowid={r._rowid}>
                {s.columns.map((col) => <td key={col} data-col={col}>{cell(r, col)}</td>)}
                {mayDelete ? <td className="bs-sub-tools"><button type="button" className="bs-sub-remove" title={t('Delete this line')} aria-label={t('Delete this line')} onClick={() => setRemoving(r)}>×</button></td> : mayAdd ? <td /> : null}
              </tr>
            ))}
          </tbody>
          {/* The new line, typed at the foot as a datasheet's new record is. */}
          {mayAdd ? (
            <tfoot>
              <tr className="bs-sub-fresh">
                {s.columns.map((col) => <td key={col} data-col={col}>{newCell(col)}</td>)}
                <td className="bs-sub-tools"><button type="button" className="bs-sub-add" title={t('Add this line')} aria-label={t('Add this line')} onClick={add}>+</button></td>
              </tr>
            </tfoot>
          ) : null}
        </table>
      </div>
      {removing ? (
        <Dialog title={t('Delete this line?')} onClose={() => setRemoving(null)} actions={<><Button label={t('Cancel')} className="bs-sub-remove-cancel" onClick={() => setRemoving(null)} /><Button primary label={t('Delete')} className="bs-sub-remove-yes" onClick={remove} /></>}>
          <p>{t('The line is deleted from {table}, and cannot be brought back.', { table: s.table })}</p>
        </Dialog>
      ) : null}
    </div>
  );
}

export function FormView({ shell, db, name, onChanged, onMakeTemplate = null, toast }) {
  const [layout, setLayout] = useState(null);
  const [count, setCount] = useState(0);
  const [at, setAt] = useState(0); // the record's place, count for a new one
  const [record, setRecord] = useState(null);
  const [draft, setDraft] = useState({});
  const [lists, setLists] = useState({});
  const [related, setRelated] = useState({});
  // A subform's field that points at another table, shown as that record's words: `${table}.${field}` -> Map(value -> words).
  const [subLists, setSubLists] = useState({});
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
    // A page of each list's choices; one whose table holds more is searched as it is typed in (SearchList).
    const lookups = layout.controls.filter((c) => c.kind === 'list' && c.lookup);
    Promise.all(lookups.map((c) => shell.base.lookupPage({ id: db.id, table: c.lookup.table, field: c.lookup.field, limit: LIST_MAX }).then((page) => [c.field, page])))
      .then((pairs) => { if (live) setLists(Object.fromEntries(pairs)); })
      .catch(() => {});
    return () => { live = false; };
  }, [shell, db.id, layout]);

  useEffect(() => {
    if (!layout) return undefined;
    let live = true;
    const pairs = [];
    for (const s of layout.subforms || []) {
      const rows = related[s.table]?.rows || [];
      for (const col of s.columns) {
        if (col === s.field) continue;
        const rel = (db.relationships || []).find((r) => r.from.table === s.table && r.from.field === col);
        const values = [...new Set(rows.map((r) => r[col]).filter((v) => v !== null && v !== undefined))];
        if (rel && values.length) pairs.push(shell.base.lookupPage({ id: db.id, table: rel.to.table, field: rel.to.field, values }).then((page) => [`${s.table}.${col}`, new Map(page.items.map((o) => [o.value, o.text]))]));
      }
    }
    Promise.all(pairs).then((done) => { if (live) setSubLists(Object.fromEntries(done)); }).catch(() => {});
    return () => { live = false; };
  }, [shell, db.id, layout, db.relationships, related]);
  const subValue = (tbl, col, v) => {
    if (v === null || v === undefined) return '';
    const words = subLists[`${tbl}.${col}`]?.get(v);
    return words === undefined ? String(v) : words;
  };

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
  /** A subform's lines read again, after one is changed, added or deleted there. */
  const reloadSub = async (s) => {
    if (!record) return;
    const page = await shell.base.rows({ id: db.id, table: s.table, where: { [s.field]: record[s.to] }, limit: 100 });
    setRelated((was) => ({ ...was, [s.table]: page }));
  };
  const [designing, setDesigning] = useState(false);
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
  // Delete asks first: a record deleted is gone, and the form shows the next.
  const [asking, setAsking] = useState(false);
  const remove = async () => {
    setAsking(false);
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
  if (designing) {
    return <FormDesigner shell={shell} db={db} layout={layout} toast={toast} onCancel={() => setDesigning(false)}
      onSaved={(summary, kept) => { setLayout(kept); setDesigning(false); onChanged?.(); }} />;
  }
  const field = (c) => {
    const v = shownValue(c.kind, draft[c.field]);
    const common = { className: `bs-fc bs-fc-${c.kind}`, 'data-field': c.field, 'aria-label': c.label, style: { left: c.x, top: c.y, width: c.w, height: c.h } };
    if (c.kind === 'check') return <input key={c.field} type="checkbox" {...common} checked={v} onChange={(e) => keep(c, e.target.checked)} />;
    if (c.kind === 'auto') return <span key={c.field} {...common}>{isNew ? t('(new)') : v}</span>;
    if (c.kind === 'list' && lists[c.field]?.more) {
      return <SearchList key={`${c.field}-${record?._rowid ?? 'new'}`} shell={shell} db={db} control={c} common={common} value={v} onPick={(value) => keep(c, value)} />;
    }
    if (c.kind === 'list') {
      return (
        <select key={`${c.field}-${record?._rowid ?? 'new'}`} {...common} className="rw-select bs-fc bs-fc-list" defaultValue={v} onChange={(e) => keep(c, e.target.value)}>
          <option value="">{t('(none)')}</option>
          {(lists[c.field]?.items || []).map((o) => <option key={String(o.value)} value={String(o.value)}>{o.text}</option>)}
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
            <span className="bs-fl" style={{ left: labelLeft(c), top: c.y + 4 }}>{c.label}</span>
            {field(c)}
          </React.Fragment>
        ))}
        {(layout.subforms || []).map((s) => (
          <SubformLines key={s.table} shell={shell} db={db} sub={s} record={record} rows={related[s.table]?.rows || []} count={related[s.table]?.count || 0}
            words={(col, v) => subValue(s.table, col, v)} toast={toast} onChanged={() => reloadSub(s).catch((e) => toast(e.message, { tone: 'bad' }))} />
        ))}
      </div>
      <div className="bs-form-bar" role="toolbar" aria-label={t('Records')}>
        <Button icon="skipBack" label="" title={t('First record')} className="bs-nav-first" disabled={at === 0} onClick={() => setAt(0)} />
        <Button icon="chevronLeft" label="" title={t('Previous record')} className="bs-nav-prev" disabled={at === 0} onClick={() => setAt(at - 1)} />
        <span className="bs-nav-place">{isNew ? t('New record') : t('Record {at} of {count}', { at: at + 1, count })}</span>
        <Button icon="chevronRight" label="" title={t('Next record')} className="bs-nav-next" disabled={at >= count - 1} onClick={() => setAt(at + 1)} />
        <Button icon="skipForward" label="" title={t('Last record')} className="bs-nav-last" disabled={at >= count - 1} onClick={() => setAt(Math.max(0, count - 1))} />
        <span className="grow" />
        {canDesign(db) ? <Button icon="pen" label={t('Design')} className="bs-form-design" title={t('Move, size, take off and put back the fields, and rename their labels')} onClick={() => setDesigning(true)} /> : null}
        {!templates.length && !isNew && canDesign(db) ? (
          // No print template for this table: where its Print buttons would be, how to make one.
          <span className="bs-form-no-print" data-no-template={layout.table}>
            <span className="bs-hint">{t('To print this record, make a print template for {table}.', { table: layout.table })}</span>
            {onMakeTemplate ? <Button icon="print" label={t('New Print Template')} className="bs-form-make-template" onClick={() => onMakeTemplate(layout.table)} /> : null}
          </span>
        ) : null}
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
        {isNew || !may(db, layout?.table, 'delete') ? null : <Button icon="trash" label={t('Delete')} className="bs-nav-delete" onClick={() => setAsking(true)} />}
      </div>
      {asking ? (
        <Dialog title={t('Delete this record?')} onClose={() => setAsking(false)} actions={<><Button label={t('Cancel')} className="bs-delete-cancel" onClick={() => setAsking(false)} /><Button primary label={t('Delete')} className="bs-delete-yes" onClick={remove} /></>}>
          <p>{t('Record {at} of {table} is deleted, and cannot be brought back.', { at: at + 1, table: layout.table })}</p>
        </Dialog>
      ) : null}
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
/* A list in a form's row: no top and bottom padding, or Chromium squeezes its words into what is left and cuts them. */
.bs-fc.bs-fc-list { padding-top: 0; padding-bottom: 0; }
.bs-fc-search { display: block; }
/* Each control's transform makes it a stacking context of its own, painted in turn, so the matches' own z-index cannot rise above the fields after it: the control being typed in rises instead. */
.bs-fc-search:focus-within { z-index: 6; }
.bs-fc-search-input { width: 100%; height: 100%; box-sizing: border-box; padding-top: 0; padding-bottom: 0; }
.bs-fc-search-list { position: absolute; z-index: 5; left: 0; right: 0; top: 100%; margin: 2px 0 0; padding: 4px 0; list-style: none; max-height: 220px; overflow: auto; background: var(--surface); border: 1px solid var(--line-strong); border-radius: var(--r-2); box-shadow: var(--shadow-2, 0 4px 12px rgba(0,0,0,.12)); }
.bs-fc-search-list li { padding: 4px 10px; cursor: pointer; font-size: 13px; }
.bs-fc-search-list li:hover, .bs-fc-search-list li[aria-selected="true"] { background: var(--accent-soft); }
.bs-subform { position: absolute; transform: translateY(48px); display: flex; flex-direction: column; border: 1px solid var(--line-soft); border-radius: var(--r-2); }
.bs-subform-head { padding: 6px 10px; font-size: 13px; font-weight: 600; border-bottom: 1px solid var(--line-soft); background: var(--sunken); }
.bs-subform-head small { color: var(--ink-3); font-weight: 400; margin-inline-start: 6px; }
.bs-subform-body { flex: 1; min-height: 0; overflow: auto; }
.bs-subform table { border-collapse: collapse; width: 100%; font-size: 12.5px; }
.bs-subform th, .bs-subform td { text-align: start; padding: 4px 8px; border-bottom: 1px solid var(--line-soft); }
.bs-subform td:has(> .bs-sub-cell, > .bs-sub-new) { padding: 2px 4px; }
.bs-sub-cell, .bs-sub-new { width: 100%; box-sizing: border-box; height: 24px; font-size: 12.5px; }
.bs-sub-text { display: block; min-height: 18px; cursor: text; border-radius: 3px; }
.bs-sub-text:hover { background: var(--hover); }
.bs-sub-text:focus-visible { outline: 2px solid var(--accent); }
.bs-sub-cell[type="checkbox"], .bs-sub-new[type="checkbox"] { width: 16px; height: 16px; accent-color: var(--accent); }
.bs-sub-number, .bs-sub-currency { text-align: end; }
.bs-sub-fresh td { background: var(--sunken); }
.bs-sub-tools { width: 24px; text-align: center; }
.bs-sub-remove, .bs-sub-add { border: 0; background: transparent; color: var(--ink-3); font-size: 15px; line-height: 1; cursor: pointer; padding: 2px 4px; border-radius: 3px; }
.bs-sub-remove:hover, .bs-sub-add:hover { color: var(--ink); background: var(--hover); }
.bs-form-bar { display: flex; align-items: center; gap: 4px; padding: 8px 16px; border-top: 1px solid var(--line-soft); background: var(--window); position: sticky; bottom: 0; }
.bs-form-print { display: inline-flex; gap: 4px; align-items: center; margin-inline-end: 8px; }
.bs-form-no-print { display: inline-flex; gap: 6px; align-items: center; margin-inline-end: 8px; }
.bs-form-no-print .bs-hint { font-size: 12px; color: var(--ink-3); }
.bs-nav-place { min-width: 140px; text-align: center; font-size: 12.5px; color: var(--ink-2); font-variant-numeric: tabular-nums; }
.bs-form-empty { padding: 24px; color: var(--ink-3); }
${FORM_DESIGN_CSS}`;
