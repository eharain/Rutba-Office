// Databases (2.9, D-2026-10-20-17): tables, their design, relationships
// and SQL over one SQLite file. Kept plain on purpose: a list of tables on
// the left and, for the one chosen, its records or its design; the
// relationships and a SQL query each a view of their own. The file is
// held by the main process (main/base.js); this window asks it for a page
// of records at a time, so a table of any size opens at once.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, Input, Select, Content, Empty, Dialog, Spinner, useToast, useMenu, t, tn } from '@rutba/office-ui';
import { AppFrame, useAppMenu, pickOpen, useFileDrop } from '../shell.js';
import { BASE_CSS } from './base/css.js';
import { QueryDesigner, SqlView, QUERY_CSS } from './base/queries.js';
import { FormView, FORM_CSS } from './base/forms.js';
import { ReportView, REPORT_CSS } from './base/reports.js';
import { RecentFiles } from '../start-page.js';
import { DocumentView, DOCUMENT_CSS } from './base/documents.js';
import { useGridFilter, FilterBar, ColumnFilter, selectionItems, FILTER_CSS } from './base/filters.js';
import { SignIn, ProtectDialog, PasswordDialog, SeeingBanner, PeopleView, PEOPLE_CSS, canDesign, may } from './base/people.js';
import { LinkDialog, LinkedBar, LinkPasswordDialog, ClashesDialog, LINKS_CSS } from './base/links.js';
import { SourceDialog, SourceSignInDialog, SourceKeysDialog, SOURCES_CSS } from './base/sources.js';
import { SourcesQuestion, TRUST_CSS } from './base/trust.js';
import { AiToolsButton, AiToolsDialog as AiToolsDatabaseDialog } from './base/ai-tools.js';

/** The ready-made databases the start screen offers, in the engine's order (packages/database/src/templates.js). */
const TEMPLATES = [
  { id: 'contacts', icon: 'contacts', name: () => t('common.contacts'), line: () => t('base.people-how-reach') },
  { id: 'inventory', icon: 'archive', name: () => t('base.inventory'), line: () => t('base.products-suppliers-stock') },
  { id: 'membership', icon: 'star', name: () => t('base.membership'), line: () => t('base.members-payments') },
  { id: 'crm', icon: 'chartColumn', name: () => t('base.customers-and-sales'), line: () => t('base.companies-people-open') },
  { id: 'library', icon: 'list', name: () => t('base.lending-library'), line: () => t('base.books-borrowers-loans') },
];

/**
 * The sample businesses the start screen offers, in the engine's order
 * (packages/templates/data-files, the templates' own repository) and in
 * groups: each with its invented records to learn from, or empty to start a
 * business of one's own from. The business's name is a made-up proper name,
 * the same in every language (tests/database-samples.test.js holds it to the
 * sample's own).
 */
const SAMPLE_GROUPS = [
  { id: 'shops', name: () => t('base.shops-and-food') },
  { id: 'trade', name: () => t('base.trade-making-building') },
  { id: 'services', name: () => t('base.services') },
  { id: 'care', name: () => t('base.health-learning') },
  { id: 'property', name: () => t('base.property-premises-staff') },
];
const SAMPLES = [
  { id: 'shop', group: 'shops', icon: 'scanReceipt', business: 'Larkspur Corner Store', kind: () => t('base.shop'), line: () => t('base.products-till-stock') },
  { id: 'food', group: 'shops', icon: 'archive', business: 'Bramblegate Wine and Deli', kind: () => t('base.food-and-beverages'), line: () => t('base.wines-beers-spirits') },
  { id: 'cafe', group: 'shops', icon: 'sun', business: 'Little Fig Cafe', kind: () => t('base.cafe-or-restaurant'), line: () => t('base.menu-costing-orders') },
  { id: 'takeaway', group: 'shops', icon: 'send', business: 'Saffron Lane Pizza and Curry', kind: () => t('base.takeaway'), line: () => t('base.phone-online-orders') },
  { id: 'florist', group: 'shops', icon: 'star', business: 'Hazel Row Flowers', kind: () => t('base.florist'), line: () => t('base.orders-card-messages') },
  { id: 'pharmacy', group: 'shops', icon: 'heal', business: 'Tansy Row Pharmacy', kind: () => t('base.pharmacy-counter'), line: () => t('base.counter-stock-batch') },
  { id: 'trading', group: 'trade', icon: 'globe', business: 'Northbank Trading', kind: () => t('base.wholesale-and'), line: () => t('base.customers-suppliers') },
  { id: 'workshop', group: 'trade', icon: 'layers', business: 'Granite and Grain Woodworks', kind: () => t('base.workshop-or-maker'), line: () => t('base.materials-bills') },
  { id: 'builder', group: 'trade', icon: 'grid', business: 'Kilnside Builders', kind: () => t('base.builder-or'), line: () => t('base.enquiries-quotes-jobs') },
  { id: 'cars', group: 'trade', icon: 'flag', business: 'Copperbeech Motors', kind: () => t('base.car-dealer'), line: () => t('base.vehicles-stock-costs') },
  { id: 'salon', group: 'services', icon: 'scissors', business: 'Marigold Hair and Beauty', kind: () => t('base.salon'), line: () => t('base.clients-bookings') },
  { id: 'barber', group: 'services', icon: 'brush', business: 'Copper Pole Barbers', kind: () => t('base.barber-shop'), line: () => t('base.walk-ins-bookings') },
  { id: 'repairs', group: 'services', icon: 'phone', business: 'Ironbridge Repairs', kind: () => t('base.repair-shop'), line: () => t('base.repair-jobs-booking') },
  { id: 'garage', group: 'services', icon: 'settings', business: 'Quarrybank Motors', kind: () => t('base.garage-mot-centre'), line: () => t('base.vehicles-job-cards') },
  { id: 'cleaning', group: 'services', icon: 'bucket', business: 'Fernhollow Cleaning', kind: () => t('base.cleaning-company'), line: () => t('base.clients-sites-contracts') },
  { id: 'agency', group: 'services', icon: 'clock', business: 'Bramble Studio', kind: () => t('base.consultancy-or'), line: () => t('base.clients-projects-time') },
  { id: 'gym', group: 'services', icon: 'sliders', business: 'Tannery Yard Fitness', kind: () => t('base.gym-fitness-studio'), line: () => t('base.members-plans-payments') },
  { id: 'clinic', group: 'care', icon: 'calendar', business: 'Cedar Lane Clinic', kind: () => t('base.clinic'), line: () => t('base.patients-appointments') },
  { id: 'school', group: 'care', icon: 'pen', business: 'Lantern Learning Centre', kind: () => t('base.school-training-centre'), line: () => t('base.students-courses-classes') },
  { id: 'charity', group: 'care', icon: 'heart', business: 'Rowanmere Community Trust', kind: () => t('base.charity-or-club'), line: () => t('base.members-renewals') },
  { id: 'rentals', group: 'property', icon: 'home', business: 'Hollybrook Lettings', kind: () => t('base.property-rentals'), line: () => t('base.properties-units-tenants') },
  { id: 'agent', group: 'property', icon: 'folderOpen', business: 'Fernleigh Estates', kind: () => t('base.estate-letting-agent'), line: () => t('base.landlords-sellers') },
  { id: 'guesthouse', group: 'property', icon: 'moon', business: 'Gullhaven Guest House', kind: () => t('base.guest-house-b'), line: () => t('base.rooms-bookings-deposits') },
  { id: 'visitors', group: 'property', icon: 'signature', business: 'Copperfield House', kind: () => t('base.visitors-front-desk'), line: () => t('base.building-passes-hosts') },
  { id: 'payroll', group: 'property', icon: 'financial', business: 'Wrenmoor Garden Centre', kind: () => t('base.staff-timesheets-payroll'), line: () => t('base.employees-holidays') },
];

export const TYPE_NAMES = () => ({
  autonumber: t('base.auto-number'),
  text: t('base.short-text'),
  longtext: t('base.long-text'),
  number: t('common.number'),
  integer: t('common.whole-number'),
  currency: t('common.currency'),
  date: t('common.date'),
  datetime: t('common.date-and-time'),
  yesno: t('base.yes-no'),
});
const TYPES = ['text', 'longtext', 'number', 'integer', 'currency', 'date', 'datetime', 'yesno'];
const PAGE = 100;
const ROW_H = 30;
const OPENS = /\.(rbase|sqlite3?|db)$/i;
/** How often a linked table on screen is read again from its engine. */
const LIVE_MS = 15000;
const IMPORTS = ['csv', 'tsv', 'txt', 'xlsx', 'xlsm', 'ods', 'dbf', 'mdb', 'accdb', 'wk1', 'wks', 'wk3', 'wk4', 'slk', 'dif'];
/** The workbooks among them: each sheet a table, or the one chosen. */
const WORKBOOKS = /\.(xlsx|xlsm|ods)$/i;

/** A value as a cell shows it. */
function shown(field, value) {
  if (value === null || value === undefined) return '';
  if (field?.type === 'currency' && typeof value === 'number') return value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 4 });
  if (field?.type === 'datetime') return String(value).replace('T', ' ');
  return String(value);
}

/** A field's default as the new row shows it: an amount with its two places (1.50). */
function shownDefault(field) {
  const n = Number(field.default);
  if (field.type === 'currency' && String(field.default).trim() !== '' && Number.isFinite(n)) return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 4 });
  return shown(field, field.default);
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
  // Searched and filtered in the database (D-2026-10-22-07); the menu offers filter by selection.
  const grid = useGridFilter();
  const [filtering, setFiltering] = useState(null);
  const menu = useMenu();
  // A field that points at another table's record shows that record's words (Semi-skimmed milk, not 24); the number is what is kept and typed.
  const pointing = useMemo(() => (db.relationships || []).filter((r) => r.from.table === table), [db.relationships, table]);
  const [words, setWords] = useState({});
  const nameRows = useCallback(async (rows) => {
    for (const r of pointing) {
      const values = [...new Set(rows.map((x) => x[r.from.field]).filter((v) => v !== null && v !== undefined && v !== ''))];
      if (!values.length) continue;
      const page = await shell.base.lookupPage({ id: db.id, table: r.to.table, field: r.to.field, values });
      setWords((w) => {
        const m = new Map(w[r.from.field] || []);
        for (const o of page.items) m.set(o.value, o.text);
        return { ...w, [r.from.field]: m };
      });
    }
  }, [shell, db.id, pointing]);

  useEffect(() => {
    let live = true;
    shell.base.fields({ id: db.id, table }).then((f) => { if (live) setFields(f); }).catch((e) => toast(e.message, { tone: 'bad' }));
    pages.current = new Map();
    setPicked(new Set());
    setEditing(null);
    return () => { live = false; };
  }, [shell, db.id, table, version, toast]);
  useEffect(() => { grid.clear(); }, [table]); // eslint-disable-line react-hooks/exhaustive-deps

  const load = useCallback(async (page) => {
    if (pages.current.has(page)) return;
    pages.current.set(page, null);
    try {
      const r = await shell.base.rows({ id: db.id, table, offset: page * PAGE, limit: PAGE, order, search: grid.search, filters: grid.filters });
      pages.current.set(page, r.rows);
      setCount(r.count);
      redraw((n) => n + 1);
      nameRows(r.rows).catch(() => {});
    } catch (e) {
      pages.current.delete(page);
      toast(e.message, { tone: 'bad' });
    }
  }, [shell, db.id, table, order, grid.search, grid.filters, toast, nameRows]);

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
      toast(tn(picked.size, 'base.count-record-deleted'));
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 6000 });
    }
  };

  const editable = fields.filter((f) => f.type !== 'autonumber');
  // What the person's role allows here (D-2026-10-22-11); the database refuses the rest in any case.
  // A copy read from a source, or of an engine table with no primary key, is read only here.
  const readOnly = Boolean(db.tables.find((x) => x.name === table)?.readOnly);
  const mayAdd = !readOnly && may(db, table, 'add');
  const mayChange = !readOnly && may(db, table, 'change');
  const mayDelete = !readOnly && may(db, table, 'delete');
  const fixed = (f) => !mayChange || Boolean(f.readOnly);
  const cols = fields.map((f) => (f.type === 'longtext' ? 280 : f.type === 'yesno' ? 90 : f.type === 'autonumber' ? 80 : 160));
  const template = `36px ${cols.map((w) => `${w}px`).join(' ')}`;
  const wide = 36 + cols.reduce((a, b) => a + b, 0);
  const sortBy = (name) => { setOrder((o) => (o === name ? `-${name}` : o === `-${name}` ? null : name)); pages.current = new Map(); };

  const cell = (row, f) => {
    if (f.type === 'yesno') {
      return <input type="checkbox" checked={Boolean(row[f.name])} aria-label={f.name} disabled={fixed(f)} onChange={(e) => save(row._rowid, f, e.target.checked)} />;
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
    const named = words[f.name]?.get(row[f.name]);
    if (named !== undefined) return <span className="bs-text bs-named" data-value={row[f.name]} title={t('base.field-value', { field: f.name, value: row[f.name] })}>{named}</span>;
    return <span className="bs-text">{shown(f, row[f.name])}</span>;
  };

  return (
    <div className="bs-sheet">
      <div className="bs-tools">
        {mayDelete ? <Button icon="trash" label={t('common.delete')} disabled={!picked.size} className="bs-delete" onClick={removePicked} /> : null}
        <span className="bs-count">{grid.active ? tn(count, 'base.count-record-found') : tn(count, 'base.count-record')}</span>
      </div>
      <FilterBar grid={grid} />
      <div className="bs-grid" ref={box} onScroll={(e) => setScrollTop(e.currentTarget.scrollTop)} role="grid" aria-rowcount={count + 1}>
        <div className="bs-row bs-head" style={{ gridTemplateColumns: template }} role="row">
          <span />
          {fields.map((f) => (
            <div key={f.name} className="bs-th-cell" role="columnheader" aria-sort={order === f.name ? 'ascending' : order === `-${f.name}` ? 'descending' : 'none'}>
              <button type="button" className="bs-th" data-field={f.name} onClick={() => sortBy(f.name)}>
                <span>{f.name}</span>
                {order === f.name ? <Icon name="chevronUp" size={12} /> : order === `-${f.name}` ? <Icon name="chevronDown" size={12} /> : null}
              </button>
              <button type="button" className={`bs-th-filter${grid.filters.some((x) => x.column === f.name) ? ' on' : ''}`} data-filter-field={f.name} title={t('base.filter-by-column', { column: f.name })} aria-label={t('base.filter-by-column', { column: f.name })} onClick={() => setFiltering({ name: f.name, type: f.type })}>
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
              <div key={i} className={`bs-row${row && picked.has(row._rowid) ? ' picked' : ''}`} role="row" data-row={i} data-rowid={row?._rowid} style={{ top: i * ROW_H, gridTemplateColumns: template }}>
                <span className="bs-pick">
                  {row ? <input type="checkbox" aria-label={t('base.pick')} checked={picked.has(row._rowid)} onChange={(e) => setPicked((p) => { const n = new Set(p); if (e.target.checked) n.add(row._rowid); else n.delete(row._rowid); return n; })} /> : null}
                </span>
                {fields.map((f) => (
                  <span
                    key={f.name}
                    className={`bs-td t-${f.type}`}
                    role="gridcell"
                    data-field={f.name}
                    data-readonly={fixed(f) || undefined}
                    onDoubleClick={() => row && f.type !== 'autonumber' && f.type !== 'yesno' && !fixed(f) && setEditing({ rowid: row._rowid, field: f.name, value: row[f.name] ?? '' })}
                    onContextMenu={(e) => { if (!row) return; e.preventDefault(); menu.open(e, selectionItems(grid, f.name, row[f.name])); }}
                  >
                    {row ? cell(row, f) : null}
                  </span>
                ))}
              </div>
            );
          })}
        </div>
        {mayAdd ? <div className="bs-row bs-new" style={{ gridTemplateColumns: template }} role="row">
          <span className="bs-pick"><Icon name="plus" size={14} /></span>
          {fields.map((f) => (
            <span key={f.name} className="bs-td">
              {f.type === 'autonumber' ? <span className="bs-auto">{t('base.new')}</span>
                : f.type === 'yesno' ? <input type="checkbox" aria-label={f.name} checked={Boolean(fresh[f.name] ?? (f.default === true || f.default === 1 || /^(1|yes|true)$/i.test(String(f.default ?? ''))))} onChange={(e) => setFresh({ ...fresh, [f.name]: e.target.checked })} />
                  // A field with a default shows it, as the record will have it unless something else is typed.
                  : <input className={`bs-new-input${f.default !== undefined && f.default !== null && f.default !== '' ? ' has-default' : ''}`} data-field={f.name} aria-label={f.name} placeholder={f.default !== undefined && f.default !== null && f.default !== '' ? shownDefault(f) : TYPE_NAMES()[f.type]} disabled={Boolean(f.readOnly)} value={fresh[f.name] ?? ''} onChange={(e) => setFresh({ ...fresh, [f.name]: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter') addRecord(); }} />}
            </span>
          ))}
        </div> : null}
      </div>
      {editable.length || !canDesign(db) ? null : <p className="bs-hint">{t('base.table-no-fields')}</p>}
      {filtering ? <ColumnFilter column={filtering} onApply={grid.put} onClose={() => setFiltering(null)} /> : null}
      {menu.node}
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
  // A change some records' values do not fit: the database names them, and they can be cleared to make it.
  const [unfit, setUnfit] = useState(null);
  const change = async (f, change) => {
    try {
      onChanged(await shell.base.changeField({ id: db.id, table, name: f.name, change }));
      setVersion((v) => v + 1);
    } catch (e) {
      if (/Clear (it|them) and change the type/.test(e.original || e.message)) setUnfit({ field: f.name, change, message: e.message });
      else toast(e.message, { tone: 'bad', ms: 7000 });
      setVersion((v) => v + 1);
    }
  };
  const clearAndChange = () => {
    const u = unfit;
    setUnfit(null);
    run(() => shell.base.changeField({ id: db.id, table, name: u.field, change: u.change, clear: true }));
  };
  return (
    <div className="bs-design">
      {unfit ? (
        <Dialog title={t('base.some-records-do')} onClose={() => setUnfit(null)} actions={<><Button label={t('common.cancel')} className="bs-unfit-cancel" onClick={() => setUnfit(null)} /><Button primary label={t('base.clear-change')} className="bs-unfit-clear" onClick={clearAndChange} /></>}>
          <p className="bs-unfit-words" role="alert">{unfit.message}</p>
          <p className="bs-hint">{t('base.clearing-leaves-values')}</p>
        </Dialog>
      ) : null}
      <table className="bs-fields">
        <thead>
          <tr><th>{t('common.field')}</th><th>{t('common.type')}</th><th>{t('common.size')}</th><th>{t('base.required')}</th><th>{t('base.unique')}</th><th>{t('common.default')}</th><th /></tr>
        </thead>
        <tbody>
          {fields.map((f) => (
            <tr key={f.name} data-field={f.name}>
              <td><Input className="rw-input bs-f-name" aria-label={t('base.field-name')} defaultValue={f.name} key={`n-${f.name}-${version}`} onBlur={(e) => { const v = e.target.value.trim(); if (v && v !== f.name) change(f, { name: v }); }} onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }} /></td>
              <td>
                {f.type === 'autonumber' ? <span className="bs-fixed">{names.autonumber}</span> : (
                  <Select className="rw-select bs-f-type" aria-label={t('common.type')} value={f.type} onChange={(e) => change(f, { type: e.target.value })}>
                    {TYPES.map((k) => <option key={k} value={k}>{names[k]}</option>)}
                  </Select>
                )}
              </td>
              <td>{f.type === 'text' ? <Input className="rw-input bs-f-size" aria-label={t('common.size')} type="number" defaultValue={f.size ?? 255} key={`s-${f.name}-${version}`} onBlur={(e) => { const v = Number(e.target.value); if (v > 0 && v !== (f.size ?? 255)) change(f, { size: v }); }} /> : null}</td>
              <td>{f.type === 'autonumber' || f.type === 'yesno' ? null : <input type="checkbox" aria-label={t('base.required')} className="bs-f-required" checked={f.required} onChange={(e) => change(f, { required: e.target.checked })} />}</td>
              <td>{f.type === 'autonumber' || f.type === 'yesno' ? null : <input type="checkbox" aria-label={t('base.unique')} className="bs-f-unique" checked={f.unique} onChange={(e) => change(f, { unique: e.target.checked })} />}</td>
              <td>{f.type === 'autonumber' ? null : <Input className="rw-input bs-f-default" aria-label={t('common.default')} defaultValue={f.default ?? ''} key={`d-${f.name}-${version}`} onBlur={(e) => { const v = e.target.value; if (v !== String(f.default ?? '')) change(f, { default: v === '' ? undefined : v }); }} />}</td>
              <td>{f.type === 'autonumber' ? null : <Button icon="close" label="" title={t('base.remove-field-data')} className="bs-f-remove" onClick={() => run(() => shell.base.removeField({ id: db.id, table, name: f.name }))} />}</td>
            </tr>
          ))}
          <tr className="bs-f-new">
            <td><Input className="rw-input bs-add-name" aria-label={t('base.new-field')} placeholder={t('base.new-field')} value={adding.name} onChange={(e) => setAdding({ ...adding, name: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter' && adding.name.trim()) run(() => shell.base.addField({ id: db.id, table, field: { ...adding, name: adding.name.trim() } }), () => setAdding({ name: '', type: adding.type })); }} /></td>
            <td>
              <Select className="rw-select bs-add-type" aria-label={t('common.type')} value={adding.type} onChange={(e) => setAdding({ ...adding, type: e.target.value })}>
                {TYPES.map((k) => <option key={k} value={k}>{names[k]}</option>)}
              </Select>
            </td>
            <td colSpan={5}>
              <Button icon="plus" label={t('base.add-field')} className="bs-add-field" disabled={!adding.name.trim()} onClick={() => run(() => shell.base.addField({ id: db.id, table, field: { ...adding, name: adding.name.trim() } }), () => setAdding({ name: '', type: adding.type }))} />
            </td>
          </tr>
        </tbody>
      </table>
      <div className="bs-design-foot">
        <Button icon="trash" label={t('base.delete-table')} className="bs-drop" onClick={() => run(() => shell.base.dropTable({ id: db.id, table }), onRemoved)} />
        <span className="bs-hint">{t('base.each-change-applied')}</span>
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
      <h2>{t('base.relationships')}</h2>
      <p className="bs-hint">{t('base.relationship-lets-many')}</p>
      <ul className="bs-rel-list">
        {db.relationships.map((r) => (
          <li key={r.id} className="bs-rel" data-relation={r.id}>
            <span><strong>{r.from.table}</strong>.{r.from.field}</span>
            <Icon name="chevronRight" size={14} />
            <span><strong>{r.to.table}</strong>.{r.to.field}</span>
            {r.cascadeDelete ? <span className="bs-chip">{t('base.deleting-one-deletes')}</span> : null}
            <Button icon="close" label="" title={t('base.remove-the')} className="bs-rel-remove" onClick={async () => { try { onChanged(await shell.base.removeRelationship({ id: db.id, relation: r.id })); } catch (e) { toast(e.message, { tone: 'bad' }); } }} />
          </li>
        ))}
        {db.relationships.length ? null : <li className="bs-hint">{t('base.no-relationships')}</li>}
      </ul>
      <div className="bs-rel-form">
        <label><span>{t('base.records-in')}</span>
          <Select className="rw-select bs-rel-from-table" aria-label={t('common.table')} value={form.fromTable} onChange={(e) => setForm({ ...form, fromTable: e.target.value, fromField: '' })}>
            <option value="">{t('common.table')}</option>
            {tables.map((n) => <option key={n} value={n}>{n}</option>)}
          </Select>
          <Select className="rw-select bs-rel-from-field" aria-label={t('common.field')} value={form.fromField} onChange={(e) => setForm({ ...form, fromField: e.target.value })}>
            <option value="">{t('common.field')}</option>
            {(fieldsOf[form.fromTable] || []).filter((f) => f.type !== 'autonumber').map((f) => <option key={f.name} value={f.name}>{f.name}</option>)}
          </Select>
        </label>
        <label><span>{t('base.point-one')}</span>
          <Select className="rw-select bs-rel-to-table" aria-label={t('common.table')} value={form.toTable} onChange={(e) => setForm({ ...form, toTable: e.target.value, toField: keysOf(e.target.value)[0]?.name || '' })}>
            <option value="">{t('common.table')}</option>
            {tables.map((n) => <option key={n} value={n}>{n}</option>)}
          </Select>
          <Select className="rw-select bs-rel-to-field" aria-label={t('common.field')} value={form.toField} onChange={(e) => setForm({ ...form, toField: e.target.value })}>
            <option value="">{t('common.field')}</option>
            {keysOf(form.toTable).map((f) => <option key={f.name} value={f.name}>{f.name}</option>)}
          </Select>
        </label>
        <label className="bs-check"><input type="checkbox" checked={form.cascadeDelete} onChange={(e) => setForm({ ...form, cascadeDelete: e.target.checked })} /><span>{t('base.deleting-one-deletes')}</span></label>
        <Button primary icon="link" label={t('base.relate')} className="bs-relate" disabled={!ready} onClick={relate} />
      </div>
    </div>
  );
}

/* ── SQL ──────────────────────────────────────────────────────────── */

/* ── the window ───────────────────────────────────────────────────── */

export default function Databases({ app, shell, boot }) {
  const toast = useToast();
  const [db, setDb] = useState(null);
  const [table, setTable] = useState(null);
  const [view, setView] = useState('data'); // data | design | relations | sql | query | form | report
  // The query, form or report open: its name, or null for a new query.
  const [item, setItem] = useState(null);
  const [formTable, setFormTable] = useState(null);
  // The New Report box's source: 'table:<name>' or 'query:<name>', null while it is shut.
  const [reportSource, setReportSource] = useState(null);
  // The New Print Template box's table, null while it is shut.
  const [docTable, setDocTable] = useState(null);
  const [busy, setBusy] = useState(false);
  // The ready-made databases that need a newer Office: id -> the version wanted.
  const [needs, setNeeds] = useState({});
  // Each ready-made one's page on office.rutba.io (D-2026-10-22-03): id -> its address.
  const [pages, setPages] = useState({});
  useEffect(() => {
    shell.base.readyMade().then((r) => {
      setNeeds(Object.fromEntries([...r.templates, ...r.samples].filter((x) => x.needs).map((x) => [x.id, x.needs])));
      setPages(Object.fromEntries([...r.templates, ...r.samples].filter((x) => x.page).map((x) => [x.id, x.page])));
    }).catch(() => {});
  }, [shell]);
  const [naming, setNaming] = useState(null);
  // Protect this database, and your own password (D-2026-10-22-11).
  const [protecting, setProtecting] = useState(false);
  const [aiTools, setAiTools] = useState(false);
  const [changingPassword, setChangingPassword] = useState(false);
  // Link to a database engine, and a link's password on a computer that has not kept it (D-2026-10-22-13).
  const [linking, setLinking] = useState(false);
  const [linkPassword, setLinkPassword] = useState(null);
  // Other sources a table is read from, a copy's clashes, and who you are where a source is read (steps 4 and 5).
  const [sourcing, setSourcing] = useState(false);
  const [clashesOf, setClashesOf] = useState(null);
  const [sourceWho, setSourceWho] = useState(null);
  // A web address's key or token typed again: kept for another address, or refused.
  const [sourceKeys, setSourceKeys] = useState(null);
  // Read from these places? Asked as a database opens whose sources this computer has not allowed (A2.18), and on Refresh.
  const [asking, setAsking] = useState(false);
  // The SQL view's text, kept while the database is open when the view is left and come back to.
  const [sqlText, setSqlText] = useState(null);
  useEffect(() => { setSqlText(null); }, [db?.id]);
  useEffect(() => { setAsking(Boolean(db?.ask?.length)); }, [db?.id, db?.access?.locked]); // eslint-disable-line react-hooks/exhaustive-deps

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
    const file = await shell.dialog.save({ title: t('base.new-database'), defaultPath: `${t('common.database')}.rbase`, filters: [{ name: t('base.rutba-database-rbase'), extensions: ['rbase'] }] });
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

  const fromTemplate = useCallback(async (tpl) => {
    const file = await shell.dialog.save({ title: t('base.new-database'), defaultPath: `${tpl.name()}.rbase`, filters: [{ name: t('base.rutba-database-rbase'), extensions: ['rbase'] }] });
    if (!file) return;
    setBusy(true);
    try {
      const s = await shell.base.createFromTemplate({ file: /\.rbase$/i.test(file) ? file : `${file}.rbase`, template: tpl.id });
      show(s);
      setView('data');
      shell.app.addRecent({ path: s.file, app: 'base' }).catch(() => {});
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    } finally {
      setBusy(false);
    }
  }, [shell, show, toast]);

  const fromSample = useCallback(async (sample, withData) => {
    const name = withData ? sample.business : sample.kind();
    const file = await shell.dialog.save({ title: t('base.new-database'), defaultPath: `${name}.rbase`, filters: [{ name: t('base.rutba-database-rbase'), extensions: ['rbase'] }] });
    if (!file) return;
    setBusy(true);
    try {
      const s = await shell.base.createFromSample({ file: /\.rbase$/i.test(file) ? file : `${file}.rbase`, sample: sample.id, withData });
      show(s);
      setView('data');
      shell.app.addRecent({ path: s.file, app: 'base' }).catch(() => {});
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    } finally {
      setBusy(false);
    }
  }, [shell, show, toast]);

  const openDialog = useCallback(async () => {
    const file = await pickOpen(shell, 'base');
    if (file) openFile(file);
  }, [shell, openFile]);

  // A workbook of several sheets: which to import, every one or one ({ file, sheets, sheet }).
  const [sheetChoice, setSheetChoice] = useState(null);
  const importFile = useCallback(async (file, sheet = null) => {
    setBusy(true);
    try {
      const r = await shell.base.importFile({ id: db.id, file, sheet });
      show(r.summary);
      setTable(r.tables[0]?.table || null);
      setView('data');
      const made = r.tables.map((x) => t('base.table-count', { table: x.table, count: x.count })).join(', ');
      // How it read the file: dates written 12/04/2025, and the file's own ID numbers kept.
      const said = r.tables.flatMap((x) => [
        ...(x.dates || []).map((d) => (d.order === 'mdy'
          ? t('base.dates-field-read-month', { field: d.field })
          : t('base.dates-field-read', { field: d.field }))),
        ...(x.key ? [t('base.table-keeps-files', { table: x.table, field: x.key })] : []),
      ]);
      toast([t('base.imported-tables', { tables: made }), ...said].join(' '), { tone: 'good', ms: said.length ? 12000 : 7000 });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 8000 });
    } finally {
      setBusy(false);
    }
  }, [shell, db, show, toast]);
  const importDialog = useCallback(async () => {
    if (!db) return;
    const files = await shell.dialog.open({ title: t('common.import'), filters: [{ name: t('base.tables-workbooks-old'), extensions: IMPORTS }] });
    if (!files?.[0]) return;
    if (WORKBOOKS.test(files[0])) {
      try {
        const { sheets } = await shell.base.importFile({ id: db.id, file: files[0], peek: true });
        if (sheets.length > 1) { setSheetChoice({ file: files[0], sheets, sheet: '' }); return; }
      } catch (e) {
        toast(e.message, { tone: 'bad', ms: 8000 });
        return;
      }
    }
    importFile(files[0]);
  }, [shell, db, toast, importFile]);

  const saveAs = useCallback(async () => {
    if (!db) return;
    const file = await shell.dialog.save({ title: t('common.save-as'), defaultPath: `${db.name}.rbase`, filters: [{ name: t('base.rutba-database-rbase'), extensions: ['rbase'] }] });
    if (!file) return;
    try {
      const s = await shell.base.saveAs({ id: db.id, file: /\.rbase$/i.test(file) ? file : `${file}.rbase` });
      show(s);
      // Another program's database made a Data Files copy (A2.20): what could not come with it, said.
      if (s.copied) {
        const said = [t('base.data-files-copy')];
        if (s.copied.left.length) said.push(t('base.left-out-no', { fields: s.copied.left.join(', ') }));
        if (s.copied.skipped) said.push(tn(s.copied.skipped, 'base.count-record-did'));
        toast(said.join(' '), { tone: 'good', ms: said.length > 1 ? 12000 : 7000 });
        shell.app.addRecent({ path: s.file, app: 'base' }).catch(() => {});
      }
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  }, [shell, db, show, toast]);
  // Another program's database: Allow changes asked first (A2.20).
  const [allowing, setAllowing] = useState(false);
  const allowChanges = async () => {
    try {
      show(await shell.base.allowChanges({ id: db.id }));
      setAllowing(false);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };

  const appMenuOpen = useRef(null);
  appMenuOpen.current = openDialog;
  const appMenu = useAppMenu({
    shell,
    appKey: 'base',
    onNew: () => shell.win.create({ app: 'base' }),
    onOpen: () => appMenuOpen.current?.(),
    saveAs: () => saveAs(),
  });
  const sideMenu = useMenu();

  useEffect(() => { if (boot?.file) openFile(boot.file); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  useFileDrop(useCallback((files) => { const f = files.find((x) => OPENS.test(x)); if (f) openFile(f); }, [openFile]));
  useEffect(() => () => { if (db) shell.base.close({ id: db.id }).catch(() => {}); }, [db?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const refresh = useCallback(async () => { if (db) show(await shell.base.summary({ id: db.id })); }, [shell, db, show]);

  // Live (D-2026-10-22-13): a linked table on screen is read again from its engine every few seconds;
  // the view changes only when the engine's records did, and not while a cell is being typed in.
  const linkedOnScreen = Boolean(db && table && (view === 'data' || view === 'design') && db.tables?.some((x) => x.name === table && x.linked && !x.blocked));
  useEffect(() => {
    if (!linkedOnScreen) return undefined;
    const timer = setInterval(async () => {
      if (document.visibilityState !== 'visible') return;
      const typing = document.activeElement?.matches?.('.bs-edit, .bs-new-input') && (document.activeElement.value || document.activeElement.classList.contains('bs-edit'));
      if (typing) return;
      try { show(await shell.base.refreshLinked({ id: db.id, table })); } catch { /* the bar says why at the next read */ }
    }, LIVE_MS);
    return () => clearInterval(timer);
  }, [shell, db?.id, table, linkedOnScreen, show]); // eslint-disable-line react-hooks/exhaustive-deps

  // Sources read again on their schedules while the database is open, on screen or not.
  const hasSources = Boolean(db?.sources?.length);
  useEffect(() => {
    if (!hasSources) return undefined;
    const timer = setInterval(async () => {
      try { const s = await shell.base.refreshDue({ id: db.id }); if (s) show(s); } catch { /* the bar says why at the next read */ }
    }, 60000);
    return () => clearInterval(timer);
  }, [shell, db?.id, hasSources, show]); // eslint-disable-line react-hooks/exhaustive-deps

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

  /** A ready-made one's page online, from its card: a small icon in its corner (D-2026-10-22-03). */
  const pageLink = (url, name) => (url ? (
    <button type="button" className="bs-page-link" data-page={url} title={t('common.about-name-online', { name })} aria-label={t('common.about-name-online', { name })}
      onClick={(e) => { e.stopPropagation(); shell.shell.openExternal({ url }).catch(() => {}); }}>
      <Icon name="globe" size={14} />
    </button>
  ) : null);

  // A newer design for a database made from a template or sample (D-2026-10-22-04), asked for when it opens.
  const [update, setUpdate] = useState(null);
  const [updating, setUpdating] = useState(false);
  useEffect(() => {
    setUpdate(null);
    if (!db?.id || !db.origin) return;
    let live = true;
    shell.base.designUpdate({ id: db.id }).then((u) => { if (live && u.available) setUpdate(u); }).catch(() => {});
    return () => { live = false; };
  }, [shell, db?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  const bringUpdate = async () => {
    setUpdating(true);
    try {
      const r = await shell.base.applyDesignUpdate({ id: db.id });
      show(r.summary);
      setUpdate(null);
      const n = r.tables.length + r.fields.length + r.added.length + r.replaced.length + r.beside.length;
      toast(r.beside.length
        ? tn(n, 'base.design-date-count-change')
        : tn(n, 'base.design-date-count'), { tone: 'good', ms: 7000 });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    } finally {
      setUpdating(false);
    }
  };
  const updateBanner = update ? (
    <div className="bs-update" role="status" data-changes={update.tables.length + update.fields.length + update.added.length + update.replaced.length + update.beside.length}>
      <Icon name="refresh" size={16} />
      <span>{tn(update.tables.length + update.fields.length + update.added.length + update.replaced.length + update.beside.length, 'base.newer-design-database')}</span>
      <Button primary label={t('base.update-the')} className="bs-update-now" disabled={updating} onClick={bringUpdate} />
      <Button ghost label={t('common.not-now')} className="bs-update-later" onClick={() => setUpdate(null)} />
    </div>
  ) : null;

  // Another program's database (A2.20): read only as it opens, nothing of ours put into it; its records changed only when the person allows it.
  const foreignBanner = db?.foreign ? (
    <div className="bs-update bs-foreign" role="status" data-foreign={db.foreign.write ? 'write' : 'read'}>
      <Icon name={db.foreign.write ? 'pen' : 'lock'} size={16} />
      <span>
        {db.foreign.write
          ? t('base.database-made-another')
          : t('base.database-made-another-program')}
        {db.foreign.left?.length ? ` ${t('base.not-shown-data', { tables: db.foreign.left.join(', ') })}` : ''}
      </span>
      {db.foreign.write ? null : <Button label={t('base.allow')} className="bs-foreign-allow" onClick={() => setAllowing(true)} />}
      <Button primary label={t('base.make-data')} className="bs-foreign-copy" onClick={saveAs} />
    </div>
  ) : null;

  const title = db ? db.name : t('common.data-files');
  if (!db) {
    return (
      <AppFrame app={app} shell={shell} title={title} menu={appMenu}>
        <Content>
          <div className="bs-start">
            <Empty
              icon="base"
              title={t('common.data-files')}
              action={(
                <span className="bs-start-actions">
                  <Button primary icon="plus" label={t('base.new-database')} className="bs-new-db" onClick={newDatabase} />
                  <Button icon="open" label={t('common.open')} className="bs-open-db" onClick={openDialog} />
                  {busy ? <Spinner /> : null}
                </span>
              )}
            >
              {t('base.tables-records-keep')}
            </Empty>
            <RecentFiles shell={shell} app="base" onOpen={openFile} open={false} embedded />
            <section className="bs-templates" aria-label={t('base.ready-made-databases')}>
              <h2>{t('base.start-ready-made')}</h2>
              <div className="bs-template-grid">
                {TEMPLATES.map((tpl) => (
                  <div key={tpl.id} className="bs-card-wrap">
                  <button type="button" className={`bs-template${needs[tpl.id] ? ' needs' : ''}`} data-template={tpl.id} data-needs={needs[tpl.id] || undefined} disabled={busy} data-tip={needs[tpl.id] ? t('base.needs-rutba-office', { version: needs[tpl.id] }) : undefined} onClick={() => (needs[tpl.id] ? shell.update.check({ manual: true }).catch(() => {}) : fromTemplate(tpl))}>
                    <span className="bs-template-icon"><Icon name={tpl.icon} size={18} /></span>
                    <span className="bs-template-words"><strong>{tpl.name()}</strong><small>{tpl.line()}</small></span>
                  </button>
                  {pageLink(pages[tpl.id], tpl.name())}
                  </div>
                ))}
              </div>
            </section>
            <section className="bs-templates bs-samples" aria-label={t('base.sample-databases')}>
              <h2>{t('base.whole-business-ready')}</h2>
              <p className="bs-samples-note">{t('base.each-comes-invented')}</p>
              {SAMPLE_GROUPS.map((g) => (
              <div key={g.id} className="bs-sample-group" data-group={g.id}>
              <h3>{g.name()}</h3>
              <div className="bs-template-grid">
                {SAMPLES.filter((s) => s.group === g.id).map((s) => (
                  <div key={s.id} className={`bs-sample${needs[s.id] ? ' needs' : ''}`} data-sample={s.id} data-needs={needs[s.id] || undefined} role="group" aria-label={s.kind()}>
                    <span className="bs-template-icon"><Icon name={s.icon} size={18} /></span>
                    <span className="bs-template-words">
                      <strong>{s.kind()}</strong>
                      <small>{s.line()}</small>
                    </span>
                    {pageLink(pages[s.id], s.business)}
                    {needs[s.id] ? (
                      <span className="bs-sample-actions">
                        <button type="button" className="bs-sample-update" onClick={() => shell.update.check({ manual: true }).catch(() => {})}>{t('common.needs-rutba-office', { version: needs[s.id] })}</button>
                      </span>
                    ) : (
                      <span className="bs-sample-actions">
                        <button type="button" className="bs-sample-data" disabled={busy} onClick={() => fromSample(s, true)} data-tip={t('base.business-invented', { business: s.business })}>{t('base.with-sample-records')}</button>
                        <button type="button" className="bs-sample-empty" disabled={busy} onClick={() => fromSample(s, false)}>{t('base.empty')}</button>
                      </span>
                    )}
                  </div>
                ))}
              </div>
              </div>
              ))}
            </section>
          </div>
        </Content>
        <style>{BASE_CSS}</style>
      </AppFrame>
    );
  }

  if (db.access?.locked) {
    return (
      <AppFrame app={app} shell={shell} title={title} menu={appMenu}>
        <SignIn shell={shell} db={db} onSigned={(s) => { show(s); setView('data'); }} />
        <style>{BASE_CSS + PEOPLE_CSS}</style>
      </AppFrame>
    );
  }
  const design = canDesign(db);
  const access = db.access || { protected: false };
  const isAdmin = access.person?.role === 'Admin';

  /* The side panel's items: opened, pinned to the top, and each section folded away (D-2026-10-22-05). */
  const KINDS = {
    table: { icon: 'sheets', attr: 'data-table', on: (n) => table === n && (view === 'data' || view === 'design'), open: (n) => { setTable(n); if (view !== 'design') setView('data'); } },
    query: { icon: 'filter', attr: 'data-query', on: (n) => view === 'query' && item === n, open: (n) => { setItem(n); setView('query'); } },
    form: { icon: 'textbox', attr: 'data-form', on: (n) => view === 'form' && item === n, open: (n) => { setItem(n); setView('form'); } },
    report: { icon: 'print', attr: 'data-report', on: (n) => view === 'report' && item === n, open: (n) => { setItem(n); setView('report'); } },
    document: { icon: 'file', attr: 'data-document', on: (n) => view === 'document' && item === n, open: (n) => { setItem(n); setView('document'); } },
  };
  const pinnedItems = db.pinned || [];
  const linkedOf = (name) => db.tables.find((x) => x.name === name && x.linked) || null;
  const isPinned = (kind, name) => pinnedItems.some((p) => p.kind === kind && p.name === name);
  const pin = (kind, name, pinned) => shell.base.pin({ id: db.id, kind, name, pinned }).then(show).catch((e) => toast(e.message, { tone: 'bad' }));
  const sideItem = (kind, name, prefix = null, count = null) => {
    const k = KINDS[kind];
    const pinned = isPinned(kind, name);
    return (
      <li key={`${prefix || kind}:${name}`} className="bs-item" data-kind={kind} data-name={name}>
        <button type="button" className={`bs-table${k.on(name) ? ' on' : ''}`} {...(prefix ? { 'data-pinned-item': name } : { [k.attr]: name })} onClick={() => k.open(name)}
          onContextMenu={(e) => {
            e.preventDefault();
            const items = [pinned ? { label: t('base.unpin'), run: () => pin(kind, name, false) } : { label: t('common.pin-top'), run: () => pin(kind, name, true) }];
            if (kind === 'table' && linkedOf(name) && design) items.push({ label: t('base.unlink-name', { name }), run: () => shell.base.unlinkTable({ id: db.id, table: name }).then((s) => { show(s); toast(linkedOf(name).source ? t('base.name-unlinked-source', { name }) : t('base.name-unlinked-engine', { name })); }).catch((err) => toast(err.message, { tone: 'bad' })) });
            sideMenu.open(e, items);
          }}>
          <Icon name={k.icon} size={14} />
          <span>{name}</span>
          {kind === 'table' && linkedOf(name) ? <span className={`bs-linked-mark${linkedOf(name).linkError ? ' bad' : ''}`} data-linked-mark={name} title={linkedOf(name).linkError || (linkedOf(name).source ? t('base.read-file-web') : linkedOf(name).mode === 'copy' ? t('base.linked-database-engine-kept') : t('base.linked-database-engine'))}><Icon name={linkedOf(name).linkError ? 'info' : 'link'} size={12} /></span> : null}
          {count !== null ? <small>{count}</small> : null}
        </button>
        <button type="button" className={`bs-pin${pinned ? ' on' : ''}`} data-pin={`${kind}:${name}`} aria-pressed={pinned} title={pinned ? t('base.unpin-name', { name }) : t('base.pin-name-top', { name })} aria-label={pinned ? t('base.unpin-name', { name }) : t('base.pin-name-top', { name })} onClick={() => pin(kind, name, !pinned)}>
          <Icon name="pin" size={13} />
        </button>
      </li>
    );
  };
  const folded = new Set(db.folded || []);
  const sideSection = (key, title, add, list) => (
    <section className={`bs-side-section${folded.has(key) ? ' folded' : ''}`} data-section={key}>
      <div className="bs-side-head">
        <button type="button" className="bs-fold" aria-expanded={!folded.has(key)} onClick={() => shell.base.fold({ id: db.id, section: key, folded: !folded.has(key) }).then(show).catch(() => {})}>
          <Icon name={folded.has(key) ? 'chevronRight' : 'chevronDown'} size={12} />
          <h2>{title}</h2>
        </button>
        {add}
      </div>
      {folded.has(key) ? null : list}
    </section>
  );

  const tab = (key, label) => <button type="button" role="tab" aria-selected={view === key} className={`bs-tab${view === key ? ' on' : ''}`} data-view={key} onClick={() => setView(key)}>{label}</button>;
  return (
    <AppFrame app={app} shell={shell} title={title} menu={appMenu}
      titleAfter={db.origin?.page ? (
        <button type="button" className="rw-btn ghost bs-about-online" data-page={db.origin.page} title={t('base.about-database-online')} aria-label={t('base.about-database-online')}
          onClick={() => shell.shell.openExternal({ url: db.origin.page }).catch(() => {})}>
          <Icon name="globe" size={13} />
        </button>
      ) : null}>
      <div className="bs-app">
        <nav className="bs-side" aria-label={t('common.tables')}>
          {pinnedItems.length ? (
            <section className="bs-side-section" data-section="pinned">
              <div className="bs-side-head"><h2><Icon name="pin" size={12} /> {t('base.pinned')}</h2></div>
              <ul className="bs-tables bs-pinned">
                {pinnedItems.map((p) => sideItem(p.kind, p.name, `pinned-${p.kind}`))}
              </ul>
            </section>
          ) : null}
          {sideSection('tables', t('common.tables'), design ? <Button icon="plus" label="" title={t('base.new-table')} className="bs-new-table" onClick={() => setNaming('')} /> : null, (
            <ul className="bs-tables">
              {db.tables.map((x) => sideItem('table', x.name, null, x.count))}
              {db.tables.length || !design ? null : <li className="bs-hint">{t('base.no-tables-yet')}</li>}
            </ul>
          ))}
          {sideSection('queries', t('base.queries'), design ? <Button icon="plus" label="" title={t('base.new-query')} className="bs-new-query" onClick={() => { setItem(null); setView('query'); }} /> : null, (
            <ul className="bs-tables bs-queries">{(db.queries || []).map((n) => sideItem('query', n))}</ul>
          ))}
          {sideSection('forms', t('base.forms'), design ? <Button icon="plus" label="" title={t('base.new-form')} className="bs-new-form" disabled={!db.tables.length} onClick={() => setFormTable(table || db.tables[0]?.name || '')} /> : null, (
            <ul className="bs-tables bs-forms">{(db.forms || []).map((n) => sideItem('form', n))}</ul>
          ))}
          {sideSection('reports', t('base.reports'), design ? <Button icon="plus" label="" title={t('base.new-report')} className="bs-new-report" disabled={!db.tables.length} onClick={() => setReportSource(`table:${table || db.tables[0]?.name || ''}`)} /> : null, (
            <ul className="bs-tables bs-reports">{(db.reports || []).map((n) => sideItem('report', n))}</ul>
          ))}
          {sideSection('documents', t('base.print-templates'), design ? <Button icon="plus" label="" title={t('base.new-print-template')} className="bs-new-document" disabled={!db.tables.length} onClick={() => setDocTable(table || db.tables[0]?.name || '')} /> : null, (
            <ul className="bs-tables bs-documents">{(db.documents || []).map((d) => sideItem('document', d.name))}</ul>
          ))}
          <div className="bs-side-foot">
            {design ? <Button icon="import" label={t('common.import')} className="bs-import" onClick={importDialog} /> : null}
            {design ? <button type="button" className={`bs-table${view === 'relations' ? ' on' : ''}`} data-view="relations" onClick={() => setView('relations')}><Icon name="link" size={14} /><span>{t('base.relationships')}</span></button> : null}
            {design || db.foreign ? <button type="button" className={`bs-table${view === 'sql' ? ' on' : ''}`} data-view="sql" onClick={() => setView('sql')}><Icon name="formula" size={14} /><span>{t('base.sql')}</span></button> : null}
            {design ? <button type="button" className="bs-table bs-link-engine" onClick={() => setLinking(true)}><Icon name="link" size={14} /><span>{t('base.link-database-engine')}</span></button> : null}
            {design ? <button type="button" className="bs-table bs-link-source" onClick={() => setSourcing(true)}><Icon name="import" size={14} /><span>{t('base.link-file-web')}</span></button> : null}
            {!access.protected && design ? <button type="button" className="bs-table bs-protect" onClick={() => setProtecting(true)}><Icon name="shield" size={14} /><span>{t('base.protect-this')}</span></button> : null}
            {(design && !access.protected) || (isAdmin && !access.person?.seeing) ? <AiToolsButton shell={shell} db={db} onOpen={() => setAiTools(true)} /> : null}
            {isAdmin && !access.person?.seeing ? <button type="button" className={`bs-table bs-people-open${view === 'people' ? ' on' : ''}`} data-view="people" onClick={() => setView('people')}><Icon name="contacts" size={14} /><span>{t('base.people-and-roles')}</span></button> : null}
            {access.person ? (
              <div className="bs-me" data-me={access.person.name}>
                <Icon name="lock" size={13} />
                <span>{t('base.signed-in-as')} <strong>{access.person.name}</strong> ({access.person.role})</span>
              </div>
            ) : null}
            {access.person ? <button type="button" className="bs-table bs-my-password" onClick={() => setChangingPassword(true)}><Icon name="lock" size={14} /><span>{t('base.change-my-password')}</span></button> : null}
            {access.person ? <button type="button" className="bs-table bs-sign-out" onClick={async () => { show(await shell.base.signOut({ id: db.id })); }}><Icon name="unlock" size={14} /><span>{t('common.sign-out')}</span></button> : null}
          </div>
          {sideMenu.node}
        </nav>
        <main className="bs-main">
          {updateBanner}
          {foreignBanner}
          <SeeingBanner shell={shell} db={db} onChanged={(s) => { show(s); setView('data'); }} />
          {(view === 'data' || view === 'design') && table ? (
            <>
              <div className="bs-tabs" role="tablist">
                <h1>{table}</h1>
                {tab('data', t('base.records'))}
                {design && !linkedOf(table) ? tab('design', t('common.design')) : null}
              </div>
              <LinkedBar shell={shell} db={db} table={table} design={design} onChanged={show} onPassword={setLinkPassword} onClashes={setClashesOf} onSourceSignIn={setSourceWho} onSourceKeys={setSourceKeys} onAsk={() => setAsking(true)} toast={toast} />
              {view === 'data' || !design || linkedOf(table)
                ? <Datasheet key={`${table}-${db.steps}-${linkedOf(table)?.readAt || ""}`} shell={shell} db={db} table={table} onChanged={refresh} toast={toast} />
                : <Design shell={shell} db={db} table={table} onChanged={show} onRemoved={() => { setTable(null); refresh(); }} toast={toast} />}
            </>
          ) : view === 'people' && isAdmin ? (
            <PeopleView shell={shell} db={db} onChanged={(s) => { show(s); if (s.access?.person?.seeing || !s.access?.protected) setView('data'); }} toast={toast} />
          ) : view === 'relations' ? (
            <Relationships shell={shell} db={db} onChanged={show} toast={toast} />
          ) : view === 'sql' ? (
            <SqlView shell={shell} db={db} text={sqlText ?? (db.tables[0]?.name ? `SELECT * FROM "${db.tables[0].name}"` : 'SELECT 1')} onText={setSqlText} onSaved={(s, n) => { show(s); setItem(n); setView('query'); }} toast={toast} />
          ) : view === 'query' ? (
            <QueryDesigner key={item || 'new'} shell={shell} db={db} name={item} onSaved={(s, n) => { show(s); setItem(n); }} toast={toast} />
          ) : view === 'form' && item ? (
            <FormView key={`${item}-${db.steps}`} shell={shell} db={db} name={item} onChanged={refresh} onMakeTemplate={(tbl) => setDocTable(tbl)} toast={toast} />
          ) : view === 'document' && item ? (
            <DocumentView key={item} shell={shell} db={db} name={item} onChanged={show} onRemoved={(s) => { show(s); setItem(null); setView('data'); }} toast={toast} />
          ) : view === 'report' && item ? (
            <ReportView key={item} shell={shell} db={db} name={item} onChanged={show} onRemoved={(s) => { show(s); setItem(null); setView('data'); }} toast={toast} />
          ) : (
            <Empty icon="sheets" title={t('base.no-table-chosen')}>{t('base.make-table-import')}</Empty>
          )}
          {busy ? <div className="bs-busy"><Spinner /></div> : null}
        </main>
      </div>
      {formTable !== null ? (
        <Dialog title={t('base.new-form')} onClose={() => setFormTable(null)} actions={<><Button label={t('common.cancel')} onClick={() => setFormTable(null)} /><Button primary label={t('base.make')} className="bs-make-form" disabled={!formTable} onClick={async () => {
          try {
            const made = await shell.base.makeForm({ id: db.id, table: formTable });
            show(made.summary);
            setItem(made.layout.name);
            setView('form');
            setFormTable(null);
          } catch (e) {
            toast(e.message, { tone: 'bad', ms: 7000 });
          }
        }} /></>}>
          <label className="bs-field"><span>{t('base.form-records')}</span>
            <Select className="rw-select bs-form-table" aria-label={t('common.table')} value={formTable} onChange={(e) => setFormTable(e.target.value)}>
              {db.tables.map((x) => <option key={x.name} value={x.name}>{x.name}</option>)}
            </Select>
          </label>
          <p className="bs-hint">{t('base.label-box-each')}</p>
        </Dialog>
      ) : null}
      {docTable !== null ? (
        <Dialog title={t('base.new-print-template')} onClose={() => setDocTable(null)} actions={<><Button label={t('common.cancel')} onClick={() => setDocTable(null)} /><Button primary label={t('base.make')} className="bs-make-document" disabled={!docTable} onClick={async () => {
          try {
            const made = await shell.base.makeDocument({ id: db.id, table: docTable });
            show(made.summary);
            setItem(made.layout.name);
            setView('document');
            setDocTable(null);
          } catch (e) {
            toast(e.message, { tone: 'bad', ms: 7000 });
          }
        }} /></>}>
          <label className="bs-field"><span>{t('base.page-each-record')}</span>
            <Select className="rw-select bs-document-table" aria-label={t('common.table')} value={docTable} onChange={(e) => setDocTable(e.target.value)}>
              {db.tables.map((x) => <option key={x.name} value={x.name}>{x.name}</option>)}
            </Select>
          </label>
          <p className="bs-hint">{t('base.invoice-receipt-payslip')}</p>
        </Dialog>
      ) : null}
      {reportSource !== null ? (
        <Dialog title={t('base.new-report')} onClose={() => setReportSource(null)} actions={<><Button label={t('common.cancel')} onClick={() => setReportSource(null)} /><Button primary label={t('base.make')} className="bs-make-report" disabled={!reportSource.split(':')[1]} onClick={async () => {
          const [kind, ...rest] = reportSource.split(':');
          const source = rest.join(':');
          try {
            const made = await shell.base.makeReport({ id: db.id, [kind]: source });
            show(made.summary);
            setItem(made.layout.name);
            setView('report');
            setReportSource(null);
          } catch (e) {
            toast(e.message, { tone: 'bad', ms: 7000 });
          }
        }} /></>}>
          <label className="bs-field"><span>{t('base.a-report-of')}</span>
            <Select className="rw-select bs-report-source" aria-label={t('common.table-or-query')} value={reportSource} onChange={(e) => setReportSource(e.target.value)}>
              <optgroup label={t('common.tables')}>{db.tables.map((x) => <option key={x.name} value={`table:${x.name}`}>{x.name}</option>)}</optgroup>
              {(db.queries || []).length ? <optgroup label={t('base.queries')}>{db.queries.map((n) => <option key={n} value={`query:${n}`}>{n}</option>)}</optgroup> : null}
            </Select>
          </label>
          <p className="bs-hint">{t('base.records-page-grouped')}</p>
        </Dialog>
      ) : null}
      {sheetChoice ? (
        <Dialog title={t('base.import-a-workbook')} onClose={() => setSheetChoice(null)} actions={<><Button label={t('common.cancel')} onClick={() => setSheetChoice(null)} /><Button primary label={t('common.import')} className="bs-import-sheets" onClick={() => { const c = sheetChoice; setSheetChoice(null); importFile(c.file, c.sheet || null); }} /></>}>
          <label className="bs-field"><span>{t('common.import')}</span>
            <Select className="rw-select bs-import-sheet" aria-label={t('common.sheet')} value={sheetChoice.sheet} onChange={(e) => setSheetChoice({ ...sheetChoice, sheet: e.target.value })}>
              <option value="">{tn(sheetChoice.sheets.length, 'base.every-sheet-count')}</option>
              {sheetChoice.sheets.map((n) => <option key={n} value={n}>{t('base.the-sheet-name', { name: n })}</option>)}
            </Select>
          </label>
          <p className="bs-hint">{t('base.each-sheets-first')}</p>
        </Dialog>
      ) : null}
      {aiTools ? <AiToolsDatabaseDialog shell={shell} db={db} onClose={() => setAiTools(false)} /> : null}
      {protecting ? <ProtectDialog shell={shell} db={db} onDone={(s) => { show(s); setProtecting(false); setView('people'); }} onClose={() => setProtecting(false)} toast={toast} /> : null}
      {linking ? <LinkDialog shell={shell} db={db} onDone={(s, first) => { show(s); setLinking(false); if (first) { setTable(first); setView('data'); } }} onClose={() => setLinking(false)} toast={toast} /> : null}
      {sourcing ? <SourceDialog shell={shell} db={db} onDone={(s, made) => { show(s); setSourcing(false); if (made) { setTable(made); setView('data'); } }} onClose={() => setSourcing(false)} toast={toast} /> : null}
      {clashesOf ? <ClashesDialog shell={shell} db={db} table={clashesOf} onChanged={show} onClose={() => setClashesOf(null)} toast={toast} /> : null}
      {asking && db.ask?.length ? <SourcesQuestion shell={shell} db={db} onDone={(s) => { show(s); setAsking(false); }} onClose={() => setAsking(false)} toast={toast} /> : null}
      {sourceKeys ? <SourceKeysDialog shell={shell} db={db} table={sourceKeys} onDone={(s) => { show(s); setSourceKeys(null); }} onClose={() => setSourceKeys(null)} toast={toast} /> : null}
      {sourceWho ? <SourceSignInDialog shell={shell} db={db} table={sourceWho} onDone={(s) => { show(s); setSourceWho(null); }} onClose={() => setSourceWho(null)} toast={toast} /> : null}
      {linkPassword ? <LinkPasswordDialog shell={shell} db={db} link={linkPassword} onDone={(s) => { show(s); setLinkPassword(null); }} onClose={() => setLinkPassword(null)} toast={toast} /> : null}
      {allowing ? (
        <Dialog title={t('base.allow-changes-database')} onClose={() => setAllowing(false)} actions={<><Button label={t('common.cancel')} onClick={() => setAllowing(false)} /><Button primary label={t('base.allow')} className="bs-foreign-allow-go" onClick={allowChanges} /></>}>
          <p className="bs-hint">{t('base.program-made-database')}</p>
          <p className="bs-hint">{t('base.rutba-office-still')}</p>
        </Dialog>
      ) : null}
      {changingPassword ? <PasswordDialog shell={shell} db={db} onClose={() => setChangingPassword(false)} toast={toast} /> : null}
      {naming !== null ? (
        <Dialog title={t('base.new-table')} onClose={() => setNaming(null)} actions={<><Button label={t('common.cancel')} onClick={() => setNaming(null)} /><Button primary label={t('base.make')} className="bs-make-table" disabled={!String(naming).trim()} onClick={makeTable} /></>}>
          <label className="bs-field"><span>{t('common.name')}</span><Input className="rw-input bs-table-name" aria-label={t('common.name')} autoFocus value={naming} onChange={(e) => setNaming(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') makeTable(); }} /></label>
          <p className="bs-hint">{t('base.starts-id-numbers')}</p>
        </Dialog>
      ) : null}
      <style>{BASE_CSS + QUERY_CSS + FORM_CSS + REPORT_CSS + DOCUMENT_CSS + FILTER_CSS + PEOPLE_CSS + LINKS_CSS + SOURCES_CSS + TRUST_CSS}</style>
    </AppFrame>
  );
}
