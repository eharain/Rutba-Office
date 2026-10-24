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
import { QueryDesigner, QUERY_CSS } from './base/queries.js';
import { FormView, FORM_CSS } from './base/forms.js';
import { ReportView, REPORT_CSS } from './base/reports.js';
import { RecentFiles } from '../start-page.js';
import { DocumentView, DOCUMENT_CSS } from './base/documents.js';
import { useGridFilter, FilterBar, ColumnFilter, selectionItems, FILTER_CSS } from './base/filters.js';
import { SignIn, ProtectDialog, PasswordDialog, SeeingBanner, PeopleView, PEOPLE_CSS, canDesign, may } from './base/people.js';
import { LinkDialog, LinkedBar, LinkPasswordDialog, ClashesDialog, LINKS_CSS } from './base/links.js';
import { SourceDialog, SourceSignInDialog, SOURCES_CSS } from './base/sources.js';

/** The ready-made databases the start screen offers, in the engine's order (packages/database/src/templates.js). */
const TEMPLATES = [
  { id: 'contacts', icon: 'contacts', name: () => t('Contacts'), line: () => t('People and how to reach them, with a phone list to print.') },
  { id: 'inventory', icon: 'archive', name: () => t('Inventory'), line: () => t('Products, their suppliers and stock movements, and what is running low.') },
  { id: 'membership', icon: 'star', name: () => t('Membership'), line: () => t('Members, their payments and the renewals coming due.') },
  { id: 'crm', icon: 'chartColumn', name: () => t('Customers and sales'), line: () => t('Companies, their people, open opportunities and each call or meeting.') },
  { id: 'library', icon: 'list', name: () => t('Lending library'), line: () => t('Books, borrowers and loans, and what is out and when it is due back.') },
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
  { id: 'shops', name: () => t('Shops and food') },
  { id: 'trade', name: () => t('Trade, making and building') },
  { id: 'services', name: () => t('Services') },
  { id: 'care', name: () => t('Health, learning and community') },
  { id: 'property', name: () => t('Property, premises and staff') },
];
const SAMPLES = [
  { id: 'shop', group: 'shops', icon: 'scanReceipt', business: 'Larkspur Corner Store', kind: () => t('Shop'), line: () => t('Products, the till, stock, deliveries, loyalty customers and running costs.') },
  { id: 'food', group: 'shops', icon: 'archive', business: 'Bramblegate Wine and Deli', kind: () => t('Food and beverages'), line: () => t('Wines, beers, spirits and fine food, stock by batch and expiry date, markdowns and age checks.') },
  { id: 'cafe', group: 'shops', icon: 'sun', business: 'Little Fig Cafe', kind: () => t('Cafe or restaurant'), line: () => t('Menu costing, orders, ingredients, suppliers, waste and the staff rota.') },
  { id: 'takeaway', group: 'shops', icon: 'send', business: 'Saffron Lane Pizza and Curry', kind: () => t('Takeaway'), line: () => t('Phone and online orders, collection and delivery, drivers, kitchen tickets and receipts.') },
  { id: 'florist', group: 'shops', icon: 'star', business: 'Hazel Row Flowers', kind: () => t('Florist'), line: () => t('Orders with card messages, deliveries, weddings and funerals, stems and waste.') },
  { id: 'pharmacy', group: 'shops', icon: 'heal', business: 'Tansy Row Pharmacy', kind: () => t('Pharmacy counter'), line: () => t('Counter stock by batch and expiry date, collections, deliveries to the housebound and the till.') },
  { id: 'trading', group: 'trade', icon: 'globe', business: 'Northbank Trading', kind: () => t('Wholesale and trading'), line: () => t('Customers, suppliers, products, orders, purchase orders, stock and invoices.') },
  { id: 'workshop', group: 'trade', icon: 'layers', business: 'Granite and Grain Woodworks', kind: () => t('Workshop or maker'), line: () => t('Materials, bills of materials, work orders, finished stock and sales.') },
  { id: 'builder', group: 'trade', icon: 'grid', business: 'Kilnside Builders', kind: () => t('Builder or tradesperson'), line: () => t('Enquiries, quotes, jobs in stages, variations, stage invoices and retentions.') },
  { id: 'cars', group: 'trade', icon: 'flag', business: 'Copperbeech Motors', kind: () => t('Car dealer'), line: () => t('Vehicles in stock with their costs, part-exchanges, sales, finance and warranties.') },
  { id: 'salon', group: 'services', icon: 'scissors', business: 'Marigold Hair and Beauty', kind: () => t('Salon'), line: () => t('Clients, bookings, stylists, services, retail stock and commission.') },
  { id: 'barber', group: 'services', icon: 'brush', business: 'Copper Pole Barbers', kind: () => t('Barber shop'), line: () => t('Walk-ins and bookings, chair rental, takings by barber and loyalty cards.') },
  { id: 'repairs', group: 'services', icon: 'phone', business: 'Ironbridge Repairs', kind: () => t('Repair shop'), line: () => t('Repair jobs from booking in to collection, parts, quotes and warranty.') },
  { id: 'garage', group: 'services', icon: 'settings', business: 'Quarrybank Motors', kind: () => t('Garage and MOT centre'), line: () => t('Vehicles, job cards, parts and labour, MOT tests and reminders.') },
  { id: 'cleaning', group: 'services', icon: 'bucket', business: 'Fernhollow Cleaning', kind: () => t('Cleaning company'), line: () => t('Clients and sites, contracts, the cleaners\' rota, hours, pay and monthly invoices.') },
  { id: 'agency', group: 'services', icon: 'clock', business: 'Bramble Studio', kind: () => t('Consultancy or agency'), line: () => t('Clients, projects, time sheets, expenses and invoices.') },
  { id: 'gym', group: 'services', icon: 'sliders', business: 'Tannery Yard Fitness', kind: () => t('Gym or fitness studio'), line: () => t('Members and plans, payments, classes and bookings, personal training and visits.') },
  { id: 'clinic', group: 'care', icon: 'calendar', business: 'Cedar Lane Clinic', kind: () => t('Clinic'), line: () => t('Patients, appointments, treatments, invoices and recalls.') },
  { id: 'school', group: 'care', icon: 'pen', business: 'Lantern Learning Centre', kind: () => t('School or training centre'), line: () => t('Students, courses, classes, attendance, results and fees.') },
  { id: 'charity', group: 'care', icon: 'heart', business: 'Rowanmere Community Trust', kind: () => t('Charity or club'), line: () => t('Members, renewals, donations, campaigns, events and volunteers.') },
  { id: 'rentals', group: 'property', icon: 'home', business: 'Hollybrook Lettings', kind: () => t('Property rentals'), line: () => t('Properties, units, tenants, leases, rent due, repairs and inspections.') },
  { id: 'agent', group: 'property', icon: 'folderOpen', business: 'Fernleigh Estates', kind: () => t('Estate and letting agent'), line: () => t('Landlords and sellers, viewings and offers, tenancies, rent collected and statements.') },
  { id: 'guesthouse', group: 'property', icon: 'moon', business: 'Gullhaven Guest House', kind: () => t('Guest house or B&B'), line: () => t('Rooms, bookings and deposits, guests, their bills and the booking sites.') },
  { id: 'visitors', group: 'property', icon: 'signature', business: 'Copperfield House', kind: () => t('Visitors at the front desk'), line: () => t('Who is in the building, passes, hosts, contractors\' inductions and deliveries.') },
  { id: 'payroll', group: 'property', icon: 'financial', business: 'Wrenmoor Garden Centre', kind: () => t('Staff, timesheets and payroll'), line: () => t('Employees, holidays and absence, timesheets, pay runs and payslips.') },
];

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
/** How often a linked table on screen is read again from its engine. */
const LIVE_MS = 15000;
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
  // Searched and filtered in the database (D-2026-10-22-07); the menu offers filter by selection.
  const grid = useGridFilter();
  const [filtering, setFiltering] = useState(null);
  const menu = useMenu();

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
    } catch (e) {
      pages.current.delete(page);
      toast(e.message, { tone: 'bad' });
    }
  }, [shell, db.id, table, order, grid.search, grid.filters, toast]);

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
    return <span className="bs-text">{shown(f, row[f.name])}</span>;
  };

  return (
    <div className="bs-sheet">
      <div className="bs-tools">
        {mayDelete ? <Button icon="trash" label={t('Delete')} disabled={!picked.size} className="bs-delete" onClick={removePicked} /> : null}
        <span className="bs-count">{grid.active ? tn(count, '{count} record found', '{count} records found') : tn(count, '{count} record', '{count} records')}</span>
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
              <button type="button" className={`bs-th-filter${grid.filters.some((x) => x.column === f.name) ? ' on' : ''}`} data-filter-field={f.name} title={t('Filter by {column}', { column: f.name })} aria-label={t('Filter by {column}', { column: f.name })} onClick={() => setFiltering({ name: f.name, type: f.type })}>
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
                  {row ? <input type="checkbox" aria-label={t('Pick')} checked={picked.has(row._rowid)} onChange={(e) => setPicked((p) => { const n = new Set(p); if (e.target.checked) n.add(row._rowid); else n.delete(row._rowid); return n; })} /> : null}
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
              {f.type === 'autonumber' ? <span className="bs-auto">{t('(new)')}</span>
                : f.type === 'yesno' ? <input type="checkbox" aria-label={f.name} checked={Boolean(fresh[f.name])} onChange={(e) => setFresh({ ...fresh, [f.name]: e.target.checked })} />
                  : <input className="bs-new-input" data-field={f.name} aria-label={f.name} placeholder={TYPE_NAMES()[f.type]} disabled={Boolean(f.readOnly)} value={fresh[f.name] ?? ''} onChange={(e) => setFresh({ ...fresh, [f.name]: e.target.value })} onKeyDown={(e) => { if (e.key === 'Enter') addRecord(); }} />}
            </span>
          ))}
        </div> : null}
      </div>
      {editable.length || !canDesign(db) ? null : <p className="bs-hint">{t('This table has no fields to fill in yet. Add some in its design.')}</p>}
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
  const [changingPassword, setChangingPassword] = useState(false);
  // Link to a database engine, and a link's password on a computer that has not kept it (D-2026-10-22-13).
  const [linking, setLinking] = useState(false);
  const [linkPassword, setLinkPassword] = useState(null);
  // Other sources a table is read from, a copy's clashes, and who you are where a source is read (steps 4 and 5).
  const [sourcing, setSourcing] = useState(false);
  const [clashesOf, setClashesOf] = useState(null);
  const [sourceWho, setSourceWho] = useState(null);

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

  const fromTemplate = useCallback(async (tpl) => {
    const file = await shell.dialog.save({ title: t('New Database'), defaultPath: `${tpl.name()}.rbase`, filters: [{ name: t('Rutba Database (.rbase)'), extensions: ['rbase'] }] });
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
    const file = await shell.dialog.save({ title: t('New Database'), defaultPath: `${name}.rbase`, filters: [{ name: t('Rutba Database (.rbase)'), extensions: ['rbase'] }] });
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
  const sideMenu = useMenu();

  useEffect(() => { if (boot?.file) openFile(boot.file); }, []); // eslint-disable-line react-hooks/exhaustive-deps
  useFileDrop(useCallback((files) => { const f = files.find((x) => OPENS.test(x)); if (f) openFile(f); }, [openFile]));
  useEffect(() => () => { if (db) shell.base.close({ id: db.id }).catch(() => {}); }, [db?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const refresh = useCallback(async () => { if (db) show(await shell.base.summary({ id: db.id })); }, [shell, db, show]);

  // Live (D-2026-10-22-13): a linked table on screen is read again from its engine every few seconds;
  // the view changes only when the engine's records did, and not while a cell is being typed in.
  const linkedOnScreen = Boolean(db && table && (view === 'data' || view === 'design') && db.tables?.some((x) => x.name === table && x.linked));
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
    <button type="button" className="bs-page-link" data-page={url} title={t('About {name} online', { name })} aria-label={t('About {name} online', { name })}
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
        ? tn(n, 'The design is up to date: {count} change. Where you had changed something, your version is kept and the newer one is beside it, named (updated).', 'The design is up to date: {count} changes. Where you had changed something, your version is kept and the newer one is beside it, named (updated).')
        : tn(n, 'The design is up to date: {count} change, your records as they were.', 'The design is up to date: {count} changes, your records as they were.'), { tone: 'good', ms: 7000 });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    } finally {
      setUpdating(false);
    }
  };
  const updateBanner = update ? (
    <div className="bs-update" role="status" data-changes={update.tables.length + update.fields.length + update.added.length + update.replaced.length + update.beside.length}>
      <Icon name="refresh" size={16} />
      <span>{tn(update.tables.length + update.fields.length + update.added.length + update.replaced.length + update.beside.length, 'A newer design of this database is ready: {count} change to its tables, queries, forms, reports or print templates. Your records stay as they are, and anything you changed keeps your version.', 'A newer design of this database is ready: {count} changes to its tables, queries, forms, reports or print templates. Your records stay as they are, and anything you changed keeps your version.')}</span>
      <Button primary label={t('Update the design')} className="bs-update-now" disabled={updating} onClick={bringUpdate} />
      <Button ghost label={t('Not now')} className="bs-update-later" onClick={() => setUpdate(null)} />
    </div>
  ) : null;

  const title = db ? db.name : t('Data Files');
  if (!db) {
    return (
      <AppFrame app={app} shell={shell} title={title} menu={appMenu}>
        <Content>
          <div className="bs-start">
            <Empty
              icon="base"
              title={t('Data Files')}
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
            <RecentFiles shell={shell} app="base" onOpen={openFile} open={false} embedded />
            <section className="bs-templates" aria-label={t('Ready-made databases')}>
              <h2>{t('Or start from a ready-made one')}</h2>
              <div className="bs-template-grid">
                {TEMPLATES.map((tpl) => (
                  <div key={tpl.id} className="bs-card-wrap">
                  <button type="button" className={`bs-template${needs[tpl.id] ? ' needs' : ''}`} data-template={tpl.id} data-needs={needs[tpl.id] || undefined} disabled={busy} data-tip={needs[tpl.id] ? t('Needs Rutba Office {version} or newer. Check for updates.', { version: needs[tpl.id] }) : undefined} onClick={() => (needs[tpl.id] ? shell.update.check({ manual: true }).catch(() => {}) : fromTemplate(tpl))}>
                    <span className="bs-template-icon"><Icon name={tpl.icon} size={18} /></span>
                    <span className="bs-template-words"><strong>{tpl.name()}</strong><small>{tpl.line()}</small></span>
                  </button>
                  {pageLink(pages[tpl.id], tpl.name())}
                  </div>
                ))}
              </div>
            </section>
            <section className="bs-templates bs-samples" aria-label={t('Sample databases')}>
              <h2>{t('Or a whole business, ready to use')}</h2>
              <p className="bs-samples-note">{t('Each comes with invented records to learn from, or empty for your own.')}</p>
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
                        <button type="button" className="bs-sample-update" onClick={() => shell.update.check({ manual: true }).catch(() => {})}>{t('Needs Rutba Office {version}', { version: needs[s.id] })}</button>
                      </span>
                    ) : (
                      <span className="bs-sample-actions">
                        <button type="button" className="bs-sample-data" disabled={busy} onClick={() => fromSample(s, true)} data-tip={t('{business}, with its invented records', { business: s.business })}>{t('With sample records')}</button>
                        <button type="button" className="bs-sample-empty" disabled={busy} onClick={() => fromSample(s, false)}>{t('Empty')}</button>
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
            const items = [pinned ? { label: t('Unpin'), run: () => pin(kind, name, false) } : { label: t('Pin to the top'), run: () => pin(kind, name, true) }];
            if (kind === 'table' && linkedOf(name) && design) items.push({ label: t('Unlink {name}', { name }), run: () => shell.base.unlinkTable({ id: db.id, table: name }).then((s) => { show(s); toast(linkedOf(name).source ? t('{name} is unlinked. Its source keeps its records.', { name }) : t('{name} is unlinked. The engine keeps its records.', { name })); }).catch((err) => toast(err.message, { tone: 'bad' })) });
            sideMenu.open(e, items);
          }}>
          <Icon name={k.icon} size={14} />
          <span>{name}</span>
          {kind === 'table' && linkedOf(name) ? <span className={`bs-linked-mark${linkedOf(name).linkError ? ' bad' : ''}`} data-linked-mark={name} title={linkedOf(name).linkError || (linkedOf(name).source ? t('Read from a file or web address') : linkedOf(name).mode === 'copy' ? t('Linked to a database engine, kept as a copy') : t('Linked to a database engine'))}><Icon name={linkedOf(name).linkError ? 'info' : 'link'} size={12} /></span> : null}
          {count !== null ? <small>{count}</small> : null}
        </button>
        <button type="button" className={`bs-pin${pinned ? ' on' : ''}`} data-pin={`${kind}:${name}`} aria-pressed={pinned} title={pinned ? t('Unpin {name}', { name }) : t('Pin {name} to the top', { name })} aria-label={pinned ? t('Unpin {name}', { name }) : t('Pin {name} to the top', { name })} onClick={() => pin(kind, name, !pinned)}>
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
        <button type="button" className="rw-btn ghost bs-about-online" data-page={db.origin.page} title={t('About this database online')} aria-label={t('About this database online')}
          onClick={() => shell.shell.openExternal({ url: db.origin.page }).catch(() => {})}>
          <Icon name="globe" size={13} />
        </button>
      ) : null}>
      <div className="bs-app">
        <nav className="bs-side" aria-label={t('Tables')}>
          {pinnedItems.length ? (
            <section className="bs-side-section" data-section="pinned">
              <div className="bs-side-head"><h2><Icon name="pin" size={12} /> {t('Pinned')}</h2></div>
              <ul className="bs-tables bs-pinned">
                {pinnedItems.map((p) => sideItem(p.kind, p.name, `pinned-${p.kind}`))}
              </ul>
            </section>
          ) : null}
          {sideSection('tables', t('Tables'), design ? <Button icon="plus" label="" title={t('New Table')} className="bs-new-table" onClick={() => setNaming('')} /> : null, (
            <ul className="bs-tables">
              {db.tables.map((x) => sideItem('table', x.name, null, x.count))}
              {db.tables.length || !design ? null : <li className="bs-hint">{t('No tables yet: make one, or import one.')}</li>}
            </ul>
          ))}
          {sideSection('queries', t('Queries'), design ? <Button icon="plus" label="" title={t('New Query')} className="bs-new-query" onClick={() => { setItem(null); setView('query'); }} /> : null, (
            <ul className="bs-tables bs-queries">{(db.queries || []).map((n) => sideItem('query', n))}</ul>
          ))}
          {sideSection('forms', t('Forms'), design ? <Button icon="plus" label="" title={t('New Form')} className="bs-new-form" disabled={!db.tables.length} onClick={() => setFormTable(table || db.tables[0]?.name || '')} /> : null, (
            <ul className="bs-tables bs-forms">{(db.forms || []).map((n) => sideItem('form', n))}</ul>
          ))}
          {sideSection('reports', t('Reports'), design ? <Button icon="plus" label="" title={t('New Report')} className="bs-new-report" disabled={!db.tables.length} onClick={() => setReportSource(`table:${table || db.tables[0]?.name || ''}`)} /> : null, (
            <ul className="bs-tables bs-reports">{(db.reports || []).map((n) => sideItem('report', n))}</ul>
          ))}
          {sideSection('documents', t('Print templates'), design ? <Button icon="plus" label="" title={t('New Print Template')} className="bs-new-document" disabled={!db.tables.length} onClick={() => setDocTable(table || db.tables[0]?.name || '')} /> : null, (
            <ul className="bs-tables bs-documents">{(db.documents || []).map((d) => sideItem('document', d.name))}</ul>
          ))}
          <div className="bs-side-foot">
            {design ? <Button icon="import" label={t('Import')} className="bs-import" onClick={importDialog} /> : null}
            {design ? <button type="button" className={`bs-table${view === 'relations' ? ' on' : ''}`} data-view="relations" onClick={() => setView('relations')}><Icon name="link" size={14} /><span>{t('Relationships')}</span></button> : null}
            {design ? <button type="button" className={`bs-table${view === 'sql' ? ' on' : ''}`} data-view="sql" onClick={() => setView('sql')}><Icon name="formula" size={14} /><span>{t('SQL')}</span></button> : null}
            {design ? <button type="button" className="bs-table bs-link-engine" onClick={() => setLinking(true)}><Icon name="link" size={14} /><span>{t('Link to a database engine')}</span></button> : null}
            {design ? <button type="button" className="bs-table bs-link-source" onClick={() => setSourcing(true)}><Icon name="import" size={14} /><span>{t('Link to a file or web address')}</span></button> : null}
            {!access.protected && design ? <button type="button" className="bs-table bs-protect" onClick={() => setProtecting(true)}><Icon name="shield" size={14} /><span>{t('Protect this database')}</span></button> : null}
            {isAdmin && !access.person?.seeing ? <button type="button" className={`bs-table bs-people-open${view === 'people' ? ' on' : ''}`} data-view="people" onClick={() => setView('people')}><Icon name="contacts" size={14} /><span>{t('People and roles')}</span></button> : null}
            {access.person ? (
              <div className="bs-me" data-me={access.person.name}>
                <Icon name="lock" size={13} />
                <span>{t('Signed in as')} <strong>{access.person.name}</strong> ({access.person.role})</span>
              </div>
            ) : null}
            {access.person ? <button type="button" className="bs-table bs-my-password" onClick={() => setChangingPassword(true)}><Icon name="lock" size={14} /><span>{t('Change my password')}</span></button> : null}
            {access.person ? <button type="button" className="bs-table bs-sign-out" onClick={async () => { show(await shell.base.signOut({ id: db.id })); }}><Icon name="unlock" size={14} /><span>{t('Sign out')}</span></button> : null}
          </div>
          {sideMenu.node}
        </nav>
        <main className="bs-main">
          {updateBanner}
          <SeeingBanner shell={shell} db={db} onChanged={(s) => { show(s); setView('data'); }} />
          {(view === 'data' || view === 'design') && table ? (
            <>
              <div className="bs-tabs" role="tablist">
                <h1>{table}</h1>
                {tab('data', t('Records'))}
                {design && !linkedOf(table) ? tab('design', t('Design')) : null}
              </div>
              <LinkedBar shell={shell} db={db} table={table} design={design} onChanged={show} onPassword={setLinkPassword} onClashes={setClashesOf} onSourceSignIn={setSourceWho} toast={toast} />
              {view === 'data' || !design || linkedOf(table)
                ? <Datasheet key={`${table}-${db.steps}-${linkedOf(table)?.readAt || ""}`} shell={shell} db={db} table={table} onChanged={refresh} toast={toast} />
                : <Design shell={shell} db={db} table={table} onChanged={show} onRemoved={() => { setTable(null); refresh(); }} toast={toast} />}
            </>
          ) : view === 'people' && isAdmin ? (
            <PeopleView shell={shell} db={db} onChanged={(s) => { show(s); if (s.access?.person?.seeing || !s.access?.protected) setView('data'); }} toast={toast} />
          ) : view === 'relations' ? (
            <Relationships shell={shell} db={db} onChanged={show} toast={toast} />
          ) : view === 'sql' ? (
            <Sql shell={shell} db={db} />
          ) : view === 'query' ? (
            <QueryDesigner key={item || 'new'} shell={shell} db={db} name={item} onSaved={(s, n) => { show(s); setItem(n); }} toast={toast} />
          ) : view === 'form' && item ? (
            <FormView key={`${item}-${db.steps}`} shell={shell} db={db} name={item} onChanged={refresh} toast={toast} />
          ) : view === 'document' && item ? (
            <DocumentView key={item} shell={shell} db={db} name={item} onChanged={show} onRemoved={(s) => { show(s); setItem(null); setView('data'); }} toast={toast} />
          ) : view === 'report' && item ? (
            <ReportView key={item} shell={shell} db={db} name={item} onChanged={show} onRemoved={(s) => { show(s); setItem(null); setView('data'); }} toast={toast} />
          ) : (
            <Empty icon="sheets" title={t('No table chosen')}>{t('Make a table, or import one from a file.')}</Empty>
          )}
          {busy ? <div className="bs-busy"><Spinner /></div> : null}
        </main>
      </div>
      {formTable !== null ? (
        <Dialog title={t('New Form')} onClose={() => setFormTable(null)} actions={<><Button label={t('Cancel')} onClick={() => setFormTable(null)} /><Button primary label={t('Make')} className="bs-make-form" disabled={!formTable} onClick={async () => {
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
          <label className="bs-field"><span>{t('A form for the records of')}</span>
            <Select className="rw-select bs-form-table" aria-label={t('Table')} value={formTable} onChange={(e) => setFormTable(e.target.value)}>
              {db.tables.map((x) => <option key={x.name} value={x.name}>{x.name}</option>)}
            </Select>
          </label>
          <p className="bs-hint">{t('A label and a box for each field, a list for each field that points at another table, and the records that point at this one below.')}</p>
        </Dialog>
      ) : null}
      {docTable !== null ? (
        <Dialog title={t('New Print Template')} onClose={() => setDocTable(null)} actions={<><Button label={t('Cancel')} onClick={() => setDocTable(null)} /><Button primary label={t('Make')} className="bs-make-document" disabled={!docTable} onClick={async () => {
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
          <label className="bs-field"><span>{t('A page for each record of')}</span>
            <Select className="rw-select bs-document-table" aria-label={t('Table')} value={docTable} onChange={(e) => setDocTable(e.target.value)}>
              {db.tables.map((x) => <option key={x.name} value={x.name}>{x.name}</option>)}
            </Select>
          </label>
          <p className="bs-hint">{t('An invoice, a receipt, a payslip or a letter: the record\'s fields at the top, its lines from a related table with their total, and your business\'s details as the letterhead when the database keeps them. Change any of it in Design.')}</p>
        </Dialog>
      ) : null}
      {reportSource !== null ? (
        <Dialog title={t('New Report')} onClose={() => setReportSource(null)} actions={<><Button label={t('Cancel')} onClick={() => setReportSource(null)} /><Button primary label={t('Make')} className="bs-make-report" disabled={!reportSource.split(':')[1]} onClick={async () => {
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
          <label className="bs-field"><span>{t('A report of')}</span>
            <Select className="rw-select bs-report-source" aria-label={t('Table or query')} value={reportSource} onChange={(e) => setReportSource(e.target.value)}>
              <optgroup label={t('Tables')}>{db.tables.map((x) => <option key={x.name} value={`table:${x.name}`}>{x.name}</option>)}</optgroup>
              {(db.queries || []).length ? <optgroup label={t('Queries')}>{db.queries.map((n) => <option key={n} value={`query:${n}`}>{n}</option>)}</optgroup> : null}
            </Select>
          </label>
          <p className="bs-hint">{t('Its records on a page, grouped by the first field that points at another table, with the amounts added up for each group and for the whole. Print it, or save it as a PDF.')}</p>
        </Dialog>
      ) : null}
      {protecting ? <ProtectDialog shell={shell} db={db} onDone={(s) => { show(s); setProtecting(false); setView('people'); }} onClose={() => setProtecting(false)} toast={toast} /> : null}
      {linking ? <LinkDialog shell={shell} db={db} onDone={(s, first) => { show(s); setLinking(false); if (first) { setTable(first); setView('data'); } }} onClose={() => setLinking(false)} toast={toast} /> : null}
      {sourcing ? <SourceDialog shell={shell} db={db} onDone={(s, made) => { show(s); setSourcing(false); if (made) { setTable(made); setView('data'); } }} onClose={() => setSourcing(false)} toast={toast} /> : null}
      {clashesOf ? <ClashesDialog shell={shell} db={db} table={clashesOf} onChanged={show} onClose={() => setClashesOf(null)} toast={toast} /> : null}
      {sourceWho ? <SourceSignInDialog shell={shell} db={db} table={sourceWho} onDone={(s) => { show(s); setSourceWho(null); }} onClose={() => setSourceWho(null)} toast={toast} /> : null}
      {linkPassword ? <LinkPasswordDialog shell={shell} db={db} link={linkPassword} onDone={(s) => { show(s); setLinkPassword(null); }} onClose={() => setLinkPassword(null)} toast={toast} /> : null}
      {changingPassword ? <PasswordDialog shell={shell} db={db} onClose={() => setChangingPassword(false)} toast={toast} /> : null}
      {naming !== null ? (
        <Dialog title={t('New Table')} onClose={() => setNaming(null)} actions={<><Button label={t('Cancel')} onClick={() => setNaming(null)} /><Button primary label={t('Make')} className="bs-make-table" disabled={!String(naming).trim()} onClick={makeTable} /></>}>
          <label className="bs-field"><span>{t('Name')}</span><Input className="rw-input bs-table-name" aria-label={t('Name')} autoFocus value={naming} onChange={(e) => setNaming(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') makeTable(); }} /></label>
          <p className="bs-hint">{t('It starts with an ID that numbers each record; add its fields in its design.')}</p>
        </Dialog>
      ) : null}
      <style>{BASE_CSS + QUERY_CSS + FORM_CSS + REPORT_CSS + DOCUMENT_CSS + FILTER_CSS + PEOPLE_CSS + LINKS_CSS + SOURCES_CSS}</style>
    </AppFrame>
  );
}
