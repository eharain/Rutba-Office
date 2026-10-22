// A Data Files report: the records of a table or a saved query laid out as a
// page, grouped by a field with each group's totals and the whole's. The bar
// above it chooses the grouping, the order and the title, each kept in the
// file as it changes; Print sends the page to the printer and Save as PDF
// writes it as it is shown, the window's own controls left out of both.

import React, { useCallback, useEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import { Button, Input, Select, t, tn } from '@rutba/office-ui';
import { canDesign } from './people.js';

const FN_LABEL = { sum: () => t('Sum'), avg: () => t('Average'), count: () => t('Count'), min: () => t('Smallest'), max: () => t('Largest') };

/** A value as the report shows it, by its field's type. */
/** A value as a page shows it; `decimals`, when a column says, fixes a number's places. */
export function shown(type, v, decimals = null) {
  if (v === null || v === undefined || v === '') return '';
  if (decimals !== null && decimals !== undefined && Number.isFinite(Number(v)) && type !== 'date' && type !== 'datetime' && type !== 'yesno') return Number(v).toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  if (type === 'currency') return Number(v).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  // A whole number is shown as Access shows one, without thousands separators: a year is 2018, not 2,018.
  if (type === 'integer') return Number(v).toLocaleString(undefined, { maximumFractionDigits: 0, useGrouping: false });
  if (type === 'number' || typeof v === 'number') return Number(v).toLocaleString(undefined, { maximumFractionDigits: 6 });
  if (type === 'yesno') return v ? t('Yes') : t('No');
  if (type === 'date') return new Date(`${String(v).slice(0, 10)}T00:00:00`).toLocaleDateString();
  if (type === 'datetime') return new Date(v).toLocaleString();
  return String(v);
}

/** The report as a page: the same for the screen, the printer and the PDF. */
function ReportPage({ data, today }) {
  const { layout, fields } = data;
  // A column may say how it is shown (an amount from a query, which has no field types of its own).
  const typeOf = { ...Object.fromEntries(fields.map((f) => [f.name, f.type])), ...Object.fromEntries(layout.columns.filter((c) => c.type).map((c) => [c.field, c.type])) };
  const totals = layout.totals || [];
  const totalRow = (values, label, cls) => (totals.length ? (
    <tr className={cls}>
      {layout.columns.map((c, i) => {
        const tt = totals.find((x) => x.field === c.field);
        const v = tt ? values[`${tt.fn}:${tt.field}`] : null;
        return (
          <td key={c.field} className={c.align === 'right' ? 'r' : ''}>
            {i === 0 && !tt ? label : null}
            {/* An average shows two places unless its column says otherwise; a count none. */}
            {tt ? <><small>{FN_LABEL[tt.fn]?.()}</small> {shown(tt.fn === 'count' ? 'number' : typeOf[tt.field], v, tt.fn === 'count' ? 0 : c.decimals ?? (tt.fn === 'avg' ? 2 : null))}</> : null}
          </td>
        );
      })}
    </tr>
  ) : null);
  return (
    <div className="bs-report-page" data-report={layout.name}>
      <header className="bs-report-head">
        <h1>{layout.title}</h1>
        <span>{today}</span>
      </header>
      <table className="bs-report-table">
        <thead>
          <tr>{layout.columns.map((c) => <th key={c.field} className={c.align === 'right' ? 'r' : ''}>{c.label}</th>)}</tr>
        </thead>
        {data.groups.map((g, gi) => (
          <tbody key={gi} className="bs-report-group" data-group={g.label ?? ''}>
            {layout.group ? (
              <tr className="bs-report-group-head"><th colSpan={layout.columns.length}>{g.label === null || g.label === undefined || g.label === '' ? t('(none)') : shown(typeOf[layout.group], g.label)}</th></tr>
            ) : null}
            {g.rows.map((r, ri) => (
              <tr key={ri}>{layout.columns.map((c) => <td key={c.field} className={c.align === 'right' ? 'r' : ''}>{shown(typeOf[c.field], r[c.field], c.decimals)}</td>)}</tr>
            ))}
            {layout.group ? totalRow(g.totals, t('Total'), 'bs-report-subtotal') : null}
          </tbody>
        ))}
        <tfoot>{totalRow(data.totals, t('Grand total'), 'bs-report-total')}</tfoot>
      </table>
      <footer className="bs-report-foot">{tn(data.count, '{count} record', '{count} records')}</footer>
    </div>
  );
}

export function ReportView({ shell, db, name, onChanged, onRemoved, toast }) {
  const [data, setData] = useState(null);
  const [missing, setMissing] = useState(false);
  const [printing, setPrinting] = useState(false);
  const [title, setTitle] = useState('');
  const today = new Date().toLocaleDateString(undefined, { day: 'numeric', month: 'long', year: 'numeric' });

  const load = useCallback(async () => {
    try {
      const d = await shell.base.reportData({ id: db.id, name });
      setData(d);
      setTitle(d.layout.title);
      setMissing(false);
    } catch (e) {
      setMissing(true);
      toast(e.message, { tone: 'bad' });
    }
  }, [shell, db.id, name, toast]);
  useEffect(() => { load(); }, [load]);

  /** A change to the layout kept in the file, then the report drawn again. */
  const change = async (patch) => {
    const layout = { ...data.layout, ...patch };
    try {
      const s = await shell.base.saveReport({ id: db.id, layout });
      onChanged?.(s);
      setData(await shell.base.reportData({ id: db.id, name }));
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  const regroup = (field) => {
    const was = data.layout.group;
    let columns = data.layout.columns.filter((c) => c.field !== field);
    // The field grouped before goes back among the columns, at the front.
    if (was && was !== field && !columns.some((c) => c.field === was)) columns = [{ field: was, label: was, align: 'left' }, ...columns];
    const totals = (data.layout.totals || []).filter((x) => x.field !== field);
    change({ group: field || null, columns, totals });
  };

  /** The page alone, outside the window's root, while the printer or the PDF takes it. */
  const withPage = async (job) => {
    setPrinting(true);
    await new Promise((r) => setTimeout(r, 80));
    try { return await job(); } finally { setPrinting(false); }
  };
  const print = () => withPage(() => shell.print.print({ silent: false })).catch((e) => toast(e.message, { tone: 'bad' }));
  const savePdf = async () => {
    const file = await shell.dialog.save({ title: t('Save as PDF'), defaultPath: `${data.layout.title || name}.pdf`, filters: [{ name: t('PDF (.pdf)'), extensions: ['pdf'] }] });
    if (!file) return;
    try {
      const { bytes } = await withPage(() => shell.print.toPDF({ pageSize: 'A4', margins: { marginType: 'none' } }));
      await shell.fs.write({ path: file, bytes });
      toast(t('Saved as {name}', { name: file.split(/[\\/]/).pop() }), { tone: 'good' });
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };
  const remove = async () => {
    try {
      const s = await shell.base.removeReport({ id: db.id, name });
      onRemoved?.(s);
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
    }
  };

  if (missing) return <div className="bs-form-empty">{t('This report is not in the database.')}</div>;
  if (!data) return null;
  const { layout, fields } = data;
  return (
    <div className="bs-report-run">
      <div className="bs-report-bar" role="toolbar" aria-label={t('Report')}>
        <label className="bs-report-opt"><span>{t('Title')}</span>
          <Input className="rw-input bs-report-title" aria-label={t('Title')} value={title} onChange={(e) => setTitle(e.target.value)}
            onBlur={() => { if (title.trim() && title !== layout.title) change({ title: title.trim() }); }}
            onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }} />
        </label>
        <label className="bs-report-opt"><span>{t('Group by')}</span>
          <Select className="rw-select bs-report-group-by" aria-label={t('Group by')} value={layout.group || ''} onChange={(e) => regroup(e.target.value)}>
            <option value="">{t('(no grouping)')}</option>
            {fields.map((f) => <option key={f.name} value={f.name}>{f.name}</option>)}
          </Select>
        </label>
        <label className="bs-report-opt"><span>{t('Sort by')}</span>
          <Select className="rw-select bs-report-sort" aria-label={t('Sort by')} value={layout.sort || ''} onChange={(e) => change({ sort: e.target.value || null })}>
            <option value="">{t('(not sorted)')}</option>
            {fields.map((f) => <option key={f.name} value={f.name}>{f.name}</option>)}
          </Select>
        </label>
        <label className="bs-check bs-report-desc-opt"><input type="checkbox" className="bs-report-desc" checked={Boolean(layout.desc)} disabled={!layout.sort} onChange={(e) => change({ desc: e.target.checked })} /><span>{t('Largest first')}</span></label>
        <span className="grow" />
        <Button icon="print" label={t('Print')} className="bs-report-print" onClick={print} />
        <Button icon="pdf" label={t('Save as PDF')} className="bs-report-pdf" onClick={savePdf} />
        {canDesign(db) ? <Button icon="trash" label="" title={t('Delete this report')} className="bs-report-delete" onClick={remove} /> : null}
      </div>
      <div className="bs-report-scroll">
        <ReportPage data={data} today={today} />
      </div>
      {printing ? createPortal(<div className="bs-report-print-root"><ReportPage data={data} today={today} /></div>, document.body) : null}
    </div>
  );
}

export const REPORT_CSS = `
.bs-report-run { flex: 1; min-height: 0; display: flex; flex-direction: column; }
.bs-report-bar { display: flex; align-items: center; flex-wrap: wrap; gap: 10px; padding: 8px 16px; border-bottom: 1px solid var(--line-soft); background: var(--window); }
.bs-report-opt { display: inline-flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--ink-2); }
.bs-report-title { width: 200px; }
.bs-report-scroll { flex: 1; min-height: 0; overflow: auto; padding: 24px; background: var(--sunken); }
.bs-report-page { box-sizing: border-box; width: 210mm; min-height: 297mm; margin: 0 auto; padding: 18mm 16mm; background: #fff; color: #1d1d1f; box-shadow: var(--shadow-2, 0 2px 12px rgba(0,0,0,.18)); font-size: 12.5px; }
.bs-report-head { display: flex; align-items: baseline; justify-content: space-between; gap: 16px; border-bottom: 2px solid #1d1d1f; padding-bottom: 6px; margin-bottom: 12px; }
.bs-report-head h1 { margin: 0; font-size: 20px; font-weight: 650; }
.bs-report-head span { color: #6e6e73; font-size: 11.5px; }
.bs-report-table { width: 100%; border-collapse: collapse; }
.bs-report-table th, .bs-report-table td { text-align: start; padding: 4px 6px; vertical-align: top; }
.bs-report-table .r { text-align: end; font-variant-numeric: tabular-nums; }
.bs-report-table thead th { font-size: 11.5px; color: #6e6e73; font-weight: 600; border-bottom: 1px solid #c7c7cc; }
.bs-report-table tbody tr td { border-bottom: 1px solid #ececf0; }
.bs-report-group { break-inside: avoid-page; }
.bs-report-group-head th { padding-top: 12px; font-size: 13.5px; font-weight: 650; border-bottom: 1px solid #c7c7cc; }
.bs-report-subtotal td { font-weight: 600; border-bottom: none !important; padding-bottom: 8px; }
.bs-report-total td { font-weight: 700; border-top: 2px solid #1d1d1f; padding-top: 6px; }
.bs-report-table small { color: #6e6e73; font-weight: 400; font-size: 10.5px; }
.bs-report-foot { margin-top: 14px; color: #6e6e73; font-size: 11.5px; }
.bs-report-print-root { display: none; }
@media print {
  html, body { background: #fff !important; }
  #root { display: none !important; }
  .bs-report-print-root { display: block !important; }
  .bs-report-print-root .bs-report-page { width: auto; min-height: 0; padding: 14mm 14mm; margin: 0; box-shadow: none; }
  @page { size: A4; margin: 0; }
}
`;
