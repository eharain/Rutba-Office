// Data Files' print templates in the window: an invoice, a receipt, a
// payslip or an agreement, one page for each record, shown at its paper's
// size. Print and Save as PDF take the pages alone. Design opens a panel
// beside them: the paper, the letterhead and the blocks (a heading, fields,
// the lines, totals, words, signatures, space), each changed in place, moved
// up or down, or taken out, the pages following every change. A form prints
// the record it shows through the same printer (usePrinter).

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Button, Icon, Input, Select, t, tn } from '@rutba/office-ui';
import { shown } from './reports.js';

const PAPER = () => [['a4', t('A4')], ['letter', t('Letter')], ['a5', t('A5')], ['receipt-80', t('Till roll, 80 mm')], ['receipt-58', t('Till roll, 58 mm')]];
const BLOCK_KINDS = () => [['heading', t('Heading')], ['fields', t('Fields')], ['lines', t('Lines')], ['totals', t('Totals')], ['text', t('Text')], ['signatures', t('Signatures')], ['spacer', t('Space')]];
/** Paper as print takes it: the CSS page size, and the page's width on screen. */
const PAGE = { a4: ['210mm 297mm', '210mm'], letter: ['8.5in 11in', '8.5in'], a5: ['148mm 210mm', '148mm'], 'receipt-80': ['80mm', '80mm'], 'receipt-58': ['58mm', '58mm'] };
const isRoll = (paper) => String(paper).startsWith('receipt');
// The shape of a total's expression, shown in an empty box: an example, not words to translate.
const TOTAL_SHAPE = '[Amount]';

/** A value as a document shows it; empty shows as nothing. */
const valueText = (type, v, decimals = null) => (v === null || v === undefined ? '' : type === 'autonumber' ? String(v) : shown(type, v, decimals));
const partsText = (parts) => parts.map((p, i) => (p.text !== undefined ? <React.Fragment key={i}>{p.text}</React.Fragment> : <React.Fragment key={i}>{valueText(p.type, p.value)}</React.Fragment>));

/** One document: the letterhead, then its blocks. */
function DocumentPage({ layout, letterhead, doc }) {
  const roll = isRoll(layout.paper);
  return (
    <div className={`bs-doc-page paper-${layout.paper || 'a4'}`} data-document={layout.name} data-key={doc.key}>
      {letterhead ? (
        <header className="bs-doc-letterhead">
          {letterhead.fields.map((f, i) => (f.value === null || f.value === '' ? null : <div key={f.name} className={i === 0 ? 'name' : ''}>{valueText(f.type, f.value)}</div>))}
        </header>
      ) : null}
      {doc.blocks.map((b, i) => {
        switch (b.type) {
          case 'heading':
            return <h1 key={i} className="bs-doc-heading">{partsText(b.parts)}</h1>;
          case 'fields':
            return (
              <dl key={i} className="bs-doc-fields" style={{ gridTemplateColumns: `repeat(${roll ? 1 : b.columns}, minmax(0, 1fr))` }}>
                {b.items.map((it, k) => <div key={k}><dt>{it.label}</dt><dd>{valueText(it.type, it.value)}</dd></div>)}
              </dl>
            );
          case 'lines':
            return (
              <table key={i} className="bs-doc-lines">
                <thead><tr>{b.columns.map((c) => <th key={c.name} className={c.align === 'right' ? 'r' : ''}>{c.label}</th>)}</tr></thead>
                <tbody>{b.rows.map((r, k) => (r._group !== undefined
                  // Grouped lines: a heading for each, then its lines, then its subtotal.
                  ? <tr key={k} className="group"><th colSpan={b.columns.length}>{r._group === '' ? t('(none)') : String(r._group)}</th></tr>
                  : r._subtotal
                    ? <tr key={k} className="subtotal">{b.columns.map((c, n) => <td key={c.name} className={c.align === 'right' ? 'r' : ''}>{n === 0 && !(c.name in r._subtotal) ? tn(r._count, '{count} line', '{count} lines') : c.name in r._subtotal ? valueText(c.type, r._subtotal[c.name], c.decimals) : ''}</td>)}</tr>
                    : <tr key={k}>{b.columns.map((c) => <td key={c.name} className={c.align === 'right' ? 'r' : ''}>{valueText(c.type, r[c.name], c.decimals)}</td>)}</tr>))}</tbody>
              </table>
            );
          case 'totals':
            return (
              <table key={i} className="bs-doc-totals"><tbody>
                {b.rows.map((r, k) => <tr key={k} className={r.strong ? 'strong' : ''}><th>{r.label}</th><td className="r">{valueText(r.type, r.value, r.decimals)}</td></tr>)}
              </tbody></table>
            );
          case 'text':
            return <p key={i} className={`bs-doc-text${b.small ? ' small' : ''}`}>{partsText(b.parts)}</p>;
          case 'signatures':
            return <div key={i} className="bs-doc-signatures">{b.labels.map((l) => <div key={l}><span className="line" /><span>{l}</span></div>)}</div>;
          case 'spacer':
            return <div key={i} style={{ height: b.size }} />;
          default:
            return null;
        }
      })}
    </div>
  );
}

function DocumentPages({ data }) {
  return data.documents.map((doc) => <DocumentPage key={doc.key} layout={data.layout} letterhead={data.letterhead} doc={doc} />);
}

/**
 * The printer a template view and a form share: the pages alone, outside the
 * window's root, at the paper's size, while the printer or the PDF takes them.
 */
export function usePrinter(shell, toast) {
  const [job, setJob] = useState(null);
  const run = useCallback(async (data, work) => {
    setJob(data);
    await new Promise((r) => setTimeout(r, 120));
    try { return await work(); } finally { setJob(null); }
  }, []);
  const print = useCallback((data) => run(data, () => shell.print.print({ silent: false })).catch((e) => toast(e.message, { tone: 'bad' })), [run, shell, toast]);
  const savePdf = useCallback(async (data, name) => {
    const file = await shell.dialog.save({ title: t('Save as PDF'), defaultPath: `${name}.pdf`, filters: [{ name: t('PDF (.pdf)'), extensions: ['pdf'] }] });
    if (!file) return null;
    try {
      const { bytes } = await run(data, () => shell.print.toPDF({ margins: { marginType: 'none' }, preferCSSPageSize: true }));
      await shell.fs.write({ path: file, bytes });
      toast(t('Saved as {name}', { name: file.split(/[\\/]/).pop() }), { tone: 'good' });
      return file;
    } catch (e) {
      toast(e.message, { tone: 'bad', ms: 7000 });
      return null;
    }
  }, [run, shell, toast]);
  const portal = job ? createPortal(<PrintRoot data={job} />, document.body) : null;
  return { print, savePdf, portal, busy: Boolean(job) };
}

/** The pages for the printer, with the paper's size as the page's. A till roll's height is the longest receipt's. */
function PrintRoot({ data }) {
  const ref = useRef(null);
  const [height, setHeight] = useState(null);
  const paper = data.layout.paper || 'a4';
  useEffect(() => {
    if (!isRoll(paper) || !ref.current) return;
    const px = Math.max(...[...ref.current.querySelectorAll('.bs-doc-page')].map((n) => n.scrollHeight), 200);
    setHeight(Math.ceil(px * 25.4 / 96) + 6);
  }, [paper, data]);
  const size = isRoll(paper) ? `${PAGE[paper][0]} ${height || 200}mm` : PAGE[paper][0];
  return (
    <div className="bs-doc-print-root" ref={ref}>
      <style>{`@media print { @page { size: ${size}; margin: 0; } }`}</style>
      <DocumentPages data={data} />
    </div>
  );
}

/* ── the design panel ──────────────────────────────────────────────── */

function move(list, i, by) {
  const j = i + by;
  if (j < 0 || j >= list.length) return list;
  const next = [...list];
  [next[i], next[j]] = [next[j], next[i]];
  return next;
}

/** A small list of rows the designer edits: add, change, take out. */
function RowList({ rows, onChange, blank, render, addLabel }) {
  return (
    <div className="bs-dd-rows">
      {rows.map((r, i) => (
        <div key={i} className="bs-dd-row">
          {render(r, (patch) => onChange(rows.map((x, k) => (k === i ? { ...x, ...patch } : x))))}
          <Button icon="close" label="" title={t('Take this out')} onClick={() => onChange(rows.filter((_, k) => k !== i))} />
        </div>
      ))}
      <Button icon="plus" label={addLabel} className="bs-dd-add-row" onClick={() => onChange([...rows, blank()])} />
    </div>
  );
}

function BlockEditor({ block, onChange, names, childTables, childNames }) {
  const set = (patch) => onChange({ ...block, ...patch });
  const nameList = <datalist id="bs-dd-names">{names.map((n) => <option key={n} value={n} />)}</datalist>;
  switch (block.type) {
    case 'heading':
      return <Input className="rw-input bs-dd-heading" aria-label={t('Heading')} value={block.text || ''} onChange={(e) => set({ text: e.target.value })} />;
    case 'text':
      return <textarea className="rw-input bs-dd-text" aria-label={t('Text')} rows={4} value={block.text || ''} placeholder={t('Words, with {Field} where a value goes')} onChange={(e) => set({ text: e.target.value })} />;
    case 'signatures':
      return <Input className="rw-input" aria-label={t('Who signs')} value={(block.labels || []).join(', ')} placeholder={t('Landlord, Tenant')} onChange={(e) => set({ labels: e.target.value.split(',').map((x) => x.trim()).filter(Boolean) })} />;
    case 'spacer':
      return <Input className="rw-input" type="number" aria-label={t('Height')} value={block.size || 16} onChange={(e) => set({ size: Number(e.target.value) || 16 })} />;
    case 'fields':
      return (
        <>
          {nameList}
          <label className="bs-dd-inline"><span>{t('Columns')}</span>
            <Select className="rw-select" value={block.columns || 2} onChange={(e) => set({ columns: Number(e.target.value) })}>{[1, 2, 3].map((n) => <option key={n} value={n}>{n}</option>)}</Select>
          </label>
          <RowList rows={block.items || []} onChange={(items) => set({ items })} blank={() => ({ field: names[0] || '' })} addLabel={t('Add a field')}
            render={(r, patch) => (
              <>
                <Input className="rw-input bs-dd-field" list="bs-dd-names" aria-label={t('Field')} value={r.field} onChange={(e) => patch({ field: e.target.value })} />
                <Input className="rw-input" aria-label={t('Label')} placeholder={t('Label')} value={r.label || ''} onChange={(e) => patch({ label: e.target.value || undefined })} />
              </>
            )} />
        </>
      );
    case 'lines':
      return (
        <>
          <label className="bs-dd-inline"><span>{t('Lines from')}</span>
            <Select className="rw-select bs-dd-lines-table" value={block.table || ''} onChange={(e) => set({ table: e.target.value, link: undefined, columns: [] })}>
              {childTables.map((x) => <option key={x} value={x}>{x}</option>)}
            </Select>
          </label>
          <datalist id="bs-dd-child-names">{(childNames[block.table] || []).map((n) => <option key={n} value={n} />)}</datalist>
          <RowList rows={block.columns || []} onChange={(columns) => set({ columns })} blank={() => ({ field: (childNames[block.table] || [])[0] || '' })} addLabel={t('Add a column')}
            render={(c, patch) => (c.running !== undefined ? (
              <>
                <Input className="rw-input" aria-label={t('Column name')} placeholder={t('Name')} value={c.as || ''} onChange={(e) => patch({ as: e.target.value })} />
                <Select className="rw-select bs-dd-running" aria-label={t('A running balance of')} value={c.running} onChange={(e) => patch({ running: e.target.value })}>
                  {(block.columns || []).filter((x) => x !== c && x.running === undefined).map((x) => x.as || x.field).filter(Boolean).map((n) => <option key={n} value={n}>{n}</option>)}
                </Select>
                <Input className="rw-input bs-dd-expr" aria-label={t('Starting from')} placeholder={t('Starting from, e.g. [Opening Balance]')} value={c.start || ''} onChange={(e) => patch({ start: e.target.value || undefined })} />
              </>
            ) : c.expr !== undefined ? (
              <>
                <Input className="rw-input" aria-label={t('Column name')} placeholder={t('Name')} value={c.as || ''} onChange={(e) => patch({ as: e.target.value })} />
                <Input className="rw-input bs-dd-expr" aria-label={t('Expression')} value={c.expr} onChange={(e) => patch({ expr: e.target.value })} />
              </>
            ) : (
              <>
                <Input className="rw-input bs-dd-field" list="bs-dd-child-names" aria-label={t('Field')} value={c.field || ''} onChange={(e) => patch({ field: e.target.value })} />
                <Input className="rw-input" aria-label={t('Label')} placeholder={t('Label')} value={c.label || ''} onChange={(e) => patch({ label: e.target.value || undefined })} />
              </>
            ))} />
          <Button icon="formula" label={t('Add a calculated column')} onClick={() => set({ columns: [...(block.columns || []), { expr: '[Quantity] * [Unit Price]', as: 'Amount', type: 'currency' }] })} />
          <Button icon="sum" label={t('Add a running balance')} className="bs-dd-add-running" disabled={!(block.columns || []).length} onClick={() => set({ columns: [...(block.columns || []), { running: (block.columns || []).map((x) => x.as || x.field).filter(Boolean).pop(), as: t('Balance') }] })} />
          <label className="bs-dd-inline"><span>{t('Only lines where')}</span>
            <Input className="rw-input bs-dd-expr bs-dd-lines-where" aria-label={t('Only lines where')} placeholder={t('e.g. Not [Paid]')} value={block.where || ''} onChange={(e) => set({ where: e.target.value || undefined })} />
          </label>
          <label className="bs-dd-inline"><span>{t('Group by')}</span>
            <Select className="rw-select bs-dd-group" value={block.group || ''} onChange={(e) => set({ group: e.target.value || undefined })}>
              <option value="">{t('(none)')}</option>
              {(block.columns || []).filter((x) => x.running === undefined).map((x) => x.as || x.field).filter(Boolean).map((n) => <option key={n} value={n}>{n}</option>)}
            </Select>
          </label>
        </>
      );
    case 'totals':
      return (
        <RowList rows={block.rows || []} onChange={(rows) => set({ rows })} blank={() => ({ label: t('Total'), expr: '' })} addLabel={t('Add a total')}
          render={(r, patch) => (
            <>
              <Input className="rw-input" aria-label={t('Label')} value={r.label} onChange={(e) => patch({ label: e.target.value })} />
              <Input className="rw-input bs-dd-expr" aria-label={t('Expression')} placeholder={TOTAL_SHAPE} value={r.expr} onChange={(e) => patch({ expr: e.target.value })} />
              <label className="bs-check"><input type="checkbox" checked={Boolean(r.strong)} onChange={(e) => patch({ strong: e.target.checked })} /><span>{t('Bold')}</span></label>
            </>
          )} />
      );
    default:
      return null;
  }
}

function Designer({ layout, onChange, tables, names, childTables, childNames }) {
  const blocks = layout.blocks || [];
  const setBlocks = (next) => onChange({ ...layout, blocks: next });
  const [adding, setAdding] = useState('');
  return (
    <aside className="bs-dd" aria-label={t('Design')}>
      <div className="bs-dd-top">
        <label className="bs-dd-inline"><span>{t('Paper')}</span>
          <Select className="rw-select bs-dd-paper" value={layout.paper || 'a4'} onChange={(e) => onChange({ ...layout, paper: e.target.value })}>
            {PAPER().map(([k, l]) => <option key={k} value={k}>{l}</option>)}
          </Select>
        </label>
        <label className="bs-dd-inline"><span>{t('Only for records where')}</span>
          <Input className="rw-input bs-dd-expr bs-dd-where" aria-label={t('Only for records where')} placeholder={t('e.g. [Status] = "Shipped"')} value={layout.where || ''} onChange={(e) => onChange({ ...layout, where: e.target.value || undefined })} />
        </label>
        <label className="bs-dd-inline"><span>{t('Letterhead')}</span>
          <Select className="rw-select bs-dd-letterhead" value={layout.letterhead?.table || ''} onChange={(e) => onChange({ ...layout, letterhead: e.target.value ? { table: e.target.value } : undefined })}>
            <option value="">{t('(none)')}</option>
            {tables.filter((x) => x !== layout.table).map((x) => <option key={x} value={x}>{x}</option>)}
          </Select>
        </label>
      </div>
      {blocks.map((b, i) => (
        <section key={i} className="bs-dd-block" data-block={b.type}>
          <header>
            <strong>{BLOCK_KINDS().find(([k]) => k === b.type)?.[1] || b.type}</strong>
            <span className="grow" />
            <Button icon="chevronUp" label="" title={t('Move up')} disabled={i === 0} onClick={() => setBlocks(move(blocks, i, -1))} />
            <Button icon="chevronDown" label="" title={t('Move down')} disabled={i === blocks.length - 1} onClick={() => setBlocks(move(blocks, i, 1))} />
            <Button icon="trash" label="" title={t('Take this block out')} className="bs-dd-remove" onClick={() => setBlocks(blocks.filter((_, k) => k !== i))} />
          </header>
          <BlockEditor block={b} names={names} childTables={childTables} childNames={childNames} onChange={(nb) => setBlocks(blocks.map((x, k) => (k === i ? nb : x)))} />
                  <label className="bs-dd-inline bs-dd-if-row"><span>{t('Show only when')}</span>
            <Input className="rw-input bs-dd-expr bs-dd-if" aria-label={t('Show only when')} placeholder={t('Always, or e.g. [Payment] = "Cash"')} value={b.if || ''} onChange={(e) => setBlocks(blocks.map((x, k) => (k === i ? { ...x, if: e.target.value || undefined } : x)))} />
          </label>
        </section>
      ))}
      <div className="bs-dd-add">
        <Select className="rw-select bs-dd-add-kind" aria-label={t('Add a block')} value={adding} onChange={(e) => setAdding(e.target.value)}>
          <option value="">{t('Add a block…')}</option>
          {BLOCK_KINDS().filter(([k]) => k !== 'lines' || childTables.length).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
        </Select>
        <Button icon="plus" label={t('Add')} className="bs-dd-add-block" disabled={!adding} onClick={() => {
          const fresh = { heading: { text: layout.title || '' }, fields: { columns: 2, items: [] }, lines: { table: childTables[0], columns: [] }, totals: { rows: [] }, text: { text: '' }, signatures: { labels: [t('Signed')] }, spacer: { size: 16 } }[adding];
          setBlocks([...blocks, { type: adding, ...fresh }]);
          setAdding('');
        }} />
      </div>
    </aside>
  );
}

/* ── the template's view ──────────────────────────────────────────── */

export function DocumentView({ shell, db, name, onChanged, onRemoved, toast }) {
  const [layout, setLayout] = useState(null);
  const [data, setData] = useState(null);
  const [error, setError] = useState('');
  const [which, setWhich] = useState('all'); // 'all', or one record's key
  const [design, setDesign] = useState(false);
  const [fieldsOf, setFieldsOf] = useState({});
  const printer = usePrinter(shell, toast);
  const saving = useRef(null);

  useEffect(() => {
    let live = true;
    shell.base.documents({ id: db.id }).then((all) => { if (live) setLayout(all.find((d) => d.name === name)?.layout || null); }).catch((e) => toast(e.message, { tone: 'bad' }));
    return () => { live = false; };
  }, [shell, db.id, name, toast]);

  // The fields of every table, for the designer's lists.
  useEffect(() => {
    let live = true;
    Promise.all(db.tables.map((x) => shell.base.fields({ id: db.id, table: x.name }).then((f) => [x.name, f]))).then((pairs) => { if (live) setFieldsOf(Object.fromEntries(pairs)); }).catch(() => {});
    return () => { live = false; };
  }, [shell, db.id, db.tables]);

  const draw = useCallback(async (l) => {
    try {
      setData(await shell.base.documentData({ id: db.id, layout: l, keys: which === 'all' ? null : [Number(which)], limit: 200 }));
      setError('');
    } catch (e) {
      setError(e.message);
    }
  }, [shell, db.id, which]);
  useEffect(() => { if (layout) draw(layout); }, [layout, draw]);

  /** A change from the designer: drawn at once, kept in the file a moment after the last one. */
  const change = (next) => {
    setLayout(next);
    clearTimeout(saving.current);
    saving.current = setTimeout(async () => {
      try { onChanged?.(await shell.base.saveDocument({ id: db.id, layout: next })); } catch (e) { setError(e.message); }
    }, 500);
  };
  useEffect(() => () => clearTimeout(saving.current), []);

  const names = useMemo(() => {
    if (!layout) return [];
    const own = (fieldsOf[layout.table] || []).map((f) => f.name);
    const through = db.relationships.filter((r) => r.from.table === layout.table).flatMap((r) => (fieldsOf[r.to.table] || []).filter((f) => f.type !== 'autonumber').map((f) => `${r.from.field}.${f.name}`));
    const head = layout.letterhead?.table ? (fieldsOf[layout.letterhead.table] || []).map((f) => `${layout.letterhead.table}.${f.name}`) : [];
    return [...own, ...through, ...head];
  }, [layout, fieldsOf, db.relationships]);
  const childTables = useMemo(() => (layout ? db.relationships.filter((r) => r.to.table === layout.table).map((r) => r.from.table) : []), [layout, db.relationships]);
  const childNames = useMemo(() => Object.fromEntries(childTables.map((c) => [c, (fieldsOf[c] || []).filter((f) => f.type !== 'autonumber').map((f) => f.name)])), [childTables, fieldsOf]);

  if (!layout) return <div className="bs-form-empty">{t('This print template is not in the database.')}</div>;
  const count = data?.documents.length || 0;
  return (
    <div className="bs-doc-run">
      <div className="bs-report-bar" role="toolbar" aria-label={t('Print template')}>
        <RecordPicker shell={shell} db={db} table={layout.table} value={which} onPick={setWhich} />
        <span className="bs-doc-count">{tn(count, '{count} page', '{count} pages')}</span>
        <span className="grow" />
        <Button icon="print" label={t('Print')} className="bs-doc-print" disabled={!count} onClick={() => printer.print(data)} />
        <Button icon="pdf" label={t('Save as PDF')} className="bs-doc-pdf" disabled={!count} onClick={() => printer.savePdf(data, layout.name)} />
        <Button icon="pen" label={t('Design')} className={`bs-doc-design${design ? ' on' : ''}`} aria-pressed={design} onClick={() => setDesign(!design)} />
        <Button icon="trash" label="" title={t('Delete this print template')} className="bs-doc-delete" onClick={async () => {
          try { onRemoved?.(await shell.base.removeDocument({ id: db.id, name })); } catch (e) { toast(e.message, { tone: 'bad', ms: 7000 }); }
        }} />
      </div>
      <div className="bs-doc-body">
        <div className="bs-report-scroll">
          {error ? <div className="bs-doc-error" role="alert">{error}</div> : null}
          {data ? <DocumentPages data={data} /> : null}
        </div>
        {design ? <Designer layout={layout} onChange={change} tables={db.tables.map((x) => x.name)} names={names} childTables={childTables} childNames={childNames} /> : null}
      </div>
      {printer.portal}
    </div>
  );
}

/** The record to print: every one, or one by its number. */
function RecordPicker({ shell, db, table, value, onPick }) {
  const [keys, setKeys] = useState([]);
  useEffect(() => {
    let live = true;
    shell.base.rows({ id: db.id, table, limit: 1000 }).then((p) => {
      if (!live) return;
      const key = Object.keys(p.rows[0] || {}).find((k) => k !== '_rowid') || 'ID';
      setKeys(p.rows.map((r) => r.ID ?? r[key]));
    }).catch(() => {});
    return () => { live = false; };
  }, [shell, db.id, table, db.steps]);
  return (
    <Select className="rw-select bs-doc-record" aria-label={t('Record')} value={value} onChange={(e) => onPick(e.target.value)}>
      <option value="all">{t('Every record')}</option>
      {keys.map((k) => <option key={k} value={String(k)}>{t('Record {n}', { n: k })}</option>)}
    </Select>
  );
}

export const DOCUMENT_CSS = `
.bs-doc-run { flex: 1; min-height: 0; display: flex; flex-direction: column; }
.bs-doc-body { flex: 1; min-height: 0; display: flex; }
.bs-doc-body > .bs-report-scroll { flex: 1; min-width: 0; display: flex; flex-direction: column; align-items: center; gap: 20px; }
.bs-doc-count { font-size: 12.5px; color: var(--ink-2); }
.bs-doc-design.on { background: var(--accent-soft, var(--hover)); }
.bs-doc-error { max-width: 210mm; padding: 10px 14px; border-radius: var(--r-2); background: color-mix(in srgb, var(--bad) 10%, var(--surface)); color: var(--ink); font-size: 13px; }
.bs-doc-page { box-sizing: border-box; flex: none; width: 210mm; min-height: 297mm; padding: 18mm 18mm; background: #fff; color: #1d1d1f; box-shadow: var(--shadow-2, 0 2px 12px rgba(0,0,0,.18)); font-size: 12.5px; line-height: 1.45; }
.bs-doc-page.paper-letter { width: 8.5in; min-height: 11in; }
.bs-doc-page.paper-a5 { width: 148mm; min-height: 210mm; padding: 12mm; }
.bs-doc-page.paper-receipt-80, .bs-doc-page.paper-receipt-58 { min-height: 0; padding: 5mm 4mm 8mm; font-size: 11.5px; font-family: ui-monospace, 'Cascadia Mono', Consolas, monospace; }
.bs-doc-page.paper-receipt-80 { width: 80mm; }
.bs-doc-page.paper-receipt-58 { width: 58mm; font-size: 10.5px; }
.bs-doc-letterhead { text-align: end; color: #4a4a4f; margin-bottom: 10mm; font-size: 11.5px; }
.bs-doc-letterhead .name { font-size: 16px; font-weight: 700; color: #1d1d1f; }
.paper-receipt-80 .bs-doc-letterhead, .paper-receipt-58 .bs-doc-letterhead { text-align: center; margin-bottom: 4mm; }
.bs-doc-heading { margin: 0 0 6mm; font-size: 24px; font-weight: 700; letter-spacing: .01em; }
.paper-receipt-80 .bs-doc-heading, .paper-receipt-58 .bs-doc-heading { font-size: 15px; text-align: center; margin-bottom: 3mm; }
.bs-doc-fields { display: grid; gap: 8px 24px; margin: 0 0 7mm; }
.bs-doc-fields dt { font-size: 10.5px; text-transform: uppercase; letter-spacing: .05em; color: #6e6e73; }
.bs-doc-fields dd { margin: 1px 0 0; white-space: pre-wrap; min-height: 1.45em; }
.bs-doc-lines { width: 100%; border-collapse: collapse; margin-bottom: 4mm; }
.bs-doc-lines th, .bs-doc-lines td { text-align: start; padding: 5px 6px; }
.bs-doc-lines th { font-size: 10.5px; text-transform: uppercase; letter-spacing: .05em; color: #6e6e73; border-bottom: 1.5px solid #1d1d1f; }
.bs-doc-lines td { border-bottom: 1px solid #e4e4e8; }
.bs-doc-lines tr.group th { padding-top: 10px; font-size: 12px; text-transform: none; letter-spacing: 0; color: #1d1d1f; border-bottom: 1px solid #c7c7cc; }
.bs-doc-lines tr.subtotal td { font-weight: 600; border-bottom: 1.5px solid #c7c7cc; }
.bs-doc-page .r { text-align: end; font-variant-numeric: tabular-nums; }
.paper-receipt-80 .bs-doc-lines th, .paper-receipt-58 .bs-doc-lines th { border-bottom: 1px dashed #1d1d1f; }
.paper-receipt-80 .bs-doc-lines td, .paper-receipt-58 .bs-doc-lines td { border: 0; padding: 2px 2px; }
.bs-doc-totals { margin: 0 0 6mm auto; border-collapse: collapse; min-width: 45%; }
.bs-doc-totals th { text-align: start; font-weight: 500; color: #4a4a4f; padding: 3px 10px 3px 0; }
.bs-doc-totals td { padding: 3px 6px; }
.bs-doc-totals tr.strong th, .bs-doc-totals tr.strong td { font-weight: 700; color: #1d1d1f; border-top: 1.5px solid #1d1d1f; padding-top: 6px; font-size: 14px; }
.paper-receipt-80 .bs-doc-totals, .paper-receipt-58 .bs-doc-totals { width: 100%; }
.bs-doc-text { margin: 0 0 5mm; white-space: pre-wrap; }
.bs-doc-text.small { font-size: 10.5px; color: #6e6e73; }
.bs-doc-signatures { display: flex; gap: 14mm; margin: 14mm 0 4mm; }
.bs-doc-signatures > div { flex: 1; display: flex; flex-direction: column; gap: 4px; font-size: 11px; color: #4a4a4f; }
.bs-doc-signatures .line { border-bottom: 1px solid #1d1d1f; height: 12mm; }
.bs-dd { flex: none; width: 360px; overflow: auto; border-inline-start: 1px solid var(--line-soft); background: var(--window); padding: 12px; display: flex; flex-direction: column; gap: 10px; }
.bs-dd-top { display: flex; flex-direction: column; gap: 8px; }
.bs-dd-inline { display: flex; align-items: center; gap: 8px; font-size: 12.5px; color: var(--ink-2); }
.bs-dd-inline > span { min-width: 74px; }
.bs-dd-block { border: 1px solid var(--line-soft); border-radius: var(--r-2); background: var(--surface); padding: 8px; display: flex; flex-direction: column; gap: 6px; }
.bs-dd-block header { display: flex; align-items: center; gap: 2px; font-size: 12.5px; }
.bs-dd-rows { display: flex; flex-direction: column; gap: 4px; }
.bs-dd-row { display: flex; gap: 4px; align-items: center; }
.bs-dd-row .rw-input { min-width: 0; flex: 1; }
.bs-dd-expr { font-family: var(--mono, ui-monospace, monospace); font-size: 12px; }
.bs-dd-text { width: 100%; box-sizing: border-box; resize: vertical; font: inherit; font-size: 12.5px; }
.bs-dd-add { display: flex; gap: 6px; }
.bs-doc-print-root { display: none; }
@media print {
  html, body { background: #fff !important; }
  #root { display: none !important; }
  .bs-doc-print-root { display: block !important; }
  .bs-doc-print-root .bs-doc-page { box-shadow: none; margin: 0; break-after: page; min-height: 0; }
  .bs-doc-print-root .bs-doc-page:last-child { break-after: auto; }
}
`;
