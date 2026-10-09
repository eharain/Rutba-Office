// The dialogs behind the Worksheets ribbon.
//
// Each one drives an engine capability that already existed and had no way in:
// conditional formatting, validation, goal seek, a what-if table, named ranges
// and find-and-replace. They are deliberately plain — a spreadsheet's power is
// in the grid, and a dialog's job is to get out of the way.

import React, { useEffect, useState } from 'react';
import { Button, Dialog, Field, Input, Select, Icon, Chip, Empty, t, tn, language } from '@rutba/office-ui';

/* ── number formats ──────────────────────────────────────────────────────── */

/**
 * The formats a person actually reaches for, with the codes a workbook stores.
 * Excel's own list is longer and mostly regional variants of these.
 */
export const NUMBER_FORMATS = [
  { label: t('General'), code: 'General' },
  { label: t('Number'), code: '#,##0.00' },
  { label: t('Number, no decimals'), code: '#,##0' },
  { label: t('Currency'), code: '"£"#,##0.00' },
  { label: t('Accounting'), code: '_-"£"* #,##0.00_-;-"£"* #,##0.00_-;_-"£"* "-"??_-;_-@_-' },
  { label: t('Percentage'), code: '0.00%' },
  { label: t('Percentage, no decimals'), code: '0%' },
  { label: t('Scientific'), code: '0.00E+00' },
  { label: t('Fraction'), code: '# ?/?' },
  { label: t('Short date'), code: 'dd/mm/yyyy' },
  { label: t('Long date'), code: 'dddd, d mmmm yyyy' },
  { label: t('Time'), code: 'hh:mm:ss' },
  { label: t('Date and time'), code: 'dd/mm/yyyy hh:mm' },
  { label: t('Duration'), code: '[h]:mm:ss' },
  { label: t('Text'), code: '@' },
];

/**
 * And the formats a window in Arabic or Urdu offers besides, as Excel in
 * those languages does: dates in the Hijri calendar (B2, and the calendar
 * byte of `[$-60401]` for Arabic month names) and numbers in the language's
 * own digits (`[$-2000000]` Arabic-Indic, `[$-3000000]` Urdu's).
 */
const REGIONAL_FORMATS = {
  ar: [
    { label: t('Hijri date'), code: 'B2dd/mm/yyyy' },
    { label: t('Hijri long date'), code: '[$-60401]d mmmm yyyy' },
    { label: t('Number in Arabic-Indic digits'), code: '[$-2000000]#,##0.00' },
  ],
  ur: [
    { label: t('Hijri date'), code: 'B2dd/mm/yyyy' },
    { label: t('Hijri long date'), code: '[$-60401]d mmmm yyyy' },
    { label: t('Number in Urdu digits'), code: '[$-3000000]#,##0.00' },
  ],
};

/** The number formats the window offers: everyone's, then its language's own. */
export const numberFormats = () => [...NUMBER_FORMATS, ...(REGIONAL_FORMATS[language()] || [])];

/* ── conditional formatting ──────────────────────────────────────────────── */

const CONDITIONS = [
  { id: 'greaterThan', label: t('Greater than'), operands: 1 },
  { id: 'lessThan', label: t('Less than'), operands: 1 },
  { id: 'between', label: t('Between'), operands: 2 },
  { id: 'equal', label: t('Equal to'), operands: 1 },
  { id: 'notEqual', label: t('Not equal to'), operands: 1 },
  { id: 'containsText', label: t('Text contains'), operands: 1, text: true },
  { id: 'duplicateValues', label: t('Duplicate values'), operands: 0 },
  { id: 'uniqueValues', label: t('Unique values'), operands: 0 },
];

const HIGHLIGHTS = [
  { label: t('Red text on light red'), colour: '9C0006', fill: 'FFC7CE' },
  { label: t('Amber text on light amber'), colour: '9C6500', fill: 'FFEB9C' },
  { label: t('Green text on light green'), colour: '006100', fill: 'C6EFCE' },
  { label: t('Bold only'), colour: null, fill: null, bold: true },
];

export function ConditionalDialog({ onClose, onApply, onClear, selection }) {
  const [rule, setRule] = useState('greaterThan');
  const [a, setA] = useState('');
  const [b, setB] = useState('');
  const [style, setStyle] = useState(0);
  const chosen = CONDITIONS.find((c) => c.id === rule);
  const look = HIGHLIGHTS[style];

  return (
    <Dialog
      title={t('Conditional formatting')}
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Clear rules')} onClick={() => onClear(true)} />
          <Button label={t('Cancel')} onClick={onClose} />
          <Button
            primary
            label={t('Apply')}
            disabled={chosen.operands > 0 && !a}
            onClick={() =>
              onApply({
                type: rule,
                operator: rule,
                values: [a, b].slice(0, chosen.operands).filter((v) => v !== ''),
                format: { fontColour: look.colour, fill: look.fill, bold: look.bold },
              })
            }
          />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('Apply to')} hint={t('The cells selected in the grid.')}>
          <Input value={selection || ''} disabled />
        </Field>
        <Field label={t('Format cells where the value is')}>
          <Select value={rule} onChange={(e) => setRule(e.target.value)}>
            {CONDITIONS.map((c) => (
              <option key={c.id} value={c.id}>{c.label}</option>
            ))}
          </Select>
        </Field>
        {chosen.operands > 0 ? (
          <div className="ml-servers" style={{ gridTemplateColumns: chosen.operands === 2 ? '1fr 1fr' : '1fr' }}>
            <Field label={chosen.operands === 2 ? t('From') : t('Value')}>
              <Input value={a} onChange={(e) => setA(e.target.value)} placeholder={chosen.text ? t('text') : '0'} autoFocus />
            </Field>
            {chosen.operands === 2 ? (
              <Field label={t('To')}>
                <Input value={b} onChange={(e) => setB(e.target.value)} placeholder="100" />
              </Field>
            ) : null}
          </div>
        ) : null}
        <Field label={t('Show them as')}>
          <Select value={String(style)} onChange={(e) => setStyle(Number(e.target.value))}>
            {HIGHLIGHTS.map((h, i) => (
              <option key={h.label} value={String(i)}>{h.label}</option>
            ))}
          </Select>
        </Field>
      </div>
    </Dialog>
  );
}

/* ── validation ──────────────────────────────────────────────────────────── */

export function ValidationDialog({ onClose, onApply, onClear, selection }) {
  const [type, setType] = useState('list');
  const [list, setList] = useState('');
  const [min, setMin] = useState('');
  const [max, setMax] = useState('');
  const [message, setMessage] = useState('');

  return (
    <Dialog
      title={t('Data validation')}
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Clear')} onClick={() => onClear(false)} />
          <Button label={t('Cancel')} onClick={onClose} />
          <Button
            primary
            label={t('Apply')}
            onClick={() =>
              onApply({
                type,
                formula1: type === 'list' ? list : min,
                formula2: type === 'list' ? undefined : max,
                prompt: message || undefined,
                allowBlank: true,
              })
            }
          />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('Apply to')}>
          <Input value={selection || ''} disabled />
        </Field>
        <Field label={t('Allow')}>
          <Select value={type} onChange={(e) => setType(e.target.value)}>
            <option value="list">{t('A list of values')}</option>
            <option value="whole">{t('Whole numbers')}</option>
            <option value="decimal">{t('Decimal numbers')}</option>
            <option value="date">{t('Dates')}</option>
            <option value="textLength">{t('Text of a given length')}</option>
          </Select>
        </Field>
        {type === 'list' ? (
          <Field label={t('Values')} hint={t('Separated by commas, or a range such as $H$1:$H$9.')}>
            <Input value={list} onChange={(e) => setList(e.target.value)} placeholder={t('Draft, In review, Approved')} autoFocus />
          </Field>
        ) : (
          <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr' }}>
            <Field label={t('Minimum')}>
              <Input value={min} onChange={(e) => setMin(e.target.value)} placeholder="0" autoFocus />
            </Field>
            <Field label={t('Maximum')}>
              <Input value={max} onChange={(e) => setMax(e.target.value)} placeholder="100" />
            </Field>
          </div>
        )}
        <Field label={t('Message when the cell is selected')}>
          <Input value={message} onChange={(e) => setMessage(e.target.value)} placeholder={t('Optional')} />
        </Field>
      </div>
    </Dialog>
  );
}

/* ── what-if ─────────────────────────────────────────────────────────────── */

export function GoalSeekDialog({ onClose, onRun, active }) {
  const [set, setSet] = useState(active || '');
  const [to, setTo] = useState('');
  const [by, setBy] = useState('');
  const [result, setResult] = useState(null);

  return (
    <Dialog
      title={t('Goal seek')}
      width={460}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Close')} onClick={onClose} />
          <Button
            primary
            label={t('Find')}
            disabled={!set || !to || !by}
            onClick={async () => setResult(await onRun({ set, to: Number(to), by }))}
          />
        </>
      }
    >
      <p style={{ marginTop: 0, fontSize: 12.5 }}>
        {t('Work backwards: name a formula cell, the answer you want from it, and the cell that may change to get there.')}
      </p>
      <div className="ml-form">
        <Field label={t('Set cell')}>
          <Input value={set} onChange={(e) => setSet(e.target.value.toUpperCase())} placeholder="B10" autoFocus />
        </Field>
        <Field label={t('To value')}>
          <Input value={to} onChange={(e) => setTo(e.target.value)} placeholder="50000" />
        </Field>
        <Field label={t('By changing cell')}>
          <Input value={by} onChange={(e) => setBy(e.target.value.toUpperCase())} placeholder="B4" />
        </Field>
        {result ? (
          <div className={`ml-result ${result.converged ? 'good' : 'bad'}`}>
            <Icon name={result.converged ? 'check' : 'info'} size={14} />
            <span>
              {result.converged
                ? t('{by} is {value} — {set} reaches {reached}', { by, value: result.value, set, reached: result.reached })
                : t('No value of {by} gets {set} there. Closest: {reached}', { by, set, reached: result.reached ?? '—' })}
            </span>
          </div>
        ) : null}
      </div>
    </Dialog>
  );
}

export function DataTableDialog({ onClose, onRun, selection }) {
  const [rowInput, setRowInput] = useState('');
  const [colInput, setColInput] = useState('');

  return (
    <Dialog
      title={t('What-if table')}
      width={480}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Fill')} disabled={!rowInput && !colInput} onClick={() => onRun({ range: selection, rowInput, colInput })} />
        </>
      }
    >
      <p style={{ marginTop: 0, fontSize: 12.5 }}>
        {t('The selection holds a formula in its corner and the values to try along its edges. Each cell is filled with what the formula gives for that combination.')}
      </p>
      <div className="ml-form">
        <Field label={t('Range')}>
          <Input value={selection || ''} disabled />
        </Field>
        <Field label={t('Row input cell')} hint={t('The cell the values along the top belong to.')}>
          <Input value={rowInput} onChange={(e) => setRowInput(e.target.value.toUpperCase())} placeholder="B3" autoFocus />
        </Field>
        <Field label={t('Column input cell')} hint={t('The cell the values down the side belong to.')}>
          <Input value={colInput} onChange={(e) => setColInput(e.target.value.toUpperCase())} placeholder="B4" />
        </Field>
      </div>
    </Dialog>
  );
}

/* ── names ───────────────────────────────────────────────────────────────── */

export function NameManager({ names, selection, onClose, onDefine, onDelete, onGoto }) {
  const [name, setName] = useState('');
  const [ref, setRef] = useState(selection || '');

  useEffect(() => setRef(selection || ''), [selection]);

  return (
    <Dialog
      title={t('Named ranges')}
      width={540}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Close')} onClick={onClose} />
          <Button primary label={t('Define')} disabled={!name || !ref} onClick={() => { onDefine(name, ref); setName(''); }} />
        </>
      }
    >
      <div className="ml-form">
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1.4fr' }}>
          <Field label={t('Name')}>
            <Input value={name} onChange={(e) => setName(e.target.value.replace(/\s/g, '_'))} placeholder={t('Revenue')} autoFocus />
          </Field>
          <Field label={t('Refers to')}>
            <Input value={ref} onChange={(e) => setRef(e.target.value)} placeholder="Sheet1!$B$2:$B$40" />{/* words-ok: a reference */}
          </Field>
        </div>
      </div>

      {names?.length ? (
        <div className="ml-import-folders" style={{ maxHeight: 240 }}>
          {names.map((n) => (
            <div key={n.name} className="ml-import-folder">
              <Icon name="find" size={13} />
              <button type="button" className="ml-linkchip name" style={{ flex: 1, textAlign: 'left' }} onClick={() => onGoto(n.name)}>
                <strong>{n.name}</strong> <span style={{ opacity: 0.7 }}>{n.ref || n.refersTo}</span>
              </button>
              <Button icon="trash" title={t('Delete {name}', { name: n.name })} onClick={() => onDelete(n.name)} />
            </div>
          ))}
        </div>
      ) : (
        <Empty icon="find" title={t('No names yet')}>
          {t('A name turns')} <code>$B$2:$B$40</code> {t('into')} <code>{t('Revenue')}</code>{t(', in every formula that uses it.')}
        </Empty>
      )}
    </Dialog>
  );
}

/* ── find and replace ────────────────────────────────────────────────────── */

export function FindDialog({ onClose, onFind, onReplace, onReplaceAll }) {
  const [find, setFind] = useState('');
  const [replace, setReplace] = useState('');
  const [note, setNote] = useState(null);

  return (
    <Dialog
      title={t('Find and replace')}
      width={480}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Close')} onClick={onClose} />
          <Button label={t('Find next')} disabled={!find} onClick={async () => setNote(await onFind(find))} />
          <Button label={t('Replace')} disabled={!find} onClick={async () => setNote(await onReplace(find, replace))} />
          <Button primary label={t('Replace all')} disabled={!find} onClick={async () => setNote(await onReplaceAll(find, replace))} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('Find')}>
          <Input value={find} onChange={(e) => setFind(e.target.value)} autoFocus />
        </Field>
        <Field label={t('Replace with')}>
          <Input value={replace} onChange={(e) => setReplace(e.target.value)} />
        </Field>
        {note ? <div className="ml-note"><Icon name="info" size={14} />{note}</div> : null}
      </div>
    </Dialog>
  );
}

/* ── pivot ───────────────────────────────────────────────────────────────── */

/** A1 → { row, col }, zero-based; null for anything that is not a cell reference. */
export function parseRef(text) {
  const m = /^\s*\$?([A-Za-z]{1,3})\$?(\d{1,7})\s*$/.exec(String(text || ''));
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
}

/** Go To: a cell reference or a defined name, the way Ctrl+G asks for one. */
export function GoToDialog({ onClose, onGo, names = [] }) {
  const [text, setText] = useState('');
  const ok = Boolean(parseRef(text)) || names.some((n) => n.name === text.trim());
  return (
    <Dialog
      title={t('Go To')}
      width={380}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Go')} disabled={!ok} onClick={() => onGo(text.trim())} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('Reference')} hint={t('A cell such as C12, or a defined name.')}>
          <Input value={text} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && ok) onGo(text.trim()); }} autoFocus placeholder="C12" />
        </Field>
        {names.length ? (
          <div className="ml-found">
            {names.slice(0, 8).map((n) => (
              <button key={n.name} type="button" className="ml-found-item" onClick={() => onGo(n.name)}>
                <span className="grow"><div className="who">{n.name}</div><div className="what">{n.ref || ''}</div></span>
              </button>
            ))}
          </div>
        ) : null}
      </div>
    </Dialog>
  );
}

/**
 * A hyperlink on the active cell: an address that opens outside the suite
 * (http, https, mailto — a bare domain or an email address is taken as
 * one), or a place in this workbook (C12, Sheet2!B4, a defined name).
 */
export function LinkDialog({ current = null, cellRef = '', onClose, onSet, onRemove }) {
  const [to, setTo] = useState(current?.href || current?.location || '');
  const [tip, setTip] = useState(current?.tooltip || '');
  const value = to.trim();
  const place = /^(?:'[^']+'|[^!'\s]+)![A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?$/.test(value) || Boolean(parseRef(value)) || /^[A-Za-z_][\w.]*$/.test(value) && !/^[A-Za-z]{1,3}\d+$/.test(value) && !/\./.test(value);
  const address = /^(https?:\/\/|mailto:)/i.test(value) ? value
    : /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(value) ? 'mailto:' + value
    : /^[\w-]+(\.[\w-]+)+(\/\S*)?$/i.test(value) ? 'https://' + value
    : null;
  const ok = Boolean(address || place);
  const submit = () => {
    if (!ok) return;
    onSet(address ? { href: address, tooltip: tip.trim() || null } : { location: value, tooltip: tip.trim() || null });
  };
  return (
    <Dialog
      title={current ? t('The link on {cell}', { cell: cellRef }) : t('A link on {cell}', { cell: cellRef })}
      width={440}
      onClose={onClose}
      actions={
        <>
          {current ? <Button label={t('Remove link')} className="sh-link-remove" onClick={onRemove} /> : null}
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={current ? t('Change') : t('Add link')} className="sh-link-ok" disabled={!ok} onClick={submit} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('Link to')} hint={t('An address such as rutba.io or hello@rutba.io, or a place in this workbook such as C12, Sheet2!B4 or a name.')}>
          <Input className="sh-link-to" value={to} onChange={(e) => setTo(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} autoFocus placeholder={t('https://…  or  Sheet2!B4')} />
        </Field>
        <Field label={t('Screen tip')} hint={t('Shown when the pointer rests on the cell. Optional.')}>
          <Input className="sh-link-tip" value={tip} onChange={(e) => setTip(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} placeholder={t('Where this goes')} />
        </Field>
        <p className="sh-link-says" style={{ margin: 0, fontSize: 12, color: 'var(--ink-3)' }}>
          {!value ? t('Nothing yet.') : address ? t('Opens {address} outside the suite.', { address }) : place ? t('Goes to {place} in this workbook.', { place: value }) : t('Not an address and not a place this workbook knows.')}
        </p>
      </div>
    </Dialog>
  );
}

/**
 * A note on a cell: the words, and who wrote them. The author is remembered
 * for the next note; the engine signs an unsigned one with the account at
 * the keyboard.
 */
export function NoteDialog({ current = null, cellRef = '', onClose, onSet, onRemove }) {
  const remembered = (() => { try { return localStorage.getItem('sheets.noteAuthor') || ''; } catch { return ''; } })();
  const [text, setText] = useState(current?.text || '');
  const [author, setAuthor] = useState(current?.author || remembered);
  const ok = text.trim().length > 0;
  const submit = () => {
    if (!ok) return;
    try { localStorage.setItem('sheets.noteAuthor', author.trim()); } catch { /* a private window */ }
    onSet({ text: text.trim(), author: author.trim() });
  };
  return (
    <Dialog
      title={current ? t('The note on {cell}', { cell: cellRef }) : t('A note on {cell}', { cell: cellRef })}
      width={440}
      onClose={onClose}
      actions={
        <>
          {current ? <Button label={t('Delete note')} className="sh-note-remove" onClick={onRemove} /> : null}
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={current ? t('Change') : t('Add note')} className="sh-note-ok" disabled={!ok} onClick={submit} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('Note')} hint={t('Shown when the pointer rests on the cell; Excel shows it the same way.')}>
          <textarea
            className="rw-input sh-note-text"
            rows={5}
            value={text}
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) submit(); }}
            autoFocus
            placeholder={t('What to say about this cell')}
            style={{ resize: 'vertical', minHeight: 90, font: 'inherit', width: '100%', boxSizing: 'border-box' }}
          />
        </Field>
        <Field label={t('Author')} hint={t('Who signs it. Left empty, the note is signed with your account\'s name.')}>
          <Input className="sh-note-author" value={author} onChange={(e) => setAuthor(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} placeholder={t('Your name')} />
        </Field>
      </div>
    </Dialog>
  );
}

/**
 * Insert → Sparklines: a small line or set of bars, drawn in one cell from
 * the numbers in a range beside it. Prefilled the way a person reaches for
 * it — the current selection as the data, and the cell past its last
 * column as where the sparkline goes (one per row, for a selection of more
 * than one) — so OK alone is usually enough.
 */
export function SparklineDialog({ type, data: dataDefault = '', at: atDefault = '', onClose, onApply }) {
  const [data, setData] = useState(dataDefault);
  const [at, setAt] = useState(atDefault);
  const ok = data.trim().length > 0 && at.trim().length > 0;
  const submit = () => { if (ok) onApply({ type, data: data.trim(), at: at.trim() }); };
  const label = type === 'column' ? t('Column sparklines') : t('Line sparklines');
  return (
    <Dialog
      title={label}
      width={380}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} className="sh-sparkline-ok" disabled={!ok} onClick={submit} />
        </>
      }
    >
      <div className="ml-form sh-sparkline">
        <Field label={t('Data range')} hint={t('The numbers each sparkline is drawn from.')}>
          <Input className="sh-sparkline-data" value={data} onChange={(e) => setData(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && ok) submit(); }} autoFocus placeholder="B2:F2" />
        </Field>
        <Field label={t('Location range')} hint={t('Where the sparkline goes — one cell, or one per row of the data.')}>
          <Input className="sh-sparkline-at" value={at} onChange={(e) => setAt(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && ok) submit(); }} placeholder="G2" />
        </Field>
      </div>
    </Dialog>
  );
}

/**
 * Insert → Header & Footer: what prints at the top and the foot of every
 * page, with Excel's codes — &P page, &N pages, &A sheet, &F file, &D date,
 * &T time — and &L, &C, &R starting the left, centre and right parts.
 */
export function HeaderFooterDialog({ current = null, onClose, onSet }) {
  const [header, setHeader] = useState(current?.header || '');
  const [footer, setFooter] = useState(current?.footer ?? '');
  const submit = () => onSet({ header: header.trim(), footer: footer.trim() });
  const picks = [[t('Page &P of &N'), t('Page &P of &N')], [t('Sheet name'), '&A'], [t('File name'), '&F'], [t('Date'), '&D'], [t('Left, centre, right'), '&L&F&C&A&R&D']];
  return (
    <Dialog
      title={t('Header and footer')}
      width={480}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Set')} className="sh-hf-ok" onClick={submit} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('Header')} hint={t('Printed at the top of every page. Codes: &P page, &N pages, &A sheet, &F file, &D date, &T time; &L, &C and &R start the left, centre and right parts.')}>
          <Input className="sh-hf-header" value={header} onChange={(e) => setHeader(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} placeholder={t('Nothing at the top')} autoFocus />
        </Field>
        <Field label={t('Footer')} hint={t('Printed at the foot of every page; empty prints nothing there.')}>
          <Input className="sh-hf-footer" value={footer} onChange={(e) => setFooter(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} placeholder={t('Nothing at the foot')} />
        </Field>
        <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
          {picks.map(([label, code]) => (
            <Button key={code} label={label} title={t('Put {code} in the footer', { code })} onClick={() => setFooter((f) => (f ? f + ' ' : '') + code)} />
          ))}
        </div>
      </div>
    </Dialog>
  );
}

/** Insert Function: Excel's categories, a pick starts `=NAME(` in the active cell. */
export function FunctionDialog({ onClose, onPick, catalogue }) {
  const categories = Object.keys(catalogue);
  const [category, setCategory] = useState(categories[0]);
  const [filter, setFilter] = useState('');
  const names = filter.trim()
    ? [...new Set(Object.values(catalogue).flat())].filter((n) => n.includes(filter.trim().toUpperCase()))
    : catalogue[category] || [];
  return (
    <Dialog title={t('Insert function')} width={480} onClose={onClose} actions={<Button label={t('Cancel')} onClick={onClose} />}>
      <div className="ml-form">
        <Field label={t('Search')}>
          <Input value={filter} onChange={(e) => setFilter(e.target.value)} placeholder={t('Type part of a name')} autoFocus />
        </Field>
        {filter.trim() ? null : (
          <Field label={t('Category')}>
            <Select value={category} onChange={(e) => setCategory(e.target.value)}>
              {categories.map((c) => <option key={c} value={c}>{c}</option>)}
            </Select>
          </Field>
        )}
        <div className="ml-found" style={{ maxHeight: 260, overflow: 'auto' }}>
          {names.map((n) => (
            <button key={n} type="button" className="ml-found-item" onClick={() => onPick(n)}>
              <span className="ml-found-logo"><Icon name="formula" size={15} /></span>
              <span className="grow"><div className="who">{n}</div></span>
            </button>
          ))}
          {names.length ? null : <Empty title={t('No function by that name')} />}
        </div>
      </div>
    </Dialog>
  );
}

/** Workbook Statistics: what the model knows, as Excel's dialog lists it. */
export function StatisticsDialog({ model, onClose }) {
  const rows = [
    [t('Sheets'), (model?.sheets || []).length],
    [t('Active sheet'), model?.activeSheet || '—'],
    [t('Cells with content'), (model?.cells || []).filter((c) => c.text !== '' && c.text != null).length + (model?.cells ? ` ${t('(on screen)')}` : '')],
    [t('Formulas'), (model?.cells || []).filter((c) => c.isFormula).length + ` ${t('(on screen)')}`],
    [t('Defined names'), (model?.names || []).length],
    [t('Tables'), (model?.tables || []).length],
    [t('Frozen'), model?.frozen && (model.frozen.rows || model.frozen.cols) ? t('{rows} rows, {cols} columns', { rows: model.frozen.rows, cols: model.frozen.cols }) : t('nothing')],
    [t('Filter'), model?.filtered ? t('on') : t('off')],
    [t('Protection'), model?.protection?.sheet ? t('sheet protected') : t('none')],
  ];
  return (
    <Dialog title={t('Workbook statistics')} width={420} onClose={onClose} actions={<Button primary label={t('Close')} onClick={onClose} />}>
      <dl className="about-list">
        {rows.map(([k, v]) => (
          <React.Fragment key={k}><dt>{k}</dt><dd>{String(v)}</dd></React.Fragment>
        ))}
      </dl>
    </Dialog>
  );
}

const SHEET_SHORTCUTS = [
  ['Ctrl+S', t('Save')], ['Ctrl+Z / Ctrl+Y', t('Undo / Redo')], ['Ctrl+C / Ctrl+X / Ctrl+V', t('Copy / Cut / Paste')],
  ['Ctrl+B / Ctrl+I / Ctrl+U', t('Bold / Italic / Underline')], ['F2', t('Edit the active cell')], ['Enter / Tab', t('Commit and move down / right')],
  ['Escape', t('Cancel the edit')], ['Delete', t('Clear the selection')], ['Shift+Arrows', t('Extend the selection')],
  ['Ctrl+Arrows', t('Jump to the edge of the data')], ['Ctrl+A', t('Select all')], ['Ctrl+D / Ctrl+R', t('Fill down / right')],
  ['Ctrl+F / Ctrl+H', t('Find / Replace')], ['Ctrl+G', t('Go To')], ['Ctrl+`', t('Show formulas')],
  ['F9', t('Calculate Now — every sheet')], ['Shift+F9', t('Calculate Sheet — this sheet')],
];

export function SheetShortcutsDialog({ onClose }) {
  return (
    <Dialog title={t('Keyboard shortcuts')} width={440} onClose={onClose} actions={<Button primary label={t('Close')} onClick={onClose} />}>
      <dl className="about-list">
        {SHEET_SHORTCUTS.map(([k, v]) => (
          <React.Fragment key={k}><dt>{k}</dt><dd>{v}</dd></React.Fragment>
        ))}
      </dl>
    </Dialog>
  );
}

/** Row height or column width, in pixels, for the rows or columns selected. */
/**
 * Data → Sort: up to three keys, each a column of the block and a direction,
 * the first deciding and the next breaking its ties — Excel's Sort dialog,
 * to the level people use.
 */
export function SortDialog({ columns, onClose, onSort }) {
  const first = columns[0]?.col ?? 0;
  const [keys, setKeys] = useState([{ col: first, ascending: true }]);
  const set = (i, patch) => setKeys((ks) => ks.map((k, j) => (j === i ? { ...k, ...patch } : k)));
  const ok = keys.length > 0 && keys.every((k) => Number.isFinite(k.col));
  return (
    <Dialog
      title={t('Sort')}
      width={440}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Sort')} className="sh-sort-ok" disabled={!ok} onClick={() => onSort(keys)} />
        </>
      }
    >
      <div className="ml-form">
        {keys.map((k, i) => (
          <div key={i} style={{ display: 'flex', gap: 8, alignItems: 'end' }}>
            <Field label={i === 0 ? t('Sort by') : t('Then by')} style={{ flex: 1 }}>
              <Select className={`sh-sort-col-${i}`} value={String(k.col)} onChange={(e) => set(i, { col: Number(e.target.value) })} style={{ width: '100%' }}>
                {columns.map((c) => <option key={c.col} value={String(c.col)}>{c.name}</option>)}
              </Select>
            </Field>
            <Field label={t('Order')}>
              <Select className={`sh-sort-dir-${i}`} value={k.ascending ? 'asc' : 'desc'} onChange={(e) => set(i, { ascending: e.target.value === 'asc' })} style={{ width: 150 }}>
                <option value="asc">{t('A to Z, small to large')}</option>
                <option value="desc">{t('Z to A, large to small')}</option>
              </Select>
            </Field>
            {i > 0 ? <Button icon="close" title={t('Remove this level')} onClick={() => setKeys((ks) => ks.filter((_, j) => j !== i))} /> : null}
          </div>
        ))}
        {keys.length < 3 ? <Button label={t('Add a level')} onClick={() => setKeys((ks) => [...ks, { col: columns[Math.min(ks.length, columns.length - 1)]?.col ?? first, ascending: true }])} /> : null}
      </div>
    </Dialog>
  );
}

/** Rename a sheet tab: Excel's rules, said in the hint; Enter renames. */
export function SheetNameDialog({ current, onClose, onRename }) {
  const [value, setValue] = useState(current || '');
  const ok = value.trim().length > 0 && value.trim().length <= 31 && !/[\[\]:*?\/\\]/.test(value);
  return (
    <Dialog
      title={t('Rename sheet')}
      width={380}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Rename')} className="sh-sheetname-ok" disabled={!ok} onClick={() => onRename(value.trim())} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('Name')} hint={t('Up to 31 characters; none of [ ] : * ? / or the backslash. Formulas that read the sheet follow the new name.')}>
          <Input className="sh-sheetname" value={value} onChange={(e) => setValue(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && ok) onRename(value.trim()); }} autoFocus />
        </Field>
      </div>
    </Dialog>
  );
}

/** Delete a sheet: asked first, as Excel asks — and here Undo puts it back. */
export function SheetDeleteDialog({ name, onClose, onDelete }) {
  return (
    <Dialog
      title={t('Delete sheet')}
      width={380}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Delete')} className="sh-sheetdelete-ok" onClick={onDelete} />
        </>
      }
    >
      <div className="ml-form">
        <p style={{ margin: 0 }}>{t('Delete the sheet "{name}" and everything on it? Undo (Ctrl+Z) puts it back.', { name })}</p>
      </div>
    </Dialog>
  );
}

export function SizeDialog({ kind, current, onClose, onApply }) {
  const [value, setValue] = useState(String(current || (kind === 'row' ? 20 : 64)));
  const n = Number(value);
  const ok = Number.isFinite(n) && n >= 4 && n <= 1000;
  return (
    <Dialog
      title={kind === 'row' ? t('Row height') : t('Column width')}
      width={340}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} disabled={!ok} onClick={() => onApply(n)} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={kind === 'row' ? t('Height (px)') : t('Width (px)')} hint={t('Applies to every selected row or column.')}>
          <Input value={value} onChange={(e) => setValue(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && ok) onApply(n); }} autoFocus />
        </Field>
      </div>
    </Dialog>
  );
}

/** PivotChart's kinds: every chart but scatter, as Excel offers them. */
const PIVOT_CHARTS = [['column', t('Column')], ['bar', t('Bar')], ['line', t('Line')], ['area', t('Area')], ['pie', t('Pie')], ['doughnut', t('Doughnut')]];

/**
 * Insert → PivotTable, and Insert → PivotChart on data (PivotChart &
 * PivotTable): the source — the list round the cursor to start with — and
 * which of its columns go down the rows, across the columns and into the
 * values, summarised how; for a chart, its kind as well.
 */
export function PivotDialog({ onClose, onCreate, list = null, sheet = '', chart = false }) {
  const columns = (list?.columns || []).map((c) => c.name);
  const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${String(sheet).replace(/'/g, "''")}'`;
  const [source, setSource] = useState(list?.ref ? `${quoted}!${list.ref}` : '');
  const firstText = columns[0] || '';
  const lastNumber = columns[columns.length - 1] || '';
  const [rows, setRows] = useState(firstText);
  const [cols, setCols] = useState('');
  const [values, setValues] = useState(columns.length > 1 ? lastNumber : '');
  const [fn, setFn] = useState('SUM');
  const [kind, setKind] = useState('column');
  const pick = (value, set, label, cls) => (
    <Field label={label}>
      <Select className={`rw-select ${cls}`} value={value} onChange={(e) => set(e.target.value)}>
        <option value="">{t('(none)')}</option>
        {columns.map((c) => <option key={c} value={c}>{c}</option>)}
      </Select>
    </Field>
  );
  const ok = Boolean(source.trim() && values && (!rows || rows !== cols));
  return (
    <Dialog
      title={chart ? t('PivotChart & PivotTable') : t('PivotTable')}
      width={chart ? 560 : 500}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button
            primary
            label={t('OK')}
            className="sh-pivot-ok"
            disabled={!ok}
            onClick={() => onCreate({
              source: source.trim(),
              rowFields: rows ? [rows] : [],
              colFields: cols ? [cols] : [],
              dataFields: [{ name: values, fn }],
              ...(chart ? { chart: { kind } } : {}),
            })}
          />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('Table or range')} hint={t('The data with its headings in the first row — the list round the cell to start with.')}>
          <Input className="rw-input sh-pivot-source" value={source} onChange={(e) => setSource(e.target.value)} placeholder="Sheet1!A1:E500" autoFocus />{/* words-ok: a reference */}
        </Field>
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr' }}>
          {pick(rows, setRows, t('Rows'), 'sh-pivot-rows')}
          {pick(cols, setCols, t('Columns'), 'sh-pivot-cols')}
        </div>
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 150px' }}>
          {pick(values, setValues, t('Values'), 'sh-pivot-values')}
          <Field label={t('Summarise by')}>
            <Select className="rw-select sh-pivot-fn" value={fn} onChange={(e) => setFn(e.target.value)}>
              {[['SUM', t('Sum')], ['COUNT', t('Count')], ['AVERAGE', t('Average')], ['MAX', t('Max')], ['MIN', t('Min')]].map(([f, label]) => (
                <option key={f} value={f}>{label}</option>
              ))}
            </Select>
          </Field>
        </div>
        {chart ? (
          <Field label={t('Chart')}>
            <div className="sh-pivot-kinds">
              {PIVOT_CHARTS.map(([k, label]) => (
                <button key={k} type="button" className={`sh-pivot-kind${kind === k ? ' on' : ''}`} data-kind={k} onClick={() => setKind(k)}>
                  <Icon name="chart" size={16} />
                  <span>{label}</span>
                </button>
              ))}
            </div>
          </Field>
        ) : null}
        <p className="rw-hint" style={{ margin: 0 }}>
          {chart
            ? t('The pivot table goes under the data and its chart beside it; the chart follows the pivot when it is refreshed or filtered.')
            : t('The pivot table goes under the data. Data → Refresh All recalculates it from the source.')}
        </p>
      </div>
    </Dialog>
  );
}

/* ── Data → Outline ──────────────────────────────────────────────────────── */

const CHECK_ROW = { display: 'flex', alignItems: 'center', gap: 8, fontSize: 12.5 };

/**
 * Group or Ungroup on a selection that is neither whole rows nor whole
 * columns: Excel asks which, and so does this.
 */
export function OutlineAxisDialog({ verb, onClose, onPick }) {
  const [axis, setAxis] = useState('row');
  const title = verb === 'ungroup' ? t('Ungroup') : t('Group');
  return (
    <Dialog
      title={title}
      width={300}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={title} className="sh-outline-ok" onClick={() => onPick(axis)} />
        </>
      }
    >
      <div className="ml-form">
        <label style={CHECK_ROW}><input type="radio" name="sh-outline-axis" className="sh-outline-rows" checked={axis === 'row'} onChange={() => setAxis('row')} /> {t('Rows')}</label>
        <label style={CHECK_ROW}><input type="radio" name="sh-outline-axis" className="sh-outline-cols" checked={axis === 'col'} onChange={() => setAxis('col')} /> {t('Columns')}</label>
      </div>
    </Dialog>
  );
}

/** Subtotal's functions in the order Excel's dialog lists them. */
const SUBTOTAL_FNS = [
  ['sum', t('Sum')], ['count', t('Count')], ['average', t('Average')], ['max', t('Max')], ['min', t('Min')], ['product', t('Product')],
  ['countNumbers', t('Count Numbers')], ['stdDev', t('StdDev')], ['stdDevp', t('StdDevp')], ['var', t('Var')], ['varp', t('Varp')],
];

/**
 * Data → Subtotal: at each change in one column, a function over the
 * columns ticked; Replace, page breaks and where the summaries sit, as
 * Excel's dialog has them; Remove All takes every subtotal off the list.
 * The last column is ticked to start with, as Excel ticks it.
 */
export function SubtotalDialog({ list, onClose, onApply, onRemoveAll }) {
  const columns = list.columns || [];
  const [by, setBy] = useState(columns[0]?.col ?? 0);
  const [fn, setFn] = useState('sum');
  const [ticked, setTicked] = useState(() => new Set(columns.length ? [columns[columns.length - 1].col] : []));
  const [replace, setReplace] = useState(true);
  const [breaks, setBreaks] = useState(false);
  const [below, setBelow] = useState(true);
  const toggle = (col) => setTicked((s) => {
    const next = new Set(s);
    if (next.has(col)) next.delete(col);
    else next.add(col);
    return next;
  });
  const ok = ticked.size > 0;
  return (
    <Dialog
      title={t('Subtotal')}
      width={400}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Remove All')} className="sh-sub-remove" disabled={!list.subtotals} onClick={onRemoveAll} />
          <span style={{ flex: 1 }} />
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} className="sh-sub-ok" disabled={!ok} onClick={() => onApply({ groupBy: by, fn, columns: [...ticked], replace, pageBreaks: breaks, summaryBelow: below })} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('At each change in')}>
          <Select className="sh-sub-by" value={String(by)} onChange={(e) => setBy(Number(e.target.value))} style={{ width: '100%' }}>
            {columns.map((c) => <option key={c.col} value={String(c.col)}>{c.name}</option>)}
          </Select>
        </Field>
        <Field label={t('Use function')}>
          <Select className="sh-sub-fn" value={fn} onChange={(e) => setFn(e.target.value)} style={{ width: '100%' }}>
            {SUBTOTAL_FNS.map(([key, name]) => <option key={key} value={key}>{name}</option>)}
          </Select>
        </Field>
        <Field label={t('Add subtotal to')}>
          <div className="sh-sub-cols">
            {columns.map((c) => (
              <label key={c.col} style={CHECK_ROW}>
                <input type="checkbox" className={`sh-sub-col sh-sub-col-${c.col}`} checked={ticked.has(c.col)} onChange={() => toggle(c.col)} /> {c.name}
              </label>
            ))}
          </div>
        </Field>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <label style={CHECK_ROW}><input type="checkbox" className="sh-sub-replace" checked={replace} onChange={(e) => setReplace(e.target.checked)} /> {t('Replace current subtotals')}</label>
          <label style={CHECK_ROW}><input type="checkbox" className="sh-sub-breaks" checked={breaks} onChange={(e) => setBreaks(e.target.checked)} /> {t('Page break between groups')}</label>
          <label style={CHECK_ROW}><input type="checkbox" className="sh-sub-below" checked={below} onChange={(e) => setBelow(e.target.checked)} /> {t('Summary below data')}</label>
        </div>
        <div style={{ fontSize: 11.5, color: 'var(--ink-3)' }}>{list.subtotals ? tn(list.subtotals, 'The list: {ref} — it has {count} subtotal row now', 'The list: {ref} — it has {count} subtotal rows now', { ref: list.ref }) : t('The list: {ref}', { ref: list.ref })}</div>
      </div>
    </Dialog>
  );
}

/* ── Data → Advanced ─────────────────────────────────────────────────────── */

/**
 * Advanced Filter, as Excel's dialog asks it: filter the list where it is
 * or copy the rows that pass to another place; the list, the criteria
 * range and the destination as references; unique records only. The
 * ranges an earlier run used are offered again, as Excel offers them.
 */
/**
 * Home → Paste → Paste Special, Excel's dialog: what of the copied cells is
 * pasted, an operation that works the copy out against what is there, and
 * Skip blanks and Transpose.
 */
export function PasteSpecialDialog({ onClose, onApply }) {
  const [what, setWhat] = useState('all');
  const [operation, setOperation] = useState('none');
  const [skipBlanks, setSkipBlanks] = useState(false);
  const [transpose, setTranspose] = useState(false);
  const apply = () => onApply({ what, operation: what === 'formats' ? 'none' : operation, skipBlanks, transpose });
  const radio = (name, value, current, set, label, disabled = false) => (
    <label style={{ ...CHECK_ROW, opacity: disabled ? 0.5 : 1 }}>
      <input type="radio" name={name} className={`sh-ps-${value}`} checked={current === value} disabled={disabled} onChange={() => set(value)} /> {label}
    </label>
  );
  return (
    <Dialog
      title={t('Paste Special')}
      width={420}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} className="sh-ps-ok" onClick={apply} />
        </>
      }
    >
      <div className="ml-form" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
        <Field label={t('Paste')}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {radio('sh-ps-what', 'all', what, setWhat, t('All'))}
            {radio('sh-ps-what', 'formulas', what, setWhat, t('Formulas'))}
            {radio('sh-ps-what', 'values', what, setWhat, t('Values'))}
            {radio('sh-ps-what', 'formats', what, setWhat, t('Formats'))}
          </div>
        </Field>
        <Field label={t('Operation')}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {radio('sh-ps-op', 'none', operation, setOperation, t('None'), what === 'formats')}
            {radio('sh-ps-op', 'add', operation, setOperation, t('Add'), what === 'formats')}
            {radio('sh-ps-op', 'subtract', operation, setOperation, t('Subtract'), what === 'formats')}
            {radio('sh-ps-op', 'multiply', operation, setOperation, t('Multiply'), what === 'formats')}
            {radio('sh-ps-op', 'divide', operation, setOperation, t('Divide'), what === 'formats')}
          </div>
        </Field>
        <label style={CHECK_ROW}><input type="checkbox" className="sh-ps-skip" checked={skipBlanks} onChange={(e) => setSkipBlanks(e.target.checked)} /> {t('Skip blanks')}</label>
        <label style={CHECK_ROW}><input type="checkbox" className="sh-ps-transpose" checked={transpose} onChange={(e) => setTranspose(e.target.checked)} /> {t('Transpose')}</label>
      </div>
    </Dialog>
  );
}

/**
 * Home → Fill → Series, Excel's dialog: down the columns or along the rows,
 * linear, growth, dates by a unit, or AutoFill, with a step and a stop.
 */
export function SeriesDialog({ rows = false, onClose, onApply }) {
  const [direction, setDirection] = useState(rows ? 'rows' : 'columns');
  const [type, setType] = useState('linear');
  const [unit, setUnit] = useState('day');
  const [step, setStep] = useState('1');
  const [stop, setStop] = useState('');
  const ok = type === 'autofill' || (step.trim() !== '' && Number.isFinite(Number(step)) && (stop.trim() === '' || Number.isFinite(Number(stop))));
  const apply = () => ok && onApply({ direction, type, unit, step: Number(step), stop: stop.trim() === '' ? null : Number(stop) });
  const enter = (e) => { if (e.key === 'Enter') apply(); };
  const radio = (name, value, current, set, label, disabled = false) => (
    <label style={{ ...CHECK_ROW, opacity: disabled ? 0.5 : 1 }}>
      <input type="radio" name={name} className={`sh-series-${value}`} checked={current === value} disabled={disabled} onChange={() => set(value)} /> {label}
    </label>
  );
  return (
    <Dialog
      title={t('Series')}
      width={440}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} className="sh-series-ok" disabled={!ok} onClick={apply} />
        </>
      }
    >
      <div className="ml-form" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 12 }}>
        <Field label={t('Series in')}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {radio('sh-series-dir', 'rows', direction, setDirection, t('Rows'))}
            {radio('sh-series-dir', 'columns', direction, setDirection, t('Columns'))}
          </div>
        </Field>
        <Field label={t('Type')}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {radio('sh-series-type', 'linear', type, setType, t('Linear'))}
            {radio('sh-series-type', 'growth', type, setType, t('Growth'))}
            {radio('sh-series-type', 'date', type, setType, t('Date'))}
            {radio('sh-series-type', 'autofill', type, setType, t('AutoFill'))}
          </div>
        </Field>
        <Field label={t('Date unit')}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {radio('sh-series-unit', 'day', unit, setUnit, t('Day'), type !== 'date')}
            {radio('sh-series-unit', 'weekday', unit, setUnit, t('Weekday'), type !== 'date')}
            {radio('sh-series-unit', 'month', unit, setUnit, t('Month'), type !== 'date')}
            {radio('sh-series-unit', 'year', unit, setUnit, t('Year'), type !== 'date')}
          </div>
        </Field>
        <Field label={t('Step value')}>
          <Input className="sh-series-step" value={step} disabled={type === 'autofill'} onChange={(e) => setStep(e.target.value)} onKeyDown={enter} autoFocus />
        </Field>
        <Field label={t('Stop value')} hint={t('Optional: with one cell selected, the series runs to here.')}>
          <Input className="sh-series-stop" value={stop} disabled={type === 'autofill'} onChange={(e) => setStop(e.target.value)} onKeyDown={enter} />
        </Field>
      </div>
    </Dialog>
  );
}

export function AdvancedFilterDialog({ list, onClose, onApply }) {
  const kept = list.filter || {};
  const [action, setAction] = useState(kept.extract ? 'copy' : 'filter');
  const [range, setRange] = useState(kept.list || list.ref || '');
  const [criteria, setCriteria] = useState(kept.criteria || '');
  const [copyTo, setCopyTo] = useState(kept.extract || '');
  const [unique, setUnique] = useState(false);
  const refOk = (t) => /^\s*(?:'[^']+'!|[A-Za-z0-9_.]+!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?\s*$/.test(t);
  const ok = refOk(range) && refOk(criteria) && (action === 'filter' || refOk(copyTo));
  const apply = () => ok && onApply({ list: range.trim(), criteria: criteria.trim(), action, copyTo: action === 'copy' ? copyTo.trim() : null, unique });
  const enter = (e) => { if (e.key === 'Enter') apply(); };
  return (
    <Dialog
      title={t('Advanced Filter')}
      width={420}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} className="sh-adv-ok" disabled={!ok} onClick={apply} />
        </>
      }
    >
      <div className="ml-form">
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <label style={CHECK_ROW}><input type="radio" name="sh-adv-action" className="sh-adv-inplace" checked={action === 'filter'} onChange={() => setAction('filter')} /> {t('Filter the list, in place')}</label>
          <label style={CHECK_ROW}><input type="radio" name="sh-adv-action" className="sh-adv-copy" checked={action === 'copy'} onChange={() => setAction('copy')} /> {t('Copy to another location')}</label>
        </div>
        <Field label={t('List range')}>
          <Input className="sh-adv-list" value={range} onChange={(e) => setRange(e.target.value)} onKeyDown={enter} placeholder={t('A1:D20, or Data!A1:D20 to copy from another sheet')} />
        </Field>
        <Field label={t('Criteria range')} hint={t('Headings from the list over rows of conditions: a row\'s conditions all hold, any one row is enough. East, >250, <>West, =B*.')}>
          <Input className="sh-adv-criteria" value={criteria} onChange={(e) => setCriteria(e.target.value)} onKeyDown={enter} placeholder="F1:G3" autoFocus />
        </Field>
        <Field label={t('Copy to')}>
          <Input className="sh-adv-to" value={copyTo} disabled={action !== 'copy'} onChange={(e) => setCopyTo(e.target.value)} onKeyDown={enter} placeholder="I1" />
        </Field>
        <label style={CHECK_ROW}><input type="checkbox" className="sh-adv-unique" checked={unique} onChange={(e) => setUnique(e.target.checked)} /> {t('Unique records only')}</label>
      </div>
    </Dialog>
  );
}

/**
 * Formulas → Evaluate Formula, Excel's dialog: the cell's reference beside
 * its formula, the part the next Evaluate works out underlined and the
 * last result in italics; Step In opens the underlined cell's own formula
 * in a box below, Step Out puts its value back. The engine replays the
 * presses from the start each time (`load(actions)`), so the dialog keeps
 * only the list of them.
 */
export function EvaluateDialog({ load, initial, onClose }) {
  const [actions, setActions] = useState([]);
  const [state, setState] = useState(initial);
  const [busy, setBusy] = useState(false);
  const press = async (action) => {
    if (busy) return;
    const next = action === 'restart' ? [] : [...actions, action];
    setBusy(true);
    try {
      const got = await load(next);
      if (got) { setActions(next); setState(got); }
    } finally {
      setBusy(false);
    }
  };
  const levels = state?.levels || [];
  const formulaText = (level) => {
    // The = sits between the columns, as Excel draws it, not in the box.
    const lead = String(level.text || '').startsWith('=') ? 1 : 0;
    const t = String(level.text || '').slice(lead);
    const shift = (span) => (span ? [span[0] - lead, span[1] - lead] : null);
    const u = shift(level.underline);
    const r = shift(level.recent);
    // A stretch of the text with the most recent result in italics — the
    // result is often inside the part underlined next (SUM({5;3}) after
    // the range was read), so the underline wraps it rather than cutting it.
    const withRecent = (from, to) => {
      if (!r || r[1] <= from || r[0] >= to) return t.slice(from, to);
      const a = Math.max(from, r[0]);
      const b = Math.min(to, r[1]);
      return [t.slice(from, a), <em key="r" className="sh-eval-recent">{t.slice(a, b)}</em>, t.slice(b, to)];
    };
    if (!u) return withRecent(0, t.length);
    return [
      <React.Fragment key="a">{withRecent(0, u[0])}</React.Fragment>,
      <u key="u" className="sh-eval-next">{withRecent(u[0], u[1])}</u>,
      <React.Fragment key="b">{withRecent(u[1], t.length)}</React.Fragment>,
    ];
  };
  return (
    <Dialog
      title={t('Evaluate Formula')}
      width={620}
      onClose={onClose}
      actions={
        <>
          {state?.done
            ? <Button primary label={t('Restart')} className="sh-eval-restart" disabled={busy} onClick={() => press('restart')} />
            : <Button primary label={t('Evaluate')} className="sh-eval-evaluate" disabled={busy || !state?.canEvaluate} onClick={() => press('evaluate')} />}
          <Button label={t('Step In')} className="sh-eval-in" disabled={busy || !state?.canStepIn} onClick={() => press('stepIn')} />
          <Button label={t('Step Out')} className="sh-eval-out" disabled={busy || !state?.canStepOut} onClick={() => press('stepOut')} />
          <Button label={t('Close')} onClick={onClose} />
        </>
      }
    >
      <div className="sh-eval">
        <div className="sh-eval-head">
          <span>{t('Reference')}</span>
          <span>{t('Evaluation')}</span>
        </div>
        {levels.map((level, i) => (
          <div key={i} className={`sh-eval-level${i === levels.length - 1 ? ' current' : ''}`} data-level={i}>
            <span className="sh-eval-ref">{level.ref}</span>
            <span className="sh-eval-eq">=</span>
            <div className="sh-eval-text" data-text={level.text}>{formulaText(level)}</div>
          </div>
        ))}
        <p className="sh-eval-message">{state?.message}</p>
      </div>
    </Dialog>
  );
}

/* ── Review → Protect ────────────────────────────────────────────────────── */

/**
 * Protect Sheet and Protect Workbook, as Excel asks them: a password that
 * is optional, typed twice when there is one (Excel's Confirm Password),
 * and — for the workbook — what is locked: its structure. Windows is
 * Excel's own greyed choice on Windows too.
 */
export function ProtectDialog({ kind, onClose, onProtect }) {
  const [password, setPassword] = useState('');
  const [again, setAgain] = useState('');
  const mismatch = password !== '' && again !== '' && again !== password;
  const ok = password === '' || again === password;
  const submit = () => ok && onProtect({ password });
  const enter = (e) => { if (e.key === 'Enter') submit(); };
  return (
    <Dialog
      title={kind === 'workbook' ? t('Protect Structure and Windows') : t('Protect Sheet')}
      width={420}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} className="sh-protect-ok" disabled={!ok} onClick={submit} />
        </>
      }
    >
      <div className="ml-form">
        {kind === 'workbook' ? (
          <Field label={t('Protect workbook for')}>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <label style={CHECK_ROW}><input type="checkbox" checked readOnly /> {t('Structure — no sheet added, deleted, renamed, moved, hidden or shown')}</label>
              <label style={{ ...CHECK_ROW, color: 'var(--ink-3)' }} data-tip={t('Windows — locking the window\'s size and place is greyed in Excel on Windows too')}><input type="checkbox" disabled /> {t('Windows')}</label>
            </div>
          </Field>
        ) : (
          <p className="sh-protect-note">{t('Locked cells take no edits while the sheet is protected, except in the ranges Allow Edit Ranges leaves open.')}</p>
        )}
        <Field label={t('Password (optional)')} hint={t('Without one, anyone can take the protection off. A password cannot be recovered if it is lost — keep it somewhere safe.')}>
          <Input type="password" className="sh-protect-password" value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={enter} autoFocus />
        </Field>
        {password ? (
          <Field label={t('Reenter password to proceed')}>
            <Input type="password" className="sh-protect-again" value={again} onChange={(e) => setAgain(e.target.value)} onKeyDown={enter} />
          </Field>
        ) : null}
        {mismatch ? <p className="sh-protect-warn">{t('The two passwords are not the same.')}</p> : null}
      </div>
    </Dialog>
  );
}

/**
 * Unprotect Sheet, Unprotect Workbook and Unlock Range: one password box,
 * Excel's sentence above it, and Excel's own words when the password is wrong.
 */
export function PasswordDialog({ title, message, error = '', onClose, onSubmit }) {
  const [password, setPassword] = useState('');
  const submit = () => onSubmit(password);
  return (
    <Dialog
      title={title}
      width={420}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('OK')} className="sh-password-ok" onClick={submit} />
        </>
      }
    >
      <div className="ml-form">
        {message ? <p className="sh-protect-note">{message}</p> : null}
        <Field label={t('Password')}>
          <Input type="password" className="sh-password" value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} autoFocus />
        </Field>
        {error ? <p className="sh-protect-warn sh-password-error">{error}</p> : null}
      </div>
    </Dialog>
  );
}

/**
 * Review → Allow Edit Ranges, Excel's dialog: the ranges a protected sheet
 * leaves open, each with its title and cells, New, Modify and Delete —
 * greyed while the sheet is protected, as Excel greys them — and the
 * New Range / Modify Range form in place: Title, Refers to cells, Range
 * password (typed twice). Protect Sheet… hands over to that dialog.
 */
export function EditRangesDialog({ ranges = [], sheetProtected, selection = '', onClose, onSave, onDelete, onProtectSheet }) {
  const [picked, setPicked] = useState(ranges[0]?.title ?? null);
  const [form, setForm] = useState(null);
  const [error, setError] = useState('');
  useEffect(() => {
    if (picked === null || !ranges.some((r) => r.title === picked)) setPicked(ranges[0]?.title ?? null);
  }, [ranges, picked]);
  const startNew = () => {
    let n = 1;
    while (ranges.some((r) => r.title.toLowerCase() === 'range' + n)) n += 1;
    setError('');
    setForm({ was: null, title: 'Range' + n, ref: selection, password: '', again: '', keep: false }); // words-ok: the new range's name, matched against 'range' + n above
  };
  const startModify = () => {
    const r = ranges.find((x) => x.title === picked);
    if (!r) return;
    setError('');
    setForm({ was: r.title, title: r.title, ref: r.ref, password: '', again: '', keep: r.hasPassword });
  };
  const set = (patch) => setForm((f) => ({ ...f, ...patch }));
  const locked = sheetProtected ? t('The sheet is protected — unprotect it to change the ranges') : null;
  const save = async () => {
    if (!form) return;
    if (form.password && form.password !== form.again) { setError(t('The two passwords are not the same.')); return; }
    const said = await onSave({ was: form.was, title: form.title.trim(), ref: form.ref.trim(), password: form.password ? form.password : form.keep ? null : '' });
    if (said) { setError(said); return; }
    setPicked(form.title.trim());
    setForm(null);
  };
  if (form) {
    return (
      <Dialog
        title={form.was === null ? t('New Range') : t('Modify Range')}
        width={440}
        onClose={() => setForm(null)}
        actions={
          <>
            <Button label={t('Cancel')} onClick={() => setForm(null)} />
            <Button primary label={t('OK')} className="sh-range-ok" disabled={!form.title.trim() || !form.ref.trim()} onClick={save} />
          </>
        }
      >
        <div className="ml-form">
          <Field label={t('Title')}>
            <Input className="sh-range-title" value={form.title} onChange={(e) => set({ title: e.target.value })} autoFocus />
          </Field>
          <Field label={t('Refers to cells')} hint={t('One range or several, such as B2:D10 F2:F10.')}>
            <Input className="sh-range-ref" value={form.ref} onChange={(e) => set({ ref: e.target.value })} placeholder="B2:D10" />
          </Field>
          <Field label={t('Range password')} hint={form.keep ? t('Leave it empty to keep the password this range has.') : t('Optional: without one, anyone can edit the range when the sheet is protected.')}>
            <Input type="password" className="sh-range-password" value={form.password} onChange={(e) => set({ password: e.target.value })} />
          </Field>
          {form.password ? (
            <Field label={t('Reenter password to proceed')}>
              <Input type="password" className="sh-range-again" value={form.again} onChange={(e) => set({ again: e.target.value })} />
            </Field>
          ) : null}
          {form.was !== null && ranges.find((r) => r.title === form.was)?.hasPassword && !form.password ? (
            <label style={CHECK_ROW}><input type="checkbox" className="sh-range-clear" checked={!form.keep} onChange={(e) => set({ keep: !e.target.checked })} /> {t('Take the password off')}</label>
          ) : null}
          {error ? <p className="sh-protect-warn">{error}</p> : null}
        </div>
      </Dialog>
    );
  }
  return (
    <Dialog
      title={t('Allow Users to Edit Ranges')}
      width={540}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Protect Sheet…')} className="sh-ranges-protect" disabled={sheetProtected} title={sheetProtected ? t('Protect Sheet — the sheet is protected already') : t('Protect Sheet — so these ranges are the only locked cells open')} onClick={onProtectSheet} />
          <span style={{ flex: 1 }} />
          <Button primary label={t('OK')} className="sh-ranges-close" onClick={onClose} />
        </>
      }
    >
      <div className="sh-ranges">
        <p className="sh-protect-note">{t('Ranges unlocked by a password when the sheet is protected:')}</p>
        <div className="sh-ranges-body">
          <div className="sh-ranges-list" role="listbox">
            <div className="sh-ranges-head"><span>{t('Title')}</span><span>{t('Refers to cells')}</span></div>
            {ranges.length ? ranges.map((r) => (
              <button
                key={r.title}
                type="button"
                role="option"
                aria-selected={r.title === picked}
                className={`sh-ranges-row${r.title === picked ? ' on' : ''}`}
                data-title={r.title}
                onClick={() => setPicked(r.title)}
                onDoubleClick={() => { setPicked(r.title); if (!sheetProtected) startModify(); }}
              >
                <span className="t">{r.hasPassword ? <Icon name="lock" size={12} /> : null}{r.title}</span>
                <span className="c">{r.ref}</span>
              </button>
            )) : <div className="sh-ranges-empty">{t('No ranges yet — New… adds one from the selection.')}</div>}
          </div>
          <div className="sh-ranges-buttons">
            <Button label={t('New…')} className="sh-ranges-new" disabled={sheetProtected} title={locked || t('New… — a range from the selection, with a title and an optional password')} onClick={startNew} />
            <Button label={t('Modify…')} className="sh-ranges-modify" disabled={sheetProtected || !picked} title={locked || t('Modify… — the picked range')} onClick={startModify} />
            <Button label={t('Delete')} className="sh-ranges-delete" disabled={sheetProtected || !picked} title={locked || t('Delete — the picked range')} onClick={() => picked && onDelete(picked)} />
          </div>
        </div>
        {sheetProtected ? <p className="sh-ranges-locked"><Icon name="lock" size={12} /> {t('The sheet is protected: the ranges stay as they are until it is unprotected.')}</p> : null}
      </div>
    </Dialog>
  );
}

/* ── View → Custom Views ─────────────────────────────────────────────────── */

/**
 * Excel's Custom Views dialog: the views kept in the workbook, Show, Close,
 * Add… and Delete; Add asks a name and whether the view keeps the print
 * settings and the hidden rows, columns and filter settings.
 */
export function CustomViewsDialog({ views = [], onClose, onShow, onAdd, onDelete }) {
  const [picked, setPicked] = useState(views[0]?.name ?? null);
  const [adding, setAdding] = useState(null);
  const [error, setError] = useState('');
  useEffect(() => {
    if (picked === null || !views.some((v) => v.name === picked)) setPicked(views[0]?.name ?? null);
  }, [views, picked]);
  const add = async () => {
    const said = await onAdd({ name: adding.name.trim(), printSettings: adding.print, hiddenRowCol: adding.hidden });
    if (said) { setError(said); return; }
    setPicked(adding.name.trim());
    setAdding(null);
  };
  if (adding) {
    return (
      <Dialog
        title={t('Add View')}
        width={420}
        onClose={() => setAdding(null)}
        actions={
          <>
            <Button label={t('Cancel')} onClick={() => setAdding(null)} />
            <Button primary label={t('OK')} className="sh-cview-ok" disabled={!adding.name.trim()} onClick={add} />
          </>
        }
      >
        <div className="ml-form">
          <Field label={t('Name')}>
            <Input className="sh-cview-name" value={adding.name} onChange={(e) => setAdding((a) => ({ ...a, name: e.target.value }))} onKeyDown={(e) => { if (e.key === 'Enter' && adding.name.trim()) add(); }} autoFocus />
          </Field>
          <Field label={t('Include in view')}>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <label style={CHECK_ROW}><input type="checkbox" className="sh-cview-print" checked={adding.print} onChange={(e) => setAdding((a) => ({ ...a, print: e.target.checked }))} /> {t('Print settings')}</label>
              <label style={CHECK_ROW}><input type="checkbox" className="sh-cview-hidden" checked={adding.hidden} onChange={(e) => setAdding((a) => ({ ...a, hidden: e.target.checked }))} /> {t('Hidden rows, columns and filter settings')}</label>
            </div>
          </Field>
          {views.some((v) => v.name.toLowerCase() === adding.name.trim().toLowerCase())
            ? <p className="sh-protect-note">{t('A view called “{name}” exists — OK replaces it.', { name: adding.name.trim() })}</p> : null}
          {error ? <p className="sh-protect-warn">{error}</p> : null}
        </div>
      </Dialog>
    );
  }
  return (
    <Dialog
      title={t('Custom Views')}
      width={500}
      onClose={onClose}
      actions={<Button label={t('Close')} className="sh-cviews-close" onClick={onClose} />}
    >
      <div className="sh-ranges">
        <div className="sh-ranges-body">
          <div className="sh-ranges-list sh-cviews-list" role="listbox">
            <div className="sh-ranges-head"><span>{t('Views')}</span><span>{t('Keeps')}</span></div>
            {views.length ? views.map((v) => (
              <button
                key={v.name}
                type="button"
                role="option"
                aria-selected={v.name === picked}
                className={`sh-ranges-row${v.name === picked ? ' on' : ''}`}
                data-view={v.name}
                onClick={() => setPicked(v.name)}
                onDoubleClick={() => onShow(v.name)}
              >
                <span className="t">{v.name}</span>
                <span className="c">{[v.printSettings ? t('print settings') : null, v.hiddenRowCol ? t('hidden rows') : null].filter(Boolean).join(', ') || t('zoom and selection')}</span>
              </button>
            )) : <div className="sh-ranges-empty">{t('No views yet — Add… keeps the way the workbook looks now.')}</div>}
          </div>
          <div className="sh-ranges-buttons">
            <Button primary label={t('Show')} className="sh-cviews-show" disabled={!picked} title={t('Show — the workbook as the picked view kept it')} onClick={() => picked && onShow(picked)} />
            <Button label={t('Add…')} className="sh-cviews-add" title={t('Add… — keep the way the workbook looks now under a name')} onClick={() => { setError(''); setAdding({ name: '', print: true, hidden: true }); }} />
            <Button label={t('Delete')} className="sh-cviews-delete" disabled={!picked} title={t('Delete — the picked view')} onClick={() => picked && onDelete(picked)} />
          </div>
        </div>
      </div>
    </Dialog>
  );
}

/* ── Data → Consolidate ──────────────────────────────────────────────────── */

export const CONSOLIDATE_FNS = [
  ['sum', t('Sum')], ['count', t('Count')], ['average', t('Average')], ['max', t('Max')], ['min', t('Min')], ['product', t('Product')],
  ['countNumbers', t('Count Numbers')], ['stdDev', t('StdDev')], ['stdDevp', t('StdDevp')], ['var', t('Var')], ['varp', t('Varp')],
];

/**
 * Excel's Consolidate dialog: the function, a reference typed (or the
 * selection, offered as the dialog opens) and added to All references,
 * Delete, labels in the top row and the left column, links to the source
 * data. It opens on the sheet's last consolidation, as Excel's does.
 */
export function ConsolidateDialog({ info, onClose, onApply }) {
  const last = info?.last || null;
  const [fn, setFn] = useState(last?.fn || 'sum');
  const [refs, setRefs] = useState(last?.refs || []);
  const [typed, setTyped] = useState(info?.selection || '');
  const [picked, setPicked] = useState(null);
  const [top, setTop] = useState(Boolean(last?.topRow));
  const [left, setLeft] = useState(Boolean(last?.leftCol));
  const [links, setLinks] = useState(Boolean(last?.links));
  const [error, setError] = useState('');
  const add = () => {
    const ref = typed.trim();
    if (!ref) return;
    if (!/^(?:'[^']+'|[^'!]+)?!?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/.test(ref)) { setError(t('"{ref}" is not a reference, such as Sheet1!$A$1:$C$5.', { ref })); return; }
    setError('');
    if (!refs.some((r) => r.toLowerCase() === ref.toLowerCase())) setRefs((list) => [...list, ref]);
    setPicked(ref);
  };
  const apply = async () => {
    const said = await onApply({ fn, refs, topRow: top, leftCol: left, links });
    if (said) setError(said);
  };
  return (
    <Dialog
      title={t('Consolidate')}
      width={500}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Close')} onClick={onClose} />
          <Button primary label={t('OK')} className="sh-cons-ok" disabled={!refs.length} onClick={apply} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={t('Function')}>
          <Select className="sh-cons-fn" value={fn} onChange={(e) => setFn(e.target.value)} style={{ width: '100%' }}>
            {CONSOLIDATE_FNS.map(([k, name]) => <option key={k} value={k}>{name}</option>)}
          </Select>
        </Field>
        <Field label={t('Reference')} hint={t('A range on this sheet or another, such as East!$A$1:$C$10 — Add puts it in the list.')}>
          <div style={{ display: 'flex', gap: 6 }}>
            <Input className="sh-cons-ref" value={typed} onChange={(e) => setTyped(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') add(); }} style={{ flex: 1 }} autoFocus />
            <Button label={t('Add')} className="sh-cons-add" onClick={add} />
          </div>
        </Field>
        <Field label={t('All references')}>
          <div className="sh-ranges-body">
            <div className="sh-ranges-list sh-cons-list" role="listbox" style={{ minHeight: 96, maxHeight: 150 }}>
              {refs.length ? refs.map((r) => (
                <button key={r} type="button" role="option" aria-selected={r === picked} className={`sh-ranges-row one${r === picked ? ' on' : ''}`} data-ref={r} onClick={() => { setPicked(r); setTyped(r); }}>
                  <span className="t">{r}</span>
                </button>
              )) : <div className="sh-ranges-empty">{t('No references yet.')}</div>}
            </div>
            <div className="sh-ranges-buttons">
              <Button label={t('Delete')} className="sh-cons-delete" disabled={!picked} title={t('Delete — the picked reference')} onClick={() => { setRefs((list) => list.filter((r) => r !== picked)); setPicked(null); }} />
            </div>
          </div>
        </Field>
        <div className="sh-cons-options">
          <div>
            <div className="sh-cons-head">{t('Use labels in')}</div>
            <label style={CHECK_ROW}><input type="checkbox" className="sh-cons-top" checked={top} onChange={(e) => setTop(e.target.checked)} /> {t('Top row')}</label>
            <label style={CHECK_ROW}><input type="checkbox" className="sh-cons-left" checked={left} onChange={(e) => setLeft(e.target.checked)} /> {t('Left column')}</label>
          </div>
          <label style={CHECK_ROW}><input type="checkbox" className="sh-cons-links" checked={links} onChange={(e) => setLinks(e.target.checked)} /> {t('Create links to source data')}</label>
        </div>
        {error ? <p className="sh-protect-warn">{error}</p> : null}
      </div>
    </Dialog>
  );
}

/* ── Data → Forecast Sheet ───────────────────────────────────────────────── */

const FORECAST_AGGREGATES = [[1, t('Average')], [2, t('Count')], [3, t('CountA')], [4, t('Max')], [5, t('Median')], [6, t('Min')], [7, t('Sum')]];

/** A date serial as the yyyy-mm-dd a date field takes, and back. */
const serialToIso = (s) => new Date(Math.round((s - 25569) * 86400000)).toISOString().slice(0, 10);
const isoToSerial = (iso) => Date.parse(iso + 'T00:00:00Z') / 86400000 + 25569;

/**
 * Excel's Create Forecast Worksheet dialog: a preview of the forecast drawn
 * as it will be, Line or Column, Forecast End, and the Options — the
 * confidence interval, seasonality detected or set, the timeline and values
 * ranges, how missing points are filled and duplicates aggregated.
 */
export function ForecastDialog({ info, preview, onClose, onCreate }) {
  const [kind, setKind] = useState('line');
  const [end, setEnd] = useState(info?.end ?? '');
  const [confOn, setConfOn] = useState(true);
  const [conf, setConf] = useState('95');
  const [autoSeason, setAutoSeason] = useState(true);
  const [season, setSeason] = useState('12');
  const [timeline, setTimeline] = useState(info?.timeline || '');
  const [values, setValues] = useState(info?.values || '');
  const [completion, setCompletion] = useState(1);
  const [aggregation, setAggregation] = useState(1);
  const [options, setOptions] = useState(true);
  const [shown, setShown] = useState({ svg: '', error: '' });
  const [error, setError] = useState('');
  const spec = () => ({
    timeline: timeline.trim(), values: values.trim(), end: Number(end), kind,
    confidence: confOn ? Number(conf) / 100 : null,
    seasonality: autoSeason ? 'auto' : Number(season), completion, aggregation,
  });
  useEffect(() => {
    let live = true;
    const t = setTimeout(async () => {
      const got = await preview(spec());
      if (live) setShown(got || { svg: '', error: '' });
    }, 160);
    return () => { live = false; clearTimeout(t); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [kind, end, confOn, conf, autoSeason, season, timeline, values, completion, aggregation]);
  const create = async () => {
    const said = await onCreate(spec());
    if (said) setError(said);
  };
  return (
    <Dialog
      title={t('Create Forecast Worksheet')}
      width={620}
      onClose={onClose}
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary label={t('Create')} className="sh-fc-create" onClick={create} />
        </>
      }
    >
      <div className="sh-fc">
        <div className="sh-fc-kinds">
          <button type="button" className={`sh-fc-kind${kind === 'line' ? ' on' : ''}`} data-kind="line" data-tip={t('Line chart')} onClick={() => setKind('line')}>
            <svg width="22" height="16" viewBox="0 0 22 16"><path d="M1 13 L6 8 L10 10 L15 4 L21 6" /></svg>
          </button>
          <button type="button" className={`sh-fc-kind${kind === 'column' ? ' on' : ''}`} data-kind="column" data-tip={t('Column chart')} onClick={() => setKind('column')}>
            <svg width="22" height="16" viewBox="0 0 22 16"><path d="M3 15V8M8 15V4M13 15V9M18 15V2" /></svg>
          </button>
        </div>
        <div className="sh-fc-preview" data-ready={shown.svg ? '1' : '0'}>
          {shown.svg ? <div dangerouslySetInnerHTML={{ __html: shown.svg }} /> : <p>{shown.error || t('Working out the forecast…')}</p>}
        </div>
        <div className="sh-fc-row">
          <Field label={t('Forecast End')}>
            {info?.isDate
              ? <Input type="date" className="sh-fc-end" value={end ? serialToIso(Number(end)) : ''} onChange={(e) => setEnd(e.target.value ? isoToSerial(e.target.value) : '')} />
              : <Input className="sh-fc-end" value={String(end)} onChange={(e) => setEnd(e.target.value)} />}
          </Field>
          <button type="button" className={`sh-fc-toggle${options ? ' open' : ''}`} onClick={() => setOptions((o) => !o)}>{options ? t('Options ▴') : t('Options ▾')}</button>
        </div>
        {options ? (
          <div className="sh-fc-options">
            <div className="sh-fc-col">
              <label style={CHECK_ROW}>
                <input type="checkbox" className="sh-fc-conf-on" checked={confOn} onChange={(e) => setConfOn(e.target.checked)} /> {t('Confidence Interval')}
                <Input className="sh-fc-conf" value={conf} disabled={!confOn} onChange={(e) => setConf(e.target.value)} style={{ width: 54, marginLeft: 6 }} />%
              </label>
              <div className="sh-cons-head">{t('Seasonality')}</div>
              <label style={CHECK_ROW}><input type="radio" name="sh-fc-season" className="sh-fc-auto" checked={autoSeason} onChange={() => setAutoSeason(true)} /> {t('Detect Automatically')}</label>
              <label style={CHECK_ROW}>
                <input type="radio" name="sh-fc-season" className="sh-fc-manual" checked={!autoSeason} onChange={() => setAutoSeason(false)} /> {t('Set Manually')}
                <Input className="sh-fc-season" value={season} disabled={autoSeason} onChange={(e) => setSeason(e.target.value)} style={{ width: 54, marginLeft: 6 }} />
              </label>
            </div>
            <div className="sh-fc-col">
              <Field label={t('Timeline Range')}><Input className="sh-fc-timeline" value={timeline} onChange={(e) => setTimeline(e.target.value)} /></Field>
              <Field label={t('Values Range')}><Input className="sh-fc-values" value={values} onChange={(e) => setValues(e.target.value)} /></Field>
              <Field label={t('Fill Missing Points Using')}>
                <Select className="sh-fc-fill" value={String(completion)} onChange={(e) => setCompletion(Number(e.target.value))} style={{ width: '100%' }}>
                  <option value="1">{t('Interpolation')}</option>
                  <option value="0">{t('Zeros')}</option>
                </Select>
              </Field>
              <Field label={t('Aggregate Duplicates Using')}>
                <Select className="sh-fc-agg" value={String(aggregation)} onChange={(e) => setAggregation(Number(e.target.value))} style={{ width: '100%' }}>
                  {FORECAST_AGGREGATES.map(([k, name]) => <option key={k} value={String(k)}>{name}</option>)}
                </Select>
              </Field>
            </div>
          </div>
        ) : null}
        {error ? <p className="sh-protect-warn">{error}</p> : null}
      </div>
    </Dialog>
  );
}
