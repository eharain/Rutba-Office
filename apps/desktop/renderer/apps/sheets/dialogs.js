// The dialogs behind the Worksheets ribbon.
//
// Each one drives an engine capability that already existed and had no way in:
// conditional formatting, validation, goal seek, a what-if table, named ranges
// and find-and-replace. They are deliberately plain — a spreadsheet's power is
// in the grid, and a dialog's job is to get out of the way.

import React, { useEffect, useState } from 'react';
import { Button, Dialog, Field, Input, Select, Icon, Chip, Empty } from '@rutba/office-ui';

/* ── number formats ──────────────────────────────────────────────────────── */

/**
 * The formats a person actually reaches for, with the codes a workbook stores.
 * Excel's own list is longer and mostly regional variants of these.
 */
export const NUMBER_FORMATS = [
  { label: 'General', code: 'General' },
  { label: 'Number', code: '#,##0.00' },
  { label: 'Number, no decimals', code: '#,##0' },
  { label: 'Currency', code: '"£"#,##0.00' },
  { label: 'Accounting', code: '_-"£"* #,##0.00_-;-"£"* #,##0.00_-;_-"£"* "-"??_-;_-@_-' },
  { label: 'Percentage', code: '0.00%' },
  { label: 'Percentage, no decimals', code: '0%' },
  { label: 'Scientific', code: '0.00E+00' },
  { label: 'Fraction', code: '# ?/?' },
  { label: 'Short date', code: 'dd/mm/yyyy' },
  { label: 'Long date', code: 'dddd, d mmmm yyyy' },
  { label: 'Time', code: 'hh:mm:ss' },
  { label: 'Date and time', code: 'dd/mm/yyyy hh:mm' },
  { label: 'Duration', code: '[h]:mm:ss' },
  { label: 'Text', code: '@' },
];

/* ── conditional formatting ──────────────────────────────────────────────── */

const CONDITIONS = [
  { id: 'greaterThan', label: 'Greater than', operands: 1 },
  { id: 'lessThan', label: 'Less than', operands: 1 },
  { id: 'between', label: 'Between', operands: 2 },
  { id: 'equal', label: 'Equal to', operands: 1 },
  { id: 'notEqual', label: 'Not equal to', operands: 1 },
  { id: 'containsText', label: 'Text contains', operands: 1, text: true },
  { id: 'duplicateValues', label: 'Duplicate values', operands: 0 },
  { id: 'uniqueValues', label: 'Unique values', operands: 0 },
];

const HIGHLIGHTS = [
  { label: 'Red text on light red', colour: '9C0006', fill: 'FFC7CE' },
  { label: 'Amber text on light amber', colour: '9C6500', fill: 'FFEB9C' },
  { label: 'Green text on light green', colour: '006100', fill: 'C6EFCE' },
  { label: 'Bold only', colour: null, fill: null, bold: true },
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
      title="Conditional formatting"
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label="Clear rules" onClick={() => onClear(true)} />
          <Button label="Cancel" onClick={onClose} />
          <Button
            primary
            label="Apply"
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
        <Field label="Apply to" hint="The cells selected in the grid.">
          <Input value={selection || ''} disabled />
        </Field>
        <Field label="Format cells where the value is">
          <Select value={rule} onChange={(e) => setRule(e.target.value)}>
            {CONDITIONS.map((c) => (
              <option key={c.id} value={c.id}>{c.label}</option>
            ))}
          </Select>
        </Field>
        {chosen.operands > 0 ? (
          <div className="ml-servers" style={{ gridTemplateColumns: chosen.operands === 2 ? '1fr 1fr' : '1fr' }}>
            <Field label={chosen.operands === 2 ? 'From' : 'Value'}>
              <Input value={a} onChange={(e) => setA(e.target.value)} placeholder={chosen.text ? 'text' : '0'} autoFocus />
            </Field>
            {chosen.operands === 2 ? (
              <Field label="To">
                <Input value={b} onChange={(e) => setB(e.target.value)} placeholder="100" />
              </Field>
            ) : null}
          </div>
        ) : null}
        <Field label="Show them as">
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
      title="Data validation"
      width={520}
      onClose={onClose}
      actions={
        <>
          <Button label="Clear" onClick={() => onClear(false)} />
          <Button label="Cancel" onClick={onClose} />
          <Button
            primary
            label="Apply"
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
        <Field label="Apply to">
          <Input value={selection || ''} disabled />
        </Field>
        <Field label="Allow">
          <Select value={type} onChange={(e) => setType(e.target.value)}>
            <option value="list">A list of values</option>
            <option value="whole">Whole numbers</option>
            <option value="decimal">Decimal numbers</option>
            <option value="date">Dates</option>
            <option value="textLength">Text of a given length</option>
          </Select>
        </Field>
        {type === 'list' ? (
          <Field label="Values" hint="Separated by commas, or a range such as $H$1:$H$9.">
            <Input value={list} onChange={(e) => setList(e.target.value)} placeholder="Draft, In review, Approved" autoFocus />
          </Field>
        ) : (
          <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr' }}>
            <Field label="Minimum">
              <Input value={min} onChange={(e) => setMin(e.target.value)} placeholder="0" autoFocus />
            </Field>
            <Field label="Maximum">
              <Input value={max} onChange={(e) => setMax(e.target.value)} placeholder="100" />
            </Field>
          </div>
        )}
        <Field label="Message when the cell is selected">
          <Input value={message} onChange={(e) => setMessage(e.target.value)} placeholder="Optional" />
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
      title="Goal seek"
      width={460}
      onClose={onClose}
      actions={
        <>
          <Button label="Close" onClick={onClose} />
          <Button
            primary
            label="Find"
            disabled={!set || !to || !by}
            onClick={async () => setResult(await onRun({ set, to: Number(to), by }))}
          />
        </>
      }
    >
      <p style={{ marginTop: 0, fontSize: 12.5 }}>
        Work backwards: name a formula cell, the answer you want from it, and the cell that may change to get there.
      </p>
      <div className="ml-form">
        <Field label="Set cell">
          <Input value={set} onChange={(e) => setSet(e.target.value.toUpperCase())} placeholder="B10" autoFocus />
        </Field>
        <Field label="To value">
          <Input value={to} onChange={(e) => setTo(e.target.value)} placeholder="50000" />
        </Field>
        <Field label="By changing cell">
          <Input value={by} onChange={(e) => setBy(e.target.value.toUpperCase())} placeholder="B4" />
        </Field>
        {result ? (
          <div className={`ml-result ${result.converged ? 'good' : 'bad'}`}>
            <Icon name={result.converged ? 'check' : 'info'} size={14} />
            <span>
              {result.converged
                ? `${by} is ${result.value} — ${set} reaches ${result.reached}`
                : `No value of ${by} gets ${set} there. Closest: ${result.reached ?? '—'}`}
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
      title="What-if table"
      width={480}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Fill" disabled={!rowInput && !colInput} onClick={() => onRun({ range: selection, rowInput, colInput })} />
        </>
      }
    >
      <p style={{ marginTop: 0, fontSize: 12.5 }}>
        The selection holds a formula in its corner and the values to try along its edges. Each cell is filled with what
        the formula gives for that combination.
      </p>
      <div className="ml-form">
        <Field label="Range">
          <Input value={selection || ''} disabled />
        </Field>
        <Field label="Row input cell" hint="The cell the values along the top belong to.">
          <Input value={rowInput} onChange={(e) => setRowInput(e.target.value.toUpperCase())} placeholder="B3" autoFocus />
        </Field>
        <Field label="Column input cell" hint="The cell the values down the side belong to.">
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
      title="Named ranges"
      width={540}
      onClose={onClose}
      actions={
        <>
          <Button label="Close" onClick={onClose} />
          <Button primary label="Define" disabled={!name || !ref} onClick={() => { onDefine(name, ref); setName(''); }} />
        </>
      }
    >
      <div className="ml-form">
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1.4fr' }}>
          <Field label="Name">
            <Input value={name} onChange={(e) => setName(e.target.value.replace(/\s/g, '_'))} placeholder="Revenue" autoFocus />
          </Field>
          <Field label="Refers to">
            <Input value={ref} onChange={(e) => setRef(e.target.value)} placeholder="Sheet1!$B$2:$B$40" />
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
              <Button icon="trash" title={`Delete ${n.name}`} onClick={() => onDelete(n.name)} />
            </div>
          ))}
        </div>
      ) : (
        <Empty icon="find" title="No names yet">
          A name turns <code>$B$2:$B$40</code> into <code>Revenue</code>, in every formula that uses it.
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
      title="Find and replace"
      width={480}
      onClose={onClose}
      actions={
        <>
          <Button label="Close" onClick={onClose} />
          <Button label="Find next" disabled={!find} onClick={async () => setNote(await onFind(find))} />
          <Button label="Replace" disabled={!find} onClick={async () => setNote(await onReplace(find, replace))} />
          <Button primary label="Replace all" disabled={!find} onClick={async () => setNote(await onReplaceAll(find, replace))} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Find">
          <Input value={find} onChange={(e) => setFind(e.target.value)} autoFocus />
        </Field>
        <Field label="Replace with">
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
      title="Go To"
      width={380}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Go" disabled={!ok} onClick={() => onGo(text.trim())} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Reference" hint="A cell such as C12, or a defined name.">
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
      title={current ? `The link on ${cellRef}` : `A link on ${cellRef}`}
      width={440}
      onClose={onClose}
      actions={
        <>
          {current ? <Button label="Remove link" className="sh-link-remove" onClick={onRemove} /> : null}
          <Button label="Cancel" onClick={onClose} />
          <Button primary label={current ? 'Change' : 'Add link'} className="sh-link-ok" disabled={!ok} onClick={submit} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Link to" hint="An address such as rutba.io or hello@rutba.io, or a place in this workbook such as C12, Sheet2!B4 or a name.">
          <Input className="sh-link-to" value={to} onChange={(e) => setTo(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} autoFocus placeholder="https://…  or  Sheet2!B4" />
        </Field>
        <Field label="Screen tip" hint="Shown when the pointer rests on the cell. Optional.">
          <Input className="sh-link-tip" value={tip} onChange={(e) => setTip(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} placeholder="Where this goes" />
        </Field>
        <p className="sh-link-says" style={{ margin: 0, fontSize: 12, color: 'var(--ink-3)' }}>
          {!value ? 'Nothing yet.' : address ? `Opens ${address} outside the suite.` : place ? `Goes to ${value} in this workbook.` : 'Not an address and not a place this workbook knows.'}
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
      title={current ? `The note on ${cellRef}` : `A note on ${cellRef}`}
      width={440}
      onClose={onClose}
      actions={
        <>
          {current ? <Button label="Delete note" className="sh-note-remove" onClick={onRemove} /> : null}
          <Button label="Cancel" onClick={onClose} />
          <Button primary label={current ? 'Change' : 'Add note'} className="sh-note-ok" disabled={!ok} onClick={submit} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Note" hint="Shown when the pointer rests on the cell; Excel shows it the same way.">
          <textarea
            className="rw-input sh-note-text"
            rows={5}
            value={text}
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) submit(); }}
            autoFocus
            placeholder="What to say about this cell"
            style={{ resize: 'vertical', minHeight: 90, font: 'inherit', width: '100%', boxSizing: 'border-box' }}
          />
        </Field>
        <Field label="Author" hint="Who signs it. Left empty, the note is signed with your account's name.">
          <Input className="sh-note-author" value={author} onChange={(e) => setAuthor(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} placeholder="Your name" />
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
  const picks = [['Page &P of &N', 'Page &P of &N'], ['Sheet name', '&A'], ['File name', '&F'], ['Date', '&D'], ['Left, centre, right', '&L&F&C&A&R&D']];
  return (
    <Dialog
      title="Header and footer"
      width={480}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Set" className="sh-hf-ok" onClick={submit} />
        </>
      }
    >
      <div className="ml-form">
        <Field label="Header" hint={'Printed at the top of every page. Codes: &P page, &N pages, &A sheet, &F file, &D date, &T time; &L, &C and &R start the left, centre and right parts.'}>
          <Input className="sh-hf-header" value={header} onChange={(e) => setHeader(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} placeholder="Nothing at the top" autoFocus />
        </Field>
        <Field label="Footer" hint={'Printed at the foot of every page; empty prints nothing there.'}>
          <Input className="sh-hf-footer" value={footer} onChange={(e) => setFooter(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') submit(); }} placeholder="Nothing at the foot" />
        </Field>
        <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
          {picks.map(([label, code]) => (
            <Button key={code} label={label} title={`Put ${code} in the footer`} onClick={() => setFooter((f) => (f ? f + ' ' : '') + code)} />
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
    <Dialog title="Insert function" width={480} onClose={onClose} actions={<Button label="Cancel" onClick={onClose} />}>
      <div className="ml-form">
        <Field label="Search">
          <Input value={filter} onChange={(e) => setFilter(e.target.value)} placeholder="Type part of a name" autoFocus />
        </Field>
        {filter.trim() ? null : (
          <Field label="Category">
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
          {names.length ? null : <Empty title="No function by that name" />}
        </div>
      </div>
    </Dialog>
  );
}

/** Workbook Statistics: what the model knows, as Excel's dialog lists it. */
export function StatisticsDialog({ model, onClose }) {
  const rows = [
    ['Sheets', (model?.sheets || []).length],
    ['Active sheet', model?.activeSheet || '—'],
    ['Cells with content', (model?.cells || []).filter((c) => c.text !== '' && c.text != null).length + (model?.cells ? ' (on screen)' : '')],
    ['Formulas', (model?.cells || []).filter((c) => c.isFormula).length + ' (on screen)'],
    ['Defined names', (model?.names || []).length],
    ['Tables', (model?.tables || []).length],
    ['Frozen', model?.frozen && (model.frozen.rows || model.frozen.cols) ? `${model.frozen.rows} rows, ${model.frozen.cols} columns` : 'nothing'],
    ['Filter', model?.filtered ? 'on' : 'off'],
    ['Protection', model?.protection?.sheet ? 'sheet protected' : 'none'],
  ];
  return (
    <Dialog title="Workbook statistics" width={420} onClose={onClose} actions={<Button primary label="Close" onClick={onClose} />}>
      <dl className="about-list">
        {rows.map(([k, v]) => (
          <React.Fragment key={k}><dt>{k}</dt><dd>{String(v)}</dd></React.Fragment>
        ))}
      </dl>
    </Dialog>
  );
}

const SHEET_SHORTCUTS = [
  ['Ctrl+S', 'Save'], ['Ctrl+Z / Ctrl+Y', 'Undo / Redo'], ['Ctrl+C / Ctrl+X / Ctrl+V', 'Copy / Cut / Paste'],
  ['Ctrl+B / Ctrl+I / Ctrl+U', 'Bold / Italic / Underline'], ['F2', 'Edit the active cell'], ['Enter / Tab', 'Commit and move down / right'],
  ['Escape', 'Cancel the edit'], ['Delete', 'Clear the selection'], ['Shift+Arrows', 'Extend the selection'],
  ['Ctrl+Arrows', 'Jump to the edge of the data'], ['Ctrl+A', 'Select all'], ['Ctrl+D / Ctrl+R', 'Fill down / right'],
  ['Ctrl+F / Ctrl+H', 'Find / Replace'], ['Ctrl+G', 'Go To'], ['Ctrl+`', 'Show formulas'],
];

export function SheetShortcutsDialog({ onClose }) {
  return (
    <Dialog title="Keyboard shortcuts" width={440} onClose={onClose} actions={<Button primary label="Close" onClick={onClose} />}>
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
      title="Sort"
      width={440}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="Sort" className="sh-sort-ok" disabled={!ok} onClick={() => onSort(keys)} />
        </>
      }
    >
      <div className="ml-form">
        {keys.map((k, i) => (
          <div key={i} style={{ display: 'flex', gap: 8, alignItems: 'end' }}>
            <Field label={i === 0 ? 'Sort by' : 'Then by'} style={{ flex: 1 }}>
              <Select className={`sh-sort-col-${i}`} value={String(k.col)} onChange={(e) => set(i, { col: Number(e.target.value) })} style={{ width: '100%' }}>
                {columns.map((c) => <option key={c.col} value={String(c.col)}>{c.name}</option>)}
              </Select>
            </Field>
            <Field label="Order">
              <Select className={`sh-sort-dir-${i}`} value={k.ascending ? 'asc' : 'desc'} onChange={(e) => set(i, { ascending: e.target.value === 'asc' })} style={{ width: 150 }}>
                <option value="asc">A to Z, small to large</option>
                <option value="desc">Z to A, large to small</option>
              </Select>
            </Field>
            {i > 0 ? <Button icon="close" title="Remove this level" onClick={() => setKeys((ks) => ks.filter((_, j) => j !== i))} /> : null}
          </div>
        ))}
        {keys.length < 3 ? <Button label="Add a level" onClick={() => setKeys((ks) => [...ks, { col: columns[Math.min(ks.length, columns.length - 1)]?.col ?? first, ascending: true }])} /> : null}
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
      title={kind === 'row' ? 'Row height' : 'Column width'}
      width={340}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary label="OK" disabled={!ok} onClick={() => onApply(n)} />
        </>
      }
    >
      <div className="ml-form">
        <Field label={kind === 'row' ? 'Height (px)' : 'Width (px)'} hint="Applies to every selected row or column.">
          <Input value={value} onChange={(e) => setValue(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && ok) onApply(n); }} autoFocus />
        </Field>
      </div>
    </Dialog>
  );
}

export function PivotDialog({ onClose, onCreate, selection, sheets }) {
  const [source, setSource] = useState(selection || '');
  const [rows, setRows] = useState('');
  const [cols, setCols] = useState('');
  const [values, setValues] = useState('');
  const [fn, setFn] = useState('SUM');

  return (
    <Dialog
      title="Pivot table"
      width={540}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button
            primary
            label="Create"
            disabled={!source || !values}
            onClick={() =>
              onCreate({
                source,
                rowFields: rows.split(',').map((s) => s.trim()).filter(Boolean),
                colFields: cols.split(',').map((s) => s.trim()).filter(Boolean),
                dataFields: values.split(',').map((s) => ({ name: s.trim(), fn })).filter((d) => d.name),
              })
            }
          />
        </>
      }
    >
      <p style={{ marginTop: 0, fontSize: 12.5 }}>
        Field names are the column headings of the source range.
      </p>
      <div className="ml-form">
        <Field label="Source range">
          <Input value={source} onChange={(e) => setSource(e.target.value)} placeholder="Sheet1!$A$1:$E$500" autoFocus />
        </Field>
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 1fr' }}>
          <Field label="Rows">
            <Input value={rows} onChange={(e) => setRows(e.target.value)} placeholder="Region" />
          </Field>
          <Field label="Columns">
            <Input value={cols} onChange={(e) => setCols(e.target.value)} placeholder="Quarter" />
          </Field>
        </div>
        <div className="ml-servers" style={{ gridTemplateColumns: '1fr 140px' }}>
          <Field label="Values">
            <Input value={values} onChange={(e) => setValues(e.target.value)} placeholder="Revenue" />
          </Field>
          <Field label="Summarise by">
            <Select value={fn} onChange={(e) => setFn(e.target.value)}>
              {['SUM', 'COUNT', 'AVERAGE', 'MAX', 'MIN'].map((f) => (
                <option key={f} value={f}>{f}</option>
              ))}
            </Select>
          </Field>
        </div>
      </div>
    </Dialog>
  );
}
