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
