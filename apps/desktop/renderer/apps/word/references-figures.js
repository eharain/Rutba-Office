// References → Insert Table of Figures, and Insert → Quick Parts → Field,
// as Word 365 lays them out. Table of Figures previews the list it will
// write for the chosen caption label; Field lists PAGE, NUMPAGES, DATE,
// TIME, FILENAME, AUTHOR and TITLE with their formats and shows the result
// before OK. The codes and results are `@rutba/ooxml/docfields`'s.

import React, { useMemo, useState } from 'react';
import { Button, Dialog, Select, Icon } from '@rutba/office-ui';
import {
  DOC_FIELDS, DATE_PICTURES, TIME_PICTURES, NUMBER_FORMATS, TEXT_FORMATS, docFieldInstr, evaluateDocField,
} from '@rutba/ooxml/docfields';
import { LEADERS } from './references-index.js';

/** The Captions group's own two, beside Insert Caption: Insert Table of Figures and Update Table. */
export function captionsExtra(refs, crossReference = null) {
  return (
    <div className="wd-refs-col">
      <Button icon="listBullet" label="Insert Table of Figures" title="Insert Table of Figures — a list of the captions of one label, with their pages" onClick={() => refs.act('insertFigures')} />
      <Button icon="refresh" label="Update Table" className="wd-refs-tof-update" disabled={!refs.info?.figures?.length} title="Update Table — the table of figures from the captions and pages as they stand" onClick={() => refs.act('updateFigures')} />
      {crossReference}
    </div>
  );
}

const LEADER_CSS = { dot: 'dotted', hyphen: 'dashed', underscore: 'solid' };

/**
 * Table of Figures: the caption label, with or without the label and
 * number, the page numbers right-aligned behind a leader — and a preview in
 * those choices, from this document's own captions when it has them.
 */
export function FiguresDialog({ labels, captions, onClose, onOk }) {
  const [label, setLabel] = useState(labels[0] || 'Figure');
  const [includeLabel, setIncludeLabel] = useState(true);
  const [pageNumbers, setPageNumbers] = useState(true);
  const [rightAlign, setRightAlign] = useState(true);
  const [leader, setLeader] = useState('dot');
  const mine = (captions[label] || []).slice(0, 6);
  const sample = mine.length
    ? mine.map((c, i) => ({ text: includeLabel ? c.text : c.bare || c.text, page: c.page ?? i + 1 }))
    : [1, 2, 3, 4, 5].map((n) => ({ text: includeLabel ? `${label} ${n}: Text` : 'Text', page: n * 2 - 1 }));

  return (
    <Dialog
      title="Table of Figures"
      width={620}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary className="wd-refs-tof-ok" label="OK" onClick={() => onOk({ label, includeLabel, pageNumbers, rightAlign: pageNumbers && rightAlign, leader })} />
        </>
      }
    >
      <div className="wd-refs-index">
        <div>
          <div className="wd-refs-section">Print Preview</div>
          <div className="wd-idx-preview wd-tof-preview">
            {sample.map((e, i) => (
              <div key={i} className={`wd-idx-line${pageNumbers && rightAlign ? ' right' : ''}`}>
                <span>{e.text}</span>
                {pageNumbers && rightAlign ? <span className="wd-idx-leader" style={{ borderBottomStyle: LEADER_CSS[leader] || 'none', borderBottomColor: LEADER_CSS[leader] ? 'currentColor' : 'transparent' }} /> : null}
                {pageNumbers ? <span>{rightAlign ? '' : ' '}{e.page}</span> : null}
              </div>
            ))}
          </div>
          {!mine.length ? <p className="wd-refs-lead wd-refs-foot-lead">No {label.toLowerCase()} captions yet: Insert Caption adds them, and Update Table brings the list up to date.</p> : null}
        </div>
        <div className="wd-refs-index-side">
          <label className="wd-refs-check"><input type="checkbox" className="wd-refs-tof-pages" checked={pageNumbers} onChange={(e) => setPageNumbers(e.target.checked)} />Show page numbers</label>
          <label className="wd-refs-check"><input type="checkbox" className="wd-refs-tof-right" checked={pageNumbers && rightAlign} disabled={!pageNumbers} onChange={(e) => setRightAlign(e.target.checked)} />Right align page numbers</label>
          <label className="wd-refs-leader">
            Tab leader
            <Select value={leader} disabled={!pageNumbers || !rightAlign} onChange={(e) => setLeader(e.target.value)}>
              {LEADERS.map(([id, text]) => <option key={id} value={id}>{text}</option>)}
            </Select>
          </label>
          <div className="wd-refs-section">General</div>
          <label className="wd-refs-leader">
            Caption label
            <Select className="rw-select wd-refs-tof-label" value={label} onChange={(e) => setLabel(e.target.value)}>
              {labels.map((l) => <option key={l} value={l}>{l}</option>)}
            </Select>
          </label>
          <label className="wd-refs-check"><input type="checkbox" className="wd-refs-tof-include" checked={includeLabel} onChange={(e) => setIncludeLabel(e.target.checked)} />Include label and number</label>
        </div>
      </div>
    </Dialog>
  );
}

const CATEGORIES = ['(All)', 'Date and Time', 'Document Information', 'Numbering'];

/**
 * Field (Insert → Quick Parts → Field): the field names in a category, the
 * formats each takes, and the result it gives right now.
 */
export function FieldDialog({ context, onClose, onOk }) {
  const [category, setCategory] = useState('(All)');
  const [name, setName] = useState('PAGE');
  const [picture, setPicture] = useState('');
  const [format, setFormat] = useState('');
  const [path, setPath] = useState(false);
  const [codes, setCodes] = useState(false);
  const list = DOC_FIELDS.filter((f) => category === '(All)' || f.category === category);
  const field = DOC_FIELDS.find((f) => f.name === name) || DOC_FIELDS[0];
  const pictures = field.formats === 'date' ? DATE_PICTURES : field.formats === 'time' ? TIME_PICTURES : null;
  const pic = pictures ? picture || pictures[0] : '';
  const instr = docFieldInstr({ name, picture: pic, format: pictures ? '' : format, path });
  const result = useMemo(() => evaluateDocField(instr, context), [instr, context]);

  const pick = (n) => {
    setName(n);
    setPicture('');
    setFormat('');
    setPath(false);
  };

  return (
    <Dialog
      title="Field"
      width={720}
      onClose={onClose}
      actions={
        <>
          <Button className="wd-refs-field-codes" label={codes ? 'Hide Codes' : 'Field Codes'} onClick={() => setCodes((c) => !c)} />
          <span className="wd-refs-foot-note" />
          <Button label="Cancel" onClick={onClose} />
          <Button primary className="wd-refs-field-ok" label="OK" onClick={() => onOk({ name, picture: pic, format: pictures ? '' : format, path })} />
        </>
      }
    >
      <div className="wd-refs-field">
        <div className="wd-refs-field-col">
          <div className="wd-refs-section">Please choose a field</div>
          <label className="wd-refs-leader">
            Categories
            <Select value={category} onChange={(e) => setCategory(e.target.value)}>
              {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
            </Select>
          </label>
          <div className="wd-refs-subhead">Field names</div>
          <div className="wd-refs-list wd-refs-fieldnames">
            {list.map((f) => (
              <button key={f.name} type="button" className={`wd-refs-row${f.name === name ? ' on' : ''}`} data-field={f.name} onClick={() => pick(f.name)} onDoubleClick={() => onOk({ name: f.name, picture: '', format: '', path: false })}>
                <span className="wd-refs-row-text">{f.name[0] + f.name.slice(1).toLowerCase()}</span>
              </button>
            ))}
          </div>
        </div>
        <div className="wd-refs-field-col">
          <div className="wd-refs-section">Field properties</div>
          {pictures ? (
            <>
              <div className="wd-refs-subhead">{field.formats === 'date' ? 'Date formats' : 'Time formats'}</div>
              <div className="wd-refs-list wd-refs-pictures">
                {pictures.map((p) => (
                  <button key={p} type="button" className={`wd-refs-row${p === pic ? ' on' : ''}`} onClick={() => setPicture(p)}>
                    <span className="wd-refs-row-text">{evaluateDocField(docFieldInstr({ name, picture: p }), context)}</span>
                  </button>
                ))}
              </div>
            </>
          ) : (
            <>
              <div className="wd-refs-subhead">Format</div>
              <div className="wd-refs-list wd-refs-formats">
                {(field.formats === 'number' ? NUMBER_FORMATS : TEXT_FORMATS).map(([id, label]) => (
                  <button key={id || 'none'} type="button" className={`wd-refs-row${id === format ? ' on' : ''}`} onClick={() => setFormat(id)}>
                    <span className="wd-refs-row-text">{label}</span>
                  </button>
                ))}
              </div>
            </>
          )}
        </div>
        <div className="wd-refs-field-col">
          <div className="wd-refs-section">Field options</div>
          {field.path ? (
            <label className="wd-refs-check"><input type="checkbox" checked={path} onChange={(e) => setPath(e.target.checked)} />Add path to filename</label>
          ) : null}
          <label className="wd-refs-check"><input type="checkbox" checked readOnly />Preserve formatting during updates</label>
          <div className="wd-refs-preview wd-refs-field-result">
            <div className="wd-refs-preview-head">{codes ? 'Field code' : 'Result'}</div>
            <div className="wd-refs-preview-body">
              {codes ? <code className="wd-refs-code">{`{${instr}}`}</code> : result ? <span className="wd-refs-result">{result}</span> : <span className="wd-refs-muted">{name === 'AUTHOR' || name === 'TITLE' ? `No ${name.toLowerCase()} in this document's properties yet.` : 'Worked out when the field is in place.'}</span>}
            </div>
          </div>
        </div>
      </div>
      <p className="wd-refs-lead wd-refs-foot-lead"><Icon name="info" size={12} /> Description: {field.description}. Update Fields (F9) works it out again, and so does printing.</p>
    </Dialog>
  );
}

export const FIGURES_CSS = `
.wd-tof-preview .wd-idx-line { padding-left: 0; text-indent: 0; margin-bottom: 3px; }
.rw-dialog[aria-label="Field"] { max-width: min(94vw, 760px); }
.wd-refs-field { display: grid; grid-template-columns: 1fr 1.1fr 1.1fr; gap: 16px; }
.wd-refs-field-col { display: flex; flex-direction: column; gap: 7px; min-width: 0; }
.wd-refs-subhead { font-size: 11.5px; font-weight: 600; color: var(--ink-2); margin-top: 2px; }
.wd-refs-field .wd-refs-list { height: 196px; }
.wd-refs-field-result { margin-top: 6px; }
.wd-refs-result { font-size: 14px; }
.wd-refs-code { font-family: Consolas, 'Cascadia Mono', monospace; font-size: 12px; color: var(--ink); background: transparent; white-space: pre-wrap; }
`;
