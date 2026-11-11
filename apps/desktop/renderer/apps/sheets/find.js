// Find in Sheets, as Excel's: a Find box at the end of the formula bar
// (Ctrl+F puts the keyboard in it; Enter finds the next match, Shift+Enter
// the one before, and it says which of how many; Esc goes back to the
// grid), and Find and Replace (Ctrl+H, or Find & Select) as a panel beside
// the grid rather than a dialog over it, with Match case, Match entire cell
// contents, Within Sheet or Workbook, and Find All's list of every match.

import React, { useEffect, useState } from 'react';
import { Button, Icon, Spacer } from '@rutba/office-ui';
import { t } from '../sheets-words.js';
import { findNames } from './find-names.js';
import { readFindResult, readFindList, findCountText, findOptions, findKey, goToMatch } from './find-model.js';

/** The Find box at the formula bar's end. `ref` is the input, for Ctrl+F. */
export function FindBox({ ref, dispatch, onBack, onReplace }) {
  const [text, setText] = useState('');
  const [note, setNote] = useState(null);
  const step = async (back) => {
    if (!text) return;
    const next = await dispatch({ op: 'findStep', text, back, ...findOptions() });
    if (next) setNote(findCountText(readFindResult(next.opResult), t));
  };
  return (
    <div className="sh-findbox" role="search">
      <Icon name="find" size={14} style={{ color: 'var(--ink-3)' }} />
      <input
        ref={ref}
        type="search"
        className="rw-input sh-findbox-input"
        spellCheck={false}
        aria-label={t('Find in sheet')}
        placeholder={t('Find in sheet')}
        title={t('Find in sheet: Enter finds the next match, Shift+Enter the one before, Esc goes back to the sheet')}
        value={text}
        onChange={(e) => { setText(e.target.value); setNote(null); }}
        onFocus={(e) => e.target.select()}
        onKeyDown={(e) => {
          if (e.key === 'Enter') { e.preventDefault(); step(e.shiftKey); return; }
          if (e.key === 'Escape') { e.preventDefault(); onBack?.(); return; }
          // Ctrl+H from the box opens Replace with the words already typed.
          if (findKey(e) === 'replace') { e.preventDefault(); e.stopPropagation(); onReplace?.(text); }
        }}
      />
      {note ? <span className="sh-findbox-count" role="status">{note}</span> : null}
    </div>
  );
}

/**
 * Find and Replace, beside the grid so the sheet stays in view and in
 * reach. `mode` is 'find' or 'replace'; `initial` the words to start with.
 */
export function FindPanel({ mode, initial = '', sheet, dispatch, onMode, onClose }) {
  const w = findNames();
  const [find, setFind] = useState(initial);
  const [replace, setReplace] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [wholeCell, setWholeCell] = useState(false);
  const [within, setWithin] = useState('sheet');
  const [note, setNote] = useState(null);
  const [list, setList] = useState(null);
  const opts = findOptions({ matchCase, wholeCell, within });
  // A changed search is a new search: the old list and its count no longer say anything.
  useEffect(() => { setNote(null); setList(null); }, [find, matchCase, wholeCell, within]);

  const step = async (back = false) => {
    if (!find) return;
    const next = await dispatch({ op: 'findStep', text: find, back, ...opts });
    if (!next) return;
    const r = readFindResult(next.opResult);
    setNote(r.found ? findCountText(r, t) : w.nothing);
  };
  const findAll = async () => {
    if (!find) return;
    const next = await dispatch({ op: 'findList', text: find, ...opts });
    if (!next) return;
    const rows = readFindList(next.opResult);
    setList(rows);
    setNote(rows.length ? t('Cells found: {count}', { count: rows.length }) : w.nothing);
  };
  const replaceOne = async () => {
    if (!find) return;
    const next = await dispatch({ op: 'replaceNext', find, replace, ...opts });
    if (!next) return;
    setList(null);
    setNote(next.opResult ? w.replacedOne : w.nothingLeft);
  };
  const replaceEvery = async () => {
    if (!find) return;
    const next = await dispatch({ op: 'replaceAll', find, replace, ...opts });
    if (!next) return;
    setList(null);
    setNote(Number(next.opResult) ? `${w.replacedAll} ${t('Cells found: {count}', { count: Number(next.opResult) })}` : w.nothing);
  };
  const keys = (e) => {
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); onClose(); return; }
    const k = findKey(e);
    if (k) { e.preventDefault(); e.stopPropagation(); onMode(k); }
  };

  return (
    <div className="sh-find" role="dialog" aria-modal="false" aria-label={w.title} onKeyDown={keys}>
      <div className="sh-find-head">
        <div className="sh-find-tabs" role="tablist">
          <button type="button" role="tab" aria-selected={mode === 'find'} className={mode === 'find' ? 'on' : undefined} onClick={() => onMode('find')}>{w.find}</button>
          <button type="button" role="tab" aria-selected={mode === 'replace'} className={mode === 'replace' ? 'on' : undefined} onClick={() => onMode('replace')}>{w.replace}</button>
        </div>
        <Spacer />
        <Button icon="close" title={w.close} onClick={onClose} />
      </div>
      <div className="sh-find-body">
        <label className="sh-find-field">
          <span>{w.find}</span>
          <input className="rw-input sh-find-what" spellCheck={false} autoFocus value={find} onFocus={(e) => e.target.select()}
            onChange={(e) => setFind(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); step(e.shiftKey); } }} />
        </label>
        {mode === 'replace' ? (
          <label className="sh-find-field">
            <span>{w.replaceWith}</span>
            <input className="rw-input sh-find-with" spellCheck={false} value={replace} onChange={(e) => setReplace(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); replaceOne(); } }} />
          </label>
        ) : null}
        <div className="sh-find-options">
          <label><input type="checkbox" className="sh-find-case" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />{w.matchCase}</label>
          <label><input type="checkbox" className="sh-find-whole" checked={wholeCell} onChange={(e) => setWholeCell(e.target.checked)} />{t('Match entire cell contents')}</label>
          <label>
            {t('Within')}
            <select className="rw-input sh-find-within" value={within} onChange={(e) => setWithin(e.target.value)}>
              <option value="sheet">{w.sheet}</option>
              <option value="workbook">{w.workbook}</option>
            </select>
          </label>
        </div>
        <div className="sh-find-actions">
          <Button className="sh-find-all" label={t('Find All')} disabled={!find} onClick={findAll} />
          <Button className="sh-find-prev" label={t('Find previous')} disabled={!find} onClick={() => step(true)} />
          <Button className="sh-find-next" primary={mode === 'find'} label={w.findNext} disabled={!find} onClick={() => step(false)} />
          {mode === 'replace' ? <Button className="sh-find-replace" label={w.replace} disabled={!find} onClick={replaceOne} /> : null}
          {mode === 'replace' ? <Button className="sh-find-replace-all" primary label={w.replaceAll} disabled={!find} onClick={replaceEvery} /> : null}
        </div>
        {note ? <div className="sh-find-note" role="status"><Icon name="info" size={14} />{note}</div> : null}
      </div>
      {list && list.length ? (
        <div className="sh-find-list" role="listbox" aria-label={t('Find All')}>
          <div className="sh-find-row sh-find-heads" aria-hidden="true">
            <span>{w.sheet}</span><span>{w.cell}</span><span>{w.value}</span><span>{w.formula}</span>
          </div>
          {list.map((m) => (
            <button type="button" role="option" aria-selected="false" key={`${m.sheet}!${m.ref}`} className="sh-find-row" data-ref={m.ref} data-sheet={m.sheet}
              onClick={() => dispatch(...goToMatch(m, sheet))}>
              <span>{m.sheet}</span><span>{m.ref}</span><span>{m.value}</span><span>{m.formula || ''}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

export const FIND_CSS = `
.sh-findbox { display: flex; align-items: center; gap: 6px; flex: 0 0 auto; }
.sh-formula .sh-findbox-input { width: 180px; font-family: inherit; font-size: 12.5px; }
.sh-findbox-count { font-size: 11.5px; color: var(--ink-3); white-space: nowrap; font-variant-numeric: tabular-nums; }
.sh-find {
  position: absolute; top: 10px; inset-inline-end: 10px; width: 360px; max-height: calc(100% - 20px); z-index: 16;
  display: flex; flex-direction: column; background: var(--surface); border: 1px solid var(--line);
  border-radius: var(--r-2); box-shadow: 0 8px 24px color-mix(in srgb, var(--ink) 20%, transparent); overflow: hidden;
}
.sh-find-head { display: flex; align-items: center; gap: 4px; padding: 6px 8px; border-bottom: 1px solid var(--line-soft); }
.sh-find-tabs { display: flex; gap: 2px; }
.sh-find-tabs button { border: 0; background: transparent; padding: 4px 10px; border-radius: var(--r-1, 4px); font: inherit; font-size: 12.5px; color: var(--ink-2); cursor: pointer; }
.sh-find-tabs button.on { background: var(--selected); color: var(--ink); font-weight: 600; }
.sh-find-body { display: flex; flex-direction: column; gap: 8px; padding: 10px; }
.sh-find-field { display: grid; grid-template-columns: 92px 1fr; align-items: center; gap: 8px; font-size: 12.5px; }
.sh-find-field .rw-input { height: 26px; }
.sh-find-options { display: flex; flex-wrap: wrap; gap: 6px 14px; font-size: 12px; color: var(--ink-2); }
.sh-find-options label { display: inline-flex; align-items: center; gap: 6px; }
.sh-find-options select { height: 24px; }
.sh-find-actions { display: flex; flex-wrap: wrap; gap: 6px; justify-content: flex-end; }
.sh-find-note { display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--ink-2); }
.sh-find-list { overflow: auto; border-top: 1px solid var(--line-soft); min-height: 0; }
.sh-find-row {
  display: grid; grid-template-columns: 1fr 60px 1.4fr 1.2fr; gap: 6px; width: 100%; padding: 5px 10px; border: 0;
  border-bottom: 1px solid var(--line-soft); background: transparent; text-align: start; font: inherit; font-size: 12px; color: var(--ink); cursor: pointer;
}
.sh-find-row span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sh-find-row:hover, .sh-find-row:focus-visible { background: var(--selected); }
.sh-find-heads { cursor: default; font-weight: 600; color: var(--ink-3); position: sticky; top: 0; background: var(--surface); }
.sh-find-heads:hover { background: var(--surface); }
`;
