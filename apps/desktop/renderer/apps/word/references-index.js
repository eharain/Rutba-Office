// References → Index: Mark Entry (Alt+Shift+X), Insert Index and Update
// Index, as Word 365 lays them out. Mark Index Entry stays open beside the
// page, as Word's does, so one entry after another can be marked; Insert
// Index previews the index it will write, in the columns, alignment, leader
// and layout chosen. The index's words are `@rutba/ooxml/wordindex`'s.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Dialog, Group, Input, Select, Icon } from '@rutba/office-ui';
import { buildIndex, lineTail } from '@rutba/ooxml/wordindex';

export const LEADERS = [
  ['none', '(none)'], ['dot', '.......'], ['hyphen', '-------'], ['underscore', '_______'],
];
const LEADER_CSS = { dot: 'dotted', hyphen: 'dashed', underscore: 'solid' };

/** References → Index, as Word lays the group out. */
export function IndexGroup({ act, hasIndex }) {
  return (
    <Group label="Index">
      <Button tall icon="flag" label={'Mark\nEntry'} className="wd-refs-2line" title="Mark Entry (Alt+Shift+X) — add the selected words to the index" onClick={() => act('markEntry')} />
      <div className="wd-refs-col">
        <Button icon="listBullet" label="Insert Index" title="Insert Index — a list of the marked words and the pages they are on" onClick={() => act('insertIndex')} />
        <Button icon="refresh" label="Update Index" disabled={!hasIndex} title="Update Index — the entries and their pages as they stand now" onClick={() => act('updateIndex')} />
      </div>
    </Group>
  );
}

/**
 * Mark Index Entry — beside the page, not over it: the main entry starts as
 * the selected words, and follows a new selection until it is typed in.
 */
export function MarkEntryPanel({ selected, bookmarks, onMark, onClose }) {
  const [main, setMain] = useState(selected || '');
  const [sub, setSub] = useState('');
  const [option, setOption] = useState('page');
  const [see, setSee] = useState('See ');
  const [bookmark, setBookmark] = useState(bookmarks[0]?.name || '');
  const [bold, setBold] = useState(false);
  const [italic, setItalic] = useState(false);
  const typed = useRef(false);
  const [note, setNote] = useState('');

  useEffect(() => {
    // A new selection is a new entry: its words, no subentry.
    if (!typed.current && selected) { setMain(selected); setSub(''); }
  }, [selected]);
  useEffect(() => {
    const onKey = (e) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);

  const spec = () => ({
    main: main.trim(), sub: sub.trim(),
    crossRef: option === 'see' ? see.trim() : '',
    bookmark: option === 'range' ? bookmark : '',
    bold, italic,
  });
  const mark = async (all) => {
    if (!main.trim()) return;
    const n = await onMark({ ...spec(), all, text: all ? selected : null });
    setNote(all ? `${n || 0} entr${n === 1 ? 'y' : 'ies'} marked.` : 'Marked. Select more words to mark another.');
    typed.current = false;
  };

  return (
    <div className="wd-refs-float" role="dialog" aria-label="Mark Index Entry">
      <div className="wd-refs-float-head">
        <span>Mark Index Entry</span>
        <button type="button" className="wd-refs-x" data-tip="Close" aria-label="Close" onClick={onClose}><Icon name="close" size={14} /></button>
      </div>
      <div className="wd-refs-float-body" onKeyDown={(e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT' && e.target.type === 'text') { e.preventDefault(); mark(false); } }}>
        <div className="wd-refs-section">Index</div>
        <div className="wd-refs-grid wd-refs-grid-tight">
          <label htmlFor="wd-refs-main">Main entry</label>
          <Input id="wd-refs-main" type="text" className="rw-input wd-refs-main" value={main} onChange={(e) => { typed.current = true; setMain(e.target.value); }} autoFocus />
          <label htmlFor="wd-refs-sub">Subentry</label>
          <Input id="wd-refs-sub" type="text" className="rw-input wd-refs-sub" value={sub} onChange={(e) => setSub(e.target.value)} />
        </div>
        <div className="wd-refs-section">Options</div>
        <div className="wd-refs-options">
          <label className="wd-refs-check">
            <input type="radio" name="wd-refs-opt" className="wd-refs-opt-see" checked={option === 'see'} onChange={() => setOption('see')} />
            Cross-reference:
          </label>
          <Input type="text" className="rw-input wd-refs-see" value={see} disabled={option !== 'see'} onChange={(e) => setSee(e.target.value)} />
          <label className="wd-refs-check wd-refs-span2">
            <input type="radio" name="wd-refs-opt" checked={option === 'page'} onChange={() => setOption('page')} />
            Current page
          </label>
          <label className="wd-refs-check">
            <input type="radio" name="wd-refs-opt" className="wd-refs-opt-range" checked={option === 'range'} disabled={!bookmarks.length} onChange={() => setOption('range')} />
            Page range
          </label>
          <Select className="rw-select wd-refs-bookmark" value={bookmark} disabled={option !== 'range' || !bookmarks.length} onChange={(e) => setBookmark(e.target.value)} title="Bookmark">
            {bookmarks.length ? bookmarks.map((b) => <option key={b.name} value={b.name}>{b.name}</option>) : <option value="">No bookmarks</option>}
          </Select>
        </div>
        <div className="wd-refs-section">Page number format</div>
        <div className="wd-refs-inline">
          <label className="wd-refs-check"><input type="checkbox" className="wd-refs-bold" checked={bold} onChange={(e) => setBold(e.target.checked)} /><b>Bold</b></label>
          <label className="wd-refs-check"><input type="checkbox" className="wd-refs-italic" checked={italic} onChange={(e) => setItalic(e.target.checked)} /><i>Italic</i></label>
        </div>
        <p className="wd-refs-lead wd-refs-note">{note || 'This box stays open so that you can mark one entry after another.'}</p>
      </div>
      <div className="wd-refs-float-foot">
        <Button primary className="wd-refs-mark" label="Mark" disabled={!main.trim()} onClick={() => mark(false)} />
        <Button className="wd-refs-markall" label="Mark All" disabled={!main.trim() || !selected || option !== 'page'} title="Mark All — the first time these words appear in every paragraph, matching case" onClick={() => mark(true)} />
        <Button label="Close" onClick={onClose} />
      </div>
    </div>
  );
}

/** Word's own print-preview entries, so the preview shows every kind of line. */
const SAMPLE = [
  { main: 'Aristotle', page: 2 },
  { main: 'Asteroid belt', crossRef: 'See Jupiter' },
  { main: 'Atmosphere', subs: ['Earth'], page: 4 },
  { main: 'Atmosphere', subs: ['Earth', 'exosphere'], page: 4 },
  { main: 'Atmosphere', subs: ['Earth', 'ionosphere'], page: 3 },
  { main: 'Atmosphere', subs: ['Earth', 'mesosphere'], page: 3 },
  { main: 'Atmosphere', subs: ['Earth', 'mesosphere'], page: 4 },
];

/** Insert Index: type, columns, right-aligned page numbers behind a leader — and a preview of all of it. */
export function IndexDialog({ current, onClose, onOk }) {
  const [runIn, setRunIn] = useState(Boolean(current?.runIn));
  const [columns, setColumns] = useState(String(current?.columns || 2));
  const [rightAlign, setRightAlign] = useState(Boolean(current?.rightAlign));
  const [leader, setLeader] = useState(current?.leader || 'dot');
  const groups = useMemo(() => buildIndex(SAMPLE, { runIn }), [runIn]);
  const right = rightAlign && !runIn;

  const line = (l, key) => {
    const tail = lineTail(l, { rightAlign: right });
    const pageSegs = tail.filter((s) => s.text !== '\t');
    const parts = [<span key="t">{l.text}</span>];
    if (runIn && l.runIn?.length) {
      parts.push(<span key="c">{lineTail(l).map((s, i) => <span key={i} style={{ fontStyle: s.italic ? 'italic' : undefined }}>{s.text}</span>)}: </span>);
      l.runIn.forEach((sub, i) => parts.push(<span key={'r' + i}>{i ? '; ' : ''}{sub.text}{lineTail(sub).map((s) => s.text).join('')}</span>));
      return <div key={key} className="wd-idx-line lvl1">{parts}</div>;
    }
    return (
      <div key={key} className={`wd-idx-line lvl${Math.min(3, l.level)}${right && l.pages.length ? ' right' : ''}`}>
        {parts}
        {right && l.pages.length ? <span className="wd-idx-leader" style={{ borderBottomStyle: LEADER_CSS[leader] || 'none', borderBottomColor: LEADER_CSS[leader] ? 'currentColor' : 'transparent' }} /> : null}
        <span className="wd-idx-pages">{pageSegs.map((s, i) => <span key={'p' + i} style={{ fontStyle: s.italic ? 'italic' : undefined }}>{right && i === 0 ? s.text.replace(/^, /, '') : s.text}</span>)}</span>
      </div>
    );
  };

  return (
    <Dialog
      title="Index"
      width={640}
      onClose={onClose}
      actions={
        <>
          <Button label="Cancel" onClick={onClose} />
          <Button primary className="wd-refs-index-ok" label="OK" onClick={() => onOk({ runIn, columns: Number(columns) || 1, rightAlign: right, leader })} />
        </>
      }
    >
      <div className="wd-refs-index">
        <div>
          <div className="wd-refs-section">Print Preview</div>
          <div className="wd-idx-preview" style={{ columnCount: Math.min(2, Number(columns) || 1) }}>
            {groups.map((g) => (
              <React.Fragment key={g.letter}>
                <div className="wd-idx-letter">{g.letter}</div>
                {g.lines.map((l, i) => line(l, g.letter + i))}
              </React.Fragment>
            ))}
          </div>
        </div>
        <div className="wd-refs-index-side">
          <div className="wd-refs-section">Type</div>
          <label className="wd-refs-check"><input type="radio" name="wd-idx-type" checked={!runIn} onChange={() => setRunIn(false)} />Indented</label>
          <label className="wd-refs-check"><input type="radio" name="wd-idx-type" className="wd-refs-runin" checked={runIn} onChange={() => setRunIn(true)} />Run-in</label>
          <div className="wd-refs-section">Columns</div>
          <Input type="number" min="1" max="4" className="rw-input wd-refs-columns" value={columns} onChange={(e) => setColumns(e.target.value.replace(/[^\d]/g, '').slice(0, 1))} />
          <label className="wd-refs-check">
            <input type="checkbox" className="wd-refs-right" checked={right} disabled={runIn} onChange={(e) => setRightAlign(e.target.checked)} />
            Right align page numbers
          </label>
          <label className="wd-refs-leader">
            Tab leader
            <Select className="rw-select wd-refs-leader-select" value={leader} disabled={!right} onChange={(e) => setLeader(e.target.value)}>
              {LEADERS.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
            </Select>
          </label>
        </div>
      </div>
      <p className="wd-refs-lead wd-refs-foot-lead">The index is written as Word's INDEX field; on this screen and on paper it is laid in one column, and Word lays it in the columns chosen.</p>
    </Dialog>
  );
}

export const INDEX_CSS = `
/* An XE field is hidden text: nothing on the page until ¶ shows it, as Word shows field codes, dotted under. */
.wd-field.wd-xe { background: none; }
.wd-page.marks .wd-xe::after { content: attr(data-xe); color: #6b7280; font-size: 0.92em; border-bottom: 1px dotted currentColor; margin: 0 1px; white-space: pre; }
.wd-refs-float { position: fixed; right: 28px; top: 168px; width: 372px; z-index: 60; background: var(--surface); border: 1px solid var(--line); border-radius: 14px; box-shadow: var(--shadow-3); display: flex; flex-direction: column; animation: rw-pop var(--slow); }
.wd-refs-float-head { display: flex; align-items: center; justify-content: space-between; padding: 12px 12px 6px 16px; font-family: var(--font-display); font-weight: 650; font-size: 14.5px; }
.wd-refs-x { border: 0; background: transparent; color: var(--ink-3); border-radius: 6px; padding: 4px; cursor: pointer; display: inline-flex; }
.wd-refs-x:hover { background: var(--hover); color: var(--ink); }
.wd-refs-float-body { padding: 0 16px 6px; display: flex; flex-direction: column; gap: 8px; }
.wd-refs-float-foot { display: flex; gap: 8px; justify-content: flex-end; padding: 10px 16px 12px; border-top: 1px solid var(--line-soft); background: var(--surface-2); border-radius: 0 0 14px 14px; }
.wd-refs-float-foot .rw-btn { padding: 6px 14px; min-height: 30px; min-width: 76px; justify-content: center; }
.wd-refs-float-foot .rw-btn:not(.primary) { border: 1px solid var(--line); background: var(--surface); }
.wd-refs-grid.wd-refs-grid-tight { grid-template-columns: 86px 1fr; max-height: none; }
.wd-refs-options { display: grid; grid-template-columns: 128px 1fr; gap: 7px 8px; align-items: center; }
.wd-refs-span2 { grid-column: 1 / span 2; }
.wd-refs-options .rw-input, .wd-refs-options .rw-select { width: 100%; box-sizing: border-box; }
.wd-refs-inline { display: flex; gap: 18px; }
.wd-refs-note { min-height: 18px; font-size: 12px; color: var(--ink-3); }
.wd-refs-index { display: grid; grid-template-columns: 1fr 190px; gap: 18px; }
.wd-refs-index-side { display: flex; flex-direction: column; gap: 8px; }
.wd-refs-index-side .rw-input { width: 64px; }
.wd-refs-leader { display: flex; flex-direction: column; gap: 5px; font-size: 12.5px; color: var(--ink-2); margin-top: 4px; }
.wd-idx-preview { border: 1px solid var(--line); border-radius: var(--r-2); background: #fff; color: #111; padding: 10px 14px; height: 208px; box-sizing: border-box; overflow: hidden; column-fill: auto; font-family: Calibri, 'Segoe UI', sans-serif; font-size: 12.5px; line-height: 1.4; column-gap: 22px; margin-top: 8px; }
.wd-idx-letter { font-weight: 700; font-family: 'Calibri Light', Calibri, sans-serif; margin: 2px 0 2px; break-after: avoid; }
.wd-idx-line { padding-left: 14px; text-indent: -14px; }
.wd-idx-line.lvl2 { padding-left: 28px; }
.wd-idx-line.lvl3 { padding-left: 42px; }
.wd-idx-line.right { display: flex; align-items: baseline; text-indent: 0; padding-left: 0; }
.wd-idx-line.right.lvl2 { padding-left: 14px; }
.wd-idx-line.right.lvl3 { padding-left: 28px; }
.wd-idx-pages { white-space: pre; }
.wd-idx-line.right > span:first-child { white-space: nowrap; }
.wd-idx-leader { flex: 1; border-bottom: 1.5px dotted currentColor; margin: 0 3px; transform: translateY(-3px); opacity: 0.8; min-width: 12px; }
.wd-refs-foot-lead { margin-top: 12px; font-size: 11.5px; color: var(--ink-3); }
`;
