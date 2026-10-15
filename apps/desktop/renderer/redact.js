// Redact, in Documents, Presentations and Worksheets: marks made as one
// goes, the selection, every match of a search, the patterns (e-mail
// addresses, telephone numbers, card numbers, a pattern of one's own) and
// areas of pictures, shown on the page as they are made, counted part by
// part across the whole file, and applied to a copy saved as a new file:
// the words gone from every part, black bars of the same length in their
// place, the picture areas blacked out, the copy read back to prove it.
//
// The window's side; the package pass is @rutba/ooxml/redact.

import React, { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Button, Field, Input, Icon, t, tn } from '@rutba/office-ui';
import { compileMarks, findSpans } from '@rutba/ooxml/redact-match';
import { openInApp } from './shell.js';

const PART_LABELS = {
  body: () => t('Body'), headers: () => t('Headers'), footers: () => t('Footers'), notes: () => t('Notes'), comments: () => t('Comments'),
  properties: () => t('Properties'), placeholders: () => t('Placeholders'), charts: () => t('Charts'), embedded: () => t('Embedded files'),
  pictures: () => t('Pictures'), slides: () => t('Slides'), masters: () => t('Masters and layouts'), cells: () => t('Cells'), workbook: () => t('Workbook'),
  links: () => t('Links'), customXml: () => t('Custom XML'), thumbnail: () => t('Thumbnail'), other: () => t('Other parts'),
};
export const partLabel = (kind) => (PART_LABELS[kind] || PART_LABELS.other)();

const PATTERN_LABELS = { email: () => t('E-mail addresses'), phone: () => t('Telephone numbers'), card: () => t('Card numbers') };

/** A mark, in words. */
export function markLabel(m) {
  if (m.pattern) return PATTERN_LABELS[m.pattern]?.() || m.pattern;
  if (m.regex != null) return t('Pattern: {pattern}', { pattern: m.regex });
  return '"' + m.text + '"';
}

const REDACT_CSS = `
.rd-hit { position: fixed; z-index: 30; pointer-events: none; outline: 1.5px solid #d32f2f; outline-offset: 0; background: rgba(211, 47, 47, 0.16); border-radius: 1px; }
.rd-pane { display: flex; flex-direction: column; gap: 14px; padding: 12px 14px 16px; font-size: 12.5px; }
.rd-lead { margin: 0; color: var(--ink-2); line-height: 1.5; }
.rd-section h4 { margin: 0 0 8px; font-size: 11.5px; font-weight: 600; letter-spacing: 0.02em; text-transform: uppercase; color: var(--ink-3); }
.rd-row { display: flex; gap: 6px; align-items: center; }
.rd-row .rw-input { flex: 1; min-width: 0; }
.rd-check { display: flex; align-items: center; gap: 8px; margin: 4px 0; cursor: pointer; }
.rd-opts { display: flex; gap: 14px; margin-top: 6px; color: var(--ink-2); }
.rd-marks { display: flex; flex-direction: column; border: 1px solid var(--line); border-radius: 8px; overflow: hidden; }
.rd-mark { display: grid; grid-template-columns: 14px 1fr auto 26px; align-items: center; gap: 8px; padding: 7px 8px 7px 10px; border-bottom: 1px solid var(--line-soft); }
.rd-mark:last-child { border-bottom: 0; }
.rd-mark .sw { width: 12px; height: 12px; border-radius: 2px; background: #111; }
.rd-mark .what { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.rd-mark .n { font-variant-numeric: tabular-nums; color: var(--ink-2); font-size: 11.5px; }
.rd-empty { padding: 12px; color: var(--ink-3); text-align: center; }
.rd-where { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 8px; }
.rd-where span { padding: 2px 8px; border-radius: 999px; background: var(--surface-2); color: var(--ink-2); font-size: 11.5px; }
.rd-apply { display: flex; flex-direction: column; gap: 8px; padding-top: 12px; border-top: 1px solid var(--line-soft); }
.rd-apply .rw-btn.primary { justify-content: center; }
.rd-note { color: var(--ink-3); font-size: 11.5px; line-height: 1.45; }
.rd-done { display: flex; flex-direction: column; gap: 6px; padding: 10px 12px; border-radius: 8px; background: color-mix(in srgb, var(--good, #1e8e3e) 9%, var(--surface)); border: 1px solid color-mix(in srgb, var(--good, #1e8e3e) 30%, transparent); }
.rd-done strong { font-size: 12.5px; }
.rd-done .path { color: var(--ink-2); word-break: break-all; font-size: 11.5px; }
.rd-area { position: fixed; z-index: 30; pointer-events: none; background: repeating-linear-gradient(135deg, rgba(211,47,47,0.34) 0 5px, rgba(211,47,47,0.12) 5px 10px); outline: 1.5px solid #d32f2f; }
.rd-area.live { background: rgba(17, 17, 17, 0.35); outline: 1.5px dashed #c62828; }
.rd-picking img.wd-image, .rd-picking .sl-pic, .rd-picking [data-redact-picture] { cursor: crosshair !important; outline: 2px dashed #c62828; outline-offset: 2px; }
`;

let cssInstalled = false;
function installCss() {
  if (cssInstalled || typeof document === 'undefined') return;
  cssInstalled = true;
  const style = document.createElement('style');
  style.id = 'rutba-redact-css';
  style.textContent = REDACT_CSS;
  document.head.appendChild(style);
}

/**
 * The marks drawn over the page's words as they are made, as Acrobat marks
 * text for redaction: a red box round each match, the words still to be
 * read until Apply. Boxes rather than a colour on the words, so the same
 * marks show over a document's page, a slide's drawing and a sheet's grid.
 * Answers the boxes, measured again as the page moves.
 */
function useMarkHighlights(rootOf, marks, version) {
  const [hits, setHits] = useState([]);
  useLayoutEffect(() => {
    const root = rootOf();
    let finders = [];
    try { finders = compileMarks(marks.filter((m) => !m.picture)); } catch { finders = []; }
    if (!root || !finders.length) { setHits([]); return undefined; }
    const ranges = [];
    // Words are read per paragraph-like element, the text nodes of each joined, so a name across two runs is found.
    const groups = new Map();
    const seen = new Map();
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const owner = n.parentElement?.closest('[data-block], [data-redact-text], text, p, li, td, div') || root;
      // Words laid out only to be measured (a page's hidden copy of its notes) are not on the page.
      if (!seen.has(owner)) seen.set(owner, Boolean(owner.closest?.('[aria-hidden="true"], .wd-notes-measure')) || getComputedStyle(owner).visibility === 'hidden');
      if (seen.get(owner)) continue;
      if (!groups.has(owner)) groups.set(owner, []);
      groups.get(owner).push(n);
    }
    for (const nodes of groups.values()) {
      const text = nodes.map((n) => n.nodeValue).join('');
      for (const [a, b] of findSpans(text, finders)) {
        let at = 0;
        const range = document.createRange();
        let started = false;
        for (const n of nodes) {
          const len = n.nodeValue.length;
          if (!started && a < at + len) { range.setStart(n, a - at); started = true; }
          if (started && b <= at + len) { range.setEnd(n, b - at); break; }
          at += len;
        }
        if (started) ranges.push(range);
      }
    }
    let raf = 0;
    const measure = () => {
      const out = [];
      for (const range of ranges) {
        const text = range.toString();
        for (const r of range.getClientRects()) if (r.width > 0.5 && r.height > 0.5) out.push({ left: r.left, top: r.top, width: r.width, height: r.height, text });
      }
      setHits((was) => (JSON.stringify(was) === JSON.stringify(out) ? was : out));
      raf = requestAnimationFrame(measure);
    };
    measure();
    return () => cancelAnimationFrame(raf);
  }, [rootOf, marks, version]);
  return hits;
}

/** A picture's bytes with its marked areas painted black, in its own format where the window can write it. */
async function blackedOut(href, rects, type) {
  // Read as bytes and decoded here, so a picture the window was handed by URL draws untainted.
  const img = await createImageBitmap(await (await fetch(href)).blob());
  const canvas = document.createElement('canvas');
  canvas.width = img.width;
  canvas.height = img.height;
  const g = canvas.getContext('2d');
  g.drawImage(img, 0, 0);
  g.fillStyle = '#000000';
  for (const r of rects) g.fillRect(Math.floor(r.x * canvas.width), Math.floor(r.y * canvas.height), Math.ceil(r.w * canvas.width) + 1, Math.ceil(r.h * canvas.height) + 1);
  const out = /jpe?g/.test(type) ? 'image/jpeg' : /webp/.test(type) ? 'image/webp' : 'image/png';
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, out, 0.92));
  return new Uint8Array(await blob.arrayBuffer());
}

/**
 * Redact for one window. `adapter`:
 *   pageRoot()        the element whose words the marks are drawn over
 *   selectionText()   the words selected now, for Mark Selection
 *   pictureAt(el)     for a picture element on the page: { part, href, type, name, el? (the drawn image, when it is not `el`) } or null
 *   pictureEls()      the picture elements on the page, for drawing the areas
 *   version           anything that changes when the page is drawn again
 */
export function useRedaction({ app, shell, doc, toast, adapter }) {
  installCss();
  const [open, setOpen] = useState(false);
  const [marks, setMarks] = useState([]);
  const [scan, setScan] = useState(null);
  const [picking, setPicking] = useState(false);
  const [done, setDone] = useState(null);
  const [busy, setBusy] = useState(false);
  const textMarks = useMemo(() => marks.filter((m) => !m.picture), [marks]);
  const id = doc?.id;

  // Counted across the whole file, a moment after each change of marks.
  useEffect(() => {
    if (!open || !id) return undefined;
    if (!textMarks.length) { setScan({ total: 0, parts: [], words: [] }); return undefined; }
    const timer = setTimeout(() => {
      shell.doc.redact({ id, action: 'scan', marks: textMarks }).then(setScan).catch((err) => toast(err.message, { tone: 'bad' }));
    }, 150);
    return () => clearTimeout(timer);
  }, [open, id, textMarks, shell, toast, doc?.version]);

  const rootOf = useCallback(() => (open ? adapter.pageRoot() : null), [open, adapter]);
  const hits = useMarkHighlights(rootOf, textMarks, adapter.version);

  const add = useCallback((mark) => {
    try { compileMarks([mark]); } catch (err) { toast(err.message, { tone: 'bad' }); return false; }
    setMarks((list) => (list.some((m) => JSON.stringify(m) === JSON.stringify(mark)) ? list : [...list, mark]));
    setDone(null);
    return true;
  }, [toast]);

  // An area of a picture: dragged over the picture while Mark Area is on.
  const [live, setLive] = useState(null);
  useEffect(() => {
    if (!picking) return undefined;
    const root = adapter.pageRoot();
    root?.classList.add('rd-picking');
    let drag = null;
    const down = (e) => {
      const el = e.target?.closest?.(adapter.pictureSelector || 'img.wd-image, [data-redact-picture]');
      const pic = el ? adapter.pictureAt(el) : null;
      if (!pic) return;
      e.preventDefault();
      e.stopPropagation();
      // The picture's own box: an app may hand back the drawn image for a press on what stands over it.
      const r = (pic.el || el).getBoundingClientRect();
      drag = { el, pic, r, x0: e.clientX, y0: e.clientY };
      setLive({ left: e.clientX, top: e.clientY, width: 0, height: 0 });
    };
    const move = (e) => {
      if (!drag) return;
      e.preventDefault();
      const x = Math.max(drag.r.left, Math.min(drag.r.right, e.clientX));
      const y = Math.max(drag.r.top, Math.min(drag.r.bottom, e.clientY));
      setLive({ left: Math.min(x, drag.x0), top: Math.min(y, drag.y0), width: Math.abs(x - drag.x0), height: Math.abs(y - drag.y0) });
    };
    const up = (e) => {
      if (!drag) return;
      e.preventDefault();
      e.stopPropagation();
      const { r, pic } = drag;
      const x = Math.max(r.left, Math.min(r.right, e.clientX));
      const y = Math.max(r.top, Math.min(r.bottom, e.clientY));
      const rect = { x: (Math.min(x, drag.x0) - r.left) / r.width, y: (Math.min(y, drag.y0) - r.top) / r.height, w: Math.abs(x - drag.x0) / r.width, h: Math.abs(y - drag.y0) / r.height };
      drag = null;
      setLive(null);
      if (rect.w < 0.01 || rect.h < 0.01) return;
      setMarks((list) => [...list, { picture: pic.part, name: pic.name || '', href: pic.href, type: pic.type, rect }]);
      setDone(null);
    };
    const key = (e) => { if (e.key === 'Escape') setPicking(false); };
    document.addEventListener('mousedown', down, true);
    document.addEventListener('mousemove', move, true);
    document.addEventListener('mouseup', up, true);
    window.addEventListener('keydown', key);
    return () => {
      root?.classList.remove('rd-picking');
      document.removeEventListener('mousedown', down, true);
      document.removeEventListener('mousemove', move, true);
      document.removeEventListener('mouseup', up, true);
      window.removeEventListener('keydown', key);
    };
  }, [picking, adapter]);

  // The areas marked so far, drawn over their pictures wherever those are now.
  const [areas, setAreas] = useState([]);
  useEffect(() => {
    const pics = marks.filter((m) => m.picture);
    if (!open || !pics.length) { setAreas([]); return undefined; }
    let raf = 0;
    const measure = () => {
      const out = [];
      for (const el of adapter.pictureEls()) {
        const pic = adapter.pictureAt(el);
        if (!pic) continue;
        const r = (pic.el || el).getBoundingClientRect();
        pics.filter((m) => m.picture === pic.part).forEach((m, i) => out.push({ key: `${pic.part}:${i}:${r.left}`, left: r.left + m.rect.x * r.width, top: r.top + m.rect.y * r.height, width: m.rect.w * r.width, height: m.rect.h * r.height }));
      }
      setAreas((was) => (JSON.stringify(was) === JSON.stringify(out) ? was : out));
      raf = requestAnimationFrame(measure);
    };
    measure();
    return () => cancelAnimationFrame(raf);
  }, [open, marks, adapter]);

  const apply = useCallback(async () => {
    if (!id || busy) return;
    const base = String(doc?.name || 'Document').replace(/\.[^.]+$/, '');
    const ext = String(doc?.name || '').split('.').pop() || { word: 'docx', slides: 'pptx', sheets: 'xlsx' }[app];
    const dir = doc?.path ? doc.path.replace(/[\\/][^\\/]*$/, '') : null;
    const sep = dir && dir.includes('\\') ? '\\' : '/';
    const target = await shell.dialog.save({ title: t('Save the redacted copy as'), defaultPath: (dir ? dir + sep : '') + t('{name} (redacted)', { name: base }) + '.' + ext, filters: [{ name: t('Redacted copy'), extensions: [ext] }] });
    if (!target) return;
    setBusy(true);
    try {
      const pictures = {};
      const byPart = new Map();
      for (const m of marks.filter((x) => x.picture)) {
        if (!byPart.has(m.picture)) byPart.set(m.picture, { href: m.href, type: m.type, rects: [] });
        byPart.get(m.picture).rects.push(m.rect);
      }
      for (const [part, p] of byPart) pictures[part] = await blackedOut(p.href, p.rects, p.type);
      const result = await shell.doc.redact({ id, action: 'apply', marks: textMarks, pictures, path: target });
      setDone(result);
      toast(t('Saved a redacted copy as {name}: {count} taken away.', { name: target.split(/[\\/]/).pop(), count: result.total }), { tone: 'good', ms: 6000 });
    } catch (err) {
      toast(err.message, { tone: 'bad', ms: 6000 });
    } finally {
      setBusy(false);
    }
  }, [id, busy, doc, app, shell, marks, textMarks, toast]);

  const overlay = open ? createPortal(
    <>
      {hits.map((h, i) => <div key={'h' + i} className="rd-hit" data-text={h.text} style={{ left: h.left, top: h.top, width: h.width, height: h.height }} />)}
      {areas.map((a) => <div key={a.key} className="rd-area" style={{ left: a.left, top: a.top, width: a.width, height: a.height }} />)}
      {live ? <div className="rd-area live" style={live} /> : null}
    </>,
    document.body,
  ) : null;

  const pane = open ? (
    <RedactPane
      app={app}
      marks={marks}
      scan={scan}
      picking={picking}
      busy={busy}
      done={done}
      onAdd={add}
      onRemove={(i) => { setMarks((list) => list.filter((_, j) => j !== i)); setDone(null); }}
      onSelection={() => {
        // A line of the selection a mark of its own: a selection over paragraphs, or cells, marks each.
        const lines = [...new Set(String(adapter.selectionText() || '').split(/\r?\n/).map((s) => s.trim()).filter(Boolean))];
        if (!lines.length) { toast(t('Select the words to mark first.'), { ms: 3500 }); return; }
        for (const words of lines) add({ text: words });
      }}
      onPicking={() => setPicking((p) => !p)}
      onApply={apply}
      onOpenCopy={() => done?.path && openInApp(shell, done.path, app)}
      onShowCopy={() => done?.path && shell.shell.showInFolder({ path: done.path })}
    />
  ) : null;

  return { open, setOpen, toggle: () => setOpen((o) => !o), pane, overlay, marks, add, apply, scan, picking, setPicking, done };
}

/** The Redact pane: Mark, the marks with what each finds, and Apply. */
export function RedactPane({ app, marks, scan, picking, busy, done, onAdd, onRemove, onSelection, onPicking, onApply, onOpenCopy, onShowCopy }) {
  const [find, setFind] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [whole, setWhole] = useState(false);
  const [custom, setCustom] = useState('');
  const has = (pattern) => marks.some((m) => m.pattern === pattern);
  const toggle = (pattern) => (has(pattern) ? onRemove(marks.findIndex((m) => m.pattern === pattern)) : onAdd({ pattern }));
  const countOf = (m) => {
    if (m.picture) return null;
    if (!scan?.words) return null;
    let finders;
    try { finders = compileMarks([m]); } catch { return 0; }
    return scan.words.filter((w) => findSpans(w, finders).length).length;
  };
  const pictureLabel = app === 'slides' ? t('Mark an area of a picture on the slide') : t('Mark an area of a picture on the page');
  // A sheet's pictures float over the grid and are not marked here: only its cells' words are.
  const pictures = app !== 'sheets';
  return (
    <div className="rd-pane">
      <p className="rd-lead">{t('Mark what must not be read. The marks show as you go; Apply saves a copy with them taken out of every part of the file, black bars in their place.')}</p>
      <div className="rd-section">
        <h4>{t('Mark')}</h4>
        <Button className="rd-mark-selection" icon="highlighter" label={t('Mark Selection')} onClick={onSelection} />
        <div className="rd-row" style={{ marginTop: 8 }}>
          <Input className="rd-find" value={find} placeholder={t('Find words to mark')} onChange={(e) => setFind(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && find.trim() && onAdd({ text: find.trim(), ...(matchCase ? { matchCase } : {}), ...(whole ? { wholeWords: whole } : {}) })) setFind(''); }} />
          <Button className="rd-find-mark" label={t('Mark All')} disabled={!find.trim()} onClick={() => { if (onAdd({ text: find.trim(), ...(matchCase ? { matchCase } : {}), ...(whole ? { wholeWords: whole } : {}) })) setFind(''); }} />
        </div>
        <div className="rd-opts">
          <label className="rd-check"><input type="checkbox" className="rd-case" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />{t('Match case')}</label>
          <label className="rd-check"><input type="checkbox" className="rd-whole" checked={whole} onChange={(e) => setWhole(e.target.checked)} />{t('Whole words only')}</label>
        </div>
      </div>
      <div className="rd-section">
        <h4>{t('Patterns')}</h4>
        {['email', 'phone', 'card'].map((p) => (
          <label key={p} className="rd-check"><input type="checkbox" className={`rd-pattern-${p}`} checked={has(p)} onChange={() => toggle(p)} />{PATTERN_LABELS[p]()}</label>
        ))}
        <div className="rd-row" style={{ marginTop: 6 }}>
          <Input className="rd-custom" value={custom} placeholder={t('A pattern of your own, such as INV-\\d{4}')} onChange={(e) => setCustom(e.target.value)} />
          <Button className="rd-custom-add" label={t('Add')} disabled={!custom.trim()} onClick={() => { if (onAdd({ regex: custom.trim() })) setCustom(''); }} />
        </div>
      </div>
      {pictures ? <div className="rd-section">
        <h4>{t('Pictures')}</h4>
        <Button className="rd-pick-area" icon="marquee" label={picking ? t('Done Marking Areas') : pictureLabel} pressed={picking} onClick={onPicking} />
        {picking ? <div className="rd-note" style={{ marginTop: 6 }}>{t('Drag over a picture to mark the part to black out. Escape when done.')}</div> : null}
      </div> : null}
      <div className="rd-section">
        <h4>{tn(marks.length, '{count} mark', '{count} marks')}</h4>
        <div className="rd-marks">
          {marks.length ? marks.map((m, i) => {
            const n = countOf(m);
            return (
              <div key={i} className="rd-mark" data-mark={i}>
                <span className="sw" />
                <span className="what" data-tip={m.picture ? m.name || m.picture : markLabel(m)}>{m.picture ? t('Area of {picture}', { picture: m.name || m.picture.split('/').pop() }) : markLabel(m)}</span>
                <span className="n">{m.picture ? t('picture') : n == null ? '…' : tn(n, '{count} found', '{count} found')}</span>
                <Button icon="close" title={t('Remove this mark')} onClick={() => onRemove(i)} />
              </div>
            );
          }) : <div className="rd-empty">{t('Nothing is marked yet.')}</div>}
        </div>
        {scan?.parts?.length ? (
          <div className="rd-where rd-where-parts">
            {Object.entries(scan.parts.reduce((acc, p) => ({ ...acc, [p.kind]: (acc[p.kind] || 0) + p.count }), {})).map(([kind, count]) => (
              <span key={kind} data-kind={kind}>{partLabel(kind)} · {count}</span>
            ))}
          </div>
        ) : null}
      </div>
      <div className="rd-apply">
        <Button primary className="rd-apply-btn" icon="shield" label={busy ? t('Applying…') : t('Apply Redactions')} disabled={busy || !marks.length} onClick={onApply} />
        <div className="rd-note">{t('The words are taken out of the body, headers and footers, notes, comments, tracked changes, fields, alt text and the file\'s properties; the open file is left as it is.')}</div>
        {done ? (
          <div className="rd-done">
            <strong>{tn(done.total, 'Redacted copy saved: {count} item taken away.', 'Redacted copy saved: {count} items taken away.')}</strong>
            <span className="path">{done.path}</span>
            {done.checked ? <span className="rd-note">{t('Read back: none of the marked words is left in any part of the copy.')}</span> : null}
            <div className="rd-row"><Button className="rd-open-copy" icon="open" label={t('Open the Copy')} onClick={onOpenCopy} /><Button icon="folderOpen" label={t('Show in Folder')} onClick={onShowCopy} /></div>
          </div>
        ) : null}
      </div>
    </div>
  );
}
