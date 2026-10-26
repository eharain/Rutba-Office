// Comments on a PDF's pages: the layer over each page where comments are
// picked and drawn, the words selected turned into the quadrilaterals a
// highlight covers, the box a note or a text box is typed into, and the
// Comments pane, as Acrobat's comment tools and pane work.
//
// What a comment looks like on the page is drawn by pdf.js from the
// appearance the engine wrote into the file (@rutba/pdf/annotations), so the
// page shows exactly what any other reader shows. This layer draws what is
// not in the file yet (a shape being dragged out), the outline of the
// comment picked, and a comment another writer saved with no appearance of
// its own, from its kind, its colour and its geometry, as readers do.
//
// Everything here works in the page's own coordinates (PDF points, y up),
// turned to and from the screen through the page's pdf.js viewport, so a
// turned or cropped page is right without a case of its own.

import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Button, Icon, t, formatWhen } from '@rutba/office-ui';
import { getPage } from '../../pdfjs.js';

/** Tools that act on the words selected. */
export const MARKUP_TOOLS = new Set(['highlight', 'underline', 'strike', 'redact-text']);
/** Tools drawn on the page with the pointer. */
export const DRAW_TOOLS = new Set(['note', 'text', 'ink', 'rect', 'ellipse', 'line', 'arrow', 'redact-area', 'stamp', 'date', 'check']);

/** Comments moved on the page by dragging them: a signature or initials, a tick, a date or another text box, a sticky note. */
export const MOVABLE_TYPES = new Set(['Stamp', 'FreeText', 'Text']);
/** Of those, the ones resized by a corner (a sticky note's icon keeps its size, as in Acrobat). */
const RESIZABLE_TYPES = new Set(['Stamp', 'FreeText']);
/** The smallest a comment is made by its corners, in points. */
const SMALLEST = 8;

/**
 * Where a comment goes when it is dragged: by its body `corner` null moves
 * it, by a corner (0 bottom left, 1 bottom right, 2 top right, 3 top left,
 * in the page's own coordinates) resizes it, the opposite corner staying
 * where it is. A stamp (a signature, a tick) keeps its proportions.
 */
export function shiftedRect(rect, from, to, { corner = null, keepShape = false } = {}) {
  const [x0, y0, x1, y1] = rect;
  if (corner == null) {
    const dx = to[0] - from[0];
    const dy = to[1] - from[1];
    return [x0 + dx, y0 + dy, x1 + dx, y1 + dy];
  }
  const fx = corner === 0 || corner === 3 ? x1 : x0;
  const fy = corner === 0 || corner === 1 ? y1 : y0;
  let w = Math.max(SMALLEST, Math.abs(to[0] - fx));
  let h = Math.max(SMALLEST, Math.abs(to[1] - fy));
  if (keepShape && x1 > x0 && y1 > y0) {
    const k = Math.max(w / (x1 - x0), h / (y1 - y0));
    w = Math.max(SMALLEST, (x1 - x0) * k);
    h = Math.max(SMALLEST, (y1 - y0) * k);
  }
  const ox = corner === 0 || corner === 3 ? -1 : 1;
  const oy = corner === 0 || corner === 1 ? -1 : 1;
  const ax = fx + ox * w;
  const ay = fy + oy * h;
  return [Math.min(fx, ax), Math.min(fy, ay), Math.max(fx, ax), Math.max(fy, ay)];
}

/** The annotation each tool writes. */
export const TOOL_TYPES = {
  highlight: 'Highlight', underline: 'Underline', strike: 'StrikeOut', note: 'Text', text: 'FreeText', ink: 'Ink',
  rect: 'Square', ellipse: 'Circle', line: 'Line', arrow: 'Arrow', 'redact-text': 'Redact', 'redact-area': 'Redact', stamp: 'Stamp',
};

/** Acrobat's colours for its comment tools, and the ones offered beside them. */
export const SWATCHES = [
  { key: 'yellow', rgb: [1, 0.92, 0.23], label: () => t('Yellow') },
  { key: 'green', rgb: [0.3, 0.75, 0.3], label: () => t('Green') },
  { key: 'blue', rgb: [0.24, 0.55, 0.95], label: () => t('Blue') },
  { key: 'pink', rgb: [0.98, 0.45, 0.72], label: () => t('Pink') },
  { key: 'red', rgb: [0.86, 0.16, 0.16], label: () => t('Red') },
  { key: 'black', rgb: [0.1, 0.1, 0.1], label: () => t('Black') },
];
export const DEFAULT_TOOL_COLOURS = {
  highlight: 'yellow', underline: 'green', strike: 'red', note: 'yellow', text: 'black', ink: 'blue', rect: 'red', ellipse: 'red', line: 'red', arrow: 'red',
};
export const swatch = (key) => SWATCHES.find((s) => s.key === key) || SWATCHES[0];
export const cssColour = (rgb) => (rgb ? `rgb(${rgb.map((v) => Math.round(Math.max(0, Math.min(1, v)) * 255)).join(' ')})` : 'transparent');

/** What a comment of each kind is called in the pane. */
export function kindName(a) {
  switch (a.type) {
    case 'Highlight': return t('Highlight');
    case 'Underline': return t('Underline');
    case 'StrikeOut': return t('Strikethrough');
    case 'Squiggly': return t('Squiggly underline');
    case 'Text': return t('Sticky note');
    case 'FreeText': return t('Text box');
    case 'Ink': return t('Drawing');
    case 'Square': return t('Rectangle');
    case 'Circle': return t('Ellipse');
    case 'Line': return a.arrow ? t('Arrow') : t('Line');
    case 'Stamp': return a.icon === 'Initials' ? t('Initials') : a.icon === 'Signature' ? t('Signature') : a.icon === 'Check' ? t('Checkmark') : t('Stamp');
    case 'Redact': return t('Marked for redaction');
    default: return a.type;
  }
}
const KIND_ICONS = {
  Highlight: 'highlighter', Underline: 'underline', StrikeOut: 'strike', Squiggly: 'underline', Text: 'note', FreeText: 'textbox', Ink: 'pen',
  Square: 'mkRect', Circle: 'mkEllipse', Line: 'mkLine', Stamp: 'signature', Redact: 'redact',
};
export const kindIcon = (a) => (a.type === 'Line' && a.arrow ? 'mkArrow' : KIND_ICONS[a.type] || 'comment');

/** The comments that stand on their own (not replies, not hidden review states). */
export const topLevel = (list) => (list || []).filter((a) => !a.replyTo && !a.hidden);

// ── the words selected, as quadrilaterals ─────────────────────────────────

/**
 * The selection's lines on each page, as QuadPoints in the page's own
 * coordinates: { [pageKey]: [[x1, y1, x2, y2, x3, y3, x4, y4], ...] }.
 * pdf.js lays a word out as several spans; the rectangles of one line are
 * joined into one, as Acrobat's highlight covers a line.
 */
export function selectionQuads(view, viewports) {
  const sel = window.getSelection();
  if (!sel || sel.isCollapsed || !sel.rangeCount || !view) return {};
  const out = {};
  const pages = [...view.querySelectorAll('.pdf-page')].map((el) => ({ el, box: el.getBoundingClientRect(), key: el.dataset.key }));
  for (let r = 0; r < sel.rangeCount; r++) {
    const range = sel.getRangeAt(r);
    const lines = new Map();
    for (const rect of range.getClientRects()) {
      if (rect.width < 0.5 || rect.height < 0.5) continue;
      const cx = rect.left + rect.width / 2;
      const cy = rect.top + rect.height / 2;
      const page = pages.find((p) => cx >= p.box.left && cx <= p.box.right && cy >= p.box.top && cy <= p.box.bottom);
      if (!page) continue;
      if (!lines.has(page.key)) lines.set(page.key, { page, rects: [] });
      lines.get(page.key).rects.push({ left: rect.left - page.box.left, right: rect.right - page.box.left, top: rect.top - page.box.top, bottom: rect.bottom - page.box.top });
    }
    for (const [key, { rects }] of lines) {
      const vp = viewports.get(key);
      if (!vp) continue;
      for (const line of joinLines(rects)) {
        const tl = vp.convertToPdfPoint(line.left, line.top);
        const tr = vp.convertToPdfPoint(line.right, line.top);
        const bl = vp.convertToPdfPoint(line.left, line.bottom);
        const br = vp.convertToPdfPoint(line.right, line.bottom);
        (out[key] ||= []).push([...tl, ...tr, ...bl, ...br].map((v) => Math.round(v * 100) / 100));
      }
    }
  }
  return out;
}

/** Rectangles that share a line (overlapping most of their height) joined into one, top to bottom. */
export function joinLines(rects) {
  const sorted = [...rects].sort((a, b) => a.top - b.top || a.left - b.left);
  const lines = [];
  for (const r of sorted) {
    const h = r.bottom - r.top;
    const line = lines.find((l) => {
      const overlap = Math.min(l.bottom, r.bottom) - Math.max(l.top, r.top);
      return overlap > Math.min(h, l.bottom - l.top) * 0.5;
    });
    if (line) {
      line.left = Math.min(line.left, r.left); line.right = Math.max(line.right, r.right);
      line.top = Math.min(line.top, r.top); line.bottom = Math.max(line.bottom, r.bottom);
    } else lines.push({ ...r });
  }
  return lines;
}

// ── the layer over a page ─────────────────────────────────────────────────

const pts = (vp, list) => {
  const out = [];
  for (let i = 0; i + 1 < list.length; i += 2) {
    const [x, y] = vp.convertToViewportPoint(list[i], list[i + 1]);
    out.push(`${Math.round(x * 10) / 10},${Math.round(y * 10) / 10}`);
  }
  return out.join(' ');
};
const quadPts = (vp, q) => pts(vp, [q[0], q[1], q[2], q[3], q[6], q[7], q[4], q[5]]);
const rectPts = (vp, r) => pts(vp, [r[0], r[3], r[2], r[3], r[2], r[1], r[0], r[1]]);

/**
 * Over one page: an outline under every comment to press, the comment picked
 * outlined, and the shape a drawing tool is dragging out. Reports each new
 * shape in the page's coordinates through `onCreate`.
 */
export function AnnotLayer({ entry, source, scale, annots, tool, colour, picked, onPick, onCreate, viewports, fields = null, onLink = null, onMove = null }) {
  const [vp, setVp] = useState(null);
  const [draft, setDraft] = useState(null);
  // A comment being dragged to a new place or size: { id, rect } in the page's coordinates.
  const [shift, setShift] = useState(null);
  const ref = useRef(null);
  useEffect(() => {
    let live = true;
    if (!source?.doc) return undefined;
    getPage(source.doc, entry.page).then((page) => {
      if (!live) return;
      const v = page.getViewport({ scale, rotation: entry.rotate });
      viewports?.set(entry.key, v);
      setVp(v);
    }).catch(() => {});
    return () => { live = false; };
  }, [source, entry.page, entry.rotate, entry.key, scale, viewports]);

  const drawing = DRAW_TOOLS.has(tool);
  const local = (e) => {
    const box = ref.current.getBoundingClientRect();
    return [e.clientX - box.left, e.clientY - box.top];
  };
  const toPdf = (p) => vp.convertToPdfPoint(p[0], p[1]);

  const onDown = (e) => {
    if (!drawing || !vp || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const start = local(e);
    let points = [start];
    setDraft({ tool, start, end: start, points });
    const move = (ev) => {
      const p = local(ev);
      if (tool === 'ink') {
        const last = points[points.length - 1];
        if (Math.hypot(p[0] - last[0], p[1] - last[1]) >= 1.5) points = [...points, p];
      }
      setDraft({ tool, start, end: p, points });
    };
    const up = (ev) => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      const end = local(ev);
      setDraft(null);
      const moved = Math.hypot(end[0] - start[0], end[1] - start[1]) > 4;
      const a = toPdf(start);
      const b = toPdf(end);
      const box = [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])];
      const base = { key: entry.key, source: entry.source, page: entry.page, client: { x: ev.clientX, y: ev.clientY } };
      if (tool === 'note' || tool === 'date' || tool === 'check') onCreate({ ...base, tool, point: a });
      else if (tool === 'text') onCreate({ ...base, tool, rect: moved ? box : [a[0], a[1] - 40, a[0] + 200, a[1]] });
      else if (tool === 'stamp') onCreate({ ...base, tool, point: a, rect: moved ? box : null });
      else if (!moved) return;
      else if (tool === 'ink') onCreate({ ...base, tool, ink: [points.concat([end]).flatMap((p) => toPdf(p).map((v) => Math.round(v * 100) / 100))] });
      else if (tool === 'line' || tool === 'arrow') onCreate({ ...base, tool, line: [...a, ...b].map((v) => Math.round(v * 100) / 100) });
      else onCreate({ ...base, tool, rect: box.map((v) => Math.round(v * 100) / 100) });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  /** Pressing a comment picks it; dragging a movable one (or a corner of it) moves or resizes it. */
  const pressComment = (e, a, corner = null) => {
    if (drawing || e.button !== 0) return;
    e.stopPropagation();
    // Pressed on the page, it is in view already: the page is not scrolled under the pointer.
    onPick(a.id, entry.key, { onPage: true });
    if (!onMove || !vp || !MOVABLE_TYPES.has(a.type)) return;
    e.preventDefault();
    const start = local(e);
    const from = toPdf(start);
    const keepShape = a.type === 'Stamp';
    let rect = a.rect;
    let far = false;
    const move = (ev) => {
      const p = local(ev);
      if (!far && Math.hypot(p[0] - start[0], p[1] - start[1]) <= 3) return;
      far = true;
      rect = shiftedRect(a.rect, from, toPdf(p), { corner, keepShape });
      setShift({ id: a.id, rect });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      if (!far) { setShift(null); return; }
      const done = onMove({ id: a.id, source: entry.source, page: entry.page, rect: rect.map((v) => Math.round(v * 100) / 100) });
      // The outline stays at the new place until the page is drawn again with the comment there.
      Promise.resolve(done).finally(() => setShift(null));
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const mine = useMemo(() => (annots || []).filter((a) => a.page === entry.page && !a.hidden && !a.replyTo && a.rect), [annots, entry.page]);
  const links = useMemo(() => (source?.nav?.links || []).filter((l) => l.page === entry.page), [source, entry.page]);
  const stroke = cssColour(colour);
  return (
    <svg
      ref={ref}
      className={`pdf-annots${drawing ? ' drawing' : ''}`}
      data-tool={tool || undefined}
      data-ready={vp ? 'yes' : undefined}
      onMouseDown={onDown}
    >
      {vp && onLink ? links.map((l, i) => (
        <polygon
          key={`l${i}`}
          className="pdf-link"
          points={rectPts(vp, l.rect)}
          data-target={l.target ?? ''}
          data-url={l.url || ''}
          onClick={(e) => { if (drawing || e.button !== 0) return; e.stopPropagation(); onLink(l, entry); }}
        >
          <title>{l.url || t('Go to page {page}', { page: (l.target ?? 0) + 1 })}</title>
        </polygon>
      )) : null}
      {vp ? mine.filter((a) => !a.hasAppearance).map((a) => <Unpainted key={`u${a.id}`} a={a} vp={vp} />) : null}
      {vp ? mine.map((a) => {
        const movable = Boolean(onMove) && MOVABLE_TYPES.has(a.type);
        const at = shift?.id === a.id ? shift.rect : a.rect;
        return (
          <g key={a.id} className={`pdf-annot${picked === a.id ? ' picked' : ''}${movable ? ' movable' : ''}${shift?.id === a.id ? ' shifting' : ''}`} data-id={a.id} data-type={a.type} onMouseDown={(e) => pressComment(e, a)}>
            {a.quads ? a.quads.map((q, i) => <polygon key={i} points={quadPts(vp, q)} />) : <polygon points={rectPts(vp, at)} />}
            {movable && picked === a.id && RESIZABLE_TYPES.has(a.type) && !drawing ? [[at[0], at[1]], [at[2], at[1]], [at[2], at[3]], [at[0], at[3]]].map(([x, y], i) => {
              const [sx, sy] = vp.convertToViewportPoint(x, y);
              return <rect key={`h${i}`} className="pdf-handle" data-handle={i} x={sx - 4} y={sy - 4} width={8} height={8} onMouseDown={(e) => pressComment(e, a, i)} />;
            }) : null}
          </g>
        );
      }) : null}
      {fields}
      {draft && vp ? <Draft draft={draft} stroke={stroke} /> : null}
    </svg>
  );
}

/** A comment saved without an appearance, drawn from what it is. */
function Unpainted({ a, vp }) {
  const colour = cssColour(a.color || [0.86, 0.16, 0.16]);
  const scale = Math.hypot(vp.transform[0], vp.transform[1]) || 1;
  const width = Math.max(1, (a.width || 1) * scale);
  const common = { className: 'pdf-unpainted', fill: 'none', stroke: colour, strokeWidth: width, opacity: a.opacity ?? 1 };
  const at = (x, y) => vp.convertToViewportPoint(x, y);
  switch (a.type) {
    case 'Square': case 'Redact': return <polygon {...common} points={rectPts(vp, a.rect)} fill={a.interior ? cssColour(a.interior) : 'none'} />;
    case 'Circle': {
      const [x0, y0] = at(a.rect[0], a.rect[1]);
      const [x1, y1] = at(a.rect[2], a.rect[3]);
      return <ellipse {...common} cx={(x0 + x1) / 2} cy={(y0 + y1) / 2} rx={Math.abs(x1 - x0) / 2} ry={Math.abs(y1 - y0) / 2} fill={a.interior ? cssColour(a.interior) : 'none'} />;
    }
    case 'Line': return a.line ? <line {...common} x1={at(a.line[0], a.line[1])[0]} y1={at(a.line[0], a.line[1])[1]} x2={at(a.line[2], a.line[3])[0]} y2={at(a.line[2], a.line[3])[1]} /> : null;
    case 'Ink': return <g>{(a.ink || []).map((s, i) => <polyline key={i} {...common} strokeLinecap="round" strokeLinejoin="round" points={pts(vp, s)} />)}</g>;
    case 'Highlight': return <g style={{ mixBlendMode: 'multiply' }}>{(a.quads || []).map((q, i) => <polygon key={i} points={quadPts(vp, q)} fill={cssColour(a.color || [1, 0.92, 0.23])} stroke="none" />)}</g>;
    case 'Underline': case 'StrikeOut': case 'Squiggly': return (
      <g>{(a.quads || []).map((q, i) => {
        const t = a.type === 'StrikeOut' ? 0.52 : 0.94;
        const p = [q[0] + (q[4] - q[0]) * t, q[1] + (q[5] - q[1]) * t, q[2] + (q[6] - q[2]) * t, q[3] + (q[7] - q[3]) * t];
        const [ax, ay] = at(p[0], p[1]);
        const [bx, by] = at(p[2], p[3]);
        return <line key={i} {...common} strokeWidth={Math.max(1, scale)} x1={ax} y1={ay} x2={bx} y2={by} />;
      })}</g>
    );
    case 'Text': {
      const [x, y] = at(a.rect[0], a.rect[3]);
      const s = 20 * scale;
      return <rect className="pdf-unpainted" x={x} y={y} width={s} height={s * 0.85} rx={2} fill={cssColour(a.color || [1, 0.85, 0.25])} stroke="#555" strokeWidth={1} />;
    }
    case 'FreeText': return <polygon {...common} points={rectPts(vp, a.rect)} strokeWidth={1} />;
    default: return null;
  }
}

function Draft({ draft, stroke }) {
  const [x0, y0] = draft.start;
  const [x1, y1] = draft.end;
  const box = { x: Math.min(x0, x1), y: Math.min(y0, y1), width: Math.abs(x1 - x0), height: Math.abs(y1 - y0) };
  const common = { fill: 'none', stroke, strokeWidth: 2, className: 'pdf-draft' };
  switch (draft.tool) {
    case 'ink': return <polyline {...common} strokeLinecap="round" strokeLinejoin="round" points={draft.points.map((p) => p.join(',')).join(' ')} />;
    case 'ellipse': return <ellipse {...common} cx={box.x + box.width / 2} cy={box.y + box.height / 2} rx={box.width / 2} ry={box.height / 2} />;
    case 'line': case 'arrow': return <line {...common} x1={x0} y1={y0} x2={x1} y2={y1} />;
    case 'redact-area': return <rect {...common} stroke="#d91a1a" fill="rgba(217,26,26,0.12)" {...box} />;
    case 'note': case 'date': case 'check': return null;
    default: return <rect {...common} {...box} />;
  }
}

// ── typing a note or a text box ───────────────────────────────────────────

/** The box a sticky note's or a text box's words are typed into, by where it was placed. */
export function NoteEditor({ draft, onPost, onCancel }) {
  const [text, setText] = useState('');
  const ref = useRef(null);
  useEffect(() => { ref.current?.focus(); }, []);
  const title = draft.tool === 'text' ? t('Text box') : t('Sticky note');
  const left = Math.max(8, Math.min(window.innerWidth - 300, draft.client.x + 8));
  const top = Math.max(8, Math.min(window.innerHeight - 200, draft.client.y + 8));
  return (
    <div className="pdf-note-editor" style={{ left, top }} data-role="note-editor" role="dialog" aria-label={title}>
      <div className="pdf-note-editor-head"><Icon name={draft.tool === 'text' ? 'textbox' : 'note'} size={14} />{title}</div>
      <textarea
        ref={ref}
        value={text}
        rows={4}
        placeholder={draft.tool === 'text' ? t('Type the words for the page') : t('Type a comment')}
        data-role="note-text"
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Escape') { e.preventDefault(); onCancel(); }
          if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); if (text.trim()) onPost(text); }
        }}
      />
      <div className="pdf-note-editor-actions">
        <Button label={t('Cancel')} onClick={onCancel} />
        <Button primary label={t('Post')} title={t('Post (Ctrl+Enter)')} data-role="note-post" disabled={!text.trim()} onClick={() => onPost(text)} />
      </div>
    </div>
  );
}

// ── the Comments pane ─────────────────────────────────────────────────────

/**
 * Every comment on the pages in this window, in page order, each with its
 * replies: press one to go to it, reply, mark it resolved or delete it.
 * `items`: [{ annot, pageKey, number, source, replies }].
 */
export function CommentsPane({ items, picked, onPick, onReply, onResolve, onDelete, onEdit, onClose }) {
  const [filter, setFilter] = useState('all');
  const shown = items.filter((it) => filter === 'all' || (filter === 'open' ? it.annot.status !== 'Completed' : it.annot.status === 'Completed'));
  const listRef = useRef(null);
  useEffect(() => {
    const el = listRef.current?.querySelector(`.pdf-comment[data-id="${picked}"]`);
    if (!el) return;
    const box = listRef.current.getBoundingClientRect();
    const r = el.getBoundingClientRect();
    if (r.top < box.top || r.bottom > box.bottom) listRef.current.scrollTop += r.top - box.top - 12;
  }, [picked]);
  const resolved = items.filter((it) => it.annot.status === 'Completed').length;
  return (
    <aside className="pdf-comments" data-role="comments-pane" aria-label={t('Comments')}>
      <div className="pdf-comments-head">
        <Icon name="comment" size={15} />
        <span className="pdf-comments-title">{t('Comments')}</span>
        <span className="pdf-comments-count" data-count={items.length}>{items.length}</span>
        <span className="spacer" />
        <Button ghost icon="close" title={t('Close the Comments pane')} onClick={onClose} />
      </div>
      <div className="pdf-comments-filter" role="tablist">
        {[['all', t('All')], ['open', t('Open')], ['resolved', t('Resolved')]].map(([key, label]) => (
          <button key={key} type="button" role="tab" aria-selected={filter === key} className={filter === key ? 'on' : ''} data-filter={key} onClick={() => setFilter(key)}>
            {label}{key === 'resolved' && resolved ? <span className="pdf-comments-n">{resolved}</span> : null}
          </button>
        ))}
      </div>
      <div className="pdf-comments-list" ref={listRef}>
        {!shown.length ? (
          <div className="pdf-comments-empty">
            {items.length ? t('No comments here.') : t('No comments yet. Highlight words, or add a sticky note, a text box or a drawing from the Comment tab.')}
          </div>
        ) : shown.map((it) => (
          <CommentCard key={`${it.pageKey}:${it.annot.id}`} item={it} picked={picked === it.annot.id} onPick={onPick} onReply={onReply} onResolve={onResolve} onDelete={onDelete} onEdit={onEdit} />
        ))}
      </div>
    </aside>
  );
}

function CommentCard({ item, picked, onPick, onReply, onResolve, onDelete, onEdit }) {
  const { annot: a, number, replies } = item;
  const [reply, setReply] = useState('');
  const [editing, setEditing] = useState(null);
  const done = a.status === 'Completed';
  const swatchColour = a.type === 'FreeText' ? [0.55, 0.57, 0.62] : a.color;
  return (
    <div
      className={`pdf-comment${picked ? ' picked' : ''}${done ? ' resolved' : ''}`}
      data-id={a.id}
      data-type={a.type}
      data-status={a.status || 'open'}
      onMouseDown={(e) => { if (!e.target.closest('textarea, input, button')) onPick(a.id, item.pageKey); }}
    >
      <div className="pdf-comment-head">
        <span className="pdf-comment-kind" style={{ '--kind': cssColour(swatchColour) }}><Icon name={kindIcon(a)} size={14} /></span>
        <span className="pdf-comment-who">{a.author || t('Unknown author')}</span>
        <span className="pdf-comment-when" title={formatWhen(a.modified, { long: true })}>{formatWhen(a.modified)}</span>
        <span className="spacer" />
        <Button ghost icon="check" className="pdf-comment-resolve" pressed={done} title={done ? t('Reopen: mark this comment as not resolved') : t('Resolve: mark this comment as done')} onClick={() => onResolve(item, !done)} />
        <Button ghost icon="trash" className="pdf-comment-delete" title={t('Delete this comment and its replies')} onClick={() => onDelete(item)} />
      </div>
      <div className="pdf-comment-meta">
        <span>{kindName(a)}</span>
        <span>·</span>
        <span>{t('Page {page}', { page: number })}</span>
        {done ? <span className="pdf-comment-done"><Icon name="check" size={11} strokeWidth={2.4} />{t('Resolved')}</span> : null}
      </div>
      {editing !== null ? (
        <div className="pdf-comment-edit">
          <textarea value={editing} rows={3} autoFocus onChange={(e) => setEditing(e.target.value)} onKeyDown={(e) => {
            if (e.key === 'Escape') setEditing(null);
            if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { onEdit(item, editing); setEditing(null); }
          }} />
          <div className="pdf-comment-actions">
            <Button label={t('Cancel')} onClick={() => setEditing(null)} />
            <Button primary label={t('Save')} onClick={() => { onEdit(item, editing); setEditing(null); }} />
          </div>
        </div>
      ) : a.contents ? (
        <div className="pdf-comment-text" onDoubleClick={() => setEditing(a.contents)} title={t('Double-click to change the words')}>{a.contents}</div>
      ) : (
        <button type="button" className="pdf-comment-add" onClick={() => setEditing('')}>{t('Add a note…')}</button>
      )}
      {replies.length ? (
        <div className="pdf-comment-replies">
          {replies.map((r) => (
            <div key={r.id} className="pdf-reply" data-id={r.id}>
              <div className="pdf-reply-head"><span className="pdf-comment-who">{r.author || t('Unknown author')}</span><span className="pdf-comment-when">{formatWhen(r.modified)}</span></div>
              <div className="pdf-comment-text">{r.contents}</div>
            </div>
          ))}
        </div>
      ) : null}
      {picked ? (
        <div className="pdf-comment-reply">
          <input
            className="rw-input"
            value={reply}
            placeholder={t('Reply')}
            aria-label={t('Reply')}
            data-role="reply-text"
            onChange={(e) => setReply(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter' && reply.trim()) { onReply(item, reply); setReply(''); } }}
          />
          <Button icon="reply" title={t('Reply (Enter)')} data-role="reply-post" disabled={!reply.trim()} onClick={() => { onReply(item, reply); setReply(''); }} />
        </div>
      ) : null}
    </div>
  );
}

/** The pane's list: each top-level comment once for each page of the window it stands on, in page order. */
export function commentItems(pages, sources) {
  const out = [];
  pages.forEach((p, i) => {
    const list = sources[p.source]?.annots || [];
    for (const a of topLevel(list)) {
      if (a.page !== p.page) continue;
      out.push({ annot: a, pageKey: p.key, number: i + 1, source: p.source, replies: list.filter((r) => r.replyTo === a.id && !r.state && !r.hidden) });
    }
  });
  return out;
}

