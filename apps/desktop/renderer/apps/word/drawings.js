// Drawings on the page: text boxes, floating pictures and shapes, groups —
// where each stands, how it is drawn, and how a reader takes hold of one.
//
// Where a drawing stands is `@rutba/doc-view/floats` — the same answer the
// paginator gives the printout, so a box dragged to a place on screen prints
// there. How the words treat it decides where it is drawn:
//
//   beside   wrapped square, tight or through: its room is an empty CSS float
//            in the paragraph that anchors it, pushed down and across to its
//            place, the lines running round it (`spacerStyles`), and the
//            drawing itself is laid on the page over that room (`DrawingLayer`,
//            layer `beside`), so two that overlap are drawn overlapping;
//   block    top and bottom, or centred: a block under the paragraph's words;
//   behind / front   no wrap: laid on the page itself, under or over every
//            paragraph (`DrawingLayer`), where its anchor's place says.
//
// A text box's words are paragraphs the engine lists in the edit space (see
// document.js `editParagraphs`), so the box draws them as blocks and the
// caret types in them like anywhere else.

import React from 'react';
import { Button, Icon, Spacer } from '@rutba/office-ui';
import { floatPlace, layerOf, sideOf } from '@rutba/doc-view/floats';
import { rectOf, pageTopOf, pageIndexAt } from './pages.js';

/** The layer a drawing lives in; a centred square-wrapped one stands under the words, as before. */
export function drawingLayer(d) {
  if (!d || d.hidden) return 'hidden';
  const layer = layerOf(d);
  if (layer === 'beside' && d.hAlign === 'center') return 'block';
  return layer;
}

/**
 * A drawing's place in the order of drawings on the page, as a CSS z-index:
 * Word's `relativeHeight` counted up from its first drawing's, so a drawing
 * brought to the front is drawn over the ones it overlaps.
 */
export function zOf(d) {
  const z = Number(d?.relativeHeight);
  return Number.isFinite(z) ? Math.max(1, Math.min(2000000000, Math.round(z - 251658240) + 1)) : 1;
}

/** The page's geometry in the terms floats.js asks for. */
export function geomOf(section, columnBoxes = null) {
  const m = section?.margins || { left: 96, right: 96, top: 96, bottom: 96 };
  const pageWidthPx = section?.widthPx || 816;
  const contentWidthPx = pageWidthPx - m.left - m.right - (m.gutter || 0);
  return {
    marginLeftPx: m.left + (m.gutter || 0), marginRightPx: m.right, marginTopPx: m.top, marginBottomPx: m.bottom,
    contentWidthPx, pageWidthPx, pageHeightPx: section?.heightPx || 1056,
    columnWidthPx: columnBoxes?.[0]?.widthPx ?? contentWidthPx,
  };
}

/** A drawing's turn and mirror as CSS — a text box's words turn with it but are never mirrored. */
export function turnCss(d, text = false) {
  const parts = [];
  if (d?.rot) parts.push(`rotate(${Math.round(d.rot * 100) / 100}deg)`);
  if (!text && (d?.flipH || d?.flipV)) parts.push(`scale(${d.flipH ? -1 : 1}, ${d.flipV ? -1 : 1})`);
  return parts.length ? parts.join(' ') : undefined;
}

/**
 * A drawing beside the words, as a float in its paragraph: at its side of
 * the column, pushed across and down to its place. `shape-outside` keeps
 * the lines above its top at the full width, so the words wrap only where
 * the drawing is. `anchorAt` is the paragraph measured on the page: its
 * top within its sheet, and how far its sides stand in from the column's.
 */
export function besideCss(d, g, anchorAt = null) {
  const place = floatPlace(d, g);
  const dist = d.dist || {};
  const col = g.columnWidthPx;
  const side = sideOf(place.x, place.widthPx, col);
  const top = besideTop(d, g, anchorAt);
  const style = {
    float: side,
    marginTop: top,
    marginBottom: Math.round(dist.b ?? 6),
    shapeOutside: `inset(${Math.max(0, top - Math.round(dist.t || 0))}px 0 0 0)`,
  };
  if (side === 'left') {
    style.marginLeft = Math.round(place.x - (anchorAt?.left ?? 0));
    style.marginRight = Math.round(dist.r ?? 12);
  } else {
    style.marginRight = Math.round(col - place.x - place.widthPx - (anchorAt?.right ?? 0));
    style.marginLeft = Math.round(dist.l ?? 12);
  }
  return style;
}

/** How far down from its paragraph's top a drawing beside the words stands — with the room it keeps above itself. */
export function besideTop(d, g, anchorAt = null) {
  const place = floatPlace(d, g);
  const down = place.yFrom === 'page' ? place.y - (anchorAt?.inPage ?? place.y) : place.y;
  return Math.max(0, Math.round(down)) + Math.round(d.dist?.t || 0);
}

/**
 * The room a drawing beside the words takes in its paragraph: an empty float
 * of its size at its place, which the lines run round. The drawing itself is
 * laid on the page over it (`DrawingLayer`, layer `beside`), so two that
 * overlap are drawn overlapping, in their order, as Word draws them — floats
 * would have been stacked side by side.
 */
export function spacerCss(d, g, anchorAt = null) {
  const place = floatPlace(d, g);
  return { ...besideCss(d, g, anchorAt), width: Math.max(1, Math.round(place.widthPx)), height: Math.max(1, Math.round(place.heightPx)), display: 'block', visibility: 'hidden', pointerEvents: 'none' };
}

/**
 * The room the drawings beside a paragraph's words take, one float to a
 * side: the box round all the ones on that side. Two floats on one side
 * would stand side by side, and the lines would run round both as if they
 * did; the printout keeps the widest reach on each side too (paginate.js).
 */
export function spacerStyles(list, g, anchorAt = null) {
  const sides = { left: [], right: [] };
  for (const d of list) {
    const place = floatPlace(d, g);
    const side = sideOf(place.x, place.widthPx, g.columnWidthPx);
    sides[side].push({ d, x: place.x, w: place.widthPx, top: besideTop(d, g, anchorAt), h: place.heightPx });
  }
  const out = [];
  for (const side of ['left', 'right']) {
    const items = sides[side];
    if (!items.length) continue;
    const L = Math.min(...items.map((i) => i.x));
    const R = Math.max(...items.map((i) => i.x + i.w));
    const T = Math.min(...items.map((i) => i.top));
    const B = Math.max(...items.map((i) => i.top + i.h));
    const topmost = items.find((i) => i.top === T).d;
    const dist = topmost.dist || {};
    const style = {
      float: side, display: 'block', visibility: 'hidden', pointerEvents: 'none',
      width: Math.max(1, Math.round(R - L)), height: Math.max(1, Math.round(B - T)),
      marginTop: Math.round(T), marginBottom: Math.round(dist.b ?? 6),
      shapeOutside: `inset(${Math.max(0, Math.round(T - (dist.t || 0)))}px 0 0 0)`,
    };
    if (side === 'left') {
      style.marginLeft = Math.round(L - (anchorAt?.left ?? 0));
      style.marginRight = Math.round(dist.r ?? 12);
    } else {
      style.marginRight = Math.round(g.columnWidthPx - R - (anchorAt?.right ?? 0));
      style.marginLeft = Math.round(dist.l ?? 12);
    }
    out.push({ side, style, ids: items.map((i) => i.d.id) });
  }
  return out;
}

/** A drawing under the words (top and bottom, centred): a block at its place across. */
export function blockCss(d, g) {
  if (d.hAlign === 'center') return { display: 'block', margin: '6px auto' };
  const place = floatPlace(d, g);
  return { display: 'block', margin: `6px 0 6px ${Math.round(place.x)}px` };
}

/**
 * Where a drawing laid over or under the words stands on the page, from the
 * page's top-left: across from the margin, down from its paragraph (or its
 * sheet, when the anchor says so).
 */
export function overlayPlace(d, g, anchorAt, geo) {
  const place = floatPlace(d, g);
  const sheetTop = anchorAt ? anchorAt.sheetTop : 0;
  const anchorTop = anchorAt ? anchorAt.top : g.marginTopPx;
  // Beside the words: where its room in the paragraph is.
  const top = drawingLayer(d) === 'beside' ? anchorTop + besideTop(d, g, anchorAt)
    : place.yFrom === 'page' ? sheetTop + place.y : anchorTop + place.y;
  return { left: Math.round(g.marginLeftPx + place.x), top: Math.round(top), width: Math.round(place.widthPx), height: Math.round(place.heightPx) };
}

/**
 * Every paragraph that anchors a drawing laid on the page or placed from
 * the sheet, measured: its top on the page, its top within its sheet, and
 * how far its sides stand in from the column. Measured after each layout
 * pass, since a pass moves paragraphs from sheet to sheet.
 */
export function measureAnchors(page, blocks, geo, paged, g) {
  const out = {};
  if (!page) return out;
  const pageRect = rectOf(page);
  for (const b of blocks || []) {
    const all = [...(b.images || []), ...(b.textBoxes || []), ...(b.groups || [])];
    if (!all.some((d) => d.anchored)) continue;
    const el = page.querySelector(`.wd-block[data-block="${b.index}"]:not([data-part="1"])`) || page.querySelector(`[data-block="${b.index}"]`);
    if (!el) continue;
    const r = rectOf(el);
    const cs = getComputedStyle(el);
    const top = r.top - pageRect.top;
    const k = paged && geo ? pageIndexAt(geo, top + 1) : 0;
    const sheetTop = paged && geo ? pageTopOf(geo, k) : 0;
    out[b.index] = {
      top: Math.round(top * 10) / 10,
      sheetTop,
      inPage: Math.round((top - sheetTop) * 10) / 10,
      left: Math.round(r.left + parseFloat(cs.paddingLeft || '0') - (pageRect.left + g.marginLeftPx)),
      right: Math.round(pageRect.left + g.marginLeftPx + g.columnWidthPx - (r.right - parseFloat(cs.paddingRight || '0'))),
    };
  }
  return out;
}

/* ── how each kind of drawing is drawn ──────────────────────────────────── */

const PICK_KEYS = ['block', 'image', 'id', 'kind'];

/** Tell the page a drawing was clicked — it keeps the selection. */
function pickEvent(e, detail) {
  e.currentTarget.dispatchEvent(new CustomEvent('wd-pick', { bubbles: true, detail: { ...Object.fromEntries(PICK_KEYS.map((k) => [k, detail[k] ?? null])), add: Boolean(e.shiftKey || e.ctrlKey || e.metaKey) } }));
}

/**
 * A text box: its frame — the fill, the outline, the margins, where its
 * words sit up and down, their direction, its turn — and its words: the
 * edit space's paragraphs (`kids`, drawn by the page's own `renderBlock`, so
 * the caret types in them) or, for a box the edit space leaves opaque, its
 * paragraphs as the file has them.
 */
export function TextBox({ box, block = null, kids = null, renderBlock = null, renderLite = null, style = null, className = '', picked = false }) {
  const ins = box.insets || { l: 9.6, t: 4.8, r: 9.6, b: 4.8 };
  const vertical = box.vert === 'vert' || box.vert === 'vert270';
  const editable = Boolean(kids && kids.length && renderBlock);
  const frame = {
    boxSizing: 'border-box',
    width: box.widthPx ? Math.round(box.widthPx) : undefined,
    minHeight: box.heightPx ? Math.round(box.heightPx) : undefined,
    ...(vertical ? { height: box.heightPx ? Math.round(box.heightPx) : undefined, writingMode: box.vert === 'vert270' ? 'sideways-lr' : 'vertical-rl' } : {}),
    backgroundColor: box.fill || 'transparent',
    border: box.line ? `${Math.max(1, Math.round(box.lineWidthPx || 1))}px solid ${box.line}` : '0',
    padding: `${Math.round(ins.t)}px ${Math.round(ins.r)}px ${Math.round(ins.b)}px ${Math.round(ins.l)}px`,
    display: 'flex', flexDirection: 'column',
    justifyContent: box.vAnchor === 'middle' ? 'center' : box.vAnchor === 'bottom' ? 'flex-end' : 'flex-start',
    transform: turnCss(box, true),
    ...(style || {}),
  };
  return (
    <div
      className={`wd-textbox wd-drawing${editable ? ' editable' : ''}${picked ? ' picked' : ''}${className ? ` ${className}` : ''}`}
      data-drawing={box.id ?? undefined}
      data-kind="textbox"
      contentEditable={editable ? undefined : false}
      suppressContentEditableWarning
      style={frame}
      onMouseDown={(e) => {
        // A press on the frame itself — its margin, not its words — takes
        // hold of the box, as a click on a text box's edge does in Word.
        if (e.button !== 0 || e.target !== e.currentTarget || box.id == null) return;
        e.preventDefault();
        pickEvent(e, { block: block ?? null, id: box.id, kind: 'textbox' });
      }}
    >
      {editable ? kids.map((k) => renderBlock(k)) : renderLite ? renderLite(box.paragraphs || []) : null}
    </div>
  );
}

/**
 * A group: one box, its members placed in it — pictures, shapes (painted by
 * the engine), text boxes with their words — each turned as it was.
 */
export function GroupBox({ group, block = null, kidsOf = null, renderBlock = null, renderLite = null, style = null, picked = false }) {
  return (
    <div
      className={`wd-group wd-drawing${picked ? ' picked' : ''}`}
      data-drawing={group.id ?? undefined}
      data-kind="group"
      contentEditable={false}
      suppressContentEditableWarning
      style={{ position: 'relative', width: Math.round(group.widthPx), height: Math.round(group.heightPx), transform: turnCss(group), ...(style || {}) }}
      onMouseDown={(e) => {
        if (e.button !== 0) return;
        if (e.target.closest?.('.wd-block')) return; // into a member box's words
        e.preventDefault();
        pickEvent(e, { block, id: group.id, kind: 'group' });
      }}
    >
      {(group.members || []).map((m, i) => {
        const at = { position: 'absolute', left: Math.round(m.xPx), top: Math.round(m.yPx), width: Math.round(m.widthPx), height: Math.round(m.heightPx) };
        if (m.kind === 'textbox') {
          const kids = kidsOf ? kidsOf(m) : null;
          return (
            <div key={i} style={at} contentEditable={kids && kids.length ? true : false} suppressContentEditableWarning>
              <TextBox box={{ ...m, widthPx: m.widthPx, heightPx: m.heightPx, id: null }} kids={kids} renderBlock={renderBlock} renderLite={renderLite} style={{ width: '100%', minHeight: '100%' }} />
            </div>
          );
        }
        return m.href
          ? <img key={i} className="wd-member" src={m.href} alt={m.name || ''} draggable={false} style={{ ...at, transform: turnCss(m) }} />
          : <div key={i} className="wd-member wd-member-empty" style={at} />;
      })}
    </div>
  );
}

/**
 * The drawings laid on the page rather than in the words — behind them or in
 * front — each at its place, under every paragraph (`behind`, above the
 * sheets) or over them (`front`).
 */
export function DrawingLayer({ layer, blocks, g, anchors, geo, pickedIds, kidsOf, renderBlock, renderLite }) {
  const items = [];
  for (const b of blocks || []) {
    if (b.box) continue;
    const at = anchors?.[b.index] ?? null;
    (b.images || []).forEach((img, i) => {
      if (drawingLayer(img) !== layer) return;
      // Beside the words: at most the column (and 640 px), as the printout draws it.
      const scale = layer === 'beside' ? Math.min(1, Math.min(g.columnWidthPx, 640) / Math.max(1, img.widthPx || 1)) : 1;
      const pos = overlayPlace({ ...img, widthPx: (img.widthPx || 0) * scale, heightPx: (img.heightPx || 0) * scale }, g, at, geo);
      items.push(
        <img
          key={`i${b.index}:${i}`}
          className={`wd-image wd-drawing wd-over ${layer}${layer === 'beside' ? ' wd-float' : ''}${pickedIds?.includes(img.id) ? ' picked' : ''}`}
          data-drawing={img.id ?? undefined}
          data-image={i}
          data-anchor={b.index}
          src={img.href}
          alt={img.name || ''}
          draggable={false}
          style={{ left: pos.left, top: pos.top, width: pos.width, height: pos.height, transform: turnCss(img), zIndex: zOf(img) }}
          // The press keeps the caret where it is; the click picks, as a picture in the words does.
          onMouseDown={(e) => { if (e.button === 0) e.preventDefault(); }}
          onClick={(e) => { e.stopPropagation(); pickEvent(e, { block: b.index, image: i, id: img.id, kind: img.kind || 'picture' }); }}
        />
      );
    });
    (b.textBoxes || []).forEach((box, i) => {
      if (drawingLayer(box) !== layer) return;
      const pos = overlayPlace(box, g, at, geo);
      items.push(
        <div key={`t${b.index}:${i}`} className={`wd-over-host ${layer}`} data-anchor={b.index} style={{ left: pos.left, top: pos.top, zIndex: zOf(box) }}>
          <TextBox box={box} block={b.index} kids={kidsOf(box)} renderBlock={renderBlock} renderLite={renderLite} picked={pickedIds?.includes(box.id)} />
        </div>
      );
    });
    (b.groups || []).forEach((grp, i) => {
      if (drawingLayer(grp) !== layer) return;
      const pos = overlayPlace(grp, g, at, geo);
      items.push(
        <div key={`g${b.index}:${i}`} className={`wd-over-host ${layer}`} data-anchor={b.index} style={{ left: pos.left, top: pos.top, zIndex: zOf(grp) }}>
          <GroupBox group={grp} block={b.index} kidsOf={kidsOf} renderBlock={renderBlock} renderLite={renderLite} picked={pickedIds?.includes(grp.id)} />
        </div>
      );
    });
  }
  if (!items.length) return null;
  return <div className={`wd-drawlayer ${layer}`} suppressContentEditableWarning>{items}</div>;
}

/* ── taking hold of a drawing ───────────────────────────────────────────── */

const HANDLES = [
  ['nw', 0, 0, 'nwse-resize'], ['n', 0.5, 0, 'ns-resize'], ['ne', 1, 0, 'nesw-resize'], ['e', 1, 0.5, 'ew-resize'],
  ['se', 1, 1, 'nwse-resize'], ['s', 0.5, 1, 'ns-resize'], ['sw', 0, 1, 'nesw-resize'], ['w', 0, 0.5, 'ew-resize'],
];

/** A drawing's unturned box on the page (page px), from its element: the centre of what is drawn and its own size. */
export function boxOnPage(page, el) {
  if (!page || !el) return null;
  const p = rectOf(page);
  const r = rectOf(el);
  const w = el.offsetWidth || r.width;
  const h = el.offsetHeight || r.height;
  const cx = r.left - p.left + r.width / 2;
  const cy = r.top - p.top + r.height / 2;
  return { left: cx - w / 2, top: cy - h / 2, width: w, height: h };
}

/**
 * The selected drawings' frames: for one, eight handles to size it by and a
 * round one above to turn it by, and its edge to drag it by; for several, a
 * frame round each. A text box being typed in wears a dashed frame, as in
 * Word. Everything happens on the page first — the drawing follows the
 * pointer — and the engine is told once, on release.
 */
export function DrawingFrame({ page, ids, editingId = null, drawings, deps, onMove, onResize, onRotate, onDrag, zoom = 1 }) {
  const [boxes, setBoxes] = React.useState([]);
  const [live, setLive] = React.useState(null);
  // A press on a selected drawing's body starts a move — heard on the page
  // before the drawing's own pick, so a pick and a drag are one gesture.
  const moveRef = React.useRef(null);
  React.useEffect(() => {
    const el = page.current;
    if (!el) return undefined;
    const down = (e) => { if (e.button === 0 && moveRef.current?.(e)) e.stopPropagation(); };
    el.addEventListener('mousedown', down, true);
    return () => el.removeEventListener('mousedown', down, true);
  }, [page]);
  const shown = editingId != null ? [editingId] : ids;
  const find = React.useCallback((id) => page.current?.querySelector(`.wd-drawing[data-drawing="${id}"]`) || null, [page]);
  const measure = React.useCallback(() => {
    if (!page.current) return;
    const next = shown.map((id) => {
      const el = find(id);
      const b = el ? boxOnPage(page.current, el) : null;
      const d = drawings.find((x) => x.id === id);
      return b ? { id, ...b, rot: d?.rot || 0, anchored: Boolean(d?.anchored), kind: d?.kind || el?.dataset.kind || 'picture' } : null;
    }).filter(Boolean);
    setBoxes((was) => (JSON.stringify(was) === JSON.stringify(next) ? was : next));
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [page, shown.join(','), drawings, find]);
  React.useLayoutEffect(() => { measure(); }, [measure, ...deps]);
  React.useEffect(() => {
    const el = page.current;
    if (!el || typeof ResizeObserver === 'undefined') return undefined;
    const ro = new ResizeObserver(() => measure());
    ro.observe(el);
    return () => ro.disconnect();
  }, [measure, page]);
  if (!boxes.length) return null;
  const single = boxes.length === 1 ? boxes[0] : null;

  /** A drag of any kind: the pointer's travel in page px, handed to `step` and then to `done`. */
  const drag = (e, step, done) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const x0 = e.clientX;
    const y0 = e.clientY;
    onDrag?.(true);
    let last = { dx: 0, dy: 0, shift: false };
    const move = (ev) => {
      last = { dx: (ev.clientX - x0) / zoom, dy: (ev.clientY - y0) / zoom, shift: ev.shiftKey };
      step(last);
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setTimeout(() => onDrag?.(false), 0);
      done(last);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const startMove = (e, b) => {
    if (!b.anchored) return;
    const el = find(b.id);
    const was = el?.style.translate || '';
    drag(e, ({ dx, dy }) => {
      if (el) el.style.translate = `${Math.round(dx)}px ${Math.round(dy)}px`;
      setLive({ id: b.id, left: b.left + dx, top: b.top + dy, width: b.width, height: b.height, rot: b.rot });
    }, ({ dx, dy }) => {
      if (el) el.style.translate = was;
      setLive(null);
      if (Math.abs(dx) < 2 && Math.abs(dy) < 2) return;
      onMove?.(b.id, { left: b.left + dx, top: b.top + dy, width: b.width, height: b.height });
    });
  };

  const startResize = (e, b, name) => {
    const el = find(b.id);
    const keepRatio = b.kind === 'picture' || b.kind === 'group';
    const ratio = b.height / Math.max(1, b.width);
    const was = { width: el?.style.width || '', height: el?.style.height || '', minHeight: el?.style.minHeight || '' };
    const sized = ({ dx, dy, shift }) => {
      // In the drawing's own turned frame: the pointer's travel turned back.
      const rad = (-(b.rot || 0) * Math.PI) / 180;
      const ux = dx * Math.cos(rad) - dy * Math.sin(rad);
      const uy = dx * Math.sin(rad) + dy * Math.cos(rad);
      let w = b.width + (name.includes('e') ? ux : name.includes('w') ? -ux : 0);
      let h = b.height + (name.includes('s') ? uy : name.includes('n') ? -uy : 0);
      const corner = name.length === 2;
      if ((keepRatio && corner) !== Boolean(shift && corner)) h = w * ratio;
      w = Math.max(12, Math.round(w));
      h = Math.max(12, Math.round(h));
      const left = b.left + (name.includes('w') ? b.width - w : 0);
      const top = b.top + (name.includes('n') ? b.height - h : 0);
      return { left, top, width: w, height: h };
    };
    drag(e, (t) => {
      const s = sized(t);
      if (el) {
        el.style.width = `${s.width}px`;
        if (b.kind === 'textbox') el.style.minHeight = `${s.height}px`; else el.style.height = `${s.height}px`;
      }
      setLive({ id: b.id, ...s, rot: b.rot });
    }, (t) => {
      if (el) { el.style.width = was.width; el.style.height = was.height; el.style.minHeight = was.minHeight; }
      setLive(null);
      if (Math.abs(t.dx) < 1 && Math.abs(t.dy) < 1) return;
      onResize?.(b.id, sized(t), { moved: name.includes('w') || name.includes('n'), from: b });
    });
  };

  const startRotate = (e, b) => {
    const el = find(b.id);
    const was = el?.style.rotate || '';
    const p = rectOf(page.current);
    const cx = b.left + b.width / 2;
    const cy = b.top + b.height / 2;
    const angleAt = (ev) => {
      const x = ev.clientX / zoom - p.left - cx;
      const y = ev.clientY / zoom - p.top - cy;
      return (Math.atan2(y, x) * 180) / Math.PI + 90;
    };
    let deg = b.rot || 0;
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    onDrag?.(true);
    const move = (ev) => {
      let a = angleAt(ev);
      if (ev.shiftKey) a = Math.round(a / 15) * 15;
      deg = ((Math.round(a) % 360) + 360) % 360;
      if (el) el.style.rotate = `${deg - (b.rot || 0)}deg`;
      setLive({ id: b.id, left: b.left, top: b.top, width: b.width, height: b.height, rot: deg });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setTimeout(() => onDrag?.(false), 0);
      if (el) el.style.rotate = was;
      setLive(null);
      if (deg !== (b.rot || 0)) onRotate?.(b.id, deg);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const frameOf = (b, main) => {
    const g = live?.id === b.id ? live : b;
    return (
      <div
        key={b.id}
        className={`wd-frame-sel${editingId === b.id ? ' editing' : ''}${main ? ' main' : ''}`}
        data-frame={b.id}
        style={{ left: g.left, top: g.top, width: g.width, height: g.height, transform: g.rot ? `rotate(${g.rot}deg)` : undefined }}
      >
        {/* The edge takes hold of a floating drawing to move it. */}
        {b.anchored ? ['t', 'r', 'b', 'l'].map((s) => <div key={s} className={`wd-frame-edge ${s}`} onMouseDown={(e) => startMove(e, b)} />) : null}
        {main ? HANDLES.map(([name, fx, fy, cursor]) => (
          <div key={name} className="wd-handle" data-handle={name} style={{ left: `calc(${fx * 100}% - 5px)`, top: `calc(${fy * 100}% - 5px)`, cursor }} onMouseDown={(e) => startResize(e, b, name)} />
        )) : null}
        {main && b.anchored ? (
          <>
            <div className="wd-rotate-stem" />
            <div className="wd-rotate" data-handle="rotate" title="Turn — hold Shift for steps of 15°" onMouseDown={(e) => startRotate(e, b)}>
              <Icon name="rotate" size={12} />
            </div>
          </>
        ) : null}
      </div>
    );
  };

  moveRef.current = (e) => {
    // The body of a selected floating picture, shape or group moves it; a
    // text box moves by its edge, since a press in its words is the caret's.
    const el = e.target.closest?.('.wd-drawing');
    const b = el ? boxes.find((x) => String(x.id) === el.dataset.drawing) : null;
    if (!b || !b.anchored || editingId === b.id) return false;
    if (b.kind === 'textbox' && e.target !== el) return false;
    startMove(e, b);
    return true;
  };

  return (
    <div className="wd-frames" contentEditable={false} aria-hidden="true" suppressContentEditableWarning>
      {boxes.map((b) => frameOf(b, Boolean(single)))}
    </div>
  );
}

/* ── the Selection Pane ─────────────────────────────────────────────────── */

const KIND_ICONS = { picture: 'picture', textbox: 'textbox', shape: 'shape', group: 'grid', chart: 'chart' };

/**
 * Layout → Selection Pane: every drawing, the one on top first — Word's
 * pane, in the suite's look (the Presentation's Layers pane): a click
 * selects (Shift or Ctrl adds one), the eye hides or shows it, a
 * double-click renames it, the arrows change the order, and Show All /
 * Hide All do what they say. A group's members are listed under it.
 */
export function SelectionPane({ drawings = [], picked = [], onPick, onToggle, act, onClose }) {
  const [renaming, setRenaming] = React.useState(null);
  const floating = drawings.filter((d) => d.anchored);
  const inline = drawings.filter((d) => !d.anchored);
  // Top-most first: the order of the heights, highest first; then the ones in the line.
  const rows = [...floating].sort((a, b) => (b.relativeHeight ?? 0) - (a.relativeHeight ?? 0)).concat(inline);
  const one = picked.length === 1 ? drawings.find((d) => d.id === picked[0]) : null;
  const ordered = [...floating].sort((a, b) => (a.relativeHeight ?? 0) - (b.relativeHeight ?? 0));
  const pos = one ? ordered.findIndex((d) => d.id === one.id) : -1;
  const row = (d, indent = 0, member = false) => (
    <div
      key={`${member ? 'm' : 'd'}${d.id}`}
      className={`wd-layer${picked.includes(d.id) ? ' active' : ''}${d.hidden ? ' off' : ''}`}
      data-drawing-row={d.id}
      style={indent ? { paddingLeft: 12 + indent * 16 } : undefined}
      onClick={(e) => (member ? null : (e.shiftKey || e.ctrlKey || e.metaKey) ? onToggle(d.id) : onPick(d.id))}
      onDoubleClick={() => (member ? null : setRenaming({ id: d.id, name: d.name || '' }))}
      title={member ? 'A member of the group above' : 'Click to select (Shift or Ctrl adds one); double-click to rename'}
    >
      {member ? <span className="wd-eye-space" /> : (
        <button
          type="button"
          className="wd-eye"
          title={d.hidden ? 'Show this drawing' : 'Hide this drawing'}
          onClick={(e) => { e.stopPropagation(); act('hideDrawing', { id: d.id, hidden: !d.hidden }); }}
        >
          <Icon name="eye" size={14} />
        </button>
      )}
      <Icon name={KIND_ICONS[d.kind] || 'shape'} size={14} />
      {renaming?.id === d.id ? (
        <input
          autoFocus
          className="rw-input wd-layer-name"
          value={renaming.name}
          onChange={(e) => setRenaming({ ...renaming, name: e.target.value })}
          onKeyDown={(e) => {
            if (e.key === 'Enter') { act('renameDrawing', { id: d.id, name: renaming.name }); setRenaming(null); }
            if (e.key === 'Escape') setRenaming(null);
          }}
          onBlur={() => setRenaming(null)}
          onClick={(e) => e.stopPropagation()}
        />
      ) : (
        <span className="wd-layer-text">
          <span className="wd-layer-title">{d.name || `${d.kind} ${d.id}`}</span>
          {d.text ? <span className="wd-layer-words">{d.text}</span> : !d.anchored && !member ? <span className="wd-layer-words">In line with text</span> : null}
        </span>
      )}
    </div>
  );
  return (
    <aside className="wd-selpane">
      <div className="wd-nav-head">
        <span>Selection</span>
        <button type="button" className="wd-nav-close" onClick={onClose} title="Close the selection pane" aria-label="Close">
          <Icon name="close" size={12} />
        </button>
      </div>
      <div className="wd-layers-tools">
        <Button label="Show All" title="Show All — every drawing drawn again" disabled={!drawings.some((d) => d.hidden)} onClick={() => act('showAllDrawings', true)} />
        <Button label="Hide All" title="Hide All — every drawing kept in the file but not drawn" disabled={!drawings.some((d) => !d.hidden)} onClick={() => act('showAllDrawings', false)} />
        <Spacer />
        <Button icon="chevronUp" title="Bring Forward" disabled={pos < 0 || pos >= ordered.length - 1} onClick={() => act('order', 'forward')} />
        <Button icon="chevronDown" title="Send Backward" disabled={pos <= 0} onClick={() => act('order', 'backward')} />
      </div>
      <div className="wd-layers">
        {rows.length ? rows.map((d) => (
          <React.Fragment key={d.id}>
            {row(d)}
            {/* A group's members, the one on top first, as the drawings above them. */}
            {d.kind === 'group' ? (d.members || []).map((m, i) => ({ m, i })).reverse().map(({ m, i }) => row({ ...m, id: m.id ?? `${d.id}.${i}`, hidden: d.hidden }, 1, true)) : null}
          </React.Fragment>
        )) : (
          <p className="wd-nav-empty">No pictures, shapes or text boxes in this document yet. Insert → Text Box puts one in.</p>
        )}
      </div>
    </aside>
  );
}

/* ── Insert → Text Box: built-in boxes of our own ───────────────────────── */

/**
 * The built-in text boxes, as specs for the engine's `insertTextBox`: a plain
 * box, a sidebar down the right-hand side, a pull quote across the column.
 * Our own designs in the document's own colours — Word's built-ins are its
 * building blocks, and these are not copies of them.
 */
export function textBoxPresets(g) {
  const col = g.columnWidthPx;
  return {
    simple: {
      name: 'Simple Text Box', widthPx: Math.round(Math.min(288, col * 0.46)), heightPx: 96, autoFit: true,
      h: { rel: 'column', offsetPx: 0 }, v: { rel: 'paragraph', offsetPx: 0 }, fill: 'FFFFFF', line: '000000', lineWidthPx: 1,
      paragraphs: [{ text: 'Type your words here. Drag the box by its edge to put it anywhere on the page.' }],
    },
    sidebar: {
      name: 'Sidebar', widthPx: Math.round(Math.min(220, col * 0.36)), heightPx: 420, autoFit: false,
      h: { rel: 'margin', align: 'right' }, v: { rel: 'paragraph', offsetPx: 0 },
      fill: 'EEF3F8', line: null, insets: { l: 16, t: 16, r: 16, b: 16 },
      paragraphs: [
        { text: 'Sidebar title', bold: true, sizePt: 14, colour: '1F3864', afterTwips: 120 },
        { text: 'A sidebar stands beside the story with the words that go with it: a summary, a list of points, a note for the reader.', sizePt: 10.5, colour: '333333' },
      ],
    },
    quote: {
      name: 'Pull Quote', widthPx: Math.round(col * 0.8), heightPx: 90, autoFit: true, wrap: 'topAndBottom',
      h: { rel: 'margin', align: 'center' }, v: { rel: 'paragraph', offsetPx: 0 },
      fill: null, line: '2B5FD9', lineWidthPx: 2, insets: { l: 18, t: 12, r: 18, b: 12 }, vAnchor: 'middle',
      paragraphs: [{ text: '“Put the line your reader should remember here.”', italic: true, sizePt: 16, colour: '2B5FD9', align: 'center' }],
    },
  };
}

export const DRAWING_CSS = `
/* Drawings on the page ------------------------------------------------------ */
.wd-drawing { cursor: default; }
/* A selected drawing wears its frame and handles; the old outline would draw a second frame round it. */
.wd-drawing[data-drawing].picked { outline: none; }
.wd-image.wd-drawing.wd-float, .wd-textbox.wd-float, .wd-group.wd-float { position: relative; z-index: 1; }
.wd-textbox { user-select: none; overflow: hidden; }
.wd-textbox.editable { user-select: text; cursor: text; }
.wd-textbox .wd-block { margin: 0; }
.wd-textbox.editable .wd-block:last-child { margin-bottom: 0; }
.wd-group { user-select: none; }
.wd-member { position: absolute; display: block; object-fit: fill; }
.wd-member-empty { outline: 1px dashed rgba(70, 120, 200, 0.35); }
/* The layers on the page: under every paragraph (above the sheet) and over them. */
.wd-drawlayer { position: absolute; left: 0; top: 0; width: 0; height: 0; }
.wd-drawlayer.behind { z-index: 0; }
.wd-drawlayer.beside { z-index: 1; }
.wd-drawlayer.front { z-index: 2; }
.wd-float-spacer { float: left; }
.wd-drawlayer > .wd-over, .wd-drawlayer > .wd-over-host { position: absolute; }
.wd-drawlayer > .wd-over-host > .wd-drawing { margin: 0; }
/* The selection: a frame, eight handles and a turn handle, Word's layout in the suite's look. */
.wd-frames { position: absolute; left: 0; top: 0; width: 0; height: 0; z-index: 4; }
.wd-frame-sel { position: absolute; box-sizing: border-box; border: 1px solid var(--accent); pointer-events: none; transform-origin: 50% 50%; }
.wd-frame-sel.editing { border-style: dashed; }
.wd-frame-edge { position: absolute; pointer-events: auto; cursor: move; }
.wd-frame-edge.t { left: 0; right: 0; top: -4px; height: 7px; }
.wd-frame-edge.b { left: 0; right: 0; bottom: -4px; height: 7px; }
.wd-frame-edge.l { top: 0; bottom: 0; left: -4px; width: 7px; }
.wd-frame-edge.r { top: 0; bottom: 0; right: -4px; width: 7px; }
.wd-frame-sel .wd-handle { z-index: 1; }
.wd-rotate-stem { position: absolute; left: 50%; top: -22px; width: 1px; height: 22px; background: var(--accent); }
.wd-rotate {
  position: absolute; left: calc(50% - 9px); top: -40px; width: 18px; height: 18px; border-radius: 50%;
  background: #fff; border: 1.5px solid var(--accent); color: var(--accent); display: grid; place-items: center;
  pointer-events: auto; cursor: grab; box-sizing: border-box;
}
/* Draw Text Box: the page takes a rectangle. */
.wd-page.drawing-box, .wd-page.drawing-box * { cursor: crosshair !important; }
.wd-drawbox { position: absolute; z-index: 5; border: 1px dashed var(--accent); background: rgba(43, 95, 217, 0.06); pointer-events: none; }
/* The Selection Pane, beside the page on the right. */
.wd-selpane {
  width: 250px; flex: none; border-left: 1px solid var(--line); background: var(--chrome);
  display: flex; flex-direction: column; min-height: 0; order: 3;
}
.wd-layers-tools { display: flex; align-items: center; gap: 2px; padding: 6px 8px; border-bottom: 1px solid var(--line-soft); }
.wd-layers { overflow: auto; flex: 1; min-height: 0; }
.wd-layer { display: flex; align-items: center; gap: 8px; padding: 6px 10px; cursor: default; border-bottom: 1px solid var(--line-soft); font-size: 12.5px; color: var(--ink); }
.wd-layer:hover { background: var(--surface-2); }
.wd-layer.active { background: var(--selected); box-shadow: inset 3px 0 0 var(--accent); }
.wd-layer.off .wd-layer-text, .wd-layer.off > svg { opacity: .4; }
.wd-layer.off .wd-eye { opacity: .35; }
.wd-eye { border: 0; background: none; color: var(--ink-2); cursor: pointer; padding: 2px; display: grid; border-radius: var(--r-1); }
.wd-eye:hover { background: var(--hover); color: var(--ink); }
.wd-eye-space { width: 18px; flex: none; }
.wd-layer-text { display: flex; flex-direction: column; min-width: 0; flex: 1; }
.wd-layer-title { font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.wd-layer-words { color: var(--ink-3); font-size: 11.5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.wd-layer-name { flex: 1; min-width: 0; height: 24px; }
`;
