/**
 * Pages on screen.
 *
 * The body is one contenteditable flow, because that is what gives the editor
 * a native caret, a native selection, an IME and the platform's spell checker.
 * But a document is not a flow, it is sheets of paper, and until now the
 * screen showed one endless sheet: a long document ran off the bottom of the
 * first page and kept going over the desk, and an explicit page break did
 * nothing anyone could see.
 *
 * This lays the flow onto pages without giving up the single surface. THE
 * BROWSER OWNS LINE BREAKING here — it is the one laying the text out, so
 * asking it where the lines fell is the only answer that agrees with what is
 * drawn. After every commit the flow's blocks are measured, and each block
 * that would cross the bottom of a page is pushed down to the top of the next
 * one by a margin; a paragraph of many lines is instead split at the line
 * that no longer fits, its remainder drawn as a second element on the next
 * page; a table splits at a row. The sheets themselves are painted behind the
 * flow, one per page, each with the header and footer it calls for.
 *
 * Nothing here moves the caret's address space: a split paragraph is still one
 * block with one index, and each part carries the character offset it starts
 * at (`data-from`), so a click on the second part maps back to the paragraph
 * the engine holds. The engine's own paginator (`doc-view/paginate.js`) still
 * lays out print and PDF; it measures with a width table and cannot see the
 * reader's fonts, which is exactly why the screen does not use it.
 *
 * The pass reads every box first and writes every margin last — one layout
 * per pass, however many pages. Interleaving the two forced a layout per
 * page, and a sixty-page report took a second per keystroke.
 */

/** The desk showing between two sheets, in CSS pixels. */
export const PAGE_GAP = 22;

/** The rule above a page's footnotes with the space round it — what `.wd-pagenotes` draws. */
export const NOTE_RULE_PX = 17;

/**
 * The geometry the pass works in: sheet height, the gap, the top and bottom
 * margins — and, when Mailings → Envelopes added an envelope in front of the
 * letter, the first sheet's own: the envelope's height, width and margins,
 * and where its sheet sits across the page (centred on the letter's).
 */
export function geometryOf(section, envelope = null) {
  if (!section) return null;
  const geo = {
    H: Math.round(section.heightPx || 1123),
    G: PAGE_GAP,
    top: Math.round(section.margins?.top ?? 96),
    bottom: Math.round(section.margins?.bottom ?? 96),
  };
  if (envelope?.heightPx) {
    const W = Math.round(envelope.widthPx);
    const left = Math.round(((section.widthPx || 794) - W) / 2);
    geo.first = {
      H: Math.round(envelope.heightPx), W, left,
      top: Math.round(envelope.margins?.top ?? 24), bottom: Math.round(envelope.margins?.bottom ?? 48),
      // How far the envelope's text sits from where the letter's would.
      dx: left + Math.round(envelope.margins?.left ?? 38) - Math.round(section.margins?.left ?? 96),
      endsAt: envelope.endsAt,
    };
  }
  return geo;
}

/** Where sheet `n` starts, down the page — every sheet the same height but an envelope first. */
export function pageTopOf(geo, n) {
  if (!geo.first) return n * (geo.H + geo.G);
  return n <= 0 ? 0 : geo.first.H + geo.G + (n - 1) * (geo.H + geo.G);
}
/** Sheet `n`'s height. */
export const pageHeightOf = (geo, n) => (geo.first && n === 0 ? geo.first.H : geo.H);
const pageMarginTop = (geo, n) => (geo.first && n === 0 ? geo.first.top : geo.top);
const pageMarginBottom = (geo, n) => (geo.first && n === 0 ? geo.first.bottom : geo.bottom);
/** The sheet (0-based) a point `y` px down the page falls on. */
export function pageIndexAt(geo, y) {
  const P = geo.H + geo.G;
  if (!geo.first) return Math.max(0, Math.floor(y / P));
  const firstP = geo.first.H + geo.G;
  return y < firstP ? 0 : 1 + Math.max(0, Math.floor((y - firstP) / P));
}

/**
 * A section's column boxes, in px from the content's left edge — the same
 * arithmetic `parseSection`'s `columnBoxes` getter does in the engine, but
 * worked from the plain fields `section` sends across the window. A getter
 * is not a value a window message necessarily keeps, so the screen works it
 * out itself from `widthPx`, `margins` and `columns`, which are.
 */
export function columnBoxesOf(section) {
  if (!section) return null;
  const contentWidthPx = section.widthPx - (section.margins?.left ?? 0) - (section.margins?.right ?? 0) - (section.margins?.gutter ?? 0);
  const cols = section.columns;
  const n = Math.max(1, cols?.count || 1);
  if (n <= 1) return [{ xPx: 0, widthPx: contentWidthPx }];
  const spacePx = cols.spacePx || 0;
  const boxes = [];
  let x = 0;
  if (cols.widths && cols.widths.length === n) {
    for (const w of cols.widths) { boxes.push({ xPx: x, widthPx: w }); x += w + spacePx; }
  } else {
    const w = Math.max(1, (contentWidthPx - spacePx * (n - 1)) / n);
    for (let i = 0; i < n; i++) { boxes.push({ xPx: x, widthPx: w }); x += w + spacePx; }
  }
  return boxes;
}

/**
 * The runs of a paragraph between two character offsets — what one part of a
 * split paragraph draws. A run cut by the boundary keeps its formatting and
 * loses the characters on the other side; an empty run stays with the part
 * its position falls in.
 */
export function sliceRuns(runs, from, to) {
  const out = [];
  let at = 0;
  for (const run of runs || []) {
    const text = run.text ?? '';
    const start = at;
    const end = at + text.length;
    at = end;
    if (text.length === 0) {
      if (start >= from && start < to) out.push(run);
      continue;
    }
    if (end <= from || start >= to) continue;
    const a = Math.max(from, start) - start;
    const b = Math.min(to, end) - start;
    out.push(a === 0 && b === text.length ? run : { ...run, text: text.slice(a, b) });
  }
  return out;
}

const isFlow = (el) =>
  el.nodeType === 1 &&
  !el.classList.contains('wd-frame') &&
  (el.classList.contains('wd-block') || el.classList.contains('wd-table') || (el.classList.contains('wd-notes') && !el.classList.contains('wd-notes-measure')));

/**
 * The push this pass gave an element last time, if it is still in force.
 * React owns the element's style; when it writes a fresh margin (the paragraph
 * was re-rendered with new spacing) ours is gone and the bookkeeping says so.
 */
function applied(el) {
  const mt = el.dataset.pushMt;
  if (mt === undefined || el.style.marginTop !== mt) return null;
  return { delta: Number(el.dataset.push) || 0, orig: el.dataset.pushOrig ?? '' };
}

/**
 * Push an element down by `delta`, as a top margin. Margins collapse, so the
 * margin written is the natural gap to the element above plus the push: the
 * collapsed result is then exactly gap + delta, wherever the gap came from.
 */
function setPush(el, delta, gap) {
  const was = applied(el);
  if (delta < 0.5) {
    if (was) el.style.marginTop = was.orig;
    if (el.dataset.pushMt !== undefined) {
      delete el.dataset.push;
      delete el.dataset.pushMt;
      delete el.dataset.pushOrig;
    }
    return;
  }
  const orig = was ? was.orig : el.style.marginTop;
  el.style.marginTop = `${Math.round((gap + delta) * 100) / 100}px`;
  el.dataset.push = String(Math.round(delta * 100) / 100);
  el.dataset.pushMt = el.style.marginTop;
  el.dataset.pushOrig = orig;
}

/** Take every push back — the view left print layout. */
export function clearPages(page) {
  for (const el of page.querySelectorAll('[data-push-mt]')) setPush(el, 0, 0);
}

/**
 * The page's zoom — View → Zoom scales the page with CSS `zoom`, and under it
 * a client rect comes back in screen pixels while the page's own layout (its
 * offsets, the geometry in `geo`, every style written back) stays in its own
 * pixels. Every rect read here is divided by it, so a pass at 150% sees the
 * same numbers as one at 100%. Read from the page itself — its drawn width
 * over its laid-out width — so nothing has to be told the level, and cached
 * for a frame because a pass reads thousands of rects.
 */
let zoomCache = { page: null, z: 1, at: -1 };
export function zoomOf(page) {
  if (!page) return 1;
  const now = typeof performance !== 'undefined' ? performance.now() : Date.now();
  if (zoomCache.page === page && now - zoomCache.at < 16) return zoomCache.z;
  const w = page.offsetWidth;
  const z = w > 0 ? page.getBoundingClientRect().width / w : 1;
  zoomCache = { page, z: Number.isFinite(z) && z > 0 ? z : 1, at: now };
  return zoomCache.z;
}
const pageOf = (node) => {
  const el = node instanceof Range ? node.startContainer : node;
  const from = el?.nodeType === 3 ? el.parentElement : el;
  return from?.closest?.('.wd-page') || null;
};
const scaled = (r, z) => (z === 1 ? r : { top: r.top / z, bottom: r.bottom / z, left: r.left / z, right: r.right / z, width: r.width / z, height: r.height / z });
/** An element's client rect in the page's own pixels, whatever the zoom. */
export function rectOf(el) {
  return scaled(el.getBoundingClientRect(), zoomOf(pageOf(el)));
}
/** A range's client rects in the page's own pixels. */
export function rectsOf(range) {
  const z = zoomOf(pageOf(range));
  return [...range.getClientRects()].map((r) => scaled(r, z));
}

/** The line boxes of a paragraph, in viewport coordinates, top to bottom. */
export function lineBoxes(el) {
  const range = document.createRange();
  range.selectNodeContents(el);
  const rects = [...rectsOf(range)].filter((r) => r.height > 0.5).sort((a, b) => a.top - b.top);
  const lines = [];
  for (const r of rects) {
    const last = lines[lines.length - 1];
    if (last && r.top < last.bottom - 1) {
      if (r.bottom > last.bottom) last.bottom = r.bottom;
    } else lines.push({ top: r.top, bottom: r.bottom });
  }
  if (!lines.length) {
    const r = rectOf(el);
    lines.push({ top: r.top, bottom: r.bottom });
  }
  return lines;
}

/**
 * The paragraph's text nodes with their character offsets — the same address
 * space the caret uses: the list marker and a text box's words are not the
 * paragraph's characters and are left out.
 */
function textNodesOf(el) {
  const nodes = [];
  let at = 0;
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
    acceptNode(n) {
      for (let p = n.parentElement; p && p !== el; p = p.parentElement) {
        if (p.classList.contains('wd-marker') || p.classList.contains('wd-textbox') || p.classList.contains('wd-shy')) return NodeFilter.FILTER_REJECT;
      }
      return NodeFilter.FILTER_ACCEPT;
    },
  });
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const len = n.nodeValue.length;
    if (!len) continue;
    nodes.push({ node: n, start: at, end: at + len });
    at += len;
  }
  return { nodes, total: at };
}

/** The top of the character at offset `o`, in viewport coordinates. */
function charTop(nodes, o) {
  let lo = 0;
  let hi = nodes.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const t = nodes[mid];
    if (o < t.start) hi = mid - 1;
    else if (o >= t.end) lo = mid + 1;
    else {
      // An equation's one character is not drawn — its MathML is, in the
      // host's shadow — so the equation's top is its host's.
      const host = t.node.parentElement?.closest?.('.wd-math');
      if (host) return rectOf(host).top;
      const r = document.createRange();
      r.setStart(t.node, o - t.start);
      r.setEnd(t.node, o - t.start + 1);
      return rectOf(r).top;
    }
  }
  return Infinity;
}

/**
 * The character offset at which the line whose top is `lineTop` begins: the
 * first character drawn at or below it. Characters go down the page in
 * offset order, so this is a binary search — a handful of measurements
 * however long the paragraph.
 */
function lineStart(text, lineTop) {
  let lo = 0;
  let hi = text.total;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (charTop(text.nodes, mid) >= lineTop - 1) hi = mid;
    else lo = mid + 1;
  }
  return lo;
}

/**
 * Where a paragraph that crosses the page's bottom should break, if it
 * should: the first line that does not fit, subject to Word's default widow
 * and orphan control — at least two lines on each side, else the whole
 * paragraph moves. `placedY` maps a viewport y to where it will be once the
 * pass's pushes above are applied; `delta` is this paragraph's own push.
 */
function splitParagraph(it, placedY, delta, lim, first) {
  const lines = lineBoxes(it.el);
  if (lines.length < 2) return null;
  const elTop = rectOf(it.el).top;
  let f = lines.findIndex((l) => placedY(l.bottom) + delta > lim + 0.5);
  if (f <= 0) return null;
  // Never leave one line alone on the next page: the split moves up a line —
  // unless the paragraph heads its page (or is taller than one) and moving it
  // up would leave nothing here, where the alternative is running over the
  // page's edge. A sheet of two picture lines that heads the page splits
  // between them.
  if (lines.length - f === 1 && (f > 1 || !first)) f -= 1;
  // Two lines on each side — unless the paragraph already heads the page,
  // where the alternative is overflowing it.
  if (f < 2 && !first) return null;
  if (f < 1) return null;
  const offset = offsetAtLine(it, lines[f].top);
  if (offset === null || offset <= it.from) return null;
  return { offset, partHeight: Math.max(0, lines[f].top - elTop) };
}

/**
 * The block offset at which the line whose top is `lineTop` begins. A line of
 * words starts at its first character. A line past the last character is a
 * picture drawn under the words: the pictures continue the paragraph's
 * address space after its text — the text's length plus the picture's index,
 * which the part carries as `data-length` and each picture as `data-image` —
 * so a split there sends that picture and the ones after it over, and the
 * ones before it stay. A text box under the words comes after every picture
 * (`data-tail`). Null when there is nothing to move.
 */
function offsetAtLine(it, lineTop) {
  const text = textNodesOf(it.el);
  const local = lineStart(text, lineTop);
  if (local < text.total) return it.from + local;
  const length = Number(it.el.dataset.length);
  if (!Number.isFinite(length)) return null;
  const pictures = [...it.el.querySelectorAll(':scope > .wd-image:not(.wd-float)')];
  const k = pictures.findIndex((img) => rectOf(img).top >= lineTop - 1);
  if (k >= 0) return length + Number(pictures[k].dataset.image);
  if (it.el.querySelector(':scope > .wd-textbox')) return Number(it.el.dataset.tail);
  return null;
}

/** The row at which a table that crosses the page's bottom should break, if any. */
function splitTable(it, placedTop, lim) {
  const rows = [...(it.el.tBodies[0]?.rows || [])];
  if (rows.length < 2) return null;
  const tableTop = rectOf(it.el).top;
  const boxes = rows.map((row) => rectOf(row));
  const r = boxes.findIndex((b) => placedTop + (b.bottom - tableTop) > lim + 0.5);
  if (r <= 0) return null;
  return { row: it.rowFrom + r, partHeight: boxes[r].top - tableTop };
}

/** Is `next` the continuation of `it` — the following part of the same paragraph or table? */
const continues = (next, it) =>
  next.kind === it.kind && (it.kind === 'p' ? next.block === it.block && next.from > it.from : it.kind === 't' && next.table === it.table && next.rowFrom > it.rowFrom);

/**
 * Room was left under a part whose continuation heads the next page: how far
 * forward its split point can move so the lines (or rows) that fit come back
 * up. -1 means all of them — the split goes away.
 */
function pullBack(it, next, room, noteCost = () => 0) {
  if (room < 4) return null;
  if (it.kind === 'p') {
    const lines = lineBoxes(next.el);
    const top = rectOf(next.el).top;
    // A line that comes back brings the footnotes it references; they need
    // their room at the foot as well.
    let c = 0;
    while (c < lines.length && lines[c].bottom - top + noteCost(lines[c].bottom) <= room - 1) c += 1;
    if (c === 0) return null;
    if (c === lines.length) return -1;
    if (lines.length - c < 2) return null;
    const offset = offsetAtLine(next, lines[c].top);
    if (offset === null || offset <= next.from) return null;
    return offset;
  }
  const rows = [...(next.el.tBodies[0]?.rows || [])];
  const tableTop = rectOf(next.el).top;
  let c = 0;
  while (c < rows.length && rectOf(rows[c]).bottom - tableTop <= room - 1) c += 1;
  if (c === 0) return null;
  if (c === rows.length) return -1;
  return next.rowFrom + c;
}

/** Keep the old array where nothing changed, so memoised blocks stay put. */
function settle(fresh, old, seen) {
  const map = {};
  let changed = false;
  for (const k of Object.keys(fresh)) {
    const v = [...new Set(fresh[k])].sort((a, b) => a - b);
    const o = old[k];
    if (o && o.length === v.length && o.every((x, j) => x === v[j])) map[k] = o;
    else {
      map[k] = v;
      changed = true;
    }
  }
  for (const k of Object.keys(old)) {
    if (k in map) continue;
    // A block the pass never saw is not mounted yet; its splits stand.
    if (!seen.has(k)) map[k] = old[k];
    else changed = true;
  }
  return { map, changed };
}

/**
 * Lay the flow onto pages: measure every top-level block, push what crosses a
 * page's bottom to the next page, split what can be split, and say how many
 * pages that made. `state` is the split points in force; the answer carries
 * the new ones and whether they differ. Pure reads first, then the writes.
 */
export function layPages(page, geo, state) {
  const ctop = (n) => pageTopOf(geo, n) + pageMarginTop(geo, n);
  const pageRect = rectOf(page);
  const els = [...page.children].filter(isFlow);

  // Footnotes go at the foot of the page their reference lands on — Word's
  // rule, and the print paginator's. Their heights come from the hidden copy
  // in the flow; a page reserves that much at its foot, plus the rule above
  // the first, and its bottom limit moves up by it. The answer says which
  // page each note is drawn on.
  const noteHeight = new Map();
  for (const el of page.querySelectorAll('.wd-notes-measure .wd-note[data-note]')) noteHeight.set(el.dataset.note, rectOf(el).height);
  const reserved = [];
  const reservedAt = (k) => reserved[k] || 0;
  const limit = (n) => pageTopOf(geo, n) + pageHeightOf(geo, n) - pageMarginBottom(geo, n) - reservedAt(n);
  const notePages = {};
  const refsOf = (el) =>
    noteHeight.size
      ? [...el.querySelectorAll('.wd-noteref[data-kind="footnote"][data-id]')]
          .filter((r) => noteHeight.has(r.dataset.id))
          .map((r) => ({ id: r.dataset.id, height: noteHeight.get(r.dataset.id), top: rectOf(r).top }))
      : [];
  const costOf = (refs, k) => (refs.length ? refs.reduce((s, r) => s + r.height, 0) + (reservedAt(k) ? 0 : NOTE_RULE_PX) : 0);
  const reserve = (it, costRefs, k, placedRefs) => {
    it.noteCost = costOf(costRefs, k);
    it.notePage = k;
    it.noteIds = placedRefs.map((r) => r.id);
    reserved[k] = reservedAt(k) + it.noteCost;
    for (const r of placedRefs) notePages[r.id] = k;
  };
  const unreserve = (it) => {
    if (!it.noteCost && !it.noteIds?.length) return;
    reserved[it.notePage] = Math.max(0, reservedAt(it.notePage) - it.noteCost);
    for (const id of it.noteIds) delete notePages[id];
    it.noteCost = 0;
    it.noteIds = [];
  };

  // Fractional boxes, from the rects: offsetTop and offsetHeight are whole
  // pixels, and with lines 22.5 px tall the rounding put a pushed paragraph
  // two or three pixels above the top of its page.
  const items = els.map((el) => {
    const was = applied(el);
    const kind = el.classList.contains('wd-block') ? 'p' : el.classList.contains('wd-table') ? 't' : 'n';
    const rect = rectOf(el);
    // A float beside the words can hang below its paragraph's last line;
    // the block is as tall as the float, or the next page cuts the picture.
    let bottom = rect.bottom;
    for (const f of el.querySelectorAll('.wd-float')) bottom = Math.max(bottom, rectOf(f).bottom);
    return {
      el,
      kind,
      top: rect.top - pageRect.top,
      height: bottom - rect.top,
      oldDelta: was ? was.delta : 0,
      block: kind === 'p' ? Number(el.dataset.block) : null,
      table: kind === 't' ? el.dataset.table || null : null,
      from: Number(el.dataset.from || 0),
      rowFrom: Number(el.dataset.rowFrom || 0),
      keep: el.dataset.keep === '1',
      keepLines: el.dataset.keeplines === '1',
      breakBefore: el.dataset.break === '1',
      delta: 0, gap: 0, placedTop: 0, placedBottom: 0, first: false, pageIndex: 0,
      newSplit: null, startOverride: null, pulled: false, forceBreak: false,
      noteCost: 0, notePage: 0, noteIds: [],
    };
  });

  let n = 0;
  let shift = 0;
  let prevBottom = geo.top;
  let first = true;
  let i = 0;
  let stop = false;

  while (i < items.length && !stop) {
    const it = items[i];
    const natTop = it.top - it.oldDelta + shift;
    it.gap = Math.max(0, natTop - prevBottom);
    let delta = 0;
    const placedY = (y) => y - pageRect.top - it.oldDelta + shift;
    const bottomAt = (d) => natTop + d + it.height;
    // A paragraph taller than a page's inside has to split somewhere: the
    // widow and orphan rules that would move it whole are waived for it.
    const tall = it.kind === 'p' && it.height > geo.H - geo.top - geo.bottom;
    // The footnotes this part references: they take room at the foot of the
    // page the part lands on, so the part has to fit above that room.
    const refs = it.kind === 'p' ? refsOf(it.el) : [];
    const trySplit = (d, heads) => {
      if (it.kind === 't') return splitTable(it, natTop + d, limit(n));
      if (it.kind !== 'p' || it.keepLines) return null;
      // Only the notes referenced on the lines that stay need room here,
      // and which lines stay depends on that room: a few rounds settle it.
      // If they do not, every note is given room and the ones above the
      // cut are drawn here — space wasted rather than words overlapped.
      const attempt = (keep) => {
        const split = splitParagraph(it, placedY, d, limit(n) - costOf(keep, n), heads || tall);
        if (!split) return null;
        const cut = rectOf(it.el).top + split.partHeight;
        return { split, above: refs.filter((r) => r.top < cut - 0.5) };
      };
      let keep = refs;
      for (let round = 0; round < 4; round++) {
        const got = attempt(keep);
        if (!got) return null;
        if (got.above.length === keep.length && got.above.every((r, j) => r === keep[j])) return { ...got.split, notes: got.above, reserveFor: got.above };
        keep = got.above;
      }
      const got = attempt(refs);
      return got ? { ...got.split, notes: got.above, reserveFor: refs } : null;
    };

    if ((it.breakBefore || it.forceBreak) && !first) {
      n += 1;
      first = true;
      delta = ctop(n) - natTop;
    }

    let split = null;
    if (bottomAt(delta) > limit(n) - costOf(refs, n) + 0.5) {
      split = trySplit(delta, first);
      if (!split && !first) {
        // A heading kept with what follows goes over with it: take the
        // previous element's placement back and place it again on the next page.
        const prev = items[i - 1];
        if (prev && prev.keep && !prev.first && !prev.pulled && !prev.newSplit && prev.startOverride === null && prev.pageIndex === n) {
          shift -= prev.delta - prev.oldDelta;
          prevBottom = i >= 2 ? items[i - 2].placedBottom : geo.top;
          unreserve(prev);
          prev.pulled = true;
          prev.forceBreak = true;
          i -= 1;
          continue;
        }
        n += 1;
        first = true;
        delta = ctop(n) - natTop;
        if (bottomAt(delta) > limit(n) - costOf(refs, n) + 0.5) split = trySplit(delta, true);
      }
    }

    it.delta = delta;
    it.placedTop = natTop + delta;
    it.first = first;
    it.pageIndex = n;
    unreserve(it);
    if (split) {
      it.newSplit = split;
      reserve(it, split.reserveFor || [], n, split.notes || []);
      it.placedBottom = it.placedTop + split.partHeight;
      const remainder = Math.max(0, it.height - split.partHeight);
      n += 1;
      // The rest of this element will be drawn at the head of the next page.
      // Until it is, everything below is measured against the unsplit
      // element; shift it to where it will land, and let the next pass
      // measure the truth.
      const remainderTop = ctop(n);
      shift += remainderTop + remainder - (it.placedTop + it.height);
      prevBottom = remainderTop + remainder;
      first = false;
    } else {
      reserve(it, refs, n, refs);
      it.placedBottom = it.placedTop + it.height;
      prevBottom = it.placedBottom;
      first = false;
      const next = items[i + 1];
      if (next && continues(next, it)) {
        const back = pullBack(it, next, limit(n) - it.placedBottom, (cutBottom) => costOf(refsOf(next.el).filter((r) => r.top < cutBottom - 0.5), n));
        if (back !== null) {
          next.startOverride = back;
          // The parts below move once the lines do; nothing measured past
          // here is worth writing. The next pass sees the new shape.
          stop = true;
        }
      }
    }
    shift += delta - it.oldDelta;
    i += 1;
  }

  for (let k = 0; k < i; k++) setPush(items[k].el, items[k].delta, items[k].gap);

  // A note the placement never reached — its reference in a table, in the
  // remainder of a split not yet drawn, or past where the pass stopped —
  // goes on the page its reference is drawn on now; the next pass sees the
  // truth. Read after the pushes are written, so it costs a layout only then.
  for (const id of noteHeight.keys()) {
    if (id in notePages) continue;
    const ref = page.querySelector(`.wd-noteref[data-kind="footnote"][data-id="${CSS.escape(id)}"]`);
    if (!ref) continue;
    notePages[id] = pageIndexAt(geo, rectOf(ref).top - pageRect.top + 1);
  }
  const oldNotes = state.notes || {};
  const notesChanged = Object.keys(notePages).length !== Object.keys(oldNotes).length || Object.keys(notePages).some((id) => oldNotes[id] !== notePages[id]);

  const fresh = { p: {}, t: {} };
  const seen = { p: new Set(), t: new Set() };
  const add = (kind, key, v) => (fresh[kind][key] ||= []).push(v);
  for (const it of items) {
    if (it.kind === 'p') {
      seen.p.add(String(it.block));
      if (it.from > 0) {
        const start = it.startOverride ?? it.from;
        if (start !== -1) add('p', it.block, start);
      }
      if (it.newSplit) add('p', it.block, it.newSplit.offset);
    } else if (it.kind === 't' && it.table) {
      seen.t.add(it.table);
      if (it.rowFrom > 0) {
        const start = it.startOverride ?? it.rowFrom;
        if (start !== -1) add('t', it.table, start);
      }
      if (it.newSplit) add('t', it.table, it.newSplit.row);
    }
  }
  const splits = settle(fresh.p, state.splits || {}, seen.p);
  const tableSplits = settle(fresh.t, state.tableSplits || {}, seen.t);

  // Stopping early (`pullBack` above) leaves everything past `i` unmeasured
  // and unpushed this pass — still exactly where the last pass drew it. The
  // count has to cover those too, or a pass that stops halfway through an
  // eleven-page document reports six: real sheets vanishing from under real
  // words for as long as the next pass takes to come round. An unprocessed
  // element's own top and height, read before this pass touched anything,
  // is where it still sits, so that stands in for its placed bottom.
  let count = n + 1;
  for (let k = 0; k < items.length; k++) {
    const bottom = k < i ? items[k].placedBottom : items[k].top + items[k].height;
    count = Math.max(count, pageIndexAt(geo, Math.max(0, bottom - 1)) + 1);
  }

  // A paragraph in a frame placed on the page is drawn out of the flow, on
  // the sheet of the words before it — an envelope's delivery address on
  // the envelope. Which sheet that is comes out of this pass.
  const frames = {};
  if (page.querySelector(':scope > .wd-frame')) {
    const byEl = new Map(items.map((x) => [x.el, x]));
    let lastPage = 0;
    for (const el of page.children) {
      if (el.classList?.contains('wd-frame')) { frames[el.dataset.block] = lastPage; continue; }
      const it = byEl.get(el);
      if (it) lastPage = it.pageIndex;
    }
  }
  const oldFrames = state.frames || {};
  const framesChanged = Object.keys(frames).length !== Object.keys(oldFrames).length || Object.keys(frames).some((k) => oldFrames[k] !== frames[k]);

  return { changed: splits.changed || tableSplits.changed || notesChanged || framesChanged, splits: splits.map, tableSplits: tableSplits.map, notes: notesChanged ? notePages : oldNotes, frames: framesChanged ? frames : oldFrames, count, processed: i, items };
}

/** The page (0-based) an element's top falls on, read after the pass has written. */
export function pageOfElement(el, geo) {
  if (!el) return 0;
  const page = el.closest('.wd-page');
  if (!page) return 0;
  const top = rectOf(el).top - rectOf(page).top;
  return pageIndexAt(geo, top + 1);
}
