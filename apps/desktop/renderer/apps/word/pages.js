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

/** The geometry the pass works in: sheet height, the gap, the top and bottom margins. */
export function geometryOf(section) {
  if (!section) return null;
  return {
    H: Math.round(section.heightPx || 1123),
    G: PAGE_GAP,
    top: Math.round(section.margins?.top ?? 96),
    bottom: Math.round(section.margins?.bottom ?? 96),
  };
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
  el.nodeType === 1 && (el.classList.contains('wd-block') || el.classList.contains('wd-table') || el.classList.contains('wd-notes'));

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

/** The line boxes of a paragraph, in viewport coordinates, top to bottom. */
function lineBoxes(el) {
  const range = document.createRange();
  range.selectNodeContents(el);
  const rects = [...range.getClientRects()].filter((r) => r.height > 0.5).sort((a, b) => a.top - b.top);
  const lines = [];
  for (const r of rects) {
    const last = lines[lines.length - 1];
    if (last && r.top < last.bottom - 1) {
      if (r.bottom > last.bottom) last.bottom = r.bottom;
    } else lines.push({ top: r.top, bottom: r.bottom });
  }
  if (!lines.length) {
    const r = el.getBoundingClientRect();
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
        if (p.classList.contains('wd-marker') || p.classList.contains('wd-textbox')) return NodeFilter.FILTER_REJECT;
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
      const r = document.createRange();
      r.setStart(t.node, o - t.start);
      r.setEnd(t.node, o - t.start + 1);
      return r.getBoundingClientRect().top;
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
  const elTop = it.el.getBoundingClientRect().top;
  let f = lines.findIndex((l) => placedY(l.bottom) + delta > lim + 0.5);
  if (f <= 0) return null;
  if (lines.length - f === 1) f -= 1;
  // Two lines on each side — unless the paragraph already heads the page,
  // where the alternative is overflowing it.
  if (f < 2 && !first) return null;
  if (f < 1) return null;
  const text = textNodesOf(it.el);
  const local = lineStart(text, lines[f].top);
  if (local <= 0) return null;
  // A "line" past the last character is a picture or a text box under the
  // words: the words stay, the picture goes over. Anything else at the end
  // means the measurement found nothing to move.
  if (local >= text.total && !it.el.querySelector('.wd-image, .wd-textbox')) return null;
  return { offset: it.from + local, partHeight: Math.max(0, lines[f].top - elTop) };
}

/** The row at which a table that crosses the page's bottom should break, if any. */
function splitTable(it, placedTop, lim) {
  const rows = [...(it.el.tBodies[0]?.rows || [])];
  if (rows.length < 2) return null;
  const tableTop = it.el.getBoundingClientRect().top;
  const boxes = rows.map((row) => row.getBoundingClientRect());
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
function pullBack(it, next, room) {
  if (room < 4) return null;
  if (it.kind === 'p') {
    const lines = lineBoxes(next.el);
    const top = next.el.getBoundingClientRect().top;
    let c = 0;
    while (c < lines.length && lines[c].bottom - top <= room - 1) c += 1;
    if (c === 0) return null;
    if (c === lines.length) return -1;
    if (lines.length - c < 2) return null;
    const text = textNodesOf(next.el);
    const local = lineStart(text, lines[c].top);
    if (local <= 0 || local >= text.total) return null;
    return next.from + local;
  }
  const rows = [...(next.el.tBodies[0]?.rows || [])];
  const tableTop = next.el.getBoundingClientRect().top;
  let c = 0;
  while (c < rows.length && rows[c].getBoundingClientRect().bottom - tableTop <= room - 1) c += 1;
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
  const P = geo.H + geo.G;
  const ctop = (n) => n * P + geo.top;
  const limit = (n) => n * P + geo.H - geo.bottom;
  const pageRect = page.getBoundingClientRect();
  const els = [...page.children].filter(isFlow);

  // Fractional boxes, from the rects: offsetTop and offsetHeight are whole
  // pixels, and with lines 22.5 px tall the rounding put a pushed paragraph
  // two or three pixels above the top of its page.
  const items = els.map((el) => {
    const was = applied(el);
    const kind = el.classList.contains('wd-block') ? 'p' : el.classList.contains('wd-table') ? 't' : 'n';
    const rect = el.getBoundingClientRect();
    return {
      el,
      kind,
      top: rect.top - pageRect.top,
      height: rect.height,
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
    const trySplit = (d, heads) =>
      it.kind === 'p' && !it.keepLines ? splitParagraph(it, placedY, d, limit(n), heads)
        : it.kind === 't' ? splitTable(it, natTop + d, limit(n))
          : null;

    if ((it.breakBefore || it.forceBreak) && !first) {
      n += 1;
      first = true;
      delta = ctop(n) - natTop;
    }

    let split = null;
    if (bottomAt(delta) > limit(n) + 0.5) {
      split = trySplit(delta, first);
      if (!split && !first) {
        // A heading kept with what follows goes over with it: take the
        // previous element's placement back and place it again on the next page.
        const prev = items[i - 1];
        if (prev && prev.keep && !prev.first && !prev.pulled && !prev.newSplit && prev.startOverride === null && prev.pageIndex === n) {
          shift -= prev.delta - prev.oldDelta;
          prevBottom = i >= 2 ? items[i - 2].placedBottom : geo.top;
          prev.pulled = true;
          prev.forceBreak = true;
          i -= 1;
          continue;
        }
        n += 1;
        first = true;
        delta = ctop(n) - natTop;
        if (bottomAt(delta) > limit(n) + 0.5) split = trySplit(delta, true);
      }
    }

    it.delta = delta;
    it.placedTop = natTop + delta;
    it.first = first;
    it.pageIndex = n;
    if (split) {
      it.newSplit = split;
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
      it.placedBottom = it.placedTop + it.height;
      prevBottom = it.placedBottom;
      first = false;
      const next = items[i + 1];
      if (next && continues(next, it)) {
        const back = pullBack(it, next, limit(n) - it.placedBottom);
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

  let count = n + 1;
  for (let k = 0; k < i; k++) count = Math.max(count, Math.floor(Math.max(0, items[k].placedBottom - 1) / P) + 1);

  return { changed: splits.changed || tableSplits.changed, splits: splits.map, tableSplits: tableSplits.map, count, processed: i, items };
}

/** The page (0-based) an element's top falls on, read after the pass has written. */
export function pageOfElement(el, geo) {
  if (!el) return 0;
  const page = el.closest('.wd-page');
  if (!page) return 0;
  const P = geo.H + geo.G;
  const top = el.getBoundingClientRect().top - page.getBoundingClientRect().top;
  return Math.max(0, Math.floor((top + 1) / P));
}
