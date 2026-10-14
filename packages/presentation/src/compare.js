// Review → Compare in Presentations, as PowerPoint does it: a reviewed copy
// of the deck merged into the open one as changes, per slide and per
// shape, the words put in and taken out, a shape added, deleted, moved,
// resized, restyled or given another picture, a slide added, deleted or
// moved, each shown on the slide and accepted (made as the reviewer has it)
// or rejected (left, or put back, as the deck had it), one at a time, a
// slide at a time or all at once.
//
// Three decks take part: the open one, a copy of it as it was when the
// review began (`base`), and the reviewer's (`reviewed`). Accepting copies an
// aspect of a shape or a slide from the reviewer's deck into the open one;
// rejecting copies it back from the base. Whether a change is accepted is
// read from the open deck itself, its aspect as the reviewer's or as the
// base's, so Undo and Redo, which put the open deck back, leave the list
// true without a word to it.

import { OoxmlPackage } from '@rutba/ooxml/package';

const unesc = (s) => String(s ?? '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const fnv = (text) => {
  let h = 0x811c9dc5;
  const s = String(text);
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(36);
};

/* ── shapes in a slide's XML ─────────────────────────────────────────────── */

const SHAPE_TAGS = ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp'];

/** The top-level shapes of a slide's tree, each with its span, id, name and the aspects compared. */
export function shapesOf(xml) {
  const tree = /<p:spTree\b[^>]*>/.exec(xml);
  if (!tree) return [];
  const out = [];
  let at = tree.index + tree[0].length;
  const end = xml.lastIndexOf('</p:spTree>');
  const opener = /<(p:sp|p:pic|p:graphicFrame|p:grpSp|p:cxnSp|mc:AlternateContent|p:contentPart)\b[^>]*?(\/?)>/g;
  while (at < end) {
    opener.lastIndex = at;
    const m = opener.exec(xml);
    if (!m || m.index >= end) break;
    const name = m[1];
    let stop;
    if (m[2]) stop = m.index + m[0].length;
    else {
      const re = new RegExp('<' + name + '\\b[^>]*?(/?)>|</' + name + '>', 'g');
      re.lastIndex = m.index;
      let depth = 0;
      let t;
      stop = -1;
      while ((t = re.exec(xml))) {
        if (t[0].startsWith('</')) { depth -= 1; if (depth === 0) { stop = t.index + t[0].length; break; } } else if (!t[1]) depth += 1;
      }
      if (stop < 0) break;
    }
    const shapeXml = xml.slice(m.index, stop);
    const nv = /<p:cNvPr\b([^>]*)/.exec(shapeXml)?.[1] ?? '';
    out.push({
      tag: name, start: m.index, end: stop, xml: shapeXml,
      id: /\bid="(\d+)"/.exec(nv)?.[1] ?? null,
      name: unesc(/\bname="([^"]*)"/.exec(nv)?.[1] ?? ''),
    });
    at = stop;
  }
  return out;
}

/** A shape's box element: `a:xfrm` in its properties, or a graphic frame's `p:xfrm`. */
function xfrmOf(shapeXml) {
  if (/^<p:graphicFrame\b/.test(shapeXml)) return /<p:xfrm\b[^>]*>[\s\S]*?<\/p:xfrm>/.exec(shapeXml)?.[0] ?? null;
  const pr = /<p:(?:spPr|grpSpPr)\b[^>]*>([\s\S]*?)<\/p:(?:spPr|grpSpPr)>/.exec(shapeXml);
  return pr ? /<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/.exec(pr[1])?.[0] ?? null : null;
}

/** Where a box puts a shape, in EMU. */
function boxOf(xfrm) {
  if (!xfrm) return null;
  const off = /<a:off\b[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/.exec(xfrm);
  const ext = /<a:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(xfrm);
  return { x: Number(off?.[1] ?? 0), y: Number(off?.[2] ?? 0), w: Number(ext?.[1] ?? 0), h: Number(ext?.[2] ?? 0), rot: Number(/\brot="(-?\d+)"/.exec(xfrm)?.[1] ?? 0) };
}

const txBodyOf = (shapeXml) => /<p:txBody\b[^>]*>[\s\S]*?<\/p:txBody>/.exec(shapeXml)?.[0] ?? null;
/** The shape's own look: its properties without the box. */
const styleOf = (shapeXml) => {
  const pr = /<p:(spPr|grpSpPr)\b[^>]*>[\s\S]*?<\/p:\1>|<p:(spPr|grpSpPr)\b[^>]*\/>/.exec(shapeXml)?.[0] ?? '';
  return pr.replace(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/, '') + (/<p:style\b[^>]*>[\s\S]*?<\/p:style>/.exec(shapeXml)?.[0] ?? '');
};
const graphicOf = (shapeXml) => /<a:graphic\b[^>]*>[\s\S]*?<\/a:graphic>/.exec(shapeXml)?.[0] ?? null;
const blipOf = (shapeXml) => /<a:blip\b[^>]*r:embed="([^"]*)"/.exec(shapeXml)?.[1] ?? null;

/**
 * Where a shape stands on its slide as the slide draws it, in pixels, rounded:
 * a placeholder that states no box of its own stands where its layout puts it.
 */
function placed(deck, slideIndex, id) {
  const g = deck.slide(slideIndex).shapes.find((s) => String(s.id) === String(id))?.geometry;
  return g ? { x: Math.round(g.x), y: Math.round(g.y), w: Math.round(g.w), h: Math.round(g.h), rot: Math.round(g.rot || 0) } : null;
}

/** A text body's paragraphs, as words. */
export function wordsOf(txBody) {
  if (!txBody) return [];
  return [...txBody.matchAll(/<a:p\b[^>]*>([\s\S]*?)<\/a:p>|<a:p\b[^>]*\/>/g)].map((p) => [...(p[1] || '').matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>|<a:br\b[^>]*\/>/g)].map((t) => (t[0].startsWith('<a:br') ? '\n' : unesc(t[1]))).join(''));
}

/** What was put in and taken out between two texts, word by word: `{ inserted: [..], deleted: [..] }`. */
export function wordDiff(before, after) {
  const a = String(before).split(/(\s+)/).filter((w) => w !== '');
  const b = String(after).split(/(\s+)/).filter((w) => w !== '');
  const n = a.length;
  const m = b.length;
  const lcs = Array.from({ length: n + 1 }, () => new Int32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) lcs[i][j] = a[i] === b[j] ? lcs[i + 1][j + 1] + 1 : Math.max(lcs[i + 1][j], lcs[i][j + 1]);
  const inserted = [];
  const deleted = [];
  let i = 0;
  let j = 0;
  let ins = '';
  let del = '';
  const flush = () => {
    if (ins.trim()) inserted.push(ins.trim());
    if (del.trim()) deleted.push(del.trim());
    ins = '';
    del = '';
  };
  while (i < n || j < m) {
    if (i < n && j < m && a[i] === b[j]) { flush(); i += 1; j += 1; } else if (j < m && (i >= n || lcs[i][j + 1] >= lcs[i + 1][j])) { ins += b[j]; j += 1; } else { del += a[i]; i += 1; }
  }
  flush();
  return { inserted, deleted };
}

/* ── the pictures and parts a shape brings with it ───────────────────────── */

/** The bytes a relationship of a slide points at in a deck, or null. */
function relTarget(deck, slidePart, rId) {
  const rel = deck.pkg.rels(slidePart).find((r) => r.Id === rId);
  if (!rel || rel.TargetMode === 'External') return rel ? { rel, external: true } : null;
  const part = OoxmlPackage.resolveTarget(slidePart, rel.Target);
  return deck.pkg.has(part) ? { rel, part, bytes: deck.pkg.read(part) } : null;
}

/**
 * A shape's XML from another deck made at home in `deck`'s slide: each
 * relationship it uses pointed at a part of this deck, the same picture
 * when this deck already has those bytes, else a copy added, and each
 * external link added again.
 */
function bringShape(fromDeck, fromPart, shapeXml, deck, part) {
  let out = shapeXml;
  for (const id of new Set([...shapeXml.matchAll(/\br:(?:embed|link|id|pict)="([^"]+)"/g)].map((m) => m[1]))) {
    const t = relTarget(fromDeck, fromPart, id);
    if (!t) continue;
    let rId;
    if (t.external) rId = deck.pkg.addRelationshipTo(part, t.rel.Type, t.rel.Target, { external: true });
    else {
      const same = deck.pkg.partNames().find((n) => n.startsWith('ppt/media/') && deck.pkg.read(n).equals(t.bytes));
      let target = same;
      if (!target) {
        const ext = t.part.split('.').pop();
        let k = 1;
        while (deck.pkg.has(`ppt/media/review${k}.${ext}`)) k += 1;
        target = `ppt/media/review${k}.${ext}`;
        const type = fromDeck.pkg.contentTypeOf(t.part);
        if (type && !deck.pkg.contentTypeOf(target)) deck.pkg.ensureDefault(ext, type);
        deck.pkg.addPart(target, t.bytes);
      }
      const rel = '../' + target.slice('ppt/'.length);
      rId = deck.pkg.rels(part).find((r) => r.Type === t.rel.Type && OoxmlPackage.resolveTarget(part, r.Target) === target)?.Id || deck.pkg.addRelationshipTo(part, t.rel.Type, rel);
    }
    out = out.replace(new RegExp('(\\br:(?:embed|link|id|pict)=")' + id + '"', 'g'), (whole, a) => a + rId + '"');
  }
  return out;
}

/* ── comparing ───────────────────────────────────────────────────────────── */

/** Each slide's id and part, in order. */
const slidesOf = (deck) => deck.slideParts.map((s, index) => ({ id: String(s.id), part: s.part, rId: s.rId, index }));

/** A slide's title, for the pane. */
function titleOf(xml) {
  const title = shapesOf(xml).find((s) => /<p:ph\b[^>]*type="(?:title|ctrTitle)"/.test(s.xml));
  return title ? wordsOf(txBodyOf(title.xml)).join(' ').trim().slice(0, 60) : '';
}

/** Slides matched between two decks: by their p:sldId, which a copy of a deck keeps; then, for the rest, by their title. */
function matchSlides(base, reviewed) {
  const a = slidesOf(base);
  const b = slidesOf(reviewed);
  const pairs = new Map();
  const taken = new Set();
  for (const s of a) {
    const hit = b.find((r) => r.id === s.id && !taken.has(r.id));
    if (hit) { pairs.set(s.id, hit); taken.add(hit.id); }
  }
  for (const s of a.filter((x) => !pairs.has(x.id))) {
    const title = titleOf(base.pkg.text(s.part));
    const hit = title && b.find((r) => !taken.has(r.id) && titleOf(reviewed.pkg.text(r.part)) === title);
    if (hit) { pairs.set(s.id, hit); taken.add(hit.id); }
  }
  return { a, b, pairs, added: b.filter((r) => !taken.has(r.id)) };
}

/** The slides that moved: those outside the longest run kept in order. */
function movedSlides(a, b, pairs) {
  const order = a.filter((s) => pairs.has(s.id)).map((s) => ({ id: s.id, at: b.indexOf(pairs.get(s.id)) }));
  const seq = order.map((o) => o.at);
  const tails = [];
  const prev = new Array(seq.length).fill(-1);
  seq.forEach((v, i) => {
    let lo = 0;
    let hi = tails.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (seq[tails[mid]] < v) lo = mid + 1; else hi = mid; }
    prev[i] = lo > 0 ? tails[lo - 1] : -1;
    tails[lo] = i;
  });
  const keep = new Set();
  for (let k = tails.length ? tails[tails.length - 1] : -1; k >= 0; k = prev[k]) keep.add(order[k].id);
  return order.filter((o) => !keep.has(o.id)).map((o) => o.id);
}

/**
 * The changes a reviewer made: `base` is the deck as it was, `reviewed` the
 * reviewer's copy. Each change: `{ id, scope, kind, slideId, shapeId?,
 * name, detail }`, kind one of slideAdded, slideDeleted, slideMoved,
 * shapeAdded, shapeDeleted, text, moved, resized, restyled, picture.
 */
export function compareDecks(base, reviewed) {
  const { a, b, pairs, added } = matchSlides(base, reviewed);
  const changes = [];
  let n = 0;
  const push = (c) => changes.push({ id: 'c' + (++n), ...c });
  for (const s of added) push({ scope: 'slide', kind: 'slideAdded', slideId: s.id, reviewedSlideId: s.id, name: titleOf(reviewed.pkg.text(s.part)), detail: { at: s.index } });
  for (const s of a.filter((x) => !pairs.has(x.id))) push({ scope: 'slide', kind: 'slideDeleted', slideId: s.id, name: titleOf(base.pkg.text(s.part)), detail: { at: s.index } });
  for (const id of movedSlides(a, b, pairs)) {
    const s = a.find((x) => x.id === id);
    push({ scope: 'slide', kind: 'slideMoved', slideId: id, reviewedSlideId: pairs.get(id).id, name: titleOf(base.pkg.text(s.part)), detail: { from: s.index, to: pairs.get(id).index } });
  }
  for (const s of a) {
    const r = pairs.get(s.id);
    if (!r) continue;
    const mine = shapesOf(base.pkg.text(s.part));
    const theirs = shapesOf(reviewed.pkg.text(r.part));
    const used = new Set();
    const counterpart = (sh) => {
      const byId = theirs.find((t) => !used.has(t) && t.id === sh.id && t.tag === sh.tag);
      const byName = byId || (sh.name ? theirs.find((t) => !used.has(t) && t.name === sh.name && t.tag === sh.tag) : null);
      if (byName) used.add(byName);
      return byName;
    };
    for (const sh of mine) {
      const t = counterpart(sh);
      const base1 = { scope: 'shape', slideId: s.id, reviewedSlideId: r.id, shapeId: sh.id, reviewedShapeId: t?.id ?? null, name: sh.name };
      if (!t) { push({ ...base1, kind: 'shapeDeleted', detail: { box: boxOf(xfrmOf(sh.xml)) } }); continue; }
      const tb = txBodyOf(sh.xml);
      const tt = txBodyOf(t.xml);
      if ((tb || tt) && tb !== tt) {
        const before = wordsOf(tb).join('\n');
        const after = wordsOf(tt).join('\n');
        const diff = wordDiff(before, after);
        push({ ...base1, kind: 'text', detail: { ...diff, formatted: before === after } });
      }
      const xa = placed(base, s.index, sh.id);
      const xb = placed(reviewed, r.index, t.id);
      if (xa && xb) {
        if (xa.x !== xb.x || xa.y !== xb.y || xa.rot !== xb.rot) push({ ...base1, kind: 'moved', detail: { from: xa, to: xb } });
        if (xa.w !== xb.w || xa.h !== xb.h) push({ ...base1, kind: 'resized', detail: { from: xa, to: xb } });
      }
      if (styleOf(sh.xml) !== styleOf(t.xml)) push({ ...base1, kind: 'restyled', detail: {} });
      if (sh.tag === 'p:graphicFrame' && graphicOf(sh.xml) !== graphicOf(t.xml) && !/r:id=/.test(graphicOf(sh.xml) || '')) push({ ...base1, kind: 'text', detail: { table: true, inserted: [], deleted: [] } });
      const ba = blipOf(sh.xml);
      const bb = blipOf(t.xml);
      if (ba && bb) {
        const pa = relTarget(base, s.part, ba)?.bytes;
        const pb = relTarget(reviewed, r.part, bb)?.bytes;
        if (pa && pb && !pa.equals(pb)) push({ ...base1, kind: 'picture', detail: {} });
      }
    }
    for (const t of theirs.filter((x) => !used.has(x))) {
      push({ scope: 'shape', kind: 'shapeAdded', slideId: s.id, reviewedSlideId: r.id, shapeId: null, reviewedShapeId: t.id, name: t.name, detail: { box: boxOf(xfrmOf(t.xml)), words: wordsOf(txBodyOf(t.xml)).join(' ').slice(0, 60) } });
    }
  }
  return changes;
}

/* ── the open deck, as each change finds it ──────────────────────────────── */

/** Where a slide is in a deck now, by its id; -1 when it is not there. */
const slideIndexOf = (deck, id) => deck.slideParts.findIndex((s) => String(s.id) === String(id));

/** A shape on a slide of a deck, by id. */
function shapeIn(deck, slideId, shapeId) {
  const at = slideIndexOf(deck, slideId);
  if (at < 0 || shapeId == null) return null;
  const part = deck.slideParts[at].part;
  const xml = deck.pkg.text(part);
  const sh = shapesOf(xml).find((x) => x.id === String(shapeId));
  return sh ? { ...sh, part, slideXml: xml, index: at } : null;
}

/** The aspect a change is about, read off a shape. */
function aspectOf(kind, sh, deck) {
  if (!sh) return null;
  switch (kind) {
    case 'text': return (txBodyOf(sh.xml) ?? '') + (sh.tag === 'p:graphicFrame' ? graphicOf(sh.xml) ?? '' : '');
    case 'moved': { const b = placed(deck, sh.index, sh.id); return b ? [b.x, b.y, b.rot].join(',') : null; }
    case 'resized': { const b = placed(deck, sh.index, sh.id); return b ? [b.w, b.h].join(',') : null; }
    case 'restyled': return styleOf(sh.xml);
    case 'picture': { const id = blipOf(sh.xml); const t = id ? relTarget(deck, sh.part, id) : null; return t?.bytes ? fnv(t.bytes.toString('latin1')) : null; }
    default: return null;
  }
}

/**
 * Whether each change is accepted now: the open deck's aspect is the
 * reviewer's. Answers the changes with `accepted` set.
 */
export function reviewState(deck, review) {
  return review.changes.map((c) => {
    let accepted = false;
    if (c.kind === 'slideAdded') accepted = slideIndexOf(deck, c.addedAs ?? '\u0000') >= 0;
    else if (c.kind === 'slideDeleted') accepted = slideIndexOf(deck, c.slideId) < 0;
    else if (c.kind === 'slideMoved') {
      const order = deck.slideParts.map((s) => String(s.id));
      accepted = order.indexOf(c.slideId) === Math.min(c.detail.to, order.length - 1);
    } else if (c.kind === 'shapeAdded') accepted = Boolean(c.addedAs && shapeIn(deck, c.slideId, c.addedAs));
    else if (c.kind === 'shapeDeleted') accepted = !shapeIn(deck, c.slideId, c.shapeId);
    else {
      const now = aspectOf(c.kind, shapeIn(deck, c.slideId, c.shapeId), deck);
      const theirs = aspectOf(c.kind, shapeIn(review.reviewed, c.reviewedSlideId, c.reviewedShapeId), review.reviewed);
      const was = aspectOf(c.kind, shapeIn(review.base, c.slideId, c.shapeId), review.base);
      accepted = now != null && now === theirs && now !== was;
    }
    const slideAt = c.kind === 'slideAdded' ? slideIndexOf(deck, c.addedAs ?? '\u0000') : slideIndexOf(deck, c.slideId);
    return { ...c, accepted, slide: slideAt };
  });
}

/* ── accepting and rejecting ─────────────────────────────────────────────── */

/** A shape's XML with one aspect taken from another shape. */
function withAspect(kind, mine, from) {
  switch (kind) {
    case 'text': {
      let out = mine;
      const tb = txBodyOf(from);
      if (txBodyOf(mine) != null && tb != null) out = out.replace(txBodyOf(mine), () => tb);
      else if (tb != null) out = out.replace(/<\/p:sp>$/, () => tb + '</p:sp>');
      if (/^<p:graphicFrame\b/.test(mine) && graphicOf(from)) out = out.replace(graphicOf(mine), () => graphicOf(from));
      return out;
    }
    case 'moved':
    case 'resized': {
      const a = xfrmOf(mine);
      const b = xfrmOf(from);
      if (!a || !b) return mine;
      const ba = boxOf(a);
      const bb = boxOf(b);
      const box = kind === 'moved' ? { ...ba, x: bb.x, y: bb.y, rot: bb.rot } : { ...ba, w: bb.w, h: bb.h };
      const next = a.replace(/<a:off\b[^>]*\/>/, () => `<a:off x="${box.x}" y="${box.y}"/>`).replace(/<a:ext\b[^>]*\/>/, () => `<a:ext cx="${box.w}" cy="${box.h}"/>`)
        .replace(/^(<(?:a|p):xfrm\b)([^>]*)>/, (whole, tag, attrs) => tag + (box.rot ? attrs.replace(/\s+rot="-?\d+"/, '') + ` rot="${box.rot}"` : attrs.replace(/\s+rot="-?\d+"/, '')) + '>');
      return mine.replace(a, () => next);
    }
    case 'restyled': {
      const prMine = /<p:(spPr|grpSpPr)\b[^>]*>[\s\S]*?<\/p:\1>|<p:(spPr|grpSpPr)\b[^>]*\/>/.exec(mine)?.[0];
      const prFrom = /<p:(spPr|grpSpPr)\b[^>]*>[\s\S]*?<\/p:\1>|<p:(spPr|grpSpPr)\b[^>]*\/>/.exec(from)?.[0];
      if (!prMine || !prFrom) return mine;
      const keepBox = xfrmOf(mine);
      let pr = prFrom.replace(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/, '');
      if (keepBox && /^<p:(?:spPr|grpSpPr)\b[^>]*>/.test(pr)) pr = pr.replace(/^(<p:(?:spPr|grpSpPr)\b[^>]*>)/, (o) => o + keepBox);
      let out = mine.replace(prMine, () => pr);
      const styleMine = /<p:style\b[^>]*>[\s\S]*?<\/p:style>/.exec(out)?.[0];
      const styleFrom = /<p:style\b[^>]*>[\s\S]*?<\/p:style>/.exec(from)?.[0];
      if (styleMine && styleFrom) out = out.replace(styleMine, () => styleFrom);
      else if (styleMine) out = out.replace(styleMine, '');
      else if (styleFrom) out = out.replace(/(<\/p:spPr>)/, (m) => m + styleFrom);
      return out;
    }
    default: return mine;
  }
}

/** One shape on the open deck's slide replaced. */
function writeShape(deck, at, sh, xml) {
  const slideXml = deck.pkg.text(sh.part);
  deck.replaceSlideXml(sh.part, slideXml.slice(0, sh.start) + xml + slideXml.slice(sh.end));
}

/** Put a change as the reviewer has it (`accept`) or as the deck had it (`reject`). */
function setChange(deck, review, c, accept) {
  const source = accept ? review.reviewed : review.base;
  switch (c.kind) {
    case 'slideAdded': {
      const here = c.addedAs != null ? slideIndexOf(deck, c.addedAs) : -1;
      if (accept && here < 0) c.addedAs = importSlide(review.reviewed, slideIndexOf(review.reviewed, c.reviewedSlideId), deck, Math.min(c.detail.at, deck.slideCount));
      else if (!accept && here >= 0) deck.removeSlide(here);
      return;
    }
    case 'slideDeleted': {
      const here = slideIndexOf(deck, c.slideId);
      if (accept && here >= 0) deck.removeSlide(here);
      else if (!accept && here < 0) restoreSlide(deck, review.base, c.slideId);
      return;
    }
    case 'slideMoved': {
      const here = slideIndexOf(deck, c.slideId);
      if (here < 0) return;
      const to = Math.min(accept ? c.detail.to : c.detail.from, deck.slideCount - 1);
      if (to !== here) deck.moveSlide(here, to);
      return;
    }
    case 'shapeAdded': {
      const here = c.addedAs ? shapeIn(deck, c.slideId, c.addedAs) : null;
      if (accept && !here) {
        const at = slideIndexOf(deck, c.slideId);
        const from = shapeIn(review.reviewed, c.reviewedSlideId, c.reviewedShapeId);
        if (at < 0 || !from) return;
        const part = deck.slideParts[at].part;
        const brought = bringShape(review.reviewed, from.part, from.xml, deck, part);
        c.addedAs = String(deck.pasteShape(at, { xml: brought, rels: [] }));
      } else if (!accept && here) deck.removeShape(here.index, c.addedAs);
      return;
    }
    case 'shapeDeleted': {
      const here = shapeIn(deck, c.slideId, c.shapeId);
      if (accept && here) deck.removeShape(here.index, c.shapeId);
      else if (!accept && !here) {
        const was = shapeIn(review.base, c.slideId, c.shapeId);
        const at = slideIndexOf(deck, c.slideId);
        if (!was || at < 0) return;
        const part = deck.slideParts[at].part;
        const xml = deck.pkg.text(part);
        // Back where it stood among the shapes, before the one that followed it.
        const after = shapesOf(review.base.pkg.text(was.part)).slice(shapesOf(review.base.pkg.text(was.part)).findIndex((s) => s.id === was.id) + 1).map((s) => s.id);
        const next = shapesOf(xml).find((s) => after.includes(s.id));
        const at2 = next ? next.start : xml.lastIndexOf('</p:spTree>');
        deck.replaceSlideXml(part, xml.slice(0, at2) + bringShape(review.base, was.part, was.xml, deck, part) + xml.slice(at2));
      }
      return;
    }
    case 'picture': {
      const here = shapeIn(deck, c.slideId, c.shapeId);
      const from = accept ? shapeIn(review.reviewed, c.reviewedSlideId, c.reviewedShapeId) : shapeIn(review.base, c.slideId, c.shapeId);
      if (!here || !from) return;
      const id = blipOf(from.xml);
      const brought = bringShape(source, from.part, `<a:blip r:embed="${id}"/>`, deck, here.part);
      const rId = /r:embed="([^"]*)"/.exec(brought)[1];
      writeShape(deck, here.index, here, here.xml.replace(/(<a:blip\b[^>]*r:embed=")[^"]*"/, (m, a) => a + rId + '"'));
      return;
    }
    case 'moved':
    case 'resized': {
      // Through the deck's own setGeometry, which writes a box for a placeholder that had none.
      const here = shapeIn(deck, c.slideId, c.shapeId);
      const from = accept ? shapeIn(review.reviewed, c.reviewedSlideId, c.reviewedShapeId) : shapeIn(review.base, c.slideId, c.shapeId);
      if (!here || !from) return;
      const now = deck.slide(here.index).shapes.find((s) => String(s.id) === here.id)?.geometry;
      const to = (accept ? review.reviewed : review.base).slide(from.index).shapes.find((s) => String(s.id) === from.id)?.geometry;
      if (!now || !to) return;
      const g = c.kind === 'moved' ? { x: to.x, y: to.y, w: now.w, h: now.h, rot: to.rot || 0 } : { x: now.x, y: now.y, w: to.w, h: to.h };
      deck.setGeometry(here.index, here.id, g);
      // Put back where the layout had it: a box it never stated is taken off again.
      if (!accept && !xfrmOf(from.xml)) {
        const after = shapeIn(deck, c.slideId, c.shapeId);
        const was = placed(review.base, from.index, from.id);
        const is = placed(deck, after.index, after.id);
        if (after && was && is && JSON.stringify(was) === JSON.stringify(is)) writeShape(deck, after.index, after, after.xml.replace(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/, '').replace(/<p:spPr><\/p:spPr>/, () => (/<p:spPr\/>/.test(from.xml) ? '<p:spPr/>' : '<p:spPr></p:spPr>')));
      }
      return;
    }
    default: {
      const here = shapeIn(deck, c.slideId, c.shapeId);
      const from = accept ? shapeIn(review.reviewed, c.reviewedSlideId, c.reviewedShapeId) : shapeIn(review.base, c.slideId, c.shapeId);
      if (!here || !from) return;
      writeShape(deck, here.index, here, withAspect(c.kind, here.xml, from.xml));
    }
  }
}

/**
 * Accept or reject changes, by id: one, a slide's (every change on it) or
 * all of them. One step Undo takes back.
 */
export function decideChanges(deck, review, ids, accept) {
  const wanted = new Set(ids);
  const snap = deck.snapshot();
  // Slides first in the order that keeps positions meaningful: deletions last when accepting.
  const list = review.changes.filter((c) => wanted.has(c.id));
  const order = (c) => (c.kind === 'slideAdded' ? 0 : c.kind === 'slideMoved' ? 2 : c.kind === 'slideDeleted' ? 3 : 1);
  list.sort((x, y) => order(x) - order(y));
  for (const c of list) setChange(deck, review, c, accept);
  deck.pushUndo(snap);
  return reviewState(deck, review);
}

/** A slide of another deck copied into this one at `at`: its layout matched by name, its pictures and parts brought along. Answers the new slide's id. */
export function importSlide(fromDeck, fromIndex, deck, at) {
  const src = fromDeck.slideParts[fromIndex];
  if (!src) throw new RangeError('no slide at index ' + fromIndex);
  const layoutRel = fromDeck.pkg.rels(src.part).find((r) => /\/slideLayout$/.test(r.Type));
  const layoutName = layoutRel ? /<p:cSld\b[^>]*\bname="([^"]*)"/.exec(fromDeck.pkg.text(OoxmlPackage.resolveTarget(src.part, layoutRel.Target)))?.[1] : null;
  // A new slide on this deck's own layout of that name, or its first, then its shapes and background from the reviewer's.
  const index = deck.insertSlide(Math.max(-1, at - 1), { layout: 'blank' });
  const part = deck.slideParts[index].part;
  const layouts = deck.layoutList();
  const mine = layouts.find((l) => l.name === layoutName) || null;
  if (mine && typeof deck.applyLayout === 'function') deck.applyLayout(index, mine.part);
  const xml = deck.pkg.text(part);
  const theirs = fromDeck.pkg.text(src.part);
  const cSld = /<p:cSld\b[\s\S]*<\/p:cSld>/.exec(theirs)?.[0];
  if (cSld) {
    const brought = bringShape(fromDeck, src.part, cSld, deck, part);
    deck.replaceSlideXml(part, xml.replace(/<p:cSld\b[\s\S]*<\/p:cSld>/, () => brought));
  }
  return String(deck.slideParts[index].id);
}

/** A slide deleted on accept put back on reject: its entry restored where it stood, its part never having left the package. */
function restoreSlide(deck, base, slideId) {
  const b = slidesOf(base);
  const at = b.findIndex((s) => s.id === slideId);
  const entry = base.slideParts[at];
  if (!entry) return;
  const pres = deck.pkg.text('ppt/presentation.xml');
  const rel = deck.pkg.rels('ppt/presentation.xml').find((r) => r.Id === entry.rId && OoxmlPackage.resolveTarget('ppt/presentation.xml', r.Target) === entry.part);
  const rId = rel ? rel.Id : deck.pkg.addRelationshipTo('ppt/presentation.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', entry.part.slice('ppt/'.length));
  const tag = `<p:sldId id="${slideId}" r:id="${rId}"/>`;
  // Before the first slide that followed it in the base and is still here.
  const following = b.slice(at + 1).map((s) => s.id).find((id) => slideIndexOf(deck, id) >= 0);
  const next = following ? pres.replace(new RegExp(`<p:sldId\\b[^>]*\\bid="${following}"[^>]*/>`), (m) => tag + m) : pres.replace('</p:sldIdLst>', () => tag + '</p:sldIdLst>');
  deck.pkg.write_('ppt/presentation.xml', next);
  deck.reloadSlides();
}

/** A review begun: the open deck's copy as it is now, the reviewer's deck, and what differs. */
export function beginReview(deck, reviewed, { DeckClass, reviewer = null, name = null } = {}) {
  const base = DeckClass.open(Buffer.from(deck.save()));
  const review = { base, reviewed, reviewer, name, changes: compareDecks(base, reviewed) };
  return review;
}
