'use strict';

/**
 * How a PDF is found one's way around: its bookmarks (the outline), its
 * named destinations, and its page labels (the "iv" or "A-3" a page is
 * called). Read here, for assemble.js to carry through a reorder, an
 * extract, a split and a merge, and for the window to show.
 *
 * Named destinations live in the catalog's /Dests dictionary (PDF 1.1) or
 * its /Names /Dests name tree (1.2 and after); page labels in the
 * /PageLabels number tree. Both trees are walked with a guard against the
 * loops a damaged file can hold.
 */

const { Name, Ref, Dict, PdfString } = require('./reader');

const num = (v, fallback = 0) => (typeof v === 'number' && Number.isFinite(v) ? v : fallback);
const keyOf = (v) => (v instanceof PdfString ? v.bytes.toString('latin1') : v instanceof Name ? v.name : typeof v === 'string' ? v : null);

/** Every entry of a name tree (or, with `numbers`, a number tree) as [key, value] in order. */
function treeEntries(file, root, numbers = false) {
  const out = [];
  const seen = new Set();
  const walk = (ref, depth) => {
    if (depth > 32) return;
    if (ref instanceof Ref) { if (seen.has(ref.key)) return; seen.add(ref.key); }
    const node = file.resolve(ref);
    if (!(node instanceof Dict)) return;
    const list = file.resolve(node.get(numbers ? 'Nums' : 'Names'));
    if (Array.isArray(list)) for (let i = 0; i + 1 < list.length; i += 2) out.push([numbers ? num(file.resolve(list[i])) : keyOf(file.resolve(list[i])), list[i + 1]]);
    const kids = file.resolve(node.get('Kids'));
    if (Array.isArray(kids)) for (const k of kids) walk(k, depth + 1);
  };
  walk(root, 0);
  return out;
}

/** The named destinations: Map of name -> explicit destination array (the source's own values). */
function namedDestinations(file) {
  if (file._named) return file._named;
  const map = new Map();
  const catalog = file.catalogOrNull();
  const explicit = (v) => {
    const d = file.resolve(v);
    if (Array.isArray(d)) return d;
    if (d instanceof Dict) { const inner = file.resolve(d.get('D')); return Array.isArray(inner) ? inner : null; }
    return null;
  };
  if (catalog) {
    const dests = file.resolve(catalog.get('Dests'));
    if (dests instanceof Dict) for (const [k, v] of dests.entries()) { const e = explicit(v); if (e) map.set(k, e); }
    const names = file.resolve(catalog.get('Names'));
    if (names instanceof Dict && names.get('Dests')) {
      for (const [k, v] of treeEntries(file, names.get('Dests'))) { const e = explicit(v); if (k !== null && e && !map.has(k)) map.set(k, e); }
    }
  }
  file._named = map;
  return map;
}

/** A destination as an explicit array: itself, or the named one it names; null when it names nothing. */
function resolveDest(file, dest) {
  const d = file.resolve(dest);
  if (Array.isArray(d)) return d;
  const key = keyOf(d);
  if (key === null) return null;
  return namedDestinations(file).get(key) || null;
}

const isNamed = (file, v) => {
  const d = file.resolve(v);
  return d instanceof PdfString || d instanceof Name;
};

/**
 * The outline as a tree: [{ title, dest, action, color, bold, italic, open, kids }],
 * `dest` an explicit destination (resolved from a name or a GoTo action),
 * `action` any other action dictionary (a web link, say), both the source's
 * own values.
 */
function readOutline(file) {
  const catalog = file.catalogOrNull();
  const root = catalog ? file.resolve(catalog.get('Outlines')) : null;
  if (!(root instanceof Dict)) return [];
  const seen = new Set();
  let budget = 100000;
  const level = (firstRef, depth) => {
    const out = [];
    let ref = firstRef;
    while (ref && budget-- > 0 && depth < 64) {
      if (ref instanceof Ref) { if (seen.has(ref.key)) break; seen.add(ref.key); }
      const item = file.resolve(ref);
      if (!(item instanceof Dict)) break;
      const t = file.resolve(item.get('Title'));
      let dest = null;
      let action = null;
      if (item.get('Dest')) dest = resolveDest(file, item.get('Dest'));
      const a = file.resolve(item.get('A'));
      if (a instanceof Dict) {
        const s = a.get('S');
        if (s instanceof Name && s.name === 'GoTo') dest = resolveDest(file, a.get('D'));
        else action = a;
      }
      const flags = num(file.resolve(item.get('F')));
      const c = file.resolve(item.get('C'));
      out.push({
        title: t instanceof PdfString ? t.text() : '',
        dest,
        action,
        color: Array.isArray(c) && c.length === 3 ? c.map((x) => num(file.resolve(x))) : null,
        italic: Boolean(flags & 1),
        bold: Boolean(flags & 2),
        open: num(file.resolve(item.get('Count'))) > 0,
        kids: level(item.get('First'), depth + 1),
      });
      ref = item.get('Next');
    }
    return out;
  };
  return level(root.get('First'), 0);
}

/**
 * Each page's label as its file defines it, or null when the file has no
 * labels: [{ style: 'D'|'R'|'r'|'A'|'a'|null, prefix, number }].
 */
function pageLabels(file) {
  const catalog = file.catalogOrNull();
  const tree = catalog ? catalog.get('PageLabels') : null;
  if (!tree) return null;
  const ranges = treeEntries(file, tree, true).filter(([k]) => Number.isInteger(k) && k >= 0).sort((a, b) => a[0] - b[0]);
  if (!ranges.length) return null;
  const count = file.pageCount;
  const out = [];
  for (let p = 0; p < count; p++) {
    let range = null;
    for (const r of ranges) if (r[0] <= p) range = r;
    if (!range) { out.push({ style: 'D', prefix: '', number: p + 1 }); continue; }
    const d = file.resolve(range[1]);
    const style = d instanceof Dict && d.get('S') instanceof Name ? d.get('S').name : null;
    const prefix = d instanceof Dict && file.resolve(d.get('P')) instanceof PdfString ? file.resolve(d.get('P')).text() : '';
    const start = d instanceof Dict ? num(file.resolve(d.get('St')), 1) : 1;
    out.push({ style, prefix, number: start + (p - range[0]) });
  }
  return out;
}

const ROMAN = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
/** A label as a reader shows it: "iv", "A-3", "C". */
function labelText({ style, prefix, number }) {
  let n = '';
  if (style === 'D') n = String(number);
  else if (style === 'R' || style === 'r') {
    let v = number;
    for (const [value, letters] of ROMAN) while (v >= value) { n += letters; v -= value; }
    if (style === 'R') n = n.toUpperCase();
  } else if (style === 'A' || style === 'a') {
    const letter = String.fromCharCode(65 + ((number - 1) % 26));
    n = letter.repeat(Math.floor((number - 1) / 26) + 1);
    if (style === 'a') n = n.toLowerCase();
  }
  return `${prefix || ''}${n}`;
}

/**
 * The links on every page: [{ page, rect, target, url }], with target the
 * index of the page a link goes to (through a name, a GoTo action or a
 * destination of its own), and url a web address it opens.
 */
function readLinks(file) {
  const pages = file.pages();
  const index = new Map(pages.map((p, i) => [p.ref ? p.ref.num : -1, i]));
  const out = [];
  pages.forEach((page, p) => {
    const list = file.resolve(page.dict.get('Annots'));
    if (!Array.isArray(list)) return;
    for (const item of list) {
      const d = file.resolve(item);
      if (!(d instanceof Dict) || !(d.get('Subtype') instanceof Name) || d.get('Subtype').name !== 'Link') continue;
      const r = file.resolve(d.get('Rect'));
      if (!Array.isArray(r) || r.length < 4) continue;
      const v = r.map((x) => num(file.resolve(x)));
      let dest = d.get('Dest') ? resolveDest(file, d.get('Dest')) : null;
      let url = null;
      const a = file.resolve(d.get('A'));
      if (a instanceof Dict && a.get('S') instanceof Name) {
        if (a.get('S').name === 'GoTo') dest = resolveDest(file, a.get('D'));
        else if (a.get('S').name === 'URI') { const u = file.resolve(a.get('URI')); url = u instanceof PdfString ? u.bytes.toString('latin1') : null; }
      }
      const target = dest && dest[0] instanceof Ref ? (index.has(dest[0].num) ? index.get(dest[0].num) : null) : typeof dest?.[0] === 'number' ? dest[0] : null;
      if (target === null && !url) continue;
      out.push({ page: p, rect: [Math.min(v[0], v[2]), Math.min(v[1], v[3]), Math.max(v[0], v[2]), Math.max(v[1], v[3])], target, url });
    }
  });
  return out;
}

/** The outline with each bookmark's page as an index of the file's pages (null when it names none). */
function outlinePages(file) {
  const index = new Map(file.pages().map((p, i) => [p.ref ? p.ref.num : -1, i]));
  const map = (nodes) => nodes.map((n) => ({
    title: n.title,
    page: n.dest && n.dest[0] instanceof Ref ? (index.has(n.dest[0].num) ? index.get(n.dest[0].num) : null) : typeof n.dest?.[0] === 'number' ? n.dest[0] : null,
    url: n.action && n.action.get('S') instanceof Name && n.action.get('S').name === 'URI' && file.resolve(n.action.get('URI')) instanceof PdfString ? file.resolve(n.action.get('URI')).bytes.toString('latin1') : null,
    bold: n.bold, italic: n.italic, color: n.color, open: n.open,
    kids: map(n.kids),
  }));
  return map(readOutline(file));
}

module.exports = { treeEntries, namedDestinations, resolveDest, isNamed, readOutline, outlinePages, readLinks, pageLabels, labelText };
