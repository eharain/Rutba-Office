// Insert → SmartArt, for all three apps: a diagram laid out from lines of
// text — a list, a process, a cycle, a hierarchy — as shapes the apps
// already write (filled shapes holding their words, and the lines, arrows
// and ring between them), put in as one group named for the layout. Office
// opens it as a group of shapes, every word editable; it is not a SmartArt
// part, which carries Office's own layout definitions. This module is the
// arithmetic only: where each shape goes, and the gallery's small pictures.

import { t } from '@rutba/office-ui/messages';

/** The gallery, in Office's names and groups. */
export const SMARTART_LAYOUTS = [
  { id: 'blockList', name: t('Basic Block List'), group: 'List' },
  { id: 'verticalList', name: t('Vertical Box List'), group: 'List' },
  { id: 'process', name: t('Basic Process'), group: 'Process' },
  { id: 'chevron', name: t('Basic Chevron Process'), group: 'Process' },
  { id: 'cycle', name: t('Basic Cycle'), group: 'Cycle' },
  { id: 'hierarchy', name: t('Hierarchy'), group: 'Hierarchy' },
];

/**
 * Office's first SmartArt style, Colored Fill — Accent 1: every shape the
 * theme's first accent with a white edge, arrows and the ring a tint of it,
 * lines a shade. `scheme` is the theme colour; the hex is Office's own
 * theme's, for a reader that has no theme to look it up in.
 */
export const SMARTART_LOOK = {
  node: { fill: '#4472C4', scheme: 'accent1', line: '#FFFFFF' },
  arrow: { fill: '#B4C7E7' },
  ring: { line: '#B4C7E7' },
  line: { line: '#2F5597' },
};

/** Lines of text as items: a line begun with a tab (or two spaces) a level under the one above. */
export function parseItems(text) {
  return String(text || '').split(/\r?\n/).filter((l) => l.trim()).map((l) => {
    const lead = /^(?:\t| {2})*/.exec(l)[0];
    const level = (lead.match(/\t| {2}/g) || []).length;
    return { text: l.trim().replace(/^[•\-*]\s+/, ''), level: Math.min(2, level) };
  });
}

/** The items as lines of text again, a level a tab in. */
export function itemsText(items) {
  return (items || []).map((i) => '\t'.repeat(i.level || 0) + i.text).join('\n');
}

/**
 * The type size (points) that fits `text` in a shape `w` by `h` px — the
 * largest from 32 down to 9 at which its words, wrapped, stand inside it,
 * as Office shrinks a SmartArt shape's words to fit.
 */
export function fitSize(text, w, h) {
  const words = String(text || '').split(/\s+/).filter(Boolean);
  const room = { w: w * 0.84, h: h * 0.8 };
  for (let pt = 32; pt > 9; pt -= 1) {
    const em = pt * (96 / 72);
    const charW = em * 0.52;
    if (words.some((word) => word.length * charW > room.w)) continue;
    let lines = 1;
    let used = 0;
    for (const word of words) {
      const len = word.length * charW;
      if (used && used + charW + len > room.w) { lines += 1; used = len; } else used += (used ? charW : 0) + len;
    }
    if (lines * em * 1.2 <= room.h) return pt;
  }
  return 9;
}

/**
 * The shapes of one layout for `items` in `box` (`{ x, y, w, h }`, px): each
 * `{ role, preset, x, y, w, h, text, size, fill, scheme, line, lineWidth }`
 * in drawing order — `role` 'node' a filled shape with its words (`size` in
 * points), 'arrow' a block arrow between two, 'ring' an unfilled circle
 * behind a cycle, 'line' a straight line (one of its sides 0).
 */
export function layoutSmartArt(layout, items, box) {
  const top = items.filter((i) => i.level === 0);
  const nodes = top.length ? top : items;
  const out = [];
  const r = (v) => Math.round(v);
  const node = (preset, x, y, w, h, text) => {
    const W = Math.max(8, r(w));
    const H = Math.max(8, r(h));
    out.push({ role: 'node', preset, x: r(x), y: r(y), w: W, h: H, text, size: fitSize(text, preset === 'ellipse' ? W * 0.72 : preset === 'chevron' ? W - H : W, preset === 'ellipse' ? H * 0.72 : H), fill: SMARTART_LOOK.node.fill, scheme: SMARTART_LOOK.node.scheme, line: SMARTART_LOOK.node.line, lineWidth: 1 });
  };
  const line = (x, y, w, h) => out.push({ role: 'line', preset: 'line', x: r(x), y: r(y), w: r(w), h: r(h), text: '', fill: null, line: SMARTART_LOOK.line.line, lineWidth: 1.5 });
  const n = Math.max(1, nodes.length);
  if (layout === 'blockList') {
    // Rows of equal boxes, as many to a row as keep them about as wide as the box is.
    const cols = Math.max(1, Math.min(n, Math.ceil(Math.sqrt((n * box.w) / (box.h * 1.6)))));
    const rows = Math.ceil(n / cols);
    const gx = box.w * 0.04;
    const gy = box.h * 0.06;
    const w = (box.w - gx * (cols - 1)) / cols;
    const h = Math.min((box.h - gy * (rows - 1)) / rows, w * 0.6);
    const oy = box.y + (box.h - (rows * h + (rows - 1) * gy)) / 2;
    nodes.forEach((it, i) => {
      const row = Math.floor(i / cols);
      const inRow = Math.min(cols, n - row * cols);
      const ox = box.x + (box.w - (inRow * w + (inRow - 1) * gx)) / 2;
      node('rect', ox + (i % cols) * (w + gx), oy + row * (h + gy), w, h, it.text);
    });
  } else if (layout === 'verticalList') {
    const gy = Math.min(box.h * 0.04, 16);
    const h = Math.min((box.h - gy * (n - 1)) / n, box.h * 0.3);
    const oy = box.y + (box.h - (n * h + (n - 1) * gy)) / 2;
    nodes.forEach((it, i) => node('roundRect', box.x, oy + i * (h + gy), box.w, h, it.text));
  } else if (layout === 'process') {
    // Boxes in a row, an arrow in each gap.
    const gap = Math.min(box.w * 0.1, 80);
    const w = (box.w - gap * (n - 1)) / n;
    const h = Math.min(box.h, w * 0.6);
    const y = box.y + (box.h - h) / 2;
    const aw = gap * 0.62;
    const ah = Math.min(h * 0.4, aw * 1.15);
    nodes.forEach((it, i) => {
      node('roundRect', box.x + i * (w + gap), y, w, h, it.text);
      if (i < n - 1) out.push({ role: 'arrow', preset: 'rightArrow', x: r(box.x + i * (w + gap) + w + (gap - aw) / 2), y: r(y + (h - ah) / 2), w: r(aw), h: r(ah), text: '', fill: SMARTART_LOOK.arrow.fill, line: null });
    });
  } else if (layout === 'chevron') {
    // Chevrons nose to tail, each point a third into the next's notch.
    const h0 = Math.min(box.h, (box.w / n) * 0.5);
    const tip = h0 / 2;
    const gap = tip * 0.12;
    const w = (box.w + (n - 1) * (tip - gap)) / n;
    const y = box.y + (box.h - h0) / 2;
    nodes.forEach((it, i) => node('chevron', box.x + i * (w - tip + gap), y, w, h0, it.text));
  } else if (layout === 'cycle') {
    // Circles round a ring, the first at the top, going clockwise, all inside the box.
    const m = Math.min(box.w, box.h);
    const size = Math.min(m * 0.34, ((2 * Math.PI * m * 0.33) / n) * 0.75);
    const ring = (m - size) / 2;
    const cx = box.x + box.w / 2;
    const cy = box.y + box.h / 2;
    if (n > 1) out.push({ role: 'ring', preset: 'ellipse', x: r(cx - ring), y: r(cy - ring), w: r(ring * 2), h: r(ring * 2), text: '', fill: null, line: SMARTART_LOOK.ring.line, lineWidth: 3 });
    nodes.forEach((it, i) => {
      const a = -Math.PI / 2 + (2 * Math.PI * i) / n;
      node('ellipse', cx + ring * Math.cos(a) - size / 2, cy + ring * Math.sin(a) - size / 2, size, size, it.text);
    });
  } else if (layout === 'hierarchy') {
    // Each top item over the items under it, a line down, across and down to each.
    const groups = [];
    for (const it of items) {
      if (it.level === 0 || !groups.length) groups.push({ head: it, kids: [] });
      else groups[groups.length - 1].kids.push(it);
    }
    const cols = Math.max(1, groups.reduce((m, g) => m + Math.max(1, g.kids.length), 0));
    const gx = box.w * 0.03;
    const w = (box.w - gx * (cols - 1)) / cols;
    const h = Math.min(box.h * 0.36, w * 0.62);
    const anyKids = groups.some((g) => g.kids.length);
    const yHead = anyKids ? box.y : box.y + (box.h - h) / 2;
    const yKid = box.y + box.h - h;
    let col = 0;
    groups.forEach((g) => {
      const span = Math.max(1, g.kids.length);
      const left = box.x + col * (w + gx);
      const hx = left + (span * w + (span - 1) * gx - w) / 2;
      if (g.kids.length) {
        const midY = (yHead + h + yKid) / 2;
        const kx = g.kids.map((_, k) => left + k * (w + gx) + w / 2);
        line(hx + w / 2, yHead + h, 0, midY - (yHead + h));
        if (kx.length > 1) line(kx[0], midY, kx[kx.length - 1] - kx[0], 0);
        kx.forEach((x) => line(x, midY, 0, yKid - midY));
      }
      node('roundRect', hx, yHead, w, h, g.head.text);
      g.kids.forEach((kid, k) => node('roundRect', left + k * (w + gx), yKid, w, h, kid.text));
      col += span;
    });
  }
  // One size for all the words, the smallest that fits, as Office sizes a diagram's shapes alike.
  const size = Math.min(...out.filter((s) => s.role === 'node').map((s) => s.size));
  for (const s of out) if (s.role === 'node') s.size = size;
  return out;
}

/** The box a diagram stands in, its items placed by layout — the shapes' own bounds. */
export function boundsOf(shapes) {
  if (!shapes.length) return { x: 0, y: 0, w: 0, h: 0 };
  const x0 = Math.min(...shapes.map((s) => s.x));
  const y0 = Math.min(...shapes.map((s) => s.y));
  const x1 = Math.max(...shapes.map((s) => s.x + s.w));
  const y1 = Math.max(...shapes.map((s) => s.y + s.h));
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

const SAMPLE = [{ text: t('Text'), level: 0 }, { text: t('Text'), level: 0 }, { text: t('Text'), level: 0 }];
const SAMPLE_TREE = [{ text: t('Text'), level: 0 }, { text: t('Text'), level: 1 }, { text: t('Text'), level: 1 }, { text: t('Text'), level: 1 }];

/** A layout drawn small, as SVG — the gallery's pictures and the preview. */
export function smartArtSvg(layout, items = [], { width = 240, height = 150, labels = true } = {}) {
  const list = items.length ? items : layout === 'hierarchy' ? SAMPLE_TREE : SAMPLE;
  const shapes = layoutSmartArt(layout, list, { x: 6, y: 6, w: width - 12, h: height - 12 });
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  const body = shapes.map((s) => {
    const { x, y, w, h } = s;
    const stroke = s.line ? ` stroke="${s.line}" stroke-width="${s.role === 'node' ? 1 : s.lineWidth || 1.5}"` : '';
    const fill = s.fill ? s.fill : 'none';
    if (s.role === 'line') return `<line x1="${x}" y1="${y}" x2="${x + w}" y2="${y + h}"${stroke}/>`;
    const m = Math.min(w, h);
    const pts = (list) => `<polygon points="${list.map(([px, py]) => `${Math.round(px * 10) / 10},${Math.round(py * 10) / 10}`).join(' ')}" fill="${fill}"${stroke}/>`;
    const shape = s.preset === 'ellipse'
      ? `<ellipse cx="${x + w / 2}" cy="${y + h / 2}" rx="${w / 2}" ry="${h / 2}" fill="${fill}"${stroke}/>`
      : s.preset === 'chevron'
        ? pts([[x, y], [x + w - m / 2, y], [x + w, y + h / 2], [x + w - m / 2, y + h], [x, y + h], [x + m / 2, y + h / 2]])
        : s.preset === 'rightArrow'
          ? pts([[x, y + h / 4], [x + w - m / 2, y + h / 4], [x + w - m / 2, y], [x + w, y + h / 2], [x + w - m / 2, y + h], [x + w - m / 2, y + h * 0.75], [x, y + h * 0.75]])
          : `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${s.preset === 'roundRect' ? Math.round(m * 0.16667 * 10) / 10 : 0}" fill="${fill}"${stroke}/>`;
    const label = labels && s.text
      ? `<text x="${x + w / 2}" y="${y + h / 2}" font-size="${Math.max(6, Math.min(12, h * 0.26))}" fill="#fff" text-anchor="middle" dominant-baseline="central" font-family="Segoe UI, sans-serif">${esc(s.text.length > 16 ? s.text.slice(0, 15) + '…' : s.text)}</text>`
      : '';
    return shape + label;
  }).join('');
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${body}</svg>`;
}
