// Tick boxes on a form, and whether each is ticked: classical image
// processing of our own, no engine. A box is a blob of ink with four straight
// sides of a like length, a square about the size of the text around it; it
// is ticked when there is ink inside it (a tick, a cross, a filled dot).
//
//   import { findCheckboxes } from '@rutba/ocr/checkboxes';
//   findCheckboxes(grey)   // [{ box: { x, y, width, height }, checked, fill }]
//
// A Chinese or Japanese character drawn as a square (口, 日) has the same
// shape; given the words read from the page, a box lying under a word in
// those scripts is left out.

import { binarise, components, textHeight, unevenPaper } from './analyse.js';

const CJK = /[㐀-鿿豈-﫿]/u;

/** The longest run of ink along a row (or a column) of a binary picture, inside a rectangle. */
function longest(binary, x0, y0, x1, y1, alongRow) {
  const { width, data } = binary;
  const lanes = alongRow ? y1 - y0 + 1 : x1 - x0 + 1;
  const out = new Int32Array(lanes);
  for (let l = 0; l < lanes; l++) {
    let best = 0, run = 0, gap = 0;
    const n = alongRow ? x1 - x0 + 1 : y1 - y0 + 1;
    for (let i = 0; i < n; i++) {
      const x = alongRow ? x0 + i : x0 + l, y = alongRow ? y0 + l : y0 + i;
      if (data[y * width + x] === 0) {
        run += 1 + gap;
        gap = 0;
        if (run > best) best = run;
      } else if (run && !gap) gap = 1; // A one-pixel break in a printed side (a scan's dropout) does not end it.
      else { run = 0; gap = 0; }
    }
    out[l] = best;
  }
  return out;
}

/** The lanes where a run nearly as long as the longest lies: the first and last of them, and how thick the first is. */
function sides(runs) {
  let most = 0;
  for (const r of runs) if (r > most) most = r;
  const strong = [];
  runs.forEach((r, i) => { if (r >= most * 0.8) strong.push(i); });
  if (strong.length < 2) return null;
  let thick = 1;
  while (thick < strong.length && strong[thick] === strong[0] + thick) thick++;
  return { most, first: strong[0], last: strong[strong.length - 1], thick };
}

/**
 * The tick boxes of a page.
 * @param {{ width, height, data }} grey one byte a pixel
 * @param {{ textHeight?: number, words?: { text, box, confidence? }[], least?: number }} [options]
 *   textHeight: the text's height in pixels (measured when not given);
 *   words: what was read from the page, to leave out squares that are characters;
 *   least: the share of a box's inside that must be ink for it to be ticked
 * @returns {{ box: { x, y, width, height }, checked: boolean, fill: number }[]} top to bottom, then along the row
 */
export function findCheckboxes(grey, options = {}) {
  const binary = grey.data.every((v) => v === 0 || v === 255) ? grey : binarise(grey, { method: unevenPaper(grey) ? 'sauvola' : 'otsu' });
  const t = options.textHeight || textHeight(binary) || 20;
  const least = options.least ?? 0.05;
  const found = [];
  for (const c of components(binary)) {
    // About the size of a letter; a tick that runs over the box can make the blob longer than the box.
    if (c.width < Math.max(8, t * 0.5) || c.height < Math.max(8, t * 0.5) || c.width > t * 4 || c.height > t * 4) continue;
    const x0 = c.x, y0 = c.y, x1 = c.x + c.width - 1, y1 = c.y + c.height - 1;
    const across = sides(longest(binary, x0, y0, x1, y1, true));
    const down = sides(longest(binary, x0, y0, x1, y1, false));
    if (!across || !down) continue;
    const w = down.last - down.first + 1, h = across.last - across.first + 1;
    // Four sides: two long rows as far apart as the long columns are long, and the other way about; square.
    if (Math.abs(w - across.most) > across.most * 0.25 || Math.abs(h - down.most) > down.most * 0.25) continue;
    if (w / h < 0.7 || w / h > 1.4 || w < Math.max(8, t * 0.5) || w > t * 3) continue;
    // Thin sides: a solid square (a bullet) is not a box.
    const stroke = Math.max(across.thick, down.thick);
    if (stroke > Math.min(w, h) * 0.3) continue;
    const box = { x: x0 + down.first, y: y0 + across.first, width: w, height: h };
    // Square corners: ink at all four, which a round letter (o, D) or an open one (n, u) does not have.
    const corner = (cx, cy) => {
      const r = stroke + 1;
      for (let y = cy - r; y <= cy + r; y++) for (let x = cx - r; x <= cx + r; x++) if (x >= 0 && y >= 0 && x < binary.width && y < binary.height && binary.data[y * binary.width + x] === 0) return true;
      return false;
    };
    const right = box.x + w - 1, bottom = box.y + h - 1;
    if (!corner(box.x, box.y) || !corner(right, box.y) || !corner(box.x, bottom) || !corner(right, bottom)) continue;
    // A box stands apart: paper beside it left and right, where a letter has the next letter of its word.
    const clear = (xa, xb) => {
      let ink = 0, n = 0;
      for (let y = box.y; y <= bottom; y++) for (let x = xa; x <= xb; x++) {
        if (x < 0 || x >= binary.width) continue;
        n++;
        if (binary.data[y * binary.width + x] === 0) ink++;
      }
      return !n || ink / n <= 0.08;
    };
    const m = Math.max(3, Math.round(w * 0.2));
    if (!clear(box.x - m - 1, box.x - 2) || !clear(right + 2, right + m + 1)) continue;
    // Run lengths alone pass a letter's straight strokes; a box's four sides each run most of its length.
    const side = (ax, ay, bx, by) => {
      let ink = 0, n = 0;
      for (let k = 0; k <= 1; k += 1 / Math.max(ax === bx ? h : w, 1)) {
        const x = Math.round(ax + (bx - ax) * k), y = Math.round(ay + (by - ay) * k);
        n++;
        let hit = false;
        for (let d = -stroke; d <= stroke && !hit; d++) {
          const xx = ax === bx ? x + d : x, yy = ax === bx ? y : y + d;
          if (xx >= 0 && yy >= 0 && xx < binary.width && yy < binary.height && binary.data[yy * binary.width + xx] === 0) hit = true;
        }
        if (hit) ink++;
      }
      return ink / n;
    };
    const t1 = across.thick >> 1, l1 = down.thick >> 1;
    if (side(box.x, box.y + t1, right, box.y + t1) < 0.85 || side(box.x, bottom - t1, right, bottom - t1) < 0.85
      || side(box.x + l1, box.y, box.x + l1, bottom) < 0.85 || side(right - l1, box.y, right - l1, bottom) < 0.85) continue;
    // The inside, a stroke and a sixth of the side clear of the sides.
    const inset = stroke + Math.round(Math.min(w, h) / 6);
    const ix0 = box.x + inset, iy0 = box.y + inset, ix1 = box.x + w - 1 - inset, iy1 = box.y + h - 1 - inset;
    let ink = 0, area = 0;
    for (let y = iy0; y <= iy1; y++) for (let x = ix0; x <= ix1; x++) { area++; if (binary.data[y * binary.width + x] === 0) ink++; }
    if (area < 4) continue;
    const fill = ink / area;
    found.push({ box, checked: fill >= least, fill: Math.round(fill * 1000) / 1000 });
  }
  const words = options.words || [];
  const under = (b, wb) => {
    const ox = Math.min(b.x + b.width, wb.x + wb.width) - Math.max(b.x, wb.x);
    const oy = Math.min(b.y + b.height, wb.y + wb.height) - Math.max(b.y, wb.y);
    return ox > 0 && oy > 0 && (ox * oy) / (b.width * b.height) > 0.5;
  };
  // A character's word is about as wide as its square for each character; a box the reading took into the
  // word of the label beside it lies under a word far wider than itself.
  // An empty box standing alone is read as 口 (mouth), which it looks exactly like: that word is the box.
  const character = (f, w) => CJK.test(w.text) && !/^[口囗]$/u.test(w.text) && (w.confidence ?? 100) >= 70 && under(f.box, w.box) && w.box.width / [...w.text].length <= f.box.width * 1.6;
  const kept = found.filter((f) => !words.some((w) => character(f, w)));
  // A form's boxes are of one size: with three or more, one under two thirds of their middle size is a letter.
  const sizes = kept.map((f) => f.box.width).sort((a, b) => a - b);
  const middle = sizes[sizes.length >> 1] || 0;
  return kept
    .filter((f) => kept.length < 3 || f.box.width >= middle * 0.66)
    .sort((a, b) => (Math.abs(a.box.y - b.box.y) < t ? a.box.x - b.box.x : a.box.y - b.box.y));
}
