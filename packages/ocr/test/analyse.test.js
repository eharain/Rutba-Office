// Page analysis and layout of our own (analyse.js, layout.js), proved on
// pages drawn here with known answers: text turned by a known angle, paper
// lit unevenly, ruled and unruled tables, two columns.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { otsu, binarise, unevenPaper, estimateSkew, rotate, unrotatePoint, scale, components, textHeight, rulings, fitToInk } from '@rutba/ocr/analyse';
import { groupLines, columns, tableCells } from '@rutba/ocr/layout';

/** A blank grey page. */
const page = (width, height, paper = 255) => ({ width, height, data: new Uint8Array(width * height).fill(paper) });
const fill = (p, x, y, w, h, v = 0) => {
  for (let yy = Math.max(0, y); yy < Math.min(p.height, y + h); yy++) for (let xx = Math.max(0, x); xx < Math.min(p.width, x + w); xx++) p.data[yy * p.width + xx] = v;
};

/**
 * Lines of text-like ink: letters 7 px wide and 14 tall (every fifth an
 * ascender of 20), 2 px apart, words of 3 to 7 letters 9 px apart.
 */
function textPage(width, height, { lines = 12, top = 40, left = 40, pitch = 36, ink = 0, paper = 255 } = {}) {
  const p = page(width, height, paper);
  let seed = 7;
  const rand = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
  for (let l = 0; l < lines; l++) {
    let x = left;
    const base = top + l * pitch;
    while (x < width - left - 70) {
      const letters = 3 + Math.floor(rand() * 5);
      for (let i = 0; i < letters; i++) {
        const tall = i % 5 === 0;
        fill(p, x, base - (tall ? 20 : 14), 7, tall ? 20 : 14, ink);
        x += 9;
      }
      x += 9;
    }
  }
  return p;
}

const inkCount = (b) => b.data.reduce((n, v) => n + (v === 0 ? 1 : 0), 0);

test('otsu: the threshold falls between the ink and the paper', () => {
  const p = textPage(400, 200, { lines: 4, ink: 40, paper: 220 });
  const t = otsu(p);
  assert.ok(t >= 40 && t < 220, `threshold ${t}`);
});

test('binarise: on paper lit from one side Otsu loses the dark side, Sauvola keeps every letter and no paper; auto chooses it', () => {
  const p = textPage(600, 300, { lines: 6 });
  const truth = binarise(p, { method: 'otsu' });
  // Light falling off to the left: paper from 110 to 250, ink 70 levels below the paper under it.
  for (let y = 0; y < p.height; y++) {
    for (let x = 0; x < p.width; x++) {
      const paper = 110 + Math.round((140 * x) / p.width);
      const i = y * p.width + x;
      p.data[i] = truth.data[i] === 0 ? paper - 70 : paper;
    }
  }
  assert.equal(unevenPaper(p), true);
  assert.equal(unevenPaper(textPage(600, 300)), false);
  const flat = binarise(p, { method: 'otsu' });
  const local = binarise(p, { method: 'sauvola' });
  const auto = binarise(p);
  assert.equal(auto.method, 'sauvola');
  const wrong = (b) => {
    let n = 0;
    for (let i = 0; i < b.data.length; i++) if (b.data[i] !== truth.data[i]) n++;
    return n / b.data.length;
  };
  assert.ok(wrong(flat) > 0.1, `Otsu wrong on ${(wrong(flat) * 100).toFixed(1)}% of pixels`);
  assert.ok(wrong(local) < 0.02, `Sauvola wrong on ${(wrong(local) * 100).toFixed(2)}% of pixels`);
});

test('estimateSkew: text turned by a known angle is measured to within a tenth of a degree, either way', () => {
  const straight = textPage(900, 600, { lines: 14 });
  assert.ok(Math.abs(estimateSkew(binarise(straight, { method: 'otsu' }))) <= 0.1);
  for (const angle of [3, -2, 0.6, -5.5]) {
    const turned = rotate(straight, angle);
    const got = estimateSkew(binarise(turned, { method: 'otsu' }));
    assert.ok(Math.abs(got - angle) <= 0.1, `turned ${angle}, measured ${got}`);
    // Turned back by what was measured, the text lies level again.
    const level = rotate(turned, -got);
    assert.ok(Math.abs(estimateSkew(binarise(level, { method: 'otsu' }))) <= 0.1);
  }
  // A page big enough to be sampled every second row (a 300 dpi scan) is measured as well, not held at 0.
  const big = textPage(2300, 1100, { lines: 28 });
  for (const angle of [1.2, -1.2]) {
    const got = estimateSkew(binarise(rotate(big, angle), { method: 'otsu' }));
    assert.ok(Math.abs(got - angle) <= 0.1, `big page turned ${angle}, measured ${got}`);
  }
});

test('rotate and unrotatePoint agree: a dot turned lands where the inverse says it came from', () => {
  const p = page(200, 120);
  fill(p, 150, 30, 3, 3, 0);
  const turned = rotate(p, 10);
  let sx = 0, sy = 0, n = 0;
  for (let y = 0; y < 120; y++) for (let x = 0; x < 200; x++) if (turned.data[y * 200 + x] < 128) { sx += x; sy += y; n++; }
  const back = unrotatePoint(sx / n, sy / n, 10, 200, 120);
  assert.ok(Math.abs(back.x - 151) < 1.2 && Math.abs(back.y - 31) < 1.2, JSON.stringify(back));
  // Counter-clockwise on screen: a dot right of centre moves up.
  assert.ok(sy / n < 31, `moved from 31 to ${(sy / n).toFixed(1)}`);
});

test('scale: double and half sizes, the page\'s mean kept', () => {
  const p = textPage(300, 150, { lines: 3 });
  const mean = (x) => x.data.reduce((a, b) => a + b, 0) / x.data.length;
  const up = scale(p, 2), down = scale(p, 0.5);
  assert.deepEqual([up.width, up.height, down.width, down.height], [600, 300, 150, 75]);
  assert.ok(Math.abs(mean(up) - mean(p)) < 3 && Math.abs(mean(down) - mean(p)) < 3);
});

test('components and textHeight: each letter one blob; the text height is the common letter height, not a rule\'s or a speck\'s', () => {
  const p = textPage(500, 200, { lines: 4 });
  fill(p, 20, 180, 460, 2);
  fill(p, 5, 5, 1, 1);
  const b = binarise(p, { method: 'otsu' });
  const blobs = components(b);
  const letters = blobs.filter((x) => x.height === 14 || x.height === 20);
  assert.ok(letters.length > 100, `${letters.length} letters`);
  assert.ok(blobs.some((x) => x.width === 460 && x.height === 2), 'the rule is one blob');
  assert.equal(textHeight(b), 14);
  // Letters that touch diagonally are one blob (eight-connected).
  const d = page(10, 10);
  fill(d, 1, 1, 2, 2); fill(d, 3, 3, 2, 2);
  assert.equal(components(binarise(d, { method: 'otsu' })).length, 1);
});

/** A grid of rules: `rows` by `cols` cells of w by h, lines `t` thick, at (x0, y0). */
function grid(p, { x0, y0, rows, cols, w, h, t = 2 }) {
  for (let r = 0; r <= rows; r++) fill(p, x0, y0 + r * h, cols * w + t, t);
  for (let c = 0; c <= cols; c++) fill(p, x0 + c * w, y0, t, rows * h + t);
}

test('rulings: the rules of a grid are found with their positions, a dropout bridged, letters and a thick block not taken for rules', () => {
  const p = page(700, 400);
  grid(p, { x0: 50, y0: 60, rows: 4, cols: 3, w: 200, h: 60 });
  fill(p, 300, 120, 2, 2, 255); // a two-pixel dropout in one rule
  fill(p, 70, 80, 7, 14); fill(p, 80, 80, 7, 14); // letters inside a cell
  fill(p, 60, 340, 400, 30); // a block, not a rule
  const { horizontal, vertical } = rulings(binarise(p, { method: 'otsu' }));
  assert.deepEqual(horizontal.map((r) => r.y), [60.5, 120.5, 180.5, 240.5, 300.5]);
  assert.ok(horizontal.every((r) => r.x0 <= 51 && r.x1 >= 650 && r.thickness === 2), JSON.stringify(horizontal));
  assert.deepEqual(vertical.map((r) => r.x), [50.5, 250.5, 450.5, 650.5]);
});

/** Words as the recogniser gives them. */
const word = (text, x, y, w = text.length * 8, h = 14) => ({ text, box: { x, y, width: w, height: h } });

test('groupLines: words with a ragged baseline gather into their lines, left to right, or right to left', () => {
  const words = [word('world', 70, 12), word('Hello', 10, 10), word('next', 10, 40), word('line', 52, 43)];
  const lines = groupLines(words);
  assert.deepEqual(lines.map((l) => l.words.map((w) => w.text).join(' ')), ['Hello world', 'next line']);
  assert.deepEqual(groupLines(words, { rtl: true }).map((l) => l.words[0].text), ['world', 'line']);
  assert.deepEqual(lines[0].box, { x: 10, y: 10, width: 100, height: 16 });
});

test('columns: two columns of text split at the gutter; a heading across both does not join them', () => {
  const words = [word('Heading across the page', 20, 0, 560)];
  for (let r = 0; r < 12; r++) {
    words.push(word('left', 20, 30 + r * 20, 100), word('text', 130, 30 + r * 20, 120));
    words.push(word('right', 330, 30 + r * 20, 110), word('side', 450, 30 + r * 20, 130));
  }
  const cols = columns(words);
  assert.equal(cols.length, 2);
  assert.ok(cols[0].x1 <= 330 && cols[1].x0 >= 250, JSON.stringify(cols));
  assert.equal(columns(words.filter((w) => w.box.x < 300)).length, 1);
});

test('tableCells: a table set with space alone becomes rows and columns, a cell\'s words joined; numbers stay in their column', () => {
  const words = [
    word('Item', 10, 10), word('Qty', 200, 10), word('Price', 300, 10),
    word('Blue', 10, 40), word('pens', 50, 40), word('12', 205, 40), word('1.50', 300, 40),
    word('Paper', 10, 70), word('3', 210, 70), word('4.75', 300, 70),
  ];
  const t = tableCells(words);
  assert.equal(t.ruled, false);
  assert.deepEqual(t.cells, [['Item', 'Qty', 'Price'], ['Blue pens', '12', '1.50'], ['Paper', '3', '4.75']]);
  const rtl = tableCells(words, undefined, { rtl: true });
  assert.deepEqual(rtl.cells[0], ['Price', 'Qty', 'Item']);
});

test('tableCells: ruling lines decide the cells, an empty cell stays empty, and a word in a cell joins only its own cell', () => {
  const rules = {
    horizontal: [0, 30, 60, 90].map((y) => ({ y, x0: 0, x1: 400, thickness: 1 })),
    vertical: [0, 150, 280, 400].map((x) => ({ x, y0: 0, y1: 90, thickness: 1 })),
  };
  const words = [
    word('Name', 5, 8), word('Town', 160, 8), word('Paid', 290, 8),
    word('Ali', 5, 38), word('Khan', 40, 38), word('Lahore', 160, 38),
    word('Sara', 5, 68), word('Kent', 160, 68), word('yes', 290, 68),
  ];
  const t = tableCells(words, rules);
  assert.equal(t.ruled, true);
  assert.deepEqual(t.cells, [['Name', 'Town', 'Paid'], ['Ali Khan', 'Lahore', ''], ['Sara', 'Kent', 'yes']]);
  assert.equal(t.columns.length, 3);
  assert.deepEqual(tableCells([]).cells, []);
});

test('cropPadded and inkIn: a part of a picture with paper round it, and the ink inside a rectangle', async () => {
  const { cropPadded, inkIn } = await import('@rutba/ocr/analyse');
  const p = page(40, 30);
  fill(p, 10, 10, 5, 4);
  const c = cropPadded(p, 8, 8, 10, 8, 2);
  assert.deepEqual([c.width, c.height], [14, 12]);
  assert.equal(c.data[0], 255);
  assert.equal(c.data[(2 + 2) * 14 + 2 + 2], 0, 'the ink where it was, moved by the crop and the margin');
  const b = binarise(p, { method: 'otsu' });
  assert.equal(inkIn(b, 0, 0, 40, 30), 20);
  assert.equal(inkIn(b, 12, 10, 10, 10), 12);
  assert.equal(inkIn(b, 30, 0, 10, 10), 0);
  const words = [word('a', 0, 0), word('b', 100, 0)];
  words[0].confidence = 40;
  words[1].confidence = 90;
  assert.deepEqual(tableCells(words).confidence, [[40, 90]]);
});

test('fitToInk: a word box read a few pixels off its letters is drawn in to them, never onto the next word; a box with no ink is kept', () => {
  // Three words on a binary line 30 px tall: letters 8 px wide, 2 px apart, words 12 px apart.
  const p = page(260, 60);
  const letters = (x, n) => { for (let i = 0; i < n; i++) fill(p, x + i * 10, 15, 8, 30); return x + n * 10 - 2; };
  const a1 = letters(20, 5); // 20..68
  const b0 = a1 + 12, b1 = letters(b0, 4); // 80..118
  const c0 = b1 + 12; letters(c0, 6); // 130..188
  const box = (x, width) => ({ x, y: 15, width, height: 30 });
  const words = [
    { text: 'moved', box: box(24, 48) }, // 4 px to the right of its letters
    { text: 'narrow', box: box(84, 30) }, // 4 px short each side
    { text: 'wide', box: box(c0 - 5, 70) }, // past both sides, the space before it reaching the last word
    { text: 'nothing', box: box(200, 40) }, // paper only
  ];
  fitToInk(p, words);
  assert.deepEqual(words[0].box, box(20, 48));
  assert.deepEqual(words[1].box, box(b0, b1 - b0));
  assert.deepEqual(words[2].box, box(c0, 58));
  assert.deepEqual(words[3].box, box(200, 40));
});

test('fitToInk: an Arabic word\'s clusters (letters joined, a dot over one) are one word, and a side moves no more than half its height', () => {
  const p = page(200, 60);
  // Two joined clusters with a dot above, read with a box missing the left cluster's outer half.
  fill(p, 40, 30, 22, 12); fill(p, 64, 26, 30, 16); fill(p, 70, 18, 4, 4);
  const words = [{ text: 'word', box: { x: 48, y: 18, width: 50, height: 24 } }];
  fitToInk(p, words);
  assert.deepEqual(words[0].box, { x: 40, y: 18, width: 54, height: 24 });
  // A cluster far beyond reach is not taken.
  const q = page(200, 60);
  fill(q, 10, 20, 80, 20);
  const far = [{ text: 'x', box: { x: 60, y: 20, width: 20, height: 20 } }];
  fitToInk(q, far);
  assert.ok(far[0].box.x >= 50, JSON.stringify(far[0].box));
});
