// Tick boxes (checkboxes.js), proved on pages drawn here with known
// answers: empty, ticked, crossed and filled boxes found and read, and the
// squares that are not boxes (a bullet, a round letter, a letter inside a
// word, a Chinese character) left out.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { findCheckboxes } from '@rutba/ocr/checkboxes';

const page = (width = 600, height = 200) => ({ width, height, data: new Uint8Array(width * height).fill(255) });
const fill = (p, x, y, w, h) => {
  for (let yy = y; yy < y + h; yy++) for (let xx = x; xx < x + w; xx++) p.data[yy * p.width + xx] = 0;
};
/** A square outline `side` across with sides `stroke` thick. */
const box = (p, x, y, side, stroke = 2) => {
  fill(p, x, y, side, stroke); fill(p, x, y + side - stroke, side, stroke);
  fill(p, x, y, stroke, side); fill(p, x + side - stroke, y, stroke, side);
};
const line = (p, x0, y0, x1, y1, thick = 3) => {
  const n = Math.max(Math.abs(x1 - x0), Math.abs(y1 - y0));
  for (let i = 0; i <= n; i++) fill(p, Math.round(x0 + ((x1 - x0) * i) / n), Math.round(y0 + ((y1 - y0) * i) / n), thick, thick);
};
/** Letter-like strokes, so the page has a text height to measure. */
const letters = (p, x, y, n = 12) => { for (let i = 0; i < n; i++) fill(p, x + i * 14, y, 8, 20); };

test('findCheckboxes: empty, ticked, crossed and filled boxes, each found where it is and read', () => {
  const p = page();
  letters(p, 20, 150);
  box(p, 40, 40, 26);
  box(p, 120, 40, 26); line(p, 125, 52, 132, 60); line(p, 132, 60, 146, 44);
  box(p, 200, 40, 26); line(p, 205, 45, 220, 60); line(p, 220, 45, 205, 60);
  box(p, 280, 40, 26); fill(p, 287, 47, 12, 12);
  const found = findCheckboxes(p, { textHeight: 20 });
  assert.equal(found.length, 4, JSON.stringify(found));
  assert.deepEqual(found.map((f) => f.checked), [false, true, true, true]);
  assert.deepEqual(found[0].box, { x: 40, y: 40, width: 26, height: 26 });
  assert.equal(found[0].fill, 0);
  assert.ok(found[3].fill > found[1].fill, 'a filled box holds more ink than a tick');
});

test('findCheckboxes: a solid bullet, a ring, a square inside a word and a box too small for the others are not boxes', () => {
  const p = page();
  letters(p, 20, 150);
  box(p, 40, 40, 26);
  box(p, 100, 40, 26);
  box(p, 160, 40, 26);
  fill(p, 230, 40, 26, 26); // a bullet
  for (let a = 0; a < 360; a += 2) fill(p, Math.round(300 + 12 * Math.cos((a * Math.PI) / 180)), Math.round(53 + 12 * Math.sin((a * Math.PI) / 180)), 2, 2); // a ring
  fill(p, 340, 42, 6, 22); box(p, 348, 40, 26); fill(p, 376, 42, 6, 22); // a square with letters touching it on both sides
  box(p, 420, 50, 12); // a small square among large boxes
  const found = findCheckboxes(p, { textHeight: 20 });
  assert.deepEqual(found.map((f) => f.box.x), [40, 100, 160]);
});

test('findCheckboxes: a square under a Chinese character the reading found is the character, not a box', () => {
  const p = page();
  letters(p, 20, 150);
  box(p, 40, 40, 26);
  box(p, 120, 40, 26); fill(p, 122, 52, 22, 2); // 日, drawn as a box with a bar
  const words = [{ text: '日期', box: { x: 118, y: 38, width: 60, height: 30 }, confidence: 92 }];
  assert.equal(findCheckboxes(p, { textHeight: 20 }).length, 2);
  const found = findCheckboxes(p, { textHeight: 20, words });
  assert.deepEqual(found.map((f) => [f.box.x, f.checked]), [[40, false]]);
  // An empty box the reading took for 口, which it looks exactly like, is still a box.
  const misread = [{ text: '口', box: { x: 38, y: 38, width: 30, height: 30 }, confidence: 90 }, ...words];
  assert.deepEqual(findCheckboxes(p, { textHeight: 20, words: misread }).map((f) => f.box.x), [40]);
});

test('findCheckboxes: a scanned look (grey paper, a one-pixel gap in a side) still reads', () => {
  const p = page();
  p.data.fill(235);
  letters(p, 20, 150);
  box(p, 40, 40, 28, 3);
  p.data[40 * p.width + 52] = 235; // a dropout in the top side
  box(p, 120, 40, 28, 3); line(p, 126, 54, 133, 62); line(p, 133, 62, 148, 46);
  const found = findCheckboxes(p);
  assert.deepEqual(found.map((f) => [f.box.x, f.checked]), [[40, false], [120, true]]);
});
