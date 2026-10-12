// Image's pixel operations, each over a small picture made here, against
// the values it should give (golden values, a level or two of tolerance for
// rounding): levels, curves, white balance, exposure, highlights and
// shadows, auto enhance, sharpen, noise reduction, vignette and red-eye.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  applyTone, levelsTable, curveTable, toneTables, invertPixels, sharpen, denoise, vignette, redEye, autoEnhance, applyPixels, histogram, isNeutral, CURVE_PRESETS,
} from '@rutba/imaging/pixels';

/** A w × h picture, each pixel from fn(x, y) → [r, g, b]. */
function picture(w, h, fn) {
  const data = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const [r, g, b] = fn(x, y);
    data.set([r, g, b, 255], (y * w + x) * 4);
  }
  return { data, width: w, height: h };
}
const flat = (w, h, rgb) => picture(w, h, () => rgb);
const px = (img, x, y) => [...img.data.subarray((y * img.width + x) * 4, (y * img.width + x) * 4 + 3)];
const close = (a, b, tol = 1) => a.length === b.length && a.every((v, i) => Math.abs(v - b[i]) <= tol);

test('levels: input points stretch the tones, gamma moves the middle, output points hold the ends', () => {
  const t = levelsTable({ inBlack: 50, inWhite: 200 });
  assert.equal(t[50], 0);
  assert.equal(t[200], 255);
  assert.equal(t[125], 128);
  assert.equal(t[20], 0);
  assert.equal(t[230], 255);
  const g = levelsTable({ gamma: 2 });
  assert.equal(g[0], 0);
  assert.equal(g[255], 255);
  assert.ok(Math.abs(g[64] - 128) <= 1, `gamma 2 lifts a quarter to the middle: ${g[64]}`);
  const o = levelsTable({ outBlack: 20, outWhite: 235 });
  assert.equal(o[0], 20);
  assert.equal(o[255], 235);
  // Golden: a 2 × 2 picture through levels 64..192.
  const img = picture(2, 2, (x, y) => [[64, 96, 128, 192][y * 2 + x], 0, 255]);
  applyTone(img, { levels: { inBlack: 64, inWhite: 192 } });
  assert.deepEqual([px(img, 0, 0)[0], px(img, 1, 0)[0], px(img, 0, 1)[0], px(img, 1, 1)[0]], [0, 64, 128, 255]);
});

test('curves: a straight line changes nothing; a point bends the tones smoothly without overshooting', () => {
  const line = curveTable(CURVE_PRESETS.linear);
  for (let i = 0; i < 256; i++) assert.equal(line[i], i);
  const lift = curveTable([[0, 0], [128, 180], [255, 255]]);
  assert.equal(lift[128], 180);
  assert.equal(lift[0], 0);
  assert.equal(lift[255], 255);
  for (let i = 1; i < 256; i++) assert.ok(lift[i] >= lift[i - 1], `rising at ${i}`);
  assert.ok(lift[64] > 64 && lift[64] < 180);
  // An S curve: darker below the middle, lighter above, monotone.
  const s = curveTable(CURVE_PRESETS.contrast);
  assert.ok(s[64] < 64 && s[192] > 192);
  for (let i = 1; i < 256; i++) assert.ok(s[i] >= s[i - 1]);
  // A curve on one channel only.
  const img = flat(1, 1, [100, 100, 100]);
  applyTone(img, { curve: { r: [[0, 0], [100, 150], [255, 255]] } });
  assert.deepEqual(px(img, 0, 0), [150, 100, 100]);
});

test('white balance: warmer raises red and lowers blue; more magenta lowers green', () => {
  const warm = flat(1, 1, [128, 128, 128]);
  applyTone(warm, { temperature: 50 });
  const [r, g, b] = px(warm, 0, 0);
  assert.ok(r > 140 && b < 116 && Math.abs(g - 128) <= 1, `${r} ${g} ${b}`);
  const cool = flat(1, 1, [128, 128, 128]);
  applyTone(cool, { temperature: -50 });
  assert.ok(px(cool, 0, 0)[2] > 140 && px(cool, 0, 0)[0] < 116);
  const magenta = flat(1, 1, [128, 128, 128]);
  applyTone(magenta, { tint: 60 });
  assert.ok(px(magenta, 0, 0)[1] < 116, String(px(magenta, 0, 0)));
});

test('exposure: one stop doubles the light, black stays black, a stop down halves it', () => {
  const t = toneTables({ exposure: 1 }).first[0];
  assert.equal(t[0], 0);
  // 128 is about 0.218 in linear light; doubled 0.436; back on the screen's curve about 175.
  assert.ok(Math.abs(t[128] - 175) <= 2, String(t[128]));
  assert.equal(t[255], 255);
  const down = toneTables({ exposure: -1 }).first[0];
  assert.ok(Math.abs(down[128] - 94) <= 2, String(down[128]));
});

test('highlights and shadows: shadows lift the dark third and leave black and white alone; highlights bring brights down', () => {
  const img = picture(3, 1, (x) => [[0, 60, 255][x], [0, 60, 255][x], [0, 60, 255][x]]);
  applyTone(img, { shadows: 80 });
  assert.equal(px(img, 0, 0)[0], 0);
  assert.ok(px(img, 1, 0)[0] > 80, String(px(img, 1, 0)));
  assert.equal(px(img, 2, 0)[0], 255);
  const bright = picture(2, 1, (x) => [[200, 128][x], [200, 128][x], [200, 128][x]]);
  applyTone(bright, { highlights: -80 });
  assert.ok(px(bright, 0, 0)[0] < 185, String(px(bright, 0, 0)));
  // The colour of a pixel is kept: its channels move together.
  const colour = flat(1, 1, [60, 30, 15]);
  applyTone(colour, { shadows: 60 });
  const [r, g, b] = px(colour, 0, 0);
  assert.ok(Math.abs(r / g - 2) < 0.1 && Math.abs(g / b - 2) < 0.15, `${r} ${g} ${b}`);
});

test('invert: each channel is 255 less itself', () => {
  const img = flat(1, 1, [10, 200, 255]);
  invertPixels(img);
  assert.deepEqual(px(img, 0, 0), [245, 55, 0]);
});

test('sharpen: an edge gains contrast, a flat area is left exactly as it was', () => {
  const edge = () => picture(8, 4, (x) => (x < 4 ? [80, 80, 80] : [160, 160, 160]));
  const img = edge();
  sharpen(img, 60, { radius: 1 });
  assert.ok(px(img, 3, 1)[0] < 80, `the dark side darker at the edge: ${px(img, 3, 1)}`);
  assert.ok(px(img, 4, 1)[0] > 160, `the light side lighter: ${px(img, 4, 1)}`);
  assert.deepEqual(px(img, 0, 1), [80, 80, 80]);
  assert.deepEqual(px(img, 7, 1), [160, 160, 160]);
  const none = edge();
  sharpen(none, 0);
  assert.deepEqual([...none.data], [...edge().data]);
});

test('noise reduction: grain in a flat patch is smoothed, a hard edge is kept', () => {
  let seed = 7;
  const rand = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
  const img = picture(16, 8, (x) => {
    const base = x < 8 ? 60 : 200;
    const n = Math.round((rand() - 0.5) * 24);
    return [base + n, base + n, base + n];
  });
  const variance = (im, x0, x1) => {
    const v = [];
    for (let y = 1; y < im.height - 1; y++) for (let x = x0; x < x1; x++) v.push(px(im, x, y)[0]);
    const m = v.reduce((a, b) => a + b, 0) / v.length;
    return v.reduce((a, b) => a + (b - m) ** 2, 0) / v.length;
  };
  const before = variance(img, 1, 6);
  denoise(img, 70);
  const after = variance(img, 1, 6);
  assert.ok(after < before * 0.6, `grain ${before.toFixed(1)} → ${after.toFixed(1)}`);
  // The edge: the two sides still about 140 apart.
  const left = px(img, 6, 4)[0];
  const right = px(img, 9, 4)[0];
  assert.ok(right - left > 110, `edge ${left} | ${right}`);
});

test('vignette: the corners darken, the middle is untouched; below zero they lighten', () => {
  const img = flat(41, 41, [200, 200, 200]);
  vignette(img, 80);
  assert.deepEqual(px(img, 20, 20), [200, 200, 200]);
  assert.ok(px(img, 0, 0)[0] < 90, String(px(img, 0, 0)));
  assert.ok(px(img, 20, 0)[0] > px(img, 0, 0)[0], 'an edge less than a corner');
  const light = flat(41, 41, [100, 100, 100]);
  vignette(light, -80);
  assert.ok(px(light, 0, 0)[0] > 200);
  assert.deepEqual(px(light, 20, 20), [100, 100, 100]);
});

test('red-eye: a red pupil inside the spot goes dark, a red pixel outside it and a grey one inside are left', () => {
  const img = picture(20, 10, (x) => (x >= 3 && x <= 6 ? [220, 40, 40] : x === 15 ? [220, 40, 40] : [120, 120, 120]));
  redEye(img, { x: 5, y: 5, radius: 4 });
  const pupil = px(img, 5, 5);
  assert.ok(pupil[0] <= 45, `pupil ${pupil}`);
  assert.deepEqual(pupil.slice(1), [40, 40]);
  assert.deepEqual(px(img, 15, 5), [220, 40, 40]);
  assert.deepEqual(px(img, 8, 5), [120, 120, 120]);
});

test('auto enhance: a flat, blue-tinged picture gets its range stretched and its cast warmed out', () => {
  const dull = picture(32, 16, (x, y) => {
    const v = 100 + Math.round(((x + y) / 46) * 60);
    return [v - 12, v, v + 14];
  });
  const auto = autoEnhance(dull);
  assert.ok(auto.levels.inBlack >= 85 && auto.levels.inBlack <= 105, JSON.stringify(auto));
  assert.ok(auto.levels.inWhite >= 150 && auto.levels.inWhite <= 170, JSON.stringify(auto));
  assert.ok(auto.temperature > 0, `warmed: ${auto.temperature}`);
  const after = applyPixels(dull, auto);
  const h = histogram(after).l;
  let lo = 0;
  while (!h[lo]) lo++;
  let hi = 255;
  while (!h[hi]) hi--;
  assert.ok(lo < 25 && hi > 230, `range ${lo}..${hi}`);
  // A picture that already fills the range and is grey is barely touched.
  const good = picture(32, 16, (x) => [x * 8, x * 8, x * 8]);
  const same = autoEnhance(good);
  assert.equal(same.levels.inBlack, 0);
  assert.equal(same.temperature, 0);
  assert.equal(same.tint, 0);
});

test('neutral settings change nothing, and all of it runs in order on one picture', () => {
  assert.ok(isNeutral({}));
  assert.ok(isNeutral({ levels: { inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 }, curve: { rgb: [[0, 0], [255, 255]] } }));
  assert.ok(!isNeutral({ vignette: 10 }));
  const img = picture(12, 12, (x, y) => [x * 20, y * 20, 128]);
  const copy = [...img.data];
  applyPixels(img, {});
  assert.deepEqual([...img.data], copy);
  applyPixels(img, { exposure: 0.5, shadows: 20, sharpen: 30, denoise: 20, vignette: 30 }, { spots: [{ x: 6, y: 6, radius: 2 }] });
  assert.ok(!close([...img.data], copy, 0));
  // Alpha is never touched.
  for (let i = 3; i < img.data.length; i += 4) assert.equal(img.data[i], 255);
});
