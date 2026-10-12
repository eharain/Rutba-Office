// The Image app's painting tools (docs/STUDIO.md 7.2, I12 and I13): the
// brush's fall-off and dab, a stroke laid out in evenly spaced dabs with the
// pen's pressure, the bucket's fill inside a region, the gradient, the clone
// stamp's offset, and the healing brush's patch blend.
import test from 'node:test';
import assert from 'node:assert/strict';
import { brushAlpha, dabSprite, dabSpacing, DabWalker, cloneOffset, cloneSourceAt, fillRegion, gradientAt, fillGradient, healPatch, hexColour } from '../packages/imaging/src/paint.js';
import { magicWand } from '../packages/imaging/src/select.js';

const px = (d, w, x, y) => [...d.slice((y * w + x) * 4, (y * w + x) * 4 + 4)];

test('I12: the brush lays all its paint inside its hardness and falls smoothly to none at its edge', () => {
  assert.equal(brushAlpha(0, 0.5), 1);
  assert.equal(brushAlpha(0.5, 0.5), 1);
  assert.equal(brushAlpha(1, 0.5), 0);
  assert.equal(brushAlpha(1.2, 0.5), 0);
  assert.ok(Math.abs(brushAlpha(0.75, 0.5) - 0.5) < 1e-9, 'half way down its edge, half');
  assert.ok(brushAlpha(0.6, 0.5) > brushAlpha(0.7, 0.5), 'falling as it goes out');
  assert.ok(brushAlpha(0.3, 0) < 1, 'a soft brush has no hard middle');
});

test('I12: a dab is a round white sprite, solid in the middle, clear at its corners, the size asked', () => {
  const { size, data } = dabSprite(20, 0.8);
  assert.equal(size, 22);
  const alpha = (x, y) => data[(y * size + x) * 4 + 3];
  assert.equal(alpha(11, 11), 255);
  assert.equal(alpha(0, 0), 0);
  assert.equal(data[(11 * size + 11) * 4], 255, 'white, so a colour is laid by tinting it');
  const covered = [...data].filter((_, i) => i % 4 === 3).reduce((n, a) => n + a / 255, 0);
  assert.ok(covered > Math.PI * 64 && covered < Math.PI * 100, `about a disc of radius 8 to 10: ${covered}`);
  const tiny = dabSprite(1, 1);
  assert.equal(tiny.size, 3);
});

test('I12: a stroke is dabs an eighth of the brush apart along the pointer\'s path, the pen\'s pressure eased between points', () => {
  assert.equal(dabSpacing(40), 4.8);
  assert.equal(dabSpacing(2), 1, 'never under a pixel');
  const walk = new DabWalker(5);
  assert.deepEqual(walk.add({ x: 0, y: 0, pressure: 0.2 }), [{ x: 0, y: 0, pressure: 0.2 }]);
  const run = walk.add({ x: 12, y: 0, pressure: 0.8 });
  assert.deepEqual(run.map((d) => d.x), [5, 10]);
  assert.ok(Math.abs(run[0].pressure - 0.45) < 1e-9);
  const next = walk.add({ x: 20, y: 0, pressure: 0.8 });
  assert.deepEqual(next.map((d) => d.x), [15, 20], 'the spacing carries across points');
  assert.deepEqual(walk.add({ x: 21, y: 0 }), [], 'a step shorter than the spacing lays nothing');
  const mouse = new DabWalker(3).add({ x: 1, y: 2 });
  assert.equal(mouse[0].pressure, 1, 'a mouse presses fully');
});

test('I13: the clone stamp copies from the Alt+clicked point, kept at the same offset through the stroke', () => {
  const offset = cloneOffset({ x: 100, y: 40 }, { x: 160, y: 90 });
  assert.deepEqual(offset, { dx: -60, dy: -50 });
  assert.deepEqual(cloneSourceAt({ x: 160, y: 90 }, offset), { x: 100, y: 40 }, 'the first dab copies the source itself');
  assert.deepEqual(cloneSourceAt({ x: 200, y: 120 }, offset), { x: 140, y: 70 }, 'and the next the same distance away');
});

test('I12: the bucket fills the region the wand finds at its tolerance, and only there', () => {
  const W = 8;
  const H = 4;
  const rgba = new Uint8ClampedArray(W * H * 4);
  for (let p = 0; p < W * H; p++) rgba.set(p % W < 4 ? [250, 250, 250, 255] : p % W === 4 ? [0, 0, 0, 255] : [244, 246, 250, 255], p * 4);
  const tight = magicWand(rgba, W, H, 0, 0, { tolerance: 4 });
  const filled = fillRegion(rgba, W, H, tight.data, hexColour('#ff0000'));
  assert.deepEqual(px(filled, W, 1, 1), [255, 0, 0, 255]);
  assert.deepEqual(px(filled, W, 4, 1), [0, 0, 0, 255], 'the black line holds it');
  const wide = magicWand(rgba, W, H, 6, 0, { tolerance: 10, contiguous: false });
  const all = fillRegion(rgba, W, H, wide.data, [0, 0, 255, 255], { opacity: 0.5 });
  assert.deepEqual(px(all, W, 0, 0), [125, 125, 252, 255], 'half the blue over the white, both sides of the line');
  assert.deepEqual(px(all, W, 4, 0), [0, 0, 0, 255]);
  // On clear pixels the colour is laid at its own strength.
  const clear = fillRegion(new Uint8ClampedArray(4), 1, 1, null, [10, 20, 30, 255], { opacity: 0.4 });
  assert.deepEqual([...clear], [10, 20, 30, 102]);
  assert.deepEqual(hexColour('#2f7de1'), [47, 125, 225, 255]);
});

test('I12: a gradient runs from its start to its end along the drag, or out from the start for radial', () => {
  const s = { x: 10, y: 0 };
  const e = { x: 30, y: 0 };
  assert.equal(gradientAt('linear', s, e, 0, 5), 0);
  assert.equal(gradientAt('linear', s, e, 20, 99), 0.5, 'across the drag it does not change');
  assert.equal(gradientAt('linear', s, e, 40, 0), 1);
  assert.equal(gradientAt('radial', s, e, 10, 10), 0.5);
  assert.equal(gradientAt('reflected', s, e, 0, 0), 0.5);
  const W = 11;
  const out = fillGradient(new Uint8ClampedArray(W * 4), W, 1, { kind: 'linear', start: { x: 0.5, y: 0 }, end: { x: 10.5, y: 0 }, from: [0, 0, 0, 255], to: [200, 100, 0, 255] });
  assert.deepEqual(px(out, W, 0, 0), [0, 0, 0, 255]);
  assert.deepEqual(px(out, W, 5, 0), [100, 50, 0, 255]);
  assert.deepEqual(px(out, W, 10, 0), [200, 100, 0, 255]);
  // To clear: the colour stays the colour and only fades.
  const fade = fillGradient(new Uint8ClampedArray(W * 4), W, 1, { start: { x: 0.5, y: 0 }, end: { x: 10.5, y: 0 }, from: [200, 0, 0, 255], to: [0, 0, 0, 0] });
  assert.deepEqual(px(fade, W, 5, 0), [200, 0, 0, 128]);
  // Only inside a region.
  const region = new Uint8Array(W).map((_, i) => (i < 3 ? 255 : 0));
  const cut = fillGradient(new Uint8ClampedArray(W * 4), W, 1, { start: { x: 0, y: 0 }, end: { x: 10, y: 0 }, from: [9, 9, 9, 255], to: [9, 9, 9, 255], region });
  assert.equal(cut[2 * 4 + 3], 255);
  assert.equal(cut[3 * 4 + 3], 0);
});

test('I13: the healing brush keeps the texture it copies and takes the light and colour round the place it heals', () => {
  const W = 40;
  const H = 40;
  // Where it is laid: a warm, darker skin with a dark blemish in the middle.
  const dst = new Uint8ClampedArray(W * H * 4);
  // What is copied: a cooler, lighter patch with a fine checker texture.
  const src = new Uint8ClampedArray(W * H * 4);
  const coverage = new Uint8Array(W * H);
  for (let y = 0; y < H; y++) {
    for (let x = 0; x < W; x++) {
      const i = (y * W + x) * 4;
      const blemish = Math.hypot(x - 20, y - 20) < 4;
      dst.set(blemish ? [40, 20, 20, 255] : [180, 120, 100, 255], i);
      const tex = (x + y) % 2 ? 12 : -12;
      src.set([120 + tex, 140 + tex, 170 + tex, 255], i);
      coverage[y * W + x] = Math.hypot(x - 20, y - 20) < 7 ? 255 : 0;
    }
  }
  const out = healPatch(src, dst, W, H, { radius: 10, coverage });
  const at = (x, y) => px(out, W, x, y);
  // At the blemish: the colour of the skin round it, give or take the texture.
  const mid = [0, 1, 2].map((c) => (at(20, 20)[c] + at(21, 20)[c]) / 2);
  assert.ok(Math.abs(mid[0] - 180) < 22 && Math.abs(mid[1] - 120) < 22 && Math.abs(mid[2] - 100) < 22, `healed to the skin round it: ${mid}`);
  assert.ok(Math.abs(at(20, 20)[0] - at(21, 20)[0]) > 12, 'the copied texture comes with it');
  assert.deepEqual(at(2, 2), [180, 120, 100, 255], 'outside the brush nothing changes');
  // Without a brush shape, a flat source over a flat place is the place.
  const flatSrc = new Uint8ClampedArray(16 * 4).fill(90);
  const flatDst = new Uint8ClampedArray(16 * 4).fill(200);
  const flat = healPatch(flatSrc, flatDst, 4, 4, { radius: 2 });
  assert.ok([...flat].every((v) => Math.abs(v - 200) <= 1));
});
