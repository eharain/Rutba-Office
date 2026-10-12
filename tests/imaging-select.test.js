// The Image app's selections and painted masks (docs/STUDIO.md 7.2, I10 and
// I11): rectangles, ellipses and lassos rasterised with smooth edges, the
// magic wand's tolerance and contiguity, select by colour, combining,
// inverting, feathering, growing and shrinking, the outline the marching ants
// follow, and a mask laid over a layer's pixels.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  emptySelection, selectAll, rectSelection, ellipseSelection, polygonSelection, magicWand, colourSelection,
  combineSelection, invertSelection, isEmptySelection, selectionBounds, featherSelection, growSelection,
  shrinkSelection, selectionOutline, applyMask, maskCoverage, blurPlane, distanceSquared,
} from '../packages/imaging/src/select.js';

const at = (s, x, y) => s.data[y * s.width + x];
const count = (s, min = 128) => s.data.reduce((n, v) => n + (v >= min ? 1 : 0), 0);

/** A picture of flat colours: `paint(x, y)` answers [r, g, b, a]. */
function picture(w, h, paint) {
  const d = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) d.set(paint(x, y), (y * w + x) * 4);
  return d;
}

test('I11: a rectangle selects whole pixels inside it and part of a pixel its edge crosses', () => {
  const s = rectSelection(10, 8, { x: 2, y: 1, width: 4, height: 3 });
  assert.equal(count(s, 255), 12);
  assert.equal(at(s, 2, 1), 255);
  assert.equal(at(s, 5, 3), 255);
  assert.equal(at(s, 6, 3), 0);
  assert.equal(at(s, 1, 1), 0);
  const half = rectSelection(10, 8, { x: 2.5, y: 1, width: 2, height: 1 });
  assert.equal(at(half, 2, 1), 128, 'half a pixel is half selected');
  assert.equal(at(half, 3, 1), 255);
  assert.equal(at(half, 4, 1), 128);
  // Dragged up and to the left, the same rectangle.
  assert.deepEqual([...rectSelection(10, 8, { x: 6, y: 4, width: -4, height: -3 }).data], [...s.data]);
  assert.deepEqual(selectionBounds(s), { x: 2, y: 1, width: 4, height: 3 });
});

test('I11: an ellipse covers about pi r squared, its rim smoothed and its middle whole', () => {
  const s = ellipseSelection(64, 64, { x: 12, y: 12, width: 40, height: 40 });
  const area = s.data.reduce((n, v) => n + v / 255, 0);
  assert.ok(Math.abs(area - Math.PI * 400) < 8, `area ${area}`);
  assert.equal(at(s, 32, 32), 255);
  assert.equal(at(s, 13, 13), 0, 'the corner of its box is outside');
  const rim = [...s.data].filter((v) => v > 0 && v < 255).length;
  assert.ok(rim > 40, 'the rim is partly selected, not stepped');
});

test('I11: a lasso fills the shape its points close, even-odd, with a smooth edge', () => {
  const tri = polygonSelection(40, 40, [[0, 0], [40, 0], [0, 40]]);
  const area = tri.data.reduce((n, v) => n + v / 255, 0);
  assert.ok(Math.abs(area - 800) < 6, `a triangle of half the square: ${area}`);
  assert.equal(at(tri, 5, 5), 255);
  assert.equal(at(tri, 35, 35), 0);
  // A square inside a square, as one path round both: the inner one is a hole.
  const ring = polygonSelection(30, 30, [[2, 2], [28, 2], [28, 28], [2, 28], [2, 2], [10, 10], [20, 10], [20, 20], [10, 20], [10, 10]]);
  assert.equal(at(ring, 5, 15), 255);
  assert.equal(at(ring, 15, 15), 0);
  assert.equal(count(polygonSelection(10, 10, [[1, 1], [5, 5]])), 0, 'two points close nothing');
});

test('I11: the magic wand takes the joined pixels within its tolerance; not contiguous, every such pixel', () => {
  // Left half dark blue, a light-blue column at 6, the right half dark blue again, a red dot.
  const W = 12;
  const H = 6;
  const blue = [20, 40, 200, 255];
  const rgba = picture(W, H, (x, y) => (x === 6 ? [120, 160, 250, 255] : x === 9 && y === 2 ? [230, 30, 30, 255] : x === 2 && y === 4 ? [30, 48, 210, 255] : blue));
  const tight = magicWand(rgba, W, H, 1, 1, { tolerance: 5, contiguous: true });
  assert.equal(count(tight), 6 * 6 - 1, 'the left half, less the pixel 10 away');
  const loose = magicWand(rgba, W, H, 1, 1, { tolerance: 12, contiguous: true });
  assert.equal(count(loose), 36, 'with more tolerance the near-blue pixel joins');
  assert.equal(at(loose, 7, 0), 0, 'the light column stops it');
  const everywhere = magicWand(rgba, W, H, 1, 1, { tolerance: 12, contiguous: false });
  assert.equal(count(everywhere), 36 + 5 * 6 - 1, 'not contiguous: both halves, not the column or the dot');
  assert.equal(at(everywhere, 9, 2), 0);
  assert.equal(count(magicWand(rgba, W, H, 99, 1)), 0, 'a click off the picture selects nothing');
  // Select by colour is the same test against a colour named.
  assert.deepEqual([...colourSelection(rgba, W, H, [230, 30, 30], { tolerance: 0 }).data].filter(Boolean).length, 1);
});

test('I11: selections add, subtract, intersect and invert; select all and nothing', () => {
  const a = rectSelection(10, 10, { x: 0, y: 0, width: 6, height: 10 });
  const b = rectSelection(10, 10, { x: 4, y: 0, width: 6, height: 10 });
  assert.equal(count(combineSelection(a, b, 'add')), 100);
  assert.equal(count(combineSelection(a, b, 'subtract')), 40);
  assert.equal(count(combineSelection(a, b, 'intersect')), 20);
  assert.equal(combineSelection(a, b, 'replace'), b);
  assert.equal(combineSelection(null, b, 'add'), b);
  assert.equal(count(invertSelection(a)), 40);
  assert.equal(count(selectAll(10, 10)), 100);
  assert.ok(isEmptySelection(emptySelection(10, 10)));
  assert.ok(!isEmptySelection(a));
  assert.equal(selectionBounds(emptySelection(4, 4)), null);
});

test('I11: feather softens the edge over about its radius, keeps the middle whole and the area the same', () => {
  const s = rectSelection(80, 40, { x: 20, y: 0, width: 40, height: 40 });
  const f = featherSelection(s, 8);
  assert.equal(at(f, 40, 20), 255, 'the middle stays selected');
  assert.equal(at(f, 2, 20), 0, 'far outside stays out');
  const edge = at(f, 20, 20);
  assert.ok(edge > 90 && edge < 170, `the edge is about half: ${edge}`);
  assert.ok(at(f, 16, 20) > 0 && at(f, 16, 20) < edge, 'outside, near: a little');
  assert.ok(at(f, 24, 20) > edge && at(f, 24, 20) < 255, 'inside, near: most');
  const area = (x) => x.data.reduce((n, v) => n + v, 0);
  assert.ok(Math.abs(area(f) - area(s)) / area(s) < 0.02, 'feathering moves the edge, it does not grow the selection');
  assert.equal(featherSelection(s, 0), s);
  // The blur behind it holds a flat plane flat.
  const flat = blurPlane(new Float32Array(100).fill(77), 10, 10, 6);
  assert.ok(flat.every((v) => Math.abs(v - 77) < 1e-3));
});

test('I11: grow and shrink move the edge by whole pixels, the corners rounded', () => {
  const s = rectSelection(40, 40, { x: 10, y: 10, width: 20, height: 20 });
  const g = growSelection(s, 3);
  assert.deepEqual(selectionBounds(g, 127), { x: 7, y: 7, width: 26, height: 26 });
  assert.equal(at(g, 7, 7), 0, 'the grown corner is round, not square');
  const k = shrinkSelection(s, 3);
  assert.deepEqual(selectionBounds(k, 127), { x: 13, y: 13, width: 14, height: 14 });
  assert.equal(growSelection(s, 0), s);
  const d = distanceSquared((i) => i === 0, 5, 1);
  assert.deepEqual([...d], [0, 1, 4, 9, 16], 'the distance under it is exact');
});

test('I11: the marching ants run round the edge of what is selected', () => {
  const s = rectSelection(10, 10, { x: 2, y: 3, width: 4, height: 2 });
  const segs = selectionOutline(s);
  assert.equal(segs.length / 4, 4, 'one run for each side');
  const len = (i) => Math.hypot(segs[i + 2] - segs[i], segs[i + 3] - segs[i + 1]);
  let total = 0;
  for (let i = 0; i < segs.length; i += 4) total += len(i);
  assert.equal(total, 12, 'the perimeter, 4 + 2 + 4 + 2');
  assert.deepEqual(segs.slice(0, 4), [2, 3, 6, 3]);
  assert.equal(selectionOutline(emptySelection(5, 5)).length, 0);
});

test('I10: a mask multiplies a layer\'s alpha; inverted, it keeps the rest; beyond its edge it shows or hides', () => {
  const rgba = new Uint8ClampedArray([10, 20, 30, 255, 10, 20, 30, 200, 10, 20, 30, 255]);
  const mask = new Uint8Array([255, 128, 0]);
  const out = applyMask(rgba, mask);
  assert.deepEqual([out[3], out[7], out[11]], [255, 100, 0]);
  assert.deepEqual([out[0], out[1], out[2]], [10, 20, 30], 'the colour is untouched');
  const inv = applyMask(rgba, mask, { invert: true });
  assert.deepEqual([inv[3], inv[7], inv[11]], [0, 100, 255]);
  assert.equal(maskCoverage(64), 64 / 255);
  assert.equal(maskCoverage(64, { invert: true }), 1 - 64 / 255);
  assert.equal(maskCoverage(0, { inside: false, outside: 1 }), 1, 'a reveal-all mask shows what lies past it');
  assert.equal(maskCoverage(255, { inside: false, outside: 0 }), 0, 'a hide-all mask hides it');
  assert.equal(maskCoverage(0, { off: true }), 1, 'a mask turned off lets everything through');
});
