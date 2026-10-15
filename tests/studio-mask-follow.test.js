// Image: a mask follows its layer (docs/STUDIO.md 7.2, I10 with I14), as
// Photoshop links them: a layer moved, scaled, turned, given free-transform
// corners or put in perspective takes its painted mask with it by the same
// map; the mask stays a rectangle while it can and lies on four corners when
// it is turned or bent; an edit that placed the mask itself, a change of the
// canvas or a layer whose place cannot be said leaves it alone; a brush on
// the mask lands through its corners; and moves and crops carry the corners.
import test from 'node:test';
import assert from 'node:assert/strict';
import { emptyImage, pictureLayer, addMask, setLayer, moveLayerBy, reframe, maskCorners, placeMask, carryMask, followMasks, canvasToMask, pictureQuad, mapPoint } from '../packages/studio/src/layers.js';
import { boxCorners, rectCorners } from '../packages/studio/src/warp.js';

const close = (a, b, tol = 1e-3) => Math.abs(a - b) <= tol;
const sameQuad = (a, b, tol = 0.01) => a.every(([x, y], i) => close(x, b[i][0], tol) && close(y, b[i][1], tol));

// A 600 × 400 canvas; a 200 × 200 picture at 100, 50; a mask on it covering
// the picture's left half (100..200 × 50..250), shown nowhere else.
function scene() {
  let doc = emptyImage({ width: 600, height: 400 });
  doc = { ...doc, patches: [...doc.patches, pictureLayer(doc, { url: 'rimg:bitmaps/pic.png', width: 200, height: 200, id: 'pic' })] };
  doc = setLayer(doc, 'pic', { fx: 100 / 600, fy: 50 / 400, fw: 200 / 600 });
  doc = addMask(doc, 'pic', { url: 'rimg:bitmaps/pic-mask.png', outside: 0 });
  doc = setLayer(doc, 'pic', { paintMask: { ...doc.patches.find((p) => p.id === 'pic').paintMask, fx: 100 / 600, fy: 50 / 400, fw: 100 / 600, fh: 200 / 400 } });
  return doc;
}
const BITMAP = { width: 200, height: 200 };
const quadOf = (d, p) => (p.type === 'image' ? pictureQuad(d, p, BITMAP) : null);
const pic = (d) => d.patches.find((p) => p.id === 'pic');

test('scaled from its corner, the layer takes its mask with it, still a rectangle', () => {
  const doc = scene();
  const next = followMasks(doc, setLayer(doc, 'pic', { fw: 300 / 600 }), quadOf);
  const m = pic(next).paintMask;
  assert.equal(m.corners, undefined, 'a scale keeps the mask a rectangle');
  assert.ok(close(m.fx * 600, 100) && close(m.fy * 400, 50) && close(m.fw * 600, 150) && close(m.fh * 400, 300), JSON.stringify(m));
  assert.equal(m.rev, 1, 'the painter is told the mask changed');
  assert.equal(m.url, 'rimg:bitmaps/pic-mask.png');
  assert.equal(m.outside, 0, 'what lies beyond it stays hidden');
});

test('turned, the mask turns with it about the layer\'s middle and lies on four corners', () => {
  const doc = scene();
  const next = followMasks(doc, setLayer(doc, 'pic', { rot: 90 }), quadOf);
  const m = pic(next).paintMask;
  assert.ok(Array.isArray(m.corners), 'corners');
  // The left half of the box, turned a quarter clockwise about (200, 150), is its top half.
  const want = boxCorners({ x: 100, y: 50, w: 100, h: 200 }, 0).map(([x, y]) => [200 - (y - 150), 150 + (x - 200)]);
  assert.ok(sameQuad(maskCorners(next, m), want), JSON.stringify(maskCorners(next, m)));
  assert.ok(close(m.fx * 600, 100, 0.01) && close(m.fy * 400, 50, 0.01) && close(m.fw * 600, 200, 0.01) && close(m.fh * 400, 100, 0.01), 'its bounds kept beside the corners');
});

test('free transform and perspective: the mask goes through the same projective map', () => {
  const doc = scene();
  const corners = [[120, 60], [330, 90], [300, 300], [90, 260]];
  const next = followMasks(doc, setLayer(doc, 'pic', { corners: corners.map(([x, y]) => [x / 600, y / 400]) }), quadOf);
  const m = pic(next).paintMask;
  // The mask's corners are the picture's top left and bottom left, and the middles of its top and bottom edges, through the map.
  const q = maskCorners(next, m);
  assert.ok(sameQuad([q[0], q[3]], [corners[0], corners[3]], 0.01), 'its outer corners are the picture\'s');
  // A projective map does not keep an edge's middle, but it keeps a point of an edge on that edge.
  const onLine = (p, a, b) => Math.abs((b[0] - a[0]) * (p[1] - a[1]) - (b[1] - a[1]) * (p[0] - a[0])) / Math.hypot(b[0] - a[0], b[1] - a[1]) < 0.01;
  assert.ok(onLine(q[1], corners[0], corners[1]) && onLine(q[2], corners[3], corners[2]), 'its inner corners on the picture\'s top and bottom edges');
  // Taken off again (Reset), the mask comes back to where it was.
  const back = followMasks(next, setLayer(next, 'pic', { corners: undefined }), quadOf);
  const mb = pic(back).paintMask;
  assert.equal(mb.corners, undefined, 'a rectangle again');
  assert.ok(close(mb.fx * 600, 100, 0.01) && close(mb.fy * 400, 50, 0.01) && close(mb.fw * 600, 100, 0.01) && close(mb.fh * 400, 200, 0.01), JSON.stringify(mb));
});

test('a long drag in many small steps lands where one big step does', () => {
  let doc = scene();
  for (let i = 1; i <= 200; i++) doc = followMasks(doc, setLayer(doc, 'pic', { rot: i * 0.45, fw: (200 + i * 0.5) / 600 }), quadOf);
  const once = followMasks(scene(), setLayer(scene(), 'pic', { rot: 90, fw: 300 / 600 }), quadOf);
  assert.ok(sameQuad(maskCorners(doc, pic(doc).paintMask), maskCorners(once, pic(once).paintMask), 0.05));
});

test('an edit that placed the mask itself, a new canvas, or a layer with no place to read leaves the mask alone', () => {
  const doc = scene();
  const moved = moveLayerBy(doc, 'pic', 0.1, 0.05);
  assert.equal(followMasks(doc, moved, quadOf), moved, 'a move carries the mask itself');
  assert.ok(close(pic(moved).paintMask.fx, 100 / 600 + 0.1, 1e-4));
  const cropped = reframe(doc, { x: 50, y: 20, width: 400, height: 300 });
  assert.equal(followMasks(doc, cropped, quadOf), cropped, 'a crop places masks itself');
  const blind = followMasks(doc, setLayer(doc, 'pic', { fw: 0.5 }), () => null);
  assert.deepEqual(pic(blind).paintMask, pic(doc).paintMask, 'no corners to go by: the mask stays');
  const plain = { ...doc, patches: doc.patches.map((p) => (p.id === 'pic' ? (({ paintMask, ...rest }) => rest)(p) : p)) };
  const scaled = setLayer(plain, 'pic', { fw: 0.5 });
  assert.equal(followMasks(plain, scaled, quadOf), scaled, 'a layer with no mask is left as it is');
});

test('moves and crops carry a mask on corners', () => {
  const turned = followMasks(scene(), setLayer(scene(), 'pic', { rot: 30 }), quadOf);
  const before = maskCorners(turned, pic(turned).paintMask);
  const moved = moveLayerBy(turned, 'pic', 30 / 600, -20 / 400);
  assert.ok(sameQuad(maskCorners(moved, pic(moved).paintMask), before.map(([x, y]) => [x + 30, y - 20]), 0.05));
  const cropped = reframe(turned, { x: 40, y: 10, width: 500, height: 380 });
  assert.ok(sameQuad(maskCorners(cropped, pic(cropped).paintMask), before.map(([x, y]) => [x - 40, y - 10]), 0.1));
});

test('a brush on a mask on corners lands through them on the mask\'s own bitmap', () => {
  const turned = followMasks(scene(), setLayer(scene(), 'pic', { rot: 90 }), quadOf);
  const m = pic(turned).paintMask;
  const bitmap = { width: 600, height: 400 };
  const map = canvasToMask(turned, m, bitmap);
  const q = maskCorners(turned, m);
  const corners = rectCorners(600, 400);
  q.forEach(([x, y], i) => {
    const p = mapPoint(map, x, y);
    assert.ok(close(p.x, corners[i][0], 0.01) && close(p.y, corners[i][1], 0.01), `corner ${i}`);
  });
  const flat = canvasToMask(scene(), pic(scene()).paintMask, bitmap);
  const p = mapPoint(flat, 150, 150);
  assert.ok(close(p.x, 300, 0.01) && close(p.y, 200, 0.01), 'an unturned mask maps as it always did');
});

test('placeMask keeps a rectangle a rectangle and anything else on corners', () => {
  const doc = scene();
  const m = pic(doc).paintMask;
  const rect = placeMask(doc, m, rectCorners(120, 80, 30, 40));
  assert.equal(rect.corners, undefined);
  assert.ok(close(rect.fx * 600, 30) && close(rect.fw * 600, 120) && close(rect.fy * 400, 40) && close(rect.fh * 400, 80));
  const bent = placeMask(doc, m, [[30, 40], [150, 50], [150, 120], [30, 120]]);
  assert.equal(bent.corners.length, 4);
  assert.ok(sameQuad(maskCorners(doc, bent), [[30, 40], [150, 50], [150, 120], [30, 120]]));
  assert.equal(carryMask(doc, m, null, rectCorners(1, 1)), m, 'nothing to go by: unchanged');
});
