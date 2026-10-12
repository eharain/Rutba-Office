// The Image app's free transform and perspective correction (docs/STUDIO.md
// 7.2, I14): the projective map through four corners, its inverse, the
// rectangle a skewed one is straightened to, the corner handles' scale, skew,
// distort and perspective, and a picture warped onto its corners.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  homography, applyHomography, invertHomography, rectCorners, quadBounds, isConvexQuad, rectifyTarget, straighten,
  boxCorners, dragCorner, warpPixels, solve,
} from '../packages/studio/src/warp.js';

const close = (a, b, tol = 1e-6) => Math.abs(a - b) <= tol;
const samePoint = (p, q, tol = 1e-6) => close(p[0], q[0], tol) && close(p[1], q[1], tol);

test('I14: a homography takes four points to four points, and its inverse brings them back', () => {
  const src = rectCorners(100, 50);
  const dst = [[10, 20], [130, 5], [150, 90], [0, 70]];
  const H = homography(src, dst);
  src.forEach((p, i) => assert.ok(samePoint(applyHomography(H, ...p), dst[i]), `corner ${i}`));
  const back = invertHomography(H);
  dst.forEach((p, i) => assert.ok(samePoint(applyHomography(back, ...p), src[i], 1e-6)));
  // A plain move is a move.
  const move = homography(src, src.map(([x, y]) => [x + 7, y - 3]));
  assert.ok(samePoint(applyHomography(move, 40, 40), [47, 37]));
  assert.equal(homography([[0, 0], [1, 1], [2, 2], [3, 3]], src), null, 'four points in a line make no map');
  assert.deepEqual(solve([[2, 0], [0, 4]], [6, 8]), [3, 2]);
});

test('I14: perspective correction straightens four points on a skewed rectangle into an upright one, and the picture with them', () => {
  // A sign photographed from the left: its far side shorter.
  const pins = [[100, 120], [300, 150], [300, 250], [100, 280]];
  assert.ok(isConvexQuad(pins));
  const rect = rectifyTarget(pins);
  // As wide as its top and bottom on average, as tall as its sides; centred where it was.
  const top = Math.hypot(200, 30);
  assert.ok(close(rect[1][0] - rect[0][0], top, 1e-9));
  assert.ok(close(rect[3][1] - rect[0][1], (160 + 100) / 2, 1e-9));
  assert.ok(close((rect[0][0] + rect[2][0]) / 2, 200) && close((rect[0][1] + rect[2][1]) / 2, 200));
  const layer = rectCorners(400, 400);
  const done = straighten(pins, layer);
  pins.forEach((p, i) => assert.ok(samePoint(applyHomography(done.H, ...p), rect[i], 1e-6), 'each pin lands on its corner'));
  assert.equal(done.corners.length, 4);
  assert.ok(!samePoint(done.corners[1], layer[1], 1), 'the picture\'s own corners move with the map');
  assert.equal(straighten([[0, 0], [10, 10], [10, 0], [0, 10]], layer), null, 'pins that cross make no rectangle');
});

test('I14: corner handles scale from the far corner, skew along an edge, distort one corner, or pull a perspective', () => {
  const q = rectCorners(100, 60, 10, 20);
  // Scale: the bottom-right corner pulled out; the top left stays.
  const s = dragCorner(q, 2, [160, 110], 'scale');
  assert.ok(samePoint(s[0], q[0]));
  assert.ok(samePoint(s[2], [160, 110]));
  assert.ok(samePoint(s[1], [160, 20]) && samePoint(s[3], [10, 110]));
  // Keep the shape: the larger stretch wins on both sides.
  const k = dragCorner(q, 2, [210, 90], 'scale', { keep: true });
  assert.ok(close((k[2][0] - k[0][0]) / (k[2][1] - k[0][1]), 100 / 60, 1e-9));
  // A turned box scales in its own frame.
  const turned = boxCorners({ x: 0, y: 0, w: 100, h: 50 }, 90);
  const ts = dragCorner(turned, 2, applyHomography(homography(rectCorners(100, 50), turned), 200, 50), 'scale');
  assert.ok(close(Math.hypot(ts[1][0] - ts[0][0], ts[1][1] - ts[0][1]), 200, 1e-6));
  // Skew: the top edge slides sideways, the bottom stays.
  const sk = dragCorner(q, 1, [140, 23], 'skew');
  assert.ok(samePoint(sk[0], [40, 20]) && samePoint(sk[1], [140, 20]) && samePoint(sk[2], q[2]) && samePoint(sk[3], q[3]));
  // Distort: one corner alone.
  const d = dragCorner(q, 3, [0, 99], 'distort');
  assert.ok(samePoint(d[3], [0, 99]) && samePoint(d[0], q[0]) && samePoint(d[2], q[2]));
  // Perspective: the corner and its partner on the edge move apart.
  const p = dragCorner(q, 0, [30, 21], 'perspective');
  assert.ok(samePoint(p[0], [30, 20]) && samePoint(p[1], [90, 20]));
  assert.deepEqual(quadBounds(q), { x: 10, y: 20, w: 100, h: 60 });
});

test('I14: a picture warped onto four corners lands inside them, its corners\' colours at the corners', () => {
  // A 20 × 20 picture: red top left, green top right, blue bottom right, white bottom left.
  const sw = 20;
  const data = new Uint8ClampedArray(sw * sw * 4);
  for (let y = 0; y < sw; y++) {
    for (let x = 0; x < sw; x++) {
      const c = y < 10 ? (x < 10 ? [255, 0, 0] : [0, 255, 0]) : (x < 10 ? [255, 255, 255] : [0, 0, 255]);
      data.set([...c, 255], (y * sw + x) * 4);
    }
  }
  const corners = [[10, 10], [70, 20], [60, 70], [5, 60]];
  const out = warpPixels({ data, width: sw, height: sw }, corners, { width: 80, height: 80 });
  const at = (x, y) => [...out.slice((y * 80 + x) * 4, (y * 80 + x) * 4 + 4)];
  assert.deepEqual(at(14, 15), [255, 0, 0, 255]);
  assert.deepEqual(at(64, 24), [0, 255, 0, 255]);
  assert.deepEqual(at(57, 64), [0, 0, 255, 255]);
  assert.deepEqual(at(10, 56), [255, 255, 255, 255]);
  assert.equal(at(2, 2)[3], 0, 'outside the corners is clear');
  assert.equal(at(78, 78)[3], 0);
  // On its own rectangle the picture is itself.
  const same = warpPixels({ data, width: sw, height: sw }, rectCorners(sw, sw), { width: sw, height: sw });
  assert.deepEqual([...same], [...data]);
});
