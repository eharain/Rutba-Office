// Image's edit list as a plan: one matrix from the source's pixels to the
// result's, so a flip after a turn mirrors what is on the screen, a crop is
// taken where the person drew it, straighten keeps the largest upright
// rectangle, and a mark stays on what it was drawn on through whatever comes
// after it.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  OPS, toCanvasPlan, resultSize, isIdentity, straightenRect, toTarget, toSource, markToOperation, spotToOperation, spotInTarget, resizeTo,
} from '@rutba/imaging/pipeline';

const SRC = { width: 400, height: 200 };
const near = (a, b, tol = 0.01) => Math.abs(a - b) <= tol;
const nearPoint = (p, q, tol = 0.01) => near(p.x, q.x, tol) && near(p.y, q.y, tol);

test('no edits: the result is the source, pixel for pixel', () => {
  const plan = toCanvasPlan(SRC, []);
  assert.deepEqual(plan.target, { width: 400, height: 200 });
  assert.deepEqual(plan.matrix, [1, 0, 0, 1, 0, 0]);
  assert.ok(isIdentity([]));
});

test('a quarter turn clockwise: the size swaps and the top-left corner goes to the top-right', () => {
  const plan = toCanvasPlan(SRC, [{ op: OPS.rotate, degrees: 90 }]);
  assert.deepEqual(plan.target, { width: 200, height: 400 });
  assert.ok(nearPoint(toTarget(plan, { x: 0, y: 0 }), { x: 200, y: 0 }));
  assert.ok(nearPoint(toTarget(plan, { x: 400, y: 200 }), { x: 0, y: 400 }));
  // Four turns are where it began, exactly.
  const four = toCanvasPlan(SRC, Array(4).fill({ op: OPS.rotate, degrees: 90 }));
  assert.deepEqual(four.target, SRC);
  assert.ok(four.matrix.every((v, i) => near(v, [1, 0, 0, 1, 0, 0][i], 1e-9)));
});

test('flip vertical turns the picture upside down; flip horizontal mirrors it', () => {
  const v = toCanvasPlan(SRC, [{ op: OPS.flip, axis: 'y' }]);
  assert.ok(nearPoint(toTarget(v, { x: 10, y: 0 }), { x: 10, y: 200 }));
  const h = toCanvasPlan(SRC, [{ op: OPS.flip, axis: 'x' }]);
  assert.ok(nearPoint(toTarget(h, { x: 0, y: 10 }), { x: 400, y: 10 }));
});

test('a flip after a turn mirrors the picture as it is shown, not as it was', () => {
  // Turned on its side, then flipped left to right on the screen: a point
  // keeps its height on the screen and moves across.
  const turned = toCanvasPlan(SRC, [{ op: OPS.rotate, degrees: 90 }]);
  const both = toCanvasPlan(SRC, [{ op: OPS.rotate, degrees: 90 }, { op: OPS.flip, axis: 'x' }]);
  const p = { x: 50, y: 30 };
  const a = toTarget(turned, p);
  const b = toTarget(both, p);
  assert.ok(near(a.y, b.y) && near(b.x, 200 - a.x), `${JSON.stringify(a)} then ${JSON.stringify(b)}`);
});

test('a crop is taken where it was drawn on the screen, even on a turned picture', () => {
  const plan = toCanvasPlan(SRC, [{ op: OPS.rotate, degrees: 90 }, { op: OPS.crop, x: 20, y: 40, width: 100, height: 60 }]);
  assert.deepEqual(plan.target, { width: 100, height: 60 });
  // What was at (20, 40) on the turned picture is now the corner.
  const turned = toCanvasPlan(SRC, [{ op: OPS.rotate, degrees: 90 }]);
  const was = toSource(turned, { x: 20, y: 40 });
  assert.ok(nearPoint(toTarget(plan, was), { x: 0, y: 0 }), JSON.stringify(toTarget(plan, was)));
  // Crops compose: a crop of a crop.
  const twice = toCanvasPlan(SRC, [{ op: OPS.crop, x: 100, y: 50, width: 200, height: 100 }, { op: OPS.crop, x: 10, y: 10, width: 50, height: 50 }]);
  assert.deepEqual(twice.sourceRect, { x: 110, y: 60, width: 50, height: 50 });
});

test('resize: to pixels, as shown, and kept through a later crop', () => {
  const half = toCanvasPlan(SRC, [{ op: OPS.resize, width: 200, height: 100 }]);
  assert.deepEqual(half.target, { width: 200, height: 100 });
  assert.ok(nearPoint(toTarget(half, { x: 400, y: 200 }), { x: 200, y: 100 }));
  // On a turned picture the width given is the width seen.
  const side = toCanvasPlan(SRC, [{ op: OPS.rotate, degrees: 90 }, { op: OPS.resize, width: 100, height: 200 }]);
  assert.deepEqual(side.target, { width: 100, height: 200 });
  const cropped = toCanvasPlan(SRC, [{ op: OPS.resize, width: 200, height: 100 }, { op: OPS.crop, x: 0, y: 0, width: 100, height: 50 }]);
  assert.deepEqual(cropped.target, { width: 100, height: 50 });
  assert.deepEqual(cropped.sourceRect, { x: 0, y: 0, width: 200, height: 100 });
  assert.deepEqual(resultSize(SRC, [{ op: OPS.resize, width: 200, height: 100 }]), { width: 200, height: 100 });
});

test('resize by pixels or percent, the ratio kept or not', () => {
  assert.deepEqual(resizeTo(SRC, { by: 'percent', width: 50 }), { width: 200, height: 100 });
  assert.deepEqual(resizeTo(SRC, { by: 'percent', width: 50, height: 25, keep: false }), { width: 200, height: 50 });
  assert.deepEqual(resizeTo(SRC, { by: 'pixels', width: 300 }), { width: 300, height: 150 });
  assert.deepEqual(resizeTo(SRC, { by: 'pixels', height: 50, changed: 'height' }), { width: 100, height: 50 });
  assert.deepEqual(resizeTo(SRC, { by: 'pixels', width: 300, height: 300, keep: false }), { width: 300, height: 300 });
});

test('straighten keeps the largest upright rectangle inside the turned picture', () => {
  assert.deepEqual(straightenRect(400, 200, 0), { width: 400, height: 200 });
  const r = straightenRect(400, 200, 5);
  // Its corners all lie inside the turned picture.
  const a = (5 * Math.PI) / 180;
  const inside = (x, y) => {
    // Turn the corner back and see that it is in the original rectangle.
    const ux = x * Math.cos(-a) - y * Math.sin(-a);
    const uy = x * Math.sin(-a) + y * Math.cos(-a);
    return Math.abs(ux) <= 200 + 1e-6 && Math.abs(uy) <= 100 + 1e-6;
  };
  for (const [sx, sy] of [[1, 1], [1, -1], [-1, 1], [-1, -1]]) assert.ok(inside((sx * r.width) / 2, (sy * r.height) / 2), JSON.stringify(r));
  // And it is the largest: one pixel more each way would not fit.
  assert.ok(!inside((r.width + 2) / 2, (r.height + 2) / 2));
  assert.ok(r.width < 400 && r.height < 200 && r.width > 330 && r.height > 150, JSON.stringify(r));
  // Symmetric, and the plan's result is that rectangle.
  assert.deepEqual(straightenRect(400, 200, -5), r);
  const plan = toCanvasPlan(SRC, [{ op: OPS.straighten, degrees: 5 }]);
  assert.deepEqual(plan.target, r);
  // The middle of the picture stays in the middle.
  assert.ok(nearPoint(toTarget(plan, { x: 200, y: 100 }), { x: r.width / 2, y: r.height / 2 }));
  assert.ok(!isIdentity([{ op: OPS.straighten, degrees: 5 }]));
  assert.ok(isIdentity([{ op: OPS.straighten, degrees: 0 }]));
});

test('a mark drawn on the result is kept in the source and stays on what it was drawn on', () => {
  const ops = [{ op: OPS.rotate, degrees: 90 }, { op: OPS.resize, width: 100, height: 200 }];
  const plan = toCanvasPlan(SRC, ops);
  const rect = markToOperation(plan, { shape: 'rect', x1: 10, y1: 20, x2: 50, y2: 40, colour: '#ff0000', width: 4 });
  assert.equal(rect.op, OPS.annotate);
  // Its centre lands back where it was drawn.
  assert.ok(nearPoint(toTarget(plan, { x: rect.cx, y: rect.cy }), { x: 30, y: 30 }, 0.02));
  // Its frame undoes the turn: the frame times the plan's turn is upright.
  const [a, b, c, d] = plan.matrix;
  const [fa, fb, fc, fd] = rect.frame;
  const upright = [a * fa + c * fb, b * fa + d * fb, a * fc + c * fd, b * fc + d * fd];
  const k = Math.sqrt(Math.abs(a * d - b * c));
  assert.ok(near(upright[0], k, 1e-3) && near(upright[1], 0, 1e-3) && near(upright[2], 0, 1e-3) && near(upright[3], k, 1e-3), JSON.stringify(upright));
  // The width given in the result's pixels is kept in the source's: half size, twice as wide.
  assert.ok(near(rect.width, 8, 0.01), String(rect.width));
  // A later crop moves it with the picture.
  const later = toCanvasPlan(SRC, [...ops, rect, { op: OPS.crop, x: 5, y: 5, width: 80, height: 80 }]);
  assert.ok(nearPoint(toTarget(later, { x: rect.cx, y: rect.cy }), { x: 25, y: 25 }, 0.02));
  assert.equal(later.overlays.length, 1);
  // Pen and arrow keep their points; text its corner and size.
  const pen = markToOperation(plan, { shape: 'pen', points: [{ x: 1, y: 2 }, { x: 3, y: 4 }], width: 2 });
  assert.equal(pen.points.length, 2);
  assert.ok(nearPoint(toTarget(plan, pen.points[1]), { x: 3, y: 4 }, 0.02));
  const text = markToOperation(plan, { shape: 'text', x1: 5, y1: 6, text: 'Hello', size: 20 });
  assert.equal(text.text, 'Hello');
  assert.ok(near(text.size, 40, 0.01));
});

test('a red-eye spot is kept in the source and comes back where it was clicked', () => {
  const plan = toCanvasPlan(SRC, [{ op: OPS.resize, width: 200, height: 100 }]);
  const spot = spotToOperation(plan, { x: 50, y: 25, radius: 6 });
  assert.deepEqual({ x: spot.x, y: spot.y, radius: spot.radius }, { x: 100, y: 50, radius: 12 });
  const back = spotInTarget(toCanvasPlan(SRC, [{ op: OPS.resize, width: 200, height: 100 }, spot]), spot);
  assert.ok(near(back.x, 50) && near(back.y, 25) && near(back.radius, 6));
  assert.equal(toCanvasPlan(SRC, [spot]).spots.length, 1);
});

test('pixel settings and adjustments merge, the last one counting', () => {
  const plan = toCanvasPlan(SRC, [{ op: OPS.pixels, exposure: 1 }, { op: OPS.pixels, sharpen: 20 }, { op: OPS.adjust, invert: 100 }]);
  assert.deepEqual(plan.pixels, { exposure: 1, sharpen: 20 });
  assert.equal(plan.filter, 'invert(100%)');
  assert.ok(isIdentity([{ op: OPS.pixels, exposure: 0, levels: { inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 } }]));
  assert.ok(!isIdentity([{ op: OPS.pixels, curve: { rgb: [[0, 0], [128, 160], [255, 255]] } }]));
});
