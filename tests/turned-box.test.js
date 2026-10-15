// A turned object's frame (packages/office-ui/src/turned-box.js): the
// corners, handles and rotate handle turn with the object about its centre,
// a point is inside by the turned outline, and a handle drag resizes along
// the object's own sides with the opposite handle staying put on the page.

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  turnPoint, framePoint, frameCorners, frameHandles, rotateHandle, frameBounds,
  insideFrame, handleCursor, resizeTurned,
} from '@rutba/office-ui/turned-box';

const near = (a, b, msg, eps = 1e-6) => assert.ok(Math.abs(a - b) < eps, `${msg}: ${a} vs ${b}`);
const nearPt = (p, q, msg, eps = 1e-6) => { near(p.x, q.x, `${msg} x`, eps); near(p.y, q.y, `${msg} y`, eps); };

test('a point turns clockwise on screen about the centre', () => {
  nearPt(turnPoint(10, 0, 0, 0, 90), { x: 0, y: 10 }, 'right of centre goes below at 90');
  nearPt(turnPoint(10, 0, 0, 0, 0), { x: 10, y: 0 }, 'no turn');
  nearPt(turnPoint(10, 0, 0, 0, 180), { x: -10, y: 0 }, 'a half turn');
});

test('the corners, the handles and the rotate handle turn with the box', () => {
  const box = { x: 100, y: 100, w: 200, h: 100 };
  // At 90 degrees the top left corner goes to the top right of the turned outline.
  const [nw, ne, se, sw] = frameCorners(box, 90);
  nearPt(nw, { x: 250, y: 50 }, 'nw');
  nearPt(ne, { x: 250, y: 250 }, 'ne');
  nearPt(se, { x: 150, y: 250 }, 'se');
  nearPt(sw, { x: 150, y: 50 }, 'sw');
  const hs = frameHandles(box, 90);
  assert.equal(hs.length, 8);
  nearPt(hs.find((h) => h.name === 'n'), { x: 250, y: 150 }, 'the top middle is on the right at 90');
  const r = rotateHandle(box, 90, 26);
  nearPt(r.stem, { x: 250, y: 150 }, 'the stem starts at the turned top edge');
  nearPt(r, { x: 276, y: 150 }, 'the rotate handle stands off the turned top edge');
  // Upright, nothing moves.
  nearPt(framePoint(box, 0, 1, 1), { x: 300, y: 200 }, 'se upright');
  const r0 = rotateHandle(box, 0, 26);
  nearPt(r0, { x: 200, y: 74 }, 'upright rotate handle above the top');
});

test('the bounds round a turned box hold every corner', () => {
  const box = { x: 0, y: 0, w: 200, h: 100 };
  const b = frameBounds(box, 30);
  for (const c of frameCorners(box, 30)) {
    assert.ok(c.x >= b.x - 1e-9 && c.x <= b.x + b.w + 1e-9 && c.y >= b.y - 1e-9 && c.y <= b.y + b.h + 1e-9);
  }
  near(b.w, 200 * Math.cos(Math.PI / 6) + 100 * Math.sin(Math.PI / 6), 'width of the bounds');
});

test('a point inside the turned outline is inside; the empty corner of the upright box is not', () => {
  const box = { x: 100, y: 100, w: 300, h: 60 };
  assert.ok(insideFrame(box, 45, 250, 130), 'the centre');
  // The upright box's top left corner is far outside a long thin box turned 45 degrees.
  assert.ok(!insideFrame(box, 45, 105, 105), 'the old corner');
  assert.ok(insideFrame(box, 0, 105, 105), 'upright, the corner is inside');
  // A point on the turned outline's far end, outside the upright box.
  const tip = framePoint(box, 45, 0.98, 0.5);
  assert.ok(tip.y > 160, 'the tip is below the upright box');
  assert.ok(insideFrame(box, 45, tip.x, tip.y), 'the turned tip is inside');
});

test('the cursors turn with the handles', () => {
  assert.equal(handleCursor('e', 0), 'ew-resize');
  assert.equal(handleCursor('e', 90), 'ns-resize');
  assert.equal(handleCursor('se', 0), 'nwse-resize');
  assert.equal(handleCursor('se', 90), 'nesw-resize');
  assert.equal(handleCursor('n', 45), 'nesw-resize');
  assert.equal(handleCursor('n', 135), 'nwse-resize');
});

test('upright, a handle drag is the plain resize', () => {
  const box = { x: 10, y: 20, w: 100, h: 50 };
  assert.deepEqual(resizeTurned(box, 0, 'se', 30, 10), { x: 10, y: 20, w: 130, h: 60 });
  assert.deepEqual(resizeTurned(box, 0, 'nw', 10, 10), { x: 20, y: 30, w: 90, h: 40 });
  assert.deepEqual(resizeTurned(box, 0, 'e', 30, 99), { x: 10, y: 20, w: 130, h: 50 });
  assert.deepEqual(resizeTurned(box, 0, 'nw', 200, 0), { x: 102, y: 20, w: 8, h: 50 }, 'never smaller than the least side');
});

test('turned, a corner drag resizes along the box\'s own sides and the opposite corner stays put', () => {
  for (const deg of [30, 90, 135, 200, 315]) {
    const box = { x: 200, y: 150, w: 160, h: 90 };
    for (const [handle, opposite] of [['se', [0, 0]], ['nw', [1, 1]], ['ne', [0, 1]], ['sw', [1, 0]]]) {
      const fixed = framePoint(box, deg, ...opposite);
      // A drag of 40 along the box's own width and 20 along its height, outwards.
      const out = { se: [40, 20], nw: [-40, -20], ne: [40, -20], sw: [-40, 20] }[handle];
      const page = turnPoint(out[0], out[1], 0, 0, deg);
      const next = resizeTurned(box, deg, handle, page.x, page.y);
      near(next.w, 200, `${deg} ${handle} width`, 1e-6);
      near(next.h, 110, `${deg} ${handle} height`, 1e-6);
      nearPt(framePoint(next, deg, ...opposite), fixed, `${deg} ${handle} the opposite corner`, 1e-6);
    }
  }
});

test('turned, a side handle changes one side and the opposite side stays put', () => {
  const box = { x: 0, y: 0, w: 100, h: 40 };
  const deg = 90;
  const fixed = framePoint(box, deg, 0, 0.5);
  // Dragging the east handle straight down the page is along the box's width at 90 degrees.
  const next = resizeTurned(box, deg, 'e', 0, 50);
  near(next.w, 150, 'width');
  near(next.h, 40, 'height');
  nearPt(framePoint(next, deg, 0, 0.5), fixed, 'the west side');
  // Across the page is across the box at 90: the east handle ignores it.
  const across = resizeTurned(box, deg, 'e', 50, 0);
  near(across.w, 100, 'no change');
});

test('a turned corner drag keeps the proportions when asked, from the opposite corner', () => {
  const box = { x: 50, y: 50, w: 200, h: 100 };
  const deg = 30;
  const fixed = framePoint(box, deg, 0, 0);
  const page = turnPoint(100, 0, 0, 0, deg);
  const next = resizeTurned(box, deg, 'se', page.x, page.y, { keep: true });
  near(next.w, 300, 'width');
  near(next.h, 150, 'height in proportion');
  nearPt(framePoint(next, deg, 0, 0), fixed, 'the opposite corner');
});
