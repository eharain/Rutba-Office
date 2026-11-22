// The renderer's part of D-2026-11-22-14: the ways in and out added to the
// envelope every layer is painted through (a wipe from each side, an iris, a
// layer growing or spinning about its own middle), and the keys a layer moves
// by reaching the layers the toolset registers and a box's own height. What
// was there before paints as it did.
import test from 'node:test';
import assert from 'node:assert/strict';
import { applyLayerPatches, paintFrame, layerBounds, withLayerStateAt } from '../packages/studio/src/index.js';
import '../packages/studio/src/toolset.js';

function stubCtx() {
  const calls = [];
  const state = {};
  const record = (name) => (...args) => { calls.push({ name, args }); };
  const ctx = { calls };
  for (const name of ['save', 'restore', 'translate', 'rotate', 'scale', 'beginPath', 'moveTo', 'lineTo', 'quadraticCurveTo', 'bezierCurveTo', 'closePath', 'ellipse', 'arc', 'arcTo', 'rect', 'roundRect', 'clip', 'fill', 'stroke', 'fillRect', 'strokeRect', 'clearRect', 'setLineDash', 'fillText', 'strokeText', 'drawImage', 'putImageData', 'setTransform']) ctx[name] = record(name);
  ctx.measureText = (s) => ({ width: String(s).length * 10 });
  ctx.createLinearGradient = () => ({ addColorStop() {} });
  ctx.createRadialGradient = () => ({ addColorStop() {} });
  ctx.getImageData = (x, y, w, h) => ({ data: new Uint8ClampedArray(w * h * 4).fill(128), width: w, height: h });
  for (const prop of ['font', 'fillStyle', 'strokeStyle', 'lineWidth', 'globalAlpha', 'filter', 'globalCompositeOperation', 'shadowColor', 'shadowBlur', 'shadowOffsetX', 'shadowOffsetY', 'textAlign', 'textBaseline', 'lineJoin', 'lineCap', 'miterLimit', 'direction']) {
    Object.defineProperty(ctx, prop, { get: () => state[prop], set: (v) => { state[prop] = v; calls.push({ name: `set:${prop}`, args: [v] }); } });
  }
  return ctx;
}

const W = 1000;
const H = 500;
function planOf(patches) {
  const plan = {
    W, H, duration: 10, fps: 30,
    theme: { bg: '#000', scrim: 'rgba(0,0,0,.7)', text: '#fff', dim: '#aaa', accent: '#ffc107' },
    layers: [], images: [], videos: {}, assets: {}, context: {},
  };
  applyLayerPatches(plan, patches);
  return plan;
}

/** A rectangle 200 by 100 in the middle of the frame, on from 2 s to 6 s: its box is x 400, y 200. */
const shape = (more = {}) => ({ id: 's', type: 'shape', geometry: 'rect', fx: 0.5, fy: 0.5, fw: 0.2, fh: 0.2, color: '#ff0000', timing: { start: 2, end: 6 }, ...more });

/** What the wrapper did before the painter drew: the calls up to the first fill. */
function before(plan, time) {
  const ctx = stubCtx();
  paintFrame(ctx, plan, time);
  // The frame is filled first (fillRect); the layer's own drawing ends in a fill or a stroke.
  const firstFill = ctx.calls.findIndex((c) => c.name === 'fill' || c.name === 'stroke');
  return ctx.calls.slice(0, firstFill < 0 ? ctx.calls.length : firstFill);
}
/** The path call (rect or arc) that the first clip closed over. */
function clipped(calls) {
  const at = calls.findIndex((c) => c.name === 'clip');
  if (at < 0) return null;
  for (let i = at - 1; i >= 0; i--) if (calls[i].name === 'rect' || calls[i].name === 'arc') return calls[i];
  return null;
}
const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 0.01, `${msg}: ${a} is not ${b}`);
const nearAll = (got, want, msg) => want.forEach((v, i) => near(got[i], v, `${msg} [${i}]`));

test('a wipe uncovers from the side its kind names, a quarter of the way at a quarter of its length', () => {
  const box = (kind) => clipped(before(planOf([shape({ enter: { kind, seconds: 2 } })]), 2.5));
  nearAll(box('wipe').args, [400, 200, 50, 100], 'from the left, as it always was');
  nearAll(box('wipe-left').args, [550, 200, 50, 100], 'from the right');
  nearAll(box('wipe-down').args, [400, 200, 200, 25], 'from the top');
  nearAll(box('wipe-up').args, [400, 275, 200, 25], 'from the bottom');
});

test('a wipe out covers back to the same side', () => {
  const box = (kind) => clipped(before(planOf([shape({ exit: { kind, seconds: 2 } })]), 5.5));
  nearAll(box('wipe').args, [400, 200, 50, 100], 'back to the left');
  nearAll(box('wipe-left').args, [550, 200, 50, 100], 'back to the right');
  nearAll(box('wipe-up').args, [400, 275, 200, 25], 'back to the bottom');
});

test('an iris opens a circle from the middle of the layer that reaches its corners when whole, and closes on the way out', () => {
  const whole = Math.hypot(200, 100) / 2;
  const arcIn = clipped(before(planOf([shape({ enter: { kind: 'iris', seconds: 2 } })]), 3));
  assert.equal(arcIn.name, 'arc');
  nearAll(arcIn.args, [500, 250, whole / 2], 'half open at half its length');
  const arcOut = clipped(before(planOf([shape({ exit: { kind: 'iris', seconds: 2 } })]), 5.5));
  nearAll(arcOut.args, [500, 250, whole / 4], 'a quarter left three quarters of the way out');
});

test('growing and spinning are about the layer\'s own middle, and fade as they go', () => {
  const grow = before(planOf([shape({ enter: { kind: 'grow', seconds: 2 } })]), 3);
  const scale = grow.find((c) => c.name === 'scale');
  nearAll(scale.args, [0.8, 0.8], 'from six tenths, half way');
  const moves = grow.filter((c) => c.name === 'translate').map((c) => c.args);
  nearAll(moves[0], [500, 250], 'to its middle');
  nearAll(moves[1], [-500, -250], 'and back');
  near(grow.find((c) => c.name === 'set:globalAlpha').args[0], 0.5, 'half there');

  const spin = before(planOf([shape({ enter: { kind: 'spin', seconds: 2 } })]), 3);
  near(spin.find((c) => c.name === 'rotate').args[0], (-45 * Math.PI) / 180, 'a quarter turn back, half undone');
  const spinOut = before(planOf([shape({ exit: { kind: 'spin', seconds: 2 } })]), 5);
  near(spinOut.find((c) => c.name === 'rotate').args[0], (45 * Math.PI) / 180, 'turning on as it goes');
  near(spinOut.find((c) => c.name === 'set:globalAlpha').args[0], 0.5, 'half gone');
});

test('outside its ways in and out a layer is painted with nothing done to it, whatever their kinds', () => {
  // What the shape's own painter does, with no way in or out at all.
  const plain = before(planOf([shape()]), 4);
  for (const kind of ['wipe-left', 'wipe-down', 'wipe-up', 'iris', 'grow', 'spin']) {
    const calls = before(planOf([shape({ enter: { kind, seconds: 1 }, exit: { kind, seconds: 1 } })]), 4);
    assert.deepEqual(calls, plain, kind);
  }
});

test('a layer the toolset registers grows by its keys, its height with its width when both are keyed', () => {
  const plan = planOf([shape({ keys: { fw: [{ t: 0, v: 0.2 }, { t: 4, v: 0.4 }], fh: [{ t: 0, v: 0.2 }, { t: 4, v: 0.4 }], fx: [{ t: 0, v: 0.5 }, { t: 4, v: 0.3 }] } })]);
  const layer = plan.layers.find((l) => l.id === 's');
  assert.deepEqual(Object.keys(layer.keys).sort(), ['fh', 'fw', 'fx'], 'a height of its own is a key the renderer keeps');
  const ctx = stubCtx();
  const at = (time) => withLayerStateAt(plan, layer, time, () => layerBounds(ctx, plan, layer));
  nearAll(Object.values(at(2)), [400, 200, 200, 100], 'at its first point');
  nearAll(Object.values(at(4)), [250, 175, 300, 150], 'half way: wider, taller and further left');
  nearAll(Object.values(at(6)), [100, 150, 400, 200], 'at its last point');
  assert.equal(layer.fw, 0.2, 'the layer itself is as it was once the frame is done');
  assert.equal(layer.fh, 0.2);
});

test('a keyed angle is seen by whoever measures the layer at a time, and the layer is left as it was', () => {
  const plan = planOf([shape({ rot: 10, keys: { rot: [{ t: 0, v: 0 }, { t: 4, v: 90 }] } })]);
  const layer = plan.layers.find((l) => l.id === 's');
  assert.equal(withLayerStateAt(plan, layer, 4, () => layer.rot), 45);
  assert.equal(layer.rot, 10);
  // Painted, it is turned by the keyed angle once.
  const turns = before(plan, 4).filter((c) => c.name === 'rotate');
  assert.equal(turns.length, 1);
  near(turns[0].args[0], (45 * Math.PI) / 180, 'the keyed angle');
});

test('a layer with no keys and the old ways in paints the calls it always did', () => {
  const plain = before(planOf([shape()]), 4);
  assert.equal(plain.filter((c) => ['clip', 'scale', 'rotate', 'set:globalAlpha', 'set:filter'].includes(c.name)).length, 0);
  const fade = before(planOf([shape({ enter: { kind: 'fade', seconds: 2 } })]), 3);
  near(fade.find((c) => c.name === 'set:globalAlpha').args[0], 0.5, 'a fade is a fade');
  assert.equal(fade.filter((c) => ['clip', 'scale', 'rotate'].includes(c.name)).length, 0);
  const zoom = before(planOf([shape({ enter: { kind: 'zoom', seconds: 2 } })]), 3);
  nearAll(zoom.find((c) => c.name === 'scale').args, [1.1, 1.1], 'a zoom is about the frame, as before');
  nearAll(zoom.find((c) => c.name === 'translate').args, [W / 2, H / 2], 'its middle is the frame\'s');
});
