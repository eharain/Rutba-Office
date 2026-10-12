// The Image app's painted masks, paint layers and corners on the studio
// document (docs/STUDIO.md 7.2, I10, I12 and I14): a mask added, changed and
// taken away; a paint layer over the whole canvas put above the layer chosen;
// a layer moved with its mask and corners; the canvas reframed with them; the
// map from the canvas to a layer's own pixels; and the renderer cutting a
// layer by its mask on a surface of its own before laying it on the frame
// with its opacity and blend, a feather and an inversion read from the patch.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  emptyImage, pictureLayer, paintLayer, isPixelLayer, insertAbove, addMask, setMask, removeMask, moveLayerBy,
  reframe, canvasSize, pictureBox, canvasToBitmap, mapPoint, affineAt, groupLayers, imagePatches,
} from '../packages/studio/src/layers.js';
import { rimgPaths } from '../packages/studio/src/rimg.js';
import { buildImagePlan } from '../packages/studio/src/image.js';
import { paintFrame } from '../packages/studio/src/index.js';
import '../packages/studio/src/toolset.js';

function project() {
  let doc = emptyImage({ width: 800, height: 400 });
  doc = { ...doc, patches: [
    pictureLayer(doc, { url: 'rimg:bitmaps/background.png', width: 800, height: 400, label: 'Background', id: 'background' }),
    { id: 'shape-1', type: 'shape', geometry: 'rect', fx: 0.5, fy: 0.5, fw: 0.2, fh: 0.2 },
  ] };
  return doc;
}

test('I12: a paint layer is a clear picture over the whole canvas, put directly above the layer chosen', () => {
  const doc = project();
  const p = paintLayer(doc);
  assert.equal(p.type, 'image');
  assert.equal(p.kind, 'paint');
  assert.equal(p.url, `rimg:bitmaps/${p.id}.png`);
  assert.deepEqual([p.fx, p.fy, p.fw], [0, 0, 1]);
  assert.ok(isPixelLayer(p) && isPixelLayer(doc.patches[0]) && !isPixelLayer(doc.patches[1]));
  const next = insertAbove(doc, 'background', p);
  assert.deepEqual(next.patches.map((x) => x.id), ['background', p.id, 'shape-1']);
  assert.deepEqual(insertAbove(doc, 'nothing', p).patches.at(-1).id, p.id, 'with nothing chosen, on top');
  // Inside a group, it joins the group.
  const grouped = groupLayers(doc, ['shape-1']).doc;
  const inGroup = insertAbove(grouped, 'shape-1', paintLayer(grouped, { id: 'paint-9' }));
  assert.equal(inGroup.patches.find((x) => x.id === 'paint-9').parent, inGroup.patches.find((x) => x.type === 'group').id);
});

test('I10: a mask is added over the whole canvas, changed, turned off and taken away; the project file keeps its bitmap', () => {
  let doc = addMask(project(), 'shape-1', { url: 'rimg:bitmaps/shape-1-mask.png', outside: 1 });
  const m = doc.patches[1].paintMask;
  assert.deepEqual(m, { url: 'rimg:bitmaps/shape-1-mask.png', fx: 0, fy: 0, fw: 1, fh: 1, outside: 1, feather: 0, invert: false, off: false });
  doc = setMask(doc, 'shape-1', { feather: 0.02, invert: true });
  assert.equal(doc.patches[1].paintMask.feather, 0.02);
  assert.equal(doc.patches[1].paintMask.invert, true);
  assert.equal(doc.patches[1].paintMask.url, m.url, 'a change keeps the rest');
  assert.equal(setMask(doc, 'background', { off: true }), doc, 'a layer with no mask is left alone');
  assert.ok(rimgPaths(doc).includes('bitmaps/shape-1-mask.png'));
  const gone = removeMask(doc, 'shape-1');
  assert.equal(gone.patches[1].paintMask, undefined);
  assert.ok(!rimgPaths(gone).includes('bitmaps/shape-1-mask.png'));
});

test('I10 and I14: a layer moves with its mask and its corners; a new canvas keeps them over the same pixels', () => {
  let doc = addMask(project(), 'background', { url: 'rimg:bitmaps/bg-mask.png' });
  doc = { ...doc, patches: doc.patches.map((p) => (p.id === 'background' ? { ...p, corners: [[0, 0], [1, 0], [1, 1], [0, 1]], paintMask: { ...p.paintMask, feather: 0.01 } } : p)) };
  const moved = moveLayerBy(doc, 'background', 0.1, -0.05).patches[0];
  assert.deepEqual([moved.fx, moved.fy], [0.1, -0.05]);
  assert.deepEqual(moved.corners[2], [1.1, 0.95]);
  assert.deepEqual([moved.paintMask.fx, moved.paintMask.fy], [0.1, -0.05]);
  // 200 px more room on the left: everything slides right by a fifth of the new width.
  const wider = canvasSize(doc, { width: 1000, height: 400, anchor: 'e' }).patches[0];
  assert.deepEqual([wider.paintMask.fx, wider.paintMask.fw, wider.paintMask.feather], [0.2, 0.8, 0.008]);
  assert.deepEqual(wider.corners, [[0.2, 0], [1, 0], [1, 1], [0.2, 1]]);
  const cut = reframe(doc, { x: 400, y: 0, width: 400, height: 400 }).patches[0];
  assert.deepEqual([cut.paintMask.fx, cut.paintMask.fw], [-1, 2]);
});

test('I12: a point on the canvas maps to a layer\'s own pixels, through a turn and through its corners', () => {
  const doc = emptyImage({ width: 1000, height: 500 });
  const bitmap = { width: 200, height: 100 };
  // A picture 400 wide (200 × 100 at twice its size) with its top left at (100, 50).
  const patch = { id: 'p', type: 'image', url: 'rimg:bitmaps/p.png', fx: 0.1, fy: 0.1, fw: 0.4 };
  assert.deepEqual(pictureBox(doc, patch, bitmap).box, { x: 100, y: 50, w: 400, h: 200 });
  const m = canvasToBitmap(doc, patch, bitmap);
  const p = mapPoint(m, 300, 150);
  assert.ok(Math.abs(p.x - 100) < 1e-9 && Math.abs(p.y - 50) < 1e-9, 'the middle of the box is the middle of the bitmap');
  const tl = mapPoint(m, 100, 50);
  assert.ok(Math.abs(tl.x) < 1e-9 && Math.abs(tl.y) < 1e-9);
  // Turned a quarter: the box's top middle is where the right middle was.
  const turned = canvasToBitmap(doc, { ...patch, rot: 90 }, bitmap);
  const q = mapPoint(turned, 300 + 100, 150);
  assert.ok(Math.abs(q.x - 100) < 1e-9 && Math.abs(q.y - 0) < 1e-9, `turned: ${JSON.stringify(q)}`);
  // On corners: each corner is a corner of the bitmap.
  const onCorners = { ...patch, corners: [[0.1, 0.1], [0.6, 0.2], [0.55, 0.9], [0.05, 0.8]] };
  const w = canvasToBitmap(doc, onCorners, bitmap);
  const c2 = mapPoint(w, 550, 450);
  assert.ok(Math.abs(c2.x - 200) < 1e-6 && Math.abs(c2.y - 100) < 1e-6);
  // The flat map about a point of a flat layer is the map itself.
  const a = affineAt(m, 300, 150);
  assert.ok(Math.abs(a[0] - 0.5) < 1e-9 && Math.abs(a[3] - 0.5) < 1e-9 && Math.abs(a[4] - (100 - 150)) < 1e-6);
});

/* ── painted: a recording context ─────────────────────────────────────── */

function stubCtx(tag) {
  const calls = [];
  const state = { globalAlpha: 1, globalCompositeOperation: 'source-over', filter: 'none' };
  const record = (name) => (...args) => { calls.push({ tag, name, args, state: { ...state } }); };
  const ctx = { calls, tag };
  for (const n of ['save', 'restore', 'translate', 'rotate', 'scale', 'beginPath', 'moveTo', 'lineTo', 'quadraticCurveTo', 'bezierCurveTo',
    'closePath', 'ellipse', 'rect', 'roundRect', 'clip', 'fill', 'stroke', 'arc', 'fillRect', 'clearRect', 'setLineDash', 'fillText', 'strokeText',
    'drawImage', 'putImageData', 'setTransform', 'resetTransform']) ctx[n] = record(n);
  ctx.measureText = (s) => ({ width: String(s).length * 10 });
  ctx.createLinearGradient = () => ({ addColorStop() {} });
  ctx.createRadialGradient = () => ({ addColorStop() {} });
  ctx.getImageData = (x, y, w, h) => ({ data: new Uint8ClampedArray(w * h * 4).fill(128), width: w, height: h });
  ctx.createImageData = (w, h) => ({ data: new Uint8ClampedArray(w * h * 4), width: w, height: h });
  for (const prop of ['font', 'fillStyle', 'strokeStyle', 'lineWidth', 'globalAlpha', 'filter', 'globalCompositeOperation', 'shadowColor', 'shadowBlur',
    'textAlign', 'textBaseline', 'lineJoin', 'lineCap', 'direction', 'imageSmoothingEnabled', 'imageSmoothingQuality']) {
    Object.defineProperty(ctx, prop, { get: () => state[prop], set: (v) => { state[prop] = v; calls.push({ tag, name: `set:${prop}`, args: [v] }); } });
  }
  return ctx;
}
const surfaces = [];
globalThis.OffscreenCanvas = class {
  constructor(w, h) { this.width = w; this.height = h; this.ctx = stubCtx(`surface${surfaces.length}`); this.ctx.canvas = this; surfaces.push(this); }
  getContext() { return this.ctx; }
};

test('I10: a masked layer is painted on its own surface, cut by its mask there, and laid on the frame with its opacity and blend', () => {
  surfaces.length = 0;
  const frame = stubCtx('frame');
  const mask = { img: { tag: 'mask-bitmap' }, width: 800, height: 400, version: 1 };
  let doc = project();
  doc = addMask(doc, 'shape-1', { url: 'rimg:bitmaps/m.png' });
  doc = setMask(doc, 'shape-1', { feather: 0.01 });
  doc = { ...doc, patches: doc.patches.map((p) => (p.id === 'shape-1' ? { ...p, opacity: 0.5, blend: 'multiply' } : p)) };
  const canvas = { width: 0, height: 0, getContext: () => frame };
  const plan = buildImagePlan({ canvas, size: doc.size, options: doc.options, layerPatches: imagePatches(doc), assets: { 'rimg:bitmaps/m.png': mask } });
  const shape = plan.layers.find((l) => l.id === 'shape-1');
  assert.deepEqual({ x: shape.paintMask.x, w: shape.paintMask.w, feather: shape.paintMask.feather, outside: shape.paintMask.outside }, { x: 0, w: 800, feather: 8, outside: 1 });
  paintFrame(frame, plan, 0);
  // The shape is filled on a surface, not on the frame.
  assert.ok(!frame.calls.some((c) => c.name === 'fill'), 'the frame never sees the shape drawn');
  const layerSurface = surfaces.find((s) => s.ctx.calls.some((c) => c.name === 'fill'));
  assert.ok(layerSurface, 'it is drawn on a surface of its own');
  const cut = layerSurface.ctx.calls.find((c) => c.name === 'drawImage' && c.state.globalCompositeOperation === 'destination-in');
  assert.ok(cut, 'then cut by its mask');
  const fills = layerSurface.ctx.calls.findIndex((c) => c.name === 'fill');
  assert.ok(layerSurface.ctx.calls.indexOf(cut) > fills, 'after it is drawn');
  // The mask surface: the bitmap laid in, blurred by half the feather.
  const blur = surfaces.find((s) => s.ctx.calls.some((c) => c.name === 'set:filter' && /blur\(4\.00px\)/.test(c.args[0])));
  assert.ok(blur, 'feathered by a blur');
  assert.ok(surfaces.some((s) => s.ctx.calls.some((c) => c.name === 'drawImage' && c.args[0] === mask.img)), 'from the painted bitmap');
  // On the frame: the layer's surface, at half strength, multiplied.
  const laid = frame.calls.find((c) => c.name === 'drawImage' && c.args[0] === layerSurface);
  assert.ok(laid, 'the surface is laid on the frame');
  assert.equal(laid.state.globalAlpha, 0.5);
  assert.equal(laid.state.globalCompositeOperation, 'multiply');
});

test('I10: a mask turned off paints the layer as if it had none; an inverted one is cut out of white', () => {
  surfaces.length = 0;
  const frame = stubCtx('frame');
  const assets = { 'rimg:bitmaps/m.png': { img: { tag: 'm' }, width: 10, height: 10 } };
  let doc = addMask(project(), 'shape-1', { url: 'rimg:bitmaps/m.png' });
  const off = setMask(doc, 'shape-1', { off: true });
  const canvas = { width: 0, height: 0, getContext: () => frame };
  paintFrame(frame, buildImagePlan({ canvas, size: off.size, options: off.options, layerPatches: imagePatches(off), assets }), 0);
  assert.ok(frame.calls.some((c) => c.name === 'fill'), 'off: drawn straight on the frame');
  assert.equal(surfaces.length, 0);
  doc = setMask(doc, 'shape-1', { invert: true });
  const frame2 = stubCtx('frame2');
  paintFrame(frame2, buildImagePlan({ canvas: { width: 0, height: 0, getContext: () => frame2 }, size: doc.size, options: doc.options, layerPatches: imagePatches(doc), assets }), 0);
  const inverted = surfaces.find((s) => s.ctx.calls.some((c) => c.name === 'drawImage' && c.state.globalCompositeOperation === 'destination-out'));
  assert.ok(inverted, 'inverted: the mask taken out of a white surface');
});
