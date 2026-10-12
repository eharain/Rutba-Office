// The Image app's layered picture (docs/STUDIO.md 7.2, I7, I9 and I15): a
// layer stack whose order, groups, visibility, opacity, blend and lock are
// arithmetic on the studio document; a canvas grown or cut from an anchor, an
// image size, and a trim to the pixels that show, each keeping every layer
// where it was in pixels; and the renderer painting a clear canvas, a text's
// opacity and blend, and an adjustment layer over only what is under it.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  emptyImage, normaliseImage, pictureLayer, layerRows, unitOf, dropLayer, stepLayer, layerToEnd, groupLayers,
  ungroupLayers, removeLayers, duplicateLayers, setLayer, isLocked, imagePatches, anchorOffset, reframe, canvasSize,
  imageSize, trimBounds, trimTo, scaleBox, resizeCentred, rotationAt, BLEND_MODES, ANCHORS, layerLabel,
} from '../packages/studio/src/layers.js';
import { buildImagePlan } from '../packages/studio/src/image.js';
import { paintFrame } from '../packages/studio/src/index.js';
import '../packages/studio/src/toolset.js';

const ids = (doc) => doc.patches.map((p) => p.id);
const rows = (doc) => layerRows(doc).map((r) => `${'  '.repeat(r.depth)}${r.id}`);

function stack() {
  let doc = emptyImage({ width: 1000, height: 500 });
  doc = { ...doc, patches: [
    pictureLayer(doc, { url: 'rimg:bitmaps/background.png', width: 1000, height: 500, label: 'Background', id: 'background' }),
    { id: 'shape-1', type: 'shape', geometry: 'rect', fx: 0.5, fy: 0.5, fw: 0.2, fh: 0.2 },
    { id: 'textblock-1', type: 'textblock', text: 'Hello', fx: 0.3, fy: 0.2, fw: 0.4, sizeFrac: 0.05 },
    { id: 'adjust-1', type: 'adjust', preset: 'mono' },
  ] };
  return doc;
}

test('I7: an image project is a studio document of a canvas in pixels, transparent, painting only its layers', () => {
  const doc = emptyImage({ width: 640, height: 480 });
  assert.equal(doc.kind, 'image');
  assert.deepEqual(doc.size, { width: 640, height: 480 });
  assert.equal(doc.options.bare, true);
  assert.equal(doc.options.background, 'transparent');
  const back = normaliseImage(JSON.parse(JSON.stringify({ ...doc, size: { width: 99999, height: 0 } })));
  assert.deepEqual(back.size, { width: 16384, height: 1 }, 'a size is held to what a canvas can draw');
  const pic = pictureLayer(doc, { url: 'rimg:bitmaps/a.png', width: 320, height: 240, centre: true });
  assert.deepEqual([pic.fx, pic.fy, pic.fw, pic.shadow], [0.25, 0.25, 0.5, false], 'a picture placed at one pixel to one pixel, centred');
});

test('I7: the panel lists the top layer first; bring forward, send backward, to front and to back move one layer', () => {
  let doc = stack();
  assert.deepEqual(rows(doc), ['adjust-1', 'textblock-1', 'shape-1', 'background']);
  doc = stepLayer(doc, 'shape-1', +1);
  assert.deepEqual(ids(doc), ['background', 'textblock-1', 'shape-1', 'adjust-1'], 'forward one place');
  doc = stepLayer(doc, 'shape-1', -1);
  doc = stepLayer(doc, 'shape-1', -1);
  assert.deepEqual(ids(doc), ['shape-1', 'background', 'textblock-1', 'adjust-1'], 'backward twice, under the picture');
  assert.equal(stepLayer(doc, 'shape-1', -1), doc, 'the bottom layer goes no lower');
  doc = layerToEnd(doc, 'shape-1', +1);
  assert.deepEqual(ids(doc), ['background', 'textblock-1', 'adjust-1', 'shape-1'], 'to the front');
  doc = layerToEnd(doc, 'shape-1', -1);
  assert.equal(ids(doc)[0], 'shape-1', 'to the back');
  assert.ok(doc.patches.every((p) => p.z === undefined), 'order is the list; no z is kept');
});

test('I7: a layer dragged above or below another lands there', () => {
  let doc = stack();
  doc = dropLayer(doc, 'background', 'textblock-1', 'above');
  assert.deepEqual(ids(doc), ['shape-1', 'textblock-1', 'background', 'adjust-1']);
  doc = dropLayer(doc, 'adjust-1', 'shape-1', 'below');
  assert.deepEqual(ids(doc), ['adjust-1', 'shape-1', 'textblock-1', 'background']);
  assert.equal(dropLayer(doc, 'adjust-1', 'adjust-1'), doc, 'onto itself is no move');
});

test('I7: layers grouped move as one, show and hide as one, fade as one, and come apart again', () => {
  let doc = stack();
  const made = groupLayers(doc, ['shape-1', 'textblock-1'], { label: 'Group' });
  doc = made.doc;
  assert.equal(made.id, 'group-1');
  assert.deepEqual(rows(doc), ['adjust-1', 'group-1', '  textblock-1', '  shape-1', 'background'], 'members under their group, one deeper');
  assert.deepEqual(unitOf(doc, 'group-1'), [1, 3]);
  assert.equal(doc.patches.find((p) => p.id === 'group-1').label, 'Group 1');
  // The group goes to the back with its members.
  doc = layerToEnd(doc, 'group-1', -1);
  assert.deepEqual(ids(doc), ['shape-1', 'textblock-1', 'group-1', 'background', 'adjust-1']);
  // A member steps within the group, and out of it past the top.
  doc = stepLayer(doc, 'shape-1', +1);
  assert.deepEqual(ids(doc), ['textblock-1', 'shape-1', 'group-1', 'background', 'adjust-1']);
  doc = stepLayer(doc, 'shape-1', +1);
  assert.deepEqual(ids(doc), ['textblock-1', 'group-1', 'shape-1', 'background', 'adjust-1'], 'the top member brought forward leaves the group, above it');
  assert.equal(doc.patches.find((p) => p.id === 'shape-1').parent, undefined);
  // Dropped beside a member, a layer joins the group; into the group, it is its top member.
  doc = dropLayer(doc, 'adjust-1', 'textblock-1', 'above');
  assert.deepEqual(rows(doc), ['background', 'shape-1', 'group-1', '  adjust-1', '  textblock-1']);
  doc = dropLayer(doc, 'background', 'group-1', 'into');
  assert.equal(doc.patches.find((p) => p.id === 'background').parent, 'group-1');
  // Hidden, faded and locked as a group: the renderer sees members only.
  doc = setLayer(doc, 'group-1', { visible: false });
  let painted = imagePatches(doc);
  assert.ok(!painted.some((p) => p.type === 'group'), 'a group is never painted itself');
  assert.equal(painted.find((p) => p.id === 'textblock-1').visible, false);
  assert.notEqual(painted.find((p) => p.id === 'shape-1').visible, false, 'a layer outside the group still shows');
  doc = setLayer(doc, 'group-1', { visible: true, opacity: 0.5, locked: true });
  doc = setLayer(doc, 'textblock-1', { opacity: 0.5 });
  painted = imagePatches(doc);
  assert.equal(painted.find((p) => p.id === 'textblock-1').opacity, 0.25, 'opacities multiply');
  assert.equal(painted.find((p) => p.id === 'adjust-1').opacity, 0.5);
  assert.equal(isLocked(doc, 'textblock-1'), true, 'a member of a locked group is locked');
  assert.equal(isLocked(doc, 'shape-1'), false);
  // Ungrouped: the members stay where they are.
  const before = ids(doc).filter((x) => x !== 'group-1');
  doc = ungroupLayers(doc, 'group-1');
  assert.deepEqual(ids(doc), before);
  assert.ok(doc.patches.every((p) => !p.parent));
});

test('I7: grouping a member of another group never splits that group; deleting a group deletes its members', () => {
  let doc = stack();
  doc = groupLayers(doc, ['shape-1', 'textblock-1']).doc;
  doc = groupLayers(doc, ['shape-1']).doc;
  for (const r of layerRows(doc)) {
    if (r.parent) {
      const [first, last] = unitOf(doc, r.parent);
      const at = doc.patches.findIndex((p) => p.id === r.id);
      assert.ok(at >= first && at < last, `${r.id} sits inside ${r.parent}'s run`);
    }
  }
  const g = doc.patches.find((p) => p.type === 'group' && doc.patches.some((q) => q.parent === p.id && q.id === 'textblock-1'));
  doc = removeLayers(doc, [g.id]);
  assert.ok(!ids(doc).includes('textblock-1'), 'its member went with it');
  assert.ok(ids(doc).includes('shape-1'));
});

test('I7: a duplicate lands above its original, nudged, and a group duplicates with its members', () => {
  let doc = stack();
  const one = duplicateLayers(doc, 'shape-1');
  assert.equal(one.id, 'shape-2');
  assert.deepEqual(ids(one.doc).slice(0, 3), ['background', 'shape-1', 'shape-2']);
  assert.equal(one.doc.patches[2].fx, 0.52);
  doc = groupLayers(doc, ['shape-1', 'textblock-1']).doc;
  const two = duplicateLayers(doc, 'group-1');
  const copies = two.doc.patches.filter((p) => p.parent === two.id);
  assert.equal(copies.length, 2, 'both members copied into the new group');
  assert.equal(two.doc.patches.filter((p) => p.type === 'group').length, 2);
});

test('I7: opacity, blend and lock are a layer\'s own; normal is no blend; every mode is a canvas mode', () => {
  let doc = stack();
  doc = setLayer(doc, 'shape-1', { opacity: 0.4, blend: 'multiply', locked: true, label: 'Box' });
  const p = doc.patches.find((x) => x.id === 'shape-1');
  assert.deepEqual([p.opacity, p.blend, p.locked, layerLabel(p)], [0.4, 'multiply', true, 'Box']);
  doc = setLayer(doc, 'shape-1', { blend: 'normal' });
  assert.equal(doc.patches.find((x) => x.id === 'shape-1').blend, undefined);
  assert.equal(BLEND_MODES[0], 'normal');
  assert.equal(BLEND_MODES.length, 16);
  assert.equal(layerLabel({ type: 'textblock', text: 'First line\nsecond' }), 'First line');
});

test('I15: a canvas grown or cut from each anchor puts the old corner where the anchor says', () => {
  const old = { width: 100, height: 60 };
  const next = { width: 160, height: 40 };
  const at = Object.fromEntries(ANCHORS.map((a) => [a, anchorOffset(old, next, a)]));
  assert.deepEqual(at.nw, { x: 0, y: 0 });
  assert.deepEqual(at.c, { x: -30, y: 10 });
  assert.deepEqual(at.se, { x: -60, y: 20 });
  assert.deepEqual(at.n, { x: -30, y: 0 });
  assert.deepEqual(at.e, { x: -60, y: 10 });
  assert.deepEqual(at.sw, { x: 0, y: 20 });
});

test('I15: Canvas size keeps every layer where it was in pixels; Image size scales them all', () => {
  let doc = stack();
  const px = (d, p) => ({ x: p.fx * d.size.width, y: p.fy * d.size.height, w: p.fw * d.size.width });
  const shapeBefore = px(doc, doc.patches[1]);
  const picBefore = px(doc, doc.patches[0]);
  const grown = canvasSize(doc, { width: 1200, height: 700, anchor: 'se' });
  assert.deepEqual(grown.size, { width: 1200, height: 700 });
  const shapeAfter = px(grown, grown.patches[1]);
  assert.ok(Math.abs(shapeAfter.x - (shapeBefore.x + 200)) < 1 && Math.abs(shapeAfter.y - (shapeBefore.y + 200)) < 1, 'grown from the bottom right: the room is added at the top left');
  assert.ok(Math.abs(shapeAfter.w - shapeBefore.w) < 1, 'the same size in pixels');
  const picAfter = px(grown, grown.patches[0]);
  assert.ok(Math.abs(picAfter.x - 200) < 1 && Math.abs(picAfter.w - picBefore.w) < 1, 'the picture moved by the room added, its pixels the same');
  const text = grown.patches[2];
  assert.ok(Math.abs(text.sizeFrac * 1200 - 0.05 * 1000) < 0.5, 'the type the same size in pixels');
  assert.equal(grown.patches[3].fw, undefined, 'an adjustment over everything still covers the whole canvas');
  const cut = canvasSize(doc, { width: 500, height: 500, anchor: 'nw' });
  assert.ok(Math.abs(cut.patches[1].fx * 500 - 500) < 1, 'cut from the top left: what was at 500 px is at 500 px');
  const scaled = imageSize(doc, { width: 500, height: 250 });
  assert.deepEqual(scaled.size, { width: 500, height: 250 });
  assert.deepEqual(scaled.patches, doc.patches, 'fractions unchanged, so everything scales with the canvas');
});

test('I15: a mask moves with its layer when the canvas changes', () => {
  let doc = stack();
  doc = setLayer(doc, 'shape-1', { mask: { shape: 'rect', fx: 0.5, fy: 0.5, fw: 0.1, fh: 0.2 } });
  const moved = reframe(doc, { x: 100, y: 50, width: 800, height: 400 });
  const m = moved.patches[1].mask;
  assert.ok(Math.abs(m.fx * 800 - 400) < 1 && Math.abs(m.fy * 400 - 200) < 1 && Math.abs(m.fw * 800 - 100) < 1 && Math.abs(m.fh * 400 - 100) < 1);
});

test('I15: trim finds the pixels that show, and trimming keeps them where they were', () => {
  const w = 20;
  const h = 10;
  const data = new Uint8ClampedArray(w * h * 4);
  const dot = (x, y, a = 255) => { data[(y * w + x) * 4 + 3] = a; };
  assert.equal(trimBounds(data, w, h), null, 'all clear: nothing to keep');
  dot(5, 2);
  dot(12, 7);
  dot(18, 0, 3);
  assert.deepEqual(trimBounds(data, w, h), { x: 5, y: 0, width: 14, height: 8 });
  assert.deepEqual(trimBounds(data, w, h, { alpha: 10 }), { x: 5, y: 2, width: 8, height: 6 }, 'nearly clear counts as clear above a threshold');
  let doc = emptyImage({ width: 20, height: 10 });
  doc = { ...doc, patches: [{ id: 'shape-1', type: 'shape', fx: 0.45, fy: 0.45, fw: 0.4, fh: 0.6 }] };
  const trimmed = trimTo(doc, { x: 5, y: 2, width: 8, height: 6 });
  assert.deepEqual(trimmed.size, { width: 8, height: 6 });
  assert.ok(Math.abs(trimmed.patches[0].fx * 8 - (0.45 * 20 - 5)) < 0.01);
  assert.equal(trimTo(doc, { x: 0, y: 0, width: 20, height: 10 }), doc, 'nothing to trim');
});

test('I8: a corner drag scales a box from the opposite corner; the rotation handle reads an angle', () => {
  const b = { x: 100, y: 100, w: 200, h: 100 };
  assert.deepEqual(scaleBox(b, 'se', 1.5), { x: 100, y: 100, w: 300, h: 150 });
  assert.deepEqual(scaleBox(b, 'nw', 0.5), { x: 200, y: 150, w: 100, h: 50 });
  const patch = resizeCentred({ width: 1000, height: 500 }, { id: 's', fw: 0.2, fh: 0.2 }, b, 'nw', 0.5);
  assert.deepEqual(patch, { id: 's', fx: 0.25, fy: 0.35, fw: 0.1, fh: 0.1 });
  assert.equal(rotationAt({ x: 0, y: 0 }, { x: 0, y: -10 }), 0);
  assert.equal(rotationAt({ x: 0, y: 0 }, { x: 10, y: 0 }), 90);
  assert.equal(rotationAt({ x: 0, y: 0 }, { x: 10, y: -9 }, { snap: true }), 45);
});

/* ── painted: a stub context that records what was done to it ─────────── */

function stubCtx() {
  const calls = [];
  const state = { globalAlpha: 1, globalCompositeOperation: 'source-over' };
  const record = (name) => (...args) => { calls.push({ name, args, state: { ...state } }); };
  const ctx = {
    calls,
    save: record('save'), restore: record('restore'), translate: record('translate'), rotate: record('rotate'), scale: record('scale'),
    beginPath: record('beginPath'), moveTo: record('moveTo'), lineTo: record('lineTo'), quadraticCurveTo: record('quadraticCurveTo'),
    bezierCurveTo: record('bezierCurveTo'), closePath: record('closePath'), ellipse: record('ellipse'), rect: record('rect'),
    roundRect: record('roundRect'), clip: record('clip'), fill: record('fill'), stroke: record('stroke'), arc: record('arc'),
    fillRect: record('fillRect'), clearRect: record('clearRect'), setLineDash: record('setLineDash'),
    fillText: record('fillText'), strokeText: record('strokeText'), drawImage: record('drawImage'), putImageData: record('putImageData'),
    measureText: (s) => ({ width: String(s).length * 10, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2 }),
    createLinearGradient: () => ({ addColorStop() {} }), createRadialGradient: () => ({ addColorStop() {} }), createPattern: () => null,
    setTransform: record('setTransform'), resetTransform: record('resetTransform'),
    getImageData: (x, y, w, h) => {
      calls.push({ name: 'getImageData', args: [x, y, w, h] });
      return { data: new Uint8ClampedArray(w * h * 4).fill(128), width: w, height: h };
    },
  };
  for (const prop of ['font', 'fillStyle', 'strokeStyle', 'lineWidth', 'globalAlpha', 'filter', 'globalCompositeOperation',
    'shadowColor', 'shadowBlur', 'shadowOffsetX', 'shadowOffsetY', 'textAlign', 'textBaseline', 'lineJoin', 'lineCap', 'miterLimit', 'direction',
    'letterSpacing', 'imageSmoothingEnabled', 'imageSmoothingQuality']) {
    Object.defineProperty(ctx, prop, { get: () => state[prop], set: (v) => { state[prop] = v; calls.push({ name: `set:${prop}`, args: [v] }); } });
  }
  return ctx;
}
// The adjustment layer hands its pixels back through a surface of its own.
globalThis.OffscreenCanvas ??= class { constructor(w, h) { this.width = w; this.height = h; this.ctx = stubCtx(); } getContext() { return this.ctx; } };
const canvasStub = () => { const ctx = stubCtx(); return { width: 0, height: 0, ctx, getContext: () => ctx }; };

test('I7: the canvas is painted clear, and a text layer takes its opacity and its blend mode', () => {
  const canvas = canvasStub();
  let doc = emptyImage({ width: 400, height: 200 });
  doc = { ...doc, patches: [{ id: 'text-1', type: 'text', text: 'Hi', fx: 0.5, fy: 0.5, sizeFrac: 0.1, opacity: 0.5, blend: 'screen' }] };
  const plan = buildImagePlan({ canvas, size: doc.size, options: doc.options, layerPatches: imagePatches(doc) });
  assert.deepEqual([plan.W, plan.H], [400, 200]);
  paintFrame(canvas.ctx, plan, 0);
  const names = canvas.ctx.calls.map((c) => c.name);
  assert.ok(names.includes('clearRect') && !names.includes('fillRect'), 'no ground painted under a transparent canvas');
  const text = canvas.ctx.calls.find((c) => c.name === 'fillText');
  assert.ok(text, 'the text is painted');
  assert.equal(text.state.globalAlpha, 0.5);
  assert.equal(text.state.globalCompositeOperation, 'screen');
});

test('I9: an adjustment layer reads and writes only the pixels painted before it, at its place in the stack', () => {
  const canvas = canvasStub();
  let doc = emptyImage({ width: 300, height: 100 });
  doc = { ...doc, patches: [
    { id: 'shape-1', type: 'shape', geometry: 'rect', fx: 0.25, fy: 0.5, fw: 0.3, fh: 0.5 },
    { id: 'adjust-1', type: 'adjust', preset: 'mono' },
    { id: 'shape-2', type: 'shape', geometry: 'ellipse', fx: 0.75, fy: 0.5, fw: 0.3, fh: 0.5 },
  ] };
  const plan = buildImagePlan({ canvas, size: doc.size, options: doc.options, layerPatches: imagePatches(doc) });
  paintFrame(canvas.ctx, plan, 0);
  const seq = canvas.ctx.calls.map((c) => c.name).filter((n) => ['getImageData', 'fill', 'drawImage'].includes(n));
  const read = seq.indexOf('getImageData');
  assert.ok(read > 0, 'the adjustment reads the canvas');
  assert.ok(seq.slice(0, read).includes('fill'), 'after the shape beneath it is painted');
  assert.ok(seq.slice(read).includes('fill'), 'and before the shape above it');
  const box = canvas.ctx.calls.find((c) => c.name === 'getImageData').args;
  assert.deepEqual(box, [0, 0, 300, 100], 'over the whole canvas when it has no region of its own');
  // Hidden, it reads nothing at all.
  const hidden = canvasStub();
  const plan2 = buildImagePlan({ canvas: hidden, size: doc.size, options: doc.options, layerPatches: imagePatches(setLayer(doc, 'adjust-1', { visible: false })) });
  paintFrame(hidden.ctx, plan2, 0);
  assert.ok(!hidden.ctx.calls.some((c) => c.name === 'getImageData'));
});
