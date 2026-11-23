// What a click on Video's stage picks (apps/desktop/renderer/apps/video/stage-pick.js):
// what stands on the frame at the playhead, an end card once it has come in,
// and a QR code that waits for its address, which the engine does not draw.
import test from 'node:test';
import assert from 'node:assert/strict';
import '../packages/studio/src/toolset.js';
import { buildPlan } from '../packages/studio/src/index.js';
import { emptyProject, addMedia, addGraphic, clipById, planPatches, OVERLAYS } from '../packages/studio/src/project.js';
import {
  BOXED_KINDS, OUTRO_SEEN_AFTER, fillsFrame, coversFrame, emptyQr, liveLayers, boundsOf, pickAt, emptyQrBoxes,
} from '../apps/desktop/renderer/apps/video/stage-pick.js';

const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 20, hasAudio: false, width: 320, height: 240 };

/** Enough of a 2D canvas for a plan to be built and measured in Node. */
function canvasStub() {
  const state = { measureText: (s) => ({ width: String(s).length * 10 }), createLinearGradient: () => ({ addColorStop() {} }), fillStyle: '' };
  const ctx = new Proxy(state, { get: (o, k) => (k in o ? o[k] : () => {}), set: (o, k, v) => { o[k] = v; return true; } });
  return { width: 0, height: 0, getContext: () => ctx, ctx };
}

function planOf(doc) {
  const canvas = canvasStub();
  const el = { readyState: 4, paused: true, currentTime: 0 };
  const plan = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos: { [A.url]: { el, img: el, width: 320, height: 240, duration: 20 } }, assets: {} });
  return { plan, ctx: canvas.ctx };
}

/** A film of twenty seconds with an end card after it and a QR code at two seconds. */
function film({ address = '' } = {}) {
  let doc = addMedia(emptyProject(), A, { at: 0 }).doc;
  const qr = addGraphic(doc, { ...OVERLAYS.qr, data: address }, { at: 2 });
  doc = qr.doc;
  const outro = addGraphic(doc, OVERLAYS.outro, { at: 0 });
  doc = outro.doc;
  return { doc, qr: qr.ids[0], outro: outro.ids[0], clip: doc.patches.find((p) => p.type === 'video').id };
}

test('an end card is picked on the frame once it has come in, and wears the whole frame', () => {
  const f = film();
  const { plan, ctx } = planOf(f.doc);
  const start = clipById(f.doc, f.outro).timing.start;
  assert.equal(start, 20, 'the end card stands after the film');
  const card = plan.layers.find((l) => l.id === f.outro);
  assert.equal(coversFrame(card), true);
  assert.equal(fillsFrame(card), false, 'it has no corner grips to make it a picture in picture');
  assert.deepEqual(boundsOf(ctx, plan, card), { x: 0, y: 0, w: plan.W, h: plan.H });

  // At its very start it has not come in: what is under it is what a click meets.
  const early = liveLayers(plan, start + 0.1);
  assert.ok(!early.some((l) => l.id === f.outro), 'still fading in');
  // From half of its fade it is the card that is clicked, anywhere on the frame.
  const at = start + OUTRO_SEEN_AFTER;
  const live = liveLayers(plan, at);
  assert.ok(live.some((l) => l.id === f.outro));
  for (const [x, y] of [[5, 5], [plan.W / 2, plan.H / 2], [plan.W - 5, plan.H - 5]]) {
    assert.equal(pickAt(ctx, plan, live, at, x, y)?.id, f.outro, `${x},${y}`);
  }
  // Before it starts and after it ends it is not there to pick.
  assert.ok(!liveLayers(plan, 10).some((l) => l.id === f.outro));
  assert.ok(!liveLayers(plan, clipById(f.doc, f.outro).timing.end + 1).some((l) => l.id === f.outro));
});

test('an end card on a locked track, or one hidden, is not picked on the frame', () => {
  const f = film();
  const { plan } = planOf(f.doc);
  const at = 20 + 1;
  assert.ok(!liveLayers(plan, at, (id) => id === f.outro).some((l) => l.id === f.outro), 'locked');
  const hidden = { ...plan, layers: plan.layers.map((l) => (l.id === f.outro ? { ...l, visible: false } : l)) };
  assert.ok(!liveLayers(hidden, at).some((l) => l.id === f.outro), 'hidden');
});

test('what stands above an end card is picked before it', () => {
  const f = film();
  const added = addGraphic(f.doc, { type: 'chart', chart: 'column', labels: ['a', 'b'], values: [1, 2], fx: 0.5, fy: 0.5, fw: 0.4, fh: 0.3 }, { at: 21 });
  const { plan, ctx } = planOf(added.doc);
  const chart = plan.layers.find((l) => l.id === added.ids[0]);
  const card = plan.layers.find((l) => l.id === f.outro);
  assert.ok(chart && (chart.z || 0) > (card.z || 0), 'the chart was added above the card');
  const live = liveLayers(plan, 21.5);
  const b = boundsOf(ctx, plan, chart);
  assert.equal(pickAt(ctx, plan, live, 21.5, b.x + b.w / 2, b.y + b.h / 2)?.id, chart.id);
  // Away from the chart, the card.
  assert.equal(pickAt(ctx, plan, live, 21.5, 10, 10)?.id, f.outro);
});

test('a QR code that waits for its address draws nothing, and is shown, picked and measured where it stands', () => {
  const f = film();
  const { plan, ctx } = planOf(f.doc);
  const code = plan.layers.find((l) => l.id === f.qr);
  assert.equal(code.visible, false, 'the engine draws no code with nothing to scan');
  assert.equal(emptyQr(code), true);
  const live = liveLayers(plan, 3);
  assert.ok(live.some((l) => l.id === f.qr), 'it is picked on the frame all the same');
  const boxes = emptyQrBoxes(ctx, plan, 3);
  assert.equal(boxes.length, 1);
  const [box] = boxes;
  assert.equal(box.id, f.qr);
  assert.equal(box.x, Math.round(plan.W * OVERLAYS.qr.fx));
  assert.equal(box.y, Math.round(plan.H * OVERLAYS.qr.fy));
  assert.equal(box.w, Math.round(plan.W * OVERLAYS.qr.fw));
  assert.equal(box.h, box.w, 'a QR code is square');
  assert.equal(pickAt(ctx, plan, live, 3, box.x + box.w / 2, box.y + box.h / 2)?.id, f.qr);
  // Outside its own time it is neither shown nor picked.
  assert.deepEqual(emptyQrBoxes(ctx, plan, 15), []);
  // On a locked track it is left alone.
  assert.deepEqual(emptyQrBoxes(ctx, plan, 3, (id) => id === f.qr), []);
});

test('a QR code with an address is drawn by the engine and has no placeholder', () => {
  const f = film({ address: 'https://rutba.io/menu' });
  const { plan, ctx } = planOf(f.doc);
  const code = plan.layers.find((l) => l.id === f.qr);
  assert.notEqual(code.visible, false);
  assert.equal(emptyQr(code), false);
  assert.deepEqual(emptyQrBoxes(ctx, plan, 3), []);
  assert.ok(liveLayers(plan, 3).some((l) => l.id === f.qr));
});

test('what lies over the frame and is seen through stays picked on the timeline', () => {
  let { doc } = film();
  for (const key of ['gradient', 'progress', 'edges']) doc = addGraphic(doc, OVERLAYS[key], { at: 2 }).doc;
  doc = addGraphic(doc, { type: 'adjust', preset: 'warm' }, { at: 2 }).doc;
  const { plan, ctx } = planOf(doc);
  const live = liveLayers(plan, 3);
  for (const type of ['gradient', 'progress', 'edges', 'adjust']) assert.ok(!live.some((l) => l.type === type), type);
  // A click on the frame away from the QR code finds the clip under them.
  assert.equal(pickAt(ctx, plan, live, 3, 20, plan.H - 20)?.type, 'video');
  assert.deepEqual(BOXED_KINDS, ['textblock', 'chart', 'qr']);
});
