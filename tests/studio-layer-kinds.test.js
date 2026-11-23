// The layer kinds Video's library adds beside media, titles, shapes, icons and
// word art (D-2026-11-20-19): a paragraph or a list, a chart, a colour
// adjustment, a QR code, the shade along the bottom, a progress bar, the fade
// at the two ends and an outro card. Each is a clip on a track of the
// timeline, and the engine paints it inside its own window.
import test from 'node:test';
import assert from 'node:assert/strict';
import '../packages/studio/src/toolset.js';
import { buildPlan, paintFrame } from '../packages/studio/src/index.js';
import { emptyProject, addMedia, addGraphic, clipById, planPatches, projectDuration, GRAPHIC_TYPES, OVERLAYS, OUTRO_SECONDS } from '../packages/studio/src/project.js';

const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 20, hasAudio: false, width: 320, height: 240 };
const film = () => addMedia(emptyProject(), A, { at: 0 }).doc;

/** Enough of a 2D canvas for a plan to be built and painted in Node; every call is kept, with the fill it was made in. */
function canvasStub() {
  const calls = [];
  const state = { calls, measureText: (s) => ({ width: String(s).length * 10 }), createLinearGradient: () => ({ addColorStop() {} }), fillStyle: '' };
  const ctx = new Proxy(state, {
    get: (o, k) => (k in o ? o[k] : (...args) => { calls.push([k, ...args, o.fillStyle]); }),
    set: (o, k, v) => { o[k] = v; return true; },
  });
  return { width: 0, height: 0, getContext: () => ctx, ctx };
}

const planOf = (doc) => {
  const canvas = canvasStub();
  const el = { readyState: 4, paused: true, currentTime: 0 };
  const plan = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos: { [A.url]: { el, img: el, width: 320, height: 240, duration: 20 } }, assets: {} });
  return { plan, canvas };
};

const SAMPLE = {
  textblock: { type: 'textblock', list: 'bullet', text: 'One\nTwo', fx: 0.5, fy: 0.55, fw: 0.66, sizeFrac: 0.035 },
  chart: { type: 'chart', chart: 'column', labels: ['a', 'b'], values: [1, 2], fx: 0.5, fy: 0.5, fw: 0.72, fh: 0.34 },
  adjust: { type: 'adjust', preset: 'warm' },
  qr: OVERLAYS.qr,
  gradient: OVERLAYS.gradient,
  progress: OVERLAYS.progress,
  edges: OVERLAYS.edges,
  outro: OVERLAYS.outro,
};

test('every kind the library adds becomes a clip on a free track above the pictures', () => {
  assert.deepEqual(GRAPHIC_TYPES.filter((k) => !['icon', 'shape', 'wordart'].includes(k)).sort(), Object.keys(SAMPLE).sort());
  let doc = film();
  for (const [type, patch] of Object.entries(SAMPLE)) {
    const r = addGraphic(doc, patch, { at: 2 });
    assert.equal(r.ids.length, 1, type);
    const clip = clipById(r.doc, r.ids[0]);
    assert.equal(clip.type, type);
    assert.ok(clip.lane >= 1, `${type} is above the pictures`);
    assert.ok(clip.timing.end > clip.timing.start, type);
    assert.equal(clip.enter.kind, 'none', type);
    assert.equal(clip.exit.kind, 'none', type);
    // The same time on the same track is never given twice.
    const clash = r.doc.patches.filter((p) => p.timing && p.type !== 'sound' && p.lane === clip.lane && p.id !== clip.id && p.timing.start < clip.timing.end && p.timing.end > clip.timing.start);
    assert.deepEqual(clash, [], `${type} shares its track with nothing`);
    doc = r.doc;
  }
  assert.deepEqual(addGraphic(doc, { type: 'caption' }).ids, [], 'a kind the library does not add is refused');
});

test('a paragraph, a chart, an adjustment and a QR code are added at the playhead for a still\'s length', () => {
  for (const type of ['textblock', 'chart', 'adjust', 'qr', 'gradient']) {
    const r = addGraphic(film(), SAMPLE[type], { at: 3 });
    assert.deepEqual(clipById(r.doc, r.ids[0]).timing, { start: 3, end: 8 }, type);
  }
  const typed = addGraphic(film(), SAMPLE.textblock, { at: 0 });
  assert.equal(clipById(typed.doc, typed.ids[0]).keepEmpty, true, 'a paragraph emptied of its words stays a clip');
});

test('a progress bar and the fade at the two ends are added over the whole film, an outro card after it', () => {
  for (const type of ['progress', 'edges']) {
    const r = addGraphic(film(), SAMPLE[type], { at: 7 });
    assert.deepEqual(clipById(r.doc, r.ids[0]).timing, { start: 0, end: 20 }, type);
  }
  const r = addGraphic(film(), SAMPLE.outro, { at: 7 });
  assert.deepEqual(clipById(r.doc, r.ids[0]).timing, { start: 20, end: 20 + OUTRO_SECONDS });
  assert.equal(projectDuration(r.doc), 20 + OUTRO_SECONDS, 'the film is longer by its outro');
  // On an empty project each still has a length to be seen for.
  const empty = addGraphic(emptyProject(), SAMPLE.progress, { at: 0 });
  assert.deepEqual(clipById(empty.doc, empty.ids[0]).timing, { start: 0, end: 5 });
});

test('the engine compiles every kind into the plan with its window', () => {
  let doc = film();
  const ids = {};
  for (const [type, patch] of Object.entries(SAMPLE)) {
    const r = addGraphic(doc, patch, { at: 2 });
    doc = r.doc;
    [ids[type]] = r.ids;
  }
  const { plan } = planOf(doc);
  for (const [type, id] of Object.entries(ids)) {
    const layer = plan.layers.find((l) => l.id === id);
    assert.ok(layer, `${type} is in the plan`);
    assert.equal(layer.type, type);
    assert.deepEqual(layer.timing, clipById(doc, id).timing, type);
    // A QR code is added with nothing to scan, so none is drawn until an address is typed.
    if (type === 'qr') assert.equal(layer.visible, false, 'a QR code with no address draws nothing');
    else assert.notEqual(layer.visible, false, type);
  }
  assert.equal(OVERLAYS.qr.data, '', 'no made-up address is put in the film');
  assert.equal(plan.layers.find((l) => l.id === ids.qr).data, '');
  // Once an address is typed the code is drawn, with that address in it.
  const typed = planOf({ ...doc, patches: doc.patches.map((p) => (p.id === ids.qr ? { ...p, data: 'https://rutba.io/menu' } : p)) }).plan.layers.find((l) => l.id === ids.qr);
  assert.notEqual(typed.visible, false);
  assert.equal(typed.data, 'https://rutba.io/menu');
  assert.equal(plan.layers.find((l) => l.id === ids.textblock).list, 'bullet');
  assert.equal(plan.layers.find((l) => l.id === ids.adjust).preset, 'warm');
});

test('a progress bar fills over its own clip, and shows nowhere else', () => {
  let doc = film();
  const r = addGraphic(doc, SAMPLE.progress);
  doc = { ...r.doc, patches: r.doc.patches.map((p) => (p.id === r.ids[0] ? { ...p, timing: { start: 4, end: 12 } } : p)) };
  const { plan, canvas } = planOf(doc);
  const bars = (t) => {
    canvas.ctx.calls.length = 0;
    paintFrame(canvas.ctx, plan, t);
    // The bar is drawn in the theme's accent along the bottom edge.
    return canvas.ctx.calls.filter((c) => c[0] === 'fillRect' && c[2] > plan.H * 0.9 && c[5] === plan.theme.accent).map((c) => c[3]);
  };
  assert.deepEqual(bars(2), [], 'before its clip there is no bar');
  assert.deepEqual(bars(4), [0]);
  assert.deepEqual(bars(8), [plan.W / 2], 'half way through its clip it is half the frame wide');
  assert.deepEqual(bars(12), [plan.W]);
  assert.deepEqual(bars(15), [], 'after its clip there is no bar');
});

test('the fade at the two ends dips at its own clip\'s ends, for the seconds it is given', () => {
  let doc = film();
  const r = addGraphic(doc, { ...SAMPLE.edges, seconds: 2 });
  doc = { ...r.doc, patches: r.doc.patches.map((p) => (p.id === r.ids[0] ? { ...p, timing: { start: 5, end: 15 } } : p)) };
  const { plan, canvas } = planOf(doc);
  const dip = (t) => {
    canvas.ctx.calls.length = 0;
    paintFrame(canvas.ctx, plan, t);
    const fill = canvas.ctx.calls.filter((c) => c[0] === 'fillRect' && /^rgba\(0,0,0,/.test(String(c[5])));
    return fill.length ? Number(String(fill[fill.length - 1][5]).match(/,([\d.]+)\)$/)[1]) : 0;
  };
  assert.equal(dip(5), 1, 'black where its clip starts');
  assert.equal(dip(6), 0.5);
  assert.equal(dip(10), 0, 'clear in the middle');
  assert.equal(dip(14), 0.5);
  assert.equal(dip(15), 1, 'black where its clip ends');
  assert.equal(dip(2), 0, 'nothing before its clip');
});

test('an outro card covers the frame from where its clip starts and says its words', () => {
  let doc = film();
  const r = addGraphic(doc, { ...SAMPLE.outro, text: 'Thank you' });
  doc = r.doc;
  const { plan, canvas } = planOf(doc);
  canvas.ctx.calls.length = 0;
  paintFrame(canvas.ctx, plan, 19);
  assert.ok(!canvas.ctx.calls.some((c) => c[0] === 'fillText' && c[1] === 'Thank you'), 'not before its clip');
  canvas.ctx.calls.length = 0;
  paintFrame(canvas.ctx, plan, 22);
  assert.ok(canvas.ctx.calls.some((c) => c[0] === 'fillText' && c[1] === 'Thank you'));
  assert.ok(canvas.ctx.calls.some((c) => c[0] === 'fillRect' && c[3] === plan.W && c[4] === plan.H && c[5] === plan.theme.bg));
});

test('the slideshow\'s own progress bar and fades, which have no window, run over the whole video as before', () => {
  const canvas = canvasStub();
  const plan = buildPlan({ canvas, images: [], body: 'Words', options: { showProgress: true, edgeFadeSeconds: 0.5 }, videos: {}, assets: {} });
  const progress = plan.layers.find((l) => l.type === 'progress');
  assert.equal(progress.timing, null);
  canvas.ctx.calls.length = 0;
  paintFrame(canvas.ctx, plan, plan.duration / 2);
  const bar = canvas.ctx.calls.filter((c) => c[0] === 'fillRect' && c[5] === plan.theme.accent && c[2] > plan.H * 0.9);
  assert.equal(bar.length, 1);
  assert.equal(bar[0][3], plan.W / 2);
  canvas.ctx.calls.length = 0;
  paintFrame(canvas.ctx, plan, 0.25);
  assert.ok(canvas.ctx.calls.some((c) => c[0] === 'fillRect' && c[5] === 'rgba(0,0,0,0.5)'), 'half way up from black a quarter second in');
});
