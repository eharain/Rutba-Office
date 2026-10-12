// The pointer of a screen recording (docs/STUDIO.md R7) and the automatic
// zoom it drives (X5): the data made small, the pointer found between its
// samples, its layer landing over the screen in the recording's group, drawn
// where the screen draws that point (through a crop, a letterbox, an inset
// and a zoom), and the zooms found from runs of clicks, made into keys on the
// screen layer that a split, a trim and an edit of one zoom all keep right.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  compactPointer, pointerAt, ringsAt, findZooms, zoomKeys, applyAutoZoom, setZooms, updateZoom, pointerOf, POINTER_LOOK, RING_SECONDS,
} from '../packages/studio/src/pointer.js';
import { placeTakes } from '../packages/studio/src/recorder.js';
import { emptyProject, clipById, planPatches, splitAt, trimStart } from '../packages/studio/src/project.js';
import { buildPlan, paintFrame, sourcePointToFrame } from '../packages/studio/src/index.js';

const near = (a, b, msg, eps = 0.01) => assert.ok(Math.abs(a - b) < eps, `${msg}: ${a} is not ${b}`);

const SCREEN = { role: 'screen', url: 'C:/rec/Screen.webm', name: 'Screen', duration: 12, width: 1920, height: 1080, hasAudio: false };
const CAMERA = { role: 'camera', url: 'C:/rec/Camera.webm', name: 'Camera', duration: 12, width: 1280, height: 720, hasAudio: true, sound: 'voice' };

/** A pointer that rests, moves across, and clicks three times near one button, then once far away. */
function data() {
  const samples = [];
  for (let i = 0; i <= 360; i++) {
    const t = i / 30;
    const x = t < 2 ? 0.2 : t < 4 ? 0.2 + (t - 2) * 0.2 : 0.6;
    samples.push([t, x, 0.5]);
  }
  return { samples, clicks: [[4.2, 0.6, 0.5], [5.5, 0.62, 0.52], [7, 0.6, 0.48], [11, 0.1, 0.9]] };
}

test('the data: rounded, sorted, a resting pointer kept as where it stopped and where it moved again', () => {
  const d = compactPointer(data());
  assert.ok(d.samples.length < 200, `${d.samples.length} samples kept of 361`);
  assert.deepEqual(d.samples[0], [0, 0.2, 0.5]);
  assert.equal(d.clicks.length, 4);
  // Where the pointer is, between samples, and across the rest.
  near(pointerAt(d.samples, 1).x, 0.2, 'resting');
  near(pointerAt(d.samples, 3).x, 0.4, 'half way across');
  near(pointerAt(d.samples, 10).x, 0.6, 'resting again');
  assert.equal(pointerAt(d.samples, 13), null, 'after the last sample');
  assert.equal(compactPointer({ samples: [[0, 'x', 1], [1, 0.5, 0.5]], clicks: [[1]] }).samples.length, 1, 'a bad sample is dropped');
  // A ring shows for its click and fades out.
  assert.deepEqual(ringsAt(d.clicks, 4.2 + RING_SECONDS / 2).map((r) => Math.round(r.p * 1000) / 1000), [0.5]);
  assert.equal(ringsAt(d.clicks, 4.1).length, 0);
});

test('on the timeline: the pointer layer over the screen and under the camera, in the recording\'s group', () => {
  const { doc, ids } = placeTakes(emptyProject(), [{ ...SCREEN, pointer: data() }, CAMERA]);
  const pointer = doc.patches.find((p) => p.type === 'pointer');
  const screen = doc.patches.find((p) => p.take === 'screen');
  const camera = doc.patches.find((p) => p.take === 'camera');
  assert.ok(ids.includes(pointer.id));
  assert.equal(screen.lane, 0);
  assert.equal(pointer.lane, 1);
  assert.equal(camera.lane, 2);
  assert.equal(pointer.group, screen.group);
  assert.equal(pointer.screenUrl, SCREEN.url);
  assert.deepEqual(pointer.timing, screen.timing);
  assert.equal(pointer.highlight, POINTER_LOOK.highlight);
  assert.equal(pointerOf(doc, screen.id).id, pointer.id);
  // Without pointer data, no layer and the camera straight above the screen.
  const plain = placeTakes(emptyProject(), [SCREEN, CAMERA]).doc;
  assert.equal(plain.patches.some((p) => p.type === 'pointer'), false);
  assert.equal(plain.patches.find((p) => p.take === 'camera').lane, 1);
});

const el = (w, h) => ({ readyState: 4, paused: true, currentTime: 0, videoWidth: w, videoHeight: h });
function planOf(doc) {
  const canvas = canvasStub();
  const s = el(1920, 1080);
  const c = el(1280, 720);
  const videos = { [SCREEN.url]: { el: s, img: s, width: 1920, height: 1080, duration: 12 }, [CAMERA.url]: { el: c, img: c, width: 1280, height: 720, duration: 12 } };
  return { canvas, plan: buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos, assets: {} }) };
}
function canvasStub() {
  const calls = [];
  const ctx = new Proxy({ calls, measureText: (s) => ({ width: String(s).length * 10 }), createLinearGradient: () => ({ addColorStop() {} }) }, {
    get: (o, k) => (k in o ? o[k] : (...args) => { calls.push([k, ...args]); }),
    set: (o, k, v) => { o[k] = v; return true; },
  });
  return { width: 0, height: 0, getContext: () => ctx, ctx };
}

test('a source point lands where the screen draws it: whole, through a crop, in an inset that fits it, and zoomed', () => {
  const { doc } = placeTakes(emptyProject(), [SCREEN]);
  const id = doc.patches.find((p) => p.take === 'screen').id;
  const at = (d, t, x, y) => { const { plan } = planOf(d); return sourcePointToFrame(plan, plan.layers.find((l) => l.id === id), t, x, y); };
  const whole = at(doc, 1, 0.25, 0.5);
  near(whole.x, 480, 'x'); near(whole.y, 540, 'y');
  const cropped = { ...doc, patches: doc.patches.map((p) => (p.id === id ? { ...p, crop: { x: 0.5, y: 0, w: 0.5, h: 1 } } : p)) };
  assert.equal(at(cropped, 1, 0.25, 0.5), null, 'cropped away');
  const c = at(cropped, 1, 0.75, 0.5);
  near(c.x, 960, 'the crop\'s middle, letterboxed to the frame\'s middle', 0.5);
  const inset = { ...doc, patches: doc.patches.map((p) => (p.id === id ? { ...p, fx: 0.5, fy: 0, fw: 0.5, fh: 0.5, fit: 'contain' } : p)) };
  const i = at(inset, 1, 0, 0);
  near(i.x, 960, 'inset left'); near(i.y, 0, 'inset top');
  const zoomed = setZooms(doc, id, [{ start: 0, end: 6, x: 0.75, y: 0.25, level: 2, points: [] }]);
  const z = at(zoomed, 3, 0.75, 0.25);
  near(z.x, 960, 'the zoom\'s centre in the frame\'s middle', 1);
  near(z.y, 540, 'and down the middle', 1);
  assert.equal(at(zoomed, 3, 0.1, 0.9), null, 'zoomed out of sight');
});

test('drawn: a highlight round the pointer and a ring for a click, on the screen where the pointer was', () => {
  const { doc } = placeTakes(emptyProject(), [{ ...SCREEN, pointer: data() }]);
  const { canvas, plan } = planOf(doc);
  paintFrame(canvas.ctx, plan, 4.3);
  const arcs = canvas.ctx.calls.filter((k) => k[0] === 'arc');
  assert.equal(arcs.length, 2, 'a ring and the highlight');
  near(arcs[0][1], 0.6 * 1920, 'the ring on the click', 1);
  near(arcs[1][1], 0.6 * 1920, 'the highlight on the pointer', 1);
  // Switched off after, nothing is drawn.
  const off = { ...doc, patches: doc.patches.map((p) => (p.type === 'pointer' ? { ...p, highlight: false, rings: false } : p)) };
  const second = planOf(off);
  paintFrame(second.canvas.ctx, second.plan, 4.3);
  assert.equal(second.canvas.ctx.calls.filter((k) => k[0] === 'arc').length, 0);
});

test('zooms: a run of clicks close together is one zoom on their middle, a lone click far off another, each held a while', () => {
  const zooms = findZooms(data().clicks, { length: 12 });
  assert.equal(zooms.length, 2);
  near(zooms[0].start, 3.4, 'a little before the first click');
  near(zooms[0].end, 8.6, 'a while after the last');
  near(zooms[0].x, 0.607, 'the middle of the run');
  assert.equal(zooms[0].points.length, 3);
  near(zooms[1].end, 12, 'held to the layer\'s end');
  // Keys: zoom 1, up to the level, held, down again; the pan follows each click.
  const keys = zoomKeys(zooms);
  assert.deepEqual(keys.zoom.slice(0, 4).map((k) => k.v), [1, 1.8, 1.8, 1]);
  assert.equal(keys.zoom[1].ease, 'in-out');
  assert.ok(keys.panX.some((k) => k.t === 5.5 && k.v === 0.62), 'the pan goes to the second click');
  // Within a crop the pan is the crop's own fraction.
  const cropped = zoomKeys([{ start: 0, end: 4, x: 0.75, y: 0.5, level: 2 }], { crop: { x: 0.5, y: 0, w: 0.5, h: 1 } });
  assert.equal(cropped.panX[0].v, 0.5);
  assert.equal(zoomKeys([]), null);
});

test('the automatic zoom on a screen layer: kept as a list and as keys, one zoom changed or taken away, and a split and a trim keep the keys in the clip\'s own time', () => {
  const { doc } = placeTakes(emptyProject(), [{ ...SCREEN, pointer: data() }]);
  const id = doc.patches.find((p) => p.take === 'screen').id;
  const zoomed = applyAutoZoom(doc, id);
  const screen = clipById(zoomed, id);
  assert.equal(screen.autoZoom.length, 2);
  assert.ok(screen.keys.zoom.length >= 8);
  const plain = placeTakes(emptyProject(), [SCREEN]).doc;
  assert.equal(applyAutoZoom(plain, plain.patches.find((p) => p.take === 'screen').id), plain, 'no clicks, nothing to zoom on');
  // Stronger, then the second taken away.
  const stronger = updateZoom(zoomed, id, 0, { level: 2.5 });
  assert.ok(clipById(stronger, id).keys.zoom.some((k) => k.v === 2.5));
  const one = updateZoom(stronger, id, 1, null);
  assert.equal(clipById(one, id).autoZoom.length, 1);
  const none = setZooms(one, id, []);
  assert.equal(clipById(none, id).keys, undefined);
  assert.equal(clipById(none, id).autoZoom, undefined);
  // Split at 6 s: the second half's zoom starts at its own 0, held mid-zoom.
  const split = splitAt(zoomed, 6, [id]);
  const second = clipById(split.doc, split.ids.find((x) => clipById(split.doc, x).type === 'video'));
  assert.equal(second.keys.zoom[0].t, 0);
  near(second.keys.zoom[0].v, 1.8, 'still zoomed where it was cut');
  near(second.autoZoom[0].end, 2.6, 'the zoom ends at its own time less the cut');
  // The pointer split with it keeps drawing from the right place.
  const pointers = split.doc.patches.filter((p) => p.type === 'pointer');
  assert.equal(pointers.length, 2);
  assert.equal(pointers.find((p) => p.timing.start === 6).offset, 6);
  // A start trimmed by a second moves the keys back a second.
  const trimmed = trimStart(zoomed, id, 1);
  near(clipById(trimmed, id).keys.zoom.find((k) => k.v > 1).t, screen.keys.zoom.find((k) => k.v > 1).t - 1, 'a second earlier');
});
