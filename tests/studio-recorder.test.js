// The screen recorder's arithmetic (docs/STUDIO.md section 8, Phase 2): the
// settings and what they cost the disk, the camera's inset and the corner it
// settles in, a drawn region as the screen layer's crop, the clock that
// pauses, the writer that streams a recording to its file a slice at a time,
// and the takes landing on Video's timeline: the screen on the main lane,
// the camera above it as a mirrored inset cut to its shape, each file's
// sound on a lane of its own, all in one group; a recovered take settled to
// the length its file really has.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  RECORDER_DEFAULTS, normaliseSettings, capSize, videoBitrate, bytesPerSecond, minutesLeft, diskState,
  insetRect, cornerOf, regionCrop, createClock, createChunkWriter, takeFileName, placeTakes, settleTakes,
  INSET_SIZES, SLICE,
} from '../packages/studio/src/recorder.js';
import { emptyProject, addMedia, clipById, laneClips, planPatches, projectMedia, splitAt, partnersOf } from '../packages/studio/src/project.js';
import { buildPlan, paintFrame } from '../packages/studio/src/index.js';

const near = (a, b, msg, eps = 0.002) => assert.ok(Math.abs(a - b) < eps, `${msg}: ${a} is not ${b}`);

test('settings: anything stored comes back as settings the recorder can use, the defaults filling the rest', () => {
  assert.deepEqual(normaliseSettings(null), { ...RECORDER_DEFAULTS });
  const s = normaliseSettings({ fps: 60, cap: 'big', quality: 'high', countdown: 7, corner: 'tl', shape: 'star', camera: false, system: 1 });
  assert.equal(s.fps, 60);
  assert.equal(s.cap, 1080, 'an unknown cap falls back');
  assert.equal(s.quality, 'high');
  assert.equal(s.countdown, 3, 'a countdown not offered falls back to three seconds');
  assert.equal(s.corner, 'tl');
  assert.equal(s.shape, 'circle');
  assert.equal(s.camera, false);
  assert.equal(s.system, true);
  assert.equal(normaliseSettings({ fps: 24 }).fps, 30, 'only 15, 30 and 60 frames a second');
});

test('quality: a screen is held to the cap on its shorter side, and the defaults keep 1080p at 30 fps under 8 Mbps', () => {
  assert.deepEqual(capSize(2560, 1440, 1080), { width: 1920, height: 1080 });
  assert.deepEqual(capSize(1920, 1080, 1080), { width: 1920, height: 1080 });
  assert.deepEqual(capSize(1080, 1920, 720), { width: 720, height: 1280 }, 'a portrait screen by its width');
  assert.deepEqual(capSize(3840, 2160, 0), { width: 3840, height: 2160 }, 'no cap keeps the source');
  assert.deepEqual(capSize(1367, 769, 1080), { width: 1368, height: 770 }, 'even numbers, as encoders want');
  const d = normaliseSettings({});
  const bits = videoBitrate({ width: 1920, height: 1080, fps: d.fps, quality: d.quality });
  assert.ok(bits < 8e6 && bits > 4e6, `1080p30 at the default quality is ${bits}`);
  assert.ok(videoBitrate({ width: 1920, height: 1080, fps: 60 }) > bits, 'more frames, more bits');
  assert.ok(videoBitrate({ width: 1920, height: 1080, fps: 30, quality: 'small' }) < bits);
  assert.ok(videoBitrate({ width: 10, height: 10, fps: 15 }) >= 250e3, 'a floor for a tiny window');
});

test('disk: what a recording costs a second, the minutes the free space holds, a warning under five minutes and a stop under twenty seconds', () => {
  const screenOnly = bytesPerSecond({ camera: false, mic: false }, { screen: { width: 1920, height: 1080 } });
  const withCamera = bytesPerSecond({ camera: true, mic: true }, { screen: { width: 1920, height: 1080 }, camera: { width: 1280, height: 720 } });
  assert.ok(withCamera > screenOnly + 300000, 'the camera and the voice add to it');
  near(screenOnly, (6220800 / 8) * 1.03, 'the screen alone', 2);
  const rate = 1e6;
  near(minutesLeft(600e6, rate), 10, 'ten minutes');
  assert.equal(diskState(600e6, rate).level, 'ok');
  assert.equal(diskState(240e6, rate).level, 'low');
  assert.equal(diskState(10e6, rate).level, 'stop');
  assert.equal(minutesLeft(5, 0), Infinity);
});

test('the camera inset: a corner of the frame, a circle as tall as it is wide on screen, and the corner it settles in after a drag', () => {
  const br = insetRect({ corner: 'br', size: 'medium', shape: 'circle', W: 1920, H: 1080 });
  near(br.fw, INSET_SIZES.medium, 'width');
  near(br.fh * 1080, br.fw * 1920, 'a circle is square in pixels', 1);
  near(br.fx + br.fw, 0.975, 'right edge inside the margin');
  near((br.fy + br.fh) * 1080, 1080 - 0.025 * 1920, 'bottom edge inside the same margin in pixels', 1);
  const tl = insetRect({ corner: 'tl', size: 'small', shape: 'rounded', W: 1920, H: 1080, aspect: 9 / 16 });
  near(tl.fx, 0.025, 'left');
  near(tl.fh * 1080, tl.fw * 1920 * (9 / 16), 'a rectangle takes the camera\'s shape', 1);
  assert.equal(cornerOf(br), 'br');
  assert.equal(cornerOf(tl), 'tl');
  assert.equal(cornerOf({ fx: 0.6, fy: 0.1, fw: 0.2, fh: 0.2 }), 'tr');
  assert.equal(cornerOf({ fx: 0.1, fy: 0.55, fw: 0.2, fh: 0.2 }), 'bl');
});

test('a region drawn on a screen is the crop its layer keeps; drawn backwards it is the same, too small or the whole screen it is none', () => {
  assert.deepEqual(regionCrop({ x: 480, y: 270, w: 960, h: 540 }, { width: 1920, height: 1080 }), { x: 0.25, y: 0.25, w: 0.5, h: 0.5 });
  assert.deepEqual(regionCrop({ x: 1440, y: 810, w: -960, h: -540 }, { width: 1920, height: 1080 }), { x: 0.25, y: 0.25, w: 0.5, h: 0.5 });
  assert.deepEqual(regionCrop({ x: -100, y: 0, w: 1060, h: 540 }, { width: 1920, height: 1080 }), { x: 0, y: 0, w: 0.5, h: 0.5 }, 'held inside the screen');
  assert.equal(regionCrop({ x: 10, y: 10, w: 5, h: 300 }, { width: 1920, height: 1080 }), null);
  assert.equal(regionCrop({ x: 0, y: 0, w: 1, h: 1 }), null);
});

test('the clock leaves the time spent paused out', () => {
  const c = createClock();
  assert.equal(c.elapsed(5), 0);
  c.start(1000);
  assert.equal(c.elapsed(3000), 2000);
  c.pause(3000);
  assert.equal(c.paused, true);
  assert.equal(c.elapsed(9000), 2000, 'still while paused');
  c.resume(9000);
  assert.equal(c.elapsed(10000), 3000);
  c.pause(10000);
  c.resume(12000);
  assert.equal(c.elapsed(12500), 3500);
});

/** A Blob-like piece of `n` bytes, each its own position mod 256. */
const piece = (n, from = 0) => ({
  size: n,
  slice: (a, b) => ({ arrayBuffer: async () => Uint8Array.from({ length: b - a }, (_, i) => (from + a + i) % 256).buffer }),
});

test('a recording streams to its file in order, a slice a message at most, and stops writing at a failure', async () => {
  const got = [];
  let slow = true;
  const w = createChunkWriter({
    slice: 1000,
    append: async (bytes) => {
      // The first write takes longest: order must still hold.
      if (slow) { slow = false; await new Promise((r) => setTimeout(r, 20)); }
      got.push(bytes);
    },
  });
  w.push(piece(2500, 0));
  w.push(piece(0));
  w.push(piece(700, 2500));
  const r = await w.done();
  assert.equal(r.written, 3200);
  assert.equal(r.failed, null);
  assert.deepEqual(got.map((b) => b.length), [1000, 1000, 500, 700]);
  const all = Buffer.concat(got.map((b) => Buffer.from(b)));
  assert.ok(all.every((v, i) => v === i % 256), 'every byte in its place');
  assert.equal(SLICE, 4 * 1024 * 1024);

  let calls = 0;
  const bad = createChunkWriter({ slice: 100, append: async () => { calls += 1; if (calls === 2) throw new Error('disk full'); } });
  bad.push(piece(300));
  bad.push(piece(300));
  const r2 = await bad.done();
  assert.equal(r2.failed.message, 'disk full');
  assert.equal(calls, 2, 'nothing written after it');
});

test('a take is named for what it is and when it began, with nothing a file name cannot hold', () => {
  const at = new Date(2026, 9, 12, 9, 5, 7);
  assert.equal(takeFileName('Screen', at, 'webm'), 'Screen 2026-10-12 09.05.07.webm');
  assert.equal(takeFileName('A/B: "c"', at, '.mp4'), 'A B c 2026-10-12 09.05.07.mp4');
});

const SCREEN = { role: 'screen', url: 'D:/rec/screen.webm', name: 'Screen', duration: 12, width: 1920, height: 1080, hasAudio: true, size: 100 };
const CAMERA = { role: 'camera', url: 'D:/rec/camera.webm', name: 'Camera', duration: 12.02, width: 1280, height: 720, hasAudio: true, shape: 'circle', mirror: true, inset: insetRect({ corner: 'br' }) };

test('takes land on the timeline: the screen on the main lane, the camera above it as an inset, each sound on its own lane, all in one group', () => {
  const { doc, ids, group } = placeTakes(emptyProject(), [CAMERA, SCREEN]);
  assert.equal(ids.length, 4);
  const screen = doc.patches.find((p) => p.take === 'screen');
  const camera = doc.patches.find((p) => p.take === 'camera');
  const computer = doc.patches.find((p) => p.take === 'computer');
  const voice = doc.patches.find((p) => p.take === 'voice');
  assert.equal(screen.lane, 0);
  assert.equal(screen.fit, 'contain');
  assert.equal(camera.lane, 1);
  assert.equal(camera.insetShape, 'circle');
  assert.equal(camera.flipX, true);
  near(camera.fw, CAMERA.inset.fw, 'inset size');
  assert.deepEqual([computer.lane, voice.lane], [0, 1], 'each sound on a lane of its own, from the first');
  assert.ok(doc.patches.every((p) => p.group === group), 'one group, one clock');
  assert.equal(screen.timing.start, 0);
  near(camera.timing.end, 12.02, 'each take its own length');
  assert.equal(doc.sources[SCREEN.url].hasAudio, true);
  assert.equal(doc.recording, undefined);
  // A split cuts all four, so they stay together.
  const split = splitAt(doc, 6, []).doc;
  assert.equal(split.patches.length, 8);
  // The renderer gets the screen under the camera.
  const order = planPatches(doc).filter((p) => p.type === 'video').map((p) => p.take);
  assert.deepEqual(order, ['screen', 'camera']);
  assert.deepEqual(projectMedia(doc).videos.map((v) => v.url).sort(), [CAMERA.url, SCREEN.url].sort());
});

test('a region is the screen\'s crop; takes go after what is there, onto free lanes; without sound or a camera they still land', () => {
  let doc = emptyProject();
  doc = addMedia(doc, { url: 'D:/a.webm', name: 'a', kind: 'video', duration: 5, hasAudio: true }).doc;
  const crop = regionCrop({ x: 0, y: 0, w: 960, h: 540 }, { width: 1920, height: 1080 });
  const r = placeTakes(doc, [{ ...SCREEN, crop, hasAudio: false }]);
  const screen = r.doc.patches.find((p) => p.take === 'screen');
  assert.deepEqual(screen.crop, { x: 0, y: 0, w: 0.5, h: 0.5 });
  assert.equal(screen.timing.start, 5, 'after the clip already there');
  assert.equal(screen.lane, 0);
  assert.equal(r.ids.length, 1);
  assert.ok(screen.group, 'a group even alone');
  // At the playhead over an existing clip: the screen goes up a lane, the camera above it.
  const r2 = placeTakes(doc, [SCREEN, CAMERA], { at: 1 });
  assert.equal(r2.doc.patches.find((p) => p.take === 'screen').lane, 1);
  assert.equal(r2.doc.patches.find((p) => p.take === 'camera').lane, 2);
  const sounds = r2.doc.patches.filter((p) => p.type === 'sound' && p.take).map((p) => p.lane).sort();
  assert.deepEqual(sounds, [1, 2], 'the sound lanes under the clip\'s own are taken');
  assert.equal(placeTakes(doc, []).doc, doc);
});

test('a recording copied while it ran comes back unfinished, and settles to the length its files really have', () => {
  const { doc } = placeTakes(emptyProject(), [{ ...SCREEN, duration: 4 }, { ...CAMERA, duration: 4 }], { unfinished: true });
  assert.equal(doc.recording.unfinished, true);
  const settled = settleTakes(doc, { [SCREEN.url]: 9.5, [CAMERA.url]: 9.48 });
  assert.equal(settled.recording, undefined);
  near(settled.patches.find((p) => p.take === 'screen').timing.end, 9.5, 'screen');
  near(settled.patches.find((p) => p.take === 'computer').timing.end, 9.5, 'its sound');
  near(settled.patches.find((p) => p.take === 'camera').timing.end, 9.48, 'camera');
  assert.equal(settled.sources[SCREEN.url].duration, 9.5);
  assert.equal(settleTakes(settled, { [SCREEN.url]: 20 }), settled, 'a finished project is left alone');
});

/** Enough of a 2D canvas for a plan to be built and painted in Node. */
function canvasStub() {
  const calls = [];
  const ctx = new Proxy({ calls, measureText: (s) => ({ width: String(s).length * 10 }), createLinearGradient: () => ({ addColorStop() {} }) }, {
    get: (o, k) => (k in o ? o[k] : (...args) => { calls.push([k, ...args]); }),
    set: (o, k, v) => { o[k] = v; return true; },
  });
  return { width: 0, height: 0, getContext: () => ctx, ctx };
}

test('the painter cuts the camera to a circle and mirrors it inside its own box, over the screen letterboxed to its crop', () => {
  const { doc } = placeTakes(emptyProject(), [{ ...SCREEN, crop: { x: 0.25, y: 0.25, w: 0.5, h: 0.5 } }, CAMERA]);
  const canvas = canvasStub();
  const el = (w, h) => ({ readyState: 4, paused: true, currentTime: 0, videoWidth: w, videoHeight: h });
  const s = el(1920, 1080);
  const c = el(1280, 720);
  const plan = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos: { [SCREEN.url]: { el: s, img: s, width: 1920, height: 1080, duration: 12 }, [CAMERA.url]: { el: c, img: c, width: 1280, height: 720, duration: 12 } }, assets: {} });
  const cam = plan.layers.find((l) => l.take === 'camera' || l.url === CAMERA.url);
  assert.equal(cam.insetShape, 'circle');
  paintFrame(canvas.ctx, plan, 2);
  const calls = canvas.ctx.calls;
  const camDraw = calls.findIndex((k) => k[0] === 'drawImage' && k[1] === c);
  const scrDraw = calls.findIndex((k) => k[0] === 'drawImage' && k[1] === s);
  assert.ok(scrDraw >= 0 && camDraw > scrDraw, 'the screen first, the camera over it');
  // The screen through its crop: the 9-argument drawImage from the middle quarter.
  assert.deepEqual(calls[scrDraw].slice(2, 6), [480, 270, 960, 540]);
  const before = calls.slice(scrDraw, camDraw).map((k) => k[0]);
  assert.ok(before.includes('ellipse') && before.includes('clip'), 'cut to a circle');
  assert.ok(calls.slice(scrDraw, camDraw).some((k) => k[0] === 'scale' && k[1] === -1), 'mirrored');
  // Turned to a rectangle, the inset keeps square corners.
  const rect = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches({ ...doc, patches: doc.patches.map((p) => (p.take === 'camera' ? { ...p, insetShape: 'rect', flipX: false } : p)) }), videos: plan.videos, assets: {} });
  canvas.ctx.calls.length = 0;
  paintFrame(canvas.ctx, rect, 2);
  const ci = canvas.ctx.calls.findIndex((k) => k[0] === 'drawImage' && k[1] === c);
  assert.ok(!canvas.ctx.calls.slice(0, ci).some((k) => k[0] === 'scale' && k[1] === -1), 'not mirrored');
  assert.equal(clipById(doc, doc.patches.find((p) => p.take === 'camera').id).lane, 1);
  assert.equal(laneClips(doc, 'v', 1).length, 1);
  assert.equal(partnersOf(doc, doc.patches[0].id).length, 3);
});

/* A WebM as a recorder writes it: no declared length, the blocks in one cluster. */
const ebmlSize = (n) => { const out = Buffer.alloc(8); out.writeBigUInt64BE(BigInt(n)); out[0] = 0x01; return out; };
const ebml = (id, ...payload) => { const body = Buffer.concat(payload); return Buffer.concat([Buffer.from(id), ebmlSize(body.length), body]); };
const ebmlUint = (id, n) => { const b = Buffer.alloc(8); b.writeBigUInt64BE(BigInt(n)); let i = 0; while (i < 7 && b[i] === 0) i++; return ebml(id, b.subarray(i)); };
const ebmlStr = (id, str) => ebml(id, Buffer.from(str, 'utf8'));

test('a screen recording of a still window, its frames only where it changed, is read as long as its sound, not a frame-gap longer', async () => {
  const { probeMedia } = await import('../packages/media/src/probe.js');
  const header = ebml([0x1a, 0x45, 0xdf, 0xa3], ebmlStr([0x42, 0x82], 'webm'));
  const tracks = ebml([0x16, 0x54, 0xae, 0x6b],
    ebml([0xae], ebmlUint([0xd7], 1), ebmlUint([0x83], 1), ebmlStr([0x86], 'V_VP9'), ebml([0xe0], ebmlUint([0xb0], 320), ebmlUint([0xba], 180))),
    ebml([0xae], ebmlUint([0xd7], 2), ebmlUint([0x83], 2), ebmlStr([0x86], 'A_OPUS')));
  const block = (track, ms) => ebml([0xa3], Buffer.from([0x80 | track, (ms >> 8) & 0xff, ms & 0xff, 0x80]), Buffer.from([0, 0, 0, 0]));
  const blocks = [];
  for (const ms of [0, 33, 66, 1000]) blocks.push([1, ms]);
  for (let ms = 0; ms <= 1200; ms += 20) blocks.push([2, ms]);
  blocks.sort((a, b) => a[1] - b[1]);
  const cluster = ebml([0x1f, 0x43, 0xb6, 0x75], ebmlUint([0xe7], 0), ...blocks.map(([tr, ms]) => block(tr, ms)));
  const file = Buffer.concat([header, ebml([0x18, 0x53, 0x80, 0x67], ebml([0x15, 0x49, 0xa9, 0x66], ebmlUint([0x2a, 0xd7, 0xb1], 1000000)), tracks, cluster)]);
  const p = probeMedia(new Uint8Array(file), { complete: true });
  near(p.duration, 1.22, 'the sound\'s end', 0.005);
  assert.ok(p.hasVideo && p.hasAudio);
});
