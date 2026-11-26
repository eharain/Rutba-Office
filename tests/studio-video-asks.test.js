// The film sessions' small Video asks for 2.28.0: the progress bar's colour,
// height and segments, and a bar that keeps covering the film as it grows
// (S1, S12); the end card's second line, logo, background and QR (S3, S13);
// a title that shrinks to fit (S6); markers written out as chapters (S2).
import test from 'node:test';
import assert from 'node:assert/strict';
import '../packages/studio/src/toolset.js';
import { buildPlan, paintFrame } from '../packages/studio/src/index.js';
import {
  emptyProject, addMedia, addGraphic, addTitle, addMarker, clipById, planPatches, projectDuration, projectMedia, trimEnd, chaptersText, OVERLAYS,
} from '../packages/studio/src/project.js';

/** The size in pixels a CSS font names. */
const fontPx = (font) => Number(/(\d+(?:\.\d+)?)px/.exec(String(font))?.[1] || 10);

function canvasStub() {
  const calls = [];
  const state = { calls, font: '10px sans-serif', measureText: (s) => ({ width: String(s).length * fontPx(state.font) * 0.5 }), createLinearGradient: () => ({ addColorStop() {} }), fillStyle: '', globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get: (o, k) => (k in o ? o[k] : (...args) => { calls.push([k, ...args, o.fillStyle]); }),
    set: (o, k, v) => { o[k] = v; return true; },
  });
  return { width: 0, height: 0, getContext: () => ctx, ctx };
}
// The QR painter draws its code once on a canvas of its own.
globalThis.document ??= { createElement: () => canvasStub() };

const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 10, hasAudio: false, width: 320, height: 240 };
const B = { url: 'D:/clips/b.webm', name: 'b.webm', kind: 'video', duration: 6, hasAudio: false, width: 320, height: 240 };
const videos = () => {
  const el = { readyState: 4, paused: true, currentTime: 0 };
  return { [A.url]: { el, img: el, width: 320, height: 240, duration: 10 }, [B.url]: { el, img: el, width: 320, height: 240, duration: 6 } };
};

function paint(doc, t, assets = {}) {
  const canvas = canvasStub();
  const plan = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos: videos(), assets });
  paintFrame(canvas.ctx, plan, t);
  return { plan, calls: canvas.ctx.calls };
}

const rectsIn = (calls, colour) => calls.filter((c) => c[0] === 'fillRect' && c[c.length - 1] === colour).map((c) => c.slice(1, 5));

test('a progress bar takes a colour and a height', () => {
  const film = addMedia(emptyProject(), A, { at: 0 }).doc;
  const r = addGraphic(film, { ...OVERLAYS.progress, color: '#ff8800', height: 0.01 });
  const { plan, calls } = paint(r.doc, 5);
  const bars = rectsIn(calls, '#ff8800');
  assert.equal(bars.length, 1);
  const [x, y, w, h] = bars[0];
  assert.equal(x, 0);
  assert.equal(h, Math.round(plan.H * 0.01));
  assert.equal(y, plan.H - h);
  assert.ok(Math.abs(w - plan.W * 0.5) < 1, 'half way through, half filled');
});

test('a progress bar in segments: one for each marker, a small gap between', () => {
  let doc = addMedia(emptyProject(), A, { at: 0 }).doc;
  doc = addMarker(doc, 4, 'Two');
  doc = addMarker(doc, 7, 'Three');
  const r = addGraphic(doc, { ...OVERLAYS.progress, color: '#00aa00', segments: 'markers' });
  const { plan, calls } = paint(r.doc, 2);
  // The track: three pieces, gaps between.
  const track = calls.filter((c) => c[0] === 'fillRect' && c[c.length - 1] !== '#00aa00' && c[2] > plan.H * 0.9);
  assert.equal(track.length, 3);
  assert.ok(track[1][1] > track[0][1] + track[0][3], 'a gap after the first piece');
  // Two seconds in: the first piece part filled, nothing in the others.
  const fill = rectsIn(calls, '#00aa00');
  assert.equal(fill.length, 1);
  assert.ok(fill[0][2] < plan.W * 0.4 && fill[0][2] > plan.W * 0.15);
});

test('a bar over the whole film keeps covering it as clips are added, until its length is changed', () => {
  const film = addMedia(emptyProject(), A, { at: 0 }).doc;
  const r = addGraphic(film, OVERLAYS.progress);
  const bar = clipById(r.doc, r.ids[0]);
  assert.equal(bar.whole, true);
  const longer = addMedia(r.doc, B, { at: 10 }).doc;
  assert.equal(projectDuration(longer), 16);
  const drawn = planPatches(longer).find((p) => p.id === bar.id);
  assert.deepEqual(drawn.timing, { start: 0, end: 16 });
  // Trimmed by hand, it keeps the length given it.
  const trimmed = trimEnd(longer, bar.id, 8);
  assert.equal(clipById(trimmed, bar.id).whole, undefined);
  assert.deepEqual(planPatches(trimmed).find((p) => p.id === bar.id).timing, { start: 0, end: 8 });
});

test('the end card: a second line, a logo, a background colour and a QR code once an address is typed', () => {
  const film = addMedia(emptyProject(), A, { at: 0 }).doc;
  const r = addGraphic(film, { ...OVERLAYS.outro, text: 'Thank you', subtitle: 'example.com', background: '#1e40af', logo: 'D:/brand/mark.png', qr: '' });
  const card = clipById(r.doc, r.ids[0]);
  assert.ok(projectMedia(r.doc).images.includes('D:/brand/mark.png'), 'its logo is loaded with the film');
  const logo = { img: { width: 200, height: 100 }, width: 200, height: 100 };
  const at = card.timing.start + 1;
  let { calls } = paint(r.doc, at, { 'D:/brand/mark.png': logo });
  const texts = calls.filter((c) => c[0] === 'fillText').map((c) => c[1]);
  assert.deepEqual(texts, ['Thank you', 'example.com']);
  assert.equal(rectsIn(calls, '#1e40af').length, 1, 'the card is the colour given');
  assert.ok(calls.some((c) => c[0] === 'drawImage' && c[1] === logo.img), 'its logo is drawn');
  const images = () => calls.filter((c) => c[0] === 'drawImage').length;
  assert.equal(images(), 1, 'no QR code while the address is blank');
  const withQr = { ...r.doc, patches: r.doc.patches.map((p) => (p.id === card.id ? { ...p, qr: 'https://example.com' } : p)) };
  ({ calls } = paint(withQr, at, { 'D:/brand/mark.png': logo }));
  assert.equal(images(), 2, 'the QR code is drawn once typed');
});

test('an end card saved by 2.27.0 draws as it did', () => {
  const film = addMedia(emptyProject(), A, { at: 0 }).doc;
  const r = addGraphic(film, { ...OVERLAYS.outro, text: 'Bye' });
  const card = clipById(r.doc, r.ids[0]);
  const { plan, calls } = paint(r.doc, card.timing.start + 1);
  assert.deepEqual(calls.filter((c) => c[0] === 'fillText').map((c) => c[1]), ['Bye']);
  const fills = calls.filter((c) => c[0] === 'fillRect' && c[3] === plan.W && c[4] === plan.H);
  assert.equal(fills[fills.length - 1][5], plan.theme.bg, 'the card in the theme\'s own colour');
  assert.equal(calls.filter((c) => c[0] === 'drawImage').length, 0, 'no logo or code it was not given');
});

test('a new title shrinks to fit the frame; one saved before keeps its size', () => {
  const long = 'A headline that runs on for a hundred and ten characters, far wider than the frame can hold at this size at all';
  const r = addTitle(emptyProject(), 'title', { at: 0, text: long });
  const title = clipById(r.doc, r.ids[0]);
  assert.equal(title.fit, true);
  const { plan, calls } = paint(r.doc, 1);
  const scale = calls.find((c) => c[0] === 'scale');
  assert.ok(scale, 'drawn smaller');
  const wide = long.length * fontPx(plan.layers.find((l) => l.type === 'text').font) * 0.5;
  assert.ok(scale[1] < 1 && scale[1] * wide <= plan.W);
  const old = { ...r.doc, patches: r.doc.patches.map((p) => ({ ...p, fit: undefined })) };
  assert.equal(paint(old, 1).calls.find((c) => c[0] === 'scale'), undefined);
  const short = addTitle(emptyProject(), 'title', { at: 0, text: 'Short' });
  assert.equal(paint(short.doc, 1).calls.find((c) => c[0] === 'scale'), undefined, 'words that fit are left as they are');
});

test('markers become chapters: 0:00 first, one line each, in the time of what is exported', () => {
  let doc = addMedia(emptyProject(), A, { at: 0 }).doc;
  assert.equal(chaptersText(doc), '', 'no markers, no chapters');
  doc = addMarker(doc, 65, 'Setup');
  doc = addMarker(doc, 12.4, '');
  doc = addMarker(doc, 3725, 'The end');
  assert.equal(chaptersText(doc, { first: 'Start', unnamed: (n) => `Part ${n}` }), '0:00 Start\n0:12 Part 2\n1:05 Setup\n1:02:05 The end\n');
  doc = addMarker(doc, 0, 'Opening');
  assert.equal(chaptersText(doc, { first: 'Start', from: 10, to: 100 }).split('\n')[0], '0:00 Start');
  assert.equal(chaptersText(doc, { first: 'Start' }).split('\n')[0], '0:00 Opening');
});

test('a bar over the whole film is saved at the film\'s length, so 2.27.0 opens it covering the film', async () => {
  const { writeRvid, readRvid } = await import('../packages/studio/src/rvid.js');
  const film = addMedia(emptyProject(), A, { at: 0 }).doc;
  const bar = addGraphic(film, OVERLAYS.progress);
  const longer = addMedia(bar.doc, B, { at: 10 }).doc;
  const back = (await readRvid(writeRvid(longer))).doc;
  assert.deepEqual(clipById(back, bar.ids[0]).timing, { start: 0, end: 16 });
  // A project with no such bar is written as it was.
  const plain = (await readRvid(writeRvid(film))).doc;
  assert.deepEqual(plain.patches, film.patches);
});

test('sound fetched from the web goes to one library folder, whichever project is open', async () => {
  const { webAudioFolder } = await import('../packages/studio/src/audio-library.js');
  const paths = { music: 'C:/Users/a/Music', videos: 'C:/Users/a/Videos', separator: '/' };
  assert.equal(webAudioFolder(paths, 'Rutba Recordings'), 'C:/Users/a/Music/Rutba Recordings');
  assert.equal(webAudioFolder({ videos: '/home/a/Videos', separator: '/' }, 'Rutba Recordings'), '/home/a/Videos/Rutba Recordings');
});
