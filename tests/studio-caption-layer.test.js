// The caption layer of the social app's video tool, in Office Video (S17 with
// S10): words typed or pasted, revealed five ways, split into one caption a
// line, with one typing speed for the film; the plain quick adds (words on a
// pill, QR, logo, footer); and a paragraph or list whose lines come in one at
// a time. The engine paints each in Node on a stub canvas.
import test from 'node:test';
import assert from 'node:assert/strict';
import '../packages/studio/src/toolset.js';
import { buildPlan, paintFrame } from '../packages/studio/src/index.js';
import {
  emptyProject, normaliseProject, addGraphic, clipById, planPatches, laneOf,
  NEW_CAPTION, CAPTION_REVEALS, captionSeconds, captionSegments, splitCaption, captionSpeed, setCaptionSpeed, quickAdd, CAPTION_SPEED,
} from '../packages/studio/src/project.js';

function canvasStub() {
  const calls = [];
  const state = { calls, measureText: (s) => ({ width: String(s).length * 10 }), createLinearGradient: () => ({ addColorStop() {} }), fillStyle: '', globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get: (o, k) => (k in o ? o[k] : (...args) => { calls.push([k, ...args, o.globalAlpha]); }),
    set: (o, k, v) => { o[k] = v; return true; },
  });
  return { width: 0, height: 0, getContext: () => ctx, ctx };
}

/** The words written on the frame at `t`, with the opacity each was written at. */
function written(doc, t) {
  const canvas = canvasStub();
  const plan = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos: {}, assets: {} });
  paintFrame(canvas.ctx, plan, t);
  return canvas.ctx.calls.filter((c) => c[0] === 'fillText').map((c) => ({ text: c[1], alpha: c[c.length - 1] }));
}

const WORDS = 'First we plan. Then we build. Then we ship.';

test('a caption is added with the person\'s words, for as long as they take to type and read', () => {
  const r = addGraphic(emptyProject(), { ...NEW_CAPTION, text: WORDS }, { at: 1 });
  const clip = clipById(r.doc, r.ids[0]);
  assert.equal(clip.type, 'caption');
  assert.equal(clip.text, WORDS);
  assert.equal(clip.timing.start, 1);
  assert.equal(clip.timing.end - clip.timing.start, captionSeconds(WORDS, { cps: 16 }));
  assert.ok(laneOf(clip) >= 1, 'over the pictures');
});

test('every reveal draws the caption, and all at once draws the whole of it at once', () => {
  for (const reveal of CAPTION_REVEALS) {
    const doc = { ...emptyProject(), patches: [{ ...NEW_CAPTION, id: 'c1', text: 'Hello world', reveal, timing: { start: 0, end: 6 }, lane: 1 }] };
    const late = written(doc, 5).map((w) => w.text).join(' ');
    assert.ok(late.includes('Hello world'), `${reveal}: the whole line is on the frame by the end`);
  }
  const doc = { ...emptyProject(), patches: [{ ...NEW_CAPTION, id: 'c1', text: 'Hello world', reveal: 'all', timing: { start: 0, end: 6 }, lane: 1 }] };
  assert.deepEqual(written(doc, 0.5).map((w) => w.text), ['Hello world']);
  const typing = { ...doc, patches: [{ ...doc.patches[0], reveal: 'type' }] };
  const early = written(typing, 0.5).map((w) => w.text).join('');
  assert.ok(early.length < 'Hello world'.length, 'typed: part of it a little after it starts');
  const words = { ...doc, patches: [{ ...doc.patches[0], reveal: 'word' }] };
  assert.equal(written(words, 0.5).map((w) => w.text).join(''), 'Hello', 'word by word: a word lands whole');
});

test('a caption keeps its own position, style, colour and size, and a film made before keeps its look', () => {
  const doc = { ...emptyProject(), patches: [{ ...NEW_CAPTION, id: 'c1', text: 'Hi', reveal: 'all', color: '#ffffff', style: 'bare', position: 'middle', sizeScale: 1.5, timing: { start: 0, end: 4 }, lane: 1 }] };
  const canvas = canvasStub();
  const plan = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos: {}, assets: {} });
  paintFrame(canvas.ctx, plan, 1);
  // Bare: no panel filled behind the words.
  assert.equal(canvas.ctx.calls.filter((c) => c[0] === 'fill').length, 0);
  assert.equal(canvas.ctx.fillStyle, '#ffffff');
  // A project saved by 2.27.0 reads back unchanged: nothing new is written into it.
  const old = { kind: 'video', patches: [{ id: 'text-1', type: 'text', text: 'Old', timing: { start: 0, end: 2 }, lane: 1 }], options: { theme: 'dark' } };
  const back = normaliseProject(old);
  assert.deepEqual(back.patches.map((p) => ({ id: p.id, type: p.type, text: p.text })), [{ id: 'text-1', type: 'text', text: 'Old' }]);
  assert.equal(back.options.charsPerSecond, normaliseProject({}).options.charsPerSecond);
});

test('the words split at the person\'s line breaks, else at sentences, never at the screen\'s wrap', () => {
  assert.deepEqual(captionSegments('One\nTwo\n\nThree'), ['One', 'Two', 'Three']);
  assert.deepEqual(captionSegments(WORDS), ['First we plan.', 'Then we build.', 'Then we ship.']);
  assert.deepEqual(captionSegments('No stop here'), ['No stop here']);
  assert.deepEqual(captionSegments('   '), []);
});

test('split makes one caption a line, one after another inside its time, each on its own track, one group', () => {
  const r = addGraphic(emptyProject(), { ...NEW_CAPTION, text: WORDS, reveal: 'karaoke', color: 'accent' }, { at: 2 });
  const was = clipById(r.doc, r.ids[0]);
  const s = splitCaption(r.doc, was.id);
  assert.equal(s.ids.length, 3);
  const lines = s.ids.map((id) => clipById(s.doc, id));
  assert.deepEqual(lines.map((c) => c.text), ['First we plan.', 'Then we build.', 'Then we ship.']);
  assert.equal(lines[0].timing.start, was.timing.start);
  assert.equal(lines[2].timing.end, was.timing.end);
  for (let i = 1; i < 3; i++) assert.ok(Math.abs(lines[i].timing.start - lines[i - 1].timing.end) < 0.002, 'each follows the one before');
  assert.equal(new Set(lines.map((c) => c.lane)).size, 3, 'a track each');
  assert.equal(new Set(lines.map((c) => c.group)).size, 1, 'one group');
  for (const c of lines) { assert.equal(c.reveal, 'karaoke'); assert.equal(c.color, 'accent'); }
  assert.equal(s.doc.patches.filter((p) => p.type === 'caption').length, 3, 'the caption it came from is gone');
  // One line: nothing to split.
  const one = addGraphic(emptyProject(), { ...NEW_CAPTION, text: 'Only one' });
  assert.deepEqual(splitCaption(one.doc, one.ids[0]).ids, []);
});

test('one typing speed for the film, kept within its bounds', () => {
  const doc = emptyProject();
  assert.equal(captionSpeed(doc), doc.options.charsPerSecond || CAPTION_SPEED.standard);
  assert.equal(captionSpeed(setCaptionSpeed(doc, 30)), 30);
  assert.equal(captionSpeed(setCaptionSpeed(doc, 1)), CAPTION_SPEED.min);
  assert.equal(captionSpeed(setCaptionSpeed(doc, 500)), CAPTION_SPEED.max);
});

test('the quick adds: words on a pill, a QR code with nothing to scan, the logo, a footer line', () => {
  let doc = emptyProject();
  const pill = addGraphic(doc, quickAdd('pill', { text: 'NEW' }));
  assert.equal(clipById(pill.doc, pill.ids[0]).pill, 'accent');
  assert.deepEqual(written(pill.doc, 0.5).map((w) => w.text), ['NEW']);
  const qr = addGraphic(doc, quickAdd('qr'));
  assert.equal(clipById(qr.doc, qr.ids[0]).data, '', 'no made-up address');
  assert.equal(quickAdd('logo'), null, 'no logo, no layer');
  const logo = addGraphic(doc, quickAdd('logo', { logo: 'D:/brand/mark.png' }));
  assert.equal(clipById(logo.doc, logo.ids[0]).url, 'D:/brand/mark.png');
  const foot = addGraphic(doc, quickAdd('footer', { text: 'example.com' }));
  const f = clipById(foot.doc, foot.ids[0]);
  assert.equal(f.type, 'text');
  assert.equal(f.baseline, 'bottom');
  assert.deepEqual(written(foot.doc, 0.5).map((w) => w.text), ['example.com']);
  doc = null;
});

test('a list brings its lines in one at a time, a gap apart (S10)', () => {
  const block = { id: 'b1', type: 'textblock', list: 'bullet', text: 'One\nTwo\nThree', fx: 0.5, fy: 0.5, fw: 0.6, sizeFrac: 0.035, lineIn: 1, timing: { start: 0, end: 8 }, lane: 1 };
  const doc = { ...emptyProject(), patches: [block] };
  const seen = (t) => written(doc, t).filter((w) => w.alpha > 0 && ['One', 'Two', 'Three'].includes(w.text)).map((w) => w.text);
  assert.deepEqual(seen(0.5), ['One']);
  assert.deepEqual(seen(1.5), ['One', 'Two']);
  assert.deepEqual(seen(3), ['One', 'Two', 'Three']);
  // Without lineIn the block shows whole, as before.
  const whole = { ...doc, patches: [{ ...block, lineIn: undefined }] };
  assert.deepEqual(written(whole, 0.1).filter((w) => ['One', 'Two', 'Three'].includes(w.text)).map((w) => w.text), ['One', 'Two', 'Three']);
});
