// A caption's typewriter with its read-aloud voice (S19): a caption read
// through Spoken lines stays a caption a line, and each line types over its
// voice's length, so the last letter lands as the voice ends; with "voice
// first" the voice starts a moment before the typing; with the words shown
// only the typing keeps the film's own speed. A title read aloud is as before.
import test from 'node:test';
import assert from 'node:assert/strict';
import '../packages/studio/src/toolset.js';
import { buildPlan, paintFrame } from '../packages/studio/src/index.js';
import { emptyProject, addGraphic, planPatches, NEW_CAPTION } from '../packages/studio/src/project.js';
import {
  placeSpokenLines, setSpokenChoice, setSpokenTyping, spokenTypingOf, spokenLayer, wordsToRead, typingWithVoice, VOICE_FIRST_SECONDS,
} from '../packages/studio/src/spoken.js';

function canvasStub() {
  const calls = [];
  const state = { calls, measureText: (s) => ({ width: String(s).length * 10 }), createLinearGradient: () => ({ addColorStop() {} }), fillStyle: '', globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get: (o, k) => (k in o ? o[k] : (...args) => { calls.push([k, ...args]); }),
    set: (o, k, v) => { o[k] = v; return true; },
  });
  return { width: 0, height: 0, getContext: () => ctx, ctx };
}

/** The letters written on the frame at `t`. */
function typed(doc, t) {
  const canvas = canvasStub();
  const plan = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos: {}, assets: {} });
  paintFrame(canvas.ctx, plan, t);
  return canvas.ctx.calls.filter((c) => c[0] === 'fillText').map((c) => c[1]).join('').replace(/\s+/g, '').length;
}

const ONE = 'We plan the week together.';
const TWO = 'Then everyone builds their part.';

function readCaption(options = {}) {
  const added = addGraphic(emptyProject(), { ...NEW_CAPTION, text: `${ONE} ${TWO}` }, { at: 1 });
  const lines = [
    { text: ONE, url: 'D:/made/sp1-1.wav', name: 'sp1-1.wav', seconds: 2 },
    { text: TWO, url: 'D:/made/sp1-2.wav', name: 'sp1-2.wav', seconds: 3 },
  ];
  return placeSpokenLines(added.doc, { from: added.ids[0], lines, voice: 'v1', ...options });
}

const lineClips = (doc, layer) => spokenLayer(doc, layer).lines.map((l) => l.text);

test('a caption is read a line at a time, at its sentences', () => {
  const added = addGraphic(emptyProject(), { ...NEW_CAPTION, text: `${ONE} ${TWO}` }, { at: 1 });
  assert.deepEqual(wordsToRead(added.doc, added.ids[0]), [ONE, TWO]);
});

test('read aloud, each caption line types over its voice, the last letter as the voice ends', () => {
  const { doc, layer } = readCaption();
  const [a, b] = lineClips(doc, layer);
  assert.equal(a.type, 'caption', 'the lines stay captions, with the caption\'s look');
  assert.equal(a.reveal, 'type');
  assert.deepEqual(a.timing, { start: 1, end: 3 });
  assert.equal(a.leadIn, 0);
  assert.equal(a.pace, 1.9);
  assert.equal(b.pace, 2.9);
  assert.equal(spokenTypingOf(doc, layer), 'with', 'with the voice unless asked');
  const letters = ONE.replace(/\s+/g, '').length;
  const half = typed(doc, a.timing.start + 1);
  assert.ok(half > letters * 0.3 && half < letters * 0.7, `half way through the voice, about half the words (${half} of ${letters})`);
  assert.equal(typed(doc, a.timing.end - 0.01), letters, 'all of it as the voice ends');
});

test('voice first: the voice starts a moment before the typing, which still ends with it', () => {
  const read = readCaption();
  const doc = setSpokenTyping(read.doc, read.layer, 'first');
  const [a] = lineClips(doc, read.layer);
  assert.equal(spokenTypingOf(doc, read.layer), 'first');
  assert.equal(a.leadIn, VOICE_FIRST_SECONDS);
  assert.equal(a.pace, 1.5);
  assert.equal(typed(doc, a.timing.start + VOICE_FIRST_SECONDS / 2), 0, 'nothing typed while the voice begins');
  assert.equal(typed(doc, a.timing.end - 0.01), ONE.replace(/\s+/g, '').length);
  // Read again, the choice is kept.
  const again = readCaption({ typing: 'first' });
  assert.equal(lineClips(again.doc, again.layer)[0].leadIn, VOICE_FIRST_SECONDS);
});

test('shown only, with no voice played, the typing keeps the film\'s own speed, and takes the voice again when played', () => {
  const read = readCaption();
  const shown = setSpokenChoice(read.doc, read.layer, { show: true, play: false });
  const [a] = lineClips(shown, read.layer);
  assert.equal(a.pace, undefined);
  assert.equal(a.leadIn, NEW_CAPTION.leadIn);
  const both = setSpokenChoice(shown, read.layer, { show: true, play: true });
  assert.equal(lineClips(both, read.layer)[0].pace, 1.9);
  assert.deepEqual(typingWithVoice({ timing: { start: 0, end: 0.5 } }, { typing: 'first' }), { leadIn: 0.25, pace: 0.15 }, 'a short line: the voice is ahead by half of it at most');
});

test('a title read aloud is read as before: text lines, no typing pace', () => {
  const doc = { ...emptyProject(), patches: [{ id: 'text-1', type: 'text', text: 'Hello\nThere', timing: { start: 0, end: 4 }, lane: 1, fx: 0.5, fy: 0.4, sizeFrac: 0.05 }] };
  const lines = [
    { text: 'Hello', url: 'D:/made/a.wav', seconds: 1 },
    { text: 'There', url: 'D:/made/b.wav', seconds: 1 },
  ];
  const r = placeSpokenLines(doc, { from: 'text-1', lines });
  const made = lineClips(r.doc, r.layer);
  assert.deepEqual(made.map((p) => p.type), ['text', 'text']);
  assert.ok(made.every((p) => p.pace === undefined && p.spoken.typing === undefined));
  assert.equal(spokenTypingOf(r.doc, r.layer), null);
  assert.equal(r.doc.patches.filter((p) => p.type === 'text').length, 2, 'the title it was read from is replaced by its lines');
});
