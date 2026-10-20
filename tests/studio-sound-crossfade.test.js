// Video's crossfade between two sound clips of their own (the 2.7 plan's
// sound items): they overlap by its length, the first fading out as the
// second fades in, what follows on the track moving up with it; taken off,
// they meet again. A picture's own sound is left to its picture's transition.
import test from 'node:test';
import assert from 'node:assert/strict';
import { emptyProject, addMedia, clipById, setTransition, clearTransition } from '../packages/studio/src/project.js';

const TONE = (name, duration) => ({ url: `D:/clips/${name}.wav`, name: `${name}.wav`, kind: 'audio', duration });
const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 6, hasAudio: true, width: 320, height: 240 };

function threeSounds() {
  let doc = emptyProject();
  const ids = [];
  for (const [name, at] of [['one', 0], ['two', 4], ['three', 8]]) {
    const r = addMedia(doc, TONE(name, 4), { at, lane: 1 });
    doc = r.doc;
    ids.push(r.ids[0]);
  }
  return { doc, ids };
}

test('two sound clips crossfade: they overlap by its length, one fading out as the other fades in', () => {
  const { doc, ids: [one, two, three] } = threeSounds();
  const out = setTransition(doc, two, { kind: 'fade', seconds: 1 });
  const a = clipById(out, one);
  const b = clipById(out, two);
  assert.deepEqual(b.timing, { start: 3, end: 7 });
  assert.deepEqual(b.enter, { kind: 'fade', seconds: 1 });
  assert.deepEqual(b.transition, { kind: 'fade', seconds: 1 });
  assert.deepEqual(a.exit, { kind: 'fade', seconds: 1 });
  // What follows on the track moves up with it.
  assert.deepEqual(clipById(out, three).timing, { start: 7, end: 11 });
});

test('taken off, the sounds meet again with no fades', () => {
  const { doc, ids: [one, two, three] } = threeSounds();
  const out = clearTransition(setTransition(doc, two, { kind: 'fade', seconds: 1 }), two);
  assert.deepEqual(clipById(out, two).timing, { start: 4, end: 8 });
  assert.deepEqual(clipById(out, three).timing, { start: 8, end: 12 });
  assert.equal(clipById(out, two).transition, undefined);
  assert.deepEqual(clipById(out, two).enter, { kind: 'fade', seconds: 0 });
  assert.deepEqual(clipById(out, one).exit, { kind: 'fade', seconds: 0 });
});

test('a crossfade is held to half of each sound', () => {
  const { doc, ids: [, two] } = threeSounds();
  const out = setTransition(doc, two, { kind: 'fade', seconds: 9 });
  assert.equal(clipById(out, two).transition.seconds, 2);
});

test('between sounds only a crossfade; a picture\'s own sound is left to its picture', () => {
  const { doc, ids: [, two] } = threeSounds();
  assert.equal(clipById(setTransition(doc, two, { kind: 'wipe', seconds: 1 }), two).transition, undefined);
  let pictures = addMedia(emptyProject(), A, { at: 0 }).doc;
  const second = addMedia(pictures, A, { at: 6 });
  pictures = second.doc;
  const sound = second.ids[1];
  assert.equal(clipById(pictures, sound).type, 'sound');
  assert.equal(setTransition(pictures, sound, { kind: 'fade', seconds: 1 }), pictures);
});
