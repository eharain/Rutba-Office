// Video's Reverse (the 2.7 plan's time items): a clip and its own sound
// played from a copy of their stretch written backward, in the same place
// and at the same length and speed; a recorded pointer drawn backward over
// the copy; and back again to the source, showing the same stretch, without
// a copy, even after the reversed clip was trimmed.
import test from 'node:test';
import assert from 'node:assert/strict';
import { emptyProject, addMedia, clipById, trimStart, setSpeed, setReversed, lengthOf } from '../packages/studio/src/project.js';

const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 6, hasAudio: true, width: 320, height: 240 };
const COPY = (seconds) => ({ url: 'D:/clips/a reversed.mp4', name: 'a reversed.mp4', kind: 'video', duration: seconds, hasAudio: true, width: 320, height: 240 });

function oneClip() {
  const { doc, ids } = addMedia(emptyProject(), A, { at: 0 });
  // The stretch from 1 s to 5 s of the source, starting at 0 on the clock.
  const out = trimStart(doc, ids[0], 1, { ripple: true });
  return { doc: out, video: ids[0], sound: ids[1] };
}

test('a reversed clip and its sound play the copy from its start, in the same place and length', () => {
  const { doc, video, sound } = oneClip();
  const before = clipById(doc, video);
  const out = setReversed(doc, video, COPY(lengthOf(before)));
  const v = clipById(out, video);
  const s = clipById(out, sound);
  assert.equal(v.url, 'D:/clips/a reversed.mp4');
  assert.equal(s.url, 'D:/clips/a reversed.mp4');
  assert.equal(v.offset, 0);
  assert.deepEqual(v.timing, before.timing);
  assert.deepEqual(v.reversed, { url: A.url, offset: before.offset, span: lengthOf(before) });
  assert.equal(out.sources['D:/clips/a reversed.mp4'].duration, lengthOf(before));
});

test('reversed again, the clip is back on its source showing the same stretch', () => {
  const { doc, video, sound } = oneClip();
  const before = clipById(doc, video);
  const there = setReversed(doc, video, COPY(lengthOf(before)));
  const back = setReversed(there, video);
  for (const id of [video, sound]) {
    const p = clipById(back, id);
    assert.equal(p.url, A.url);
    assert.equal(p.offset, before.offset);
    assert.equal(p.reversed, undefined);
  }
});

test('a reversed clip trimmed at its start comes back showing the stretch it showed', () => {
  const { doc, video } = oneClip();
  const before = clipById(doc, video);
  const span = lengthOf(before);
  let out = setReversed(doc, video, COPY(span));
  // The first second of the reversed copy is the last second of the stretch: trimmed away.
  out = trimStart(out, video, before.timing.start + 1, { ripple: false });
  const trimmed = clipById(out, video);
  assert.equal(trimmed.offset, 1);
  out = setReversed(out, video);
  const p = clipById(out, video);
  // What is left is the source's stretch from its start to a second before its end.
  assert.equal(p.offset, before.offset);
  assert.equal(lengthOf(p), span - 1);
});

test('a reversed clip keeps its speed, and its length on the clock with it', () => {
  const { doc, video } = oneClip();
  const fast = setSpeed(doc, video, 2);
  const clip = clipById(fast, video);
  const out = setReversed(fast, video, COPY(lengthOf(clip) * 2));
  const v = clipById(out, video);
  assert.equal(v.speed, 2);
  assert.deepEqual(v.timing, clip.timing);
  assert.equal(v.reversed.span, lengthOf(clip) * 2);
  const back = clipById(setReversed(out, video), video);
  assert.equal(back.offset, clip.offset);
});

test('a pointer recorded with the screen is drawn backward over the copy, and forward again after', () => {
  const { doc, video } = oneClip();
  const v = clipById(doc, video);
  const pointer = { id: 'pointer-1', type: 'pointer', name: 'Pointer', screenUrl: A.url, offset: v.offset, timing: { ...v.timing }, lane: 1, group: v.group, samples: [], clicks: [] };
  const withPointer = { ...doc, patches: [...doc.patches, pointer] };
  const out = setReversed(withPointer, video, COPY(lengthOf(v)));
  const p = clipById(out, 'pointer-1');
  assert.equal(p.backward, true);
  assert.equal(p.screenUrl, 'D:/clips/a reversed.mp4');
  const back = clipById(setReversed(out, video), 'pointer-1');
  assert.equal(back.backward, undefined);
  assert.equal(back.screenUrl, A.url);
});

test('a sound clip, a still or a reverse with no copy changes nothing', () => {
  const { doc, video, sound } = oneClip();
  assert.equal(setReversed(doc, sound, COPY(4)), doc);
  assert.equal(setReversed(doc, video), doc);
  assert.equal(setReversed(doc, 'nothing', COPY(4)), doc);
});
