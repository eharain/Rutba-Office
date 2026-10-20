// Video's library (the 2.7 plan's library and panels): sound effects made
// from sines and noise, the same every time, as WAV files; and a shape or an
// icon put on the lowest free track above the pictures, showing at once.
import test from 'node:test';
import assert from 'node:assert/strict';
import { SOUND_EFFECTS, effectSamples, effectWav, wavBytes } from '../packages/studio/src/effects.js';
import { emptyProject, addMedia, addGraphic, clipById, planPatches } from '../packages/studio/src/project.js';

test('every effect is made at its length, peaking under full scale, with no click at either end', () => {
  for (const { key, seconds } of SOUND_EFFECTS) {
    const s = effectSamples(key, 16000);
    assert.equal(s.length, Math.round(seconds * 16000), key);
    let peak = 0;
    for (const v of s) peak = Math.max(peak, Math.abs(v));
    assert.ok(peak > 0.7 && peak <= 0.8001, `${key} peaks at ${peak}`);
    assert.equal(Math.abs(s[0]), 0, key);
    assert.ok(Math.abs(s[s.length - 1]) < 1e-6, key);
  }
});

test('an effect is the same every time', () => {
  assert.deepEqual(effectSamples('whoosh', 8000), effectSamples('whoosh', 8000));
  assert.equal(effectSamples('nothing'), null);
});

test('an effect is a 16-bit mono WAV', () => {
  const bytes = effectWav('ding', 8000);
  const v = new DataView(bytes.buffer);
  assert.equal(String.fromCharCode(...bytes.subarray(0, 4)), 'RIFF');
  assert.equal(String.fromCharCode(...bytes.subarray(8, 12)), 'WAVE');
  assert.equal(v.getUint16(22, true), 1);
  assert.equal(v.getUint32(24, true), 8000);
  assert.equal(v.getUint16(34, true), 16);
  assert.equal(v.getUint32(40, true), Math.round(1.1 * 8000) * 2);
  assert.equal(wavBytes(new Float32Array([1, -1]), 8000).length, 48);
});

test('a shape goes on the lowest free track above the pictures, at the playhead, showing at once', () => {
  const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 6, hasAudio: false, width: 320, height: 240 };
  let doc = addMedia(emptyProject(), A, { at: 0 }).doc;
  const first = addGraphic(doc, { type: 'shape', geometry: 'star', fx: 0.5, fy: 0.5, fw: 0.36 }, { at: 1, seconds: 3 });
  const star = clipById(first.doc, first.ids[0]);
  assert.equal(star.type, 'shape');
  assert.equal(star.lane, 1);
  assert.deepEqual(star.timing, { start: 1, end: 4 });
  assert.equal(star.enter.kind, 'none');
  assert.equal(star.exit.kind, 'none');
  // An icon over the same time goes on the next track up.
  const second = addGraphic(first.doc, { type: 'icon', name: 'heart', fx: 0.5, fy: 0.5, fw: 0.16 }, { at: 2 });
  const heart = clipById(second.doc, second.ids[0]);
  assert.equal(heart.lane, 2);
  assert.deepEqual(heart.timing, { start: 2, end: 7 });
  assert.ok(planPatches(second.doc).some((p) => p.type === 'icon' && p.name === 'heart'));
  doc = second.doc;
  assert.deepEqual(addGraphic(doc, { type: 'video' }).ids, []);
});
