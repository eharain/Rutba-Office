// Video's sound: fades as long as asked, shortened alike only when together
// they are longer than the clip; volume to +12 dB, in decibels.
import test from 'node:test';
import assert from 'node:assert/strict';
import { soundFades, MAX_GAIN, gainToDb, dbToGain } from '@rutba/studio';

test('fades: as long as asked, with no hidden third-of-clip cap', () => {
  assert.deepEqual(soundFades(4, 3, 10), { fadeIn: 4, fadeOut: 3, shortened: false });
  assert.deepEqual(soundFades(0, 0, 10), { fadeIn: 0, fadeOut: 0, shortened: false });
});

test('fades: together longer than the clip, both shortened alike to meet in the middle', () => {
  const f = soundFades(6, 6, 6);
  assert.equal(f.shortened, true);
  assert.equal(f.fadeIn, 3);
  assert.equal(f.fadeOut, 3);
  const g = soundFades(8, 2, 5);
  assert.equal(g.fadeIn, 4);
  assert.equal(g.fadeOut, 1);
});

test('volume: decibels and back, and the ceiling at +12 dB', () => {
  assert.equal(Math.round(gainToDb(1) * 10) / 10, 0);
  assert.equal(gainToDb(0), -Infinity);
  assert.equal(Math.round(dbToGain(6) * 100) / 100, 2);
  assert.equal(dbToGain(-Infinity), 0);
  assert.equal(Math.round(gainToDb(MAX_GAIN)), 12);
});

test('speed: a clip set to twice its speed takes half the time, its sound with it, and what follows on its lane moves up', async () => {
  const { setSpeed, speedOf } = await import('@rutba/studio/project');
  const doc = {
    sources: { 'a.webm': { duration: 20 } },
    patches: [
      { id: 'v1', type: 'video', url: 'a.webm', offset: 0, timing: { start: 0, end: 10 }, lane: 0, group: 'g1' },
      { id: 's1', type: 'sound', url: 'a.webm', offset: 0, timing: { start: 0, end: 10 }, lane: 0, group: 'g1' },
      { id: 'v2', type: 'video', url: 'a.webm', offset: 10, timing: { start: 10, end: 14 }, lane: 0 },
    ],
  };
  const out = setSpeed(doc, 'v1', 2);
  const by = Object.fromEntries(out.patches.map((p) => [p.id, p]));
  assert.equal(speedOf(by.v1), 2);
  assert.equal(speedOf(by.s1), 2);
  assert.deepEqual([by.v1.timing.start, by.v1.timing.end], [0, 5]);
  assert.deepEqual([by.s1.timing.start, by.s1.timing.end], [0, 5]);
  assert.deepEqual([by.v2.timing.start, by.v2.timing.end], [5, 9]);
  const back = setSpeed(out, 'v1', 0.5);
  const b = Object.fromEntries(back.patches.map((p) => [p.id, p]));
  assert.deepEqual([b.v1.timing.start, b.v1.timing.end], [0, 20]);
  assert.equal(b.v2.timing.start, 20);
});
