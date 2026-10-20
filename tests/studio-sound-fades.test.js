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
