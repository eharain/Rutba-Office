// Video's speed with the pitch kept, and its pitch shift (the 2.7 plan's
// sound items): a tone stretched longer or shorter keeps its frequency, and
// the rate and stretch a clip is played with follow its speed, its semitones
// and whether the pitch is kept.
import test from 'node:test';
import assert from 'node:assert/strict';
import { stretchChannels, stretchedBuffer, soundRate, pitchFactor, PITCH_STEPS } from '../packages/studio/src/stretch.js';

const RATE = 16000;
const tone = (hz, seconds) => {
  const n = Math.round(seconds * RATE);
  const a = new Float32Array(n);
  for (let i = 0; i < n; i++) a[i] = Math.sin((2 * Math.PI * hz * i) / RATE) * 0.6;
  return a;
};
/** The frequency of a steady tone, from its upward zero crossings away from the ends. */
const frequency = (a) => {
  const from = Math.round(a.length * 0.1);
  const to = Math.round(a.length * 0.9);
  let n = 0;
  for (let i = from + 1; i < to; i++) if (a[i - 1] < 0 && a[i] >= 0) n++;
  return n / ((to - from) / RATE);
};

test('a tone made twice as long keeps its frequency', () => {
  const [out] = stretchChannels([tone(440, 1)], RATE, 2);
  assert.equal(out.length, 2 * RATE);
  assert.ok(Math.abs(frequency(out) - 440) < 6, `heard ${frequency(out)} Hz`);
});

test('a tone made half as long keeps its frequency', () => {
  const [out] = stretchChannels([tone(300, 2)], RATE, 0.5);
  assert.equal(out.length, RATE);
  assert.ok(Math.abs(frequency(out) - 300) < 6, `heard ${frequency(out)} Hz`);
});

test('a stretched tone stays at its level, with no gaps', () => {
  const [out] = stretchChannels([tone(440, 1)], RATE, 1.5);
  let peak = 0;
  // Every 20 ms window has sound in it.
  for (let w = Math.round(RATE * 0.05); w + 320 < out.length - RATE * 0.05; w += 320) {
    let m = 0;
    for (let i = w; i < w + 320; i++) m = Math.max(m, Math.abs(out[i]));
    assert.ok(m > 0.3, `quiet at ${w / RATE} s`);
    peak = Math.max(peak, m);
  }
  assert.ok(peak < 0.8, `peak ${peak}`);
});

test('two channels are stretched alike', () => {
  const left = tone(440, 1);
  const right = Float32Array.from(left, (v) => -v);
  const [l, r] = stretchChannels([left, right], RATE, 1.25);
  assert.equal(l.length, r.length);
  for (let i = 0; i < l.length; i += 97) assert.ok(Math.abs(l[i] + r[i]) < 1e-5);
});

test('a ratio of 1 is the sound as it was', () => {
  const a = tone(440, 0.5);
  const [out] = stretchChannels([a], RATE, 1);
  assert.deepEqual(out, a);
});

test('the same stretch of a buffer asked twice is made once', () => {
  const data = tone(440, 2);
  const buffer = { sampleRate: RATE, length: data.length, duration: 2, numberOfChannels: 1, getChannelData: () => data };
  const a = stretchedBuffer(buffer, 0.5, 1, 2);
  const b = stretchedBuffer(buffer, 0.5, 1, 2);
  assert.equal(a, b);
  assert.equal(a.length, 2 * RATE);
  assert.notEqual(stretchedBuffer(buffer, 0.5, 1, 1.5), a);
});

test('speed with the pitch kept stretches; without, the pitch goes with the speed; semitones move the pitch alone', () => {
  assert.deepEqual(soundRate({ speed: 2, keepPitch: false }), { rate: 2, ratio: 1 });
  assert.deepEqual(soundRate({ speed: 2, keepPitch: true }), { rate: 1, ratio: 0.5 });
  assert.deepEqual(soundRate({ speed: 0.5 }), { rate: 1, ratio: 2 });
  const up = soundRate({ speed: 1, semitones: 12 });
  assert.equal(up.rate, 2);
  assert.equal(up.ratio, 2);
  const both = soundRate({ speed: 2, semitones: 12, keepPitch: false });
  assert.equal(both.rate, 4);
  assert.equal(both.ratio, 2);
  assert.equal(pitchFactor(-12), 0.5);
  assert.ok(PITCH_STEPS.includes(0) && PITCH_STEPS[0] === -12 && PITCH_STEPS.at(-1) === 12);
});
