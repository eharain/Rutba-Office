// Record's voice is brought to the WAV's rate as it comes and kept as 16-bit
// samples: pushed in chunks of any size it comes out as it does in one, the
// joins seamless, the length right, and the WAV the same as before.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Downsampler } from '../apps/desktop/renderer/apps/slides/downsample.js';
import { wavOf, wavOfPcm, toPcm } from '../apps/desktop/renderer/apps/slides/sounds.js';

const sine = (n, rate, hz = 440) => Float32Array.from({ length: n }, (_, i) => 0.8 * Math.sin((2 * Math.PI * hz * i) / rate));

test('chunks of any size come out as one chunk does, at the rate asked for', () => {
  const input = sine(48000, 48000);
  const whole = new Downsampler(48000, 22050);
  whole.push(input);
  const once = whole.take();
  const pieces = new Downsampler(48000, 22050);
  for (let at = 0, k = 0; at < input.length; k++) {
    const size = [4096, 1, 777, 3000, 2][k % 5];
    pieces.push(input.subarray(at, at + size));
    at += size;
  }
  const chunked = pieces.take();
  assert.equal(chunked.length, once.length);
  assert.deepEqual([...chunked], [...once]);
  assert.ok(Math.abs(once.length - 22050) <= 1, `one second is ${once.length} samples`);
  // Each sample is the straight line between the two around it.
  const step = 48000 / 22050;
  for (const i of [0, 1, 500, 12345, once.length - 2]) {
    const p = i * step;
    const a = Math.floor(p);
    const expected = toPcm(input[a] * (1 - (p - a)) + (input[a + 1] ?? input[a]) * (p - a));
    assert.ok(Math.abs(once[i] - expected) <= 1, `sample ${i}`);
  }
  assert.equal(pieces.take().length, 0, 'taken, it starts again');
});

test('a WAV of samples already made is the WAV wavOf makes', () => {
  const samples = sine(1000, 22050);
  assert.deepEqual([...wavOfPcm(Int16Array.from(samples, toPcm), 22050)], [...wavOf(samples, 22050)]);
});
