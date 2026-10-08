// Record's voice brought down to the rate it is kept at, as it comes.

import { toPcm } from './sounds.js';

/**
 * The microphone's samples brought to the WAV's rate as they come, by
 * straight lines between them, and kept as the 16-bit values the WAV will
 * hold: a quarter of what the raw samples took, for a voice that runs as long
 * as a slide is up. The fraction and the last sample carry from one chunk to
 * the next, so the joins are seamless.
 */
export class Downsampler {
  constructor(from, to) {
    this.step = from / to;
    this.pos = 0;
    this.last = 0;
    this.parts = [];
    this.length = 0;
  }
  push(chunk) {
    const n = chunk.length;
    if (!n) return;
    const out = new Int16Array(Math.ceil((n - this.pos) / this.step) + 1);
    let k = 0;
    let p = this.pos;
    // Position -1 is the last sample of the chunk before.
    while (p <= n - 1) {
      const a = Math.floor(p);
      const t = p - a;
      const va = a < 0 ? this.last : chunk[a];
      const vb = a + 1 < n ? chunk[a + 1] : va;
      out[k++] = toPcm(va * (1 - t) + vb * t);
      p += this.step;
    }
    this.pos = p - n;
    this.last = chunk[n - 1];
    this.parts.push(out.subarray(0, k));
    this.length += k;
  }
  /** Everything so far, as one run of samples; and a fresh start. */
  take() {
    const all = new Int16Array(this.length);
    let at = 0;
    for (const part of this.parts) { all.set(part, at); at += part.length; }
    this.parts = [];
    this.length = 0;
    return all;
  }
}
