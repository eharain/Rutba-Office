// Good pictures in every format decodePicture reads, made here for the
// fuzzer (tools/fuzz-open.js) to damage: each a small gradient with a
// see-through column, written by the suite's own writers or, for formats
// it only reads, by the few lines below that lay the format out.

import { encodePng } from '../packages/imaging/src/codecs/png.js';
import { encodeBmp } from '../packages/imaging/src/codecs/bmp.js';

/** A small picture to write: colour across, down and both, one column half clear. */
export function sample(w = 24, h = 16, { alpha = true } = {}) {
  const rgba = new Uint8Array(w * h * 4);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = (y * w + x) * 4;
      rgba[i] = Math.round((x * 255) / (w - 1));
      rgba[i + 1] = Math.round((y * 255) / (h - 1));
      rgba[i + 2] = (x * 16 + y * 8) & 255;
      rgba[i + 3] = alpha && x === 2 ? 128 : 255;
    }
  }
  return rgba;
}

/** Source name -> bytes; every name begins `picture.` and ends with the format's extension. */
export function pictureSources() {
  const w = 24;
  const h = 16;
  const rgba = sample(w, h);
  return {
    'picture.png': Buffer.from(encodePng(rgba, w, h)),
    'picture.bmp': Buffer.from(encodeBmp(rgba, w, h)),
  };
}
