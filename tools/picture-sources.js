// Good pictures in every format decodePicture reads, made here for the
// fuzzer (tools/fuzz-open.js) to damage: each a small gradient with a
// see-through column, written by the suite's own writers or, for formats
// it only reads, by the few lines below that lay the format out.

import { encodePng } from '../packages/imaging/src/codecs/png.js';
import { encodeBmp } from '../packages/imaging/src/codecs/bmp.js';
import { encodeTiff } from '../packages/imaging/src/codecs/tiff.js';
import { encodeIco } from '../packages/imaging/src/encode.js';
import { encodeQoi } from '../packages/imaging/src/codecs/qoi.js';
import { encodePcx } from '../packages/imaging/src/codecs/pcx.js';
import { encodeHdr } from '../packages/imaging/src/codecs/hdr.js';

/** Netpbm: a binary PPM (P6), or a PAM (P7) with alpha. */
export function pnmOf(rgba, w, h, { pam = false } = {}) {
  const depth = pam ? 4 : 3;
  const head = pam
    ? `P7\nWIDTH ${w}\nHEIGHT ${h}\nDEPTH 4\nMAXVAL 255\nTUPLTYPE RGB_ALPHA\nENDHDR\n`
    : `P6\n# made for the fuzzer\n${w} ${h}\n255\n`;
  const out = new Uint8Array(head.length + w * h * depth);
  for (let i = 0; i < head.length; i++) out[i] = head.charCodeAt(i);
  for (let i = 0; i < w * h; i++) for (let c = 0; c < depth; c++) out[head.length + i * depth + c] = rgba[i * 4 + c];
  return out;
}

/** Targa: 32-bit BGRA, run-length packed (type 10), rows from the top. */
export function tgaOf(rgba, w, h) {
  const out = [0, 0, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, w & 255, w >> 8, h & 255, h >> 8, 32, 0x28];
  const px = (i) => [rgba[i * 4 + 2], rgba[i * 4 + 1], rgba[i * 4], rgba[i * 4 + 3]];
  const same = (a, b) => a.every((v, k) => v === b[k]);
  for (let i = 0; i < w * h;) {
    let n = 1;
    while (i + n < w * h && n < 128 && same(px(i + n), px(i))) n++;
    if (n > 1) { out.push(0x80 | (n - 1), ...px(i)); i += n; continue; }
    let m = 1;
    while (i + m < w * h && m < 128 && !same(px(i + m), px(i + m - 1))) m++;
    out.push(m - 1);
    for (let k = 0; k < m; k++) out.push(...px(i + k));
    i += m;
  }
  return Uint8Array.from(out);
}

/** Linear light for a Radiance file: the picture's colours, brighter to the right, past white. */
export function lightOf(rgba, w, h) {
  const light = new Float32Array(w * h * 3);
  for (let i = 0; i < w * h; i++) for (let c = 0; c < 3; c++) light[i * 3 + c] = (rgba[i * 4 + c] / 255) ** 2.2 * (1 + 7 * ((i % w) / w));
  return light;
}

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
    // two pages, so the page walk is damaged too
    'picture.tif': Buffer.from(encodeTiff([{ rgba, width: w, height: h }, { rgba: sample(8, 8), width: 8, height: 8 }])),
    'picture.ico': Buffer.from(encodeIco([16, 8].map((s) => ({ width: s, height: s, png: encodePng(sample(s, s), s, s) })))),
    'picture.qoi': Buffer.from(encodeQoi(rgba, w, h)),
    'picture.ppm': Buffer.from(pnmOf(rgba, w, h)),
    'picture.pam': Buffer.from(pnmOf(rgba, w, h, { pam: true })),
    'picture.tga': Buffer.from(tgaOf(rgba, w, h)),
    'picture.pcx': Buffer.from(encodePcx(rgba, w, h, { alpha: true })),
    'picture.hdr': Buffer.from(encodeHdr(lightOf(rgba, w, h), w, h)),
    'picture.cur': Buffer.from(encodeIco([{ width: 16, height: 16, png: encodePng(sample(16, 16), 16, 16) }], { cursor: true, hotspots: [{ x: 2, y: 3 }] })),
  };
}
