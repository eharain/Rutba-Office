// One way to read a picture.
//
// decodePicture(bytes, name) answers a picture's pixels, whatever its
// format, in the main process, Node's workers and the windows alike: pure
// JavaScript, no canvas. Everything that draws a picture Chromium cannot
// (Pictures' stage over rutba://file, Image's open, a picture inside a
// document) asks here, and gets either the picture or a sentence a person
// can act on.
//
// Every reader takes untrusted files, so each is bounded before it
// allocates: the file's size, the number of pages, and each page's pixels
// are counted against LIMITS first, and a damaged file is a refusal, never a
// hang or a crash (tools/fuzz-open.js holds them to it).
//
// The formats Chromium draws itself (JPEG, GIF, WebP, AVIF, SVG) are not
// decoded here: a caller can lend a decoder for them (`system`, Windows' own
// codecs; `jpeg`, for a JPEG inside a TIFF), and without one they are
// refused in a sentence.

import { PictureError, bad, asBytes, ascii, LIMITS } from './codecs/bytes.js';
import { isPng, pngInfo, decodePng, encodePng } from './codecs/png.js';
import { isBmp, bmpInfo, decodeBmp } from './codecs/bmp.js';

export { LIMITS, PictureError, encodePng };

const extOf = (name) => {
  const m = /\.([a-z0-9]+)$/i.exec(String(name || ''));
  return m ? m[1].toLowerCase() : '';
};

/**
 * The readers, by format: how to tell one from its first bytes, its name
 * when it has no mark of its own, what it is without decoding it, and its
 * pixels.
 */
const READERS = {
  png: { exts: ['png', 'apng'], sniff: isPng, info: pngInfo, decode: (b, o) => decodePng(b, o.limits) },
  bmp: { exts: ['bmp', 'dib'], sniff: isBmp, info: bmpInfo, decode: (b, o) => decodeBmp(b, o) },
};

/** Formats Chromium draws in an <img> itself: a window shows these from the file. */
export const BROWSER_DRAWS = new Set(['png', 'jpeg', 'gif', 'webp', 'avif', 'bmp', 'ico', 'svg']);

/** What a file's first bytes say it is, for formats nobody here reads. */
function foreign(b) {
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'jpeg';
  if (ascii(b, 0, 4) === 'GIF8') return 'gif';
  if (ascii(b, 0, 4) === 'RIFF' && ascii(b, 8, 4) === 'WEBP') return 'webp';
  if (ascii(b, 4, 4) === 'ftyp') {
    const brand = ascii(b, 8, 4);
    return /avif|avis/.test(brand) ? 'avif' : /hei|hev|mif1|msf1/.test(brand) ? 'heic' : null;
  }
  if (b[0] === 0x49 && b[1] === 0x49 && b[2] === 0xbc) return 'jxr';
  if ((b[0] === 0xff && b[1] === 0x0a) || ascii(b, 4, 8) === 'JXL \r\n\x87\n') return 'jxl';
  if (ascii(b, 4, 8) === 'jP  \r\n\x87\n' || (b[0] === 0xff && b[1] === 0x4f && b[2] === 0xff && b[3] === 0x51)) return 'jp2';
  if (/<svg[\s>]/i.test(ascii(b, 0, 1024))) return 'svg';
  return null;
}

/**
 * The format of a picture: its bytes first, its name where the format has
 * no mark of its own (TGA), or null.
 */
export function pictureFormat(bytes, name = '') {
  const b = asBytes(bytes);
  for (const [format, rd] of Object.entries(READERS)) if (rd.sniff?.(b, name)) return format;
  const other = foreign(b);
  if (other) return other;
  const ext = extOf(name);
  for (const [format, rd] of Object.entries(READERS)) if (!rd.sniff && rd.exts.includes(ext)) return format;
  return null;
}

/** The formats read here, and the extensions they go by. */
export function pictureFormats() {
  return Object.fromEntries(Object.entries(READERS).map(([f, rd]) => [f, rd.exts.slice()]));
}

/** Whether a file of this name is one a window cannot draw by itself and must be served decoded. */
export function needsDecoding(name) {
  const ext = extOf(name);
  for (const [format, rd] of Object.entries(READERS)) if (rd.exts.includes(ext)) return !BROWSER_DRAWS.has(format);
  return false;
}

const NAMES = { jpeg: 'JPEG', gif: 'GIF', webp: 'WebP', avif: 'AVIF', heic: 'HEIC', jxr: 'JPEG XR', jxl: 'JPEG XL', jp2: 'JPEG 2000', svg: 'SVG' };

function cannot(format, name) {
  const what = format ? `${NAMES[format] || format.toUpperCase()} pictures` : `${extOf(name) ? `.${extOf(name)} files` : 'this kind of file'}`;
  return bad(`Rutba Office cannot draw ${what} here.`);
}

/** Size and pages, from the headers only. */
export function pictureInfo(bytes, name = '', opts = {}) {
  const b = asBytes(bytes);
  const format = pictureFormat(b, name);
  const rd = READERS[format];
  if (!rd) cannot(format, name);
  return guarded(format, () => ({ pages: 1, ...rd.info(b, { limits: { ...LIMITS, ...(opts.limits || {}) } }) }));
}

/** A thrown error that is not a refusal is a fault of the reader; it is said as a refusal, the fault kept as its cause. */
function guarded(format, run) {
  try {
    return run();
  } catch (err) {
    if (err instanceof PictureError || err?.picture) throw err;
    const what = NAMES[format] || String(format || 'picture').toUpperCase();
    throw new PictureError(`This ${what} picture could not be read: it may be damaged.`, { cause: err });
  }
}

/* ── turning, shrinking ────────────────────────────────────────────────── */

/** Pixels turned upright as an EXIF or TIFF orientation (1-8) says. */
export function orient(rgba, width, height, orientation) {
  if (!(orientation >= 2 && orientation <= 8)) return { rgba, width, height };
  const swap = orientation >= 5;
  const w = swap ? height : width;
  const h = swap ? width : height;
  const out = new Uint8Array(rgba.length);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      let nx;
      let ny;
      switch (orientation) {
        case 2: nx = width - 1 - x; ny = y; break;
        case 3: nx = width - 1 - x; ny = height - 1 - y; break;
        case 4: nx = x; ny = height - 1 - y; break;
        case 5: nx = y; ny = x; break;
        case 6: nx = height - 1 - y; ny = x; break;
        case 7: nx = height - 1 - y; ny = width - 1 - x; break;
        default: nx = y; ny = width - 1 - x; break; // 8
      }
      const s = (y * width + x) * 4;
      const d = (ny * w + nx) * 4;
      out[d] = rgba[s];
      out[d + 1] = rgba[s + 1];
      out[d + 2] = rgba[s + 2];
      out[d + 3] = rgba[s + 3];
    }
  }
  return { rgba: out, width: w, height: h };
}

/** Pixels brought inside `max` on their longer side by averaging (a box filter), never enlarged. */
export function shrink(rgba, width, height, max) {
  const s = Math.max(width, height) / max;
  if (!(s > 1)) return { rgba, width, height };
  const w = Math.max(1, Math.round(width / s));
  const h = Math.max(1, Math.round(height / s));
  const out = new Uint8Array(w * h * 4);
  for (let y = 0; y < h; y++) {
    const y0 = Math.floor((y * height) / h);
    const y1 = Math.max(y0 + 1, Math.floor(((y + 1) * height) / h));
    for (let x = 0; x < w; x++) {
      const x0 = Math.floor((x * width) / w);
      const x1 = Math.max(x0 + 1, Math.floor(((x + 1) * width) / w));
      let r = 0;
      let g = 0;
      let b = 0;
      let a = 0;
      for (let yy = y0; yy < y1; yy++) {
        for (let xx = x0; xx < x1; xx++) {
          const i = (yy * width + xx) * 4;
          const al = rgba[i + 3];
          r += rgba[i] * al;
          g += rgba[i + 1] * al;
          b += rgba[i + 2] * al;
          a += al;
        }
      }
      const d = (y * w + x) * 4;
      const n = (y1 - y0) * (x1 - x0);
      if (a) {
        out[d] = Math.round(r / a);
        out[d + 1] = Math.round(g / a);
        out[d + 2] = Math.round(b / a);
      }
      out[d + 3] = Math.round(a / n);
    }
  }
  return { rgba: out, width: w, height: h };
}

/* ── the one door ──────────────────────────────────────────────────────── */

/**
 * A picture's pixels.
 *
 * @param {Uint8Array | ArrayBuffer | string} bytes the file (base64 accepted)
 * @param {string} [name] its name, for the formats with no mark of their own
 * @param {object} [opts]
 * @param {number} [opts.page] which page, from 0 (a multi-page TIFF, an icon's sizes)
 * @param {number} [opts.fit] bring the longer side within this many pixels
 * @param {boolean} [opts.png] answer a PNG of it too
 * @param {boolean} [opts.orient] turn it upright by its orientation tag (default true)
 * @param {object} [opts.limits] LIMITS to hold it to instead
 * @param {(bytes: Uint8Array, name: string) => (Uint8Array | null)} [opts.system]
 *   a decoder lent for formats not read here (Windows' own codecs): bytes in, a PNG out, or null
 * @param {(bytes: Uint8Array) => ({ width, height, rgba } | null)} [opts.jpeg]
 *   a JPEG decoder lent for JPEG inside a TIFF
 * @returns {{ format, width, height, pages, page, rgba: Uint8Array, png?: Uint8Array, exif: object }}
 */
export function decodePicture(bytes, name = '', opts = {}) {
  const b = asBytes(bytes);
  const limits = { ...LIMITS, ...(opts.limits || {}) };
  if (b.length > limits.bytes) bad(`This picture is ${Math.round(b.length / 1048576)} MB, larger than Rutba Office reads.`);
  if (b.length < 8) bad('This file is too short to be a picture.');
  const format = pictureFormat(b, name);
  const page = Math.max(0, Math.floor(Number(opts.page) || 0));
  const rd = READERS[format];
  let pic;
  if (rd) {
    pic = guarded(format, () => rd.decode(b, { page, limits, jpeg: opts.jpeg || null }));
  } else {
    const png = typeof opts.system === 'function' ? opts.system(b, name) : null;
    if (!png) cannot(format, name);
    pic = guarded('png', () => decodePng(asBytes(png), limits));
  }
  let { rgba, width, height } = pic;
  if (opts.orient !== false && pic.orientation > 1) ({ rgba, width, height } = orient(rgba, width, height, pic.orientation));
  if (opts.fit > 0) ({ rgba, width, height } = shrink(rgba, width, height, opts.fit));
  const out = { format: format || 'system', width, height, pages: pic.pages || 1, page, rgba, exif: pic.exif || {} };
  if (opts.png) out.png = encodePng(rgba, width, height, { level: opts.level ?? 6 });
  return out;
}
