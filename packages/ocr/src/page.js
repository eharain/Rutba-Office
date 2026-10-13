// One recognition, start to finish, on the thread that holds the engine:
// the picture as it came (pixels, or a PNG file's bytes) made grey, looked
// at by our own analysis (analyse.js), prepared, read, and the answer put
// back into the picture's own coordinates.
//
// The preparation, each step skipped when it would change nothing:
//   turned back level when the text is turned by two degrees or more (the
//     engine reads a smaller tilt as well itself, and turning a soft scan
//     by resampling it costs more than the tilt: docs/OCR.md);
//   scaled when the text is very small (a screenshot) or very large (a poster,
//     a 600 dpi scan), towards the height the models read best;
//   handed over as ink and paper (Sauvola) when the paper is lit unevenly,
//     and otherwise as grey, which the engine thresholds itself.

import { checkLanguages } from './languages.js';
import { isPng, decodePng } from './png.js';
import { toGrey } from './pixels.js';
import { binarise, estimateSkew, rotate, unrotatePoint, scale, textHeight, unevenPaper } from './analyse.js';
import { fromEngine } from './results.js';
import { PSM } from './engine.js';

/** Pixels from whatever was handed over. */
export function picture(image) {
  if (image instanceof Uint8Array || image instanceof ArrayBuffer) {
    const bytes = image instanceof ArrayBuffer ? new Uint8Array(image) : image;
    if (isPng(bytes)) return decodePng(bytes);
    throw Object.assign(new Error('Text can be read from pixels or a PNG picture.'), { code: 'OCR_IMAGE' });
  }
  const { width, height, data } = image || {};
  if (!(width > 0 && height > 0 && data && data.length >= width * height)) {
    throw Object.assign(new Error('The picture to read has no pixels.'), { code: 'OCR_IMAGE' });
  }
  return { width, height, data: data instanceof Uint8Array ? data : new Uint8Array(data.buffer ?? data, data.byteOffset ?? 0, data.length), channels: image.channels };
}

/** Text smaller than this (in pixels, a lower-case letter's height) is scaled up, larger than LARGE down. */
const SMALL = 10;
const LARGE = 64;
const TARGET = 20;
/** A page turned by this many degrees or more is turned back before reading. */
const TURN = 2;
/** Analysis looks at a copy of at most this many pixels; the answer does not need more. */
const LOOK = 3e6;

/**
 * What to do to a grey picture before reading it: the angle to turn it back,
 * the factor to scale it by, and whether to hand over ink and paper.
 */
export function plan(grey, { deskew = true, rescale = true } = {}) {
  const shrink = Math.min(1, Math.sqrt(LOOK / (grey.width * grey.height)));
  const look = shrink < 1 ? scale(grey, shrink) : grey;
  const uneven = unevenPaper(look);
  const binary = binarise(look, { method: uneven ? 'sauvola' : 'otsu' });
  const angle = deskew ? estimateSkew(binary) : 0;
  const height = textHeight(binary) / shrink;
  let factor = 1;
  if (rescale && height > 0 && height < SMALL) factor = Math.min(4, TARGET / height);
  else if (rescale && height > LARGE) factor = Math.max(0.25, TARGET * 1.6 / height);
  return { angle, turn: Math.abs(angle) >= TURN ? angle : 0, factor: Math.round(factor * 100) / 100, uneven, textHeight: Math.round(height) };
}

/**
 * @param engine from createEngine
 * @param {{ image, languages?, dpi?, mode?: 'auto'|'block'|'line'|'sparse', deskew?: boolean, rescale?: boolean }} job
 */
export function recognizePage(engine, { image, languages, dpi = 0, mode = 'auto', deskew = true, rescale = true }) {
  const langs = checkLanguages(languages);
  const started = Date.now();
  const pixels = picture(image);
  const grey = toGrey(pixels);
  const how = plan(grey, { deskew, rescale });
  let prepared = how.turn ? rotate(grey, -how.turn) : grey;
  if (how.factor !== 1) prepared = scale(prepared, how.factor);
  if (how.uneven) prepared = binarise(prepared, { method: 'sauvola' });
  const { json } = engine.read(prepared, { languages: langs, dpi: dpi ? dpi * how.factor : 0, psm: PSM[mode] ?? PSM.auto });
  // A box in what was read, back in the picture: unscaled, then turned as the picture is turned.
  const map = (b) => {
    const x0 = b.x / how.factor, y0 = b.y / how.factor, x1 = (b.x + b.width) / how.factor, y1 = (b.y + b.height) / how.factor;
    if (!how.turn) return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
    const corners = [[x0, y0], [x1, y0], [x0, y1], [x1, y1]].map(([x, y]) => unrotatePoint(x, y, -how.turn, grey.width, grey.height));
    const xs = corners.map((c) => c.x), ys = corners.map((c) => c.y);
    return { x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) };
  };
  const result = fromEngine(json, { width: pixels.width, height: pixels.height, languages: langs, angle: how.angle, map });
  result.prepared = { turned: how.turn, factor: how.factor, binarised: how.uneven, textHeight: how.textHeight };
  result.ms = Date.now() - started;
  return result;
}
