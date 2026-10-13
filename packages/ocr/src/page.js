// One recognition, start to finish, on the thread that holds the engine:
// the picture as it came (pixels, or a PNG file's bytes) made grey, read,
// and the answer put back into the picture's own coordinates.

import { checkLanguages } from './languages.js';
import { isPng, decodePng } from './png.js';
import { toGrey } from './pixels.js';
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

/**
 * @param engine from createEngine
 * @param {{ image, languages?, dpi?, mode?: 'auto'|'block'|'line'|'sparse' }} job
 */
export function recognizePage(engine, { image, languages, dpi = 0, mode = 'auto' }) {
  const langs = checkLanguages(languages);
  const started = Date.now();
  const pixels = picture(image);
  const grey = toGrey(pixels);
  const { json } = engine.read(grey, { languages: langs, dpi, psm: PSM[mode] ?? PSM.auto });
  const result = fromEngine(json, { width: pixels.width, height: pixels.height, languages: langs });
  result.ms = Date.now() - started;
  return result;
}
