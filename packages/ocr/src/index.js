// @rutba/ocr: the text in a picture, read on this computer.
//
//   import { recognize } from '@rutba/ocr';
//   const { text, lines, words } = await recognize(pixels, { languages: ['eng'] });
//
// Pixels are { width, height, data } (a canvas's ImageData will do) or a PNG
// file's bytes. Recognition runs on a pool of threads (host.js), never on
// the caller's; the engine and its six models ship inside the suite and
// nothing is fetched. See docs/OCR.md.

import { createOcrHost } from './host.js';

export { createOcrHost, defaultThreads, OcrCancelledError } from './host.js';
export { LANGUAGES, CODES, UNSPACED, checkLanguages, languagesFor, OcrLanguageError } from './languages.js';

let shared = null;

/** The one pool a process shares. */
export function ocrHost() {
  shared ||= createOcrHost();
  return shared;
}

/** Read the text in a picture on the shared pool. */
export function recognize(image, options) {
  return ocrHost().recognize(image, options);
}
