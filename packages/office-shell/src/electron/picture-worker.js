// The thread a picture Chromium cannot draw is decoded on (protocol.js):
// a 40-megapixel TIFF takes a second or two to unpack, and on the main
// process that second would hold every window still.
//
// It is handed a file's path and the page wanted, reads the file itself,
// and answers a PNG of that page (decodePicture), or the page count, or the
// sentence decodePicture refused with.

import fs from 'node:fs';
import path from 'node:path';
import { parentPort } from 'node:worker_threads';
import { decodePicture, pictureInfo } from '@rutba/imaging/decode';

/** A JPEG inside a TIFF needs the main process's decoder: said so, and the main process decodes that one itself. */
const NEEDS_JPEG = { needsJpeg: true };

parentPort.on('message', ({ id, file, page = 0, fit = 0, info = false }) => {
  try {
    const bytes = fs.readFileSync(file);
    if (info) {
      parentPort.postMessage({ id, info: pictureInfo(bytes, path.basename(file)) });
      return;
    }
    const pic = decodePicture(bytes, path.basename(file), {
      page,
      fit,
      png: true,
      level: 3,
      jpeg: () => {
        throw NEEDS_JPEG;
      },
    });
    const png = pic.png;
    parentPort.postMessage({ id, png, width: pic.width, height: pic.height, pages: pic.pages, page: pic.page }, [png.buffer]);
  } catch (err) {
    if (err === NEEDS_JPEG || err?.cause === NEEDS_JPEG) parentPort.postMessage({ id, needsJpeg: true });
    else parentPort.postMessage({ id, error: err?.picture ? err.message : `This picture could not be read: ${err?.message || err}` });
  }
});
