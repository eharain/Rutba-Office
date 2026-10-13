// The recogniser itself, as one thread holds it: Tesseract 5 in WebAssembly
// (vendor/core, the engine tesseract.js 7 ships) with the fast models in
// vendor/tessdata, both read from the suite's own files. Nothing here can
// reach the network: the engine is handed its WebAssembly as bytes and its
// models are written into its own memory file system, so it never looks for
// either anywhere else.
//
// Synchronous once started: a thread reads one page at a time.

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { pgm } from './pixels.js';
import { modelOf } from './languages.js';

const BUILD = 'tesseract-core-simd-lstm';

/** Inside an installed app the engine and models sit beside the archive, unpacked. */
function unpacked(p) {
  const u = p.replace(/app\.asar([\\/])/, 'app.asar.unpacked$1');
  return u !== p && fs.existsSync(u) ? u : p;
}

export const vendorDir = () => unpacked(fileURLToPath(new URL('../vendor/', import.meta.url)));

/** Page segmentation: let the engine find the blocks and lines (3), or read one block (6), one line (7). */
export const PSM = Object.freeze({ auto: 3, block: 6, line: 7, sparse: 11 });

export async function createEngine({ vendor = vendorDir() } = {}) {
  const require = createRequire(import.meta.url);
  const Core = require(path.join(vendor, 'core', `${BUILD}.cjs`));
  const tess = await Core({
    wasmBinary: fs.readFileSync(path.join(vendor, 'core', `${BUILD}.wasm`)),
    print() {},
    printErr() {},
  });
  const written = new Set();
  let api = null;
  let current = '';

  function use(languages) {
    const key = languages.map(modelOf).join('+');
    if (key === current && api) return;
    for (const model of languages.map(modelOf)) {
      if (written.has(model)) continue;
      tess.FS.writeFile(`/${model}.traineddata`, fs.readFileSync(path.join(vendor, 'tessdata', `${model}.traineddata`)));
      written.add(model);
    }
    if (api) {
      api.End();
      tess.destroy(api);
    }
    api = new tess.TessBaseAPI();
    // 1: the LSTM recogniser, the only one the fast models carry.
    if (api.Init('/', key, 1) !== 0) {
      api = null;
      current = '';
      throw Object.assign(new Error(`The text recognition model for ${key} could not be loaded.`), { code: 'OCR_MODEL' });
    }
    current = key;
  }

  return {
    build: BUILD,
    /** The models loaded now, by file name. */
    get models() { return current ? current.split('+') : []; },

    /**
     * Read a grey picture. @returns the engine's JSON (blocks, paragraphs,
     * lines, words with bbox and confidence) and its plain text.
     */
    read(grey, { languages, dpi = 0, psm = PSM.auto } = {}) {
      use(languages);
      tess.FS.writeFile('/input', pgm(grey));
      api.SetImageFile(1, 0);
      if (dpi) api.SetSourceResolution(Math.round(dpi));
      api.SetPageSegMode(psm);
      api.Recognize(null);
      const json = JSON.parse(api.GetJSONText());
      const text = api.GetUTF8Text();
      api.Clear();
      return { json, text };
    },

    end() {
      if (api) {
        api.End();
        tess.destroy(api);
      }
      api = null;
      current = '';
    },
  };
}
