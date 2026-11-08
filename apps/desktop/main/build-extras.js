// Code a build carries only when it was made with a switch that names it.
//
// build/bundle.js writes the list (build/out/extras.json) as it builds; a build
// made without a switch lists nothing and leaves the extras' code out of the
// archive altogether (electron-builder.config.cjs). So there is nothing here
// that a person can turn on: this file only loads what the build put in.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const LIST = path.join(here, '..', 'build', 'out', 'extras.json');

function listed() {
  try {
    const { extras } = JSON.parse(fs.readFileSync(LIST, 'utf8'));
    return Array.isArray(extras) ? extras : [];
  } catch {
    return [];
  }
}

let loading = null;

/** The extras this build carries, loaded once: each a module with describe() and register(). */
export function buildExtras() {
  loading ||= (async () => {
    const out = [];
    for (const name of listed()) {
      try {
        out.push(await import(pathToFileURL(path.join(here, 'extras', name, 'index.js')).href));
      } catch (err) {
        console.error(`[extras] ${name}: ${err?.message || err}`);
      }
    }
    return out;
  })();
  return loading;
}
