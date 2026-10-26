// Where a pack's modules find what they import (docs/APP-PACKS.md 5.1.7).
//
// A pack's main module and the files beside it live in the pack's own
// archive, in the packs folder. What they import by a relative path is the
// pack's own; a package the pack carries is in its own node_modules, which
// Node finds by itself, walking up from the importing file. A package the
// core has and the pack does not (the shared engines, the office-shell
// package, React's server parts if ever) is not under any folder Node walks
// from there, so an import the pack's own folders cannot answer is answered
// as if the core's main module had made it: the core's copy, the one every
// other module of the run shares.
//
// Only bare names are looked for in the core, and only for a module inside
// the packs folder. A relative import a pack cannot answer stays an error:
// it names a file the pack should have carried.

import module from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

/** A specifier that names a package rather than a file or a URL. */
export const isBare = (s) => !/^(\.{1,2}\/|\/|[a-zA-Z]:[\\/]|node:|file:|data:)/.test(s) && s !== 'electron';

const folderUrl = (dir) => pathToFileURL(path.resolve(dir) + path.sep).href.toLowerCase();

let state = null;

/**
 * From now on, a module inside `packsDir` that imports a package it cannot
 * find gets the core's: resolved from `coreFile`. Once per run; a second call
 * names another folder (a check run's).
 * @param {{ packsDir: string, coreFile: string }} o
 */
export function installPackResolution({ packsDir, coreFile }) {
  const next = { prefix: folderUrl(packsDir), core: pathToFileURL(coreFile).href };
  if (state) {
    Object.assign(state, next);
    return;
  }
  state = next;
  module.registerHooks({
    resolve(specifier, context, nextResolve) {
      try {
        return nextResolve(specifier, context);
      } catch (err) {
        const from = String(context.parentURL || '').toLowerCase();
        if (!isBare(specifier) || !state || !from.startsWith(state.prefix)) throw err;
        return nextResolve(specifier, { ...context, parentURL: state.core });
      }
    },
  });
}
