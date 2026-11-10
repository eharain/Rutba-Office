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
// Node's own search does not stop at the pack: it climbs every folder above
// the packs folder. A node_modules up there (a build run inside this
// repository, a developer's home folder) would hand the pack a second copy of
// a core module, whose top level then runs again. So what that search finds
// outside the packs folder is thrown away, and the core is asked instead.
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

/** Where a worker thread finds what this thread was told, the environment being copied to it. */
export const WORKER_RESOLUTION = 'RUTBA_PACK_RESOLUTION';
/** The worker options this run's threads are started with, read by name so that a pack need not import the core for them. */
export const WORKER_ARGV = '__rutbaPackWorkerArgv';

/**
 * From now on, a module inside `packsDir` that imports a package it cannot
 * find gets the core's: resolved from `coreFile`. Once per run; a second call
 * names another folder (a check run's).
 * @param {{ packsDir: string, coreFile: string }} o
 */
export function installPackResolution({ packsDir, coreFile }) {
  const next = { prefix: folderUrl(packsDir), core: pathToFileURL(coreFile).href };
  // Hooks are per thread, and a worker's imports run before any of its code:
  // a thread a pack starts is told to install these first (see packWorkerOptions).
  process.env[WORKER_RESOLUTION] = JSON.stringify({ packsDir, coreFile });
  globalThis[WORKER_ARGV] = ['--import', new URL('./pack-resolve-worker.js', import.meta.url).href];
  if (state) {
    Object.assign(state, next);
    return;
  }
  state = next;
  module.registerHooks({
    resolve(specifier, context, nextResolve) {
      const from = String(context.parentURL || '').toLowerCase();
      const inPack = Boolean(state) && isBare(specifier) && from.startsWith(state.prefix);
      // A require() is resolved from the module that made it whatever the
      // context says, so the core's own require answers it.
      const fromCore = () => (context.conditions?.includes('require')
        ? { url: pathToFileURL(module.createRequire(state.core).resolve(specifier)).href, shortCircuit: true }
        : nextResolve(specifier, { ...context, parentURL: state.core }));
      let found;
      try {
        found = nextResolve(specifier, context);
      } catch (err) {
        if (!inPack) throw err;
        return fromCore();
      }
      if (!inPack || !found.url?.startsWith('file:') || found.url.toLowerCase().startsWith(state.prefix)) return found;
      return fromCore();
    },
  });
}
