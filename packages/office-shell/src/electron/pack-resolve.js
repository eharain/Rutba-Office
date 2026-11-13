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
//
// An add-on whose code runs (the speech engine) is held tighter: it carries
// everything it imports, and is handed what it may have of the core in its
// register() context. So a module inside an add-ons folder resolves only to
// files inside its own add-on's folder (or Node's built-ins); a lookup that
// leaves it, by a bare name climbing to a node_modules above or by a relative
// path, is refused, and the core is not asked. Node names a module by its
// real path, so a folder reached through a junction or a link is held by
// its real path too.

import fs from 'node:fs';
import module from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

/** A specifier that names a package rather than a file or a URL. */
export const isBare = (s) => !/^(\.{1,2}\/|\/|[a-zA-Z]:[\\/]|node:|file:|data:)/.test(s) && s !== 'electron';

const folderUrl = (dir) => pathToFileURL(path.resolve(dir) + path.sep).href.toLowerCase();

/** The packs folder and the core's main module, once a run has packs that carry code. */
let state = null;
/** Add-ons folders, as given, and the URL prefixes each is held by (as given and real). */
const closed = new Map();
let hooked = false;

/** Where a worker thread finds what this thread was told, the environment being copied to it. */
export const WORKER_RESOLUTION = 'RUTBA_PACK_RESOLUTION';
/** The worker options this run's threads are started with, read by name so that a pack need not import the core for them. */
export const WORKER_ARGV = '__rutbaPackWorkerArgv';
/** What a thread or a process of this run imports first, so that its imports are held as this one's are. */
export const RESOLUTION_IMPORT = new URL('./pack-resolve-worker.js', import.meta.url).href;

/** The add-on folder a module URL is inside of, as a URL prefix, or null. */
function addonHome(from) {
  for (const prefixes of closed.values()) {
    for (const prefix of prefixes) {
      if (!from.startsWith(prefix)) continue;
      const id = from.slice(prefix.length).split('/')[0];
      // A file lying in the add-ons folder itself belongs to no add-on: it has nothing to import.
      return id && from.length > prefix.length + id.length ? `${prefix}${id}/` : prefix + '\0';
    }
  }
  return null;
}

/**
 * Whether a module file is inside an add-ons folder this run holds, by its
 * real path as Node will name it: so an engine is imported only when what it
 * imports will be held.
 * @param {string} file
 */
export function isHeld(file) {
  let real;
  try {
    real = fs.realpathSync.native(file);
  } catch {
    return false;
  }
  return addonHome(pathToFileURL(real).href.toLowerCase()) !== null;
}

const outside = (specifier) => Object.assign(
  new Error(`Cannot import ${specifier}: an add-on imports only what it carries in its own folder`),
  { code: 'ERR_MODULE_NOT_FOUND' },
);

// Hooks are per thread, and a worker's imports run before any of its code: a
// thread a pack or an add-on starts is told to install these first.
function tellWorkers() {
  process.env[WORKER_RESOLUTION] = JSON.stringify({ ...(state?.given || {}), closed: [...closed.keys()] });
  globalThis[WORKER_ARGV] = ['--import', RESOLUTION_IMPORT];
}

function hook() {
  if (hooked) return;
  hooked = true;
  module.registerHooks({
    resolve(specifier, context, nextResolve) {
      const from = String(context.parentURL || '').toLowerCase();
      const home = closed.size ? addonHome(from) : null;
      if (home) {
        const found = nextResolve(specifier, context);
        if (found.url?.startsWith('file:') && !found.url.toLowerCase().startsWith(home)) throw outside(specifier);
        return found;
      }
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

/**
 * From now on, a module inside `packsDir` that imports a package it cannot
 * find gets the core's: resolved from `coreFile`. Once per run; a second call
 * names another folder (a check run's).
 * @param {{ packsDir: string, coreFile: string }} o
 */
export function installPackResolution({ packsDir, coreFile }) {
  state = { prefix: folderUrl(packsDir), core: pathToFileURL(coreFile).href, given: { packsDir, coreFile } };
  tellWorkers();
  hook();
}

/**
 * From now on, a module inside an add-on in `dir` (the add-ons folder) finds
 * only files of its own add-on. Before the engine is imported; a folder named
 * twice is held once.
 * @param {string} dir
 */
export function closeAddonsFolder(dir) {
  const given = path.resolve(dir);
  let real = given;
  try {
    real = fs.realpathSync.native(given);
  } catch {
    // not made yet: held as given
  }
  closed.set(given, [...new Set([folderUrl(given), folderUrl(real)])]);
  tellWorkers();
  hook();
}
