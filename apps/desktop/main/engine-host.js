// The speech engine add-on, loaded (C6; private design:
// office-pack-compatibility-design.md 1.2 and 2.2).
//
// The engine is the one add-on whose code runs. Its `main` is imported as an
// app pack's is, after addons.js has checked it (the release key, exactly
// this Office's version, its speech and host APIs, every file's hash), and
// its register(context) is handed only what the contract's engineHost.context
// lists: the verified add-ons (list, find) and Electron's utilityProcess, in
// which the engine runs its model. What it answers must be the contract's
// speech namespace with exactly the contract's calls, and nothing else: an
// engine can serve speech and cannot change any other namespace.
//
// It has a time budget at start (C13): an engine that takes too long, throws,
// or answers the wrong namespaces is not used this run, and Read Aloud keeps
// the Windows voice. Nothing here says so in a dialog; the Voices and
// languages page shows why.
//
// The model runs in a process the engine forks, and module hooks belong to
// one process: the utilityProcess it is handed starts every child with
// pack-resolve's hooks imported first and this run's resolution in its
// environment, so the child's imports are held to the engine's folder as the
// engine's own are (pack-resolve.js).

import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { RESOLUTION_IMPORT, WORKER_RESOLUTION } from '@rutba/office-shell/electron/pack-resolve';

/** The engine's import and register() together, at start. */
export const ENGINE_BUDGET_MS = 5000;

const within = (promise, ms, what) => {
  let timer;
  return Promise.race([
    Promise.resolve(promise).finally(() => clearTimeout(timer)),
    new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(`${what} took longer than ${ms / 1000} seconds`)), ms); }),
  ]);
};

/** A path a manifest may name: relative, forward slashes, no step out. */
const safePath = (p) => typeof p === 'string' && p.length > 0 && !p.startsWith('/') && !p.includes('\\') && !/^[a-zA-Z]:/.test(p) && !p.split('/').some((s) => s === '..' || s === '.' || s === '');

/**
 * Electron's utilityProcess as an engine is handed it: fork() alone, each
 * child importing the resolution hooks before anything else, with this run's
 * resolution in its environment whatever environment the engine gives.
 * @param {{ fork: Function }|null|undefined} utilityProcess
 * @param {object} [o]
 * @param {string} [o.hooks]   the module imported first
 * @param {NodeJS.ProcessEnv} [o.env]  this process's environment
 */
export function heldUtilityProcess(utilityProcess, { hooks = RESOLUTION_IMPORT, env = process.env } = {}) {
  if (typeof utilityProcess?.fork !== 'function') return utilityProcess;
  return Object.freeze({
    fork(modulePath, args, options = {}) {
      const given = options || {};
      const told = env[WORKER_RESOLUTION];
      return utilityProcess.fork(modulePath, args, {
        ...given,
        env: { ...(given.env || env), ...(told ? { [WORKER_RESOLUTION]: told } : {}) },
        execArgv: ['--import', hooks, ...(Array.isArray(given.execArgv) ? given.execArgv : [])],
      });
    },
  });
}

/**
 * The context an engine is handed: each dotted name the contract lists,
 * taken from what the core has, and nothing else.
 */
export function engineContext({ registry, electron, names }) {
  const have = {
    addons: { list: () => registry.list(), find: (name) => registry.find(name) },
    electron: { utilityProcess: heldUtilityProcess(electron?.utilityProcess) },
  };
  const out = {};
  for (const name of names || []) {
    const [top, key] = String(name).split('.');
    if (!have[top] || !(key in have[top])) continue;
    out[top] ||= {};
    out[top][key] = have[top][key];
  }
  for (const k of Object.keys(out)) Object.freeze(out[k]);
  return Object.freeze(out);
}

/**
 * @param {object} o
 * @param {object} o.registry     this start's add-ons (addons.js)
 * @param {object} o.contract     the run's add-ons contract
 * @param {object} [o.electron]   Electron's main-process module
 * @param {number} [o.budgetMs]
 * @param {(url: string) => Promise<object>} [o.importer]
 * @returns {Promise<{ namespaces: object, methods: object, engine: string|null, problem: string|null, stop: () => void }>}
 */
export async function loadEngine({ registry, contract, electron = null, budgetMs = ENGINE_BUDGET_MS, importer = (url) => import(url) }) {
  const none = (problem = null, engine = null) => ({ namespaces: {}, methods: {}, engine, problem, stop: () => {} });
  const engine = registry?.engine?.();
  if (!engine) return none();
  const sp = contract.speech || {};
  const main = engine.manifest.main;
  if (!safePath(main) || !Object.hasOwn(engine.manifest.files || {}, main)) return none(`${engine.manifest.name || engine.id} names no main module among its files.`, engine.id);
  const url = pathToFileURL(path.join(engine.dir, ...main.split('/'))).href;
  let service;
  try {
    const mod = await within(importer(url), budgetMs, 'Loading the speech engine');
    if (typeof mod?.register !== 'function') return none('The speech engine has no register().', engine.id);
    const context = engineContext({ registry, electron, names: contract.engineHost?.context });
    const answered = await within(mod.register(context), budgetMs, 'Starting the speech engine');
    const names = Object.keys(answered || {});
    if (names.length !== 1 || names[0] !== sp.namespace) {
      return none(`The speech engine answered ${names.join(', ') || 'nothing'}, not ${sp.namespace} alone.`, engine.id);
    }
    service = answered[sp.namespace];
    const missing = (sp.calls || []).filter((c) => typeof service?.[c] !== 'function');
    if (missing.length) return none(`The speech engine does not answer ${missing.join(', ')}.`, engine.id);
  } catch (err) {
    return none(`The speech engine could not start: ${err?.message || err}`, engine.id);
  }
  // Exactly the contract's calls, each answered as a promise; anything else the service has stays inside it.
  const wrapped = Object.fromEntries((sp.calls || []).map((c) => [c, async (args) => service[c](args)]));
  return {
    namespaces: { [sp.namespace]: wrapped },
    methods: { [sp.namespace]: [...sp.calls] },
    engine: engine.id,
    problem: null,
    stop: () => {
      try {
        service.stop?.();
      } catch {
        // already stopped
      }
    },
  };
}
