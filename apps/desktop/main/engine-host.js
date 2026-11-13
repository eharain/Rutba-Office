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
// one process: the utilityProcess it is handed starts every child as the
// core's own first module (pack-resolve-child.js), with this run's resolution
// and the engine's module in its environment. That file puts the hooks in and
// only then imports the engine's module, so the child's imports are held to
// the engine's folder as the engine's own are (pack-resolve.js), with nothing
// resting on execArgv. A module outside a held add-on is not started at all.

import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { CHILD_BOOTSTRAP, CHILD_MODULE, WORKER_RESOLUTION, isHeld } from '@rutba/office-shell/electron/pack-resolve';
import { safePath } from '@rutba/office-shell/electron/addons-compat';

/** The engine's import and register() together, at start. */
export const ENGINE_BUDGET_MS = 5000;

const within = (promise, ms, what) => {
  let timer;
  return Promise.race([
    Promise.resolve(promise).finally(() => clearTimeout(timer)),
    new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(`${what} took longer than ${ms / 1000} seconds`)), ms); }),
  ]);
};

/**
 * Electron's utilityProcess as an engine is handed it: fork() alone. Each
 * child starts as the core's bootstrap, which holds its imports and then
 * imports the module asked for; this run's resolution is in its environment
 * whatever environment the engine gives. What an engine asks Node to run
 * before its module (execArgv, NODE_OPTIONS) is not passed on: nothing runs
 * before the hooks. A module that is not inside a held add-on is refused.
 * @param {{ fork: Function }|null|undefined} utilityProcess
 * @param {object} [o]
 * @param {string} [o.bootstrap]  the file every child starts as
 * @param {NodeJS.ProcessEnv} [o.env]  this process's environment
 * @param {(file: string) => boolean} [o.held]
 */
export function heldUtilityProcess(utilityProcess, { bootstrap = CHILD_BOOTSTRAP, env = process.env, held = isHeld } = {}) {
  if (typeof utilityProcess?.fork !== 'function') return utilityProcess;
  return Object.freeze({
    fork(modulePath, args, options = {}) {
      const asked = modulePath instanceof URL || /^file:/i.test(String(modulePath)) ? fileURLToPath(modulePath) : String(modulePath);
      const target = path.resolve(asked);
      const told = env[WORKER_RESOLUTION];
      if (!told || !held(target)) throw new Error(`${target} is not inside an add-on whose imports are held, so it was not started.`);
      const { execArgv: _execArgv, ...given } = options || {};
      // Windows names in its environment are the same in any case: every spelling of NODE_OPTIONS goes.
      const childEnv = Object.fromEntries(Object.entries(given.env || env).filter(([k]) => k.toUpperCase() !== 'NODE_OPTIONS'));
      return utilityProcess.fork(bootstrap, args, {
        ...given,
        env: { ...childEnv, [WORKER_RESOLUTION]: told, [CHILD_MODULE]: target },
        execArgv: [],
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
 * @param {(file: string) => boolean} [o.held]  whether the module is in an add-ons folder whose imports are held (pack-resolve.js)
 * @returns {Promise<{ namespaces: object, methods: object, engine: string|null, problem: string|null, stop: () => void }>}
 */
export async function loadEngine({ registry, contract, electron = null, budgetMs = ENGINE_BUDGET_MS, importer = (url) => import(url), held = isHeld }) {
  const none = (problem = null, engine = null) => ({ namespaces: {}, methods: {}, engine, problem, stop: () => {} });
  const engine = registry?.engine?.();
  if (!engine) return none();
  const sp = contract.speech || {};
  const main = engine.manifest.main;
  if (!safePath(main) || !Object.hasOwn(engine.manifest.files || {}, main)) return none(`${engine.manifest.name || engine.id} names no main module among its files.`, engine.id);
  const file = path.join(engine.dir, ...main.split('/'));
  // Not imported unless what it imports will be held to its folder: closed shut, never open.
  if (!held(file)) return none(`${engine.manifest.name || engine.id} is not in an add-ons folder whose imports are held, so it was not started.`, engine.id);
  const url = pathToFileURL(file).href;
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
