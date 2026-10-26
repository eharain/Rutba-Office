// The apps as packs, on the core's side (docs/APP-PACKS.md 5.1): which apps
// this run has, what each of them offers, and how they are registered.
//
// Every app's code is still in the core's archive in this version, beside a
// register module of its own (<id>-pack.js, BUILT_IN below). Which of them
// run is decided here. With no packs folder (a copy run from the source, the
// portable build) every app is built in, as before. With one (an installed
// copy: resources/packs, written by the installer and by Add or remove apps;
// or the folder a check run names) only the packs found there and signed
// with a key the core holds are loaded, and an app with no pack is absent: no
// IPC channel, no window, no tile, no file type, no AI tool. A pack's
// manifest names its main module; one whose code is still in the core
// registers the core's own (`builtIn` in the context it is handed), one that
// carries its code registers that.
//
// A pack may not declare a namespace the core has, or one another pack owns
// (the contract's NAMESPACE_PACK); such a pack is refused with the others.
// One whose register fails is refused too, with what it said, and its
// channels are left out.

import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { PACKS, NAMESPACE_PACK, packIpc, packEvents, coreMethods, coreEvents } from '@rutba/office-shell/contract';
import { findPacks, publicKeyOf } from '@rutba/office-shell/electron/packs';
import { applyRemovals, pendingRemovals } from '@rutba/office-shell/electron/pack-install';
import { APPS, APP_EXTENSIONS, OLD_EXTENSIONS, WRITABLE, NEW_DOCUMENTS, fileAssociations } from '@rutba/office-formats/registry';
import { TOOL_PACK, toolsOf } from './pack-tools.js';

/** Each app's register module, in the core's archive in this version. */
export const BUILT_IN = {
  word: () => import('./word-pack.js'),
  sheets: () => import('./sheets-pack.js'),
  slides: () => import('./slides-pack.js'),
  pdf: () => import('./pdf-pack.js'),
  pictures: () => import('./pictures-pack.js'),
  image: () => import('./image-pack.js'),
  video: () => import('./video-pack.js'),
  recorder: () => import('./recorder-pack.js'),
  mail: () => import('./mail-pack.js'),
  calendar: () => import('./calendar-pack.js'),
  contacts: () => import('./contacts-pack.js'),
  base: () => import('./base-pack.js'),
  talk: () => import('./talk-pack.js'),
  ai: () => import('./ai-pack.js'),
  ocr: () => import('./ocr-pack.js'),
};

/**
 * The packs whose features a pack offers when they are installed and hides
 * when not (the map's calls and windows across packs): never a reason to
 * refuse it.
 */
const USES = {
  word: ['contacts', 'mail'],
  sheets: ['ocr', 'pdf', 'pictures'],
  pdf: ['ocr'],
  pictures: ['image', 'video', 'pdf', 'ocr'],
  image: ['ocr'],
  recorder: ['video', 'pictures', 'image'],
  mail: ['calendar', 'contacts', 'pdf', 'ocr'],
  calendar: ['mail'],
  contacts: ['mail'],
  talk: ['contacts', 'mail'],
  ai: ['ocr', 'word', 'sheets', 'slides', 'pdf'],
};

/** The icon a kind of file wears: PDF's own wherever it is claimed, the app's otherwise. */
export const fileTypeIcon = (ext, owner) => (ext === 'pdf' ? 'pdf' : owner);

/**
 * What a pack is, as its manifest says it: the build writes this into each
 * pack (build/make-app-packs.js), and a core with every app built in reads it
 * from here, so the two cannot differ.
 * @param {string} id
 * @param {{ toolNames?: string[] }} [o] every AI tool's name, to say which are AI Tools' own
 */
export function describePack(id, { toolNames = Object.keys(TOOL_PACK) } = {}) {
  const known = PACKS.find((p) => p.id === id);
  if (!known) throw new Error(`There is no app pack called ${id}.`);
  const app = APPS[id] || null;
  return {
    id,
    name: known.name,
    needs: [],
    uses: USES[id] || [],
    ipc: packIpc(id),
    events: packEvents(id),
    tiles: app ? [{ app: id, icon: `tiles/${id}.ico` }] : [],
    fileTypes: fileAssociations({ apps: [id] }).map(({ app: owner, ...a }) => ({ ...a, icon: `filetypes/${fileTypeIcon(a.ext, owner)}.ico` })),
    tools: toolsOf(id, toolNames),
    registry: app
      ? {
          app,
          extensions: APP_EXTENSIONS[id] || [],
          old: OLD_EXTENSIONS[id] || [],
          writable: WRITABLE[id] || [],
          newDocuments: NEW_DOCUMENTS.filter((n) => n.app === id),
        }
      : null,
  };
}

/** The keys packs must be signed with, from the file the build writes into the core (build/out/pack-keys.json). */
export function readPackKeys(file) {
  try {
    const list = JSON.parse(fs.readFileSync(file, 'utf8'))?.keys;
    return (Array.isArray(list) ? list : []).map((pem) => publicKeyOf(String(pem)));
  } catch {
    return [];
  }
}

/** Each pack's size as the build measured it, from the same file: what Add or remove apps says of a pack not here yet. */
export function readPackSizes(file) {
  try {
    const sizes = JSON.parse(fs.readFileSync(file, 'utf8'))?.sizes;
    return sizes && typeof sizes === 'object' ? sizes : {};
  } catch {
    return {};
  }
}

/** A refusal as a window and the log say it. */
const said = ({ id, reason }) => ({ id, name: PACKS.find((p) => p.id === id)?.name || id, reason });

/**
 * The packs this run loads.
 *
 * @param {object} o
 * @param {string|null} o.dir     the packs folder, or null: every app built in
 * @param {*[]} [o.keys]          the public keys packs must be signed with
 * @param {string} o.version      the core's version
 * @param {string} [o.tilesDir]   where the built-in apps' tiles are (resources/apps)
 * @returns {Plan}
 */
export function planPacks({ dir = null, keys = [], version, tilesDir = null }) {
  if (!dir) {
    const packs = PACKS.map(({ id }) => {
      const d = describePack(id);
      const tile = tilesDir ? path.join(tilesDir, `${id}.ico`) : null;
      return { ...d, main: null, windows: null, archive: null, source: 'built-in', tiles: d.tiles.map((t) => ({ ...t, icon: tile })) };
    });
    return plan({ source: 'built-in', dir: null, packs, refused: [] });
  }

  // What Add or remove apps wrote down to go, gone before anything is read.
  let held = [];
  try {
    held = applyRemovals(dir).held;
  } catch {
    held = pendingRemovals(dir);
  }
  if (!keys.length) {
    const ids = fs.existsSync(dir) ? fs.readdirSync(dir).filter((f) => f.endsWith('.pack.json')).map((f) => f.slice(0, -'.pack.json'.length)) : [];
    return plan({ source: 'files', dir, packs: [], refused: ids.map((id) => said({ id, reason: 'This Rutba Office holds no key to check its app packs with, so none was loaded. Install Rutba Office again.' })) });
  }
  const found = findPacks({ dir, publicKey: keys, version });
  const refused = found.refused.map(said);
  const packs = [];
  const core = new Set(Object.keys(coreMethods()));
  const taken = new Map();
  for (const p of found.packs) {
    if (held.includes(p.id)) continue;
    const clash = Object.keys(p.ipc).find((ns) => core.has(ns) || (NAMESPACE_PACK[ns] && NAMESPACE_PACK[ns] !== p.id) || taken.has(ns));
    if (clash) {
      const whose = core.has(clash) ? 'the core' : `the ${PACKS.find((x) => x.id === (taken.get(clash) || NAMESPACE_PACK[clash]))?.name || taken.get(clash) || NAMESPACE_PACK[clash]} pack`;
      refused.push(said({ id: p.id, reason: `The ${p.name} pack was not loaded: it declares the ${clash} namespace, which is ${whose}'s.` }));
      continue;
    }
    // A pack refused for what it needs is refused in turn by findPacks; one left out here is not, yet, needed by anything.
    for (const ns of Object.keys(p.ipc)) taken.set(ns, p.id);
    const windows = p.windows && p.archive ? path.relative(p.archive, p.windows).split(path.sep).join('/') : null;
    packs.push({ ...p, source: 'files', windows, tiles: p.tiles });
  }
  return plan({ source: 'files', dir, packs, refused });
}

/**
 * @typedef {object} Plan
 * @property {'built-in'|'files'} source
 * @property {string|null} dir
 * @property {object[]} packs      the packs loaded, in the order to register them
 * @property {object[]} refused    { id, name, reason }
 * @property {() => string[]} ids
 * @property {(id) => boolean} has
 * @property {(id, reason) => void} drop   a pack that failed to register, refused after all
 * @property {() => object} offer   what the window's preload is told: methods, events, packs, refused
 */
function plan(p) {
  p.ids = () => p.packs.map((x) => x.id);
  p.has = (id) => p.packs.some((x) => x.id === id);
  p.drop = (id, reason) => {
    const at = p.packs.findIndex((x) => x.id === id);
    if (at < 0) return;
    p.packs.splice(at, 1);
    p.refused.push(said({ id, reason }));
  };
  p.methods = () => {
    const out = coreMethods();
    for (const x of p.packs) for (const [ns, names] of Object.entries(x.ipc || {})) out[ns] = [...names];
    return out;
  };
  p.events = () => {
    const out = new Set(coreEvents());
    for (const x of p.packs) for (const e of x.events || []) out.add(e);
    return [...out];
  };
  p.offer = () => ({
    methods: p.methods(),
    events: p.events(),
    // A pack that carries its own window module is served from its archive (rutba://app/packs/<id>/...).
    packs: p.packs.map((x) => ({ id: x.id, name: x.name, windows: x.windows ? `packs/${x.id}/${x.windows}` : null, code: x.code || null })),
    refused: p.refused.map(({ id, name, reason }) => ({ id, name, reason })),
  });
  return p;
}

/**
 * How one pack reaches another's service without importing it: `has('mail')`,
 * and `service('mail')`, which answers undefined while Mail is absent or not
 * yet registered, so every caller copes with it not being there.
 */
export function packView({ plan: p, services }) {
  return {
    has: (id) => p.has(id),
    service: (ns) => (p.has(NAMESPACE_PACK[ns] || '') ? services[ns] : undefined),
  };
}

/**
 * Registers every pack of the plan, in its order, and answers the namespaces
 * they serve. A pack serves only what its signed manifest declared; one that
 * throws, or serves more, is dropped from the plan with the reason.
 *
 * @param {Plan} p
 * @param {object} context   handed to each pack's register: the core's stores, services and hooks
 * @param {{ services?: object, record?: (kind, err) => void }} [o] services: the namespaces so far, filled in as each pack registers (what packView reads)
 */
export async function registerPacks(p, context, { services: served = {}, record = () => {} } = {}) {
  for (const pack of [...p.packs]) {
    const builtIn = BUILT_IN[pack.id] ? async (core) => (await BUILT_IN[pack.id]()).register(core) : null;
    try {
      const mod = pack.main ? await import(pathToFileURL(pack.main).href) : await BUILT_IN[pack.id]?.();
      // Whose code runs: the pack's own, or the core's (a pack whose main hands over to builtIn says so by calling it).
      pack.code = pack.main ? 'own' : 'core';
      if (typeof mod?.register !== 'function') throw new Error('its main module has no register function');
      const handOver = builtIn ? (core) => { pack.code = 'core'; return builtIn(core); } : null;
      const out = (await mod.register({ ...context, pack: { id: pack.id, name: pack.name }, builtIn: handOver })) || {};
      const extra = Object.keys(out).filter((ns) => !pack.ipc?.[ns]);
      if (extra.length) throw new Error(`it serves ${extra.join(', ')}, which its manifest does not declare`);
      Object.assign(served, out);
    } catch (err) {
      record(`pack ${pack.id}`, err);
      p.drop(pack.id, `The ${pack.name} pack could not be started: ${err?.message || err}. Install it again with Add or remove apps.`);
      for (const ns of Object.keys(pack.ipc || {})) delete served[ns];
    }
  }
  return served;
}

/**
 * One pack's services alone, when it is in the plan: for a program that
 * needs one app's code and not the rest (the AI tools run as Node want text
 * recognition). Null when the pack is not installed or would not start.
 * @param {Plan} p
 * @param {string} id
 * @param {object} [context]   what its register is handed
 */
export async function packService(p, id, context = {}) {
  const pack = p.packs.find((x) => x.id === id);
  if (!pack) return null;
  try {
    const mod = pack.main ? await import(pathToFileURL(pack.main).href) : await BUILT_IN[id]?.();
    const builtIn = BUILT_IN[id] ? async (core) => (await BUILT_IN[id]()).register(core) : null;
    return (await mod?.register?.({ ...context, pack: { id: pack.id, name: pack.name }, builtIn })) || null;
  } catch {
    return null;
  }
}

/**
 * One pack's main module, when it is in the plan: its own, or the core's
 * register module for a pack whose code is the core's. Null when the pack is
 * not installed or its module would not load. For a program that runs one
 * app's code (the AI tools run as Node run AI Tools' program).
 * @param {Plan} p
 * @param {string} id
 */
export async function packModule(p, id) {
  const pack = p.packs.find((x) => x.id === id);
  if (!pack) return null;
  try {
    const own = pack.main ? await import(pathToFileURL(pack.main).href) : null;
    // A main module that only registers hands back to the core's (a pack whose code is the core's).
    if (own && Object.keys(own).some((k) => k !== 'register')) return own;
    return (await BUILT_IN[id]?.()) || own;
  } catch {
    return null;
  }
}

/** The pack ids in the launcher's order, the apps' first. */
export const PACK_ORDER = PACKS.map((p) => p.id);
