// Renderer-side facade.
//
// Apps import this, never `window.rutbaOffice` directly. Two reasons: the shape
// is checked once here instead of at every call site, and when a second backend
// arrives (Tauri) only this file learns about it.

import { METHODS, EVENTS, PACKS, NAMESPACE_PACK } from './contract.js';

function missing() {
  throw new Error(
    'Rutba Office shell is not attached. The renderer must run inside the desktop shell.'
  );
}

function bridge() {
  const b = globalThis.rutbaOffice;
  if (!b) missing();
  return b;
}

/** True when running inside a shell backend rather than a bare browser tab. */
export const hasShell = () => Boolean(globalThis.rutbaOffice);

/**
 * Errors come back wearing Electron's clothes.
 *
 * A rejected `ipcRenderer.invoke` arrives as "Error invoking remote method
 * 'rutba-office/doc:apply': Error: unknown format: bold" — which puts the
 * transport, the channel name and the word "remote" in front of the only part
 * a person can act on, and makes an in-process call between the window and the
 * application's own backend read like a network failure. Nothing here is
 * remote; the wording is stripped so the message says what went wrong.
 */
const CHANNEL_NOISE = /^Error invoking remote method '[^']*':\s*(?:Error:\s*)?/;

/**
 * What the window does to a message before a person reads it: the windows put
 * an engine's English into their own language (renderer/language.js sets it).
 */
let inLanguage = (message) => message;
export function setErrorWords(fn) {
  inLanguage = typeof fn === 'function' ? fn : (message) => message;
}

function unwrap(error, ns, name) {
  const message = String(error?.message ?? error ?? 'something went wrong');
  const clean = message.replace(CHANNEL_NOISE, '').trim();
  const out = new Error(inLanguage(clean || message) || clean || message);
  // The engine's own words, for code that decides by them rather than shows them.
  out.original = clean || message;
  out.name = error?.name && error.name !== 'Error' ? error.name : 'OfficeError';
  out.operation = `${ns}.${name}`;
  out.cause = error;
  return out;
}

/**
 * What this window's bridge offers: the methods and events of this run and
 * the packs it loaded (preload.js asks for them as the window starts). Outside
 * the desktop shell, a test or a bare page, there is no bridge and the
 * whole contract stands in.
 */
const offered = (() => {
  const c = globalThis.rutbaOffice?.__contract;
  return c && c.methods && Array.isArray(c.events) ? c : null;
})();

/** The pack a namespace belongs to, named as a person reads it. */
const packName = (id) => PACKS.find((p) => p.id === id)?.name || id;

/**
 * A call into a pack that is not installed: its methods are not on the
 * bridge, so the call is answered here, in plain words, and never reaches
 * the main process (which has no channel for it either).
 */
function absent(ns, name) {
  const id = NAMESPACE_PACK[ns] || ns;
  const words = `${packName(id)} is not installed on this computer. Add it with Add or remove apps.`;
  const out = new Error(inLanguage(words) || words);
  out.original = words;
  out.name = 'OfficeError';
  out.code = 'absent';
  out.pack = id;
  out.operation = `${ns}.${name}`;
  return out;
}

const namespaces = {};
const offeredMethods = offered ? offered.methods : METHODS;
// The contract's namespaces and any a loaded pack's manifest adds to them.
for (const ns of new Set([...Object.keys(METHODS), ...Object.keys(offeredMethods)])) {
  namespaces[ns] = {};
  const here = new Set(offeredMethods[ns] || []);
  for (const name of new Set([...(METHODS[ns] || []), ...here])) {
    namespaces[ns][name] = here.has(name)
      ? async (payload) => {
          try {
            return await bridge()[ns][name](payload ?? {});
          } catch (error) {
            throw unwrap(error, ns, name);
          }
        }
      : async () => {
          throw absent(ns, name);
        };
  }
}

/** The packs this run loaded, by id; every pack outside the desktop shell. */
export function installedPacks() {
  return offered ? (offered.packs || []).map((p) => p.id) : PACKS.map((p) => p.id);
}

/** Whether a pack is installed and loaded in this run. */
export function isInstalled(id) {
  return installedPacks().includes(id);
}

/** The packs on this computer that did not load, and why: [{ id, name, reason }]. */
export function refusedPacks() {
  return offered ? [...(offered.refused || [])] : [];
}

/** A loaded pack's window module, when it carries one of its own (a URL), or null for an app built into the core. */
export function packWindows(id) {
  return offered ? (offered.packs || []).find((p) => p.id === id)?.windows || null : null;
}

export const app = namespaces.app;
export const win = namespaces.win;
export const fs = namespaces.fs;
export const dialog = namespaces.dialog;
export const shell = namespaces.shell;
export const store = namespaces.store;
export const doc = namespaces.doc;
export const mail = namespaces.mail;
export const update = namespaces.update;
export const present = namespaces.present;
export const oauth = namespaces.oauth;
export const announce = namespaces.announce;
export const defaults = namespaces.defaults;
export const discover = namespaces.discover;
export const print = namespaces.print;
export const clipboard = namespaces.clipboard;
export const calendar = namespaces.calendar;
export const contacts = namespaces.contacts;
export const dav = namespaces.dav;
export const thumbs = namespaces.thumbs;
export const capture = namespaces.capture;
export const pdf = namespaces.pdf;
export const ocr = namespaces.ocr;
export const ai = namespaces.ai;
export const live = namespaces.live;
export const talk = namespaces.talk;
export const base = namespaces.base;
export const profiles = namespaces.profiles;
export const packs = namespaces.packs;

export function on(event, handler) {
  if (!EVENTS.includes(event) && !offered?.events.includes(event)) throw new Error(`unknown event: ${event}`);
  // An event of a pack that is not installed never comes: listening to it is listening to nothing.
  if (offered && !offered.events.includes(event)) return () => {};
  return bridge().on(event, handler);
}

/** A dropped File's path on disk, or null for one that is not on disk. */
export function pathOf(file) {
  return hasShell() ? bridge().pathOf?.(file) ?? null : null;
}

export function boot() {
  return hasShell() ? bridge().boot() : { app: 'home', file: null, platform: 'web', arch: '', versions: {} };
}

export default { ...namespaces, app, win, fs, dialog, shell, store, doc, mail, oauth, present, update, announce, defaults, discover, print, clipboard, calendar, contacts, dav, thumbs, capture, pdf, ocr, ai, live, talk, base, profiles, packs, on, boot, pathOf, hasShell, installedPacks, isInstalled, refusedPacks, packWindows };
