// Renderer-side facade.
//
// Apps import this, never `window.rutbaOffice` directly. Two reasons: the shape
// is checked once here instead of at every call site, and when a second backend
// arrives (Tauri) only this file learns about it.

import { METHODS, EVENTS } from './contract.js';

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

const namespaces = {};
for (const [ns, names] of Object.entries(METHODS)) {
  namespaces[ns] = {};
  for (const name of names) {
    namespaces[ns][name] = (payload) => bridge()[ns][name](payload ?? {});
  }
}

export const app = namespaces.app;
export const win = namespaces.win;
export const fs = namespaces.fs;
export const dialog = namespaces.dialog;
export const shell = namespaces.shell;
export const secrets = namespaces.secrets;
export const store = namespaces.store;
export const doc = namespaces.doc;
export const mail = namespaces.mail;
export const print = namespaces.print;
export const clipboard = namespaces.clipboard;

export function on(event, handler) {
  if (!EVENTS.includes(event)) throw new Error(`unknown event: ${event}`);
  return bridge().on(event, handler);
}

export function boot() {
  return hasShell() ? bridge().boot() : { app: 'home', file: null, platform: 'web', arch: '', versions: {} };
}

export default { app, win, fs, dialog, shell, secrets, store, doc, mail, print, clipboard, on, boot, hasShell };
