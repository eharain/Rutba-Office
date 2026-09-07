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

function unwrap(error, ns, name) {
  const message = String(error?.message ?? error ?? 'something went wrong');
  const clean = message.replace(CHANNEL_NOISE, '').trim();
  const out = new Error(clean || message);
  out.name = error?.name && error.name !== 'Error' ? error.name : 'OfficeError';
  out.operation = `${ns}.${name}`;
  out.cause = error;
  return out;
}

const namespaces = {};
for (const [ns, names] of Object.entries(METHODS)) {
  namespaces[ns] = {};
  for (const name of names) {
    namespaces[ns][name] = async (payload) => {
      try {
        return await bridge()[ns][name](payload ?? {});
      } catch (error) {
        throw unwrap(error, ns, name);
      }
    };
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
export const update = namespaces.update;
export const print = namespaces.print;
export const clipboard = namespaces.clipboard;

export function on(event, handler) {
  if (!EVENTS.includes(event)) throw new Error(`unknown event: ${event}`);
  return bridge().on(event, handler);
}

export function boot() {
  return hasShell() ? bridge().boot() : { app: 'home', file: null, platform: 'web', arch: '', versions: {} };
}

export default { app, win, fs, dialog, shell, secrets, store, doc, mail, update, print, clipboard, on, boot, hasShell };
