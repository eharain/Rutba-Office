// The studio's checks start from a setup of their own: one written to the store
// with the sources and settings the check wants, the one the window opens on.

import { STORE, normaliseIndex, normaliseSetup, withSetup, markUsed, setupSettings, globalSettings } from '@rutba/studio/setup';
import { materialiseScenes, normaliseStack } from '@rutba/studio/stack';
import { BrowserWindow } from 'electron';

export const CHECK_SETUP = 'setup-check';

/** The three things kept in the store for a setup called `name`, opened in `app`. */
export function checkSetup({ app, stack, studio, name = 'Check' }) {
  const setup = normaliseSetup({ id: CHECK_SETUP, name, app, usedAt: 1, stack: materialiseScenes(normaliseStack(stack)), settings: setupSettings(studio) }, { app });
  return { setup, index: markUsed(withSetup(normaliseIndex({}), setup), CHECK_SETUP, app, 1), global: globalSettings(studio) };
}

const KEYS = [STORE.index, STORE.setup(CHECK_SETUP), STORE.global];

/** The setup written through a window's store; answers what was there, for restoreSetup. */
export async function seedSetup(js, o) {
  const { setup, index, global } = checkSetup(o);
  const kept = await js(`Promise.all(${JSON.stringify(KEYS)}.map((key) => window.rutbaOffice.store.get({ key, fallback: null })))`);
  await js(`Promise.all([
    window.rutbaOffice.store.set({ key: ${JSON.stringify(STORE.index)}, value: ${JSON.stringify(index)} }),
    window.rutbaOffice.store.set({ key: ${JSON.stringify(STORE.setup(CHECK_SETUP))}, value: ${JSON.stringify(setup)} }),
    window.rutbaOffice.store.set({ key: ${JSON.stringify(STORE.global)}, value: ${JSON.stringify(global)} }),
  ]).then(() => 1)`);
  return kept;
}

export const restoreSetup = (js, kept) => js(`Promise.all(${JSON.stringify(KEYS)}.map((key, i) => window.rutbaOffice.store.set({ key, value: ${JSON.stringify(kept)}[i] }))).then(() => 1)`);

/** The setup as it is kept now. */
export const readSetup = (js) => js(`window.rutbaOffice.store.get({ key: ${JSON.stringify(STORE.setup(CHECK_SETUP))}, fallback: null })`);

/** The setup written straight to the settings store, before a window opens. */
export function seedSetupStore(stores, o) {
  const { setup, index, global } = checkSetup(o);
  stores.settings.set(STORE.index, index);
  stores.settings.set(STORE.setup(CHECK_SETUP), setup);
  stores.settings.set(STORE.global, global);
}

export const CHECK_SOUNDS = (mic = true) => [{ id: 'mic-1', kind: 'mic', ref: '', name: '', live: mic, muted: false, gain: 1 }, { id: 'system-1', kind: 'system', ref: '', name: '', live: false, muted: false, gain: 1 }];

/** The windows there are now, for recorderOpened to tell the new one by. */
export const windowsNow = () => new Set(BrowserWindow.getAllWindows().map((w) => w.id));

/** The Recorder window another window has just opened (A7), once it is drawn: { rw, rjs }. */
export async function recorderOpened(until, known, ms = 15000) {
  const fresh = () => BrowserWindow.getAllWindows().find((w) => !known.has(w.id) && !w.isDestroyed() && w.webContents.getURL().includes('app=recorder')) || null;
  await until(() => Boolean(fresh()), 'the Recorder window', ms);
  const rw = fresh();
  const rjs = (code) => rw.webContents.executeJavaScript(code);
  await until(() => rjs(`Boolean(document.querySelector('.st-panels'))`).catch(() => false), 'the Recorder drawn', ms);
  return { rw, rjs };
}

/** Pressed Record, and recording. */
export const RECORDING = `document.querySelector('.st-status')?.dataset.phase === 'recording'`;
export const PRESS_RECORD = `document.querySelector('.st-record')?.click(), 1`;

/** Between the panels, from a check. `h` has js (a code string run in the window) and until. */
export const PANEL = (p) => `document.querySelector('.st-panels [data-panel="${p}"]')?.click(), 1`;
export const IN_PANEL = (p) => `Boolean(document.querySelector('.st-panel[data-panel="${p}"]'))`;
