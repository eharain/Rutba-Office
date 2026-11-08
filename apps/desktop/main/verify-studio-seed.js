// The studio's checks start from a setup of their own: one written to the store
// with the sources and settings the check wants, the one the window opens on.

import { STORE, normaliseIndex, normaliseSetup, withSetup, markUsed, setupSettings, globalSettings } from '@rutba/studio/setup';
import { materialiseScenes, normaliseStack } from '@rutba/studio/stack';

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

/** Between the panels, from a check. `h` has js (a code string run in the window) and until. */
export const PANEL = (p) => `document.querySelector('.st-panels [data-panel="${p}"]')?.click(), 1`;
export const IN_PANEL = (p) => `Boolean(document.querySelector('.st-panel[data-panel="${p}"]'))`;
