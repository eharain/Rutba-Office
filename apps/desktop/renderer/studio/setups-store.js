// Where the setups are kept: the list under one key, each setup under its own,
// so a save rewrites one setup and the Recorder and Live Broadcast open at once
// do not overwrite each other's. The shared settings are a key of their own.

import { STORE, normaliseIndex, normaliseSetup, withSetup, withoutSetup, markUsed, SETUP_EXT, setupId, uniqueName, readSetupFile, placeMedia, writeSetupFile, rematch, menuOrder } from '@rutba/studio/setup';
import { t } from '@rutba/office-ui';

export function createSetupStore(shell) {
  const get = (key, fallback = null) => shell.store.get({ key, fallback }).catch(() => fallback);
  const set = (key, value) => shell.store.set({ key, value }).catch(() => {});
  let chain = Promise.resolve();

  // One after another within this window: the list is read as it is now before each change.
  const index = () => get(STORE.index).then(normaliseIndex);
  const changeIndex = (fn) => {
    chain = chain.then(async () => {
      const next = fn(await index());
      await set(STORE.index, next);
      return next;
    });
    return chain;
  };

  return {
    index,
    changeIndex,
    get,
    set,
    /** One setup as stored, or null. */
    load: async (id) => {
      const raw = await get(STORE.setup(id));
      return raw ? normaliseSetup(raw) : null;
    },
    save: (setup) => set(STORE.setup(setup.id), setup),
    /** A new setup written, and put in the list. */
    add: async (setup) => {
      await set(STORE.setup(setup.id), setup);
      return changeIndex((i) => withSetup(i, setup));
    },
    rename: async (setup) => {
      await set(STORE.setup(setup.id), setup);
      return changeIndex((i) => withSetup(i, setup));
    },
    remove: async (id) => {
      await shell.store.delete({ key: STORE.setup(id) }).catch(() => {});
      return changeIndex((i) => withoutSetup(i, id));
    },
    used: (id, app) => changeIndex((i) => markUsed(i, id, app, Date.now())),
    global: () => get(STORE.global),
    setGlobal: (value) => set(STORE.global, value),
  };
}

/** Where a setup's pictures and music are kept once opened from a file. */
const mediaDir = (paths, id) => `${paths.userData}${paths.separator}studio-media${paths.separator}${id}`;

/**
 * A `.rstudio` file read and its pictures and music put where this profile
 * keeps them; the sources found again on this computer. Answers
 * { setup, found, missing, mediaCount, streams }, ready to be added.
 */
export async function openSetupFile({ shell, path, store, sources, devices, app }) {
  const { bytes } = await shell.fs.read({ path });
  const { setup, media } = await readSetupFile(bytes);
  const list = await store.index();
  const id = setupId(list.setups.map((x) => x.id));
  const paths = await shell.app.paths();
  const dir = mediaDir(paths, id);
  const placed = {};
  for (const m of media) {
    const to = `${dir}${paths.separator}${m.name}`;
    await shell.fs.write({ path: to, bytes: m.bytes });
    placed[m.entry] = to;
  }
  const taken = menuOrder(list).map((x) => x.name);
  const named = { ...placeMedia(setup, placed), id, name: uniqueName(setup.name || t('studio.new-setup'), taken), app, usedAt: Date.now(), tools: false, streams: [], liveOn: setup.liveOn };
  const matched = rematch(named, sources, devices);
  return { setup: matched.setup, found: matched.found, missing: matched.missing, mediaCount: media.length, streams: setup.streams, savedApp: setup.app };
}

/** A setup written to a `.rstudio` file, with its media. Answers the files that could not be included. */
export async function saveSetupFile({ shell, path, setup, streams }) {
  const read = async (p) => (await shell.fs.read({ path: p })).bytes;
  const { bytes, absent } = await writeSetupFile(setup, { read, streams });
  await shell.fs.write({ path, bytes });
  return absent;
}

export const SETUP_FILTERS = () => [{ name: t('studio.setups-store.rutba-studio-setups'), extensions: [SETUP_EXT] }];
