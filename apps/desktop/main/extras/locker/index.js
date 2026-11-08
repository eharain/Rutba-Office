// The Locker in the running app: the window's calls, the idle timer, and
// watching the folders and the temporary copies while it is unlocked. This
// module is loaded only by a build made with RUTBA_LOCKER=1 (build-extras.js).

import fs from 'node:fs';
import path from 'node:path';
import { createLocker } from './locker.js';
import { helloProvider, fakeHello } from './hello.js';
import { LockerError } from './format.js';

const METHODS = {
  locker: [
    'state',       // () -> { exists, open, ways, hello, idleMinutes, lockAt, folders: [{ id, name, root, choices, summary }] }
    'setup',       // ({ password }) -> { recovery }
    'unlock',      // ({ password?, recovery?, hello? }) -> state
    'lock',        // () -> state
    'addFolder',   // ({ root, choices, password?, recovery? }) -> { id, root, joined }
    'setChoices',  // ({ id, patch }) -> folder
    'removeFolder',// ({ id }) -> { removed, failed }
    'list',        // ({ id }) -> entries
    'lockWaiting', // ({ id }) -> { locked, failed }
    'why',         // ({ id, rel }) -> { state, why }
    'open',        // ({ id, rel }) -> { path, app }
    'exportCopy',  // ({ id, rel, dest, overwrite? }) -> { size }
    'exportShare', // ({ id, rel, dest, password }) -> { size }
    'openShared',  // ({ file, dest, password }) -> { size }
    'importFile',  // ({ id, source, rel? }) -> { size }
    'unlockForGood', // ({ id, rel }) -> { dest }
    'versions',    // ({ id, rel }) -> [{ version, current, label }]
    'restore',     // ({ id, rel, version }) -> { size }
    'readRules',   // ({ id }) -> text
    'writeRules',  // ({ id, text }) -> void
    'changePassword',  // ({ current, next }) -> { failed }
    'replaceRecovery', // ({ current }) -> { recovery }
    'enableHello',  // () -> void
    'disableHello', // () -> void
    'setIdle',      // ({ minutes }) -> state
    'touch',        // () -> void   the window was used
  ],
};
const EVENTS = ['locker:state', 'locker:changed'];

export function describe() {
  return { methods: METHODS, events: EVENTS, pack: { id: 'locker', name: 'Locker', windows: null, code: 'core' } };
}

/** The Office app that opens a file, or the system for anything else. */
async function appForFile(file) {
  const { appFor, kindFromExtension } = await import('@rutba/office-formats');
  return appFor(kindFromExtension(file)) || null;
}

export async function register({ stores, broadcast, electron }) {
  const dir = path.join(stores.dir, 'locker');
  const locker = createLocker({ dir });
  const useFake = !electron.isPackaged && process.env.RUTBA_LOCKER_FAKE_HELLO === '1';
  const hello = useFake ? fakeHello() : helloProvider();

  let lockTimer = null;
  let lockAt = null;
  const watchers = new Map();
  let openWatcher = null;
  const idleMs = () => Number(stores.settings.get('locker.idleMinutes', 10)) * 60000;

  const stopLockTimer = () => {
    if (lockTimer) clearTimeout(lockTimer);
    lockTimer = null;
    lockAt = null;
  };

  const startLockTimer = () => {
    stopLockTimer();
    if (!locker.isOpen()) return;
    const ms = idleMs();
    lockAt = Date.now() + ms;
    lockTimer = setTimeout(() => { doLock('idle').catch(() => {}); }, ms);
  };

  async function stateOut() {
    const folders = [];
    if (locker.isOpen()) {
      for (const f of await locker.folders()) {
        let summary = null;
        try {
          summary = await locker.summary(f.id);
        } catch {
          // The folder's drive is away: it shows as such.
        }
        folders.push({ id: f.id, name: path.basename(f.root), root: f.root, choices: { hidden: f.hidden, afterLock: f.afterLock, keepDir: f.keepDir, keepVersions: f.keepVersions, docs: f.docs }, summary });
      }
    }
    return {
      exists: await locker.exists(),
      open: locker.isOpen(),
      ways: await locker.ways().catch(() => []),
      hello: await hello.available().catch(() => false),
      helloOn: (await locker.ways().catch(() => [])).includes('hello'),
      idleMinutes: idleMs() / 60000,
      lockAt,
      folders,
    };
  }

  async function publish() {
    try {
      broadcast('locker:state', await stateOut());
    } catch {
      // A window on its way out.
    }
  }

  /** Files that arrived are locked; files changed in a temporary copy are written back. */
  async function sweep(id) {
    if (!locker.isOpen()) return;
    try {
      const out = await locker.lockWaiting(id);
      if (out.locked.length || out.failed.length) broadcast('locker:changed', { id, ...out });
    } catch {
      // A folder that has gone away is shown as it is on the next scan.
    }
  }

  const debounced = new Map();
  const later = (key, fn) => {
    clearTimeout(debounced.get(key));
    debounced.set(key, setTimeout(fn, 1500));
  };

  function startWatching() {
    stopWatching();
    if (!locker.isOpen()) return;
    locker.folders().then((list) => {
      for (const f of list) {
        try {
          const w = fs.watch(f.root, { recursive: true }, () => later(f.id, () => sweep(f.id)));
          w.on('error', () => {});
          watchers.set(f.id, w);
        } catch {
          // A drive that does not support watching: the next start's sweep locks what arrived.
        }
      }
    }).catch(() => {});
    try {
      fs.mkdirSync(path.join(dir, 'open'), { recursive: true });
      openWatcher = fs.watch(path.join(dir, 'open'), { recursive: true }, (_e, name) => {
        if (!name) return;
        const full = path.join(dir, 'open', name);
        if (!locker.isOpenCopy(full)) return;
        later(`open:${full}`, () => locker.saveOpen(full).then((r) => { if (r.changed) broadcast('locker:changed', { saved: true }); }).catch((err) => broadcast('locker:changed', { problem: err?.message || String(err) })));
      });
      openWatcher.on('error', () => {});
    } catch {
      openWatcher = null;
    }
  }

  function stopWatching() {
    for (const w of watchers.values()) w.close();
    watchers.clear();
    if (openWatcher) openWatcher.close();
    openWatcher = null;
    for (const t of debounced.values()) clearTimeout(t);
    debounced.clear();
  }

  async function afterUnlock() {
    startLockTimer();
    startWatching();
    await publish();
    // Whatever arrived while it was locked, or was still waiting.
    for (const f of await locker.folders()) later(f.id, () => sweep(f.id));
  }

  async function doLock(reason) {
    await locker.lock();
    stopLockTimer();
    stopWatching();
    await publish();
    broadcast('locker:changed', { locked: reason });
  }

  // Locks with the computer: on sleep and when Windows locks.
  electron.powerMonitor?.on?.('suspend', () => { doLock('sleep').catch(() => {}); });
  electron.powerMonitor?.on?.('lock-screen', () => { doLock('screen').catch(() => {}); });

  const touch = () => { if (locker.isOpen()) startLockTimer(); };

  async function open({ id, rel }) {
    touch();
    const done = await locker.openPlain(id, rel);
    return { ...done, app: await appForFile(done.path) };
  }

  const service = {
    state: stateOut,
    setup: async ({ password }) => {
      const out = await locker.setup(password);
      startLockTimer();
      startWatching();
      await publish();
      return out;
    },
    unlock: async ({ password, recovery, hello: useHello }) => {
      if (useHello) await locker.unlock({ secret: await hello.secret() });
      else await locker.unlock({ password, recovery });
      await afterUnlock();
      return stateOut();
    },
    lock: async () => {
      await doLock('asked');
      return stateOut();
    },
    addFolder: async ({ root, choices, password = null, recovery = null }) => {
      touch();
      const out = await locker.addFolder(root, choices, { password, recovery });
      startWatching();
      await publish();
      return out;
    },
    setChoices: async ({ id, patch }) => {
      touch();
      const f = await locker.setChoices(id, patch);
      await publish();
      return f;
    },
    removeFolder: async ({ id }) => {
      touch();
      const out = await locker.removeFolder(id);
      startWatching();
      await publish();
      return out;
    },
    list: async ({ id }) => {
      touch();
      return locker.scan(id);
    },
    lockWaiting: async ({ id, ...opts }) => {
      touch();
      const out = await locker.lockWaiting(id, opts);
      await publish();
      return out;
    },
    why: ({ id, rel }) => locker.why(id, rel),
    open,
    exportCopy: async ({ id, rel, dest, overwrite, allowConflict }) => {
      touch();
      return locker.exportCopy(id, rel, dest, { overwrite, allowConflict });
    },
    exportShare: async ({ id, rel, dest, password }) => {
      touch();
      return locker.exportShare(id, rel, dest, password);
    },
    openShared: ({ file, dest, password }) => locker.openShared(file, dest, password),
    importFile: async ({ id, source, rel }) => {
      touch();
      const out = await locker.importFile(id, source, rel);
      await publish();
      return out;
    },
    unlockForGood: async ({ id, rel }) => {
      touch();
      const out = await locker.unlockForGood(id, rel);
      await publish();
      return out;
    },
    versions: ({ id, rel }) => locker.versions(id, rel),
    restore: async ({ id, rel, version }) => {
      touch();
      const out = await locker.restore(id, rel, version);
      await publish();
      return out;
    },
    readRules: ({ id }) => locker.readRules(id),
    writeRules: async ({ id, text }) => {
      await locker.writeRules(id, text);
      await publish();
    },
    changePassword: async ({ current, next }) => {
      if (!locker.isOpen()) throw new LockerError('locked', 'Unlock the locker first.');
      touch();
      return locker.changePassword(current, next);
    },
    replaceRecovery: async ({ current }) => {
      if (!locker.isOpen()) throw new LockerError('locked', 'Unlock the locker first.');
      touch();
      return locker.replaceRecovery(current);
    },
    enableHello: async () => {
      if (!locker.isOpen()) throw new LockerError('locked', 'Unlock the locker first.');
      touch();
      await locker.enableHello(await hello.enroll());
      await publish();
    },
    disableHello: async () => {
      await locker.disableHello();
      await publish();
    },
    setIdle: async ({ minutes }) => {
      const m = Math.max(1, Math.min(240, Math.floor(Number(minutes) || 10)));
      stores.settings.set('locker.idleMinutes', m);
      startLockTimer();
      return stateOut();
    },
    touch: async () => {
      touch();
    },
  };

  return { locker: service };
}
