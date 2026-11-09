// The Locker in the running app: the window's calls, the idle timer, and
// watching the folders and the temporary copies while it is unlocked. This
// module is loaded only by a build made with RUTBA_LOCKER=1 (build-extras.js).

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { createLocker } from './locker.js';
import { helloProvider, fakeHello } from './hello.js';
import { LockerError } from './format.js';
import { OPEN_DIR } from './rules.js';

const METHODS = {
  locker: [
    'state',       // () -> { exists, open, ways, hello, idleMinutes, lockAt, folders: [{ id, name, root, choices, summary }] }
    'setup',       // ({ password }) -> { recovery }
    'unlock',      // ({ password?, recovery?, hello? }) -> state
    'lock',        // () -> state
    'addFolder',   // ({ root, choices, password?, recovery? }) -> { id, root, joined }
    'makeHomeFolder', // () -> { id, root, made }
    'setAdvanced', // ({ on }) -> state
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
    'putBack',     // ({ id, rel }) -> { target, size }
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

export async function register({ stores, broadcast, electron, windows = () => [] }) {
  const { powerMonitor } = await import('electron');
  const dir = path.join(stores.dir, 'locker');
  const locker = createLocker({ dir });
  // A check run never puts a Windows Hello prompt on the desktop.
  const useFake = !electron.isPackaged && (process.env.RUTBA_LOCKER_FAKE_HELLO === '1' || process.env.RUTBA_OFFICE_VERIFY_APPS === '1');
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
        folders.push({ id: f.id, name: path.basename(f.root), root: f.root, choices: { hidden: f.hidden, afterLock: f.afterLock, keepDir: f.keepDir, keepVersions: f.keepVersions, docs: f.docs, openIn: f.openIn || 'beside' }, summary });
      }
    }
    return {
      exists: await locker.exists(),
      open: locker.isOpen(),
      ways: await locker.ways().catch(() => []),
      hello: await hello.available().catch(() => false),
      helloOn: (await locker.ways().catch(() => [])).includes('hello'),
      idleMinutes: idleMs() / 60000,
      advanced: stores.settings.get('locker.advanced', false) === true,
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

  /** Something moved in a working folder: the copy written back once the program has finished. A file still being written is tried again. */
  function noteOpen(full, tries = 0) {
    const dest = locker.openCopyFor(full);
    if (!dest) return;
    later(`open:${dest}`, () => locker.saveOpen(dest).then((r) => {
      if (r.changed) broadcast('locker:changed', { saved: true });
    }).catch((err) => {
      if (err?.code === 'ENOENT') return;
      if ((err?.code === 'EBUSY' || err?.code === 'EPERM' || err?.code === 'EACCES') && tries < 20) noteOpen(full, tries + 1);
      else broadcast('locker:changed', { problem: err?.message || String(err) });
    }));
  }

  function startWatching() {
    stopWatching();
    if (!locker.isOpen()) return;
    locker.folders().then((list) => {
      for (const f of list) {
        try {
          const w = fs.watch(f.root, { recursive: true }, (_e, name) => {
            if (name && name.split(/[\\/]/).includes(OPEN_DIR)) noteOpen(path.join(f.root, name));
            else later(f.id, () => sweep(f.id));
          });
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
        if (name) noteOpen(path.join(dir, 'open', name));
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
    const { left } = await locker.lock();
    stopLockTimer();
    stopWatching();
    await publish();
    broadcast('locker:changed', { locked: reason, left });
  }

  // Locks with the computer: on sleep and when Windows locks.
  powerMonitor.on('suspend', () => { doLock('sleep').catch(() => {}); });
  powerMonitor.on('lock-screen', () => { doLock('screen').catch(() => {}); });

  // Copies whose program has closed are cleared as soon as it has. While the
  // locker is locked this also clears what a program was still holding.
  const shownInWindow = (file) => windows().some((w) => {
    try {
      return !w.isDestroyed() && decodeURIComponent(w.webContents.getURL()).includes(file);
    } catch {
      return false;
    }
  });
  let reaping = false;
  const reaper = setInterval(() => {
    if (reaping || !locker.hasOpenCopies()) return;
    reaping = true;
    locker.reapClosed({ shown: shownInWindow })
      .then((closed) => { if (closed.length) broadcast('locker:changed', { closed }); })
      .catch(() => {})
      .finally(() => { reaping = false; });
  }, 3000);
  reaper.unref?.();

  // Quitting is a lock too: changes written back and copies removed before Office goes.
  let quitting = false;
  electron.on?.('before-quit', (e) => {
    if (quitting || !locker.hasOpenCopies()) return;
    quitting = true;
    e.preventDefault();
    Promise.race([doLock('quit'), new Promise((r) => { setTimeout(r, 8000); })]).catch(() => {}).finally(() => electron.quit());
  });

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
      if (useHello) {
        let secret;
        try {
          secret = await hello.secret(await locker.helloId());
        } catch (err) {
          // Pressing Cancel is an answer, not a fault: the window stays as it was.
          if (err?.code === 'hello-cancelled') return { ...(await stateOut()), cancelled: true };
          throw err;
        }
        try {
          await locker.unlock({ secret });
        } finally {
          secret.fill(0);
        }
      } else await locker.unlock({ password, recovery });
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
    makeHomeFolder: async () => {
      touch();
      const out = await locker.makeHomeFolder(electron.getPath('home'));
      startWatching();
      await publish();
      return out;
    },
    setAdvanced: async ({ on }) => {
      stores.settings.set('locker.advanced', on === true);
      await publish();
      return stateOut();
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
    putBack: async ({ id, rel }) => {
      touch();
      const out = await locker.putBack(id, rel);
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
      const before = await locker.helloId();
      const id = crypto.randomBytes(8).toString('hex');
      let secret;
      try {
        secret = await hello.enroll(id);
      } catch (err) {
        // Cancelling between making the key and signing with it leaves a key nothing uses.
        await hello.forget(id);
        if (err?.code === 'hello-cancelled') return;
        throw err;
      }
      try {
        await locker.enableHello(secret, id);
      } finally {
        secret.fill(0);
      }
      if (before) await hello.forget(before);
      await publish();
    },
    disableHello: async () => {
      const id = await locker.helloId();
      await locker.disableHello();
      if (id) await hello.forget(id);
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
