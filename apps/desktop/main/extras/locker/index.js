// The Locker in the running app: the window's calls, the idle timer, and
// watching the folders and the temporary copies while it is unlocked. This
// module is loaded through build-extras.js.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { createLocker } from './locker.js';
import { helloProvider, fakeHello } from './hello.js';
import { touchIdProvider, noQuickUnlock, quickUnlockName } from './touchid.js';
import { LockerError } from './format.js';
import { OPEN_DIR } from './rules.js';
import { addressShows, addressApp } from './window-address.js';
import { createMemoryOpener } from './memory.js';
import { createMemoryFiles, setMemoryFiles, setLockedCopyCheck, unsavedOf } from '../../memory-files.js';

const METHODS = {
  locker: [
    'state',       // () -> { exists, open, ways, hello, idleMinutes, lockAt, shownCopies, folders: [{ id, name, root, choices, summary }] }
    'setup',       // ({ password }) -> { recovery }
    'unlock',      // ({ password?, recovery?, hello? }) -> state
    'lock',        // () -> state
    'lookAtFolder', // ({ root }) -> { locked, opens }
    'addFolder',   // ({ root, choices, password?, recovery? }) -> { id, root, joined }
    'makeHomeFolder', // () -> { id, root, made }
    'setAdvanced', // ({ on }) -> state
    'setChoices',  // ({ id, patch }) -> folder
    'removeFolder',// ({ id, leaveDamaged }) -> { removed, failed, damaged }
    'list',        // ({ id }) -> entries
    'lockWaiting', // ({ id }) -> { locked, failed }
    'why',         // ({ id, rel }) -> { state, why }
    'open',        // ({ id, rel }) -> { path, app, inMemory, plainWhy? }   plainWhy: 'too-large' | 'by-path' | 'kind'
    'reopen',      // ({ keys }) -> [{ key, path, app, unsaved }]   the files that were open when it locked, or that a crash left
    'dismissReopen', // ({ keys }) -> void
    'exportCopy',  // ({ id, rel, dest, overwrite? }) -> { size }
    'exportShare', // ({ id, rel, dest, password }) -> { size }
    'shareForMail', // ({ id, rel, password }) -> { dest, size }
    'openShared',  // ({ file, dest, password }) -> { size }
    'importShared', // ({ file, id, password }) -> { rel }
    'importFile',  // ({ id, source, rel? }) -> { size }
    'unlockForGood', // ({ id, rel }) -> { dest }
    'renameFile',  // ({ id, rel, name, state? }) -> { rel }
    'deleteFile',  // ({ id, rel, state? }) -> { removed }
    'lockOne',     // ({ id, rel }) -> { target, size }
    'setWhenLocked', // ({ keepCopies?, relock? }) -> state
    'versions',    // ({ id, rel }) -> [{ version, current, label }]
    'restore',     // ({ id, rel, version }) -> { size }
    'readRules',   // ({ id }) -> text
    'writeRules',  // ({ id, text }) -> void
    'changePassword',  // ({ current, next, recovery }) -> { failed }
    'replaceRecovery', // ({ current }) -> { recovery }
    'enableHello',  // () -> void
    'disableHello', // () -> void
    'setIdle',      // ({ minutes }) -> state
    'touch',        // () -> { lockAt }   the window was used
    'safetyEstimate',    // ({ root }) -> { files, bytes, needed, free, enough }
    'checkSafetyCopy',   // ({ id }) -> { files, matched, changed, missing, unreadable }
    'removeSafetyCopy',  // ({ id, accept? }) -> { removed, check }
    'discardSafetyCopy', // ({ copyId }) -> { removed }
    'restoreSafetyCopy', // ({ copyId, dest }) -> { dest, restored, failed }   works with the locker shut
    'leavePlain',        // ({ id, rel }) -> { rel, state }   a file that kept changing during the safety copy, left out
  ],
};
const EVENTS = ['locker:state', 'locker:changed'];

export function describe() {
  return { methods: METHODS, events: EVENTS, pack: { id: 'locker', name: 'Locker', windows: null, code: 'core' } };
}

/**
 * The system's own safe store for a small secret, through Electron's
 * safeStorage on every system: DPAPI on Windows, the Keychain on macOS, the
 * secret service on Linux. Not taken for one where Linux has no keyring and
 * Electron would use a fixed key: then there are no safety copies, and the
 * window says why, rather than a key kept plain.
 */
export function systemSeal(safeStorage) {
  return {
    name: 'system',
    available: async () => {
      try {
        if (process.platform === 'linux' && safeStorage.getSelectedStorageBackend?.() === 'basic_text') return false;
        return safeStorage.isEncryptionAvailable();
      } catch {
        return false;
      }
    },
    seal: async (bytes) => safeStorage.encryptString(Buffer.from(bytes).toString('base64')),
    open: async (bytes) => Buffer.from(safeStorage.decryptString(Buffer.from(bytes)), 'base64'),
  };
}

/** The Office app that opens a file, or the system for anything else. */
async function appForFile(file) {
  const { appFor, kindFromExtension } = await import('@rutba/office-formats');
  return appFor(kindFromExtension(file)) || null;
}

export async function register({ stores, broadcast, electron, windows = () => [] }) {
  const { powerMonitor, safeStorage, systemPreferences } = await import('electron');
  const dir = path.join(stores.dir, 'locker');
  // The safety copies' key: the system's safe store, the same call on every system; and the profile's password over it where it has one.
  const locker = createLocker({ dir, vault: stores.vault || null, protector: systemSeal(safeStorage) });
  await locker.loadOpenRecord().catch(() => {});
  await locker.tidySafety().catch(() => {});
  // Locked files opened in memory (memory.js): the table every app's file calls ask first. Markers a
  // run that did not end cleanly left are swept; nothing behind them was ever on the disk.
  const table = createMemoryFiles({ root: path.join(dir, 'mem') });
  table.sweep();
  setMemoryFiles(table);
  // The plain copies it still makes count as the Locker's too: their windows save only back into them.
  // The list too: the document thread is sent it, as it has no Locker of its own to ask.
  setLockedCopyCheck((p) => locker.isOpenCopy(p), () => locker.openCopyPaths());
  const memory = createMemoryOpener({ locker, table, recoveryDir: path.join(dir, 'recovery'), unsavedOf });
  /** Whether an Office window shows a file from the Locker: a plain copy, or one opened in memory. */
  const anyShown = () => locker.openCopyPaths().some(shownInWindow) || memory.paths().some(shownInWindow);
  // A check run never puts a Windows Hello prompt on the desktop.
  const useFake = !electron.isPackaged && (process.env.RUTBA_LOCKER_FAKE_HELLO === '1' || process.env.RUTBA_OFFICE_VERIFY_APPS === '1');
  // The quick way in (owner, 13 Nov): Windows Hello on Windows, Touch ID on a Mac, both off until turned on;
  // none on Linux, where the password and the recovery key are the ways in.
  const quick = quickUnlockName();
  const hello = useFake ? fakeHello()
    : quick === 'hello' ? helloProvider()
      : quick === 'touchid' ? touchIdProvider({ systemPreferences, safeStorage, dir })
        : noQuickUnlock();

  let lockTimer = null;
  let lockAt = null;
  /** Files being opened just now: the idle lock waits for them. */
  let opening = 0;
  /** The copy handed over last, until a window shows it: { path, at }. */
  let handed = null;
  const HANDED_GRACE_MS = 30000;
  const watchers = new Map();
  let openWatcher = null;
  const idleMs = () => Number(stores.settings.get('locker.idleMinutes', 10)) * 60000;
  const keepCopies = () => stores.settings.get('locker.keepCopies', false) === true;
  const relock = () => stores.settings.get('locker.relock', false) === true;

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
    lockTimer = setTimeout(() => {
      // A document from the Locker open in an Office window, one being opened, or one just
      // handed over whose window is still on its way, is the Locker in use, and locking would
      // take its file away. The copy handed over stops counting once a window has shown it.
      if (handed && shownInWindow(handed.path)) handed = null;
      const onItsWay = handed && Date.now() - handed.at < HANDED_GRACE_MS;
      if (opening > 0 || onItsWay || anyShown()) startLockTimer();
      else doLock('idle').catch(() => {});
    }, ms);
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
        folders.push({ id: f.id, name: path.basename(f.root), root: f.root, choices: { hidden: f.hidden, afterLock: f.afterLock, keepDir: f.keepDir, keepVersions: f.keepVersions, docs: f.docs }, safetyWanted: f.safetyWanted === true, safetyBusy: f.safetyBusy || null, summary });
      }
    }
    return {
      exists: await locker.exists(),
      open: locker.isOpen(),
      ways: await locker.ways().catch(() => []),
      hello: await hello.available().catch(() => false),
      // Which one it is, so the window names it: 'hello' (Windows Hello) or 'touchid' (Touch ID).
      quick,
      helloOn: (await locker.ways().catch(() => [])).includes('hello'),
      idleMinutes: idleMs() / 60000,
      advanced: stores.settings.get('locker.advanced', false) === true,
      keepCopies: keepCopies(),
      relock: relock(),
      lockAt,
      // How many of its open copies an Office window is showing: what the idle lock and closing the window wait for.
      shownCopies: locker.isOpen() ? locker.openCopyPaths().filter(shownInWindow).length + memory.paths().filter(shownInWindow).length : 0,
      // The files that were open when it locked, and what a crash left: offered once it is unlocked (owner, 13 Nov).
      reopen: locker.isOpen() ? await memory.offers().catch(() => []) : [],
      folders,
      // Shown with the locker shut too: the password screen offers a restore from them.
      safety: await locker.safetyCopies().catch(() => []),
      safetyInfo: await locker.safetyInfo().catch(() => ({ available: false, scope: 'account' })),
    };
  }

  async function publish() {
    try {
      broadcast('locker:state', await stateOut());
    } catch {
      // A window on its way out.
    }
  }

  /** Something changed in a folder: a locked file copied in is made the folder's own, and the window's list brought up to date. A plain file is locked only when the person asks. */
  async function changed(id) {
    if (!locker.isOpen()) return;
    await locker.adoptCopies(id).catch(() => {});
    if (locker.isOpen()) broadcast('locker:changed', { id });
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
      // The file went, or the copy was cleared before this late notice arrived.
      if (err?.code === 'ENOENT' || !locker.isOpenCopy(dest)) return;
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
            else later(f.id, () => changed(f.id));
          });
          w.on('error', () => {});
          watchers.set(f.id, w);
        } catch {
          // A drive that does not support watching: the list is read again when the folder is next shown.
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
    for (const f of await locker.folders()) later(f.id, () => changed(f.id));
  }

  /** How long each kind of lock can spend locking files again: a computer going to sleep gives only moments. */
  const LOCK_BUDGET = { sleep: 4000, quit: 6000 };

  async function doLock(reason) {
    closedWithDocs = false;
    // Files opened in memory (owner, 13 Nov): unsaved work sealed as a recovery copy while the keys are still
    // here, each window closed without asking, and every byte let go. The next unlock offers them again.
    const budget = LOCK_BUDGET[reason] ?? 30000;
    const sealing = await memory.sealOpen({ deadline: Date.now() + budget / 2 }).catch(() => []);
    closeWindowsShowing(sealing);
    const { left, open: stillOpen, notLocked } = await locker.lock({ keepCopies: keepCopies(), relock: relock(), budgetMs: budget, shown: shownInWindow });
    await memory.closeAll().catch(() => {});
    stopLockTimer();
    handed = null;
    stopWatching();
    await publish();
    broadcast('locker:changed', { locked: reason, left, open: stillOpen, notLocked });
  }

  // Locks with the computer: on sleep and when Windows locks.
  powerMonitor.on('suspend', () => { doLock('sleep').catch(() => {}); });
  powerMonitor.on('lock-screen', () => { doLock('screen').catch(() => {}); });

  // Copies whose program has closed are cleared as soon as it has. While the
  // locker is locked this also clears what a program was still holding.
  // The file is read from the window's address as a parameter and compared as a
  // path (window-address.js): a space is written "+" there.
  const shownInWindow = (file) => windows().some((w) => {
    try {
      return !w.isDestroyed() && addressShows(w.webContents.getURL(), file);
    } catch {
      return false;
    }
  });

  /** The windows showing these files, closed at once: their work is sealed already, and asking would keep it open. */
  function closeWindowsShowing(files) {
    if (!files.length) return;
    for (const w of windows()) {
      try {
        if (!w.isDestroyed() && files.some((f) => addressShows(w.webContents.getURL(), f))) w.destroy();
      } catch {
        /* closed already */
      }
    }
  }

  // Closing the Locker's window locks it. A document from it still open in an
  // Office window would lose its file, so the lock waits for that to close.
  const isLockerWindow = (w) => {
    try {
      return !w.isDestroyed() && addressApp(w.webContents.getURL()) === 'locker';
    } catch {
      return false;
    }
  };
  let closedWithDocs = false;
  function lockIfClosed() {
    if (!locker.isOpen() || windows().some(isLockerWindow)) {
      closedWithDocs = false;
      return;
    }
    if (anyShown()) {
      closedWithDocs = true;
      return;
    }
    doLock('closed').catch(() => {});
  }
  electron.on?.('browser-window-created', (_e, w) => {
    let mine = false;
    w.webContents.once('did-finish-load', () => { mine = isLockerWindow(w); });
    w.once('closed', () => { if (mine) setTimeout(lockIfClosed, 0); });
  });

  let reaping = false;
  const reaper = setInterval(() => {
    if (closedWithDocs) lockIfClosed();
    if (memory.hasOpen()) memory.reap({ shown: shownInWindow }).then((gone) => { if (gone.length) broadcast('locker:changed', { closed: gone }); }, () => {});
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
    if (quitting || (!locker.hasOpenCopies() && !memory.hasOpen() && !(locker.isOpen() && relock()))) return;
    quitting = true;
    e.preventDefault();
    Promise.race([doLock('quit'), new Promise((r) => { setTimeout(r, 8000); })]).catch(() => {})
      .then(() => memory.closeAll()).catch(() => {}).finally(() => electron.quit());
  });

  const touch = () => { if (locker.isOpen()) startLockTimer(); };

  async function open({ id, rel }) {
    touch();
    opening += 1;
    try {
      // One of Office's own kinds opens in memory: no plain copy is written. Anything else, or a document
      // over the size limit, opens from a plain copy in the profile as before, and the window says so.
      const inMemory = await memory.open(id, rel);
      if (inMemory.path) {
        handed = { path: inMemory.path, at: Date.now() };
        return { path: inMemory.path, inMemory: true, note: inMemory.note, app: await appForFile(inMemory.path) };
      }
      const done = await locker.openPlain(id, rel);
      handed = { path: done.path, at: Date.now() };
      return { ...done, inMemory: false, plainWhy: inMemory.why, app: await appForFile(done.path) };
    } finally {
      opening -= 1;
      // The file handed over is the Locker used: its window gets the whole idle time to open in, not what the opening left.
      touch();
    }
  }

  /** The files offered after an unlock, opened again: each in memory, with its unsaved work when it had some. */
  async function reopen({ keys = [] } = {}) {
    touch();
    opening += 1;
    const out = [];
    try {
      for (const key of keys) {
        try {
          const done = await memory.reopen(key);
          if (!done.path) continue;
          handed = { path: done.path, at: Date.now() };
          out.push({ key, path: done.path, unsaved: done.unsaved, note: done.note || null, beside: done.beside || null, app: await appForFile(done.path) });
        } catch (err) {
          out.push({ key, failed: err?.message || String(err) });
        }
      }
      return out;
    } finally {
      opening -= 1;
      touch();
    }
  }

  const service = {
    state: stateOut,
    reopen,
    dismissReopen: async ({ keys = [] } = {}) => {
      for (const key of keys) await memory.dismiss(key);
      await publish();
    },
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
          await locker.unlock({ secret }, { shown: shownInWindow });
        } finally {
          secret.fill(0);
        }
      } else await locker.unlock({ password, recovery }, { shown: shownInWindow });
      await afterUnlock();
      return stateOut();
    },
    lock: async () => {
      await doLock('asked');
      return stateOut();
    },
    lookAtFolder: async ({ root }) => {
      touch();
      return locker.lookAtFolder(root);
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
    setWhenLocked: async ({ keepCopies: keep, relock: again }) => {
      if (typeof keep === 'boolean') stores.settings.set('locker.keepCopies', keep);
      if (typeof again === 'boolean') stores.settings.set('locker.relock', again);
      await publish();
      return stateOut();
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
    removeFolder: async ({ id, leaveDamaged = false }) => {
      touch();
      const out = await locker.removeFolder(id, { leaveDamaged });
      startWatching();
      await publish();
      return out;
    },
    // Not a use of the Locker: a folder that changes by itself, a cloud sync for one, refreshes this list.
    list: ({ id }) => locker.scan(id),
    lockWaiting: async ({ id, ...opts }) => {
      touch();
      const out = await locker.lockWaiting(id, opts);
      await publish();
      return out;
    },
    why: ({ id, rel }) => locker.why(id, rel),
    open,
    exportCopy: async ({ id, rel, dest, overwrite, allowConflict, lockedRel, list }) => {
      touch();
      return locker.exportCopy(id, rel, dest, { overwrite, allowConflict, lockedRel, list });
    },
    exportShare: async ({ id, rel, dest, password }) => {
      touch();
      return locker.exportShare(id, rel, dest, password);
    },
    // A locked copy for an email, kept in the profile so a draft can still find it. Only copies a week old are cleared.
    shareForMail: async ({ id, rel, password }) => {
      touch();
      const outgoing = path.join(dir, 'outgoing');
      for (const name of await fs.promises.readdir(outgoing).catch(() => [])) {
        const at = path.join(outgoing, name);
        const st = await fs.promises.stat(at).catch(() => null);
        if (st && Date.now() - st.mtimeMs > 7 * 86400000) await fs.promises.rm(at, { recursive: true, force: true }).catch(() => {});
      }
      const place = path.join(outgoing, crypto.randomBytes(8).toString('hex'));
      await fs.promises.mkdir(place, { recursive: true });
      const dest = path.join(place, `${path.posix.basename(rel)}.rlock`);
      const out = await locker.exportShare(id, rel, dest, password);
      return { dest, size: out.size };
    },
    openShared: ({ file, dest, password }) => locker.openShared(file, dest, password),
    importShared: async ({ file, id, password }) => {
      touch();
      const out = await locker.importShared(file, id, password);
      await publish();
      return out;
    },
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
    renameFile: async ({ id, rel, name, state }) => {
      touch();
      const out = await locker.renameFile(id, rel, name, state);
      await publish();
      return out;
    },
    deleteFile: async ({ id, rel, state, lockedRel, list }) => {
      touch();
      const out = await locker.deleteFile(id, rel, state, { lockedRel, list });
      await publish();
      return out;
    },
    lockOne: async ({ id, rel }) => {
      touch();
      const out = await locker.lockOne(id, rel);
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
    changePassword: async ({ current, next, recovery = null }) => {
      if (!locker.isOpen()) throw new LockerError('locked', 'Unlock the locker first.');
      touch();
      return locker.changePassword(current, next, { recovery });
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
    safetyEstimate: async ({ root }) => {
      touch();
      return locker.safetyEstimate(root);
    },
    checkSafetyCopy: async ({ id }) => {
      touch();
      return locker.checkSafetyCopy(id);
    },
    removeSafetyCopy: async ({ id, accept = false }) => {
      touch();
      const out = await locker.removeSafetyCopy(id, { accept: accept === true });
      await publish();
      return out;
    },
    discardSafetyCopy: async ({ copyId }) => {
      touch();
      const out = await locker.discardSafetyCopy(copyId);
      await publish();
      return out;
    },
    restoreSafetyCopy: async ({ copyId, dest }) => {
      touch();
      return locker.restoreSafetyCopy(copyId, dest);
    },
    leavePlain: async ({ id, rel }) => {
      touch();
      const out = await locker.leavePlain(id, String(rel || ''));
      await publish();
      return out;
    },
    // The person clicked or typed in the Locker's window.
    touch: async () => {
      touch();
      return { lockAt };
    },
  };

  return { locker: service };
}
