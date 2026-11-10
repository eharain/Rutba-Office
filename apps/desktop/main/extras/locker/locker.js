// The Locker as the app and the command line use it: the keys (keys.js), the
// folders the person has added (kept in state.json in Office's own profile),
// and each folder's operations (folder.js), one at a time per folder.

import crypto from 'node:crypto';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createKeys, writeJsonAtomic, FOLDER_KEYS, passwordKey, newKdf } from './keys.js';
import { createFolder, wipeFile, MODE_SHARED, newHeader, sharedKey } from './folder.js';
import { LockerError, openLocked, readHeader, fileHash, MODE_FOLDER } from './format.js';
import { sealFile, writePlainFrom, sweepWorking } from './place.js';
import { OPEN_DIR } from './rules.js';

const OLD_DIR_NAME = '.rlocker-old';
import { makeWorkingDir, isHeld as programHolds, wipeWorking, wipeCopy, findWorkingDirs } from './working.js';

export const DEFAULT_CHOICES = { hidden: true, afterLock: 'delete', keepDir: null, keepVersions: 3, docs: 'rutba', openIn: 'beside' };

/** What a person may set for a folder, and nothing else. */
function cleanChoices(c = {}) {
  const out = {};
  if (typeof c.hidden === 'boolean') out.hidden = c.hidden;
  if (c.afterLock === 'delete' || c.afterLock === 'keep') out.afterLock = c.afterLock;
  if (typeof c.keepDir === 'string' || c.keepDir === null) out.keepDir = c.keepDir;
  if (Number.isInteger(c.keepVersions) && c.keepVersions >= 1 && c.keepVersions <= 3) out.keepVersions = c.keepVersions;
  if (c.docs === 'rutba' || c.docs === 'own') out.docs = c.docs;
  if (c.openIn === 'beside' || c.openIn === 'profile') out.openIn = c.openIn;
  return out;
}

/**
 * @param {object} o
 * @param {string} o.dir   the locker's folder in Office's profile
 * @param {object} [o.cost] scrypt cost (tests pass a cheap one)
 * @param {object} [o.hooks] for the tests
 */
/** How long a mark from another live process is believed without being touched. */
const IN_USE_FOR = 120000;

const UNSAVED_COPY = 'This file is open with changes that could not be saved. Save them somewhere else or close it, then try again.';

export function createLocker({ dir, cost, hooks = null }) {
  const keys = createKeys({ dir, cost });
  const stateFile = path.join(dir, 'state.json');
  const stageDir = path.join(dir, 'stage');
  const openDir = path.join(dir, 'open');
  const queues = new Map();
  const openFilesRecord = path.join(dir, 'open.json');
  const openFiles = new Map(); // working path -> { path, sub, folderId, rel, sha256, seen }
  const isHeld = hooks?.isHeld || programHolds;

  /** Kept on disk before a plain copy exists, so a crash can always be cleaned up after. */
  const recordOpenFiles = () => writeJsonAtomic(openFilesRecord, [...openFiles.values()]);

  // Which locker a working folder belongs to: on a folder a cloud drive shares, another computer's may be in use there.
  const owner = crypto.createHash('sha256').update(`${os.hostname()}|${path.resolve(dir)}`).digest('hex').slice(0, 12);
  const ours = (sub) => !path.basename(sub).includes('-') || path.basename(sub).startsWith(`${owner}-`);

  async function state() {
    let text;
    try {
      text = await fsp.readFile(stateFile, 'utf8');
    } catch (err) {
      if (err.code === 'ENOENT') return { folders: [] };
      throw new LockerError('format', 'The list of locked folders could not be read.');
    }
    try {
      return JSON.parse(text);
    } catch {
      // Never taken for an empty list: the next save would forget every folder.
      throw new LockerError('format', 'The list of locked folders could not be read.');
    }
  }
  const saveState = (s) => writeJsonAtomic(stateFile, s);

  /** One operation at a time on a folder: two windows and the watcher never interleave. */
  function queued(id, fn) {
    const prev = queues.get(id) || Promise.resolve();
    const next = prev.catch(() => {}).then(fn);
    queues.set(id, next);
    return next;
  }

  /** True while the locker is closing: nothing new starts, and what is running stops at its next file. */
  let closing = false;
  let locking = null;
  /** The work on the keys themselves that is not one folder's, kept from running across a lock. */
  const exclusive = (fn) => (closing ? Promise.reject(new LockerError('locked', 'Unlock the locker first.')) : queued('*', async () => {
    if (closing) throw new LockerError('locked', 'Unlock the locker first.');
    return fn();
  }));

  /** Whether the person chose, at the last lock, to keep the plain copies until the next unlock. */
  let keepingCopies = false;

  /** Plain copies that came along in a copied folder are nobody's open file. */
  async function clearCopiedWorking(root) {
    const known = new Set([...openFiles.values()].map((e) => path.resolve(e.sub)));
    for (const d of await findWorkingDirs(root).catch(() => [])) {
      if (!known.has(path.resolve(d)) && ours(d)) await wipeWorking(d);
    }
  }

  function needOpen() {
    if (!keys.isOpen() || closing) throw new LockerError('locked', 'Unlock the locker first.');
  }

  async function folderRecord(id) {
    const s = await state();
    const f = s.folders.find((x) => x.id === id);
    if (!f) throw new LockerError('state', 'That folder is not in the locker.');
    return f;
  }

  /** Tried and not opened by any key this locker knows: not tried again until something about them changes. */
  const strangersTried = new Map();

  /** A name not in use in the folder: Report (copy).pdf, then Report (copy 2).pdf. */
  function freeRel(list, rel) {
    const taken = new Set(list.filter((e) => e.state !== 'damaged').map((e) => e.rel.toLowerCase()));
    if (!taken.has(rel.toLowerCase())) return rel;
    const ext = path.posix.extname(rel);
    const stem = rel.slice(0, rel.length - ext.length);
    for (let n = 1; ; n += 1) {
      const next = `${stem} (copy${n > 1 ? ` ${n}` : ''})${ext}`;
      if (!taken.has(next.toLowerCase())) return next;
    }
  }

  /** The keys that might open what was copied into a folder: every folder of this locker's, and any key file that came with the copy. */
  async function keysFor(rec, strangers, own) {
    const found = keys.knownFolderKeys();
    const dirs = new Set();
    for (const e of strangers) {
      let d = path.posix.dirname(e.lockedRel);
      while (d && d !== '.') {
        dirs.add(d);
        d = path.posix.dirname(d);
      }
    }
    for (const d of dirs) {
      const file = await keys.readFolderKeysFile(path.join(rec.root, ...d.split('/'), FOLDER_KEYS)).catch(() => null);
      const key = file ? await keys.openFolderKeys(file).catch(() => null) : null;
      if (key) found.push(key);
    }
    const seen = new Set([own.toString('hex')]);
    return found.filter((k) => {
      const hex = k.toString('hex');
      if (seen.has(hex)) return false;
      seen.add(hex);
      return true;
    });
  }

  /** What a copied folder brought with it that the host has no use for once its files are the host's own. */
  async function clearCopiedLeftovers(rec, keep) {
    const walk = async (dir, rel) => {
      const names = await fsp.readdir(dir, { withFileTypes: true }).catch(() => []);
      for (const e of names) {
        if (!e.isDirectory() || e.name === OPEN_DIR || e.name === OLD_DIR_NAME || e.name.endsWith('.d')) continue;
        await walk(path.join(dir, e.name), rel ? rel + '/' + e.name : e.name);
      }
      if (!rel || keep(rel + '/')) return;
      if (!names.some((e) => e.name === FOLDER_KEYS)) return;
      await fsp.rm(path.join(dir, FOLDER_KEYS), { force: true }).catch(() => {});
      await fsp.rm(path.join(dir, OLD_DIR_NAME), { recursive: true, force: true }).catch(() => {});
    };
    await walk(rec.root, '');
  }

  async function folderEngine(f) {
    needOpen();
    let key = keys.folderKey(f.id);
    if (!key) {
      // The profile's store lost it (or never had it): the folder's own key file.
      const file = await keys.readFolderKeysFile(path.join(f.root, FOLDER_KEYS));
      if (file) key = await keys.openFolderKeys(file);
      if (!key) throw new LockerError('bad-key', 'The locker cannot open this folder. Enter the password that locked it.');
      keys.rememberFolderKey(f.id, key);
    }
    // Its own copy: the keys are wiped when the locker closes, and work under way must never see them turn to zeros.
    return createFolder({ root: f.root, key: Buffer.from(key), folderId: Buffer.from(f.id, 'hex'), stageDir, choices: f, hooks });
  }

  const refusedWhileClosing = () => Promise.reject(new LockerError('locked', 'Unlock the locker first.'));
  const withFolder = (id, fn) => (closing ? refusedWhileClosing() : queued(id, async () => fn(await folderEngine(await folderRecord(id)))));

  async function changedSince(e) {
    try {
      return (await fileHash(e.path)).sha256 !== e.sha256;
    } catch (err) {
      // Gone is not changed; a file that cannot be read may well be.
      return err.code !== 'ENOENT';
    }
  }

  /** A copy's changes written back and the copy removed. A copy that could not be written back stays, to be tried again. */
  async function finishOpen(e) {
    if (await fsp.stat(e.path).then(() => true, () => false)) {
      try {
        await api.saveOpen(e.path);
      } catch {
        return false;
      }
    }
    if (!(await wipeWorking(e.sub))) return false;
    openFiles.delete(path.resolve(e.path));
    await recordOpenFiles().catch(() => {});
    return true;
  }

  /** A copy removed when it locked, written out again for the window still showing it. */
  async function putBackCopy(e) {
    await withFolder(e.folderId, async (folder) => {
      const found = await folder.entryFor(e.rel);
      if (found?.state !== 'locked') return;
      const o = await folder.opener(found);
      try {
        e.sha256 = (await writePlainFrom(o, e.path)).sha256;
      } finally {
        await o.close();
      }
    });
  }

  /** A file with a copy open cannot be renamed or deleted under whoever has it. */
  function refuseIfOpen(id, rel) {
    for (const e of openFiles.values()) {
      if (e.folderId === id && e.rel.toLowerCase() === rel.toLowerCase()) throw new LockerError('open', 'This file is open. Close it first.');
    }
  }

  /**
   * The copies of a file or folder that is leaving the locker: changes saved,
   * copies removed. One a program holds is cleared at the next unlock. A copy
   * with changes that could not be saved stays on the list and is answered: it
   * is the only place those changes are.
   */
  async function closeCopies(which) {
    const unsaved = [];
    for (const [key, e] of [...openFiles.entries()]) {
      if (!which(e)) continue;
      if (e.sha256 && !(await finishOpen(e)) && await changedSince(e)) {
        unsaved.push(e.rel);
        continue;
      }
      if (openFiles.delete(key)) await recordOpenFiles().catch(() => {});
    }
    return unsaved;
  }

  /**
   * A folder is left alone if its file already has this locker's password and
   * recovery key and opens to the right key. A folder that cannot be reached
   * now (a moved folder, a drive not plugged in) is not recreated. What other
   * lockers put in a key file stays.
   */
  async function healFolders() {
    const s = await state();
    const failed = [];
    for (const f of s.folders) {
      const file = path.join(f.root, FOLDER_KEYS);
      const key = keys.folderKey(f.id);
      if (!key) continue;
      try {
        if (!(await fsp.stat(f.root).then((st) => st.isDirectory(), () => false))) {
          failed.push({ id: f.id, message: 'That folder could not be found.' });
          continue;
        }
        const existing = await keys.readFolderKeysFile(file);
        // Another locked folder's key file, restored here or arrived from a cloud: rewriting it under this folder's id would lock its owners out.
        if (existing?.id && existing.id !== f.id) {
          failed.push({ id: f.id, message: 'This folder holds the key file of a different locked folder, so it was left as it is.' });
          continue;
        }
        const opens = existing ? await keys.openFolderKeys(existing) : null;
        if (!opens || !opens.equals(key) || !(await keys.folderKeysCurrent(existing))) {
          await keys.writeFolderKeysFile(file, await keys.mergedFolderKeys(existing, key, f.id));
        }
      } catch (err) {
        failed.push({ id: f.id, message: err?.message || String(err) });
      }
    }
    return { failed };
  }

  // One program at a time on a profile. The list of open copies, and the queues
  // that keep two operations on a folder apart, belong to the process that holds
  // them: the command line unlocking beside the app would clear the copies the
  // app's windows are using. A crash leaves the mark, so it counts only while
  // that process is alive and has touched it lately.
  const inUse = path.join(dir, 'in-use');
  let beat = null;
  const isRunning = (pid) => {
    try {
      process.kill(pid, 0);
      return true;
    } catch (err) {
      return err.code === 'EPERM';
    }
  };
  async function refuseIfInUse() {
    const held = await fsp.readFile(inUse, 'utf8').then(JSON.parse).catch(() => null);
    if (!held || held.pid === process.pid || !isRunning(held.pid)) return;
    const age = Date.now() - (await fsp.stat(inUse).then((s) => s.mtimeMs, () => 0));
    if (age < IN_USE_FOR) throw new LockerError('busy', 'The Locker is open in another program. Lock it there first.');
  }
  async function markInUse() {
    await fsp.writeFile(inUse, JSON.stringify({ pid: process.pid }));
    clearInterval(beat);
    beat = setInterval(() => {
      const now = new Date();
      fsp.utimes(inUse, now, now).catch(() => {});
    }, IN_USE_FOR / 4);
    beat.unref();
  }
  async function clearInUse() {
    if (!beat) return;
    clearInterval(beat);
    beat = null;
    await fsp.rm(inUse, { force: true }).catch(() => {});
  }

  const api = {
    exists: keys.exists,
    isOpen: keys.isOpen,
    ways: keys.ways,
    keys,

    async setup(password) {
      if (!password) throw new LockerError('password', 'Choose a password for the locker.');
      await fsp.mkdir(dir, { recursive: true });
      const made = await keys.setup(password);
      await markInUse();
      return made;
    },

    /** Opened, and anything a crash left in the working folder cleared. A wrong password changes nothing. */
    async unlock(how, { shown } = {}) {
      await refuseIfInUse();
      await keys.unlock(how);
      await markInUse();
      keepingCopies = false;
      await sweepWorking(stageDir).catch(() => {});
      await api.recoverOpenFiles({ shown }).catch(() => {});
      await api.healFolderKeys().catch(() => {});
    },

    /**
     * Locked: changes in the plain copies written back while the keys are still
     * here, then the keys forgotten and the copies wiped. Answers the files
     * that could not be removed: a program still has them, or they hold changes
     * that could not be written back.
     *
     * The person's two choices: keepCopies leaves the plain copies until the
     * next unlock, and relock locks the files they had unlocked, for as long as
     * budgetMs allows, while the keys are still here.
     *
     * A document an Office window is showing loses its file but keeps the folder
     * the file was in: that window's Save goes there, and what it saves is locked
     * in at the next unlock.
     *
     * @param {{ keepCopies?: boolean, relock?: boolean, budgetMs?: number, shown?: (file: string) => boolean }} [o]
     * @returns {Promise<{ left: string[], open: string[], notLocked: string[], kept: number }>}
     */
    lock(options = {}) {
      locking ||= api.closeLocker(options).finally(() => { locking = null; });
      return locking;
    },

    /** The work of `lock`: one at a time, whichever of a sleep, an idle timer and a press asks first. */
    async closeLocker({ keepCopies = false, relock = false, budgetMs = 30000, shown } = {}) {
      const notLocked = [];
      if (keys.isOpen()) {
        for (const e of [...openFiles.values()]) {
          if (e.sha256) await api.saveOpen(e.path).catch(() => {});
        }
        if (relock) {
          const deadline = Date.now() + budgetMs;
          for (const f of (await state()).folders) {
            const out = await api.lockWaiting(f.id, { deadline }).catch(() => ({ failed: [] }));
            notLocked.push(...out.failed.map((x) => x.rel));
          }
        }
      }
      // Nothing new starts, what is running stops at its next file, and only then do the keys go.
      closing = true;
      try {
        await Promise.allSettled([...queues.values()]);
        keys.lock();
      } finally {
        closing = false;
      }
      await clearInUse();
      if (keepCopies) {
        keepingCopies = openFiles.size > 0;
        await recordOpenFiles().catch(() => {});
        return { left: [], open: [], notLocked, kept: openFiles.size };
      }
      return { ...(await api.wipeOpenFiles({ shown })), notLocked, kept: 0 };
    },

    /** Temporary unlocked copies, gone, except any holding work that is not in the locked file yet. The folder of one an Office window is showing stays. */
    async wipeOpenFiles({ shown = () => false } = {}) {
      const left = [];
      const open = [];
      for (const [key, e] of [...openFiles.entries()]) {
        if (e.sha256 && shown(e.path)) {
          e.seen = true;
          if (await changedSince(e) || !(await wipeCopy(e.path))) left.push(e.rel);
          else open.push(e.rel);
          continue;
        }
        if (e.sha256 && await changedSince(e)) {
          left.push(e.rel);
          continue;
        }
        if (await wipeWorking(e.sub)) openFiles.delete(key);
        else left.push(e.rel);
      }
      await recordOpenFiles().catch(() => {});
      const known = new Set([...openFiles.values()].map((e) => path.resolve(e.sub)));
      for (const sub of await fsp.readdir(openDir).catch(() => [])) {
        const d = path.join(openDir, sub);
        if (!known.has(path.resolve(d))) await wipeWorking(d);
      }
      return { left, open };
    },

    /** The copies a crash or a quit left, known again before anyone unlocks, so unchanged ones can be cleared. */
    async loadOpenRecord() {
      let saved = [];
      try {
        saved = JSON.parse(await fsp.readFile(openFilesRecord, 'utf8'));
      } catch {
        // Nothing was open.
      }
      for (const e of saved) if (!openFiles.has(path.resolve(e.path))) openFiles.set(path.resolve(e.path), e);
    },

    /**
     * After a crash, or a lock that could not finish: changes found in a plain
     * copy become the file's next version, and the copy is removed. Whatever
     * else is in a working folder is not a person's file.
     */
    async recoverOpenFiles({ shown = () => false } = {}) {
      await api.loadOpenRecord();
      const list = (await state()).folders;
      for (const [key, e] of [...openFiles.entries()]) {
        const inLocker = list.some((f) => f.id === e.folderId);
        if (!e.sha256 || !inLocker) {
          if (await wipeWorking(e.sub)) openFiles.delete(key);
          continue;
        }
        // Still open in an Office window: what it saved while locked goes in now. With nothing saved its copy comes back, for the window's next Save and anything it reads.
        if (shown(e.path)) {
          e.seen = true;
          const saved = await fsp.stat(e.path).then(() => true, () => false);
          await (saved ? api.saveOpen(e.path) : putBackCopy(e)).catch(() => {});
          continue;
        }
        if (await isHeld(e.path)) {
          e.seen = true;
          continue;
        }
        await finishOpen(e);
      }
      await recordOpenFiles().catch(() => {});
      const known = new Set([...openFiles.values()].map((e) => path.resolve(e.sub)));
      for (const f of list) {
        for (const d of await findWorkingDirs(f.root)) {
          if (!known.has(path.resolve(d)) && ours(d)) await wipeWorking(d);
        }
      }
    },

    /**
     * Copies whose program has let go, written back and removed. A copy counts
     * as closed only once it has been seen in use: a program that reads a file
     * and lets go at once (Notepad) is still showing it, and its copy waits for
     * the lock.
     *
     * @param {{ shown?: (file: string) => boolean }} o  whether an Office window has the file
     * @returns {Promise<string[]>} the files closed
     */
    async reapClosed({ shown = () => false } = {}) {
      const closed = [];
      for (const [key, e] of [...openFiles.entries()]) {
        if (!e.sha256) continue;
        const inUse = shown(e.path) || await isHeld(e.path);
        if (inUse) {
          e.seen = true;
          continue;
        }
        if (!keys.isOpen()) {
          // The person chose to keep them until the next unlock.
          if (keepingCopies) continue;
          // Locked meanwhile: unchanged copies can go, changed ones wait for the keys.
          if (!(await changedSince(e)) && await wipeWorking(e.sub)) {
            openFiles.delete(key);
            closed.push(e.rel);
          }
        } else if (e.seen && await finishOpen(e)) {
          closed.push(e.rel);
        }
      }
      if (closed.length) await recordOpenFiles().catch(() => {});
      return closed;
    },

    /** A new password, by the old one or, when that is lost, the recovery key. */
    changePassword: (current, next, how = {}) => exclusive(async () => {
      if (!next) throw new LockerError('password', 'Choose a new password.');
      await keys.changePassword(current, next, how);
      return healFolders();
    }),

    /** A new recovery key, reaching every folder's own key file: the old one stops opening them. */
    replaceRecovery: (password) => exclusive(async () => {
      const { recovery } = await keys.replaceRecovery(password);
      const { failed } = await healFolders();
      return { recovery, failed };
    }),
    helloId: () => keys.helloId(),
    enableHello: (secret, id) => keys.enableHello(secret, id),
    disableHello: () => keys.disableHello(),

    /** Each folder's own key file brought up to date with the password and recovery key. */
    healFolderKeys: () => exclusive(healFolders),

    async folders() {
      const s = await state();
      return s.folders.map((f) => ({ ...f }));
    },

    /**
     * What a folder already is, before anything is asked or written: locked
     * (it has a key file), and whether this locker opens it with nothing typed.
     * A key file that cannot be read is an error, never a folder with no key.
     *
     * @returns {Promise<{ locked: boolean, opens: boolean, lockedBelow?: boolean }>}
     */
    async lookAtFolder(root) {
      needOpen();
      const existing = await keys.readFolderKeysFile(path.join(path.resolve(root), FOLDER_KEYS));
      if (!existing) return { locked: false, opens: false, lockedBelow: await hasLockedBelow(path.resolve(root)) };
      const key = keys.folderKey(existing.id) || await keys.openFolderKeys(existing);
      return { locked: true, opens: Boolean(key) };
    },

    /**
     * A folder added to the locker. A folder already locked elsewhere is
     * recognised by its key file, and opens with the password that locked it.
     *
     * @returns {Promise<{ id, root, joined: boolean }>}
     */
    addFolder: (root, choices = {}, how = {}) => exclusive(() => api.addOneFolder(root, choices, how)),

    async addOneFolder(root, choices = {}, { password = null, recovery = null } = {}) {
      needOpen();
      const resolved = path.resolve(root);
      const st = await fsp.stat(resolved).catch(() => null);
      if (!st?.isDirectory()) throw new LockerError('state', 'That folder could not be found.');
      const s = await state();
      if (s.folders.some((f) => path.resolve(f.root).toLowerCase() === resolved.toLowerCase())) throw new LockerError('exists', 'That folder is already in the locker.');
      const merged = { ...DEFAULT_CHOICES, ...cleanChoices(choices) };
      if (merged.afterLock === 'keep' && !merged.keepDir) throw new LockerError('choice', 'Choose where to keep the plain files.');
      if (merged.keepDir && isInside(path.resolve(merged.keepDir), resolved)) throw new LockerError('choice', 'The folder for the plain files must be outside the locked folder.');
      const keyFile = path.join(resolved, FOLDER_KEYS);
      let id;
      let joined = false;
      // Not caught: writing a new key over one that could not be read would lose the folder.
      const existing = await keys.readFolderKeysFile(keyFile);
      if (existing) {
        id = existing.id;
        const twin = s.folders.find((f) => f.id === id);
        if (twin) {
          // Its key file names a folder this locker already has. If that folder has gone from where it was, this is the same one moved.
          const stillThere = await keys.readFolderKeysFile(path.join(twin.root, FOLDER_KEYS)).catch(() => null);
          if (stillThere?.id !== id) {
            twin.root = resolved;
            await saveState(s);
            await clearCopiedWorking(resolved);
            return { id, root: resolved, joined: true, moved: true };
          }
          // Otherwise it is a copy: a folder of its own, under the same key, so what is done to one is never done to the other.
          const key = keys.folderKey(id) || await keys.openFolderKeys(existing, { password, recovery }).catch(() => null);
          if (!key) throw new LockerError('bad-password', 'That password does not open this folder.');
          id = crypto.randomBytes(16).toString('hex');
          keys.rememberFolderKey(id, key);
          await keys.writeFolderKeysFile(keyFile, keys.wrapForFolder(key, id));
          s.folders.push({ id, root: resolved, ...merged, added: new Date().toISOString() });
          await saveState(s);
          await clearCopiedWorking(resolved);
          await api.adoptCopies(id).catch(() => {});
          return { id, root: resolved, joined: true, copied: true };
        }
        let key = keys.folderKey(id);
        if (!key && (password || recovery)) key = await keys.openFolderKeys(existing, { password, recovery });
        if (!key) key = await keys.openFolderKeys(existing);
        if (!key) throw new LockerError('bad-password', recovery ? 'That recovery key does not open this folder.' : 'That password does not open this folder.');
        keys.rememberFolderKey(id, key);
        joined = true;
      } else {
        id = crypto.randomBytes(16).toString('hex');
        const key = crypto.randomBytes(32);
        // The folder's own key file first: without it the folder could not open elsewhere.
        await keys.writeFolderKeysFile(keyFile, keys.wrapForFolder(key, id));
        keys.rememberFolderKey(id, key);
      }
      s.folders.push({ id, root: resolved, ...merged, added: new Date().toISOString() });
      await saveState(s);
      await clearCopiedWorking(resolved);
      await api.adoptCopies(id).catch(() => {});
      return { id, root: resolved, joined };
    },

    /**
     * Locked files copied into this folder from another locked folder of this
     * locker, or from elsewhere in the folder itself, made its own: they would
     * otherwise show as damaged, or share one file's id. Nothing is ever plain
     * in the folder for this; a file that no known key opens is left as it is.
     *
     * @returns {Promise<{ adopted: string[] }>}
     */
    adoptCopies: (id) => withFolder(id, async (host) => {
      const rec = await folderRecord(id);
      const list = await host.scan();
      // A folder inside this one that is in the locker in its own right keeps its own files and key.
      const inner = (await state()).folders
        .filter((f) => f.id !== id && isInside(path.resolve(f.root), path.resolve(rec.root)))
        .map((f) => path.relative(rec.root, f.root).split(path.sep).join('/') + '/');
      const theirs = (rel) => inner.some((p) => rel.startsWith(p));
      const strangers = list.filter((e) => e.state === 'damaged' && !theirs(e.lockedRel));
      const firstOf = new Map();
      const twins = [];
      for (const e of [...list].filter((x) => x.fileId && x.state === 'locked' && !theirs(x.lockedRel)).sort((a, b) => a.lockedRel.length - b.lockedRel.length || a.lockedRel.localeCompare(b.lockedRel))) {
        if (firstOf.has(e.fileId)) twins.push(e);
        else firstOf.set(e.fileId, e);
      }
      const adopted = [];
      const take = async (from, e, taken) => {
        const stage = await from.toStage(e);
        try {
          await fsp.utimes(stage, new Date(e.modified), new Date(e.modified)).catch(() => {});
          const rel = freeRel(taken, e.rel);
          await host.adopt(rel, stage, e.path);
          taken.push({ rel, state: 'locked' });
          adopted.push(rel);
        } finally {
          await fsp.rm(stage, { force: true }).catch(() => {});
        }
      };
      const others = list.filter((e) => e.state !== 'damaged' && !twins.includes(e));
      for (const e of twins) await take(host, e, others).catch(() => {});
      if (strangers.length) {
        const signature = strangers.map((e) => e.path).sort().join('|');
        if (strangersTried.get(id) !== signature) {
          const own = keys.folderKey(id);
          for (const key of await keysFor(rec, strangers, own)) {
            const other = createFolder({ root: rec.root, key, folderId: Buffer.from(id, 'hex'), stageDir, choices: rec, hooks });
            const wanted = new Set(strangers.map((e) => e.lockedRel));
            for (const e of await other.scan()) {
              if (e.state === 'locked' && wanted.has(e.lockedRel)) await take(other, e, others).catch(() => {});
            }
          }
          const after = (await host.scan()).filter((e) => e.state === 'damaged').map((e) => e.path).sort().join('|');
          if (after) strangersTried.set(id, after);
          else strangersTried.delete(id);
        }
      }
      if (adopted.length) {
        const damaged = (await host.scan()).filter((e) => e.state === 'damaged').map((e) => e.lockedRel);
        await clearCopiedLeftovers(rec, (dir) => theirs(dir) || damaged.some((r) => r.startsWith(dir)));
      }
      return { adopted };
    }),

    /**
     * The ready-made folder, in the person's own home folder and so outside
     * the folders a cloud drive normally takes. Asked again, it is the same one.
     *
     * @returns {Promise<{ id, root, made: boolean }>}
     */
    async makeHomeFolder(home) {
      needOpen();
      const root = path.join(home, 'Locker');
      const known = (await state()).folders.find((f) => path.resolve(f.root).toLowerCase() === root.toLowerCase());
      if (known) return { id: known.id, root: known.root, made: false };
      await fsp.mkdir(root, { recursive: true });
      const { id } = await api.addFolder(root);
      return { id, root, made: true };
    },

    /** A folder's choices changed. */
    async setChoices(id, patch) {
      const clean = cleanChoices(patch);
      const s = await state();
      const f = s.folders.find((x) => x.id === id);
      if (!f) throw new LockerError('state', 'That folder is not in the locker.');
      const next = { ...f, ...clean };
      if (next.afterLock === 'keep' && !next.keepDir) throw new LockerError('choice', 'Choose where to keep the plain files.');
      if (next.keepDir && isInside(path.resolve(next.keepDir), f.root)) throw new LockerError('choice', 'The folder for the plain files must be outside the locked folder.');
      Object.assign(f, clean);
      await saveState(s);
      return f;
    },

    /**
     * A folder taken out of the locker: every file unlocked first, for good,
     * and only then forgotten. A file that will not unlock stops it, and the
     * folder stays in the locker so nothing is left unreachable.
     */
    async removeFolder(id, { leaveDamaged = false } = {}) {
      const unsaved = await closeCopies((e) => e.folderId === id);
      if (unsaved.length) return { removed: false, failed: unsaved.map((rel) => ({ rel, message: UNSAVED_COPY })) };
      return withFolder(id, async (folder) => {
        const failed = [];
        const entries = await folder.scan();
        // A file that will not open is not gone: it may be a newer version's, or in use, or still on its way from the cloud.
        const damaged = entries.filter((e) => e.state === 'damaged');
        if (damaged.length && !leaveDamaged) return { removed: false, damaged: damaged.length, failed: damaged.map((e) => ({ rel: e.rel, message: e.message })) };
        // Which of two versions becomes the plain file is the person's to say, and nothing is changed until they have.
        const twoOf = entries.filter((e) => e.state === 'conflict');
        if (twoOf.length) return { removed: false, conflicts: new Set(twoOf.map((e) => e.rel)).size, failed: twoOf.map((e) => ({ rel: e.rel, message: 'This file has two versions. Compare them first.' })) };
        for (const e of entries) {
          if (e.state !== 'locked') continue;
          try {
            await folder.unlockForGood(e.rel);
          } catch (err) {
            failed.push({ rel: e.rel, message: err?.message || String(err) });
          }
        }
        if (failed.length) return { removed: false, failed };
        const s = await state();
        const f = s.folders.find((x) => x.id === id);
        // Files left locked still need the key file that opens them.
        if (!damaged.length) {
          await fsp.rm(path.join(f.root, FOLDER_KEYS), { force: true }).catch(() => {});
          await fsp.rm(path.join(f.root, '.rlocker-old'), { recursive: true, force: true }).catch(() => {});
        }
        keys.forgetFolderKey(id);
        s.folders = s.folders.filter((x) => x.id !== id);
        await saveState(s);
        return { removed: true, failed: [] };
      });
    },

    scan: (id) => withFolder(id, (f) => f.scan()),
    why: (id, rel) => withFolder(id, (f) => f.why(rel)),
    lockWaiting: (id, opts) => withFolder(id, (f) => f.lockWaiting({ ...opts, shouldStop: () => closing })),
    lockOne: (id, rel, opts) => withFolder(id, (f) => f.lockOne(rel, opts)),
    exportCopy: (id, rel, dest, opts) => withFolder(id, (f) => f.exportCopy(rel, dest, opts)),
    /** Back to a plain file where it is. It stays plain until the person locks it again. */
    unlockForGood: async (id, rel) => {
      if ((await closeCopies((e) => e.folderId === id && e.rel === rel)).length) throw new LockerError('open', UNSAVED_COPY);
      return withFolder(id, (f) => f.unlockForGood(rel));
    },
    renameFile: async (id, rel, name, state) => {
      refuseIfOpen(id, rel);
      return withFolder(id, (f) => f.rename(rel, name, state));
    },
    deleteFile: async (id, rel, state, how) => {
      refuseIfOpen(id, rel);
      return withFolder(id, (f) => f.removeFile(rel, state, how));
    },
    versions: (id, rel) => withFolder(id, (f) => f.versions(rel)),
    restore: (id, rel, version) => withFolder(id, (f) => f.restore(rel, version)),
    readRules: (id) => withFolder(id, (f) => f.readRules()),
    writeRules: (id, text) => withFolder(id, (f) => f.writeRules(text)),

    /** A file brought in from elsewhere: locked into the folder at `rel`; the original is not touched. */
    importFile: (id, source, rel) => withFolder(id, (f) => f.saveFrom(rel || path.basename(source), source)),

    /** One file locked with a password of its own, for someone else. */
    async exportShare(id, rel, dest, password) {
      if (!password) throw new LockerError('password', 'Choose a password for the shared file.');
      if ((await state()).folders.some((f) => isInside(path.resolve(dest), path.resolve(f.root)))) throw new LockerError('choice', 'Save the copy for sending outside your locked folders.');
      return withFolder(id, async (folder) => {
        const e = await folder.entryFor(rel);
        if (e?.state === 'conflict') throw new LockerError('conflict', 'This file has two versions. Compare them first.');
        if (!e || e.state === 'waiting' || e.state === 'excluded' || e.state === 'damaged') throw new LockerError('state', 'That file is not locked.');
        const kdf = newKdf(cost);
        const key = await passwordKey(password, kdf);
        const headerBuf = newHeader({ mode: MODE_SHARED, scrypt: { salt: Buffer.from(kdf.salt, 'hex'), N: kdf.N, r: kdf.r, p: kdf.p } });
        const tmp = await folder.toStage(e);
        try {
          await sealFile({ source: tmp, target: dest, rootKey: key, headerBuf, meta: { name: e.name, shared: true }, stageDir });
          return { dest, size: e.size };
        } finally {
          await wipeFile(tmp);
          key.fill(0);
        }
      });
    },

    /** A shared file opened with its own password, into a plain file. */
    async openShared(file, dest, password) {
      const header = await readHeader(file);
      if (header.mode !== MODE_SHARED) throw new LockerError('state', 'This file was not locked for sharing.');
      const key = await sharedKey(password, header);
      const o = await openLocked(file, key).catch((err) => {
        if (err.code === 'bad-key') throw new LockerError('bad-password', 'That password does not open this file.');
        throw err;
      });
      try {
        return await writePlainFrom(o, dest);
      } finally {
        await o.close();
        key.fill(0);
      }
    },

    /**
     * A file someone sent, locked into a folder of this locker under the name it
     * arrived with, or "name (copy)" if that is taken. The received file is left as it is.
     */
    async importShared(file, id, password) {
      const name = path.basename(file).replace(/\.rlock$/i, '') || 'Received file';
      return withFolder(id, async (folder) => {
        const tmp = path.join(stageDir, `.rl-${crypto.randomBytes(12).toString('hex')}.part`);
        try {
          await api.openShared(file, tmp, password);
          const rel = freeRel(await folder.scan(), name);
          await folder.saveFrom(rel, tmp);
          return { rel };
        } finally {
          await wipeFile(tmp);
        }
      });
    },

    /**
     * A temporary plain copy of a locked file for a program to open, in a
     * hidden folder beside the file, or in the profile if the person chose
     * that for this folder. Answers the path.
     */
    async openPlain(id, rel) {
      const rec = await folderRecord(id);
      return withFolder(id, async (folder) => {
        const e = await folder.entryFor(rel);
        if (e?.state === 'conflict') throw new LockerError('conflict', 'This file has two versions. Compare them first.');
        if (!e || (e.state !== 'locked')) throw new LockerError('state', 'That file is not locked.');
        const top = rec.openIn === 'profile' ? openDir : path.join(path.dirname(path.join(rec.root, ...e.rel.split('/'))), OPEN_DIR);
        const sub = await makeWorkingDir(top, owner);
        const dest = path.join(sub, e.name);
        const key = path.resolve(dest);
        const entry = { path: dest, sub, folderId: id, rel: e.rel, sha256: null, seen: false };
        openFiles.set(key, entry);
        const o = await folder.opener(e);
        try {
          await recordOpenFiles();
          const done = await writePlainFrom(o, dest);
          entry.sha256 = done.sha256;
          await recordOpenFiles();
          return { path: dest, size: done.size };
        } catch (err) {
          openFiles.delete(key);
          await wipeWorking(sub);
          await recordOpenFiles().catch(() => {});
          throw err;
        } finally {
          await o.close();
        }
      });
    },

    /** The working copy written back as the file's next version if it has changed. */
    async saveOpen(dest) {
      const key = path.resolve(dest);
      const open = openFiles.get(key);
      if (!open) throw new LockerError('state', 'That file was not opened from the locker.');
      return withFolder(open.folderId, async (folder) => {
        const h = await fileHash(dest);
        if (h.sha256 === open.sha256) return { changed: false };
        const sealed = await folder.saveFrom(open.rel, dest);
        open.sha256 = sealed.sha256;
        await recordOpenFiles();
        return { changed: true, ...sealed };
      });
    },

    /** Whether a path is one of the working copies. */
    isOpenCopy: (p) => openFiles.has(path.resolve(p)),

    /** The working copy a path belongs to, the path being anything in its folder: programs save by writing a new file and renaming it over the old. */
    openCopyFor(p) {
      const full = path.resolve(p);
      for (const e of openFiles.values()) {
        if (isInside(full, path.resolve(e.sub))) return e.path;
      }
      return null;
    },

    hasOpenCopies: () => openFiles.size > 0,
    /** The plain copies open now, for asking whether an Office window has one. */
    openCopyPaths: () => [...openFiles.values()].filter((e) => e.sha256).map((e) => e.path),

    /** What a folder's state amounts to, for its row. */
    async summary(id) {
      const list = await api.scan(id);
      const count = (s) => list.filter((e) => e.state === s).length;
      return { locked: count('locked'), waiting: count('waiting'), excluded: count('excluded'), conflict: count('conflict'), damaged: count('damaged'), total: list.length };
    },

    /** The most recent locked copy's header-only fact: no secret. */
    headerOf: (p) => readHeader(p).then((h) => ({ format: h.version, mode: h.mode === MODE_FOLDER ? 'folder' : 'shared', fileId: h.fileId.toString('hex') })),

    /** The state file's path, for a test. */
    get stateFile() { return stateFile; },
    get stageDir() { return stageDir; },
  };
  return api;
}

/** Whether anything under this folder is already locked: a key file or a locked file, looked for a few levels down. */
async function hasLockedBelow(root) {
  let left = 5000;
  const go = async (dir, depth) => {
    for (const e of await fsp.readdir(dir, { withFileTypes: true }).catch(() => [])) {
      if (left-- <= 0) return false;
      if (e.name === FOLDER_KEYS || e.name.endsWith('.rlock') || e.name.endsWith('.rlock.d')) return true;
      if (e.isDirectory() && depth < 6 && !e.name.startsWith('.') && await go(path.join(dir, e.name), depth + 1)) return true;
    }
    return false;
  };
  return go(root, 0);
}

function isInside(p, root) {
  const rel = path.relative(root, p);
  return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
}

