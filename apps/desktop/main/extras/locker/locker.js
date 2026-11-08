// The Locker as the app and the command line use it: the keys (keys.js), the
// folders the person has added (kept in state.json in Office's own profile),
// and each folder's operations (folder.js), one at a time per folder.

import crypto from 'node:crypto';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { createKeys, writeJsonAtomic, FOLDER_KEYS, passwordKey, newKdf } from './keys.js';
import { createFolder, wipeFile, MODE_SHARED, newHeader, sharedKey } from './folder.js';
import { LockerError, openLocked, readHeader, MODE_FOLDER } from './format.js';
import { sealFile, writePlainFrom, sweepWorking } from './place.js';

export const DEFAULT_CHOICES = { hidden: true, afterLock: 'delete', keepDir: null, keepVersions: 3, docs: 'rutba' };

/** What a person may set for a folder, and nothing else. */
function cleanChoices(c = {}) {
  const out = {};
  if (typeof c.hidden === 'boolean') out.hidden = c.hidden;
  if (c.afterLock === 'delete' || c.afterLock === 'keep') out.afterLock = c.afterLock;
  if (typeof c.keepDir === 'string' || c.keepDir === null) out.keepDir = c.keepDir;
  if (Number.isInteger(c.keepVersions) && c.keepVersions >= 1 && c.keepVersions <= 3) out.keepVersions = c.keepVersions;
  if (c.docs === 'rutba' || c.docs === 'own') out.docs = c.docs;
  return out;
}

/**
 * @param {object} o
 * @param {string} o.dir   the locker's folder in Office's profile
 * @param {object} [o.cost] scrypt cost (tests pass a cheap one)
 * @param {object} [o.hooks] for the tests
 */
export function createLocker({ dir, cost, hooks = null }) {
  const keys = createKeys({ dir, cost });
  const stateFile = path.join(dir, 'state.json');
  const stageDir = path.join(dir, 'stage');
  const openDir = path.join(dir, 'open');
  const queues = new Map();
  const openFiles = new Map(); // working path -> { folderId, rel, sha256 }

  async function state() {
    try {
      return JSON.parse(await fsp.readFile(stateFile, 'utf8'));
    } catch {
      return { folders: [] };
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

  function needOpen() {
    if (!keys.isOpen()) throw new LockerError('locked', 'Unlock the locker first.');
  }

  async function folderRecord(id) {
    const s = await state();
    const f = s.folders.find((x) => x.id === id);
    if (!f) throw new LockerError('state', 'That folder is not in the locker.');
    return f;
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
    return createFolder({ root: f.root, key, folderId: Buffer.from(f.id, 'hex'), stageDir, choices: f, hooks });
  }

  const withFolder = (id, fn) => queued(id, async () => fn(await folderEngine(await folderRecord(id))));

  const api = {
    exists: keys.exists,
    isOpen: keys.isOpen,
    ways: keys.ways,
    keys,

    async setup(password) {
      if (!password) throw new LockerError('password', 'Choose a password for the locker.');
      await fsp.mkdir(dir, { recursive: true });
      return keys.setup(password);
    },

    /** Opened, and anything a crash left in the working folder cleared. A wrong password changes nothing. */
    async unlock(how) {
      await keys.unlock(how);
      await sweepWorking(stageDir).catch(() => {});
      await api.healFolderKeys().catch(() => {});
    },

    /** Locked: keys forgotten, every temporary plain copy wiped. */
    async lock() {
      keys.lock();
      await api.wipeOpenFiles();
    },

    /** Temporary unlocked copies, gone. */
    async wipeOpenFiles() {
      openFiles.clear();
      let names = [];
      try {
        names = await fsp.readdir(openDir);
      } catch {
        return;
      }
      for (const sub of names) {
        const d = path.join(openDir, sub);
        const files = await fsp.readdir(d).catch(() => []);
        for (const f of files) await wipeFile(path.join(d, f));
        await fsp.rm(d, { recursive: true, force: true }).catch(() => {});
      }
    },

    async changePassword(current, next) {
      if (!next) throw new LockerError('password', 'Choose a new password.');
      await keys.changePassword(current, next);
      return api.healFolderKeys();
    },

    replaceRecovery: (password) => keys.replaceRecovery(password),
    enableHello: (secret) => keys.enableHello(secret),
    disableHello: () => keys.disableHello(),

    /** Each folder's own key file brought up to date with the password and recovery key: a folder is left alone if its file already opens. */
    async healFolderKeys() {
      const s = await state();
      const failed = [];
      for (const f of s.folders) {
        const file = path.join(f.root, FOLDER_KEYS);
        const key = keys.folderKey(f.id);
        if (!key) continue;
        try {
          const existing = await keys.readFolderKeysFile(file);
          const opens = existing ? await keys.openFolderKeys(existing) : null;
          if (!opens || !opens.equals(key)) await keys.writeFolderKeysFile(file, keys.wrapForFolder(key, f.id));
        } catch (err) {
          failed.push({ id: f.id, message: err?.message || String(err) });
        }
      }
      return { failed };
    },

    async folders() {
      const s = await state();
      return s.folders.map((f) => ({ ...f }));
    },

    /**
     * A folder added to the locker. A folder already locked elsewhere is
     * recognised by its key file, and opens with the password that locked it.
     *
     * @returns {Promise<{ id, root, joined: boolean }>}
     */
    async addFolder(root, choices = {}, { password = null, recovery = null } = {}) {
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
      const existing = await keys.readFolderKeysFile(keyFile).catch(() => null);
      if (existing) {
        id = existing.id;
        let key = keys.folderKey(id);
        if (!key && (password || recovery)) key = await keys.openFolderKeys(existing, { password, recovery });
        if (!key) key = await keys.openFolderKeys(existing);
        if (!key) throw new LockerError('bad-password', 'That password does not open this folder.');
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
      return { id, root: resolved, joined };
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
    async removeFolder(id) {
      return withFolder(id, async (folder) => {
        const failed = [];
        for (const e of await folder.scan()) {
          if (e.state !== 'locked' && e.state !== 'conflict') continue;
          try {
            await folder.unlockForGood(e.rel);
          } catch (err) {
            failed.push({ rel: e.rel, message: err?.message || String(err) });
          }
        }
        if (failed.length) return { removed: false, failed };
        const s = await state();
        const f = s.folders.find((x) => x.id === id);
        await fsp.rm(path.join(f.root, FOLDER_KEYS), { force: true }).catch(() => {});
        await fsp.rm(path.join(f.root, '.rlocker-old'), { recursive: true, force: true }).catch(() => {});
        keys.forgetFolderKey(id);
        s.folders = s.folders.filter((x) => x.id !== id);
        await saveState(s);
        return { removed: true, failed: [] };
      });
    },

    scan: (id) => withFolder(id, (f) => f.scan()),
    why: (id, rel) => withFolder(id, (f) => f.why(rel)),
    lockWaiting: (id, opts) => withFolder(id, (f) => f.lockWaiting(opts)),
    lockOne: (id, rel, opts) => withFolder(id, (f) => f.lockOne(rel, opts)),
    exportCopy: (id, rel, dest, opts) => withFolder(id, (f) => f.exportCopy(rel, dest, opts)),
    /** Back to a plain file for good: it stays plain, even if the rules say to lock it. */
    unlockForGood: (id, rel) => withFolder(id, async (f) => {
      const out = await f.unlockForGood(rel);
      const s = await state();
      const rec = s.folders.find((x) => x.id === id);
      rec.released = [...new Set([...(rec.released || []), rel])];
      await saveState(s);
      return out;
    }),
    versions: (id, rel) => withFolder(id, (f) => f.versions(rel)),
    restore: (id, rel, version) => withFolder(id, (f) => f.restore(rel, version)),
    readRules: (id) => withFolder(id, (f) => f.readRules()),
    writeRules: (id, text) => withFolder(id, (f) => f.writeRules(text)),

    /** A file brought in from elsewhere: locked into the folder at `rel`; the original is not touched. */
    importFile: (id, source, rel) => withFolder(id, (f) => f.saveFrom(rel || path.basename(source), source)),

    /** One file locked with a password of its own, for someone else. */
    async exportShare(id, rel, dest, password) {
      if (!password) throw new LockerError('password', 'Choose a password for the shared file.');
      return withFolder(id, async (folder) => {
        const e = await folder.entryFor(rel);
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
     * A temporary plain copy of a locked file for an app to open, in the
     * profile's working folder, wiped on lock. Answers the path.
     */
    async openPlain(id, rel) {
      return withFolder(id, async (folder) => {
        const e = await folder.entryFor(rel);
        if (!e || (e.state !== 'locked')) throw new LockerError('state', 'That file is not locked.');
        const sub = path.join(openDir, crypto.randomBytes(8).toString('hex'));
        await fsp.mkdir(sub, { recursive: true });
        const dest = path.join(sub, e.name);
        const o = await folder.opener(e);
        try {
          const done = await writePlainFrom(o, dest);
          openFiles.set(path.resolve(dest), { folderId: id, rel: e.rel, sha256: done.sha256 });
          return { path: dest, size: done.size };
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
        const { fileHash } = await import('./format.js');
        const h = await fileHash(dest);
        if (h.sha256 === open.sha256) return { changed: false };
        const sealed = await folder.saveFrom(open.rel, dest);
        openFiles.set(key, { ...open, sha256: sealed.sha256 });
        return { changed: true, ...sealed };
      });
    },

    /** Whether a path is one of the working copies. */
    isOpenCopy: (p) => openFiles.has(path.resolve(p)),

    /** What a folder's state amounts to, for its row. */
    async summary(id) {
      const list = await api.scan(id);
      const count = (s) => list.filter((e) => e.state === s).length;
      return { locked: count('locked'), waiting: count('waiting'), excluded: count('excluded'), conflict: count('conflict'), damaged: count('damaged'), total: list.length };
    },

    /** The most recent locked copy's header-only fact: no secret. */
    headerOf: (p) => readHeader(p).then((h) => ({ mode: h.mode === MODE_FOLDER ? 'folder' : 'shared', fileId: h.fileId.toString('hex') })),

    /** The state file's path, for a test. */
    get stateFile() { return stateFile; },
    get stageDir() { return stageDir; },
  };
  return api;
}

function isInside(p, root) {
  const rel = path.relative(root, p);
  return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
}

