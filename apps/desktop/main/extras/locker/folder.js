// One locked folder: what is in it, and the operations on its files.
//
// What is on disk in the folder:
//   the plain files not yet locked (`waiting`) or excluded by the rules
//   `<name>.rlock` (or a random name, when names are hidden) one locked file
//   `<name>.rlock.d`  a large file in pieces (pieces.js)
//   `.rlocker-old/<fileId>/`  earlier versions of smaller files, also locked
//   `.rutbalocker` the rules, `.rutbalocker-keys` the folder's sealed key
//
// The rules of this module, and the tests that hold it to them:
//   - a locked copy is read back and matched before a plain file is touched;
//   - a plain file is deleted only when the person chose that, and only if it
//     still is what was locked; otherwise it is left alone;
//   - nothing here deletes a locked file except "unlock for good" and pruning
//     versions the person asked not to keep, and those only after the plain
//     copy (or newer version) has been checked.

import crypto from 'node:crypto';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { newHeader, readHeader, openLocked, verifyLocked, fileHash, LockerError, HEADER_SIZE, MODE_FOLDER, MODE_SHARED } from './format.js';
import { sealFile, writePlain, writePlainFrom, isWorkingFile, workingName } from './place.js';
import { sealPieces, openPieces, listsIn, readItem, removeNewestList } from './pieces.js';
import { parseRules, decide, isLockerOwn, OPEN_DIR } from './rules.js';

export const OLD_DIR = '.rlocker-old';
export const LARGE = 50 * 1024 * 1024;
const LOCK_EXT = '.rlock';
const PIECES_EXT = '.rlock.d';

const posix = (p) => p.split(path.sep).join('/');
const randomName = () => `${crypto.randomBytes(16).toString('hex')}${LOCK_EXT}`;

/**
 * A file moved under a working name, where it stays until it has gone, or null
 * if it cannot be. A program can hold a file so that it may be written but not
 * deleted. Windows then opens it for writing and refuses only the delete, so
 * overwriting first would leave that program a file of zeros; a move needs the
 * same permission a delete does. And a stop part way through must not leave
 * zeros under the file's own name, to be locked as its next version.
 */
async function movedAside(file) {
  const moved = path.join(path.dirname(file), workingName());
  try {
    await fsp.rename(file, moved);
    return moved;
  } catch {
    return null;
  }
}

/**
 * A file's bytes overwritten and removed, as far as a program can: an SSD may
 * keep the old blocks, and the Locker says so. A file another program holds is
 * left exactly as it is, and still there for the caller to notice.
 */
export async function wipeFile(file) {
  let target = file;
  try {
    // Never through a link: what it points at is not this file.
    const st = await fsp.lstat(file);
    if (st.isFile()) {
      target = await movedAside(file);
      if (!target) return;
      const fh = await fsp.open(target, 'r+');
      try {
        const zeros = Buffer.alloc(1024 * 1024);
        for (let at = 0; at < st.size; at += zeros.length) await fh.write(zeros, 0, Math.min(zeros.length, st.size - at), at);
        await fh.sync();
      } finally {
        await fh.close();
      }
    }
  } catch {
    // A file that cannot be overwritten is removed as it is, below.
  }
  await fsp.rm(target, { force: true }).catch(() => {});
}

const BAD_NAME = /[<>:"/\\|?*\x00-\x1f]|[. ]$/;
const RESERVED_NAME = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i;

/** A name a person typed, if it can be a file's name here. */
function checkedName(typed) {
  const name = String(typed ?? '').trim();
  if (!name || BAD_NAME.test(name) || RESERVED_NAME.test(name) || name.endsWith(LOCK_EXT) || isLockerOwn(name) || isWorkingFile(name)) {
    throw new LockerError('name', 'That name cannot be used for a file.');
  }
  return name;
}

/** A locked file's own record of its name, if it is a name and not a way out of the folder. */
function recordedName(name) {
  if (typeof name !== 'string' || !name || name === '.' || name === '..' || /[/\\:\x00]/.test(name)) throw new LockerError('format', 'This locked file records a name that cannot be used.');
  return name;
}

/**
 * @param {object} o
 * @param {string} o.root       the folder's path
 * @param {Buffer} o.key        the folder's key
 * @param {Buffer} o.folderId   16 bytes
 * @param {string} o.stageDir   a working folder in Office's profile
 * @param {object} o.choices    { hidden, afterLock: 'delete'|'keep', keepDir, keepVersions }
 * @param {object} [o.hooks]    for the tests
 * @param {number} [o.large]    the size over which a file is kept as pieces: LARGE, smaller in the tests
 *                              so the pieces are made in seconds (12 Nov 2026)
 */
export function createFolder({ root, key, folderId, stageDir, choices, hooks = null, large: largeAt = LARGE }) {
  const opt = () => ({ hidden: false, afterLock: 'keep', keepDir: null, keepVersions: 3, ...choices });
  const abs = (rel) => path.join(root, ...rel.split('/'));

  /** Every file under the folder, with the rules files found along the way. */
  async function walk() {
    const files = [];
    const lockedFiles = [];
    const pieceDirs = [];
    const sets = [];
    const stray = [];
    const go = async (dir, rel) => {
      let names;
      try {
        names = await fsp.readdir(dir, { withFileTypes: true });
      } catch {
        return;
      }
      for (const e of names.sort((a, b) => a.name.localeCompare(b.name))) {
        if (e.name === '.rutbalocker' && e.isFile()) {
          try {
            sets.push({ dir: rel, rules: parseRules(await fsp.readFile(path.join(dir, e.name), 'utf8')) });
          } catch {
            // An unreadable rules file says nothing; it is shown as it is when edited.
          }
        }
      }
      for (const e of names) {
        const p = path.join(dir, e.name);
        const r = rel ? `${rel}/${e.name}` : e.name;
        if (e.isDirectory()) {
          if (e.name === OLD_DIR || e.name === OPEN_DIR) continue;
          if (e.name.endsWith(PIECES_EXT)) {
            pieceDirs.push({ path: p, rel: r });
            continue;
          }
          await go(p, r);
        } else if (e.isFile()) {
          if (isWorkingFile(e.name)) {
            stray.push(p);
            continue;
          }
          if (e.name.endsWith(LOCK_EXT)) lockedFiles.push({ path: p, rel: r });
          else if (!isLockerOwn(e.name)) files.push({ path: p, rel: r });
        }
      }
    };
    await go(root, '');
    sets.sort((a, b) => a.dir.length - b.dir.length);
    return { files, lockedFiles, pieceDirs, sets, stray };
  }

  const rulesFor = (sets, rel) => sets.filter((s) => s.dir === '' || rel.startsWith(`${s.dir}/`) || rel === s.dir);

  /**
   * What the folder holds: every file with its state and the reason.
   *
   * @returns {Promise<Array<{ rel, name, state: 'locked'|'waiting'|'excluded'|'conflict'|'damaged', why, size, modified, path, pieces, fileId, otherOf, message }>>}
   */
  async function scan() {
    const { files, lockedFiles, pieceDirs, sets } = await walk();
    const entries = [];
    const byId = new Map();
    // Keyed by directory too: a cloud's conflict copy sits beside the file, and a copy someone made elsewhere is not a conflict.
    const note = (fileId, entry) => {
      const key = `${fileId}|${path.posix.dirname(entry.lockedRel)}`;
      if (!byId.has(key)) byId.set(key, []);
      byId.get(key).push(entry);
    };
    for (const { path: p, rel } of lockedFiles) {
      const dir = rel.includes('/') ? rel.slice(0, rel.lastIndexOf('/')) : '';
      try {
        const h = await readHeader(p);
        const o = await openLocked(p, key);
        const meta = o.meta;
        await o.close();
        const name = recordedName(meta.name);
        const realRel = opt().hidden || !rel.endsWith(`${name}${LOCK_EXT}`) ? (dir ? `${dir}/${name}` : name) : rel.slice(0, -LOCK_EXT.length);
        const entry = { rel: realRel, name, state: 'locked', size: meta.size, modified: meta.modified, path: p, lockedRel: rel, sha256: meta.sha256, fileId: h.fileId.toString('hex'), why: null };
        entries.push(entry);
        note(entry.fileId, entry);
      } catch (err) {
        entries.push({ rel, name: path.basename(p), state: 'damaged', newer: err?.code === 'newer' || undefined, why: null, message: err?.message || String(err), path: p, lockedRel: rel });
      }
    }
    for (const { path: p, rel } of pieceDirs) {
      const dir = rel.includes('/') ? rel.slice(0, rel.lastIndexOf('/')) : '';
      try {
        const item = await readItem(p);
        const opened = await openPieces({ dir: p, folderKey: key });
        const lists = await listsIn(p);
        const name = recordedName(opened.meta.name);
        const realRel = dir ? `${dir}/${name}` : name;
        const entry = { rel: realRel, name, state: 'locked', size: opened.meta.size, modified: opened.meta.modified, path: p, lockedRel: rel, sha256: opened.meta.sha256, fileId: item.fileId, pieces: opened.meta.pieces, why: null };
        if (lists.length > 1 && lists[0].seq === lists[1].seq) {
          entry.conflictNote = true;
          // Two computers' versions in the one folder: each has its own row, to be saved or removed on its own.
          const other = await openPieces({ dir: p, folderKey: key, version: 1 }).catch(() => null);
          if (other) entry.twoLists = [{ file: lists[0].file, size: opened.meta.size, modified: opened.meta.modified }, { file: lists[1].file, size: other.meta.size, modified: other.meta.modified }];
        }
        entries.push(entry);
        note(entry.fileId, entry);
      } catch (err) {
        entries.push({ rel, name: path.basename(p), state: 'damaged', newer: err?.code === 'newer' || undefined, why: null, message: err?.message || String(err), path: p, lockedRel: rel });
      }
    }
    // Two locked copies with one file's id: the cloud made a conflict copy. Shown side by side, never merged.
    for (const group of byId.values()) {
      if (group.length > 1 || group.some((e) => e.conflictNote)) {
        for (const e of group) {
          e.state = 'conflict';
          e.versions = Math.max(group.length, 2);
        }
      }
    }
    const lockedRels = new Set(entries.map((e) => e.rel.toLowerCase()));
    for (const { path: p, rel } of files) {
      const d = decide(rel, rulesFor(sets, rel));
      const st = await fsp.stat(p).catch(() => null);
      if (!st) continue;
      if (d.lock) {
        entries.push({ rel, name: path.basename(rel), state: 'waiting', why: d.why, kind: d.kind, size: st.size, modified: st.mtimeMs, path: p, beside: lockedRels.has(rel.toLowerCase()) || undefined });
      } else {
        entries.push({ rel, name: path.basename(rel), state: 'excluded', why: d.why, kind: d.kind, size: st.size, modified: st.mtimeMs, path: p });
      }
    }
    entries.sort((a, b) => a.rel.localeCompare(b.rel));
    return entries;
  }

  /**
   * The plain files the rules would lock, from the names and sizes alone: no
   * locked file is opened, so no key is needed. For the safety copy's estimate.
   *
   * @returns {Promise<Array<{ rel, path, size, modified }>>}
   */
  async function plainWaiting() {
    const { files, sets } = await walk();
    const out = [];
    for (const { path: p, rel } of files) {
      if (!decide(rel, rulesFor(sets, rel)).lock) continue;
      const st = await fsp.stat(p).catch(() => null);
      if (st) out.push({ rel, path: p, size: st.size, modified: st.mtimeMs });
    }
    return out;
  }

  /** What a file's state is and why, in words, for `why` and the table. */
  async function why(rel) {
    const list = await scan();
    const e = list.find((x) => x.rel.toLowerCase() === rel.toLowerCase());
    if (e) return { state: e.state, why: e.why, kind: e.kind, entry: e };
    const { sets } = await walk();
    const d = decide(rel, rulesFor(sets, rel));
    return { state: d.lock ? 'waiting' : 'excluded', why: d.why, kind: d.kind };
  }

  const entryFor = async (rel) => {
    const list = await scan();
    const e = list.find((x) => x.rel.toLowerCase() === rel.toLowerCase());
    return e || null;
  };

  /** Whether a plain file still holds what was hashed: size, time, then the bytes themselves. */
  async function plainStillIs(p, { size, sha256, mtimeMs }) {
    let st;
    try {
      st = await fsp.stat(p);
    } catch {
      return false;
    }
    if (st.size !== size || (mtimeMs !== undefined && st.mtimeMs !== mtimeMs)) return false;
    const h = await fileHash(p);
    return h.size === size && h.sha256 === sha256;
  }

  /** The locked copy read back, every byte, and its record matched against what was locked. */
  async function readBackMatches(entryPath, sha256, size) {
    const isDir = await fsp.stat(entryPath).then((s) => s.isDirectory());
    if (isDir) {
      const o = await openPieces({ dir: entryPath, folderKey: key });
      const hasher = crypto.createHash('sha256');
      let n = 0;
      for await (const c of o.frames()) {
        hasher.update(c);
        n += c.length;
      }
      return n === size && hasher.digest('hex') === sha256;
    }
    const r = await verifyLocked(entryPath, key);
    return r.size === size && r.sha256 === sha256;
  }

  /** What to do with a plain file once it is safely locked, as the person chose. */
  async function afterLocked(plain, sealed, targetPath, options) {
    const o = opt();
    const verdict = options.after ?? o.afterLock;
    if (verdict === 'none') return 'kept';
    await hooks?.at?.('before-delete');
    if (!(await readBackMatches(targetPath, sealed.sha256, sealed.size))) {
      throw new LockerError('failed', 'The locked copy did not read back as the original, so the plain file was left where it was.');
    }
    if (!(await plainStillIs(plain, { size: sealed.size, sha256: sealed.sha256, mtimeMs: sealed.modified }))) {
      throw new LockerError('busy', 'This file changed after it was locked. The plain file was left where it was, and the new version will be locked next.');
    }
    const removePlain = async () => {
      await wipeFile(plain);
      if (await fsp.access(plain).then(() => true, () => false)) throw new LockerError('busy', 'This file is in use. It is locked, and its plain copy could not be removed yet.');
    };
    if (verdict === 'delete') {
      await removePlain();
      return 'deleted';
    }
    if (verdict === 'keep') {
      if (!o.keepDir) throw new LockerError('choice', 'Choose where to keep the plain files first.');
      const rel = posix(path.relative(root, plain));
      let dest = path.join(o.keepDir, ...rel.split('/'));
      await fsp.mkdir(path.dirname(dest), { recursive: true });
      for (let i = 2; ; i += 1) {
        try {
          await fsp.access(dest);
          const ext = path.extname(rel);
          dest = path.join(o.keepDir, `${rel.slice(0, rel.length - ext.length)} (${i})${ext}`);
        } catch {
          break;
        }
      }
      await fsp.copyFile(plain, dest, fsp.constants.COPYFILE_EXCL);
      const copy = await fileHash(dest);
      if (copy.sha256 !== sealed.sha256) {
        await fsp.rm(dest, { force: true }).catch(() => {});
        throw new LockerError('failed', 'The plain file could not be moved to the folder you chose, so it was left here.');
      }
      await removePlain();
      return 'moved';
    }
    return 'kept';
  }

  /** Earlier versions of a smaller file, oldest first: the files themselves in `.rlocker-old/<id>/`. */
  async function oldVersions(fileId) {
    const dir = path.join(root, OLD_DIR, fileId);
    let names = [];
    try {
      names = await fsp.readdir(dir);
    } catch {
      return [];
    }
    return names.filter((n) => n.endsWith(LOCK_EXT)).sort().map((n) => path.join(dir, n));
  }

  /** The present locked copy kept as a version: a checked copy, before anything replaces it. */
  async function keepAsVersion(entry) {
    const o = opt();
    if (o.keepVersions <= 1) return;
    const dir = path.join(root, OLD_DIR, entry.fileId);
    await fsp.mkdir(dir, { recursive: true });
    const name = `${String(Date.now()).padStart(15, '0')}-${crypto.randomBytes(3).toString('hex')}${LOCK_EXT}`;
    const dest = path.join(dir, name);
    await fsp.copyFile(entry.path, dest, fsp.constants.COPYFILE_EXCL);
    const [a, b] = await Promise.all([fileHash(entry.path), fileHash(dest)]);
    if (a.sha256 !== b.sha256) {
      await fsp.rm(dest, { force: true }).catch(() => {});
      throw new LockerError('failed', 'The earlier version could not be kept, so the file was not changed.');
    }
  }

  async function pruneVersions(fileId) {
    const o = opt();
    const old = await oldVersions(fileId);
    const excess = old.length - Math.max(0, o.keepVersions - 1);
    for (let i = 0; i < excess; i += 1) await fsp.rm(old[i], { force: true }).catch(() => {});
    if (old.length && old.length === excess) await fsp.rmdir(path.join(root, OLD_DIR, fileId)).catch(() => {});
  }

  /** A smaller file's earlier versions, and the copy it is now, made the first versions in its new folder of pieces, oldest first. */
  async function carryVersions(existing, dir) {
    await fsp.mkdir(stageDir, { recursive: true });
    for (const version of [...await oldVersions(existing.fileId), existing.path]) {
      const tmp = path.join(stageDir, workingName());
      try {
        const { meta } = await writePlain({ locked: version, dest: tmp, rootKey: key, overwrite: true });
        if (meta.modified) await fsp.utimes(tmp, new Date(meta.modified), new Date(meta.modified)).catch(() => {});
        await sealPieces({ source: tmp, dir, folderKey: key, folderId, stageDir, meta: { name: meta.name, by: meta.by }, keep: opt().keepVersions, hooks });
      } finally {
        await wipeFile(tmp);
      }
    }
  }

  /**
   * A plain file sealed into the folder: as a new locked file, or as the
   * next version of the locked file of the same name. Nothing the file was
   * before is touched until the new copy has been written and checked.
   */
  async function sealInto(plain, realRel, existing, meta) {
    const st = await fsp.stat(plain);
    // A file already in pieces stays in pieces though it has shrunk: its earlier versions are in that folder.
    const useLarge = st.size > largeAt || Boolean(existing?.pieces);
    await fsp.mkdir(path.dirname(abs(realRel)), { recursive: true });
    if (useLarge) {
      const dir = existing?.pieces ? existing.path : path.join(path.dirname(abs(realRel)), `${opt().hidden ? randomName() : `${path.basename(realRel)}${LOCK_EXT}`}.d`);
      if (existing && !existing.pieces) await carryVersions(existing, dir);
      const out = await sealPieces({ source: plain, dir, folderKey: key, folderId, stageDir, meta, keep: opt().keepVersions, hooks });
      // A smaller file that grew: its old single-file copy goes, once the pieces are in place and read back.
      if (existing && !existing.pieces) {
        if (await readBackMatches(dir, out.sha256, out.size)) {
          await fsp.rm(existing.path, { force: true }).catch(() => {});
          await fsp.rm(path.join(root, OLD_DIR, existing.fileId), { recursive: true, force: true }).catch(() => {});
        }
      }
      const stAfter = await fsp.stat(plain);
      return { ...out, modified: stAfter.mtimeMs, target: dir };
    }
    let target;
    let headerBuf;
    let previous = null;
    if (existing && !existing.pieces) {
      target = existing.path;
      const fh = await fsp.open(target, 'r');
      try {
        headerBuf = Buffer.alloc(HEADER_SIZE);
        await fh.read(headerBuf, 0, HEADER_SIZE, 0);
      } finally {
        await fh.close();
      }
      await keepAsVersion(existing);
      previous = target;
    } else {
      target = path.join(path.dirname(abs(realRel)), opt().hidden ? randomName() : `${path.basename(realRel)}${LOCK_EXT}`);
      if (await fsp.access(target).then(() => true, () => false)) throw new LockerError('exists', 'A locked file of that name is already there.');
      headerBuf = newHeader({ mode: MODE_FOLDER, folderId });
    }
    const out = await sealFile({ source: plain, target, rootKey: key, headerBuf, meta, stageDir, previous, hooks });
    if (existing) await pruneVersions(existing.fileId);
    return { ...out, target };
  }

  /**
   * One waiting file locked: sealed, read back, and only then the plain file
   * dealt with as the person chose. Answers what happened to the plain file.
   */
  async function lockOne(rel, options = {}) {
    const plain = abs(rel);
    const list = await scan();
    const entry = list.find((x) => x.rel.toLowerCase() === rel.toLowerCase() && x.path === plain);
    if (!entry || entry.state !== 'waiting') throw new LockerError('state', 'This file is not waiting to be locked.');
    const existing = list.find((x) => x.rel.toLowerCase() === rel.toLowerCase() && x.state === 'locked') || null;
    const before = await fsp.stat(plain);
    if (existing) {
      // A plain copy beside its own locked copy (a crash, or an export): the same bytes need no new version.
      const h = await fileHash(plain);
      if (h.sha256 === existing.sha256 && h.size === existing.size) {
        const verdict = await afterLocked(plain, { sha256: existing.sha256, size: existing.size, modified: before.mtimeMs }, existing.path, options);
        return { plain: verdict, target: existing.path, sha256: existing.sha256, size: existing.size, reused: 0, frames: 0 };
      }
    }
    const meta = { name: path.basename(rel), by: 'lock' };
    const sealed = await sealInto(plain, rel, existing, meta);
    const after = await fsp.stat(plain);
    if (after.size !== before.size || after.mtimeMs !== before.mtimeMs) {
      throw new LockerError('busy', 'This file changed while it was being locked; its plain copy was left and it stays waiting.');
    }
    const plainWent = await afterLocked(plain, { sha256: sealed.sha256, size: sealed.size, modified: after.mtimeMs }, sealed.target, options);
    return { plain: plainWent, target: sealed.target, sha256: sealed.sha256, size: sealed.size, reused: sealed.reused, frames: sealed.frames };
  }

  /** Every waiting file, one at a time. Answers what happened to each, and what would not go. A deadline, or the locker closing, leaves the rest as they are. */
  async function lockWaiting({ deadline = Infinity, shouldStop = () => false, ...options } = {}) {
    // Left by a stop part way, at any depth: what they hold may be plain, so each is overwritten before it goes.
    for (const p of (await walk()).stray) await wipeFile(p).catch(() => {});
    const list = await scan();
    const done = [];
    const failed = [];
    for (const e of list.filter((x) => x.state === 'waiting')) {
      if (shouldStop()) {
        failed.push({ rel: e.rel, code: 'stopped', message: 'The locker closed before this file was locked. It stays as it was.' });
        continue;
      }
      if (Date.now() > deadline) {
        failed.push({ rel: e.rel, code: 'time', message: 'There was not time to lock this file.' });
        continue;
      }
      try {
        done.push({ rel: e.rel, ...(await lockOne(e.rel, options)) });
      } catch (err) {
        failed.push({ rel: e.rel, code: err?.code || 'error', message: err?.message || String(err) });
      }
    }
    return { locked: done, failed };
  }

  /**
   * A locked file copied in from elsewhere made this folder's own: sealed afresh
   * under this folder's key and a new id, read back, and only then does the
   * copy it came from go.
   */
  async function adopt(rel, source, oldPath) {
    const st = await fsp.stat(source);
    const large = st.size > largeAt;
    const dirOf = path.dirname(abs(rel));
    const meta = { name: path.basename(rel), by: 'save' };
    await fsp.mkdir(dirOf, { recursive: true });
    const temp = large ? `${randomName()}.d` : randomName();
    let target = path.join(dirOf, temp);
    const sealed = large ? await sealPieces({ source, dir: target, folderKey: key, folderId, stageDir, meta, keep: opt().keepVersions, hooks })
      : await sealFile({ source, target, rootKey: key, headerBuf: newHeader({ mode: MODE_FOLDER, folderId }), meta, stageDir, previous: null, hooks });
    if (!(await readBackMatches(target, sealed.sha256, sealed.size))) throw new LockerError('failed', 'The locked copy did not read back as the original.');
    const wanted = opt().hidden ? target : path.join(dirOf, `${path.basename(rel)}${LOCK_EXT}${large ? '.d' : ''}`);
    const same = path.resolve(wanted).toLowerCase() === path.resolve(oldPath).toLowerCase();
    if (same && large) await fsp.rm(oldPath, { recursive: true, force: true });
    if (wanted !== target) {
      await fsp.rename(target, wanted);
      target = wanted;
    }
    if (!same) await fsp.rm(oldPath, { recursive: true, force: true });
    return { target };
  }

  /** A locked file (or one in pieces) opened for reading: `frames()`, `meta`, `close()`. */
  async function opener(entry, file = null) {
    if (entry.pieces) return openPieces({ dir: entry.path, folderKey: key, file });
    return openLocked(entry.path, key);
  }

  /** A plain copy written elsewhere: nothing about the locked file changes. */
  async function exportCopy(rel, dest, { overwrite = false, allowConflict = false, lockedRel = null, list = null } = {}) {
    const e = lockedRel ? (await scan()).find((x) => x.lockedRel === lockedRel && x.rel.toLowerCase() === rel.toLowerCase()) || null : await entryFor(rel);
    if (!e || e.state === 'waiting' || e.state === 'excluded' || e.state === 'damaged') throw new LockerError('state', 'That file is not locked.');
    if (e.state === 'conflict' && !allowConflict) throw new LockerError('conflict', 'This file has two versions. Compare them first.');
    const o = await opener(e, list);
    try {
      return await writePlainFrom(o, dest, { overwrite });
    } finally {
      await o.close();
    }
  }

  /** The current plain bytes in a working file, for a shared copy or for reading: under stageDir. */
  async function toStage(entry) {
    await fsp.mkdir(stageDir, { recursive: true });
    const tmp = path.join(stageDir, `.rl-${crypto.randomBytes(12).toString('hex')}.part`);
    const o = await opener(entry);
    try {
      await writePlainFrom(o, tmp, { overwrite: true });
    } finally {
      await o.close();
    }
    return tmp;
  }

  /**
   * A new version of a locked file from a plain file elsewhere (a save from
   * Office, an import): the source is not touched.
   */
  async function saveFrom(rel, source) {
    const list = await scan();
    const existing = list.find((x) => x.rel.toLowerCase() === rel.toLowerCase() && (x.state === 'locked' || x.state === 'conflict')) || null;
    if (existing?.state === 'conflict') throw new LockerError('conflict', 'This file has two versions. Compare them first.');
    const meta = { name: path.basename(rel), by: 'save' };
    const sealed = await sealInto(source, rel, existing, meta);
    if (!(await readBackMatches(sealed.target, sealed.sha256, sealed.size))) throw new LockerError('failed', 'The locked copy did not read back as the original.');
    return sealed;
  }

  /**
   * A locked file back to a plain one, for good: the plain file written and
   * checked first, the locked copy and its versions removed only after.
   */
  async function unlockForGood(rel) {
    const e = await entryFor(rel);
    if (e?.state === 'conflict') throw new LockerError('conflict', 'This file has two versions. Compare them first.');
    if (!e || e.state !== 'locked') throw new LockerError('state', 'That file is not locked.');
    const dest = abs(e.rel);
    const o = await opener(e);
    let done;
    try {
      done = await writePlainFrom(o, dest);
    } finally {
      await o.close();
    }
    const check = await fileHash(dest);
    if (check.sha256 !== e.sha256) throw new LockerError('failed', 'The unlocked copy did not match the locked file, so the locked copy was left in place.');
    await fsp.rm(e.path, { recursive: true, force: true });
    if (e.fileId) await fsp.rm(path.join(root, OLD_DIR, e.fileId), { recursive: true, force: true }).catch(() => {});
    return { dest, size: done.size };
  }

  /** A row of the list, by what the window showed: a plain file can sit beside a locked one of the same name. */
  async function rowFor(rel, state) {
    const list = await scan();
    return list.find((x) => x.rel.toLowerCase() === rel.toLowerCase() && (!state || x.state === state)) || null;
  }

  /**
   * A file given a new name in the same place. A locked one is sealed again
   * under its own id, so it keeps its earlier versions, and its name on the
   * disk follows if the folder shows names.
   */
  async function rename(rel, newName, state) {
    const e = await rowFor(rel, state);
    if (!e) throw new LockerError('state', 'That file is not in this folder.');
    if (e.state === 'conflict') throw new LockerError('conflict', 'This file has two versions. Compare them first.');
    if (e.state === 'damaged') throw new LockerError('state', 'That file cannot be read, so it cannot be renamed.');
    const name = checkedName(newName);
    if (name === e.name) return { rel: e.rel };
    const dir = path.posix.dirname(e.rel);
    const newRel = dir === '.' ? name : `${dir}/${name}`;
    const sameFile = newRel.toLowerCase() === e.rel.toLowerCase();
    const onDisk = (p) => fsp.access(p).then(() => true, () => false);
    if (!sameFile && ((await entryFor(newRel)) || await onDisk(abs(newRel)))) throw new LockerError('exists', 'A file of that name is already here.');
    if (e.state !== 'locked') {
      await fsp.rename(e.path, abs(newRel));
      return { rel: newRel };
    }
    const diskName = path.basename(e.path);
    const named = diskName === `${e.name}${LOCK_EXT}` || diskName === `${e.name}${LOCK_EXT}.d`;
    const newPath = named ? path.join(path.dirname(e.path), `${name}${LOCK_EXT}${e.pieces ? '.d' : ''}`) : e.path;
    if (!sameFile && newPath !== e.path && await onDisk(newPath)) throw new LockerError('exists', 'A file of that name is already here.');
    const tmp = await toStage(e);
    try {
      if ((await fileHash(tmp)).sha256 !== e.sha256) throw new LockerError('failed', 'The locked file could not be read back, so it was not renamed.');
      if (e.modified) await fsp.utimes(tmp, new Date(e.modified), new Date(e.modified));
      const meta = { name, by: 'rename' };
      if (e.pieces) {
        await sealPieces({ source: tmp, dir: e.path, folderKey: key, folderId, stageDir, meta, keep: opt().keepVersions, hooks });
      } else {
        const fh = await fsp.open(e.path, 'r');
        const headerBuf = Buffer.alloc(HEADER_SIZE);
        try {
          await fh.read(headerBuf, 0, HEADER_SIZE, 0);
        } finally {
          await fh.close();
        }
        await sealFile({ source: tmp, target: e.path, rootKey: key, headerBuf, meta, stageDir, previous: e.path, hooks });
      }
      if (newPath !== e.path) await fsp.rename(e.path, newPath);
      if (!(await readBackMatches(newPath, e.sha256, e.size))) throw new LockerError('failed', 'The renamed file did not read back as the original.');
      return { rel: newRel };
    } finally {
      await wipeFile(tmp);
    }
  }

  /** A file gone for good, with the earlier versions kept of it. */
  async function removeFile(rel, state, { lockedRel = null, list = null } = {}) {
    const e = lockedRel ? (await scan()).find((x) => x.lockedRel === lockedRel && x.rel.toLowerCase() === rel.toLowerCase()) || null : await rowFor(rel, state);
    if (!e) throw new LockerError('state', 'That file is not in this folder.');
    // One of two versions may be removed by naming which; the earlier versions belong to the one that stays.
    if (e.state === 'conflict' && !lockedRel) throw new LockerError('conflict', 'This file has two versions. Compare them first.');
    if (e.conflictNote) {
      // Both versions of a large file live in the one folder: only the one named goes.
      await removeNewestList(e.path, key, list);
      return { removed: e.rel };
    }
    if (e.state === 'waiting' || e.state === 'excluded') await wipeFile(e.path);
    else await fsp.rm(e.path, { recursive: true, force: true });
    if (await fsp.access(e.path).then(() => true, () => false)) throw new LockerError('busy', 'That file is in use, so it could not be deleted.');
    if (e.fileId && e.state !== 'conflict') await fsp.rm(path.join(root, OLD_DIR, e.fileId), { recursive: true, force: true }).catch(() => {});
    return { removed: e.rel };
  }

  /** The kept earlier versions of a file, newest first: a label, and the order to ask for one by. */
  async function versions(rel) {
    const e = await entryFor(rel);
    if (!e) throw new LockerError('state', 'That file is not in this folder.');
    if (e.state === 'conflict') throw new LockerError('conflict', 'This file has two versions. Compare them first.');
    if (e.state !== 'locked') throw new LockerError('state', 'That file is not locked.');
    if (e.pieces) {
      const lists = await listsIn(e.path);
      return lists.map((l, i) => ({ version: i, current: i === 0, label: `#${l.seq}` }));
    }
    const old = await oldVersions(e.fileId);
    const list = [{ version: 0, current: true, label: 'Now' }];
    [...old].reverse().forEach((p, i) => list.push({ version: i + 1, current: false, path: p, label: path.basename(p).split('-')[0] }));
    return list;
  }

  /** An earlier version put back as the newest, the present one kept. */
  async function restore(rel, version) {
    const e = await entryFor(rel);
    if (!e) throw new LockerError('state', 'That file is not in this folder.');
    const list = await versions(rel);
    const v = list.find((x) => x.version === Number(version));
    if (!v || v.current) throw new LockerError('state', 'That version is not kept.');
    const tmp = path.join(stageDir, `.rl-${crypto.randomBytes(12).toString('hex')}.part`);
    await fsp.mkdir(stageDir, { recursive: true });
    try {
      if (e.pieces) {
        const o = await openPieces({ dir: e.path, folderKey: key, version: v.version });
        await writePlainFrom(o, tmp, { overwrite: true });
      } else {
        await writePlain({ locked: v.path, dest: tmp, rootKey: key, overwrite: true });
      }
      const check = await fileHash(tmp);
      return await saveFrom(rel, tmp).then((r) => ({ ...r, restored: check.sha256 }));
    } finally {
      await wipeFile(tmp);
    }
  }

  /** The folder's rules file as it stands, for the editor. */
  async function readRules() {
    try {
      return await fsp.readFile(path.join(root, '.rutbalocker'), 'utf8');
    } catch {
      return '';
    }
  }

  /** The rules file written whole. */
  async function writeRules(text) {
    const tmp = path.join(root, `.rl-${crypto.randomBytes(12).toString('hex')}.part`);
    await fsp.writeFile(tmp, text);
    await fsp.rename(tmp, path.join(root, '.rutbalocker'));
  }

  /** The key this folder's work has been holding, overwritten once that work is over. */
  const forget = () => key.fill(0);

  return { root, scan, plainWaiting, readsBackAs: readBackMatches, why, lockOne, lockWaiting, exportCopy, toStage, saveFrom, unlockForGood, rename, removeFile, adopt, versions, restore, readRules, writeRules, entryFor, opener, forget };
}

/** A shared file's key: from its password, by the scrypt settings its header carries. */
export async function sharedKey(password, header) {
  const { passwordKey } = await import('./keys.js');
  return passwordKey(password, { N: header.scrypt.N, r: header.scrypt.r, p: header.scrypt.p, salt: header.scrypt.salt.toString('hex') });
}

export { MODE_SHARED, newHeader };
