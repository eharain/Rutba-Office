// Removing a folder and everything in it, without ever going through a link
// (D-2026-11-13-37).
//
// Office ships on Electron 44, whose Node 24 follows a Windows junction in a
// recursive fs.rm or fs.rmSync: removing a folder that holds a junction, or a
// junction itself, deletes what the junction points at. A person who has put
// a junction inside a profile, a temp folder or a pack's folder would lose
// whatever it leads to. Node 22, which the unit tests run on, does not, so a
// test of fs.rm proved nothing about the runtime Office ships.
//
// This walks the tree itself. Every entry is looked at with lstat, which
// reports a junction as a link on every Node Office runs on; a link (a
// symbolic link or a junction, to a file or a folder) is unlinked where it
// stands and never entered. Only real folders are walked into. A path that is
// itself a link is unlinked, never followed. Product code removes trees only
// through here (tools/check-recursive-removals.mjs refuses the rest).
//
// What is missing is no fault: as with force: true, a tree already gone, or
// an entry removed by someone else meanwhile, is simply not there to remove.
// Windows' passing locks (a scanner, the indexer, a handle closing) are what
// `maxRetries` waits out, as fs.rm's own option of that name does.

import fs from 'node:fs';
import path from 'node:path';

/** The errors a retry may outlast: Windows holding a file a moment longer, a folder not empty yet, handles running out. */
const PASSING = new Set(['EBUSY', 'EPERM', 'ENOTEMPTY', 'EMFILE', 'ENFILE']);

const missing = (err) => err?.code === 'ENOENT' || err?.code === 'ENOTDIR';

/** A folder the size of a disk is never a tree to remove: a mistake upstream, refused before anything goes. */
function refuseRoot(target) {
  if (typeof target !== 'string' || !target.trim()) throw new TypeError('removeTree: a path is needed');
  const full = path.resolve(target);
  if (full === path.parse(full).root) throw new Error(`removeTree: refusing to remove a root: ${full}`);
  return full;
}

const sleepSync = (ms) => {
  try {
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
  } catch {
    const until = Date.now() + ms;
    while (Date.now() < until) { /* no Atomics: wait the plain way */ }
  }
};
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/* ── at once ─────────────────────────────────────────────────────────────── */

/** A link or a file, removed where it stands. A link to a folder (a junction, a folder symlink) goes as itself. */
function unlinkSync(p, st) {
  try {
    fs.unlinkSync(p);
  } catch (err) {
    if (missing(err)) return;
    if (st.isSymbolicLink() && (err.code === 'EPERM' || err.code === 'EISDIR')) {
      // A folder link on a system where unlink will not take one: rmdir, never recursive, removes the link alone.
      fs.rmdirSync(p);
      return;
    }
    if (err.code === 'EPERM' || err.code === 'EACCES') {
      // A read-only file on Windows: made writable, then removed, as fs.rm does.
      try {
        fs.chmodSync(p, 0o666);
      } catch {
        throw err;
      }
      fs.unlinkSync(p);
      return;
    }
    throw err;
  }
}

function removeOnceSync(p) {
  let st;
  try {
    st = fs.lstatSync(p);
  } catch (err) {
    if (missing(err)) return;
    throw err;
  }
  if (!st.isDirectory() || st.isSymbolicLink()) {
    unlinkSync(p, st);
    return;
  }
  let names;
  try {
    names = fs.readdirSync(p);
  } catch (err) {
    if (missing(err)) return;
    throw err;
  }
  for (const name of names) removeOnceSync(path.join(p, name));
  try {
    fs.rmdirSync(p);
  } catch (err) {
    if (!missing(err)) throw err;
  }
}

/**
 * Remove a file, a link or a folder and everything in it, synchronously,
 * entering no link. Missing is no fault.
 * @param {string} target
 * @param {{ maxRetries?: number, retryDelay?: number }} [o]  as fs.rm's: tries again on Windows' passing locks, waiting longer each time
 */
export function removeTreeSync(target, { maxRetries = 0, retryDelay = 100 } = {}) {
  const full = refuseRoot(target);
  for (let attempt = 0; ; attempt++) {
    try {
      removeOnceSync(full);
      return;
    } catch (err) {
      if (attempt >= maxRetries || !PASSING.has(err?.code)) throw err;
      sleepSync(retryDelay * (attempt + 1));
    }
  }
}

/* ── in the background ───────────────────────────────────────────────────── */

async function unlinkAsync(p, st) {
  try {
    await fs.promises.unlink(p);
  } catch (err) {
    if (missing(err)) return;
    if (st.isSymbolicLink() && (err.code === 'EPERM' || err.code === 'EISDIR')) {
      await fs.promises.rmdir(p);
      return;
    }
    if (err.code === 'EPERM' || err.code === 'EACCES') {
      try {
        await fs.promises.chmod(p, 0o666);
      } catch {
        throw err;
      }
      await fs.promises.unlink(p);
      return;
    }
    throw err;
  }
}

/** How many entries of one folder are removed at a time. */
const AT_ONCE = 8;

async function removeOnce(p) {
  let st;
  try {
    st = await fs.promises.lstat(p);
  } catch (err) {
    if (missing(err)) return;
    throw err;
  }
  if (!st.isDirectory() || st.isSymbolicLink()) {
    await unlinkAsync(p, st);
    return;
  }
  let names;
  try {
    names = await fs.promises.readdir(p);
  } catch (err) {
    if (missing(err)) return;
    throw err;
  }
  for (let i = 0; i < names.length; i += AT_ONCE) {
    await Promise.all(names.slice(i, i + AT_ONCE).map((name) => removeOnce(path.join(p, name))));
  }
  try {
    await fs.promises.rmdir(p);
  } catch (err) {
    if (!missing(err)) throw err;
  }
}

/**
 * Remove a file, a link or a folder and everything in it, entering no link.
 * Missing is no fault.
 * @param {string} target
 * @param {{ maxRetries?: number, retryDelay?: number }} [o]  as fs.rm's
 * @returns {Promise<void>}
 */
export async function removeTree(target, { maxRetries = 0, retryDelay = 100 } = {}) {
  const full = refuseRoot(target);
  for (let attempt = 0; ; attempt++) {
    try {
      await removeOnce(full);
      return;
    } catch (err) {
      if (attempt >= maxRetries || !PASSING.has(err?.code)) throw err;
      await sleep(retryDelay * (attempt + 1));
    }
  }
}
