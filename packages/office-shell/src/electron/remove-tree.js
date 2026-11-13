// A folder removed with everything in it, never through a link.
//
// From Node 24 (Electron 44), fs.rm with `recursive` follows a Windows
// junction: one inside the folder, and one that is the path itself, so what
// it points at is deleted with it. Node 22 did not. This walks the tree with
// lstat instead: a link (a symbolic link or a junction, which lstat names
// one) is unlinked where it is, and never entered. A path that is not there,
// or that goes while the walk is under way (another removal of the same
// tree), is not an error, as with rm's `force`. Removals of the same tree
// running at once need `retries`: on Windows a file one of them is deleting
// answers EPERM to the others until it has gone. A filesystem root, a
// \\?\GLOBALROOT or \\.\GLOBALROOT path, or no path at all, is refused with
// nothing touched.
//
// Kept apart so that every part of Office that clears a folder can use the
// same one (`@rutba/office-shell/electron/remove-tree`).

import { createRequire } from 'node:module';
import path from 'node:path';

const require = createRequire(import.meta.url);
const realFs = process.versions.electron ? require('original-fs') : require('node:fs');

const windows = process.platform === 'win32';
/** Already removed: not there, or (off Windows) a folder on the way that is now a file. */
const gone = (err) => err?.code === 'ENOENT' || (!windows && err?.code === 'ENOTDIR');
/** Errors worth a moment's wait and another try: a file a scanner or an indexer has open for a moment. */
const passing = (err) => ['EBUSY', 'EPERM', 'ENOTEMPTY', 'EACCES'].includes(err?.code);
/** \\?\UNC\server\share, which path.parse does not take for a root. */
const LONG_UNC_ROOT = /^\\\\[?.]\\UNC\\[^\\]+(\\[^\\]*)?\\*$/i;
/** \\?\GLOBALROOT and \\.\GLOBALROOT: the object manager's namespace, devices and all, never a folder of ours. */
const GLOBALROOT = /^[\\/]{2}[?.][\\/]GLOBALROOT([\\/]|$)/i;

/**
 * The path to remove, resolved ("D:" is that drive's working folder), or an
 * error for no path, a GLOBALROOT path, or a filesystem root: C:\,
 * \\server\share\, \\?\C:\, \\?\UNC\server\share\ or /.
 */
function treeToRemove(p) {
  if (typeof p !== 'string' || p === '') throw new TypeError(`removeTree: no path to remove (${p === '' ? 'an empty path' : typeof p})`);
  if (GLOBALROOT.test(p)) throw new Error(`removeTree: ${p} is in GLOBALROOT, not a folder to remove`);
  const full = path.resolve(p);
  if (GLOBALROOT.test(full)) throw new Error(`removeTree: ${full} is in GLOBALROOT, not a folder to remove`);
  if (path.dirname(full) === full || LONG_UNC_ROOT.test(full)) throw new Error(`removeTree: ${full} is a filesystem root, not a folder to remove`);
  return full;
}

/** `retries` and `delayMs`, each a finite number of 0 or more: a NaN wait would hold Atomics.wait for good. */
function checkOptions(retries, delayMs) {
  for (const [name, v] of [['retries', retries], ['delayMs', delayMs]]) {
    if (typeof v !== 'number' || !Number.isFinite(v) || v < 0) throw new TypeError(`removeTree: ${name} must be a finite number of 0 or more, not ${String(v)}`);
  }
}

/**
 * What a failed step on `p` means: nothing, when `p` has gone, and the error
 * otherwise. On Windows a file or folder another removal is already deleting
 * answers EPERM until it has gone, so EPERM there is checked with lstat.
 */
function settle(fs, p, err) {
  if (gone(err)) return;
  if (windows && err?.code === 'EPERM') {
    try {
      fs.lstatSync(p);
    } catch (e) {
      if (gone(e)) return;
    }
  }
  throw err;
}

async function settleAsync(fsp, p, err) {
  if (gone(err)) return;
  if (windows && err?.code === 'EPERM') {
    try {
      await fsp.lstat(p);
    } catch (e) {
      if (gone(e)) return;
    }
  }
  throw err;
}

function removeOnce(fs, p) {
  let st;
  try {
    st = fs.lstatSync(p);
  } catch (err) {
    return settle(fs, p, err);
  }
  if (st.isSymbolicLink()) {
    // unlink takes a junction or a symbolic link to a folder on Windows as well; rmdir where it will not.
    // Elsewhere rmdir never removes a link (it answers ENOTDIR, which would read as gone).
    try {
      fs.unlinkSync(p);
    } catch (err) {
      if (!windows || (err?.code !== 'EPERM' && err?.code !== 'EISDIR')) return settle(fs, p, err);
      try {
        fs.rmdirSync(p);
      } catch (again) {
        settle(fs, p, again);
      }
    }
    return;
  }
  if (!st.isDirectory()) {
    try {
      fs.unlinkSync(p);
    } catch (err) {
      // A file marked read-only, off Windows: made writable, then removed. On
      // Windows unlink clears read-only itself, so EPERM there is something else.
      if (windows || err?.code !== 'EPERM') return settle(fs, p, err);
      try {
        fs.chmodSync(p, 0o666);
        fs.unlinkSync(p);
      } catch (again) {
        settle(fs, p, again);
      }
    }
    return;
  }
  // A read-only folder: on Windows its files can go as it is; elsewhere it is made writable first.
  if (!windows && !(st.mode & 0o200)) {
    try {
      fs.chmodSync(p, st.mode | 0o700);
    } catch (err) {
      return settle(fs, p, err);
    }
  }
  let names;
  try {
    names = fs.readdirSync(p);
  } catch (err) {
    return settle(fs, p, err);
  }
  for (const name of names) removeOnce(fs, path.join(p, name));
  try {
    fs.rmdirSync(p);
  } catch (err) {
    settle(fs, p, err);
  }
}

/** Holds the thread for `ms`. Where it may not be held, the busy error stands. */
function pauseSync(ms, busy) {
  try {
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
  } catch {
    throw busy;
  }
}

/**
 * Removes `p` and everything in it, unlinking links rather than entering them.
 * Throws, with nothing touched, for no path, a filesystem root, or `retries`
 * or `delayMs` that is not a finite number of 0 or more.
 * @param {string} p
 * @param {{ retries?: number, delayMs?: number, fs?: typeof import('node:fs') }} [o]
 *   `retries`: tries again after a busy file this many times, waiting
 *   `delayMs * n` before the n-th try again, as rm's retryDelay does. The
 *   wait holds the thread (Atomics.wait); where the thread may not be held,
 *   use removeTreeAsync, as this then throws the busy error without waiting.
 */
export function removeTree(p, { retries = 0, delayMs = 100, fs = realFs } = {}) {
  const full = treeToRemove(p);
  checkOptions(retries, delayMs);
  for (let i = 0; ; i += 1) {
    try {
      return removeOnce(fs, full);
    } catch (err) {
      if (i >= retries || !passing(err)) throw err;
      pauseSync(delayMs * (i + 1), err);
    }
  }
}

async function removeOnceAsync(fsp, p) {
  let st;
  try {
    st = await fsp.lstat(p);
  } catch (err) {
    return settleAsync(fsp, p, err);
  }
  if (st.isSymbolicLink()) {
    try {
      await fsp.unlink(p);
    } catch (err) {
      if (!windows || (err?.code !== 'EPERM' && err?.code !== 'EISDIR')) return settleAsync(fsp, p, err);
      try {
        await fsp.rmdir(p);
      } catch (again) {
        await settleAsync(fsp, p, again);
      }
    }
    return;
  }
  if (!st.isDirectory()) {
    try {
      await fsp.unlink(p);
    } catch (err) {
      if (windows || err?.code !== 'EPERM') return settleAsync(fsp, p, err);
      try {
        await fsp.chmod(p, 0o666);
        await fsp.unlink(p);
      } catch (again) {
        await settleAsync(fsp, p, again);
      }
    }
    return;
  }
  if (!windows && !(st.mode & 0o200)) {
    try {
      await fsp.chmod(p, st.mode | 0o700);
    } catch (err) {
      return settleAsync(fsp, p, err);
    }
  }
  let names;
  try {
    names = await fsp.readdir(p);
  } catch (err) {
    return settleAsync(fsp, p, err);
  }
  for (const name of names) await removeOnceAsync(fsp, path.join(p, name));
  try {
    await fsp.rmdir(p);
  } catch (err) {
    await settleAsync(fsp, p, err);
  }
}

/**
 * removeTree, without holding the thread: for a large folder, one a file of
 * which may still be closing, or wherever the thread may not be held.
 * @param {string} p
 * @param {{ retries?: number, delayMs?: number, fs?: typeof import('node:fs') }} [o]
 *   `retries` and `delayMs` as removeTree's: `delayMs * n` before the n-th try again
 */
export async function removeTreeAsync(p, { retries = 0, delayMs = 100, fs = realFs } = {}) {
  const full = treeToRemove(p);
  checkOptions(retries, delayMs);
  for (let i = 0; ; i += 1) {
    try {
      return await removeOnceAsync(fs.promises, full);
    } catch (err) {
      if (i >= retries || !passing(err)) throw err;
      await new Promise((r) => setTimeout(r, delayMs * (i + 1)));
    }
  }
}
