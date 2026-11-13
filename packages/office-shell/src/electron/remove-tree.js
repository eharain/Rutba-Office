// A folder removed with everything in it, never through a link.
//
// From Node 24 (Electron 44), fs.rm with `recursive` follows a Windows
// junction: one inside the folder, and one that is the path itself, so what
// it points at is deleted with it. Node 22 did not. This walks the tree with
// lstat instead: a link (a symbolic link or a junction, which lstat names
// one) is unlinked where it is, and never entered. A path that is not there
// is not an error, as with rm's `force`.
//
// Kept apart so that every part of Office that clears a folder can use the
// same one (`@rutba/office-shell/electron/remove-tree`).

import { createRequire } from 'node:module';
import path from 'node:path';

const require = createRequire(import.meta.url);
const realFs = process.versions.electron ? require('original-fs') : require('node:fs');

const gone = (err) => err?.code === 'ENOENT';
/** Errors worth a moment's wait and another try: a file a scanner or an indexer has open for a moment. */
const passing = (err) => ['EBUSY', 'EPERM', 'ENOTEMPTY', 'EACCES'].includes(err?.code);

/** A link removed itself: unlink takes a junction or a symbolic link to a folder on Windows as well. */
function unlinkLink(fs, p) {
  try {
    fs.unlinkSync(p);
  } catch (err) {
    if (gone(err)) return;
    if (err?.code !== 'EPERM' && err?.code !== 'EISDIR') throw err;
    fs.rmdirSync(p);
  }
}

function removeOnce(fs, p) {
  let st;
  try {
    st = fs.lstatSync(p);
  } catch (err) {
    if (gone(err)) return;
    throw err;
  }
  if (st.isSymbolicLink()) return unlinkLink(fs, p);
  if (!st.isDirectory()) {
    try {
      fs.unlinkSync(p);
    } catch (err) {
      if (gone(err)) return;
      // A file marked read-only: made writable, then removed.
      if (err?.code !== 'EPERM') throw err;
      fs.chmodSync(p, 0o666);
      fs.unlinkSync(p);
    }
    return;
  }
  for (const name of fs.readdirSync(p)) removeOnce(fs, path.join(p, name));
  try {
    fs.rmdirSync(p);
  } catch (err) {
    if (!gone(err)) throw err;
  }
}

const pauseSync = (ms) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);

/**
 * Removes `p` and everything in it, unlinking links rather than entering them.
 * @param {string} p
 * @param {{ retries?: number, delayMs?: number, fs?: typeof import('node:fs') }} [o]
 *   `retries`: tries again after a busy file this many times, `delayMs` apart
 */
export function removeTree(p, { retries = 0, delayMs = 100, fs = realFs } = {}) {
  for (let i = 0; ; i += 1) {
    try {
      return removeOnce(fs, p);
    } catch (err) {
      if (i >= retries || !passing(err)) throw err;
      pauseSync(delayMs);
    }
  }
}

async function removeOnceAsync(fsp, p) {
  let st;
  try {
    st = await fsp.lstat(p);
  } catch (err) {
    if (gone(err)) return;
    throw err;
  }
  if (st.isSymbolicLink()) {
    try {
      await fsp.unlink(p);
    } catch (err) {
      if (gone(err)) return;
      if (err?.code !== 'EPERM' && err?.code !== 'EISDIR') throw err;
      await fsp.rmdir(p);
    }
    return;
  }
  if (!st.isDirectory()) {
    try {
      await fsp.unlink(p);
    } catch (err) {
      if (gone(err)) return;
      if (err?.code !== 'EPERM') throw err;
      await fsp.chmod(p, 0o666);
      await fsp.unlink(p);
    }
    return;
  }
  for (const name of await fsp.readdir(p)) await removeOnceAsync(fsp, path.join(p, name));
  try {
    await fsp.rmdir(p);
  } catch (err) {
    if (!gone(err)) throw err;
  }
}

/**
 * removeTree, without holding the thread: for a large folder, or one a file
 * of which may still be closing.
 * @param {string} p
 * @param {{ retries?: number, delayMs?: number, fs?: typeof import('node:fs') }} [o]
 */
export async function removeTreeAsync(p, { retries = 0, delayMs = 100, fs = realFs } = {}) {
  for (let i = 0; ; i += 1) {
    try {
      return await removeOnceAsync(fs.promises, p);
    } catch (err) {
      if (i >= retries || !passing(err)) throw err;
      await new Promise((r) => setTimeout(r, delayMs));
    }
  }
}
