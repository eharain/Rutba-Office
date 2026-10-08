// The recent-files list, as list operations with nothing Electron about
// them — so a rename or a removal can be checked without a window.
//
// The Electron-backed store (electron/store.js) does the reading, the
// writing to disk and the file-system side of a rename; this file only says
// what happens to the array of entries.

import path from 'node:path';

/** Take one entry out of the list by path. An unknown path is a no-op. */
export function removeEntry(list, p) {
  return list.filter((r) => r.path !== p);
}

/**
 * Point the entry at `p` to a new path, and rename it to match — the store
 * only knows the file moved, not why. An unknown path is a no-op.
 */
export function renameEntry(list, p, to) {
  return list.map((r) => (r.path === p ? { ...r, path: to, name: path.basename(to) } : r));
}

/**
 * The entries whose files are still there, up to `limit`, looked for all at
 * once and without waiting on any one for longer than `timeoutMs`: a file on a
 * network drive that has gone away held the window's Recent list — and the
 * main process — for as long as the system took to give up on each one. A
 * file that does not answer in time is listed; opening it says if it is gone.
 */
export async function presentEntries(list, { exists, limit = 30, timeoutMs = 1500 }) {
  // Each file's wait ends when it answers or when its time is up, whichever
  // comes first, and its timer goes with it.
  const look = (p) => new Promise((resolve) => {
    const timer = setTimeout(() => resolve(true), timeoutMs);
    Promise.resolve()
      .then(() => exists(p))
      .then((there) => { clearTimeout(timer); resolve(Boolean(there)); }, () => { clearTimeout(timer); resolve(false); });
  });
  const there = await Promise.all(list.map((r) => look(r.path)));
  return list.filter((_, i) => there[i]).slice(0, limit);
}
