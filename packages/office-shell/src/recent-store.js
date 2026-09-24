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
