// The network places a window may draw from over rutba://file and
// rutba://thumb.
//
// A local file is served as asked: a window can read one through the bridge
// already. A path on a network share is another matter. Only looking at it
// hands the share the person's Windows sign-in, so a share that a document
// or a page named could take it without a click. Such a path is served only
// when the person reached it in this run: a file opened from there, a file
// or folder picked in a dialog, a folder browsed.

import path from 'node:path';

/**
 * Whether a path goes through the network: `\\server\share`, `//server/share`,
 * and the `\\?\` and `\\.\` forms, on Windows. Elsewhere a path that starts
 * with two slashes is a local one.
 */
export function isNetworkPath(p, platform = process.platform) {
  return platform === 'win32' && /^(\\\\|\/\/)/.test(String(p || ''));
}

const granted = new Set();

const keyOf = (p) => path.win32.normalize(String(p)).replace(/[\\/]+$/, '').toLowerCase();

/** The person reached this place themselves: it, and everything in it, may be drawn from. */
export function grantPlace(p, platform = process.platform) {
  if (p && isNetworkPath(p, platform)) granted.add(keyOf(p));
}

/** Whether rutba://file and rutba://thumb may serve a path. */
export function mayServe(p, platform = process.platform) {
  if (!isNetworkPath(p, platform)) return true;
  let at = keyOf(p);
  for (;;) {
    if (granted.has(at)) return true;
    const up = keyOf(path.win32.dirname(at));
    if (up === at || !up) return false;
    at = up;
  }
}

/** For tests: forget every grant. */
export function forgetPlaces() {
  granted.clear();
}
