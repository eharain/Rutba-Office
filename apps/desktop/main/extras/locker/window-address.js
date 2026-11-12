// What an Office window's address says, read as the address was written.
// windows.js builds it with URLSearchParams, which writes a space as "+";
// decodeURIComponent leaves the "+", so a path with a space in it, a file's
// own name or a user folder such as C:\Users\Sam Khan, never matched a test
// on the decoded text, and the Locker took a document's copy from under it.

import path from 'node:path';

/** One parameter of a window's address, or null. */
export function addressParam(url, key) {
  try {
    return new URL(url).searchParams.get(key);
  } catch {
    return null;
  }
}

/** Whether two paths name the same file, as windows.js findByFile compares them; case does not count on Windows. */
export function sameFile(a, b, platform = process.platform) {
  if (!a || !b) return false;
  const x = path.resolve(a);
  const y = path.resolve(b);
  return platform === 'win32' ? x.toLowerCase() === y.toLowerCase() : x === y;
}

/** Whether the window at this address shows this file. */
export function addressShows(url, file, platform = process.platform) {
  return sameFile(addressParam(url, 'file'), file, platform);
}

/** Whether the window at this address is the given app's. */
export function addressApp(url) {
  return addressParam(url, 'app');
}
