// Where an opened locked file lives while a program has it: Office's own
// profile folder, `<profile>/locker/open/<owner>-<random>/Name.ext`. The
// profile is the person's own on every system, so the copy is private to them
// without Office rewriting any permissions. Written back through the locked
// file's own save, and removed when the program lets go, when the locker locks,
// and at the next unlock after a crash.
//
// 2.24.0 and the first 2.25.0 builds put the copy in a hidden folder beside the
// file instead (`.rutba-open`, OPEN_DIR). Those folders are only ever cleared
// now: findWorkingDirs finds them at the next unlock.

import crypto from 'node:crypto';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { removeTreeAsync } from '@rutba/office-shell/electron/remove-tree';
import { wipeFile, OLD_DIR } from './folder.js';
import { LockerError } from './format.js';
import { OPEN_DIR } from './rules.js';

/** Room kept free on the profile's drive beyond the file itself, for the program's own saves. */
const SPARE_BYTES = 64 * 1024 * 1024;

/**
 * Whether the profile's drive has room for a plain copy of `size` bytes.
 * Refused with a sentence that says how much is needed; a drive that cannot
 * say how much it has is left to the write itself.
 */
export async function roomFor(dir, size, statfs = fsp.statfs) {
  let free = null;
  try {
    const s = await statfs(dir);
    free = Number(s.bavail) * Number(s.bsize);
  } catch {
    return;
  }
  if (!Number.isFinite(free) || free >= size + SPARE_BYTES) return;
  const mb = (n) => `${Math.max(1, Math.ceil(n / (1024 * 1024)))} MB`;
  throw new LockerError('space', `There is not enough free space on the drive that holds your Office profile to open this file. It needs about ${mb(size + SPARE_BYTES)}, and ${mb(free)} is free. Free some space there, then open it again.`);
}

/**
 * A new empty working folder inside `top` (the profile's open folder), named
 * for the locker that made it, readable by this account alone where the
 * system keeps such modes (0700); on Windows the profile's own permissions
 * already keep it the person's.
 */
export async function makeWorkingDir(top, owner = '') {
  await fsp.mkdir(top, { recursive: true, mode: 0o700 });
  // A link or junction in its place is refused: nothing goes where it points.
  if (!(await fsp.lstat(top)).isDirectory()) {
    throw new LockerError('restrict', 'The locker could not make its working folder in your Office profile, so the file was not opened. Something other than its own folder is in the way.');
  }
  if (process.platform !== 'win32') await fsp.chmod(top, 0o700).catch(() => {});
  const sub = path.join(top, `${owner ? `${owner}-` : ''}${crypto.randomBytes(8).toString('hex')}`);
  await fsp.mkdir(sub, { mode: 0o700 });
  return sub;
}

/** Whether a program has this file open. Office leaves a `~$` file beside it; most others refuse to share it for writing. */
export async function isHeld(file) {
  const names = await fsp.readdir(path.dirname(file)).catch(() => []);
  if (names.some((n) => n.startsWith('~$'))) return true;
  try {
    const fh = await fsp.open(file, 'r+');
    await fh.close();
    return false;
  } catch (err) {
    return err.code === 'EBUSY' || err.code === 'EPERM' || err.code === 'EACCES';
  }
}

async function filesIn(dir) {
  const out = [];
  const entries = await fsp.readdir(dir, { withFileTypes: true }).catch(() => []);
  for (const e of entries) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) out.push(...await filesIn(p));
    else out.push(p);
  }
  return out;
}

const gone = (p) => fsp.stat(p).then(() => false, (err) => err.code === 'ENOENT');

/** One copy overwritten and removed, its folder left for the window that may still save into it. Answers whether it has really gone. */
export async function wipeCopy(file) {
  await wipeFile(file).catch(() => {});
  return gone(file);
}

/** A working folder and everything a program put in it, overwritten and removed. Answers whether it has really gone. */
export async function wipeWorking(sub) {
  for (const f of await filesIn(sub)) await wipeFile(f);
  await removeTreeAsync(sub).catch(() => {});
  const removed = await gone(sub);
  if (removed && path.basename(path.dirname(sub)) === OPEN_DIR) await fsp.rmdir(path.dirname(sub)).catch(() => {});
  return removed;
}

/** The working folders under a locked folder, however they came to be there. */
export async function findWorkingDirs(root) {
  const out = [];
  const go = async (dir) => {
    const entries = await fsp.readdir(dir, { withFileTypes: true }).catch(() => []);
    for (const e of entries) {
      if (!e.isDirectory() || e.name === OLD_DIR || e.name.endsWith('.rlock.d')) continue;
      const p = path.join(dir, e.name);
      if (e.name !== OPEN_DIR) {
        await go(p);
        continue;
      }
      const subs = await fsp.readdir(p).catch(() => []);
      for (const s of subs) out.push(path.join(p, s));
      if (!subs.length) await fsp.rmdir(p).catch(() => {});
    }
  };
  await go(root);
  return out;
}
