// Where an opened locked file lives while a program has it: a hidden folder
// beside the file, `.rutba-open/<random>/Name.ext`, so any program can open it
// and it sits on the same drive. Written back through the locked file's own
// save, and removed when the program lets go, when the locker locks, and at the
// next unlock after a crash.

import { execFile } from 'node:child_process';
import crypto from 'node:crypto';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { wipeFile, OLD_DIR } from './folder.js';
import { OPEN_DIR } from './rules.js';

const run = (cmd, args) => new Promise((resolve) => {
  execFile(cmd, args, { windowsHide: true, timeout: 10000 }, (err) => resolve(!err));
});

/** The folder hidden, left out of Windows search, and readable by this person only. A drive with no permissions (a memory stick) simply has none. */
async function restrict(dir) {
  if (process.platform !== 'win32') {
    await fsp.chmod(dir, 0o700).catch(() => {});
    return;
  }
  await run('attrib', ['+h', '+i', dir]);
  const user = process.env.USERNAME;
  if (!user) return;
  const who = process.env.USERDOMAIN ? `${process.env.USERDOMAIN}\\${user}` : user;
  await run('icacls', [dir, '/inheritance:r', '/grant:r', `${who}:(OI)(CI)F`]);
}

/** A new empty working folder inside `top`. `top` is restricted before anything goes in it. */
export async function makeWorkingDir(top) {
  let made = false;
  try {
    await fsp.mkdir(top);
    made = true;
  } catch (err) {
    if (err.code === 'ENOENT') {
      await fsp.mkdir(top, { recursive: true });
      made = true;
    } else if (err.code !== 'EEXIST') {
      throw err;
    }
  }
  if (made) await restrict(top);
  const sub = path.join(top, crypto.randomBytes(8).toString('hex'));
  await fsp.mkdir(sub);
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
  await fsp.rm(sub, { recursive: true, force: true }).catch(() => {});
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
