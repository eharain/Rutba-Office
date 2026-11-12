// Where an opened locked file lives while a program has it: a hidden folder
// beside the file, `.rutba-open/<random>/Name.ext`, so any program can open it
// and it sits on the same drive. Written back through the locked file's own
// save, and removed when the program lets go, when the locker locks, and at the
// next unlock after a crash.

import { execFile } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { wipeFile, OLD_DIR } from './folder.js';
import { LockerError } from './format.js';
import { OPEN_DIR } from './rules.js';

/** A Windows program by its full path, so a program of the same name earlier on the PATH is never run. */
const system32 = (name) => path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', name);

/** Runs a program and answers whether it succeeded, with what it printed. */
const exec = (cmd, args, env = null) => new Promise((resolve) => {
  execFile(cmd, args, { windowsHide: true, timeout: 30000, ...(env ? { env: { ...process.env, ...env } } : {}) }, (err, stdout) => resolve({ ok: !err, out: String(stdout || '') }));
});

let sidHere = null;
/** The security ID of the account this program runs as: Windows' own name for it, which needs no lookup by name. */
async function currentSid(sys) {
  if (sys.sid) return sys.sid();
  if (!sidHere) {
    const r = await sys.exec(system32('whoami.exe'), ['/user', '/fo', 'csv', '/nh']);
    const m = r.ok ? r.out.match(/"(S-1-[0-9-]+)"/) : null;
    if (!m) return null;
    sidHere = m[1];
  }
  return sidHere;
}

/**
 * Whether a folder's permissions let in this account alone: none taken from
 * the folder above, and every entry for this account. Read as Windows writes
 * permissions down (SDDL), which no language setting changes.
 */
export function onlyFor(sddl, sid) {
  // The list must run to the end of the line, or to the audit part: one cut short by an entry not read here never passes.
  const dacl = /D:([A-Z]*)((?:\([^)]*\))*)(?=$|S:)/.exec(String(sddl || '').trim());
  if (!dacl || !dacl[1].includes('P')) return false;
  const aces = [...dacl[2].matchAll(/\(([^)]*)\)/g)].map((m) => m[1].split(';'));
  // Windows writes the computer's own Administrator account (number 500) as LA, not by its number.
  const mine = (who) => who === sid || (who === 'LA' && /^S-1-5-21-\d+-\d+-\d+-500$/.test(sid));
  return aces.length > 0 && aces.every(([type, flags, , , , who]) => type === 'A' && !/ID/.test(flags) && mine(who));
}

/** The folder's permissions as SDDL, or null if they could not be read. */
async function readSddl(dir, sys) {
  const out = path.join(os.tmpdir(), `.rl-acl-${crypto.randomBytes(8).toString('hex')}.txt`);
  try {
    const r = await sys.exec(system32('icacls.exe'), [dir, '/save', out]);
    if (!r.ok) return null;
    // icacls writes the file in UTF-16: the folder's name on one line, its permissions on the next.
    return fs.readFileSync(out).toString('utf16le').split(/\r?\n/)[1] || null;
  } catch {
    return null;
  } finally {
    try { fs.rmSync(out, { force: true }); } catch { /* held by a scanner: only a name and permissions, in this account's own temp */ }
  }
}

/** The drive's file system (NTFS, FAT32, exFAT...), or null if Windows could not say, as for a network share. */
async function driveFormat(dir) {
  // The drive goes in by the environment, never inside the command.
  const r = await exec(system32(path.join('WindowsPowerShell', 'v1.0', 'powershell.exe')), ['-NoProfile', '-NonInteractive', '-Command', '[Console]::Out.Write(([IO.DriveInfo]::new($env:RL_ROOT)).DriveFormat)'], { RL_ROOT: path.parse(path.resolve(dir)).root });
  return r.ok ? r.out.trim() || null : null;
}

/** File systems that keep no permissions at all: a folder on them is open to anyone who has the drive, as every file on it is. */
const NO_PERMISSIONS = new Set(['FAT', 'FAT12', 'FAT16', 'FAT32', 'exFAT']);

/** What restrict() asks of Windows; a test hands in its own. */
export const SYSTEM = { exec, driveFormat };

/**
 * The folder hidden, left out of Windows search, and open to this account
 * only, granted by its security ID and read back to be sure. If that cannot be
 * done on a drive that keeps permissions, nothing is put in it: the open is
 * refused, rather than a plain copy left where others could read it. A drive
 * with no permissions at all (a FAT or exFAT memory stick) has none to set.
 */
export async function restrict(dir, sys = SYSTEM) {
  if (process.platform !== 'win32' && !sys.windows) {
    await fsp.chmod(dir, 0o700).catch(() => {});
    return;
  }
  await sys.exec(system32('attrib.exe'), ['+h', '+i', dir]);
  const sid = await currentSid(sys);
  if (sid) {
    // Already this account's alone (a second file opened beside the first): left as it is.
    if (onlyFor(await readSddl(dir, sys), sid)) return;
    const set = await sys.exec(system32('icacls.exe'), [dir, '/inheritance:r', '/grant:r', `*${sid}:(OI)(CI)F`]);
    if (set.ok && onlyFor(await readSddl(dir, sys), sid)) return;
  }
  const format = await sys.driveFormat(dir);
  if (format && NO_PERMISSIONS.has(format)) return;
  throw new LockerError('restrict', 'The locker could not make a private folder for this file here, so it was not opened. Windows would not limit the folder to your account. Try again, or open it from a folder on another drive.');
}

/**
 * A new empty working folder inside `top`, named for the locker that made it.
 * `top` is restricted to this account before anything goes in it, each time,
 * so a folder left by an earlier version or an earlier failure is made private
 * too. One made here and refused is taken away again.
 */
export async function makeWorkingDir(top, owner = '', sys = SYSTEM) {
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
  // A link or junction in its place is refused: its target's permissions are never rewritten, and nothing goes in it.
  if (!made && !(await fsp.lstat(top)).isDirectory()) {
    throw new LockerError('restrict', 'The locker could not make a private folder for this file here, so it was not opened. Something other than its own folder is in the way.');
  }
  try {
    await restrict(top, sys);
  } catch (err) {
    if (made) await fsp.rmdir(top).catch(() => {});
    throw err;
  }
  const sub = path.join(top, `${owner ? `${owner}-` : ''}${crypto.randomBytes(8).toString('hex')}`);
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
