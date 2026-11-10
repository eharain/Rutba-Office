// How a locked file reaches its place, and how a plain file is written when
// the person asked for one. This is the only code that changes a file inside
// a locked folder, and the order of its steps is the Locker's promise:
//
//   write new, flush, verify, then replace.
//
// `sealFile` puts a new locked copy beside the old one without ever touching
// the old one until the new one has been read back and matched, and it swaps
// them in with one rename. A crash at any point leaves either the old file
// whole, or the new one whole: never a half-written file under the real name,
// and never nothing.
//
// The bytes are staged in Office's own profile folder first (`stageDir`), so
// the half-written file never reaches a synced folder. If the two are on
// different drives, the finished, verified file is copied beside its target
// under a temporary name, flushed and checked once more, and only then
// renamed into place. The rename is the only moment the real name changes.
//
// A rename that fails (a drive client or an antivirus holding the file) is
// tried a few times and then given up on. The target is left exactly as it
// was. Nothing is ever deleted here on the way to a failure.

import crypto from 'node:crypto';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { sealToStage, verifyLocked, openLocked, fileHash, LockerError } from './format.js';

const RETRY_CODES = new Set(['EPERM', 'EBUSY', 'EACCES', 'ENOTEMPTY']);
const SUFFIX = '.part';

/** A name for a working file: random, says nothing about the file. */
const tempName = () => `.rl-${crypto.randomBytes(12).toString('hex')}${SUFFIX}`;
export const workingName = tempName;

/** Whether a name is one of these working files: the scan never lists them and the sweep may remove them. */
export const isWorkingFile = (name) => /^\.rl-[0-9a-f]{24}\.part$/.test(name);

/** The system's rename, retried while something else holds the file. */
async function renameRetry(rename, from, to, retries) {
  for (let i = 0; ; i += 1) {
    try {
      await rename(from, to);
      return;
    } catch (err) {
      if (!RETRY_CODES.has(err?.code) || i >= retries) throw err;
      await new Promise((r) => setTimeout(r, 30 * 2 ** i));
    }
  }
}

/** The bytes of a file flushed to the disk, where the operating system lets a directory be asked. */
async function flushDir(dir) {
  try {
    const d = await fsp.open(dir, 'r');
    try {
      await d.sync();
    } finally {
      await d.close();
    }
  } catch {
    // Windows will not open a directory this way; a rename there is journalled.
  }
}

/**
 * Sealed, verified and swapped in. The source is not touched and the target's
 * old content is untouched until the moment of the rename.
 *
 * @param {object} o
 * @param {string} o.source      the plain file
 * @param {string} o.target      the locked file's place
 * @param {Buffer} o.rootKey     the folder's key
 * @param {Buffer} o.headerBuf   the file's header
 * @param {object} o.meta        what the metadata frame says besides the hash and size
 * @param {string} o.stageDir    a folder in Office's own profile
 * @param {string} [o.previous]  the present locked copy, for frames that have not changed
 * @param {{ at?: (step: string) => any, rename?: Function, retries?: number }} [o.hooks] for the tests
 * @returns {Promise<{ size, sha256, reused, frames, modified }>}
 */
export async function sealFile({ source, target, rootKey, headerBuf, meta, stageDir, previous = null, hooks = null }) {
  const at = async (step) => { await hooks?.at?.(step); };
  const rename = hooks?.rename || fsp.rename;
  const retries = hooks?.retries ?? 5;
  await fsp.mkdir(stageDir, { recursive: true });
  const stage = path.join(stageDir, tempName());
  let sealed = null;
  try {
    sealed = await sealToStage({ source, stagePath: stage, rootKey, headerBuf, meta, previous, hooks });
    await at('staged');
    // Read back: the new file, every byte, before anything else happens.
    const read = await verifyLocked(stage, rootKey);
    if (read.sha256 !== sealed.sha256 || read.size !== sealed.size) throw new LockerError('failed', 'The locked copy did not read back as the original.');
    await at('verified');

    await placeStaged(stage, target, async (near) => {
      const again = await verifyLocked(near, rootKey);
      if (again.sha256 !== sealed.sha256 || again.size !== sealed.size) throw new LockerError('failed', 'The locked copy did not read back as the original.');
    }, { rename, retries, at });
    await at('swapped');
    await fsp.rm(stage, { force: true }).catch(() => {});
    return sealed;
  } catch (err) {
    await fsp.rm(stage, { force: true }).catch(() => {});
    throw err;
  }
}

/**
 * A finished, verified file swapped into its place: renamed when the stage is
 * on the same drive, otherwise copied beside the target, flushed, checked by
 * `check`, and renamed. The target's old content stays untouched until the rename.
 */
export async function placeStaged(stage, target, check, { rename = fsp.rename, retries = 5, at = async () => {} } = {}) {
  try {
    await renameRetry(rename, stage, target, retries);
  } catch (err) {
    if (err?.code !== 'EXDEV') throw err;
    await copyIn(stage, target, check, { rename, retries, at });
  }
  await flushDir(path.dirname(target));
}

/** The finished file carried across drives: copied beside its target, flushed, read back, then renamed. */
async function copyIn(stage, target, check, { rename, retries, at }) {
  const dir = path.dirname(target);
  const near = path.join(dir, tempName());
  try {
    await fsp.copyFile(stage, near, fsp.constants.COPYFILE_EXCL);
    const fh = await fsp.open(near, 'r+');
    try {
      await fh.sync();
    } finally {
      await fh.close();
    }
    await at('copied');
    await check(near);
    await renameRetry(rename, near, target, retries);
  } catch (err) {
    await fsp.rm(near, { force: true }).catch(() => {});
    throw err;
  }
}

/**
 * The plain bytes of a locked file written to `dest`: into a working file
 * beside it, flushed, its hash checked against the record, and renamed into
 * place. A destination that already exists is refused unless `overwrite`.
 *
 * @param {object} o
 * @param {string} o.locked   the locked file
 * @param {string} o.dest     where the plain file goes
 * @param {Buffer} o.rootKey
 * @param {boolean} [o.overwrite]
 * @returns {Promise<{ size, sha256, meta }>}
 */
export async function writePlain({ locked, dest, rootKey, overwrite = false, hooks = null }) {
  if (!overwrite) {
    try {
      await fsp.access(dest);
      throw new LockerError('exists', 'A file of that name is already there.');
    } catch (err) {
      if (err.code === 'exists') throw err;
    }
  }
  const opened = await openLocked(locked, rootKey);
  try {
    return await writePlainFrom(opened, dest, { overwrite: true, hooks });
  } finally {
    await opened.close();
  }
}

/**
 * The bytes of something that opens as frames written to `dest` the same
 * way: a working file beside it, flushed, its hash checked, then renamed. The
 * caller closes `opened`. Takes a locked file (openLocked) or one in pieces.
 */
export async function writePlainFrom(opened, dest, { overwrite = false, hooks = null } = {}) {
  const at = async (step) => { await hooks?.at?.(step); };
  const rename = hooks?.rename || fsp.rename;
  const dir = path.dirname(dest);
  await fsp.mkdir(dir, { recursive: true });
  if (!overwrite) {
    try {
      await fsp.access(dest);
      throw new LockerError('exists', 'A file of that name is already there.');
    } catch (err) {
      if (err.code === 'exists') throw err;
    }
  }
  const part = path.join(dir, tempName());
  let out = null;
  try {
    out = await fsp.open(part, 'wx');
    let pos = 0;
    for await (const chunk of opened.frames()) {
      await out.write(chunk, 0, chunk.length, pos);
      pos += chunk.length;
    }
    await out.sync();
    await out.close();
    out = null;
    await at('written');
    const check = await fileHash(part);
    if (check.sha256 !== opened.meta.sha256 || check.size !== opened.meta.size) throw new LockerError('failed', 'The unlocked copy did not match the locked file.');
    // A file already there is replaced only by the finished, checked copy.
    await renameRetry(rename, part, dest, hooks?.retries ?? 5);
    await flushDir(dir);
    await at('swapped');
    return { size: check.size, sha256: check.sha256, meta: opened.meta };
  } catch (err) {
    await out?.close().catch(() => {});
    await fsp.rm(part, { force: true }).catch(() => {});
    throw err;
  }
}

/**
 * Left-over working files in a folder, removed: whatever is there is not a
 * person's file. `remove` is how: one may hold a plain copy, and is overwritten first.
 */
export async function sweepWorking(dir, remove = (p) => fsp.rm(p, { force: true })) {
  let names = [];
  try {
    names = await fsp.readdir(dir);
  } catch {
    return 0;
  }
  let n = 0;
  for (const name of names) {
    if (!isWorkingFile(name)) continue;
    await remove(path.join(dir, name)).catch(() => {});
    n += 1;
  }
  return n;
}
