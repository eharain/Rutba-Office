// Adding and removing app packs (docs/APP-PACKS.md 5.4), on disk.
//
// A pack is added in two steps. Its three things (the archive, the folder of
// files kept beside it and the signed manifest) are copied into a staging
// folder of their own and checked there by the loader, with the core's keys
// and version, exactly as a start would check them; only a pack that passes
// is copied into the packs folder, its manifest last, so a copy cut short
// leaves no manifest and so no pack. It loads at the next start.
//
// A pack is removed at the next start, before the loader runs: Windows will
// not delete a file a running Electron holds open, and a loaded pack's
// archive is held. Until then it is written down in the packs folder
// (pending.json), and a pack written down to go is not loaded even when its
// files could not be deleted yet (another profile's Rutba Office still
// running holds them): it is tried again at the start after.
//
// Only the packs folder is touched: the person's documents, mail, settings
// and profile data are not here, and stay.
//
// No Electron import; under Electron the files are handled with original-fs,
// which sees an .asar as the file it is rather than as a folder.

import path from 'node:path';
import { createRequire } from 'node:module';
import { findPacks } from './packs.js';

const require = createRequire(import.meta.url);
const fs = process.versions.electron ? require('original-fs') : require('node:fs');

const ID = /^[a-z][a-z0-9-]{0,31}$/;
const PENDING = 'pending.json';

/** A pack's three things, by name in its folder. */
export function packFiles(id) {
  return { manifest: `${id}.pack.json`, archive: `${id}.asar`, unpacked: `${id}.asar.unpacked` };
}

/** Bytes of a pack as it lies in a folder, 0 when it is not there. */
export function packSize(dir, id) {
  const f = packFiles(id);
  const sizeOf = (p) => {
    try {
      const st = fs.statSync(p);
      if (!st.isDirectory()) return st.size;
      return fs.readdirSync(p).reduce((n, name) => n + sizeOf(path.join(p, name)), 0);
    } catch {
      return 0;
    }
  };
  return sizeOf(path.join(dir, f.manifest)) + sizeOf(path.join(dir, f.archive)) + sizeOf(path.join(dir, f.unpacked));
}

/** The packs a folder holds, by the manifests in it. */
export function packsIn(dir) {
  try {
    return fs.readdirSync(dir).filter((f) => f.endsWith('.pack.json')).map((f) => f.slice(0, -'.pack.json'.length)).filter((id) => ID.test(id)).sort();
  } catch {
    return [];
  }
}

/** Whether this process may write in a folder: tried, since on Windows the read-only flag says nothing of who may write. */
export function canWriteIn(dir) {
  const probe = path.join(dir, `.write-${process.pid}-${Date.now()}`);
  try {
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(probe, '');
    fs.rmSync(probe, { force: true });
    return true;
  } catch {
    return false;
  }
}

/* ── removal at the next start ─────────────────────────────────────── */

function readPending(dir) {
  try {
    const p = JSON.parse(fs.readFileSync(path.join(dir, PENDING), 'utf8'));
    return { remove: Array.isArray(p?.remove) ? p.remove.filter((id) => ID.test(String(id))) : [] };
  } catch {
    return { remove: [] };
  }
}

function writePending(dir, pending) {
  const file = path.join(dir, PENDING);
  if (!pending.remove.length) {
    fs.rmSync(file, { force: true });
    return;
  }
  fs.writeFileSync(file, `${JSON.stringify({ remove: [...new Set(pending.remove)].sort() }, null, 2)}\n`);
}

/** The packs written down to go at the next start. */
export function pendingRemovals(dir) {
  return readPending(dir).remove;
}

/** A pack written down to go at the next start. */
export function markRemoval(dir, id) {
  if (!ID.test(String(id))) throw new Error(`${id} is not the name of a pack.`);
  const pending = readPending(dir);
  if (!pending.remove.includes(id)) pending.remove.push(id);
  writePending(dir, pending);
  return pending.remove;
}

/** A removal taken back before the start that would have made it. */
export function unmarkRemoval(dir, id) {
  const pending = readPending(dir);
  pending.remove = pending.remove.filter((x) => x !== id);
  writePending(dir, pending);
  return pending.remove;
}

/** One pack's files deleted, its manifest first so that what is left of it is no pack at all. */
function deletePack(dir, id) {
  const f = packFiles(id);
  fs.rmSync(path.join(dir, f.manifest), { force: true });
  fs.rmSync(path.join(dir, f.archive), { force: true });
  fs.rmSync(path.join(dir, f.unpacked), { recursive: true, force: true });
}

/**
 * At start, before the loader: the packs written down to go are deleted.
 * One whose files are still held stays written down, and is answered in
 * `held` so the start leaves it out all the same.
 * @returns {{ removed: string[], held: string[] }}
 */
export function applyRemovals(dir) {
  const pending = readPending(dir);
  const removed = [];
  const held = [];
  for (const id of pending.remove) {
    try {
      deletePack(dir, id);
      removed.push(id);
    } catch {
      held.push(id);
    }
  }
  if (removed.length) writePending(dir, { remove: held });
  return { removed, held };
}

/* ── adding: staged, checked, then put in place ────────────────────── */

/** Copies a file or a folder, without following links. */
function copy(from, to) {
  const st = fs.lstatSync(from);
  if (st.isSymbolicLink()) throw new Error(`${path.basename(from)} is a link; a pack holds files only.`);
  if (st.isDirectory()) {
    fs.mkdirSync(to, { recursive: true });
    for (const name of fs.readdirSync(from)) copy(path.join(from, name), path.join(to, name));
    return;
  }
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(from, to);
}

/**
 * A pack copied from a folder that holds it (a folder of packs, or one a
 * release's pack file was unpacked into) into a staging folder of its own,
 * and checked there. Answers the pack as the loader describes it; throws a
 * sentence for a person when it is missing or does not pass.
 *
 * @param {object} o
 * @param {string} o.from      the folder the pack is in
 * @param {string} o.id        the pack wanted
 * @param {string} o.staging   a folder for this pack alone; emptied first
 * @param {*} o.publicKey      the core's keys
 * @param {string} o.version   the core's version
 * @param {string[]} [o.present] the packs installed, which this one may need
 * @param {string} [o.name]    the pack's name for the words, until its manifest is believed
 */
export function stagePack({ from, id, staging, publicKey, version, present = [], name = id }) {
  if (!ID.test(String(id))) throw new Error(`${id} is not the name of a pack.`);
  const f = packFiles(id);
  if (!fs.existsSync(path.join(from, f.manifest)) || !fs.existsSync(path.join(from, f.archive))) {
    throw new Error(`The ${name} pack is not in ${from}.`);
  }
  fs.rmSync(staging, { recursive: true, force: true });
  fs.mkdirSync(staging, { recursive: true });
  copy(path.join(from, f.manifest), path.join(staging, f.manifest));
  copy(path.join(from, f.archive), path.join(staging, f.archive));
  if (fs.existsSync(path.join(from, f.unpacked))) copy(path.join(from, f.unpacked), path.join(staging, f.unpacked));
  const { packs, refused } = findPacks({ dir: staging, publicKey, version, present });
  const bad = refused.find((r) => r.id === id);
  // The loader's words, said of an install rather than a start.
  if (bad) throw Object.assign(new Error(bad.reason.replace(/It was not loaded\.$/, 'It was not installed.')), { code: bad.code });
  const pack = packs.find((p) => p.id === id);
  if (!pack) throw new Error(`The ${name} pack could not be checked.`);
  return pack;
}

/**
 * A staged pack put in place in the packs folder: its archive and files
 * first, its manifest last. A pack of that name already there (one that
 * would not load, being installed again) is replaced.
 */
export function installStaged({ staging, id, dir }) {
  const f = packFiles(id);
  fs.mkdirSync(dir, { recursive: true });
  deletePack(dir, id);
  copy(path.join(staging, f.archive), path.join(dir, f.archive));
  if (fs.existsSync(path.join(staging, f.unpacked))) copy(path.join(staging, f.unpacked), path.join(dir, f.unpacked));
  copy(path.join(staging, f.manifest), path.join(dir, f.manifest));
  unmarkRemoval(dir, id);
  fs.rmSync(staging, { recursive: true, force: true });
  return { id, size: packSize(dir, id) };
}
