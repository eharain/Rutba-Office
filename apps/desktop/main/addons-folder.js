// The add-ons folder: where the languages, dictionaries, fonts, voices and
// speech engine a person adds are kept, apart from the program, its app
// packs and the profile (private design: office-pack-compatibility-design.md 1.1).
//
//   Installed (for this person or for everyone)  %LOCALAPPDATA%\Rutba Office\addons\
//   Portable                                     addons\ beside the program
//
// Each is written as the person. In 2.26.0 an install for everyone keeps
// add-ons per person too, each person on the computer adding their own
// (D-2026-11-13-35): a folder for everyone, written through the
// administrator's prompt, could be made or planted by any user before that
// prompt ran, so it waits for a later release. Nothing reads
// %ProgramData%\Rutba Office\addons.
//
// An Office update replaces the program folder and leaves this one alone; an
// uninstall asks whether to remove it. Inside:
//
//   <id>/  <id>.pack.json        an add-on and its signed envelope (addons.js reads them)
//   <id>.previous/               the version kept for rollback: <id>/ and <id>.pack.json
//   .staging/<office version>/   packs staged for that version, checked, each <id>/ beside its envelope;
//                                .parts/ inside it holds downloads in progress
//   .lock                        two Office processes kept apart
//   state.json                   the highest feed sequence seen per Office version, removals waiting
//
// Nothing here checks a pack: addons.js does, at every start.

import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const fs = process.versions.electron ? require('original-fs') : require('node:fs');

/**
 * Every folder the add-ons code removes goes through Office's one removeTree
 * (D-2026-11-13-37), which never goes through a link: fs.rm does on Node 24,
 * the one Electron 44 runs. Handed on from here so the add-ons code names it
 * in one place.
 */
import { removeTree } from '@rutba/office-shell/electron/remove-tree';
export { removeTree };

const ID = /^[a-z][a-z0-9-]{0,31}$/;
export const STATE_FILE = 'state.json';
export const LOCK_FILE = '.lock';
export const STAGING = '.staging';
/** Inside a version's staging folder: downloads in progress, never moved in. */
export const PARTS = '.parts';

/**
 * Where this copy keeps its add-ons, and who may write there.
 * @param {object} o
 * @param {string} o.execPath       the program
 * @param {boolean} o.isPackaged    a copy run from the source has none unless one is named
 * @param {NodeJS.ProcessEnv} [o.env]
 * @param {string} [o.platform]
 * @param {string} [o.appData]      the operating system's per-person data folder, off Windows
 * @returns {{ dir: string, scope: 'person'|'portable'|'named' } | null}
 */
export function addonsFolder({ execPath, isPackaged, env = process.env, platform = process.platform, appData = null }) {
  if (env.RUTBA_ADDONS_DIR) return { dir: path.resolve(env.RUTBA_ADDONS_DIR), scope: 'named' };
  if (!isPackaged) return null;
  if (env.PORTABLE_EXECUTABLE_DIR) return { dir: path.join(env.PORTABLE_EXECUTABLE_DIR, 'addons'), scope: 'portable' };
  // Installed for this person or for everyone alike: the person's own (D-2026-11-13-35).
  if (platform === 'win32') {
    return { dir: path.join(env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local'), 'Rutba Office', 'addons'), scope: 'person' };
  }
  return { dir: path.join(appData || path.join(os.homedir(), '.local', 'share'), 'Rutba Office', 'addons'), scope: 'person' };
}

/** Whether this process can write in a folder. */
export function canWrite(dir) {
  try {
    fs.mkdirSync(dir, { recursive: true });
    const probe = path.join(dir, `.write-${process.pid}-${Date.now()}`);
    fs.writeFileSync(probe, '');
    fs.rmSync(probe, { force: true });
    return true;
  } catch {
    return false;
  }
}

const alive = (pid) => {
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return err?.code === 'EPERM';
  }
};

/**
 * A lock older than this is taken over whatever process it names: a change
 * holds the folder for moments, and Windows gives a gone process's number to
 * another program, which would otherwise keep every change out while it runs.
 */
export const STALE_LOCK_MS = 60 * 1000;
/** A takeover left half-way by a process that stopped in it. */
const TAKEOVER_STALE_MS = 10 * 1000;

/**
 * A stale lock removed by one process alone, and only while it is still the
 * lock judged stale: two processes taking over at once must not both hold the
 * folder, the second removing the first's new lock. Answers whether it may
 * try to take the folder again.
 */
function takeOver(file, seen, now) {
  const guard = `${file}.takeover`;
  try {
    fs.writeFileSync(guard, '', { flag: 'wx' });
  } catch (err) {
    if (err?.code === 'EEXIST') {
      try {
        if (now() - fs.statSync(guard).mtimeMs > TAKEOVER_STALE_MS) fs.rmSync(guard, { force: true });
      } catch {
        // let go of meanwhile
      }
    }
    return false;
  }
  try {
    let still = null;
    try {
      still = fs.readFileSync(file, 'utf8');
    } catch {
      still = null;
    }
    if (still !== null && still !== seen) return false;
    fs.rmSync(file, { force: true });
    return true;
  } finally {
    fs.rmSync(guard, { force: true });
  }
}

/**
 * Takes the folder for this process, so two Office processes never change it
 * at once. Answers a release function, or null when another live process
 * holds it. A lock left by a process that has gone, or older than
 * STALE_LOCK_MS, is taken over; one still being written is not.
 */
export function lockFolder(dir, { pid = process.pid, isAlive = alive, now = Date.now, staleMs = STALE_LOCK_MS } = {}) {
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, LOCK_FILE);
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      fs.writeFileSync(file, JSON.stringify({ pid, at: new Date(now()).toISOString() }), { flag: 'wx' });
      const ours = () => {
        try {
          return JSON.parse(fs.readFileSync(file, 'utf8')).pid === pid;
        } catch {
          return false;
        }
      };
      const release = () => {
        try {
          if (ours()) fs.rmSync(file, { force: true });
        } catch {
          // already gone
        }
      };
      // Written again between the steps of a long hold, so that it is never taken for stale while it is in use.
      release.refresh = () => {
        try {
          if (ours()) fs.writeFileSync(file, JSON.stringify({ pid, at: new Date(now()).toISOString() }));
        } catch {
          // taken over meanwhile: the next step's lock is the holder's
        }
      };
      return release;
    } catch (err) {
      if (err?.code !== 'EEXIST') return null;
      let seen;
      try {
        seen = fs.readFileSync(file, 'utf8');
      } catch {
        continue; // let go of meanwhile
      }
      let holder = null;
      let at = NaN;
      try {
        const j = JSON.parse(seen);
        holder = Number.isInteger(j?.pid) ? j.pid : null;
        at = Date.parse(j?.at);
      } catch {
        holder = null;
      }
      if (!Number.isFinite(at)) {
        try {
          at = fs.statSync(file).mtimeMs;
        } catch {
          continue;
        }
      }
      const stale = now() - at > staleMs;
      // A lock that cannot be read yet is being written: held, until it is stale.
      const live = holder === null ? true : holder !== pid && isAlive(holder);
      if (live && !stale) return null;
      if (!takeOver(file, seen, now)) return null;
    }
  }
  return null;
}

/** Whether a path is a link (a symbolic link, or a junction, which Node names one): never cleared through. */
export function isLink(p) {
  try {
    return fs.lstatSync(p).isSymbolicLink();
  } catch {
    return false;
  }
}

/**
 * Whether `p` is inside the add-ons folder with no link on the way to it, so
 * that clearing what is in it never reaches a folder elsewhere. The add-ons
 * folder itself may be a link (AppData moved to another disk by a junction).
 */
export function linkFree(dir, p) {
  const rel = path.relative(dir, p);
  if (!rel || rel.startsWith('..') || path.isAbsolute(rel)) return false;
  let at = dir;
  for (const part of rel.split(/[\\/]/).filter(Boolean)) {
    at = path.join(at, part);
    if (isLink(at)) return false;
  }
  return true;
}

/** How long a change waits for another Office to let go of the folder, in all. */
export const HOLD_WAIT_MS = 300;

/**
 * One change to the folder made while this process holds it (lockFolder), so
 * that another Office's start does not move add-ons in or out half-way
 * through it. Waits a moment for another holder, then gives up.
 * @returns {{ done: true, value: any } | { done: false }}
 */
export function whileLocked(dir, fn, { waitMs = HOLD_WAIT_MS, step = 50, lock = lockFolder, pause = sleep } = {}) {
  let release = lock(dir);
  for (let waited = 0; !release && waited < waitMs; waited += step) {
    pause(step);
    release = lock(dir);
  }
  if (!release) return { done: false };
  try {
    return { done: true, value: fn() };
  } finally {
    release();
  }
}

/** Waits, holding the thread: only for the moment another Office takes to write a file. */
function sleep(ms) {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

/** What the folder remembers: { seen: { [office]: sequence }, removals: [id] }. */
export function readState(dir) {
  try {
    const s = JSON.parse(fs.readFileSync(path.join(dir, STATE_FILE), 'utf8'));
    return {
      seen: s?.seen && typeof s.seen === 'object' ? Object.fromEntries(Object.entries(s.seen).filter(([, n]) => Number.isInteger(n))) : {},
      removals: Array.isArray(s?.removals) ? s.removals.filter((id) => ID.test(String(id))) : [],
      feeds: s?.feeds && typeof s.feeds === 'object' ? Object.fromEntries(Object.entries(s.feeds).filter(([, h]) => /^[0-9a-f]{64}$/.test(String(h)))) : {},
    };
  } catch {
    return { seen: {}, removals: [], feeds: {} };
  }
}

/** Writes the state whole, so a stop half-way leaves the old one. */
export function writeState(dir, state) {
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, STATE_FILE);
  const part = `${file}.${process.pid}.part`;
  fs.writeFileSync(part, `${JSON.stringify({ seen: state.seen || {}, removals: state.removals || [], feeds: state.feeds || {} }, null, 2)}\n`);
  fs.renameSync(part, file);
}

/** Every path one add-on has in the folder. */
export const pathsOf = (dir, id) => ({
  folder: path.join(dir, id),
  envelope: path.join(dir, `${id}.pack.json`),
  previous: path.join(dir, `${id}.previous`),
});

/** Asks for an add-on to go at the next start, when its files are no longer open. */
export function markRemoval(dir, id) {
  if (!ID.test(String(id))) throw new Error(`${id} is not an add-on.`);
  const s = readState(dir);
  if (!s.removals.includes(id)) s.removals.push(id);
  writeState(dir, s);
}

export function unmarkRemoval(dir, id) {
  const s = readState(dir);
  s.removals = s.removals.filter((x) => x !== id);
  writeState(dir, s);
}

/**
 * The removals asked for, done before the loader runs: the envelope first, so
 * a stop half-way leaves a folder with no envelope, which is never loaded.
 * Answers { removed, held }: held are those that could not be removed now.
 */
export function applyRemovals(dir) {
  const s = readState(dir);
  const removed = [];
  const held = [];
  for (const id of s.removals) {
    const p = pathsOf(dir, id);
    try {
      removeTree(p.envelope);
      removeTree(p.folder);
      removeTree(p.previous);
      removed.push(id);
    } catch {
      held.push(id);
    }
  }
  if (removed.length) writeState(dir, { ...s, removals: held });
  return { removed, held };
}

/**
 * Puts the kept previous version back in place of the current one (a pack
 * withdrawn by its makers). Answers false when there is none to go back to.
 */
export function rollBack(dir, id) {
  const p = pathsOf(dir, id);
  const prevEnvelope = path.join(p.previous, `${id}.pack.json`);
  const prevFolder = path.join(p.previous, id);
  if (!fs.existsSync(prevEnvelope) || !fs.existsSync(prevFolder)) return false;
  removeTree(p.envelope);
  removeTree(p.folder);
  fs.renameSync(prevFolder, p.folder);
  fs.renameSync(prevEnvelope, p.envelope);
  removeTree(p.previous);
  return true;
}

/**
 * Moves a checked add-on from a staging folder into place, keeping the one it
 * replaces as the previous. The current folder moves first: on Windows a
 * folder with a file open in it (an engine loaded) cannot be renamed, and
 * then nothing has moved. Each later step that fails puts back what the
 * earlier ones moved, so the add-on in place stays whole and the staged one
 * stays staged, for the next start. The new envelope moves last, so until
 * it is there the new folder is never loaded.
 * @param {string} dir      the add-ons folder
 * @param {string} from     a folder holding <id>/ and <id>.pack.json
 */
export function putInPlace(dir, from, id) {
  const p = pathsOf(dir, id);
  const newFolder = path.join(from, id);
  const newEnvelope = path.join(from, `${id}.pack.json`);
  if (!fs.existsSync(newFolder) || !fs.existsSync(newEnvelope)) throw new Error(`${id} is not staged whole.`);
  const prevFolder = path.join(p.previous, id);
  const prevEnvelope = path.join(p.previous, `${id}.pack.json`);
  const undo = [];
  const step = (fn, back) => {
    fn();
    if (back) undo.unshift(back);
  };
  // The version kept before is moved aside, not deleted, until the new one is in: a rename that fails puts it back.
  const aside = `${p.previous}.going`;
  let asideMade = false;
  try {
    if (fs.existsSync(p.envelope) && fs.existsSync(p.folder)) {
      removeTree(aside);
      if (fs.existsSync(p.previous)) {
        step(() => fs.renameSync(p.previous, aside), () => fs.renameSync(aside, p.previous));
        asideMade = true;
      }
      step(() => fs.mkdirSync(p.previous, { recursive: true }), () => removeTree(p.previous));
      step(() => fs.renameSync(p.folder, prevFolder), () => fs.renameSync(prevFolder, p.folder));
      step(() => fs.renameSync(p.envelope, prevEnvelope), () => fs.renameSync(prevEnvelope, p.envelope));
    } else {
      // Half an add-on (a stop half-way before): nothing of it to keep.
      removeTree(p.folder);
      removeTree(p.envelope);
    }
    step(() => fs.renameSync(newFolder, p.folder), () => fs.renameSync(p.folder, newFolder));
    step(() => fs.renameSync(newEnvelope, p.envelope), null);
  } catch (err) {
    for (const back of undo) {
      try { back(); } catch { /* as much as can be put back */ }
    }
    throw err;
  }
  if (asideMade) {
    try {
      removeTree(aside);
    } catch {
      // a file of it open: let go of at a later start (letGo)
    }
  }
}

/** Drops the kept previous version once the new one has loaded. */
export function forgetPrevious(dir, id) {
  // A link is let go of itself, never what it points at (removeTree).
  removeTree(pathsOf(dir, id).previous);
}

/** How long a download nobody went back to is kept for resuming. */
export const PARTS_KEEP_MS = 7 * 24 * 60 * 60 * 1000;

/** The newest change to anything in a folder, or 0. */
function newestIn(at) {
  let newest = 0;
  const walk = (p) => {
    const st = fs.lstatSync(p);
    newest = Math.max(newest, st.mtimeMs);
    if (st.isDirectory()) for (const n of fs.readdirSync(p)) walk(path.join(p, n));
  };
  try {
    walk(at);
  } catch {
    // gone while looked at
  }
  return newest;
}

/**
 * What the folder keeps that nothing will use, let go of at a start with the
 * folder held: a kept previous version whose add-on is gone, that cannot be
 * read, or that is for another Office (rollBack goes back only within this
 * one, so an update leaves every older version's behind), and, in this
 * version's staging (or the `partsIn` folders named), downloads nothing has
 * added to for PARTS_KEEP_MS.
 * @returns {{ forgotten: string[], parts: string[] }}
 */
export function letGo(dir, { office, now = Date.now(), keepPartsMs = PARTS_KEEP_MS, partsIn = [path.join(stagingFor(dir, office), PARTS)] } = {}) {
  const forgotten = [];
  const parts = [];
  let names = [];
  try {
    names = fs.readdirSync(dir);
  } catch {
    return { forgotten, parts };
  }
  for (const name of names) {
    // A previous version moved aside by putInPlace and not let go of then (a file of it was open).
    if (name.endsWith('.previous.going') && ID.test(name.slice(0, -'.previous.going'.length))) {
      if (isLink(path.join(dir, name))) continue;
      try {
        removeTree(path.join(dir, name));
        forgotten.push(name.slice(0, -'.previous.going'.length));
      } catch {
        // still open
      }
      continue;
    }
    if (!name.endsWith('.previous')) continue;
    const id = name.slice(0, -'.previous'.length);
    if (!ID.test(id)) continue;
    const p = pathsOf(dir, id);
    // Not put there by Office: left alone.
    if (isLink(p.previous)) continue;
    let version = null;
    try {
      version = JSON.parse(fs.readFileSync(path.join(p.previous, `${id}.pack.json`), 'utf8'))?.manifest?.version || null;
    } catch {
      version = null;
    }
    if (fs.existsSync(p.envelope) && version === office && fs.existsSync(path.join(p.previous, id))) continue;
    try {
      forgetPrevious(dir, id);
      forgotten.push(id);
    } catch {
      // a file of it open: let go of at a later start
    }
  }
  for (const at of partsIn) {
    // Downloads are cleared only in the add-ons folder itself, never through a link out of it.
    if (!linkFree(dir, at)) continue;
    let inProgress = [];
    try {
      inProgress = fs.readdirSync(at);
    } catch {
      inProgress = [];
    }
    for (const name of inProgress) {
      if (isLink(path.join(at, name))) continue;
      if (now - newestIn(path.join(at, name)) < keepPartsMs) continue;
      try {
        removeTree(path.join(at, name));
        parts.push(name);
      } catch {
        // still being written
      }
    }
  }
  return { forgotten, parts };
}

/** The staging folder for one Office version. */
export const stagingFor = (dir, office) => path.join(dir, STAGING, office);
