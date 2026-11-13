// The add-ons folder: where the languages, dictionaries, fonts, voices and
// speech engine a person adds are kept, apart from the program, its app
// packs and the profile (private design: office-pack-compatibility-design.md 1.1).
//
//   Installed for this person  %LOCALAPPDATA%\Rutba Office\addons\   written as the person
//   Installed for everyone     %ProgramData%\Rutba Office\addons\    written through the administrator's prompt
//   Portable                   addons\ beside the program               written as the person
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
import { execFileSync } from 'node:child_process';

const require = createRequire(import.meta.url);
const fs = process.versions.electron ? require('original-fs') : require('node:fs');

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
 * @returns {{ dir: string, scope: 'person'|'machine'|'portable'|'named' } | null}
 */
export function addonsFolder({ execPath, isPackaged, env = process.env, platform = process.platform, appData = null }) {
  if (env.RUTBA_ADDONS_DIR) return { dir: path.resolve(env.RUTBA_ADDONS_DIR), scope: 'named' };
  if (!isPackaged) return null;
  if (env.PORTABLE_EXECUTABLE_DIR) return { dir: path.join(env.PORTABLE_EXECUTABLE_DIR, 'addons'), scope: 'portable' };
  if (platform === 'win32') {
    const programFiles = [env.ProgramFiles, env['ProgramFiles(x86)'], env.ProgramW6432].filter(Boolean).map((p) => path.resolve(p).toLowerCase() + path.sep);
    const forEveryone = programFiles.some((p) => path.resolve(execPath).toLowerCase().startsWith(p));
    if (forEveryone) return { dir: path.join(env.ProgramData || 'C:\\ProgramData', 'Rutba Office', 'addons'), scope: 'machine' };
    return { dir: path.join(env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local'), 'Rutba Office', 'addons'), scope: 'person' };
  }
  return { dir: path.join(appData || path.join(os.homedir(), '.local', 'share'), 'Rutba Office', 'addons'), scope: 'person' };
}

/**
 * The permissions of an install for everyone's add-ons folder, set by the
 * administrator's step each time it runs: SYSTEM and Administrators may
 * change it, everyone signed in may only read it, nothing inherited from
 * ProgramData (whose own let any user make files), and the owner of a file
 * gets no more than that. What is already in it is reset to the same. The
 * arguments for icacls, in the order they are run.
 */
export function machineFolderAcl(dir) {
  return [
    [dir, '/inheritance:r', '/grant:r', '*S-1-5-18:(OI)(CI)F', '*S-1-5-32-544:(OI)(CI)F', '*S-1-5-32-545:(OI)(CI)RX', '*S-1-3-4:(OI)(CI)RX', '/C', '/Q'],
    [path.join(dir, '*'), '/reset', '/T', '/C', '/Q'],
  ];
}

/** Sets machineFolderAcl() with icacls (Windows); answers whether it held. */
export function secureMachineFolder(dir, { run = (args) => execFileSync(icaclsPath(), args, { windowsHide: true, stdio: 'ignore' }) } = {}) {
  try {
    fs.mkdirSync(dir, { recursive: true });
    const [root, inside] = machineFolderAcl(dir);
    run(root);
    // An empty folder has nothing inside for the second to reset.
    if (fs.readdirSync(dir).length) run(inside);
    return true;
  } catch {
    return false;
  }
}
const icaclsPath = () => path.join(process.env.SystemRoot || process.env.WINDIR || 'C:\\Windows', 'System32', 'icacls.exe');

/** Whether this process can write in a folder (an install for everyone's, as the person, cannot). */
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
 * Takes the folder for this process, so two Office processes never change it
 * at once. Answers a release function, or null when another live process
 * holds it. A lock left by a process that has gone is taken over.
 */
export function lockFolder(dir, { pid = process.pid, isAlive = alive } = {}) {
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, LOCK_FILE);
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      fs.writeFileSync(file, JSON.stringify({ pid, at: new Date().toISOString() }), { flag: 'wx' });
      return () => {
        try {
          if (JSON.parse(fs.readFileSync(file, 'utf8')).pid === pid) fs.rmSync(file, { force: true });
        } catch {
          // already gone
        }
      };
    } catch (err) {
      if (err?.code !== 'EEXIST') return null;
      let holder = null;
      try {
        holder = JSON.parse(fs.readFileSync(file, 'utf8')).pid;
      } catch {
        holder = null;
      }
      if (holder && holder !== pid && isAlive(holder)) return null;
      fs.rmSync(file, { force: true });
    }
  }
  return null;
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
      fs.rmSync(p.envelope, { force: true });
      fs.rmSync(p.folder, { recursive: true, force: true });
      fs.rmSync(p.previous, { recursive: true, force: true });
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
  fs.rmSync(p.envelope, { force: true });
  fs.rmSync(p.folder, { recursive: true, force: true });
  fs.renameSync(prevFolder, p.folder);
  fs.renameSync(prevEnvelope, p.envelope);
  fs.rmSync(p.previous, { recursive: true, force: true });
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
  try {
    if (fs.existsSync(p.envelope) && fs.existsSync(p.folder)) {
      fs.rmSync(p.previous, { recursive: true, force: true });
      fs.mkdirSync(p.previous, { recursive: true });
      step(() => fs.renameSync(p.folder, prevFolder), () => fs.renameSync(prevFolder, p.folder));
      step(() => fs.renameSync(p.envelope, prevEnvelope), () => fs.renameSync(prevEnvelope, p.envelope));
    } else {
      // Half an add-on (a stop half-way before): nothing of it to keep.
      fs.rmSync(p.folder, { recursive: true, force: true });
      fs.rmSync(p.envelope, { force: true });
    }
    step(() => fs.renameSync(newFolder, p.folder), () => fs.renameSync(p.folder, newFolder));
    step(() => fs.renameSync(newEnvelope, p.envelope), null);
  } catch (err) {
    for (const back of undo) {
      try { back(); } catch { /* as much as can be put back */ }
    }
    throw err;
  }
}

/** Drops the kept previous version once the new one has loaded. */
export function forgetPrevious(dir, id) {
  fs.rmSync(pathsOf(dir, id).previous, { recursive: true, force: true });
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
 * version's staging (or the `partsIn` folders named: an install for
 * everyone downloads into the person's), downloads nothing has added to for
 * PARTS_KEEP_MS.
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
    if (!name.endsWith('.previous')) continue;
    const id = name.slice(0, -'.previous'.length);
    if (!ID.test(id)) continue;
    const p = pathsOf(dir, id);
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
    let inProgress = [];
    try {
      inProgress = fs.readdirSync(at);
    } catch {
      inProgress = [];
    }
    for (const name of inProgress) {
      if (now - newestIn(path.join(at, name)) < keepPartsMs) continue;
      try {
        fs.rmSync(path.join(at, name), { recursive: true, force: true });
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
