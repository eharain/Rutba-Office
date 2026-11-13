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

/** What the folder remembers: { seen: { [office]: sequence }, removals: [id] }. */
export function readState(dir) {
  try {
    const s = JSON.parse(fs.readFileSync(path.join(dir, STATE_FILE), 'utf8'));
    return {
      seen: s?.seen && typeof s.seen === 'object' ? Object.fromEntries(Object.entries(s.seen).filter(([, n]) => Number.isInteger(n))) : {},
      removals: Array.isArray(s?.removals) ? s.removals.filter((id) => ID.test(String(id))) : [],
    };
  } catch {
    return { seen: {}, removals: [] };
  }
}

/** Writes the state whole, so a stop half-way leaves the old one. */
export function writeState(dir, state) {
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, STATE_FILE);
  const part = `${file}.${process.pid}.part`;
  fs.writeFileSync(part, `${JSON.stringify({ seen: state.seen || {}, removals: state.removals || [] }, null, 2)}\n`);
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
 * replaces as the previous. The envelope moves last, so until it is there the
 * new folder is never loaded.
 * @param {string} dir      the add-ons folder
 * @param {string} from     a folder holding <id>/ and <id>.pack.json
 */
export function putInPlace(dir, from, id) {
  const p = pathsOf(dir, id);
  const newFolder = path.join(from, id);
  const newEnvelope = path.join(from, `${id}.pack.json`);
  if (!fs.existsSync(newFolder) || !fs.existsSync(newEnvelope)) throw new Error(`${id} is not staged whole.`);
  if (fs.existsSync(p.envelope) && fs.existsSync(p.folder)) {
    fs.rmSync(p.previous, { recursive: true, force: true });
    fs.mkdirSync(p.previous, { recursive: true });
    fs.renameSync(p.envelope, path.join(p.previous, `${id}.pack.json`));
    fs.renameSync(p.folder, path.join(p.previous, id));
  } else {
    fs.rmSync(p.envelope, { force: true });
    fs.rmSync(p.folder, { recursive: true, force: true });
  }
  fs.renameSync(newFolder, p.folder);
  fs.renameSync(newEnvelope, p.envelope);
}

/** Drops the kept previous version once the new one has loaded. */
export function forgetPrevious(dir, id) {
  fs.rmSync(pathsOf(dir, id).previous, { recursive: true, force: true });
}

/** The staging folder for one Office version. */
export const stagingFor = (dir, office) => path.join(dir, STAGING, office);
