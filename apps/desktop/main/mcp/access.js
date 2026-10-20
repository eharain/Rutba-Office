// Who may use the AI tools, where, and what they did (D-2026-10-16-06, -07).
//
// Off unless the person turns them on in Rutba Office: for this session, which
// ends when the Rutba Office that turned them on closes, or always. The
// setting is read afresh for every call, so switching it off takes effect at
// once, mid-conversation. Every call is kept in a log the person can read in
// the AI Tools window: the tool, the files it named, whether it was done and
// how long it took, and never a file's contents.

import fs from 'node:fs';
import path from 'node:path';

export const MODE_KEY = 'ai.tools.mode';
export const SESSION_KEY = 'ai.tools.sessionPid';
export const FOLDERS_KEY = 'ai.tools.folders';
export const MODES = ['off', 'session', 'always'];
export const LOG_FILE = 'ai-tools-log.jsonl';
const LOG_KEEP = 500;

/** Whether a process is running. */
export function alive(pid) {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return err.code === 'EPERM';
  }
}

/**
 * Whether the tools answer now, and if not, the sentence that says how to
 * turn them on. `settings` is the settings object as saved (settings.json).
 */
export function openNow(settings, { isAlive = alive } = {}) {
  // AI Tools is an add-on (D-2026-10-20-19): not added in Choose apps, the tools answer nothing.
  const choice = settings?.['apps.choice'];
  if (choice && !(Array.isArray(choice.addons) && choice.addons.includes('ai'))) {
    return { open: false, mode: 'off', reason: 'AI Tools is not added to this copy of Rutba Office. To add it, open Rutba Office, choose Choose apps in the launcher, and turn on AI Tools under Add-ons.' };
  }
  const mode = MODES.includes(settings?.[MODE_KEY]) ? settings[MODE_KEY] : 'off';
  if (mode === 'always') return { open: true, mode };
  if (mode === 'session' && isAlive(Number(settings?.[SESSION_KEY]))) return { open: true, mode };
  return {
    open: false,
    mode: mode === 'session' ? 'off' : mode,
    reason: 'Rutba Office’s AI tools are off. To let an AI app on this computer use them, open Rutba Office and choose AI Tools from the app menu, then turn them on for this session or always.',
  };
}

/** The folders the tools may use, as saved: absolute paths, resolved. */
export function foldersOf(settings) {
  const list = Array.isArray(settings?.[FOLDERS_KEY]) ? settings[FOLDERS_KEY] : [];
  return [...new Set(list.filter((f) => typeof f === 'string' && path.isAbsolute(f)).map((f) => path.resolve(f)))];
}

const inside = (folder, p) => {
  const rel = path.relative(folder, p);
  return rel === '' || (!!rel && !rel.startsWith('..') && !path.isAbsolute(rel));
};
const same = (a, b) => (process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b);
const insideAny = (folders, p) => folders.some((f) => (process.platform === 'win32' ? inside(f.toLowerCase(), p.toLowerCase()) : inside(f, p)) || same(f, p));

/**
 * The full path of `given` if it lies in one of `folders`, else an error
 * that says which folders it may use. A relative path is taken from the first
 * folder. A path through a link that leads out of the folders is refused
 * too: what counts is where the file really is.
 */
export function resolveIn(folders, given) {
  if (!folders.length) throw new Error('No folder has been given to the AI tools yet. In Rutba Office, choose AI Tools from the app menu and add a folder.');
  const raw = String(given || '').trim();
  if (!raw) throw new Error('A file or folder name is needed.');
  const full = path.resolve(path.isAbsolute(raw) ? raw : path.join(folders[0], raw));
  // Where it really is: the deepest part of the path that exists, its links followed.
  let probe = full;
  while (!fs.existsSync(probe) && path.dirname(probe) !== probe) probe = path.dirname(probe);
  let real = full;
  try {
    real = path.join(fs.realpathSync.native(probe), path.relative(probe, full));
  } catch {
    real = full;
  }
  const realFolders = folders.map((f) => { try { return fs.realpathSync.native(f); } catch { return f; } });
  if (!insideAny(folders, full) || !insideAny(realFolders, real)) {
    throw new Error(`The AI tools may only use these folders: ${folders.join('; ')}. ${raw} is outside them.`);
  }
  return full;
}

/** A path for a new file: in the folders, and not a file already there (the tools never overwrite). */
export function newFileIn(folders, given, extensions) {
  const full = resolveIn(folders, given);
  const ext = path.extname(full).toLowerCase();
  if (extensions && !extensions.includes(ext)) throw new Error(`The file must end in ${extensions.join(' or ')}.`);
  if (fs.existsSync(full)) throw new Error(`${path.basename(full)} already exists, and the AI tools never overwrite a file. Choose another name.`);
  return full;
}

/** Keep one call in the log, the newest last, the oldest dropped past LOG_KEEP. */
export function logCall(dir, entry) {
  const file = path.join(dir, LOG_FILE);
  try {
    fs.mkdirSync(dir, { recursive: true });
    const lines = fs.existsSync(file) ? fs.readFileSync(file, 'utf8').split('\n').filter(Boolean) : [];
    lines.push(JSON.stringify(entry));
    fs.writeFileSync(file, `${lines.slice(-LOG_KEEP).join('\n')}\n`);
  } catch {
    // A log that cannot be written does not stop the work, and is not the work.
  }
}

export const CONNECTIONS_FILE = 'ai-tools-connections.json';

/**
 * Who is connected: one entry per AI app that has started the tools and not
 * yet gone, with what it said it is (its name and version), where it connects
 * from, when it connected and when it last asked for something. Kept by the
 * tools' process (node.js), read by the AI Tools window, which leaves out an
 * entry whose process has gone (alive).
 */
export function writeConnection(dir, pid, entry) {
  const file = path.join(dir, CONNECTIONS_FILE);
  let all = {};
  try {
    all = JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    all = {};
  }
  if (entry) all[pid] = { ...(all[pid] || {}), ...entry, pid };
  else delete all[pid];
  try {
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(file, JSON.stringify(all, null, 2));
  } catch {
    // Who is connected is shown, not needed: the work goes on.
  }
}

/** The AI apps connected now, the newest first. */
export function readConnections(dir, { isAlive = alive } = {}) {
  try {
    const all = JSON.parse(fs.readFileSync(path.join(dir, CONNECTIONS_FILE), 'utf8'));
    return Object.values(all).filter((c) => isAlive(Number(c.pid))).sort((a, b) => String(b.since).localeCompare(String(a.since)));
  } catch {
    return [];
  }
}

/** The log, newest first. */
export function readLog(dir, limit = 100) {
  try {
    return fs.readFileSync(path.join(dir, LOG_FILE), 'utf8').split('\n').filter(Boolean).map((l) => JSON.parse(l)).reverse().slice(0, limit);
  } catch {
    return [];
  }
}
