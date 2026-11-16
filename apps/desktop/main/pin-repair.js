// Pins and shortcuts that still point at a copy of Office that is gone.
//
// An installer that moved Office from one scope to the other (everyone to
// this person, or back) could leave the old folder with only the program in
// it, or with nothing: a pin to "C:\Program Files\Rutba Office\Rutba
// Office.exe" then started a program that quit at once, and a click on the
// Mail pin opened nothing (owner, 16 Nov). Pinning again does not help:
// Windows keeps the old shortcut files, the ones under User Pinned, and the
// ImplicitAppShortcuts copy it made for the app's own window.
//
// So at start the installed copy looks at the person's own pins and
// shortcuts, and a shortcut to "Rutba Office.exe" in another folder, where
// that program is not there or is there without the app beside it, is
// pointed at this copy. Its arguments (`--app=mail`, `--profile=...`), its
// app id and its name stay as they were; its icon follows when it named the
// dead folder, its start folder when it named it or is gone.
//
// Only the installer's own folders count, for everyone and for this person:
// a copy elsewhere is the person's own choice of folder, or a build, or on a
// drive or share that is only away for now, and a pin to it is left alone,
// as is a pin to another working copy. And only once for each version and
// folder of this copy, unless the installer's other folder holds a copy that
// is only its program: a start after that reads no shortcuts.

import { win32 as path } from 'node:path';

export const PROGRAM = 'Rutba Office.exe';

/** The copy the pins were last looked at for (pinStamp). */
export const PINS_REPAIRED_KEY = 'pins.repairedFor';

/** The folders where the person's pins and shortcuts are, and how deep to look in each. */
export function shortcutPlaces({ appData, desktop }) {
  const pinned = path.join(appData, 'Microsoft', 'Internet Explorer', 'Quick Launch', 'User Pinned');
  return [
    // TaskBar\*.lnk, StartMenu\*.lnk and ImplicitAppShortcuts\<id>\*.lnk.
    { dir: pinned, depth: 2 },
    { dir: path.join(appData, 'Microsoft', 'Internet Explorer', 'Quick Launch'), depth: 0 },
    { dir: path.join(appData, 'Microsoft', 'Windows', 'Start Menu', 'Programs'), depth: 2 },
    { dir: desktop, depth: 0 },
  ];
}

const tidy = (p) => path.normalize(p).replace(/[\\/]+$/, '').toLowerCase();
const same = (a, b) => tidy(a) === tidy(b);

/** The folders the installer puts Office in: for everyone, and for this person. */
export function installerFolders(env) {
  const roots = [env.ProgramW6432, env.ProgramFiles, env.LOCALAPPDATA && path.join(env.LOCALAPPDATA, 'Programs')].filter(Boolean);
  return [...new Set(roots.map((r) => path.join(r, 'Rutba Office')))];
}

/** What a start records once its look at the pins is done. */
/**
 * What a start records once its look at the pins is done: this copy's
 * version and program, and the installer's other folders that hold only a
 * program. A start whose stamp is the one recorded looks at nothing, so a
 * folder this copy cannot clear (all-users, from a copy for this person) is
 * not a reason to read every shortcut at every start; a new one is.
 */
export function pinStamp({ version, exe, folders, exists }) {
  const husks = folders.filter((f) => !same(f, path.dirname(exe)) && exists(path.join(f, PROGRAM)) && isDeadCopy(path.join(f, PROGRAM), { exists }));
  return [version, tidy(exe), ...husks.map(tidy).sort()].join('|');
}

/** Whether this start looks at the pins: its stamp is not the one recorded. */
export const pinsDue = ({ last, ...here }) => last !== pinStamp(here);

/** Whether a program path is a copy of Office that no longer runs: gone, or with no app beside it. */
export function isDeadCopy(target, { exists }) {
  if (!exists(target)) return true;
  return !exists(path.join(path.dirname(target), 'resources', 'app.asar'));
}

/**
 * What to change in one shortcut so it starts this copy, or null to leave it.
 * @param {{ target?: string, icon?: string, cwd?: string }} link as read
 * @param {string} exe this copy's program
 */
export function repairOf(link, exe, { exists, folders }) {
  const target = link?.target;
  if (!target || path.basename(target).toLowerCase() !== PROGRAM.toLowerCase()) return null;
  const deadDir = path.dirname(target);
  if (same(deadDir, path.dirname(exe))) return null;
  if (!folders.some((f) => same(f, deadDir))) return null;
  if (!isDeadCopy(target, { exists })) return null;
  const change = { target: exe };
  // Windows takes an icon only with its index: given alone, the old one stays.
  if (link.icon && same(path.dirname(link.icon), deadDir)) Object.assign(change, { icon: exe, iconIndex: link.iconIndex ?? 0 });
  if (link.cwd && (same(link.cwd, deadDir) || !exists(link.cwd))) change.cwd = path.dirname(exe);
  return change;
}

/** Every .lnk in a folder and its subfolders, as deep as asked. */
function linksIn(dir, depth, list) {
  let entries;
  try { entries = list(dir); } catch { return []; }
  const found = [];
  for (const e of entries) {
    const full = path.join(dir, e.name);
    if (e.isDirectory) { if (depth > 0) found.push(...linksIn(full, depth - 1, list)); }
    else if (/\.lnk$/i.test(e.name)) found.push(full);
  }
  return found;
}

/**
 * Points every pin and shortcut to a dead copy of Office at this one.
 * @param {{
 *   exe: string,
 *   places: { dir: string, depth: number }[],
 *   list: (dir: string) => { name: string, isDirectory: boolean }[],
 *   read: (file: string) => { target?: string, icon?: string, cwd?: string },
 *   write: (file: string, change: object) => boolean,
 *   exists: (file: string) => boolean,
 *   folders: string[],
 * }} io with the installer's folders (installerFolders)
 * @returns {{ changed: { file: string, from: string }[], failed: number }} the shortcuts changed, and how many could not be
 */
export function repairPins({ exe, places, list, read, write, exists, folders }) {
  const changed = [];
  let failed = 0;
  const seen = new Set();
  for (const { dir, depth } of places) {
    for (const file of linksIn(dir, depth, list)) {
      const key = file.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      let link;
      try { link = read(file); } catch { continue; }
      const change = repairOf(link, exe, { exists, folders });
      if (!change) continue;
      try { if (write(file, change)) changed.push({ file, from: link.target }); else failed += 1; } catch { failed += 1; }
    }
  }
  return { changed, failed };
}

/** The same, on this computer, with Electron's shortcut reader and writer. */
export function repairPinsHere({ shell, fs, exe, appData, desktop, env }) {
  return repairPins({
    exe,
    folders: installerFolders(env),
    places: shortcutPlaces({ appData, desktop }),
    list: (dir) => fs.readdirSync(dir, { withFileTypes: true }).map((e) => ({ name: e.name, isDirectory: e.isDirectory() })),
    read: (file) => shell.readShortcutLink(file),
    write: (file, change) => shell.writeShortcutLink(file, 'update', change),
    exists: (file) => fs.existsSync(file),
  });
}
