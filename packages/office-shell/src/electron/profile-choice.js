// Which profile this run of Rutba Office is (profiles.js), settled before
// anything asks Electron for the profile folder.
//
// The first app.getPath('userData') fixes the folder for the whole run, and
// asked before the app was named it fixed it under the package's own name
// (2.15.0 opened an empty profile that way; template-set-choice.js). So the
// folder is set here, explicitly, from the list: the command line's
// `--profile`, else for a file or link handed over the profile used last,
// else the one marked to open first. Chromium keeps its own data in the
// same folder, and Electron's one-instance lock is taken on it, so each
// profile is its own running Rutba Office and two can run side by side.
//
// The base folder is the one a check or a developer names with
// --user-data-dir, else %APPDATA%\Rutba Office as it has always been.

import fs from 'node:fs';
import path from 'node:path';
import { app } from 'electron';
import { chooseProfile, profileFolder, profileArg, DEFAULT_ID } from '../profiles.js';
import { parseLink } from '../links.js';

/** Whether the command line hands Office a file or a rutba-office:// link to open. */
function handsSomething(argv) {
  const skip = profileArg(argv).at;
  return argv.slice(1).some((a, i) => {
    if (i + 1 === skip) return false;
    if (parseLink(a)) return true;
    if (String(a).startsWith('-')) return false;
    try {
      return fs.statSync(a).isFile();
    } catch {
      return false;
    }
  });
}

let settled = null;

/**
 * The profile this run is, worked out once: { base, id, name, colour, folder,
 * lock, count, isDefault, baseNamed (the base was named on the command line) }. Sets the
 * profile's folder as Electron's userData and sessionData before anything
 * else reads them.
 */
export function settleProfile({ name = 'Rutba Office' } = {}) {
  if (settled) return settled;
  const named = app.commandLine.getSwitchValue('user-data-dir');
  const base = named || path.join(app.getPath('appData'), name);
  const { profile, list } = chooseProfile({ base, argv: process.argv, handed: handsSomething(process.argv) });
  const folder = profileFolder(base, profile);
  try {
    fs.mkdirSync(folder, { recursive: true });
  } catch {
    // Electron makes it, or says why it cannot.
  }
  app.setPath('userData', folder);
  app.setPath('sessionData', folder);
  settled = { base, baseNamed: Boolean(named), id: profile.id, name: profile.name, colour: profile.colour, folder, lock: profile.lock || null, rekey: profile.rekey || null, count: list.profiles.length, isDefault: profile.id === DEFAULT_ID };
  return settled;
}

/** The profile settled at the start (settleProfile), or null before it is. */
export const settledProfile = () => settled;
