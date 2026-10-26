// Which profile folder the AI tools work in (node.js), apart so tests can ask
// it without starting the tools.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { readProfiles, findProfile, profileFolder, profileArg } from '@rutba/office-shell/profiles';

/** The folder Rutba Office keeps its profiles in, as Electron names it for the app. */
export function baseDir(env = process.env) {
  if (process.platform === 'win32') return path.join(env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming'), 'Rutba Office');
  if (process.platform === 'darwin') return path.join(os.homedir(), 'Library', 'Application Support', 'Rutba Office');
  return path.join(env.XDG_CONFIG_HOME || path.join(os.homedir(), '.config'), 'Rutba Office');
}

/**
 * The profile folder the tools work in, where its settings are: the one
 * RUTBA_OFFICE_PROFILE names; else the profile `--profile <id or name>`
 * names (D-2026-10-26-02: the AI Tools window puts the profile it was opened
 * in on the line); else the first profile, Default, which is the base folder
 * as before profiles. A profile that is not there is an error, so an AI app
 * never works in the wrong one.
 */
export function profileDir({ argv = process.argv.slice(2), env = process.env } = {}) {
  if (env.RUTBA_OFFICE_PROFILE) return env.RUTBA_OFFICE_PROFILE;
  const base = baseDir(env);
  const asked = profileArg(argv).value;
  if (!asked) return base;
  const profile = findProfile(readProfiles(base), asked);
  if (!profile) throw new Error(`Rutba Office has no profile "${asked}". Open AI Tools in the profile you want and copy its lines again.`);
  return profileFolder(base, profile);
}

/**
 * The entry of the profile whose folder this is, from its base folder's
 * list: Default's for a base folder, one under Profiles for the rest; null
 * when the folder is not a profile's. Read afresh for each call, so a
 * password set while the tools run counts at once.
 */
export function profileEntryOf(folder) {
  const full = path.resolve(folder);
  if (fs.existsSync(path.join(full, 'profiles.json')) || !/^profiles$/i.test(path.basename(path.dirname(full)))) {
    return findProfile(readProfiles(full), 'default');
  }
  return findProfile(readProfiles(path.dirname(path.dirname(full))), path.basename(full));
}
