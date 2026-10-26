// Which template set this run uses, decided before anything reads templates:
// imported first by main.js, so the gallery's and Data Files' entries
// (@rutba/templates/active and data-files/active) and the document worker
// all read the same set. A newer set Office downloaded and verified
// (template-sets.js) is used when this Office is new enough for it; else
// the set Office was built with. RUTBA_TEMPLATES_DIR names the folder.

import fs from 'node:fs';
import path from 'node:path';
import { app } from 'electron';
import { compareVersions, usableBy } from '@rutba/templates/versions';
import { readManifest, MANIFEST_PATH } from '@rutba/templates/manifest';
import { settleProfile } from '@rutba/office-shell/electron/profile-choice';

/**
 * The profile, settled first of all (profile-choice.js): which one this run
 * is, and its folder set as Electron's before anything asks for it. This
 * module runs before the shell names the app "Rutba Office", and the first
 * app.getPath('userData') fixes the profile where the name then says. Asked
 * that early it fixed it under the package's own name, and 2.15.0 opened an
 * empty profile there, every setting and mail account seemingly gone. A
 * folder named on the command line (the checks' own) is the base folder.
 */
export const PROFILE_NAME = 'Rutba Office';
export const PROFILE = settleProfile({ name: PROFILE_NAME });

/** The folder where downloaded sets are kept: the base folder's, shared by every profile. */
export const SETS_DIR = () => path.join(PROFILE.base, 'templates');

/** The set Office was built with, read from this package's own manifest. */
export const bundledManifest = () => readManifest(MANIFEST_PATH);

function choose() {
  // A check stands in a set of its own, or none.
  if (process.env.RUTBA_TEMPLATES_DIR !== undefined) return process.env.RUTBA_TEMPLATES_DIR || null;
  try {
    const root = SETS_DIR();
    const { set } = JSON.parse(fs.readFileSync(path.join(root, 'current.json'), 'utf8'));
    const dir = path.join(root, set);
    const m = readManifest(path.join(dir, 'manifest.json'));
    const office = process.env.RUTBA_OFFICE_VERSION || app.getVersion();
    if (m.set === set && compareVersions(m.set, bundledManifest().set) > 0 && usableBy(office, m.minOffice) && fs.existsSync(path.join(dir, 'data-files', 'index.js'))) return dir;
  } catch {
    // No set downloaded, or one this Office cannot use: the one it was built with.
  }
  return null;
}

const chosen = choose();
if (chosen) process.env.RUTBA_TEMPLATES_DIR = chosen;
else delete process.env.RUTBA_TEMPLATES_DIR;

/** The folder of the downloaded set in use, or null for the bundled one. */
export const CHOSEN_SET_DIR = chosen;
