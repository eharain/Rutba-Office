// A private copy of the profile, kept by Windows for an AI app that runs as
// a Windows app package (MSIX), set aside before the tools read anything.
//
// Windows gives a packaged app's processes, and every program they start,
// their own copy of AppData: a file such a process writes in a folder that is
// in that copy lands under %LOCALAPPDATA%\Packages\<package>\LocalCache\
// Roaming, and from then on hides the real file from it. An AI app run that
// way starts the tools inside its package, so once a copy of the profile
// folder is there (a Rutba Office once started from inside the package, or an
// earlier run of the tools), the tools read that copy's settings.json for
// ever, and not the settings the person makes in Rutba Office: AI Tools
// turned on and its folders given were never seen (2026-10-27).
//
// So at start the tools find out whether their profile folder has such a
// copy (a file written in the folder that appears in a package's copy), and
// if so move the copy aside, whole, beside itself: nothing in it is deleted,
// and the tools then read and write the real profile.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';

/** The packages folder Windows keeps packaged apps' copies in. */
const packagesDir = (env) => path.join(env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local'), 'Packages');

/**
 * The folder in a package's copy of AppData that stands for `folder`, when
 * this process writes there rather than in the folder itself; else null.
 *
 * @param {object} o
 * @param {string} o.folder    the profile folder, as the tools name it
 * @param {object} [o.env]
 * @param {string} [o.packages] the packages folder (tests)
 * @param {Function} [o.writeMarker] writes the marker file (tests stand in for Windows' redirection)
 */
export function findPrivateCopy({ folder, env = process.env, packages = packagesDir(env), writeMarker = (file) => fs.writeFileSync(file, '') }) {
  const roaming = env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
  const rel = path.relative(roaming, folder);
  if (!rel || rel.startsWith('..') || path.isAbsolute(rel)) return null;
  const name = `.rutba-where-${process.pid}-${crypto.randomBytes(4).toString('hex')}`;
  const marker = path.join(folder, name);
  try {
    writeMarker(marker);
  } catch {
    return null;
  }
  let found = null;
  try {
    for (const pkg of fs.readdirSync(packages)) {
      const copy = path.join(packages, pkg, 'LocalCache', 'Roaming', rel);
      if (fs.existsSync(path.join(copy, name))) {
        found = { package: pkg, dir: copy };
        fs.rmSync(path.join(copy, name), { force: true });
        break;
      }
    }
  } catch {
    // No packages folder: not a packaged app's process.
  }
  fs.rmSync(marker, { force: true });
  return found;
}

/**
 * The private copy moved aside, whole, beside itself, so the folder reads
 * through to the real profile again. Answers what was done, or null when
 * there was no copy; a copy that cannot be moved (a Rutba Office running
 * inside the package holds it) is answered with `held`.
 */
export function setAsidePrivateCopy(o) {
  if ((o.platform || process.platform) !== 'win32') return null;
  const found = findPrivateCopy(o);
  if (!found) return null;
  const aside = `${found.dir}.set-aside-${new Date().toISOString().replace(/[:.]/g, '-')}`;
  try {
    fs.renameSync(found.dir, aside);
    return { ...found, aside };
  } catch (err) {
    return { ...found, held: err.code || 'failed' };
  }
}

/** What the tools say of it, on standard error, for the AI app's log. */
export function privateCopyNotice(done) {
  if (!done) return null;
  if (done.held) return `Windows keeps a private copy of this Rutba Office profile for ${done.package}, the app these tools run in, and it could not be put aside (${done.held}): the tools may not see the settings made in Rutba Office. Close any Rutba Office started from that app and start the tools again.`;
  return `Windows kept a private copy of this Rutba Office profile for ${done.package}, the app these tools run in, which hid the settings made in Rutba Office. It was put aside, unchanged, in ${done.aside}, and the tools now read the real profile.`;
}
