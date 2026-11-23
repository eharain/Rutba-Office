// The licences of what the window code is built from.
//
// The packages the program loads as they are keep their licence files inside
// its archive (electron-builder.config.cjs `files`). The ones built into the
// window code do not: esbuild writes one file out of many and leaves the
// licence comments out (bundle.js `legalComments`), so nothing of React's or
// LiveKit's notice would reach a person who installs a copy. This writes them
// out as one text file the installers put beside LICENSE.txt.
//
// Which packages: every one a build read a file of (esbuild's own record of
// its inputs), and the packages each of those depends on, since a package may
// come ready-made with its dependencies already inside it (LiveKit's client
// does). A package named here that the build left out costs a paragraph; a
// package built in and not named is a licence not kept.
//
// A package with no licence file stops the build, unless its text is kept in
// build/notice-texts/<name>@<version>.txt: for that version only, so a newer
// one is looked at again.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
export const KEPT_TEXTS = path.join(here, 'notice-texts');

const LICENCE_FILE = /^(licen[cs]e|copying|notice)([.-].*)?$/i;
const posix = (p) => p.split(path.sep).join('/');

/** The folder of the installed package a file belongs to, or null for a file of our own. */
export function packageDirOf(file) {
  const parts = posix(path.resolve(file)).split('/');
  const at = parts.lastIndexOf('node_modules');
  if (at < 0 || at + 1 >= parts.length - 1) return null;
  const scoped = parts[at + 1].startsWith('@');
  if (scoped && at + 2 >= parts.length - 1) return null;
  return parts.slice(0, at + (scoped ? 3 : 2)).join('/');
}

/** Where `name` is installed for a package in `from`, as Node would find it, or null. */
function whereInstalled(from, name) {
  for (let dir = from; ; dir = path.dirname(dir)) {
    if (path.basename(dir) !== 'node_modules') {
      const at = path.join(dir, 'node_modules', name);
      if (fs.existsSync(path.join(at, 'package.json'))) return at;
    }
    if (path.dirname(dir) === dir) return null;
  }
}

/**
 * The packages in the folders given and all they depend on: their
 * `dependencies`, not the optional ones (built in only if installed and
 * imported, and then an input of their own) nor their peers (the same).
 * Answers [{ name, version, license, dir }], one per name and version, sorted.
 */
export function packagesFrom(dirs) {
  const seen = new Set();
  const found = new Map();
  const visit = (dir) => {
    const real = posix(fs.realpathSync(dir));
    if (seen.has(real)) return;
    seen.add(real);
    const pkg = JSON.parse(fs.readFileSync(path.join(real, 'package.json'), 'utf8'));
    const id = `${pkg.name}@${pkg.version}`;
    if (!found.has(id)) found.set(id, { name: pkg.name, version: pkg.version, license: typeof pkg.license === 'string' ? pkg.license : pkg.license?.type || '', dir: real });
    for (const dep of Object.keys(pkg.dependencies || {})) {
      const at = whereInstalled(real, dep);
      if (!at) throw new Error(`${id} depends on ${dep}, which is not installed: its licence cannot be read`);
      visit(at);
    }
  };
  for (const dir of dirs) visit(dir);
  return [...found.values()].sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : a.version < b.version ? -1 : a.version > b.version ? 1 : 0));
}

const tidy = (text) => text.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n').replace(/[ \t]+$/gm, '').replace(/^\n+|\n+$/g, '');

/** The licence text of one package: the files it carries, or the text kept for this version. */
export function licenceTextOf(pkg, kept = KEPT_TEXTS) {
  const files = fs.readdirSync(pkg.dir, { withFileTypes: true }).filter((e) => e.isFile() && LICENCE_FILE.test(e.name)).map((e) => e.name).sort();
  const texts = files.map((f) => tidy(fs.readFileSync(path.join(pkg.dir, f), 'utf8'))).filter(Boolean);
  if (texts.length) return texts.join('\n\n');
  const own = path.join(kept, ...`${pkg.name}@${pkg.version}.txt`.split('/'));
  if (fs.existsSync(own)) {
    const text = tidy(fs.readFileSync(own, 'utf8'));
    if (text) return text;
  }
  throw new Error(`${pkg.name} ${pkg.version} is built into the window code and carries no licence file: put its licence text in ${posix(path.relative(path.resolve(here, '..', '..', '..'), own))}`);
}

const RULE = '='.repeat(78);
const HEAD = `Rutba Office: third-party notices

The window code of Rutba Office is built from the packages below and from the
packages built into them. Each is given with its version, its licence as the
package states it, and the licence text it carries.

Beside this file: LICENSE.txt is Rutba Office's own licence,
LICENSE.electron.txt is Electron's and LICENSES.chromium.html is Chromium's.
The packages the program loads as they are keep their licence files inside
the program's archive, and the bill of materials published with each release
lists every part with its version, licence and hash.`;

/** The notices file's text for the packages given (packagesFrom). The same packages give the same text. */
export function noticesText(packages, kept = KEPT_TEXTS) {
  const parts = packages.map((p) => `${RULE}\n${p.name} ${p.version}\nLicence: ${p.license || 'not stated in its package.json'}\n${RULE}\n\n${licenceTextOf(p, kept)}`);
  return `${[HEAD, ...parts].join('\n\n\n')}\n`;
}

/**
 * The notices for a build, from esbuild's records of what each bundle read
 * (`metafile`, its paths from the folder the build ran in).
 * @param {{ inputs: Record<string, unknown> }[]} metafiles
 * @param {{ cwd?: string, kept?: string }} [how]
 */
export function noticesFor(metafiles, { cwd = process.cwd(), kept = KEPT_TEXTS } = {}) {
  const dirs = new Set();
  for (const m of metafiles) for (const input of Object.keys(m.inputs || {})) {
    // A made-up input (esbuild's own helpers, a data: file) is nobody's package.
    if (/^[\w-]+:/.test(input) && !/^[A-Za-z]:[\\/]/.test(input)) continue;
    const dir = packageDirOf(path.resolve(cwd, input));
    if (dir) dirs.add(dir);
  }
  return noticesText(packagesFrom([...dirs].sort()), kept);
}
