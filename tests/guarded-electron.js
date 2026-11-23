// A real Electron started by a test, with no error box able to reach the
// screen of the person at this computer.
//
// Electron starts on the guard's few lines, which turn an error box and any
// fault left uncaught into a line of output and an exit. Only then is the
// test's own probe read, so a probe that cannot be read, or that fails, is a
// failed test and nothing more. Both files are proved to read as scripts
// before Electron is started on them. The guard asks Electron for no folder,
// so a probe finds the profile named as it would be without it.
//
// A probe opens no window and ends with app.exit(): what it found goes to
// its output, which `runGuarded` hands back.

import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

export const GUARD = `const { app, dialog } = require('electron');
const fail = (err) => { console.error('FAULT=' + ((err && err.stack) || err)); app.exit(1); };
dialog.showErrorBox = (title, content) => fail(new Error(title + ': ' + content));
process.on('uncaughtException', fail);
process.on('unhandledRejection', fail);
import('./probe.mjs').catch(fail);
`;

/** The Electron the desktop app is built on. */
export const electronPath = () => createRequire(path.join(root, 'apps', 'desktop', 'package.json'))('electron');

/** The probe written into `dir` behind the guard, as an application named `name`. */
export function writeGuarded(dir, probe, { name = '@rutba/office-desktop' } = {}) {
  fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ name, type: 'module', main: 'main.cjs' }));
  fs.writeFileSync(path.join(dir, 'main.cjs'), GUARD);
  fs.writeFileSync(path.join(dir, 'probe.mjs'), probe);
  for (const file of ['main.cjs', 'probe.mjs']) execFileSync(process.execPath, ['--check', path.join(dir, file)], { encoding: 'utf8', timeout: 60000 });
}

/** Electron started on the guarded probe in `dir`; its output. A probe that fails throws, with what it said. */
export function runGuarded(dir, args = [], { env = process.env, timeout = 60000 } = {}) {
  const clean = { ...env };
  delete clean.ELECTRON_RUN_AS_NODE;
  return execFileSync(electronPath(), [dir, ...args], { env: clean, encoding: 'utf8', timeout });
}
