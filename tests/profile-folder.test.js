// The profile stays where Office keeps it. 2.15.0's first import
// (template-set-choice.js) asked Electron for the profile folder before the
// shell named the app "Rutba Office", which fixed the profile under the
// package's own name: Office opened an empty profile, every setting and mail
// account seemingly gone. Here the real module runs first in a real Electron,
// as main.js runs it, and the profile must still be "Rutba Office".

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const choice = path.join(root, 'apps', 'desktop', 'main', 'template-set-choice.js');

test('the first import leaves the profile where Office names it: Rutba Office', () => {
  const electron = createRequire(path.join(root, 'apps', 'desktop', 'package.json'))('electron');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'profile-'));
  try {
    // Named as the package is, as a packaged Office is before the shell names it.
    fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ name: '@rutba/office-desktop', type: 'module', main: 'main.mjs' }));
    fs.writeFileSync(path.join(dir, 'main.mjs'), `import { app } from 'electron';
await import(${JSON.stringify(pathToFileURL(choice).href)});
app.setName('Rutba Office');
console.log('PROFILE=' + app.getPath('userData'));
app.exit(0);
`);
    const env = { ...process.env };
    delete env.RUTBA_TEMPLATES_DIR;
    delete env.ELECTRON_RUN_AS_NODE;
    const out = execFileSync(electron, [dir], { env, encoding: 'utf8', timeout: 60000 });
    const profile = /PROFILE=(.*)/.exec(out)?.[1]?.trim();
    assert.ok(profile, out);
    // The base folder, or a profile made in it (profiles.js) when one is marked to open first.
    assert.match(profile, /[\\/]Rutba Office([\\/]Profiles[\\/][0-9a-f]{8})?$/, profile);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

// Profiles (D-2026-10-26-02): the run is the profile `--profile` names, and
// its folder, under the base folder's Profiles, is Electron's before anything
// asks; a folder named with --user-data-dir is the base, as the checks name it.
test('the profile named on the command line is the folder of the run, for Electron and Chromium both', () => {
  const electron = createRequire(path.join(root, 'apps', 'desktop', 'package.json'))('electron');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'profile-'));
  const base = fs.mkdtempSync(path.join(os.tmpdir(), 'profile-base-'));
  try {
    fs.writeFileSync(path.join(base, 'profiles.json'), JSON.stringify({ version: 1, openFirst: 'default', lastUsed: 'default', profiles: [{ id: 'default', name: 'Default' }, { id: 'a1b2c3d4', name: 'Work', colour: '#1a9f7a' }] }));
    fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ name: '@rutba/office-desktop', type: 'module', main: 'main.mjs' }));
    fs.writeFileSync(path.join(dir, 'main.mjs'), `import { app } from 'electron';
const { PROFILE } = await import(${JSON.stringify(pathToFileURL(choice).href)});
app.setName('Rutba Office');
console.log('PROFILE=' + JSON.stringify({ userData: app.getPath('userData'), sessionData: app.getPath('sessionData'), id: PROFILE.id, base: PROFILE.base, count: PROFILE.count }));
app.exit(0);
`);
    const env = { ...process.env };
    delete env.RUTBA_TEMPLATES_DIR;
    delete env.ELECTRON_RUN_AS_NODE;
    const run = (...args) => JSON.parse(/PROFILE=(.*)/.exec(execFileSync(electron, [dir, `--user-data-dir=${base}`, ...args], { env, encoding: 'utf8', timeout: 60000 }))[1]);
    const work = run('--profile', 'work');
    assert.equal(work.id, 'a1b2c3d4');
    assert.equal(work.count, 2);
    assert.equal(path.resolve(work.userData), path.join(path.resolve(base), 'Profiles', 'a1b2c3d4'));
    assert.equal(path.resolve(work.sessionData), path.resolve(work.userData));
    const first = run();
    assert.equal(first.id, 'default');
    assert.equal(path.resolve(first.userData), path.resolve(base), 'Default is the base folder, as before profiles');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
    fs.rmSync(base, { recursive: true, force: true });
  }
});
