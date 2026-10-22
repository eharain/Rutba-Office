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
    assert.equal(path.basename(profile), 'Rutba Office', profile);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
