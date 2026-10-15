/**
 * The Linux build (.github/workflows/linux.yml) can only run on Linux, so
 * what can be held here is held: the installer configuration passes
 * electron-builder's own schema, names the program and the Debian package
 * as the workflow starts and installs them, and the workflow builds both
 * packages, starts each under a virtual display and publishes nothing.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(path.join(root, 'apps', 'desktop', 'package.json'));
// A Windows checkout may give it CRLF line endings.
const workflow = fs.readFileSync(path.join(root, '.github', 'workflows', 'linux.yml'), 'utf8').replace(/\r\n/g, '\n');

test('the installer configuration passes electron-builder\'s schema, with the Linux names the workflow uses', async () => {
  const config = await require('./electron-builder.config.cjs')();
  const { validateConfiguration } = require('app-builder-lib/out/util/config/config');
  const { DebugLogger } = require('builder-util');
  await validateConfiguration(config, new DebugLogger(false));
  assert.deepEqual(config.linux.target.map((t) => t.target), ['AppImage', 'deb']);
  assert.equal(config.linux.executableName, 'rutba-office');
  assert.equal(config.deb.packageName, 'rutba-office');
  assert.equal(config.linux.desktop.entry.StartupWMClass, config.linux.executableName, 'the window is matched to its launcher by the program\'s name');
  assert.match(workflow, /linux-unpacked\/rutba-office /);
  assert.match(workflow, /"\/opt\/Rutba Office\/rutba-office"/, 'where the .deb installs it: /opt/<productName>/<executableName>');
  assert.match(workflow, /dpkg -s rutba-office/);
  assert.match(workflow, /\/usr\/share\/applications\/rutba-office\.desktop/);
});

test('the workflow runs by hand only, builds both packages, starts each under Xvfb and publishes nothing', () => {
  assert.match(workflow, /^on:\n  workflow_dispatch:\n\n/m);
  assert.doesNotMatch(workflow, /push:|workflow_run:|release:/);
  assert.match(workflow, /npm run dist:linux/);
  assert.equal((workflow.match(/xvfb-run -a/g) || []).length, 3);
  assert.match(workflow, /RUTBA_OFFICE_SMOKE: '1'/);
  assert.doesNotMatch(workflow, /publish-release|contents: write|GH_TOKEN/);
});
