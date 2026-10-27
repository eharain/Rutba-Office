// The AI tools under an AI app run as a Windows app package (apps/desktop/
// main/mcp/private-copy.js): a private copy of the profile that Windows keeps
// for the package hid the real settings, so every call was refused while AI
// Tools was on (2026-10-27). The copy is found by a file written in the
// profile folder appearing in it, and moved aside whole, nothing deleted.
// Windows' redirection is stood in for by where the test writes the marker.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { findPrivateCopy, setAsidePrivateCopy, privateCopyNotice } from '../apps/desktop/main/mcp/private-copy.js';

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-private-copy-'));
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));

let n = 0;
/** A computer: a real profile, and a packages folder with an app's copy of it, or none. */
function computer({ copy = true } = {}) {
  const root = path.join(tmp, `pc-${++n}`);
  const env = { APPDATA: path.join(root, 'Roaming'), LOCALAPPDATA: path.join(root, 'Local') };
  const folder = path.join(env.APPDATA, 'Rutba Office');
  fs.mkdirSync(folder, { recursive: true });
  fs.writeFileSync(path.join(folder, 'settings.json'), JSON.stringify({ 'apps.choice': { addons: ['ai'] }, 'ai.tools.mode': 'always' }));
  const packages = path.join(env.LOCALAPPDATA, 'Packages');
  fs.mkdirSync(path.join(packages, 'Other.App_1'), { recursive: true });
  const privateDir = path.join(packages, 'An.AiApp_abc', 'LocalCache', 'Roaming', 'Rutba Office');
  if (copy) {
    fs.mkdirSync(privateDir, { recursive: true });
    fs.writeFileSync(path.join(privateDir, 'settings.json'), JSON.stringify({ 'apps.choice': { addons: [] } }));
  }
  // Inside the package a file written in a folder its copy has lands in the copy.
  const writeMarker = (file) => fs.writeFileSync(copy ? path.join(privateDir, path.basename(file)) : file, '');
  return { env, folder, packages, privateDir, writeMarker };
}

test('a private copy of the profile is found by where a file written in the folder lands, and the marker is gone', () => {
  const pc = computer();
  const found = findPrivateCopy({ folder: pc.folder, env: pc.env, packages: pc.packages, writeMarker: pc.writeMarker });
  assert.deepEqual(found, { package: 'An.AiApp_abc', dir: pc.privateDir });
  assert.deepEqual(fs.readdirSync(pc.privateDir), ['settings.json'], 'the marker is taken away again');
  assert.deepEqual(fs.readdirSync(pc.folder), ['settings.json']);
});

test('with no copy, or outside a package, nothing is found and nothing is left behind', () => {
  const pc = computer({ copy: false });
  assert.equal(findPrivateCopy({ folder: pc.folder, env: pc.env, packages: pc.packages, writeMarker: pc.writeMarker }), null);
  assert.deepEqual(fs.readdirSync(pc.folder), ['settings.json']);
  assert.equal(findPrivateCopy({ folder: pc.folder, env: pc.env, packages: path.join(tmp, 'no-packages'), writeMarker: pc.writeMarker }), null);
  // A profile named outside AppData (RUTBA_OFFICE_PROFILE, a check run) is never looked for.
  assert.equal(findPrivateCopy({ folder: path.join(tmp, 'elsewhere'), env: pc.env, packages: pc.packages, writeMarker: () => assert.fail('no marker') }), null);
});

test('the copy is moved aside whole, beside itself, so the real settings are read, and the notice says where it went', () => {
  const pc = computer();
  const done = setAsidePrivateCopy({ folder: pc.folder, env: pc.env, packages: pc.packages, writeMarker: pc.writeMarker, platform: 'win32' });
  assert.equal(done.package, 'An.AiApp_abc');
  assert.equal(fs.existsSync(pc.privateDir), false);
  assert.ok(done.aside.startsWith(`${pc.privateDir}.set-aside-`));
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(done.aside, 'settings.json'), 'utf8')), { 'apps.choice': { addons: [] } }, 'kept, unchanged');
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(pc.folder, 'settings.json'), 'utf8'))['apps.choice'], { addons: ['ai'] }, 'the real profile untouched');
  assert.match(privateCopyNotice(done), /^Windows kept a private copy of this Rutba Office profile for An\.AiApp_abc, the app these tools run in, which hid the settings made in Rutba Office\. It was put aside, unchanged, in .+, and the tools now read the real profile\.$/);
  assert.equal(privateCopyNotice(null), null);
  // Only on Windows.
  assert.equal(setAsidePrivateCopy({ folder: pc.folder, env: pc.env, packages: pc.packages, writeMarker: () => assert.fail('no marker'), platform: 'linux' }), null);
});

test('a copy that cannot be moved is said plainly, and nothing is changed', () => {
  const pc = computer();
  const held = privateCopyNotice({ package: 'An.AiApp_abc', dir: pc.privateDir, held: 'EBUSY' });
  assert.match(held, /could not be put aside \(EBUSY\): the tools may not see the settings made in Rutba Office\. Close any Rutba Office started from that app and start the tools again\.$/);
});
