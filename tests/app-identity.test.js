// Who Rutba Office is to Windows (office-shell app-identity.js, B18): only
// the person's installed copy hands Windows a notification. Electron writes
// to the person's Start menu and registry the first time one is asked of it,
// and a copy run from the source, a portable copy or a built copy never
// installed must not: the first left an Electron shortcut carrying Rutba
// Office's id, the others would turn the installed copy's own shortcut
// towards themselves.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { appIdentity, APP_ID, TOAST_ACTIVATOR, TOASTS_SWITCH } from '../packages/office-shell/src/electron/app-identity.js';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (...parts) => fs.readFileSync(path.join(root, ...parts), 'utf8');

test('the installed copy keeps the application\'s id and one press-back id of its own', () => {
  const id = appIdentity({ packaged: true, installed: true, platform: 'win32', env: {} });
  assert.deepEqual(id, { appUserModelId: 'co.techstyle.rutba.office', toastActivatorClsid: TOAST_ACTIVATOR, toasts: true });
  assert.equal(APP_ID, 'co.techstyle.rutba.office');
  // A class id as Windows writes one: Electron takes no other shape.
  assert.match(TOAST_ACTIVATOR, /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/);
  // The switch is for work on the source; it changes nothing for the installed copy.
  assert.deepEqual(appIdentity({ packaged: true, installed: true, platform: 'win32', env: { [TOASTS_SWITCH]: '1' } }), id);
});

test('a copy run from the source, a portable copy and a built copy never installed hand Windows none', () => {
  for (const [what, run] of [
    ['run from the source', { packaged: false, installed: false }],
    ['portable, or built and never installed', { packaged: true, installed: false }],
  ]) {
    const id = appIdentity({ ...run, platform: 'win32', env: {} });
    assert.deepEqual(id, { appUserModelId: APP_ID, toastActivatorClsid: null, toasts: false }, what);
  }
  // A packaged copy is never talked into it: it carries the installed copy's product name.
  assert.equal(appIdentity({ packaged: true, installed: false, platform: 'win32', env: { [TOASTS_SWITCH]: '1' } }).toasts, false);
});

test('a copy run from the source shows them when told to, under an id of its own', () => {
  const id = appIdentity({ packaged: false, installed: false, platform: 'win32', env: { [TOASTS_SWITCH]: '1' } });
  assert.deepEqual(id, { appUserModelId: 'co.techstyle.rutba.office.dev', toastActivatorClsid: null, toasts: true });
  for (const value of ['', '0', 'true', 'yes', undefined]) {
    assert.equal(appIdentity({ packaged: false, installed: false, platform: 'win32', env: { [TOASTS_SWITCH]: value } }).toasts, false, `"${value}" is not the switch`);
  }
  assert.equal(appIdentity({ packaged: false, installed: false, platform: 'win32' }).toasts, false, 'no environment handed in');
});

test('anywhere but Windows every copy shows them, as it always did', () => {
  for (const platform of ['darwin', 'linux']) {
    for (const run of [{ packaged: false, installed: false }, { packaged: true, installed: false }]) {
      assert.deepEqual(appIdentity({ ...run, platform, env: {} }), { appUserModelId: APP_ID, toastActivatorClsid: null, toasts: true });
    }
  }
});

test('the shell and the desktop take their say from it, and the installer writes the same id', () => {
  const shell = read('packages', 'office-shell', 'src', 'electron', 'main.js');
  assert.ok(shell.includes('const identity = appIdentity({'), 'the shell asks');
  assert.ok(shell.includes('app.setAppUserModelId(appUserModelId);'));
  assert.ok(shell.includes('app.setToastActivatorCLSID?.(identity.toastActivatorClsid);'), 'the press-back id is handed to Electron');
  assert.ok(!/setAppUserModelId\('co\./.test(shell), 'the id is written in one place only');
  const desktop = read('apps', 'desktop', 'main', 'main.js');
  assert.ok(desktop.includes('toasts: identity.toasts'), 'the notifier\'s parts are told whether this copy shows any');
  assert.ok(desktop.includes('readToastBlock({ appId: identity.appUserModelId })'), 'Windows is asked about the id this copy goes under');
  const exported = JSON.parse(read('packages', 'office-shell', 'package.json')).exports['./electron/app-identity'];
  assert.ok(exported && fs.existsSync(path.join(root, 'packages', 'office-shell', exported)));
  // electron-builder writes the installer's shortcuts under the build's appId.
  assert.ok(read('apps', 'desktop', 'electron-builder.config.cjs').includes(`appId: '${APP_ID}',`), 'the build carries the same id');
});
