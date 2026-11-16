// Only the person's installed copy takes the rutba-office:// links and their
// pins: a built copy run as a test took the links on the owner's computer
// (16 Nov), so a sign-in would have come back to the test build.

import test from 'node:test';
import assert from 'node:assert/strict';
import { isInstalledCopy } from '../packages/office-shell/src/electron/installed-copy.js';

const INSTALLED = 'C:\\Users\\a\\AppData\\Local\\Programs\\Rutba Office\\Rutba Office.exe';
const BUILT = 'D:\\office\\apps\\desktop\\release\\win-unpacked\\Rutba Office.exe';
const exists = (f) => f.toLowerCase() === 'c:\\users\\a\\appdata\\local\\programs\\rutba office\\uninstall rutba office.exe';
const run = (over = {}) => ({ isPackaged: true, platform: 'win32', exe: INSTALLED, argv: [INSTALLED, '--app=mail', '--profile=default'], env: {}, exists, ...over });

test('the installed copy, started by a pin, is the person\'s copy', () => {
  assert.equal(isInstalledCopy(run()), true);
});

test('a built copy never installed, with no uninstaller beside it, is not', () => {
  assert.equal(isInstalledCopy(run({ exe: BUILT, argv: [BUILT, '--app=mail'] })), false);
});

test('a test run of the installed copy is not: a scratch profile folder, or a check run\'s switch', () => {
  assert.equal(isInstalledCopy(run({ argv: [INSTALLED, '--app=mail', '--user-data-dir=D:\\scratch\\pin-1'] })), false);
  assert.equal(isInstalledCopy(run({ argv: [INSTALLED, '--user-data-dir', 'D:\\scratch'] })), false);
  for (const k of ['RUTBA_OFFICE_VERIFY_EDIT', 'RUTBA_OFFICE_VERIFY_APPS', 'RUTBA_OFFICE_VERIFY_CORPUS', 'RUTBA_OFFICE_SMOKE', 'RUTBA_SCREENS_FEATURES', 'RUTBA_PROFILE_PEER']) {
    assert.equal(isInstalledCopy(run({ env: { [k]: '1' } })), false, k);
  }
});

test('a portable copy, a copy run from source, and another system are not', () => {
  assert.equal(isInstalledCopy(run({ env: { PORTABLE_EXECUTABLE_DIR: 'E:\\' } })), false);
  assert.equal(isInstalledCopy(run({ isPackaged: false })), false);
  assert.equal(isInstalledCopy(run({ platform: 'darwin' })), false);
});
