// Whether Windows itself lets Rutba Office show a notification
// (apps/desktop/main/windows-toasts.js, B18). reg.exe is a stand-in here:
// what it would print is handed in, and nothing of the computer is read.

import test from 'node:test';
import assert from 'node:assert/strict';
import { dword, blockOf, readToastBlock, WINDOWS_PAGE } from '../apps/desktop/main/windows-toasts.js';

// What reg.exe prints for one value.
const printed = (key, name, hex) => `\r\n${key}\r\n    ${name}    REG_DWORD    0x${hex}\r\n\r\n`;
const ALL = 'HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\PushNotifications';
const OFF = printed(ALL, 'ToastEnabled', '0');
const ON = printed(ALL, 'ToastEnabled', '1');
const RULE = printed('HKEY_CURRENT_USER\\Software\\Policies\\Microsoft\\Windows\\CurrentVersion\\PushNotifications', 'NoToastApplicationNotification', '1');
const APP_OFF = printed('HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings\\co.techstyle.rutba.office', 'Enabled', '0');

test('a number is read from what reg.exe printed, and anything else is no number', () => {
  assert.equal(dword(OFF, 'ToastEnabled'), 0);
  assert.equal(dword(ON, 'ToastEnabled'), 1);
  assert.equal(dword(printed(ALL, 'ToastEnabled', 'ff'), 'ToastEnabled'), 255);
  assert.equal(dword(printed(ALL, 'toastenabled', '0'), 'ToastEnabled'), 0, 'Windows does not compare names by their case');
  assert.equal(dword(OFF, 'Enabled'), null, 'a name that only ends alike is another value');
  assert.equal(dword(`${ALL}\r\n    ToastEnabled    REG_SZ    0\r\n`, 'ToastEnabled'), null, 'text is no number');
  assert.equal(dword('', 'ToastEnabled'), null);
  assert.equal(dword(null, 'ToastEnabled'), null, 'reg.exe did not answer');
  assert.equal(dword(undefined, 'ToastEnabled'), null);
});

test('what stands in the way is named, the widest first, and only a plain "off" is one', () => {
  assert.equal(blockOf({}), null);
  assert.equal(blockOf(), null);
  assert.equal(blockOf({ all: ON }), null);
  assert.equal(blockOf({ all: OFF }), 'all');
  assert.equal(blockOf({ app: APP_OFF }), 'app');
  assert.equal(blockOf({ all: OFF, app: APP_OFF }), 'all');
  assert.equal(blockOf({ policyUser: RULE, all: OFF, app: APP_OFF }), 'policy');
  assert.equal(blockOf({ policyMachine: RULE }), 'policy');
  assert.equal(blockOf({ policyUser: RULE.replace('0x1', '0x0') }), null, 'a rule that is there and not set');
  assert.equal(blockOf({ app: APP_OFF.replace('0x0', '0x1') }), null);
});

test('the four places are asked, and the answer is what they hold', async () => {
  const asked = [];
  const reg = (holds) => async (args) => {
    asked.push(args);
    const found = Object.entries(holds).find(([key]) => args[1].toLowerCase() === key.toLowerCase());
    return found ? found[1] : null;
  };
  const POLICY = 'Software\\Policies\\Microsoft\\Windows\\CurrentVersion\\PushNotifications';
  const SWITCH = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\PushNotifications';
  const APP = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings\\co.techstyle.rutba.office';

  assert.equal(await readToastBlock({ appId: 'co.techstyle.rutba.office', platform: 'win32', reg: reg({ [SWITCH]: OFF }) }), 'all');
  assert.deepEqual(asked, [
    ['query', `HKCU\\${POLICY}`, '/v', 'NoToastApplicationNotification'],
    ['query', `HKLM\\${POLICY}`, '/v', 'NoToastApplicationNotification'],
    ['query', SWITCH, '/v', 'ToastEnabled'],
    ['query', APP, '/v', 'Enabled'],
  ], 'each is a read of one value, and nothing is written');

  assert.equal(await readToastBlock({ appId: 'co.techstyle.rutba.office', platform: 'win32', reg: reg({ [APP]: APP_OFF }) }), 'app');
  assert.equal(await readToastBlock({ appId: 'co.techstyle.rutba.office', platform: 'win32', reg: reg({ [`HKLM\\${POLICY}`]: RULE, [SWITCH]: OFF }) }), 'policy');
  assert.equal(await readToastBlock({ appId: 'co.techstyle.rutba.office', platform: 'win32', reg: reg({ [SWITCH]: ON }) }), null);
  assert.equal(await readToastBlock({ appId: 'co.techstyle.rutba.office', platform: 'win32', reg: reg({}) }), null, 'nothing kept there is no block');
});

test('a reg.exe that fails, an id that is no id, and any system but Windows are no block', async () => {
  const failing = async () => { throw new Error('no reg.exe'); };
  assert.equal(await readToastBlock({ appId: 'co.techstyle.rutba.office', platform: 'win32', reg: failing }), null);
  const thrown = () => { throw new Error('at once'); };
  assert.equal(await readToastBlock({ appId: 'co.techstyle.rutba.office', platform: 'win32', reg: thrown }), null);

  // An id with anything but letters, digits, dots and dashes is never put in a key's name.
  const asked = [];
  const reg = async (args) => { asked.push(args[1]); return null; };
  for (const appId of ['a\\..\\b', 'a b', 'a"b', '', null, undefined]) {
    asked.length = 0;
    assert.equal(await readToastBlock({ appId, platform: 'win32', reg }), null);
    assert.equal(asked.length, 3, `"${appId}" is not asked about`);
  }

  let calls = 0;
  for (const platform of ['darwin', 'linux']) assert.equal(await readToastBlock({ appId: 'co.techstyle.rutba.office', platform, reg: async () => { calls += 1; return OFF; } }), null);
  assert.equal(calls, 0);
});

test('Windows\' own page is the one Windows documents', () => {
  assert.equal(WINDOWS_PAGE, 'ms-settings:notifications');
});
