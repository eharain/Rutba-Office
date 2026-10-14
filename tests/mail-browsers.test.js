// The browsers a sign-in can open in, read as Windows lists them, and how an
// account signs in, told from the provider table and the server's greeting.

import test from 'node:test';
import assert from 'node:assert/strict';
import { parseRegistryBrowsers, listBrowsers, openIn } from '../apps/desktop/main/browsers.js';
import { authOf } from '../apps/desktop/main/mail-discover.js';

const REG = [
  'HKEY_LOCAL_MACHINE\\SOFTWARE\\Clients\\StartMenuInternet',
  '    (Default)    REG_SZ    IEXPLORE.EXE',
  '',
  'HKEY_LOCAL_MACHINE\\SOFTWARE\\Clients\\StartMenuInternet\\Google Chrome',
  '    (Default)    REG_SZ    Google Chrome',
  '',
  'HKEY_LOCAL_MACHINE\\SOFTWARE\\Clients\\StartMenuInternet\\Google Chrome\\Capabilities\\URLAssociations',
  '    http    REG_SZ    ChromeHTML',
  '    https    REG_SZ    ChromeHTML',
  '',
  'HKEY_LOCAL_MACHINE\\SOFTWARE\\Clients\\StartMenuInternet\\Google Chrome\\shell\\open\\command',
  '    (Default)    REG_SZ    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"',
  '',
  'HKEY_LOCAL_MACHINE\\SOFTWARE\\Clients\\StartMenuInternet\\Microsoft Edge',
  '    (Default)    REG_SZ    Microsoft Edge',
  '',
  'HKEY_LOCAL_MACHINE\\SOFTWARE\\Clients\\StartMenuInternet\\Microsoft Edge\\Capabilities\\URLAssociations',
  '    https    REG_SZ    MSEdgeHTM',
  '',
  'HKEY_LOCAL_MACHINE\\SOFTWARE\\Clients\\StartMenuInternet\\Microsoft Edge\\shell\\open\\command',
  '    (Default)    REG_SZ    "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"',
  '',
  'HKEY_LOCAL_MACHINE\\SOFTWARE\\Clients\\StartMenuInternet\\IEXPLORE.EXE',
  '    (Default)    REG_SZ    Internet Explorer',
  '',
  'HKEY_LOCAL_MACHINE\\SOFTWARE\\Clients\\StartMenuInternet\\IEXPLORE.EXE\\shell\\open\\command',
  '    (Default)    REG_SZ    "C:\\Program Files\\Internet Explorer\\iexplore.exe"',
  '',
].join('\r\n');

test('the browsers Windows lists are read with the program each opens, Internet Explorer left out', () => {
  const list = parseRegistryBrowsers(REG);
  assert.deepEqual(list.map((b) => [b.id, b.name, b.exe, b.progId]), [
    ['Google Chrome', 'Google Chrome', 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe', 'ChromeHTML'],
    ['Microsoft Edge', 'Microsoft Edge', 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe', 'MSEdgeHTM'],
  ]);
});

test('the default browser is the one whose https handler the user chose, listed first', async () => {
  const reg = async (args) => {
    if (args.includes('/v')) return '\r\nHKEY_CURRENT_USER\\...\\UserChoice\r\n    ProgId    REG_SZ    MSEdgeHTM\r\n';
    return args[1].startsWith('HKLM\\SOFTWARE\\Clients') ? REG : '';
  };
  const list = await listBrowsers({ platform: 'win32', reg, exists: () => true });
  assert.deepEqual(list.map((b) => [b.name, b.isDefault]), [['Microsoft Edge', true], ['Google Chrome', false]]);
  // A program no longer on the disk is not offered.
  assert.equal((await listBrowsers({ platform: 'win32', reg, exists: (f) => !/chrome/.test(f) })).length, 1);
});

test('a sign-in page opens in the browser chosen, or the default, and only an https address', async () => {
  const opened = [];
  await openIn('https://accounts.example/sign-in', null, { openDefault: async (url) => opened.push(['default', url]) });
  await openIn('https://accounts.example/sign-in', 'gone', { openDefault: async (url) => opened.push(['default', url]), list: [] });
  assert.deepEqual(opened, [['default', 'https://accounts.example/sign-in'], ['default', 'https://accounts.example/sign-in']]);
  await assert.rejects(openIn('file:///C:/Windows/System32/calc.exe', null, { openDefault: async () => {} }), /secure web address/);
});

test('how an account signs in: its provider\'s page, an app password, or its password, and what the server offers', () => {
  assert.deepEqual(authOf({ label: 'Google', oauth: 'google' }, null), { kind: 'oauth', provider: 'Google', mechanisms: [] });
  assert.equal(authOf({ label: 'Yahoo', appPassword: { url: 'x' } }, null).kind, 'app-password');
  const dovecot = authOf(null, { banner: '* OK [CAPABILITY IMAP4rev1 SASL-IR AUTH=PLAIN AUTH=XOAUTH2] Dovecot ready.' });
  assert.deepEqual(dovecot, { kind: 'password', provider: null, mechanisms: ['PLAIN', 'XOAUTH2'], oauthOffered: true });
  assert.equal(authOf(null, { banner: '* OK ready', auth: 'password-cleartext' }).oauthOffered, false);
  assert.equal(authOf(null, { banner: '', auth: 'OAuth2' }).oauthOffered, true);
});
