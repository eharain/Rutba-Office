// The browsers a sign-in can open in, read as Windows lists them, and how an
// account signs in, told from the provider table and the server's greeting.

import test from 'node:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { parseRegistryBrowsers, listBrowsers, readBrowsers, browserList, regRunner, openIn } from '../apps/desktop/main/browsers.js';
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

// A stand-in for reg.exe, run as a program of its own the way reg.exe is: it
// answers as Windows would, except while the computer is "busy" (a file
// beside it), when it hangs as reg.exe did beside a gate's other lanes, until
// the time limit kills it.
const STUB = [
  "const fs = require('node:fs');",
  'const [busy, reg, ...args] = process.argv.slice(2);',
  'if (fs.existsSync(busy)) setTimeout(() => {}, 60000);',
  "else if (args.includes('/v')) process.stdout.write('HKEY_CURRENT_USER UserChoice\\r\\n    ProgId    REG_SZ    MSEdgeHTM\\r\\n');",
  "else if (/^HKLM.SOFTWARE.Clients/.test(args[1])) process.stdout.write(fs.readFileSync(reg, 'utf8'));",
  'else process.exit(1); // the key is not there, as HKCU\'s often is not',
].join('\n');

test('a reg.exe that times out leaves the list empty and not kept; it is asked again, and kept once it answers', async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-browsers-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const stub = path.join(dir, 'reg.cjs');
  const busy = path.join(dir, 'busy');
  fs.writeFileSync(stub, STUB);
  fs.writeFileSync(path.join(dir, 'reg.txt'), REG);
  const runner = regRunner({ file: process.execPath, before: [stub, busy, path.join(dir, 'reg.txt')], timeout: 1500 });
  let reads = 0;
  const list = browserList({ read: () => { reads += 1; return readBrowsers({ platform: 'win32', reg: runner, exists: () => true }); } });

  fs.writeFileSync(busy, '');
  const timedOut = await readBrowsers({ platform: 'win32', reg: runner, exists: () => true });
  assert.deepEqual(timedOut, { list: [], complete: false }, 'a read that timed out is no answer');
  assert.deepEqual(await list.get(), [], 'the busy moment: nothing listed');
  assert.equal(reads, 1);

  fs.rmSync(busy);
  // Asked again (the dialog opened again), and this time reg.exe answers.
  const [a, b] = await Promise.all([list.get(), list.get()]);
  assert.equal(reads, 2, 'two asks at once share one read');
  assert.deepEqual(a.map((x) => [x.name, x.isDefault]), [['Microsoft Edge', true], ['Google Chrome', false]]);
  assert.deepEqual(b, a);
  // Kept: the next dialog asks no reg.exe at all.
  assert.deepEqual(await list.get(), a);
  assert.equal(reads, 2);
});

test('a list is kept only when it is whole and not empty', async () => {
  const answers = [
    () => ({ list: [{ id: 'Google Chrome', name: 'Google Chrome', isDefault: false }], complete: false }),
    () => ({ list: [], complete: true }),
    () => { throw new Error('reg.exe could not be started'); },
    () => ({ list: [{ id: 'Microsoft Edge', name: 'Microsoft Edge', isDefault: true }], complete: true }),
    () => ({ list: [], complete: true }),
  ];
  let reads = 0;
  const list = browserList({ read: async () => answers[reads++]() });
  assert.equal((await list.get()).length, 1, 'a short list is shown');
  assert.deepEqual(await list.get(), [], 'but asked again: empty this time');
  assert.deepEqual(await list.get(), [], 'a read that fails gives an empty list, not an error');
  assert.equal((await list.get())[0].name, 'Microsoft Edge');
  assert.equal((await list.get())[0].name, 'Microsoft Edge');
  assert.equal(reads, 4, 'the whole list is kept');
  list.forget();
  await list.get().catch(() => {});
  assert.equal(reads, 5, 'forgotten, it is read again');
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
