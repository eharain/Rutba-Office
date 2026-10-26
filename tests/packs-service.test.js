// Add or remove apps (apps/desktop/main/packs-service.js, docs/APP-PACKS.md
// 5.4): a pack is added from the release's file or a folder of packs, checked
// in a staging folder with the core's key and version before it is put in
// place, and loads at the next start; a pack removed goes at the next start;
// a pack that is not good is refused in plain words and nothing is put in
// place. The packs are made for the test with a key thrown away after it.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { makeKeyPair } from '../tools/make-pack-key.mjs';
import { makeAppPacks } from '../apps/desktop/build/make-app-packs.js';
import { planPacks, readPackKeys } from '../apps/desktop/main/app-packs.js';
import { createPacksService, packFileName, elevatedScript } from '../apps/desktop/main/packs-service.js';
import { packFiles, pendingRemovals } from '../packages/office-shell/src/electron/pack-install.js';

const VERSION = '9.9.9';
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-packs-service-'));
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));
const pair = makeKeyPair();
const keysFile = path.join(tmp, 'pack-keys.json');
const release = path.join(tmp, 'release');
const built = await makeAppPacks({ privateKey: crypto.createPrivateKey(pair.privatePem), version: VERSION, only: ['word', 'sheets', 'mail', 'contacts', 'ai'], release, root: path.join(tmp, 'packs'), keysFile, installerFile: path.join(tmp, 'installer-packs.nsh') });
const keys = readPackKeys(keysFile);

let installs = 0;
/** A copy installed with these packs, and its first start. */
function installed(ids) {
  const dir = path.join(tmp, `installed-${++installs}`);
  fs.mkdirSync(dir, { recursive: true });
  for (const id of ids) {
    const f = packFiles(id);
    for (const name of [f.archive, f.unpacked, f.manifest]) {
      if (fs.existsSync(path.join(built.out, name))) fs.cpSync(path.join(built.out, name), path.join(dir, name), { recursive: true });
    }
  }
  return { dir, plan: planPacks({ dir, keys, version: VERSION }) };
}
const service = (o) => createPacksService({ version: VERSION, keys, staging: path.join(tmp, `staging-${crypto.randomUUID()}`), broadcast: (e, p) => (o.events ||= []).push([e, p]), ...o });
const stateOf = (list, id) => list.apps.find((a) => a.id === id).state;

test('the list says what is installed, what is not, and each pack\'s size', () => {
  const { dir, plan } = installed(['word', 'contacts']);
  const list = service({ plan, dir, sizes: { mail: 4096 } }).list();
  assert.equal(list.managed, true);
  assert.equal(stateOf(list, 'word'), 'installed');
  assert.equal(stateOf(list, 'mail'), 'absent');
  assert.ok(list.apps.find((a) => a.id === 'word').size > 0);
  assert.equal(list.apps.find((a) => a.id === 'mail').size, 4096, 'one not here: the size the build measured');
  assert.equal(list.restart, false);
});

test('a pack is added from a folder of the release\'s pack files, checked first, and loads at the next start', async () => {
  const { dir, plan } = installed(['word', 'contacts']);
  const o = { plan, dir, source: release };
  const packs = service(o);
  const after = await packs.add({ id: 'mail' });
  assert.equal(stateOf(after, 'mail'), 'adding');
  assert.equal(after.restart, true);
  assert.equal(o.events.at(-1)[0], 'packs:changed');
  assert.ok(fs.existsSync(path.join(dir, 'mail.pack.json')));
  // The next start loads it.
  assert.deepEqual(planPacks({ dir, keys, version: VERSION }).ids().sort(), ['contacts', 'mail', 'word']);
  // Added in this run and taken back: its files go at once.
  await packs.remove({ id: 'mail' });
  assert.equal(fs.existsSync(path.join(dir, 'mail.pack.json')), false);
  assert.equal(fs.existsSync(path.join(dir, 'mail.asar')), false);
});

test('a pack is added from the release of this very version, fetched and checked', async () => {
  const { dir, plan } = installed(['word']);
  const asked = [];
  const fetchImpl = async (url) => {
    asked.push(url);
    const bytes = fs.readFileSync(path.join(release, packFileName(VERSION, 'sheets')));
    return new Response(bytes, { status: 200 });
  };
  const packs = service({ plan, dir, source: null, fetchImpl });
  await packs.add({ id: 'sheets' });
  assert.deepEqual(asked, [`https://github.com/eharain/Rutba-Office/releases/download/v${VERSION}/Rutba-Office-${VERSION}-sheets.pack`]);
  assert.ok(planPacks({ dir, keys, version: VERSION }).has('sheets'));

  // Offline, or the file not in the release: plain words, nothing put in place.
  const offline = service({ plan, dir, source: null, fetchImpl: async () => { throw new Error('ENOTFOUND'); } });
  await assert.rejects(offline.add({ id: 'mail' }), /^Error: The Mail pack could not be fetched: this computer could not reach the release\. Try again when it is online\.$/);
  const missing = service({ plan, dir, source: null, fetchImpl: async () => new Response('no', { status: 404 }) });
  await assert.rejects(missing.add({ id: 'mail' }), /could not be fetched from the release of Rutba Office 9\.9\.9 \(404\)/);
  assert.equal(fs.existsSync(path.join(dir, 'mail.pack.json')), false);
});

test('a pack that is not good is refused and nothing is put in place', async () => {
  const { dir, plan } = installed(['word']);
  // A folder of loose packs where one file was changed after signing.
  const spoilt = path.join(tmp, 'spoilt');
  fs.cpSync(built.out, spoilt, { recursive: true });
  const archive = path.join(spoilt, 'mail.asar');
  const bytes = fs.readFileSync(archive);
  bytes[bytes.length - 1] ^= 0xff;
  fs.writeFileSync(archive, bytes);
  await assert.rejects(service({ plan, dir, source: spoilt }).add({ id: 'mail' }), /Mail/);
  assert.equal(fs.existsSync(path.join(dir, 'mail.pack.json')), false);
  assert.equal(fs.existsSync(path.join(dir, 'mail.asar')), false);

  // Signed with another key.
  const other = makeKeyPair();
  await assert.rejects(service({ plan, dir, source: release, keys: readPackKeysOf(other.publicPem) }).add({ id: 'mail' }), /^Error: The mail pack is not signed by Rutba, or was changed after it was signed\. It was not installed\.$/);
  // Of another version: the release's file is not there by its name, and a loose pack is refused by the loader.
  await assert.rejects(service({ plan, dir, source: release, version: '9.9.8' }).add({ id: 'mail' }), /The Mail pack is not in /);
  await assert.rejects(service({ plan, dir, source: built.out, version: '9.9.8' }).add({ id: 'mail' }), /9\.9\.9/);
  assert.equal(fs.existsSync(path.join(dir, 'mail.pack.json')), false);

  // Not an app, already installed, not in the folder.
  const packs = service({ plan, dir, source: release });
  await assert.rejects(packs.add({ id: 'nonsense' }), /^Error: There is no app called nonsense to add\.$/);
  await assert.rejects(packs.add({ id: 'word' }), /^Error: Documents is already installed\.$/);
  await assert.rejects(packs.add({ id: 'talk' }), /The Talk pack is not in /);
});

function readPackKeysOf(pem) {
  const file = path.join(tmp, `keys-${crypto.randomUUID()}.json`);
  fs.writeFileSync(file, JSON.stringify({ keys: [pem] }));
  return readPackKeys(file);
}

test('a pack removed goes at the next start, its files then deleted; Keep takes it back; the last app stays', async () => {
  const { dir, plan } = installed(['word', 'contacts', 'ai']);
  const packs = service({ plan, dir, source: release });
  const after = await packs.remove({ id: 'contacts' });
  assert.equal(stateOf(after, 'contacts'), 'removing');
  assert.equal(after.restart, true);
  assert.deepEqual(pendingRemovals(dir), ['contacts']);
  assert.ok(fs.existsSync(path.join(dir, 'contacts.pack.json')), 'still there while this run holds it');
  assert.equal(stateOf(await packs.keep({ id: 'contacts' }), 'contacts'), 'installed');
  assert.deepEqual(pendingRemovals(dir), []);
  await packs.remove({ id: 'contacts' });
  // Add after Remove: it simply stays.
  assert.equal(stateOf(await packs.add({ id: 'contacts' }), 'contacts'), 'installed');
  await packs.remove({ id: 'contacts' });
  // Documents is the last app; AI Tools is not an app of its own.
  await assert.rejects(packs.remove({ id: 'word' }), /^Error: Keep one app at least\. To remove Rutba Office altogether, uninstall it from Windows Settings\.$/);
  await assert.rejects(packs.remove({ id: 'mail' }), /^Error: Mail is not installed\.$/);

  const next = planPacks({ dir, keys, version: VERSION });
  assert.deepEqual(next.ids().sort(), ['ai', 'word']);
  assert.equal(fs.existsSync(path.join(dir, 'contacts.asar')), false);
  assert.deepEqual(pendingRemovals(dir), []);
});

test('with every app built in there is nothing to add or remove, and it says so', async () => {
  const plan = planPacks({ dir: null, version: VERSION });
  const packs = service({ plan, dir: null });
  const list = packs.list();
  assert.equal(list.managed, false);
  assert.ok(list.apps.every((a) => a.state === 'installed'));
  await assert.rejects(packs.add({ id: 'mail' }), /^Error: Every app is built into this copy of Rutba Office, so there is nothing to add or remove\. The installed copy lets you choose\.$/);
  await assert.rejects(packs.remove({ id: 'mail' }), /Every app is built into this copy/);
});

test('Restart starts Rutba Office again; a copy installed for everyone is changed by a script run with permission', async () => {
  const { dir, plan } = installed(['word']);
  const calls = [];
  const app = { isPackaged: false, relaunch: () => calls.push('relaunch'), quit: () => calls.push('quit') };
  assert.deepEqual(await service({ plan, dir, app }).restart(), { restarting: true, elevated: false });
  assert.deepEqual(calls, ['relaunch', 'quit']);

  const script = elevatedScript({
    pid: 1234,
    dir: 'C:\\Program Files\\Rutba Office\\resources\\packs',
    exe: 'C:\\Program Files\\Rutba Office\\Rutba Office.exe',
    removals: [{ id: 'mail', types: [{ ext: 'eml', name: 'Email Message' }] }],
    additions: [{ id: 'word', staging: 'C:\\Temp\\it\'s\\word', types: [{ ext: 'docx', name: 'Word Document', description: 'Word Document', icon: 'C:\\x.ico' }] }],
  });
  assert.match(script, /^\$ErrorActionPreference = 'Continue'\r\nWait-Process -Id 1234 /);
  assert.match(script, /Remove-Item -LiteralPath 'C:\\Program Files\\Rutba Office\\resources\\packs\\mail\.pack\.json'/);
  assert.match(script, /Remove-ItemProperty -Path 'HKLM:\\Software\\Classes\\\.eml\\OpenWithProgids' -Name 'Email Message'/);
  assert.match(script, /Copy-Item -LiteralPath 'C:\\Temp\\it''s\\word\\word\.asar'/, 'a quote in a path is doubled');
  // The manifest is copied last.
  assert.ok(script.indexOf('word.pack.json') > script.indexOf('word.asar.unpacked'));
  assert.match(script, /Start-Process -FilePath 'C:\\Program Files\\Rutba Office\\Rutba Office\.exe'\r\n$/);
});
