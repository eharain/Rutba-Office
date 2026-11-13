// The add-ons at start (addons-start.js: the failed-start count, the time
// budget, staged packs moved in), one file fetched (addons-download.js), and
// the add-ons service against a release made here (addons-service.js: the
// feed, adding with what is needed, removing, new revisions, withdrawals,
// and the update gate's two questions). Keys made here and thrown away.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { canonical } from '../packages/office-shell/src/electron/packs.js';
import { FEED_WORDS } from '../packages/office-shell/src/electron/addons-compat.js';
import { envelopeOf, fileEntry, manifests, CONTRACT_FILE } from '../tools/make-addon-vectors.mjs';
import { startAddons, applyStaged, FAILED_STARTS } from '../apps/desktop/main/addons-start.js';
import { fetchToFile, allowedUrl, releaseUrl, RELEASES, DownloadError } from '../apps/desktop/main/addons-download.js';
import { createAddonsService, closure } from '../apps/desktop/main/addons-service.js';
import { stagingFor, markRemoval, readState, PARTS } from '../apps/desktop/main/addons-folder.js';
import { writeZip, ZipEntry } from '../packages/ooxml/src/zip.js';

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-addons-svc-'));
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));
const file = JSON.parse(fs.readFileSync(CONTRACT_FILE, 'utf8'));
const OFFICE = file.office;
const NEXT = '99.0.0';
const packagesKey = crypto.generateKeyPairSync('ed25519');
const releaseKey = crypto.generateKeyPairSync('ed25519');
const pem = (k) => k.publicKey.export({ type: 'spki', format: 'pem' }).toString();
const contract = { ...file, keys: { packages: [pem(packagesKey)], release: [pem(releaseKey)] } };
const M = manifests(contract);
let n = 0;
const fresh = (name = 'd') => {
  const d = path.join(tmp, `${name}${(n += 1)}`);
  fs.mkdirSync(d, { recursive: true });
  return d;
};
const keyFor = (m) => (m.kind === 'engine' ? releaseKey : packagesKey);
const engineContent = { 'main/index.js': 'export function register() { return {}; }\n' };

/** An add-on written into a folder, signed. */
function put(at, manifest, content = { 'data.txt': 'test\n' }) {
  const m = { ...manifest, files: Object.fromEntries(Object.entries(content).map(([p, b]) => [p, fileEntry(b)])) };
  fs.rmSync(path.join(at, m.id), { recursive: true, force: true });
  for (const [p, b] of Object.entries(content)) {
    fs.mkdirSync(path.dirname(path.join(at, m.id, p)), { recursive: true });
    fs.writeFileSync(path.join(at, m.id, p), b);
  }
  fs.writeFileSync(path.join(at, `${m.id}.pack.json`), JSON.stringify(envelopeOf(m, keyFor(m).privateKey)));
  return m;
}

/**
 * A release as the packages repository publishes one, served by a fetch of
 * our own: feed.json, its .sig, each envelope and each file as an asset.
 * `served` counts what was asked for.
 */
function release({ version, packs, sequence = 1, state = 'partial', retired = [], withdrawn = [], key = packagesKey }) {
  const assets = new Map();
  const base = releaseUrl(version);
  const entries = packs.map(({ manifest, content = { 'data.txt': 'test\n' } }) => {
    const m = { ...manifest, version, files: Object.fromEntries(Object.entries(content).map(([p, b]) => [p, fileEntry(b)])) };
    const env = Buffer.from(JSON.stringify(envelopeOf(m, keyFor(m).privateKey)));
    assets.set(`${base}${m.id}.pack.json`, env);
    const files = {};
    for (const [p, b] of Object.entries(content)) {
      const url = `${base}${m.id}--${p.replace(/\//g, '--')}`;
      assets.set(url, Buffer.from(b));
      files[p] = { url, ...fileEntry(b) };
    }
    return {
      id: m.id, kind: m.kind, version, revision: m.revision, name: m.name, requires: m.requires, provides: m.provides, licence: { spdx: 'MIT' },
      size: Object.values(files).reduce((s, f) => s + f.size, 0), withdrawn: withdrawn.includes(m.id),
      envelope: { url: `${base}${m.id}.pack.json`, ...fileEntry(env) }, files,
    };
  });
  const feed = { format: 1, readiness: 1, office: version, sequence, state, retired, date: '2026-11-13T00:00:00Z', packs: entries };
  assets.set(`${base}feed.json`, Buffer.from(JSON.stringify(feed)));
  assets.set(`${base}feed.json.sig`, Buffer.from(JSON.stringify({ signature: crypto.sign(null, Buffer.from(FEED_WORDS + canonical(feed), 'utf8'), key.privateKey).toString('base64') })));
  return { feed, assets };
}

function server(...releases) {
  const assets = new Map(releases.flatMap((r) => [...r.assets]));
  const served = [];
  const fetchImpl = async (url, init = {}) => {
    served.push(url);
    const body = assets.get(url);
    if (!body) return new Response('no', { status: 404 });
    const range = init.headers?.range;
    if (range) {
      const from = Number(/bytes=(\d+)-/.exec(range)[1]);
      return new Response(body.subarray(from), { status: 206, headers: { etag: '"x"' } });
    }
    return new Response(body, { status: 200, headers: { etag: '"x"' } });
  };
  return { fetchImpl, served, set: (url, bytes) => assets.set(url, bytes) };
}

const settings = () => {
  const m = new Map();
  return { get: (k, d) => (m.has(k) ? m.get(k) : d), set: (k, v) => m.set(k, v), all: () => Object.fromEntries(m) };
};

/* ── the start ─────────────────────────────────────────────────────── */

test('a start with no folder, or with --no-addons, has none', () => {
  assert.equal(startAddons({ folder: null, contract, office: OFFICE, memoryFile: path.join(tmp, 'm0.json') }).registry.list().length, 0);
  const dir = fresh();
  put(dir, M.lang);
  const r = startAddons({ folder: { dir, scope: 'person' }, contract, office: OFFICE, memoryFile: path.join(tmp, 'm1.json'), argv: ['--no-addons'] });
  assert.equal(r.skipped, 'switch');
  assert.equal(r.registry.list().length, 0);
});

test('two starts that do not reach the first window skip add-ons at the third, until Try again or a change', () => {
  const dir = fresh();
  put(dir, M.lang);
  const memoryFile = path.join(fresh(), 'starts.json');
  const go = () => startAddons({ folder: { dir, scope: 'person' }, contract, office: OFFICE, memoryFile, argv: [] });
  for (let i = 0; i < FAILED_STARTS; i += 1) assert.equal(go().skipped, null);
  const third = go();
  assert.equal(third.skipped, 'failed-starts');
  assert.match(third.said, /Started without voices and languages/);
  assert.deepEqual([...third.registry.installed], ['lang-xx'], 'still counted by the update gate');
  assert.equal(go().skipped, 'failed-starts', 'and stays skipped');
  third.tryAgain();
  const again = go();
  assert.equal(again.skipped, null);
  again.reached();
  assert.equal(go().skipped, null);
  go();
  assert.equal(go().skipped, 'failed-starts');
  put(dir, { ...M.lang, revision: 2 }); // the add-ons changed: the count starts again
  assert.equal(go().skipped, null);
});

test('what is not checked within the time is left for the next start', () => {
  const dir = fresh();
  put(dir, M.lang);
  put(dir, M.dict);
  let t = 0;
  const r = startAddons({ folder: { dir, scope: 'person' }, contract, office: OFFICE, memoryFile: path.join(fresh(), 's.json'), argv: [], budgetMs: 1, now: () => (t += 1) });
  assert.ok(r.registry.deferred.length >= 1);
  assert.match(r.said, /next start/);
});

test('staged packs are moved in with this version\'s rules, unchanged files taken from the add-on in place', () => {
  const dir = fresh();
  put(dir, M.lang, { 'data.txt': 'test\n', 'big.bin': 'unchanged' });
  const staging = stagingFor(dir, OFFICE);
  fs.mkdirSync(staging, { recursive: true });
  // Revision 2 changes data.txt only: big.bin is not staged.
  put(staging, { ...M.lang, revision: 2 }, { 'data.txt': 'two\n', 'big.bin': 'unchanged' });
  fs.rmSync(path.join(staging, 'lang-xx', 'big.bin'));
  // A pack with a broken signature is never moved in.
  put(staging, M.dict);
  const env = JSON.parse(fs.readFileSync(path.join(staging, 'dict-xx.pack.json'), 'utf8'));
  fs.writeFileSync(path.join(staging, 'dict-xx.pack.json'), JSON.stringify({ ...env, manifest: { ...env.manifest, revision: 9 } }));
  fs.mkdirSync(path.join(staging, PARTS, 'other'), { recursive: true });
  fs.mkdirSync(stagingFor(dir, NEXT), { recursive: true });
  fs.mkdirSync(stagingFor(dir, '0.0.1'), { recursive: true });

  const r = applyStaged({ dir, contract, office: OFFICE, platform: 'win32-x64' });
  assert.deepEqual(r.moved, ['lang-xx']);
  assert.deepEqual(r.refused.map((x) => [x.id, x.code]), [['dict-xx', 'signature']]);
  assert.equal(fs.readFileSync(path.join(dir, 'lang-xx', 'data.txt'), 'utf8'), 'two\n');
  assert.equal(fs.readFileSync(path.join(dir, 'lang-xx', 'big.bin'), 'utf8'), 'unchanged');
  assert.ok(fs.existsSync(path.join(dir, 'lang-xx.previous', 'lang-xx.pack.json')), 'the one replaced is kept');
  assert.ok(!fs.existsSync(path.join(staging, 'dict-xx.pack.json')), 'nothing staged is tried twice');
  assert.ok(fs.existsSync(path.join(staging, PARTS, 'other')), 'downloads in progress stay');
  assert.ok(fs.existsSync(stagingFor(dir, NEXT)), 'a newer version\'s staging stays');
  assert.ok(!fs.existsSync(stagingFor(dir, '0.0.1')), 'an older version\'s goes');
});

test('the start moves staged packs in and does removals before it loads', () => {
  const dir = fresh();
  put(dir, M.lang);
  put(dir, M.dict);
  markRemoval(dir, 'dict-xx');
  put(stagingFor(dir, OFFICE), { ...M.lang, revision: 3 });
  const r = startAddons({ folder: { dir, scope: 'person' }, contract, office: OFFICE, memoryFile: path.join(fresh(), 's.json'), argv: [], platform: 'win32-x64' });
  assert.deepEqual(r.applied.moved, ['lang-xx']);
  assert.deepEqual(r.removed, ['dict-xx']);
  assert.equal(r.registry.find('lang-xx').manifest.revision, 3);
  assert.equal(r.registry.find('dict-xx'), null);
});

/* ── one file ──────────────────────────────────────────────────────── */

test('only the packages releases are fetched from, or files under a named folder', () => {
  assert.ok(allowedUrl(`${RELEASES}v1.0.0/feed.json`));
  assert.ok(!allowedUrl('https://github.com/eharain/other/releases/download/v1/x'));
  assert.ok(!allowedUrl(`${RELEASES}v1.0.0/../x`));
  assert.ok(!allowedUrl('file:///C:/x/y', {}));
  const share = fresh('share');
  assert.ok(allowedUrl(new URL(`file:///${share.replace(/\\/g, '/')}/v1/x`).href, { source: share }));
});

test('a file is fetched whole, resumed after a cut, and never kept when its hash is wrong', async () => {
  const dir = fresh();
  const bytes = crypto.randomBytes(5000);
  const url = `${RELEASES}v1.0.0/x--data.bin`;
  const s = server();
  s.set(url, bytes);
  const to = path.join(dir, 'data.bin');
  const e = fileEntry(bytes);
  // A cut: the first 1200 bytes are there, with the validator.
  fs.writeFileSync(`${to}.part`, bytes.subarray(0, 1200));
  fs.writeFileSync(`${to}.part.json`, JSON.stringify({ url, validator: '"x"' }));
  let counted = 0;
  await fetchToFile({ url, ...e, to, fetchImpl: s.fetchImpl, onBytes: (k) => { counted += k; } });
  assert.deepEqual(fs.readFileSync(to), bytes);
  assert.equal(counted, 5000 - 1200, 'only the rest was fetched');
  assert.ok(!fs.existsSync(`${to}.part`));

  const bad = path.join(dir, 'bad.bin');
  await assert.rejects(fetchToFile({ url, size: e.size, sha512: fileEntry('other').sha512, to: bad, fetchImpl: s.fetchImpl }), (err) => err instanceof DownloadError && err.code === 'hash');
  assert.ok(!fs.existsSync(bad) && !fs.existsSync(`${bad}.part`));
  await assert.rejects(fetchToFile({ url: 'https://example.com/x', ...e, to: bad, fetchImpl: s.fetchImpl }), (err) => err.code === 'address');
  await assert.rejects(fetchToFile({ url, size: 10, sha512: e.sha512, to: bad, fetchImpl: s.fetchImpl }), (err) => err.code === 'size');
});

/* ── the service ───────────────────────────────────────────────────── */

const service = ({ dir, scope = 'person', srv, stores = { settings: settings() }, start = null, ...more }) => createAddonsService({
  folder: dir ? { dir, scope } : null,
  start: start || startAddons({ folder: dir ? { dir, scope } : null, contract, office: OFFICE, memoryFile: path.join(fresh(), 's.json'), argv: [], platform: 'win32-x64' }),
  contract, office: OFFICE, stores, fetchImpl: srv.fetchImpl, source: null, platform: 'win32-x64', ...more,
});
const until = async (cond) => {
  for (let i = 0; i < 200 && !cond(); i += 1) await new Promise((r) => setTimeout(r, 5));
  assert.ok(cond(), 'in time');
};

test('closure follows requires by id or by what a pack provides', () => {
  const packs = [{ id: 'voice-a', requires: ['model-x'] }, { id: 'model-x', requires: ['speech'] }, { id: 'speech-engine', provides: ['speech'] }, { id: 'other' }];
  assert.deepEqual(closure(['voice-a'], packs), ['voice-a', 'model-x', 'speech-engine']);
});

test('the feed is fetched, verified and taken once per sequence; a bad signature is not', async () => {
  const dir = fresh();
  const rel = release({ version: OFFICE, packs: [{ manifest: M.lang }], sequence: 5 });
  const srv = server(rel);
  const svc = service({ dir, srv });
  const l = await svc.feed();
  assert.deepEqual(l.feed, { office: OFFICE, sequence: 5, state: 'partial', from: 'network' });
  assert.equal(l.offered[0].status, 'add');
  assert.deepEqual(readState(dir).seen, { [OFFICE]: 5 });

  const forged = release({ version: OFFICE, packs: [{ manifest: M.lang }], sequence: 6, key: crypto.generateKeyPairSync('ed25519') });
  const other = service({ dir: fresh(), srv: server(forged) });
  assert.equal((await other.feed()).feed, null);
});

test('the same feed is taken again at the next start; another with the same sequence is not', async () => {
  const dir = fresh();
  const rel = release({ version: OFFICE, packs: [{ manifest: M.lang }], sequence: 3 });
  const first = service({ dir, srv: server(rel) });
  assert.equal((await first.feed()).feed.from, 'network');
  // The next start: a new service on the same folder, the same release.
  const second = service({ dir, srv: server(rel) });
  const again = await second.feed();
  assert.equal(again.feed?.from, 'network', 'taken again, not refused as older');
  assert.equal(again.offered[0].status, 'add');
  // Another feed at that sequence (a replay with other packs): refused.
  const other = release({ version: OFFICE, packs: [{ manifest: M.lang }, { manifest: M.dict }], sequence: 3 });
  assert.equal((await service({ dir, srv: server(other) }).feed()).feed, null);
  // Where the folder cannot be written, the profile keeps the hash.
  const stores = { settings: settings() };
  const person = (srv) => service({ dir: fresh(), scope: 'machine', srv, stores });
  assert.equal((await person(server(rel)).feed()).feed?.from, 'network');
  assert.equal((await person(server(rel)).feed()).feed?.from, 'network', 'again, from the profile');
});

test('with the network gone, the feed built into the release lists what can be added', async () => {
  const rel = release({ version: OFFICE, packs: [{ manifest: M.lang }] });
  const feed = JSON.parse(rel.assets.get(`${releaseUrl(OFFICE)}feed.json`));
  const sig = JSON.parse(rel.assets.get(`${releaseUrl(OFFICE)}feed.json.sig`));
  const svc = service({ dir: fresh(), srv: server(), bundled: { feed, sig } });
  assert.equal((await svc.feed()).feed.from, 'bundled');
});

test('adding a voice brings what it needs, fetched and put in place for the next start', async () => {
  const dir = fresh();
  const rel = release({ version: OFFICE, packs: [{ manifest: M.engine, content: engineContent }, { manifest: M.model }, { manifest: M.voice }, { manifest: M.lang }] });
  const srv = server(rel);
  const events = [];
  const svc = service({ dir, srv, broadcast: (name, p) => events.push([name, p]) });
  const r = await svc.add({ id: 'voice-test' });
  assert.deepEqual(r.adding.sort(), ['model-test', 'speech-engine', 'voice-test']);
  await until(() => ['speech-engine', 'model-test', 'voice-test'].every((id) => fs.existsSync(path.join(dir, `${id}.pack.json`))));
  await until(() => svc.list().restart);
  assert.ok(events.some(([e]) => e === 'addons:changed'));
  // The next start uses them.
  const next = startAddons({ folder: { dir, scope: 'person' }, contract, office: OFFICE, memoryFile: path.join(fresh(), 's.json'), argv: [], platform: 'win32-x64' });
  assert.deepEqual(next.registry.list().map((a) => a.id), ['model-test', 'speech-engine', 'voice-test']);
  assert.ok(!srv.served.some((u) => u.includes('lang-xx')), 'nothing not asked for');
});

test('an install for everyone stages in the person\'s folder and asks for the administrator at Restart', async () => {
  const dir = fresh();
  const person = fresh('person');
  const rel = release({ version: OFFICE, packs: [{ manifest: M.lang }] });
  const asked = [];
  const svc = service({
    dir, scope: 'machine', srv: server(rel), personStaging: person,
    start: { registry: startAddons({ folder: { dir, scope: 'machine' }, contract, office: OFFICE, memoryFile: path.join(fresh(), 's.json'), argv: [] }).registry, writable: false },
    elevate: async (args, o) => { asked.push([args, o]); return { ok: true }; }, app: { relaunch() { calls.push('relaunch'); }, quit() { calls.push('quit'); }, exit() { calls.push('exit'); } },
    pid: 4242,
  });
  const calls = [];
  await svc.add({ id: 'lang-xx' });
  await until(() => fs.existsSync(path.join(person, OFFICE, 'lang-xx.pack.json')));
  assert.ok(!fs.existsSync(path.join(dir, 'lang-xx.pack.json')), 'the machine folder is not the person\'s to write');
  assert.deepEqual(svc.list().staged, ['lang-xx']);
  assert.deepEqual(await svc.restart(), { restarting: true, elevated: true });
  // The administrator's step waits for this process to quit before it moves anything, then starts Office again as the person.
  assert.deepEqual(asked, [[[`--addons-apply-staged=${person}`, '--addons-after=4242', '--addons-relaunch'], { wait: false }]]);
  assert.deepEqual(calls, ['quit'], 'quit, never exit: the quit steps run; the step brings Office back');
});

test('an envelope put in the folder by hand counts for nothing, and a start never writes an install for everyone\'s folder', async () => {
  const dir = fresh();
  put(dir, M.lang);
  // Junk, and a pack signed with no key this Office trusts: neither is installed, so neither can hold an update.
  fs.writeFileSync(path.join(dir, 'junk-xx.pack.json'), JSON.stringify({ format: 1, keyId: 'x', manifest: { id: 'junk-xx', kind: 'voice', version: OFFICE }, signature: 'AAAA' }));
  const stranger = crypto.generateKeyPairSync('ed25519');
  const theirs = put(dir, { ...M.lang, id: 'lang-yy' });
  fs.writeFileSync(path.join(dir, 'lang-yy.pack.json'), JSON.stringify(envelopeOf(theirs, stranger.privateKey)));
  // And a signed envelope under another id's name.
  fs.copyFileSync(path.join(dir, 'lang-xx.pack.json'), path.join(dir, 'lang-zz.pack.json'));
  const svc = service({ dir, srv: server(release({ version: OFFICE, packs: [{ manifest: M.lang }] })) });
  assert.deepEqual(svc.installed().map((p) => p.id), ['lang-xx']);

  // For everyone: the start neither makes the folder nor moves staged packs in as the person.
  const machine = path.join(fresh(), 'not-yet');
  const s = startAddons({ folder: { dir: machine, scope: 'machine' }, contract, office: OFFICE, memoryFile: path.join(fresh(), 's.json'), argv: [] });
  assert.equal(s.writable, false);
  assert.ok(!fs.existsSync(machine), 'not even probed');
  const there = fresh();
  put(stagingFor(there, OFFICE), M.lang);
  startAddons({ folder: { dir: there, scope: 'machine' }, contract, office: OFFICE, memoryFile: path.join(fresh(), 's.json'), argv: [] });
  assert.ok(!fs.existsSync(path.join(there, 'lang-xx.pack.json')), 'staged packs wait for the administrator\'s step');
});

test('the administrator\'s step sets the folder\'s permissions before it moves anything, and moves nothing when it cannot', async () => {
  const { applyStagedRun } = await import('../apps/desktop/main/addons-wiring.js');
  const { machineFolderAcl } = await import('../apps/desktop/main/addons-folder.js');
  const programFiles = fresh('pf');
  const programData = fresh('pd');
  const env = { ProgramFiles: programFiles, ProgramData: programData };
  const execPath = path.join(programFiles, 'Rutba Office', 'Rutba Office.exe');
  const dir = path.join(programData, 'Rutba Office', 'addons');
  put(dir, M.dict);
  const secured = [];
  const run = (secure) => applyStagedRun({ argv: ['--addons-remove=dict-xx'], execPath, isPackaged: true, office: OFFICE, releaseKeys: [], env, platform: 'win32', secure, wait: () => true, relaunch: () => {} });
  assert.equal(run(() => false), 1);
  assert.ok(fs.existsSync(path.join(dir, 'dict-xx.pack.json')), 'nothing done');
  assert.equal(run((d) => { secured.push(d); return true; }), 0);
  assert.deepEqual(secured, [dir]);
  assert.ok(!fs.existsSync(path.join(dir, 'dict-xx.pack.json')), 'then the removal');
  // Administrators and SYSTEM may change it; users, and a file's owner, only read; nothing inherited.
  const [root, inside] = machineFolderAcl(dir);
  assert.deepEqual(root.slice(1, 7), ['/inheritance:r', '/grant:r', '*S-1-5-18:(OI)(CI)F', '*S-1-5-32-544:(OI)(CI)F', '*S-1-5-32-545:(OI)(CI)RX', '*S-1-3-4:(OI)(CI)RX']);
  assert.deepEqual(inside.slice(1, 3), ['/reset', '/T']);
});

test('Restart is not over unsaved work, quits rather than exits, and a no from Windows forgets nothing', async () => {
  const calls = [];
  const app = { relaunch() { calls.push('relaunch'); }, quit() { calls.push('quit'); }, exit() { calls.push('exit'); } };
  let windows = [{ name: 'Budget.xlsx', dirty: true }, { name: 'Notes.docx', dirty: false }];
  const svc = service({ dir: fresh(), srv: server(), app, openWindows: () => windows });
  const stopped = await svc.restart();
  assert.equal(stopped.restarting, false);
  assert.deepEqual(stopped.unsaved, ['Budget.xlsx']);
  assert.deepEqual(calls, [], 'nothing closed');
  windows = [];
  assert.deepEqual(await svc.tryAgain(), { restarting: true });
  assert.deepEqual(calls, ['relaunch', 'quit']);

  // An install for everyone, the administrator's prompt declined: the removals wait for the next try.
  const stores = { settings: settings() };
  stores.settings.set('addons.removeAtRestart', ['dict-xx']);
  calls.length = 0;
  const dir = fresh();
  const machine = service({
    dir, scope: 'machine', srv: server(), stores, personStaging: fresh('person'), app,
    start: { registry: startAddons({ folder: { dir, scope: 'machine' }, contract, office: OFFICE, memoryFile: path.join(fresh(), 's.json'), argv: [] }).registry, writable: false },
    elevate: async () => ({ ok: false, code: 5 }),
  });
  const declined = await machine.restart();
  assert.equal(declined.restarting, false);
  assert.ok(declined.declined);
  assert.deepEqual(stores.settings.get('addons.removeAtRestart'), ['dict-xx'], 'kept');
  assert.deepEqual(calls, []);
});

test('a download paused keeps its part and is resumed by Add; Cancel takes its files away', async () => {
  const dir = fresh();
  const big = crypto.randomBytes(300000).toString('base64');
  const rel = release({ version: OFFICE, packs: [{ manifest: M.lang, content: { 'data.txt': big } }] });
  const srv = server(rel);
  const slow = async (url, init) => {
    const r = await srv.fetchImpl(url, init);
    if (!url.endsWith('data.txt')) return r;
    const all = Buffer.from(await r.arrayBuffer());
    let at = 0;
    return new Response(new ReadableStream({
      async pull(c) {
        if (init.signal?.aborted) { c.error(new Error('aborted')); return; }
        await new Promise((x) => setTimeout(x, 2));
        if (at >= all.length) { c.close(); return; }
        c.enqueue(all.subarray(at, at + 16384));
        at += 16384;
      },
    }), { status: r.status, headers: { etag: '"x"' } });
  };
  const svc = createAddonsService({ folder: { dir, scope: 'person' }, start: { registry: null, writable: true }, contract, office: OFFICE, stores: { settings: settings() }, fetchImpl: slow, platform: 'win32-x64' });
  await svc.add({ id: 'lang-xx' });
  await until(() => (svc.list().jobs[0]?.done || 0) > 0);
  svc.pause({ id: 'lang-xx' });
  await until(() => svc.list().jobs[0]?.state === 'paused');
  const part = path.join(stagingFor(dir, OFFICE), PARTS, 'lang-xx', 'lang-xx', 'data.txt.part');
  await until(() => fs.existsSync(part));
  await svc.add({ id: 'lang-xx' });
  await until(() => fs.existsSync(path.join(dir, 'lang-xx.pack.json')));

  // Cancel, on a fresh addition.
  const dir2 = fresh();
  const svc2 = createAddonsService({ folder: { dir: dir2, scope: 'person' }, start: { registry: null, writable: true }, contract, office: OFFICE, stores: { settings: settings() }, fetchImpl: slow, platform: 'win32-x64' });
  await svc2.add({ id: 'lang-xx' });
  await until(() => (svc2.list().jobs[0]?.done || 0) > 0);
  svc2.cancel({ id: 'lang-xx' });
  assert.ok(!fs.existsSync(path.join(stagingFor(dir2, OFFICE), PARTS, 'lang-xx')));
  assert.deepEqual(svc2.list().jobs, []);
});

test('removing says what depends on it, and is done at the next start; cancel takes it back', async () => {
  const dir = fresh();
  put(dir, M.lang);
  put(dir, M.dict);
  const svc = service({ dir, srv: server() });
  const r = svc.remove({ id: 'lang-xx' });
  assert.deepEqual(r.dependents, ['dict-xx']);
  assert.equal(r.installed.find((p) => p.id === 'lang-xx').state, 'removing');
  svc.cancel({ id: 'lang-xx' });
  assert.deepEqual(readState(dir).removals, []);
});

test('a language in use is not removed, English never is, and the list says how much each has (L6)', () => {
  const dir = fresh();
  put(dir, M.lang);
  let inUse = ['lang-xx'];
  const languages = { inUse: () => inUse, coverageOf: (id) => (id === 'lang-xx' ? { found: 9, of: 10 } : null), known: () => [{ tag: 'xx', pack: 'lang-xx', name: 'Test' }] };
  const svc = service({ dir, srv: server(), languages });
  assert.throws(() => svc.remove({ id: 'lang-xx' }), /A window is open in this language/);
  assert.throws(() => svc.remove({ id: 'lang-en' }), /English is part of Rutba Office/);
  const listed = svc.list();
  const xx = listed.installed.find((p) => p.id === 'lang-xx');
  assert.deepEqual(xx.coverage, { found: 9, of: 10 });
  assert.equal(xx.inUse, true);
  assert.deepEqual(listed.languages.map((l) => l.pack), ['lang-xx']);
  inUse = [];
  assert.equal(svc.remove({ id: 'lang-xx' }).installed.find((p) => p.id === 'lang-xx').state, 'removing');
});

test('a new revision is taken by itself with updates on; a withdrawn one is rolled back or turned off', async () => {
  const dir = fresh();
  put(dir, M.lang);
  put(dir, M.dict);
  // lang-xx revision 2 is out; dict-xx's revision 1 is withdrawn, with nothing kept to go back to.
  const rel = release({ version: OFFICE, packs: [{ manifest: { ...M.lang, revision: 2 }, content: { 'data.txt': 'two\n' } }, { manifest: M.dict }], withdrawn: ['dict-xx'] });
  const stores = { settings: settings() };
  const svc = service({ dir, srv: server(rel), stores });
  const out = await svc.check();
  assert.deepEqual(out.updated, ['lang-xx']);
  assert.deepEqual(out.turnedOff, ['dict-xx']);
  await until(() => JSON.parse(fs.readFileSync(path.join(dir, 'lang-xx.pack.json'), 'utf8')).manifest.revision === 2);
  assert.deepEqual(stores.settings.get('addons.off'), ['dict-xx']);
  const next = startAddons({ folder: { dir, scope: 'person' }, contract, office: OFFICE, memoryFile: path.join(fresh(), 's.json'), argv: [], off: stores.settings.get('addons.off') });
  assert.equal(next.registry.find('dict-xx'), null);

  // Withdrawn with the previous kept: rolled back.
  const rel3 = release({ version: OFFICE, packs: [{ manifest: { ...M.lang, revision: 2 }, content: { 'data.txt': 'two\n' } }], withdrawn: ['lang-xx'], sequence: 2 });
  const svc3 = service({ dir, srv: server(rel3), stores });
  assert.deepEqual((await svc3.check()).rolledBack, ['lang-xx']);
  assert.equal(JSON.parse(fs.readFileSync(path.join(dir, 'lang-xx.pack.json'), 'utf8')).manifest.revision, 1);

  // Updates off: a new revision waits for a click.
  const off = { settings: settings() };
  off.settings.set('addons.autoUpdate', false);
  const rel4 = release({ version: OFFICE, packs: [{ manifest: { ...M.lang, revision: 5 } }], sequence: 3 });
  assert.deepEqual((await service({ dir, srv: server(rel4), stores: off }).check()).waiting, ['lang-xx']);
});

test('the update gate: no add-ons, an offer with no request', async () => {
  const srv = server();
  const svc = service({ dir: fresh(), srv });
  const r = await svc.readiness(NEXT);
  assert.equal(r.offer, true);
  assert.deepEqual(srv.served, []);
});

test('the update gate: an add-on not yet published for the version waits, by name and kind', async () => {
  const dir = fresh();
  put(dir, M.lang);
  put(dir, M.dict);
  const next = release({ version: NEXT, packs: [{ manifest: M.lang }] });
  const svc = service({ dir, srv: server(next) });
  const r = await svc.readiness(NEXT);
  assert.equal(r.offer, false);
  assert.deepEqual(r.waitingFor, [{ id: 'dict-xx', why: 'unpublished', name: 'Test dictionary', kind: 'dictionary' }]);
  // No feed at all: waiting, said as not checked.
  const none = await service({ dir, srv: server() }).readiness(NEXT);
  assert.equal(none.offer, false);
  assert.ok(none.feedError);
});

test('the update gate: ready add-ons are staged for the new version, which moves them in at its first start', async () => {
  const dir = fresh();
  put(dir, M.lang, { 'data.txt': 'test\n', 'big.bin': 'same' });
  const next = release({ version: NEXT, packs: [{ manifest: M.lang, content: { 'data.txt': 'new\n', 'big.bin': 'same' } }] });
  const srv = server(next);
  const svc = service({ dir, srv });
  assert.equal((await svc.readiness(NEXT)).offer, true);
  const s = await svc.stageFor(NEXT);
  assert.deepEqual(s.staged, ['lang-xx']);
  assert.ok(!srv.served.some((u) => u.endsWith('big.bin')), 'an unchanged file is not fetched again');
  // The new version's first start.
  const later = { ...contract, office: NEXT };
  const r = applyStaged({ dir, contract: later, office: NEXT, platform: 'win32-x64' });
  assert.deepEqual(r.moved, ['lang-xx']);
  assert.equal(JSON.parse(fs.readFileSync(path.join(dir, 'lang-xx.pack.json'), 'utf8')).manifest.version, NEXT);
});

test('the update service waits for add-ons, unless Update anyway was chosen for that version; a portable copy is only told', async () => {
  const { gateStep } = await import('../apps/desktop/main/update-gate.js');
  assert.equal(gateStep({ answer: { offer: true }, version: NEXT }), 'download');
  assert.equal(gateStep({ answer: { offer: false }, version: NEXT }), 'wait');
  assert.equal(gateStep({ answer: { offer: false }, version: NEXT, anyway: NEXT }), 'download');
  assert.equal(gateStep({ answer: { offer: false }, version: NEXT, anyway: '98.0.0' }), 'wait', 'Update anyway is for one version only');
  assert.equal(gateStep({ answer: { offer: true }, version: NEXT, portable: true }), 'notice');
  assert.equal(gateStep({ answer: null, version: NEXT }), 'wait');
});

test('the installer\'s run checks staged add-ons with the shipped contract, does removals, and says when it could not do all', async () => {
  const { applyStagedRun } = await import('../apps/desktop/main/addons-wiring.js');
  const dir = fresh();
  put(dir, M.dict);
  const person = fresh('person');
  put(path.join(person, OFFICE), M.lang);
  const lines = [];
  const code = applyStagedRun({
    argv: [`--addons-apply-staged=${person}`, '--addons-remove=dict-xx'], isPackaged: true, office: OFFICE, releaseKeys: [],
    env: { RUTBA_ADDONS_DIR: dir }, log: (l) => lines.push(l),
  });
  // The shipped contract trusts no packages key yet, so a pack signed here is refused: checked, never taken on trust.
  assert.equal(code, 1);
  assert.match(lines.join('\n'), /refused lang-xx \(signature\)/);
  assert.ok(!fs.existsSync(path.join(dir, 'lang-xx.pack.json')));
  assert.ok(!fs.existsSync(path.join(dir, 'dict-xx.pack.json')), 'the removal was done');
  assert.ok(!fs.existsSync(path.join(person, OFFICE)), 'the person\'s staging is cleared');
  assert.equal(applyStagedRun({ argv: [], isPackaged: true, office: OFFICE, releaseKeys: [], env: {} }), null, 'any other run is not this');
});

test('adding a language brings the font this computer lacks for it, and the first start of a version looks once (L8)', async () => {
  const font = { id: 'font-xxxx', name: 'Test font', kind: 'font', version: OFFICE, revision: 1, format: 1, files: M.lang.files };
  const rel = release({ version: OFFICE, packs: [{ manifest: M.lang }, { manifest: font }] });
  const asked = [];
  const fontsFor = async ({ tags, installed, offered }) => {
    asked.push({ tags, installed: [...installed], offered });
    return offered.filter((id) => !installed.includes(id));
  };
  const dir = fresh();
  const stores = { settings: settings() };
  const svc = service({ dir, srv: server(rel), stores, languages: { fontsFor } });
  const r = await svc.add({ id: 'lang-xx' });
  assert.deepEqual(r.adding, ['lang-xx']);
  await until(() => fs.existsSync(path.join(dir, 'font-xxxx.pack.json')));
  assert.deepEqual(asked[0].tags, ['lang-xx']);
  assert.deepEqual(asked[0].offered, ['font-xxxx']);
  assert.ok(asked[0].installed.includes('lang-xx'), 'the language counts as here once it is');

  // A voice brings no font check.
  asked.length = 0;
  const other = service({ dir: fresh(), srv: server(release({ version: OFFICE, packs: [{ manifest: M.lang }, { manifest: font }] })), languages: { fontsFor } });
  await other.add({ id: 'font-xxxx' });
  await new Promise((res) => setTimeout(res, 30));
  assert.equal(asked.length, 0);

  // The first start of a version: the languages here are looked at once.
  const here = fresh();
  put(here, { ...M.lang, version: OFFICE });
  const first = service({ dir: here, srv: server(rel), stores: { settings: settings() }, languages: { fontsFor } });
  assert.deepEqual(await first.fontsCheck(), ['font-xxxx']);
  await until(() => fs.existsSync(path.join(here, 'font-xxxx.pack.json')));
  assert.deepEqual(await first.fontsCheck(), [], 'once per version');
  // With no languages here, or no check given, nothing.
  assert.deepEqual(await service({ dir: fresh(), srv: server(rel), languages: { fontsFor } }).fontsCheck(), []);
  assert.deepEqual(await service({ dir: here, srv: server(rel) }).fontsCheck(), []);
});

test('an add-on from a file: a .zip of the envelope and its folder, checked with this Office\'s rules (L5)', async () => {
  const zipOf = (files) => writeZip(Object.entries(files).map(([name, data]) => {
    const e = new ZipEntry({ name, method: 8 });
    e.data = Buffer.from(data);
    return e;
  }));
  const packed = (manifest, { key = packagesKey, content = { 'data.txt': 'test\n' }, extra = {} } = {}) => {
    const m = { ...manifest, files: Object.fromEntries(Object.entries(content).map(([p, b]) => [p, fileEntry(b)])) };
    return { [`${m.id}.pack.json`]: JSON.stringify(envelopeOf(m, key.privateKey)), ...Object.fromEntries(Object.entries(content).map(([p, b]) => [`${m.id}/${p}`, b])), ...extra };
  };
  const at = fresh('zips');
  const file = (name, files) => {
    const f = path.join(at, name);
    fs.writeFileSync(f, files instanceof Buffer ? files : zipOf(files));
    return f;
  };
  const dir = fresh();
  const svc = service({ dir, srv: server() });
  const r = await svc.addFromFile({ file: file('good.zip', packed(M.lang)) });
  assert.equal(r.added, 'lang-xx');
  assert.ok(fs.existsSync(path.join(dir, 'lang-xx.pack.json')), 'put in place');
  assert.ok(svc.list().restart);
  assert.ok(!fs.existsSync(path.join(stagingFor(dir, OFFICE), PARTS, 'from-file-lang-xx')), 'nothing left behind');
  // The next start uses it.
  const next = startAddons({ folder: { dir, scope: 'person' }, contract, office: OFFICE, memoryFile: path.join(fresh(), 's.json'), argv: [], platform: 'win32-x64' });
  assert.deepEqual(next.registry.list().map((a) => a.id), ['lang-xx']);

  const refused = async (name, files, pattern) => {
    const d = fresh();
    await assert.rejects(service({ dir: d, srv: server() }).addFromFile({ file: file(name, files) }), pattern, name);
    assert.ok(!fs.existsSync(path.join(d, 'lang-xx.pack.json')), `${name} not added`);
  };
  await refused('not-a-zip.zip', Buffer.from('hello'), /not an add-on/);
  await refused('two.zip', { ...packed(M.lang), ...packed(M.dict) }, /one add-on/);
  await refused('outside.zip', packed(M.lang, { extra: { 'elsewhere/x.txt': 'x' } }), /may not: elsewhere\/x\.txt/);
  await refused('climbing.zip', packed(M.lang, { extra: { 'lang-xx/../x.txt': 'x' } }), /may not/);
  await refused('forged.zip', packed(M.lang, { key: crypto.generateKeyPairSync('ed25519') }), /not signed by Rutba/);
  const tampered = packed(M.lang);
  tampered['lang-xx/data.txt'] = 'changed\n';
  await refused('tampered.zip', tampered, /file data\.txt has been changed since it was signed/);
  // A file in its folder that the signed manifest does not list is not written, nor anything else.
  await refused('unlisted.zip', packed(M.lang, { extra: { 'lang-xx/unlisted.dll': 'x' } }), /may not: lang-xx\/unlisted\.dll/);
  // Names Windows would read as a device, a stream or trimmed, even signed.
  for (const bad of ['CON.txt', 'fonts/aux', 'data.txt.', 'data.txt ', 'a:b.txt', 'com1.dat']) {
    await refused(`named-${bad.replace(/\W/g, '_')}.zip`, packed(M.lang, { content: { [bad]: 'x' } }), /may not/);
  }

  // Nothing written before the envelope is usable: not even the staging folder.
  const untouched = fresh();
  await assert.rejects(service({ dir: untouched, srv: server() }).addFromFile({ file: file('forged2.zip', packed(M.lang, { key: crypto.generateKeyPairSync('ed25519') })) }));
  assert.ok(!fs.existsSync(stagingFor(untouched, OFFICE)), 'no staging made');

  // Not an older revision over a newer one here.
  const newer = fresh();
  put(newer, { ...M.lang, revision: 2 });
  await assert.rejects(service({ dir: newer, srv: server() }).addFromFile({ file: file('older.zip', packed(M.lang)) }), /newer than the one in this file/);
  assert.equal(JSON.parse(fs.readFileSync(path.join(newer, 'lang-xx.pack.json'), 'utf8')).manifest.revision, 2, 'kept');
  await assert.rejects(service({ dir: null, srv: server() }).addFromFile({ file: file('good2.zip', packed(M.lang)) }), /no add-ons folder/);
});

test('the update gate counts the built-in language in use as its pack: the update waits for it and brings it (L7)', async () => {
  const dir = fresh();
  const languages = { builtInInUse: () => ['lang-xx'] };
  // Nothing added here, but the person reads the windows in a built-in language.
  const without = release({ version: NEXT, packs: [{ manifest: M.dict }] });
  const waits = await service({ dir, srv: server(without), languages }).readiness(NEXT);
  assert.equal(waits.offer, false);
  assert.deepEqual(waits.waitingFor.map((w) => [w.id, w.why]), [['lang-xx', 'unpublished']]);
  // Its pack published for the next version: the update goes ahead, and brings it.
  const withIt = release({ version: NEXT, packs: [{ manifest: M.lang }], sequence: 2 });
  const svc = service({ dir, srv: server(withIt), languages });
  assert.equal((await svc.readiness(NEXT)).offer, true);
  assert.deepEqual((await svc.stageFor(NEXT)).staged, ['lang-xx']);
  const r = applyStaged({ dir, contract: { ...contract, office: NEXT }, office: NEXT, platform: 'win32-x64' });
  assert.deepEqual(r.moved, ['lang-xx']);
  // Added already: counted once, as itself. And a person in English holds nothing.
  assert.equal((await service({ dir, srv: server(release({ version: NEXT, packs: [{ manifest: M.lang }], sequence: 3 })), languages }).readiness(NEXT)).offer, true);
  assert.equal((await service({ dir: fresh(), srv: server(without), languages: { builtInInUse: () => [] } }).readiness(NEXT)).offer, true);
});

test('a feed entry whose id cannot be an add-on\'s is never offered, added or made into a folder', async () => {
  const dir = fresh();
  const bad = { ...M.dict, id: 'x/../../escape', requires: [] };
  const rel = release({ version: OFFICE, packs: [{ manifest: M.lang }, { manifest: bad }] });
  const svc = service({ dir, srv: server(rel) });
  const l = await svc.feed();
  assert.deepEqual(l.offered.map((p) => p.id), ['lang-xx']);
  await assert.rejects(svc.add({ id: bad.id }), /not an add-on/);
  assert.ok(!fs.existsSync(path.join(dir, '..', 'escape')) && !fs.existsSync(path.join(stagingFor(dir, OFFICE), '..', 'escape')));
});

test('the names an add-on may list for its files, as Windows writes them', async () => {
  const { safePath, isAddonId } = await import('../packages/office-shell/src/electron/addons-compat.js');
  for (const ok of ['data.txt', 'fonts/Noto Nastaliq.ttf', 'model/v1.0/weights.bin', 'console.txt', 'auxiliary/a', 'com10.txt']) assert.ok(safePath(ok), ok);
  for (const no of ['', '/abs', 'a\b', 'C:x', 'a:stream', 'a/../b', './a', 'a//b', 'a.', 'a ', 'CON', 'con.txt', 'Nul.dat', 'COM1', 'lpt9.log', 'dir/PRN/x', 'conin$', 'a<b', 'a|b', 'a?b', 'a*b', 'a"b', 'tab\there', 'x'.repeat(241)]) assert.ok(!safePath(no), JSON.stringify(no));
  assert.ok(isAddonId('lang-ur') && !isAddonId('Lang') && !isAddonId('a/b') && !isAddonId(7) && !isAddonId('x'.repeat(33)));
});
