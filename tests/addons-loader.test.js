// The add-ons on disk (packages/office-shell/src/electron/addons.js) and the
// folder that holds them (apps/desktop/main/addons-folder.js): which are
// used, which are refused and why, the previous kept for rollback, what
// each requires, removals and the lock. The add-ons are made and signed for
// the test with a key thrown away with its folder.
import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { findAddons, noAddons, checkFolder } from '../packages/office-shell/src/electron/addons.js';
import { findPacks } from '../packages/office-shell/src/electron/packs.js';
import { envelopeOf, fileEntry, manifests, CONTRACT_FILE } from '../tools/make-addon-vectors.mjs';
import {
  addonsFolder, lockFolder, readState, writeState, markRemoval, unmarkRemoval, applyRemovals, rollBack, putInPlace, forgetPrevious, stagingFor, canWrite,
} from '../apps/desktop/main/addons-folder.js';
import { applyStaged } from '../apps/desktop/main/addons-start.js';

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-addons-'));
test.after(() => fs.rmSync(tmp, { recursive: true, force: true }));
const file = JSON.parse(fs.readFileSync(CONTRACT_FILE, 'utf8'));
const OFFICE = file.office;
const packagesKey = crypto.generateKeyPairSync('ed25519');
const releaseKey = crypto.generateKeyPairSync('ed25519');
const pem = (k) => k.publicKey.export({ type: 'spki', format: 'pem' }).toString();
const contract = { ...file, keys: { packages: [pem(packagesKey)], release: [pem(releaseKey)] } };
const M = manifests(contract);

let n = 0;
const freshDir = () => {
  const d = path.join(tmp, `f${(n += 1)}`);
  fs.mkdirSync(d, { recursive: true });
  return d;
};

/** Writes an add-on into a folder: its files from `content`, the manifest listing them, the envelope signed. */
function put(dir, manifest, { content = { 'data.txt': 'test\n' }, key = manifest.kind === 'engine' ? releaseKey : packagesKey, at = dir } = {}) {
  const files = Object.fromEntries(Object.entries(content).map(([p, b]) => [p, fileEntry(b)]));
  const m = { ...manifest, files };
  const folder = path.join(at, m.id);
  fs.rmSync(folder, { recursive: true, force: true });
  for (const [p, b] of Object.entries(content)) {
    fs.mkdirSync(path.dirname(path.join(folder, p)), { recursive: true });
    fs.writeFileSync(path.join(folder, p), b);
  }
  fs.writeFileSync(path.join(at, `${m.id}.pack.json`), JSON.stringify(envelopeOf(m, key.privateKey)));
  return m;
}
const engineContent = { 'main/index.js': 'export function register() { return {}; }\n' };
const find = (dir, more = {}) => findAddons({ dir, contract, office: OFFICE, platform: 'win32-x64', ...more });
const ids = (r) => r.list().map((a) => a.id);
const codes = (r) => Object.fromEntries(r.refused.map((x) => [x.id, x.code]));

test('no folder, or an empty one, has no add-ons', () => {
  assert.deepEqual(ids(find(null)), []);
  assert.deepEqual(ids(find(path.join(tmp, 'never-made'))), []);
  assert.deepEqual(ids(noAddons()), []);
  assert.equal(noAddons().engine(), null);
});

test('good add-ons are used, found by id or by what they provide', () => {
  const dir = freshDir();
  put(dir, M.lang);
  put(dir, M.dict);
  const r = find(dir);
  assert.deepEqual(ids(r), ['dict-xx', 'lang-xx']);
  assert.deepEqual(r.refused, []);
  assert.equal(r.find('language:xx').id, 'lang-xx');
  assert.equal(r.find('lang-xx').kind, 'language');
  assert.equal(r.find('nothing'), null);
  assert.deepEqual([...r.installed].sort(), ['dict-xx', 'lang-xx']);
  assert.ok(Object.isFrozen(r.list()[0]));
});

test('speech data is used only beside an engine that reads it', () => {
  const dir = freshDir();
  put(dir, M.model);
  put(dir, M.voice);
  assert.deepEqual(codes(find(dir)), { 'model-test': 'requires', 'voice-test': 'requires' });
  put(dir, M.engine, { content: engineContent });
  const r = find(dir);
  assert.deepEqual(ids(r), ['model-test', 'speech-engine', 'voice-test']);
  assert.equal(r.engine().id, 'speech-engine');
  // An engine for another computer is not one, and its data waits.
  assert.deepEqual(codes(find(dir, { platform: 'linux-x64' })), { 'model-test': 'requires', 'speech-engine': 'version', 'voice-test': 'requires' });
});

test('the engine signed with the packages key is refused, and a language with the release key', () => {
  const dir = freshDir();
  put(dir, M.engine, { content: engineContent, key: packagesKey });
  put(dir, M.lang, { key: releaseKey });
  assert.deepEqual(codes(find(dir)), { 'lang-xx': 'key-for-kind', 'speech-engine': 'key-for-kind' });
});

test('a changed, missing or extra file, or a link, refuses the add-on', () => {
  const dir = freshDir();
  put(dir, M.lang);
  fs.writeFileSync(path.join(dir, 'lang-xx', 'data.txt'), 'TEST\n');
  assert.deepEqual(codes(find(dir)), { 'lang-xx': 'files' });

  put(dir, M.lang);
  fs.writeFileSync(path.join(dir, 'lang-xx', 'extra.js'), 'nothing');
  assert.deepEqual(codes(find(dir)), { 'lang-xx': 'files' });

  put(dir, M.lang);
  fs.rmSync(path.join(dir, 'lang-xx', 'data.txt'));
  assert.deepEqual(codes(find(dir)), { 'lang-xx': 'files' });

  put(dir, M.lang);
  fs.rmSync(path.join(dir, 'lang-xx'), { recursive: true });
  assert.deepEqual(codes(find(dir)), { 'lang-xx': 'files' });

  put(dir, M.lang);
  const outside = freshDir();
  try {
    fs.symlinkSync(outside, path.join(dir, 'lang-xx', 'more'), 'junction');
  } catch {
    return; // no links on this computer
  }
  assert.deepEqual(codes(find(dir)), { 'lang-xx': 'link' });
});

test('a model is checked by size at start, its bytes when read', () => {
  const dir = freshDir();
  put(dir, M.engine, { content: engineContent });
  put(dir, M.model);
  fs.writeFileSync(path.join(dir, 'model-test', 'data.txt'), 'TEST\n');
  assert.deepEqual(codes(find(dir)), {});
  fs.writeFileSync(path.join(dir, 'model-test', 'data.txt'), 'TESTS\n');
  assert.deepEqual(codes(find(dir)), { 'model-test': 'files' });
});

test('a manifest naming a path outside its folder is refused', () => {
  const dir = freshDir();
  for (const bad of ['../x', '/x', 'a\\b', 'C:/x', 'a//b', './a']) {
    assert.equal(checkFolder(dir, { id: 'x', kind: 'language', files: { [bad]: { size: 1, sha512: 'x' } } })?.code, 'files', bad);
  }
});

test('an envelope for another id is refused', () => {
  const dir = freshDir();
  put(dir, M.lang);
  fs.renameSync(path.join(dir, 'lang-xx.pack.json'), path.join(dir, 'lang-yy.pack.json'));
  fs.renameSync(path.join(dir, 'lang-xx'), path.join(dir, 'lang-yy'));
  assert.deepEqual(codes(find(dir)), { 'lang-yy': 'envelope' });
});

test('a pack whose need is not usable waits, and so does one needing it', () => {
  const dir = freshDir();
  put(dir, M.dict);
  put(dir, { ...M.lang, id: 'lang-zz', provides: [], requires: ['dict-xx'] });
  const r = find(dir);
  assert.deepEqual(codes(r), { 'dict-xx': 'requires', 'lang-zz': 'requires' });
  assert.deepEqual(ids(r), []);
});

test('the highest version that passes is used, the other kept as the previous, and rollback puts it back', () => {
  const dir = freshDir();
  put(dir, { ...M.lang, revision: 1 });
  const staged = path.join(stagingFor(dir, OFFICE), 'lang-xx-new');
  fs.mkdirSync(staged, { recursive: true });
  put(dir, { ...M.lang, revision: 2 }, { at: staged, content: { 'data.txt': 'two\n' } });
  putInPlace(dir, staged, 'lang-xx');
  let r = find(dir);
  assert.equal(r.find('lang-xx').manifest.revision, 2);
  assert.deepEqual(r.find('lang-xx').previous, { version: OFFICE, revision: 1 });

  // The new one broken: the previous is used, and nothing is reported as a fault.
  fs.writeFileSync(path.join(dir, 'lang-xx', 'data.txt'), 'broken\n');
  r = find(dir);
  assert.equal(r.find('lang-xx').manifest.revision, 1);
  assert.equal(r.find('lang-xx').slot, 'previous');
  assert.deepEqual(r.refused, []);

  assert.ok(rollBack(dir, 'lang-xx'));
  r = find(dir);
  assert.equal(r.find('lang-xx').manifest.revision, 1);
  assert.equal(r.find('lang-xx').previous, null);
  assert.equal(rollBack(dir, 'lang-xx'), false, 'nothing kept to go back to');
});

test('a staged add-on with no envelope is never put in place, and the previous can be forgotten', () => {
  const dir = freshDir();
  put(dir, M.lang);
  const staged = freshDir();
  fs.mkdirSync(path.join(staged, 'lang-xx'));
  assert.throws(() => putInPlace(dir, staged, 'lang-xx'), /not staged whole/);
  put(dir, { ...M.lang, revision: 3 }, { at: staged });
  putInPlace(dir, staged, 'lang-xx');
  assert.ok(fs.existsSync(path.join(dir, 'lang-xx.previous')));
  forgetPrevious(dir, 'lang-xx');
  assert.ok(!fs.existsSync(path.join(dir, 'lang-xx.previous')));
});

test('a folder that cannot be moved (a file in it open) leaves the add-on in place whole and the staged one staged', { skip: process.platform !== 'win32' && 'Windows refuses to rename a folder with a file open in it' }, () => {
  const dir = freshDir();
  put(dir, { ...M.lang, revision: 1 });
  const staged = freshDir();
  put(dir, { ...M.lang, revision: 2 }, { at: staged, content: { 'data.txt': 'two\n' } });
  const fd = fs.openSync(path.join(dir, 'lang-xx', 'data.txt'), 'r');
  try {
    assert.throws(() => putInPlace(dir, staged, 'lang-xx'));
    assert.ok(fs.existsSync(path.join(dir, 'lang-xx.pack.json')), 'the envelope in place is kept');
    assert.ok(fs.existsSync(path.join(dir, 'lang-xx', 'data.txt')), 'and its folder');
    assert.ok(fs.existsSync(path.join(staged, 'lang-xx.pack.json')) && fs.existsSync(path.join(staged, 'lang-xx')), 'the staged one is still staged');
    assert.equal(find(dir).find('lang-xx').manifest.revision, 1);
  } finally {
    fs.closeSync(fd);
  }
  putInPlace(dir, staged, 'lang-xx');
  assert.equal(find(dir).find('lang-xx').manifest.revision, 2, 'moved in once nothing holds it');
});

test('a staged add-on that could not be moved in now stays staged for the next start', { skip: process.platform !== 'win32' && 'Windows refuses to rename a folder with a file open in it' }, () => {
  const dir = freshDir();
  put(dir, { ...M.lang, revision: 1 });
  put(stagingFor(dir, OFFICE), { ...M.lang, revision: 2 }, { content: { 'data.txt': 'two\n' } });
  const fd = fs.openSync(path.join(dir, 'lang-xx', 'data.txt'), 'r');
  let r;
  try {
    r = applyStaged({ dir, contract, office: OFFICE });
  } finally {
    fs.closeSync(fd);
  }
  assert.deepEqual(r.kept, ['lang-xx']);
  assert.deepEqual(r.moved, []);
  assert.ok(fs.existsSync(path.join(stagingFor(dir, OFFICE), 'lang-xx.pack.json')), 'kept staged');
  assert.equal(find(dir).find('lang-xx').manifest.revision, 1, 'the one in place is whole');
  r = applyStaged({ dir, contract, office: OFFICE });
  assert.deepEqual(r.moved, ['lang-xx']);
  assert.equal(find(dir).find('lang-xx').manifest.revision, 2);
});

test('the administrator\'s step waits for Office to quit, then starts it again as the person', async () => {
  const { waitForExit, applyStagedRun } = await import('../apps/desktop/main/addons-wiring.js');
  let t = 0;
  let left = 3;
  assert.equal(waitForExit(1, 60_000, { alive: () => left-- > 0, now: () => t }), true);
  assert.equal(waitForExit(1, 1000, { alive: () => true, now: () => (t += 600) }), false, 'not for ever');
  const order = [];
  const dir = freshDir();
  applyStagedRun({
    argv: ['--addons-apply-staged=', '--addons-after=99', '--addons-relaunch'], isPackaged: true, office: OFFICE, releaseKeys: [],
    env: { RUTBA_ADDONS_DIR: dir }, log: (l) => order.push(`log ${l.split(' ')[0]}`),
    wait: (pid) => { order.push(`wait ${pid}`); return true; }, relaunch: () => order.push('relaunch'),
  });
  assert.equal(order[0], 'wait 99');
  assert.equal(order[order.length - 1], 'relaunch');
});

test('removals asked for go at the next start, envelope first; ids skipped are not loaded', () => {
  const dir = freshDir();
  put(dir, M.lang);
  put(dir, M.dict);
  markRemoval(dir, 'dict-xx');
  markRemoval(dir, 'dict-xx');
  assert.deepEqual(readState(dir).removals, ['dict-xx']);
  assert.deepEqual(ids(find(dir, { skip: readState(dir).removals })), ['lang-xx']);
  unmarkRemoval(dir, 'dict-xx');
  markRemoval(dir, 'dict-xx');
  assert.deepEqual(applyRemovals(dir), { removed: ['dict-xx'], held: [] });
  assert.ok(!fs.existsSync(path.join(dir, 'dict-xx.pack.json')));
  assert.ok(!fs.existsSync(path.join(dir, 'dict-xx')));
  assert.deepEqual(readState(dir).removals, []);
  assert.throws(() => markRemoval(dir, '../x'));
});

test('the state keeps the highest feed seen per version, and a damaged one reads as empty', () => {
  const dir = freshDir();
  assert.deepEqual(readState(dir), { seen: {}, removals: [], feeds: {} });
  writeState(dir, { seen: { '2.25.1': 4 }, removals: [] });
  assert.deepEqual(readState(dir).seen, { '2.25.1': 4 });
  fs.writeFileSync(path.join(dir, 'state.json'), '{ not json');
  assert.deepEqual(readState(dir), { seen: {}, removals: [], feeds: {} });
  fs.writeFileSync(path.join(dir, 'state.json'), JSON.stringify({ seen: { a: 'x', b: 2 }, removals: ['ok', '../no'] }));
  assert.deepEqual(readState(dir), { seen: { b: 2 }, removals: ['ok'], feeds: {} });
});

test('the lock keeps a second process out, and is taken over from one that has gone', () => {
  const dir = freshDir();
  const release = lockFolder(dir, { pid: 1001, isAlive: () => true });
  assert.equal(typeof release, 'function');
  assert.equal(lockFolder(dir, { pid: 1002, isAlive: () => true }), null);
  const taken = lockFolder(dir, { pid: 1002, isAlive: () => false });
  assert.equal(typeof taken, 'function');
  release(); // the first no longer holds it, so the second keeps it
  assert.ok(fs.existsSync(path.join(dir, '.lock')));
  taken();
  assert.ok(!fs.existsSync(path.join(dir, '.lock')));
  assert.ok(canWrite(dir));
});

test('the folder is by install: for this person, for everyone, portable, or named', () => {
  const env = { LOCALAPPDATA: 'C:\\Users\\p\\AppData\\Local', ProgramData: 'C:\\ProgramData', ProgramFiles: 'C:\\Program Files' };
  const win = (execPath, more = {}) => addonsFolder({ execPath, isPackaged: true, env: { ...env, ...more }, platform: 'win32' });
  assert.deepEqual(win('C:\\Users\\p\\AppData\\Local\\Programs\\Rutba Office\\Rutba Office.exe'), { dir: path.join(env.LOCALAPPDATA, 'Rutba Office', 'addons'), scope: 'person' });
  assert.deepEqual(win('C:\\Program Files\\Rutba Office\\Rutba Office.exe'), { dir: path.join(env.ProgramData, 'Rutba Office', 'addons'), scope: 'machine' });
  assert.deepEqual(win('E:\\Office\\Rutba Office.exe', { PORTABLE_EXECUTABLE_DIR: 'E:\\Office' }), { dir: path.join('E:\\Office', 'addons'), scope: 'portable' });
  assert.equal(addonsFolder({ execPath: 'x', isPackaged: false, env: {} }), null, 'a copy run from the source has none');
  assert.equal(addonsFolder({ execPath: 'x', isPackaged: false, env: { RUTBA_ADDONS_DIR: tmp } }).scope, 'named');
});

test('an add-on put among the app packs is refused for its kind', () => {
  const dir = freshDir();
  put(dir, M.dict);
  const r = findPacks({ dir, publicKey: packagesKey.publicKey, version: OFFICE });
  assert.deepEqual(r.refused.map((x) => [x.id, x.code]), [['dict-xx', 'kind']]);
});
