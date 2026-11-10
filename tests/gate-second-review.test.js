// The gate's own honesty, after the second review of it.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { VERIFY_APPS, SHRINKING_ENV, porcelain, dirtyForRecord, wantedForChanged, machineCovers, floorsNow, belowRecorded, gateSection } from '../tools/gate.js';

test('a machine lane that ran for one app does not stand for all of them', () => {
  const one = { ok: true, of: 9, scope: ['word'] };
  assert.equal(machineCovers(one, ['word']), true);
  assert.equal(machineCovers(one, VERIFY_APPS), false);
});

test('a machine lane that failed, is missing, or came from --serial stands for nothing', () => {
  assert.equal(machineCovers({ ok: false, scope: VERIFY_APPS }, ['word']), false);
  assert.equal(machineCovers(null, ['word']), false);
  assert.equal(machineCovers({ ok: true }, ['word']), false, 'a state from before scopes were kept');
  assert.equal(machineCovers({ ok: true, serial: true, scope: VERIFY_APPS }, VERIFY_APPS), false);
});

test('a real change not staged is seen whole, the first line included', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-gate-'));
  try {
    const git = (...a) => execFileSync('git', ['-c', 'user.name=t', '-c', 'user.email=t@t', '-c', 'core.hooksPath=', '-c', 'core.autocrlf=false', ...a], { cwd: dir, encoding: 'utf8' });
    git('init', '-q');
    fs.mkdirSync(path.join(dir, 'docs', 'releases'), { recursive: true });
    fs.writeFileSync(path.join(dir, 'docs', 'releases', 'v2.0.1.md'), 'a');
    fs.writeFileSync(path.join(dir, 'other.txt'), 'a');
    git('add', '.');
    git('commit', '-q', '-m', 'x');
    fs.writeFileSync(path.join(dir, 'docs', 'releases', 'v2.0.1.md'), 'b');
    const said = porcelain(dir);
    assert.ok(said.startsWith(' M '), JSON.stringify(said));
    assert.deepEqual(dirtyForRecord(said, '2.0.1'), [], 'the note a record writes does not stop the record');
    fs.writeFileSync(path.join(dir, 'other.txt'), 'b');
    assert.equal(dirtyForRecord(porcelain(dir), '2.0.1').length, 1);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('--changed with no tag to compare with checks every app', () => {
  assert.deepEqual(wantedForChanged('', []), VERIFY_APPS);
  assert.deepEqual(wantedForChanged('v2.0.0', ['docs/a.md']), []);
});

test('a shell that shortens the corpus cannot shorten the gate', () => {
  assert.ok(SHRINKING_ENV.includes('RUTBA_CORPUS_LIMIT'));
});

const counts = {
  test: { tests: 100, passed: 96, failed: 0, skipped: 4 },
  'test:perf': { tests: 10, passed: 10, failed: 0, skipped: 0 },
  'fuzz:open': { passed: 300, of: 300 },
  corpus: { passed: 147, of: 147 },
  'verify:packs': { passed: 20, of: 20 },
};
const state = { apps: { mail: { of: 50 }, word: { of: 30 } }, machine: { of: 12 } };

test('one app losing checks is seen even when another gained more', () => {
  const was = floorsNow(counts, state);
  const now = floorsNow(counts, { apps: { mail: { of: 40 }, word: { of: 90 } }, machine: { of: 12 } });
  assert.deepEqual(belowRecorded(now, was), ['mail window checks: 40, down from 50']);
});

test('a skipped test does not count as run', () => {
  const was = floorsNow(counts, state);
  const more = { ...counts, test: { tests: 100, passed: 90, failed: 0, skipped: 10 } };
  assert.deepEqual(belowRecorded(floorsNow(more, state), was), ['engine tests that ran rather than skipped: 100, down from 106']);
});

test('fuzzing, app packs and the corpus each have a floor, and a part that stopped running falls through it', () => {
  const was = floorsNow(counts, state);
  const { corpus, ...rest } = counts;
  assert.ok(corpus);
  assert.deepEqual(belowRecorded(floorsNow(rest, state), was), ['corpus: none, down from 147']);
});

test('a --serial run is held to the window checks as a whole, its machine checks being inside the apps', () => {
  const was = floorsNow(counts, state);
  const serial = (mail, word) => floorsNow(counts, { apps: { mail: { of: mail }, word: { of: word } }, machine: { ok: true, of: 0, serial: true } });
  assert.deepEqual(belowRecorded(serial(56, 36), was), [], 'the twelve machine checks ran inside mail and word');
  assert.deepEqual(belowRecorded(serial(50, 36), was), ["window checks, the machine lane's among them: 86, down from 92"]);
  assert.deepEqual(belowRecorded(serial(45, 47), was), ['mail window checks: 45, down from 50'], 'one app still may not lose what another gained');
  // A run in lanes still holds the machine lane on its own.
  assert.deepEqual(belowRecorded(floorsNow(counts, { apps: { mail: { of: 56 }, word: { of: 36 } }, machine: { of: 0 } }), was), ['machine lane: 0, down from 12']);
});

test('the key --resume keeps changes with any byte of the tree: untracked files and binary ones too', async () => {
  const { treeKey } = await import('../tools/gate.js');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-gate-key-'));
  try {
    const git = (...a) => execFileSync('git', ['-c', 'user.name=t', '-c', 'user.email=t@t', '-c', 'core.hooksPath=', '-c', 'core.autocrlf=false', ...a], { cwd: dir, encoding: 'utf8' });
    git('init', '-q');
    fs.writeFileSync(path.join(dir, 'picture.png'), Buffer.from([0x89, 0x50, 0x4e, 0x47, 0, 1, 2, 3]));
    git('add', '.');
    git('commit', '-q', '-m', 'x');
    const keys = new Set([treeKey(dir)]);
    const changed = (what) => { const k = treeKey(dir); assert.ok(!keys.has(k), what); keys.add(k); };
    fs.writeFileSync(path.join(dir, 'picture.png'), Buffer.from([0x89, 0x50, 0x4e, 0x47, 0, 1, 2, 4]));
    changed('a binary file changed');
    fs.writeFileSync(path.join(dir, 'picture.png'), Buffer.from([0x89, 0x50, 0x4e, 0x47, 0, 1, 2, 5]));
    changed('the same binary file changed again: the status line is the same');
    fs.writeFileSync(path.join(dir, 'new.js'), 'one');
    changed('an untracked file added');
    fs.writeFileSync(path.join(dir, 'new.js'), 'two');
    changed('the untracked file edited: its name and status are the same');
    assert.equal(treeKey(dir), treeKey(dir), 'the same tree gives the same key');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('with nothing on record there is nothing to fall below', () => {
  assert.deepEqual(belowRecorded(floorsNow(counts, state), null), []);
});

test('fewer checks by hand are said in the release note', () => {
  const note = gateSection(counts, { version: '9.9.9', date: '2026-11-10', platform: 'win32', arch: 'x64', fewer: ['mail window checks: 40, down from 50'] });
  assert.match(note.replace(/\s+/g, ' '), /Fewer checks than the gate before, allowed by hand: mail window checks: 40, down from 50/);
  assert.doesNotMatch(gateSection(counts, { version: '9.9.9', date: '2026-11-10', platform: 'win32', arch: 'x64' }), /Fewer checks/);
});

test('what makes the shipped build means every app; a test or a note does not', async () => {
  const { appsForPath } = await import('../tools/gate.js');
  for (const f of ['tools/make-packs.mjs', 'tools/sbom.mjs', 'tools/publish-release.js', '.github/workflows/release.yml']) assert.deepEqual(appsForPath(f), ['shell'], f);
  for (const f of ['tools/gate.js', 'tests/a.test.js', 'docs/a.md', 'README.md']) assert.deepEqual(appsForPath(f), [], f);
});

test('what is installed is held to the lock file', async () => {
  const { lockDrift } = await import('../tools/gate.js');
  const lock = { packages: { '': { version: '1.0.0' }, 'node_modules/a': { version: '1.0.0' }, 'node_modules/b': { version: '2.0.0' }, 'node_modules/c': { version: '1.0.0', optional: true }, 'node_modules/d': { version: '1.0.0' }, 'node_modules/@x/w': { link: true } } };
  const here = { 'node_modules/a': '1.0.0', 'node_modules/b': '2.1.0' };
  assert.deepEqual(lockDrift(lock, (at) => here[at] ?? null), ['node_modules/b is 2.1.0, and the lock file says 2.0.0', 'node_modules/d is missing']);
});

test('a floors file taken away stops the record, and one brought up to date does not', () => {
  assert.equal(dirtyForRecord(' D tools/gate-floors.json\n', '1.0.0').length, 1);
  assert.equal(dirtyForRecord(' M tools/gate-floors.json\n', '1.0.0').length, 0);
});

test('every setting that narrows a window check is cleared for a gate', () => {
  for (const name of ['RUTBA_A11Y_APPS', 'RUTBA_DARK_APPS', 'RUTBA_POLISH_APPS', 'RUTBA_RIBBON_APPS', 'RUTBA_SCREENS_ONLY']) assert.ok(SHRINKING_ENV.includes(name), name);
});

test('a patch that changed no app says nothing of application checks', () => {
  const note = gateSection({ test: { tests: 1, passed: 1, failed: 0, skipped: 0 }, 'verify:apps': { passed: 5, of: 5 } }, { version: '1.0.1', date: '2026-11-10', platform: 'win32', arch: 'x64', apps: [] });
  assert.doesNotMatch(note, /undefined|application checks/);
});
