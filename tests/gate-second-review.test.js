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

test('with nothing on record there is nothing to fall below', () => {
  assert.deepEqual(belowRecorded(floorsNow(counts, state), null), []);
});

test('fewer checks by hand are said in the release note', () => {
  const note = gateSection(counts, { version: '9.9.9', date: '2026-11-10', platform: 'win32', arch: 'x64', fewer: ['mail window checks: 40, down from 50'] });
  assert.match(note.replace(/\s+/g, ' '), /Fewer checks than the gate before, allowed by hand: mail window checks: 40, down from 50/);
  assert.doesNotMatch(gateSection(counts, { version: '9.9.9', date: '2026-11-10', platform: 'win32', arch: 'x64' }), /Fewer checks/);
});
