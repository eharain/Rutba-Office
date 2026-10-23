// The gate puts its own counts on the record: what each pass counted is read
// from the lines it ends with, written as the release note's Gate section in
// the note's plain style, and brought into the README's counts.
import test from 'node:test';
import assert from 'node:assert/strict';
import { countsOf, engineCounts, gateSection, withGateSection, withReadmeCounts, localDate, wrap, isPatch } from '../tools/gate.js';

const counts = {
  test: countsOf('test', 'ℹ tests 1800\nℹ pass 1794\nℹ fail 0\nℹ skipped 6\n...\nℹ tests 70\nℹ pass 70\nℹ fail 0\nℹ skipped 0\n'),
  'verify:edit': countsOf('verify:edit', '\u001b[32mok\u001b[0m\n12/12 editing checks passed\n'),
  'verify:apps': countsOf('verify:apps', 'ok   home: ...\n1190/1190 application checks passed\n'),
  smoke: countsOf('smoke', 'ok   home  2494 ms\n10/10 windows rendered clean\n'),
};

test('each pass\'s counts are read from its last lines, the two test runs added together', () => {
  assert.deepEqual(counts.test, { tests: 1870, passed: 1864, failed: 0, skipped: 6 });
  assert.deepEqual(counts['verify:edit'], { passed: 12, of: 12 });
  assert.deepEqual(counts['verify:apps'], { passed: 1190, of: 1190 });
  assert.deepEqual(counts.smoke, { passed: 10, of: 10 });
  assert.equal(countsOf('verify:apps', 'nothing counted'), null);
});

test('the Gate section is written into the note, or written again in its place, with no dashes', () => {
  const section = gateSection(counts, { version: '1.30.0', date: '2026-10-09', platform: 'win32', arch: 'x64' });
  assert.equal(section, `## Gate\n\n${wrap('Run on Windows (x64) on 2026-10-09, before 1.30.0 was released: 1,870 engine tests, 1,864 passing and 6 skipped where a package they need is not installed; 12 of 12 editing checks; 1,190 of 1,190 application checks; and 10 of 10 windows painted clean.')}\n`);
  assert.ok(section.split('\n').every((l) => l.length <= 75), 'wrapped as the notes are');
  assert.equal(localDate(new Date(2026, 9, 9, 0, 30)), '2026-10-09', 'the day as this computer has it, past midnight');
  assert.doesNotMatch(section, /[–—]/);
  const note = 'Lead.\n\n## Fixed\n\n- A thing.\n';
  const once = withGateSection(note, section);
  assert.equal(once, `${note}\n${section}`);
  const later = gateSection({ ...counts, smoke: { passed: 9, of: 10 } }, { version: '1.30.0', date: '2026-10-10', platform: 'win32', arch: 'x64' });
  const again = withGateSection(`${once}\n## After\n\nMore.\n`, later);
  assert.equal((again.match(/## Gate/g) || []).length, 1, 'put in again, not twice');
  assert.match(again, /9 of 10 windows/);
  assert.match(again, /## After\n\nMore\./, 'what follows is kept');
});

test('the README\'s counts are brought to what the gate counted', () => {
  const readme = 'npm test          # 1,807 tests, no network needed\n\nnpm test                # the engine suite: 1,807 checks, no windows\n\n`npm run gate` is what runs before a release: 1,807 engine tests, 12 editing checks,\n1,174 application checks (most of them pressing the ribbon) and 10 window captures.\n';
  const out = withReadmeCounts(readme, counts);
  assert.match(out, /# 1,870 tests, no network needed/);
  assert.match(out, /the engine suite: 1,870 checks/);
  assert.match(out, /before a release: 1,870 engine tests, 12 editing checks,\n1,190 application checks/);
  assert.match(out, /and 10 window captures/);
});

test('the TAP reporter\'s counts are read as the spec reporter\'s are, as a gate run into a log prints them', () => {
  assert.deepEqual(countsOf('test', '# Subtest: x\nok 1 - x\n1..2209\n# tests 2209\n# suites 0\n# pass 2203\n# fail 0\n# cancelled 0\n# skipped 6\n'), { tests: 2209, passed: 2203, failed: 0, skipped: 6 });
  assert.deepEqual(countsOf('test:perf', '# tests 4\n# pass 4\n# fail 0\n# skipped 0\n'), { tests: 4, passed: 4, failed: 0, skipped: 0 });
});

test('the budgets, run alone after the suite, count among the engine tests but not in npm test\'s own count', () => {
  const withPerf = { ...counts, 'test:perf': countsOf('test:perf', 'ℹ tests 4\nℹ pass 4\nℹ fail 0\nℹ skipped 0\n') };
  assert.deepEqual(withPerf['test:perf'], { tests: 4, passed: 4, failed: 0, skipped: 0 });
  assert.deepEqual(engineCounts(withPerf), { tests: 1874, passed: 1868, failed: 0, skipped: 6 });
  assert.match(gateSection(withPerf, { version: '1.43.0', date: '2026-10-12', platform: 'win32', arch: 'x64' }), /1,874\s+engine\s+tests,\s+1,868\s+passing/);
  const readme = 'npm test          # 1,807 tests, no network needed\n\n`npm run gate` is what runs before a release: 1,807 engine tests, 12 editing checks,\n';
  const out = withReadmeCounts(readme, withPerf);
  assert.match(out, /# 1,870 tests, no network needed/);
  assert.match(out, /before a release: 1,874 engine tests/);
});

test('a patch release goes on the record with the apps it changed, a minor release only in full (D-2026-10-23-01)', () => {
  assert.equal(isPatch('2.17.1'), true);
  assert.equal(isPatch('2.17.12'), true);
  assert.equal(isPatch('2.17.0'), false);
  assert.equal(isPatch('2.18.0-dev'), false);
  const counts = { test: { tests: 3000, passed: 2991, failed: 0, skipped: 9 }, 'verify:apps': { passed: 120, of: 120 }, smoke: { passed: 14, of: 14 } };
  assert.match(gateSection(counts, { version: '2.17.2', date: '2026-10-23', platform: 'win32', arch: 'x64', apps: ['base', 'mail'] }), /120 of 120 application checks, of base and mail, the apps it\s+changed/);
  assert.match(gateSection(counts, { version: '2.17.2', date: '2026-10-23', platform: 'win32', arch: 'x64', apps: ['base'] }), /of base, the app it changed/);
  assert.doesNotMatch(gateSection(counts, { version: '2.18.0', date: '2026-10-23', platform: 'win32', arch: 'x64' }), /it changed/);
});
