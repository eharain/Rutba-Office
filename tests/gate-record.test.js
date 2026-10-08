// The gate puts its own counts on the record: what each pass counted is read
// from the lines it ends with, written as the release note's Gate section in
// the note's plain style, and brought into the README's counts.
import test from 'node:test';
import assert from 'node:assert/strict';
import { countsOf, gateSection, withGateSection, withReadmeCounts, localDate, wrap } from '../tools/gate.js';

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
  const readme = 'npm test          # 1,807 tests, no network needed\n\nnpm test                # the engine suite — 1,807 checks, no windows\n\n`npm run gate` is what runs before a release: 1,807 engine tests, 12 editing checks,\n1,174 application checks — most of them pressing the ribbon — and 10 window captures.\n';
  const out = withReadmeCounts(readme, counts);
  assert.match(out, /# 1,870 tests, no network needed/);
  assert.match(out, /the engine suite — 1,870 checks/);
  assert.match(out, /before a release: 1,870 engine tests, 12 editing checks,\n1,190 application checks/);
  assert.match(out, /and 10 window captures/);
});
