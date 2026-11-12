// What makes a gate's "passed" mean something (peer review of the testing
// process, 10 Nov 2026): every part of the window checks belongs to a chunk
// that runs, a run cannot quietly be smaller than the last, and what goes on
// the record names the commit that was checked.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { VERIFY_APPS, SHRINKING_ENV, RERUN_LIMIT, floorsFrom, belowFloor, dirtyForRecord, gateSection, appsForPath, wrap } from '../tools/gate.js';
import { gatedCommitOf, gateProblem } from '../tools/publish-release.js';

const SHA = 'a'.repeat(40);

test('every app a window check names is a chunk the gate runs', () => {
  const source = fs.readFileSync(new URL('../apps/desktop/main/verify-apps.js', import.meta.url), 'utf8');
  const named = new Set([...source.matchAll(/\bpart\('([a-z]+)'/g)].map((m) => m[1]));
  assert.ok(named.size > 10, 'the scan found the parts');
  assert.deepEqual([...named].filter((app) => !VERIFY_APPS.includes(app)), []);
});

test('the Locker has its own chunk, and its own files name it', () => {
  assert.ok(VERIFY_APPS.includes('locker'));
  assert.deepEqual(appsForPath('apps/desktop/main/extras/locker/format.js'), ['locker']);
  assert.deepEqual(appsForPath('apps/desktop/renderer/apps/locker.js'), ['locker']);
  assert.deepEqual(appsForPath('apps/desktop/main/verify-locker.js'), ['locker']);
});

test('what would shrink a run is not inherited from the shell', () => {
  for (const name of ['RUTBA_VERIFY_ONLY', 'RUTBA_VERIFY_APPS', 'RUTBA_VERIFY_PART_IDS', 'RUTBA_LOCKER']) assert.ok(SHRINKING_ENV.includes(name), name);
});

test('a run that counts fewer checks than the README says the last did is seen', () => {
  const readme = 'what runs before a release: 3,784 engine tests, 12 editing checks,\n2,552 application checks and 16 window captures';
  const floors = floorsFrom(readme);
  assert.deepEqual(floors, { tests: 3784, edit: 12, apps: 2552, smoke: 16 });
  const full = { test: { tests: 3600, passed: 3600, failed: 0, skipped: 0 }, 'test:perf': { tests: 184, passed: 184, failed: 0, skipped: 0 }, 'verify:edit': { passed: 12, of: 12 }, 'verify:apps': { passed: 2552, of: 2552 }, smoke: { passed: 16, of: 16 } };
  assert.deepEqual(belowFloor(full, floors), []);
  assert.deepEqual(belowFloor({ ...full, 'verify:apps': { passed: 400, of: 400 } }, floors), ['application checks: 400, down from 2552']);
  assert.deepEqual(belowFloor({ ...full, test: { tests: 40, passed: 40, failed: 0, skipped: 0 } }, floors), ['engine tests: 224, down from 3784']);
  assert.deepEqual(belowFloor({ ...full, smoke: undefined }, floors), ['window captures: none, down from 16']);
});

test('only the note and the README may be changed when a run goes on the record', () => {
  assert.deepEqual(dirtyForRecord(' M README.md\n M docs/releases/v2.25.0.md\n', '2.25.0'), []);
  assert.deepEqual(dirtyForRecord(' M README.md\n M packages/formula/src/parser.js\n', '2.25.0'), [' M packages/formula/src/parser.js']);
  assert.deepEqual(dirtyForRecord('?? tests/new.test.js\n', '2.25.0'), ['?? tests/new.test.js']);
});

test('the rerun limit is small enough to mean a fault', () => {
  assert.ok(RERUN_LIMIT > 0 && RERUN_LIMIT <= 10);
});

test('the record names the commit, and a reran check still counts as having failed first', () => {
  const counts = { test: { tests: 10, passed: 10, failed: 0, skipped: 0 }, 'verify:apps': { passed: 9, of: 10 } };
  const section = gateSection(counts, { version: '9.9.9', date: '2026-11-10', platform: 'win32', arch: 'x64', reran: ['word: a check'], commit: SHA });
  assert.match(section, new RegExp(`\\nCommit checked: ${SHA}\\n`));
  assert.match(section.replace(/\s+/g, ' '), /9 of 10 application checks/);
  assert.equal(gatedCommitOf(`Words.\n\n${section}`), SHA);
  assert.equal(section.split('\n')[2], wrap('Run on Windows (x64) on 2026-11-10, before 9.9.9 was released: 10 engine tests, 10 passing; and 9 of 10 application checks.').split('\n')[0]);
});

test('a release goes up only if it is what the gate checked', () => {
  const note = `Words.\n\n## Gate\n\nRun on Windows.\n\nCommit checked: ${SHA}\n`;
  const tag = 'v9.9.9';
  assert.equal(gateProblem({ note, tag, changed: ['README.md', 'docs/releases/v9.9.9.md'] }), null);
  // The numbers the record gave checks seen for the first time are its own too, as the gate's record step has them.
  assert.equal(gateProblem({ note, tag, changed: ['README.md', 'tools/check-ids.json', 'tools/gate-floors.json'] }), null);
  assert.equal(gateProblem({ note, tag, changed: [] }), null);
  assert.match(gateProblem({ note, tag, changed: ['packages/formula/src/parser.js'] }), /changed since the gate checked/);
  assert.match(gateProblem({ note, tag, changed: null }), /not in this checkout/);
  assert.match(gateProblem({ note: 'Words only.', tag, changed: [] }), /no Gate section/);
  assert.match(gateProblem({ note: 'Words.\n\n## Gate\n\nRun on Windows.\n', tag, changed: [] }), /no Gate section/);
});
