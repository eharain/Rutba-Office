// Where the gate's engine tests and fuzz passes run (PM, 11 Nov 2026): on the
// lab machine when the gate has a runner for it, here with --local, never here
// in its place when it cannot be reached; the tests that touch Windows itself
// always here; and what each found read into the failure list.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { FUZZERS, VM_PASSES, VM_FAULTS, vmConfig, vmCommand, engineFiles, readHosts, splitEngine, osReach, moduleTouches, unlisted, vmFault, snapshotOf, parseTap, readRun, testArgs, engineFindings, mergeFindings, engineParts, isEngineId, hostsSentence, labLabel } from '../tools/gate-hosts.js';
import { FUZZERS as GATE_FUZZERS, gateSection } from '../tools/gate.js';
import { recordRun, emptyLedger, openParts } from '../tools/gate-ledger.js';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

test('the runner is used only when set, and --local brings the passes here for one run', () => {
  assert.deepEqual(vmConfig({}, []).on, false);
  const on = vmConfig({ RUTBA_GATE_VM: '/x/run.sh', RUTBA_GATE_VM_HOST: 'lab' }, []);
  assert.equal(on.on, true);
  assert.equal(on.host, 'lab');
  const here = vmConfig({ RUTBA_GATE_VM: '/x/run.sh', RUTBA_GATE_VM_HOST: 'lab' }, ['--local']);
  assert.equal(here.on, false);
  assert.match(here.said, /--local/);
  // No setting falls back to this computer by itself: there is none to set.
  assert.equal(vmConfig({ RUTBA_GATE_VM: '/x/run.sh', RUTBA_GATE_VM_FALLBACK: 'local' }, []).on, true);
  const [prog, args] = vmCommand(on, 'D:\\work\\office', 'vm-1234', ['node', 'tools/gate-vm.js'], {});
  assert.equal(prog, 'bash');
  assert.deepEqual(args.slice(0, 3), ['/x/run.sh', path.sep === '\\' ? 'D:/work/office' : 'D:\\work\\office', 'vm-1234']);
  assert.deepEqual(vmCommand(on, '/r', 'l', [], { RUTBA_GATE_VM_BASH: 'C:/Git/bin/bash.exe' })[0], 'C:/Git/bin/bash.exe');
});

test('the fuzz passes are one list, and the lab machine runs nothing with a window', () => {
  assert.equal(GATE_FUZZERS, FUZZERS);
  assert.deepEqual(VM_PASSES, ['test', 'fuzz:open', 'fuzz:pdf', 'fuzz:smime', 'fuzz:grammar']);
  for (const p of ['verify:apps', 'verify:edit', 'smoke', 'office:reopen', 'verify:packs', 'test:perf']) assert.equal(VM_PASSES.includes(p), false, p);
});

test("the runner's own exit codes are the machine's fault; a command's are not", () => {
  for (const code of [3, 4, 5, 6]) assert.equal(vmFault(code, '[vm] office at abc1234, snapshot 1234abcd\n'), VM_FAULTS[code]);
  // The command itself ended 4 after it counted: that is its own verdict.
  assert.equal(vmFault(4, 'TAP version 13\n# tests 3\n# pass 2\n# fail 1\n'), null);
  assert.equal(vmFault(1, '[gate-vm] test exit 1 in 30 s\n'), null);
  assert.equal(vmFault(0, ''), null);
  assert.equal(vmFault(1, '[vm] npm ci failed, see ~/npm-ci-office.log\n'), VM_FAULTS[6]);
  assert.equal(vmFault(255, '[vm] push to the VM failed\n'), VM_FAULTS[4]);
  assert.equal(snapshotOf('[vm] office at 78043841, snapshot 9f00aa12\nmore'), '9f00aa12');
  assert.equal(snapshotOf('nothing'), null);
  // The runner says the route it took to the machine (the LAN or the tailnet, 11 Nov).
  assert.equal(snapshotOf('[vm] office via 192.0.2.7 at bd2c2355, snapshot 9dabd0b2\r\nmore'), '9dabd0b2');
});

const FILES = ['tests/a.test.js', 'tests/b.test.js', 'packages/pdf/test/c.test.js'];
const TAP = `TAP version 13
# Subtest: top ok
ok 1 - top ok
  ---
  duration_ms: 2.4
  type: 'test'
  location: '/home/lab/work/office/tests/a.test.js:3:1'
  ...
# Subtest: suite
    # Subtest: inner fails
    not ok 1 - inner fails
      ---
      duration_ms: 1.9
      type: 'test'
      location: '/home/lab/work/office/tests/a.test.js:4:38'
      failureType: 'testCodeFailure'
      error: '1 == 2'
      code: 'ERR_ASSERTION'
      stack: |-
        TestContext.<anonymous> (file:///home/lab/work/office/tests/a.test.js:4:71)
      ...
    # Subtest: inner ok
    ok 2 - inner ok
      ---
      duration_ms: 0.3
      type: 'test'
      location: '/home/lab/work/office/tests/a.test.js:4:90'
      ...
    1..2
not ok 2 - suite
  ---
  duration_ms: 4.5
  type: 'test'
  location: '/home/lab/work/office/tests/a.test.js:4:1'
  failureType: 'subtestsFailed'
  error: '1 subtest failed'
  ...
# Subtest: skipped one
ok 3 - skipped one # SKIP
  ---
  duration_ms: 0.2
  type: 'test'
  location: '/home/lab/work/office/tests/a.test.js:5:6'
  ...
# Subtest: a \\# in a name
ok 4 - a \\# in a name
  ---
  duration_ms: 2.1
  type: 'test'
  location: 'D:\\\\work\\\\office\\\\packages\\\\pdf\\\\test\\\\c.test.js:9:1'
  ...
# Subtest: tests/b.test.js
not ok 5 - tests/b.test.js
  ---
  duration_ms: 242.9
  type: 'test'
  location: '/home/lab/work/office/tests/b.test.js:1:1'
  failureType: 'testCodeFailure'
  exitCode: 1
  error: 'test failed'
  ...
1..5
# tests 7
# suites 0
# pass 4
# fail 2
# cancelled 0
# skipped 1
# todo 0
`;

test('TAP is read into each test that failed itself, placed in its file and line', () => {
  const t = parseTap(TAP, FILES);
  assert.deepEqual(t.totals, { tests: 7, passed: 4, failed: 2, skipped: 1 });
  assert.deepEqual(t.failures.map((f) => [f.name, f.at, f.detail]), [
    ['tests/a.test.js > suite > inner fails', 'tests/a.test.js:4', '1 == 2'],
    ['tests/b.test.js: the file ran to its end', 'tests/b.test.js:1', 'test failed'],
  ]);
  // The parent that failed for its child is not a failure of its own; the skipped one is neither.
  assert.equal(t.failures.some((f) => f.name.endsWith('> suite')), false);
  assert.ok(t.passed.includes('tests/a.test.js > suite > inner ok'));
  assert.ok(t.passed.includes('packages/pdf/test/c.test.js > a # in a name'), t.passed.join('\n'));
  assert.equal(t.passed.some((n) => /skipped one/.test(n)), false);
  // A file that ran to its end passes by the name its failure would have, so a later pass closes that entry.
  assert.ok(t.passed.includes('tests/a.test.js: the file ran to its end'));
  assert.equal(t.passed.includes('tests/b.test.js: the file ran to its end'), false);
  assert.deepEqual(t.ran.sort(), ['packages/pdf/test/c.test.js', 'tests/a.test.js', 'tests/b.test.js']);
  // A run that never got to its totals says so.
  assert.equal(parseTap('TAP version 13\n# Subtest: x\n', FILES).totals, null);
});

// What tools/gate-reporter.js writes: a line a test, with its file, which TAP gives only for a failure.
const ev = (event, name, nesting, file, extra = {}) => JSON.stringify({ event, name, nesting, file, line: extra.line ?? 3, skip: Boolean(extra.skip), todo: false, failureType: extra.failureType || null, error: extra.error || null });
const EVENTS = [
  ev('start', 'top ok', 0, '/w/office/tests/a.test.js'),
  ev('pass', 'top ok', 0, '/w/office/tests/a.test.js'),
  ev('start', 'suite', 0, '/w/office/tests/a.test.js'),
  ev('start', 'inner fails', 1, '/w/office/tests/a.test.js'),
  ev('fail', 'inner fails', 1, '/w/office/tests/a.test.js', { line: 4, failureType: 'testCodeFailure', error: '1 == 2' }),
  ev('start', 'inner ok', 1, '/w/office/tests/a.test.js'),
  ev('pass', 'inner ok', 1, '/w/office/tests/a.test.js'),
  ev('fail', 'suite', 0, '/w/office/tests/a.test.js', { failureType: 'subtestsFailed' }),
  ev('start', 'skipped one', 0, '/w/office/tests/a.test.js'),
  ev('pass', 'skipped one', 0, '/w/office/tests/a.test.js', { skip: true }),
  ev('start', 'c works', 0, 'D:\\w\\office\\packages\\pdf\\test\\c.test.js'),
  ev('pass', 'c works', 0, 'D:\\w\\office\\packages\\pdf\\test\\c.test.js'),
  ev('start', 'tests/b.test.js', 0, '/w/office/tests/b.test.js'),
  ev('fail', 'tests/b.test.js', 0, '/w/office/tests/b.test.js', { line: 1, failureType: 'testCodeFailure', error: 'test failed' }),
  'not a line of ours',
].join('\n');

test("the reporter's lines place every test in its file, those that passed too", () => {
  const r = readRun(TAP, EVENTS, FILES);
  assert.deepEqual(r.totals, { tests: 7, passed: 4, failed: 2, skipped: 1 });
  assert.deepEqual(r.failures.map((f) => [f.name, f.at, f.detail]), [
    ['tests/a.test.js > suite > inner fails', 'tests/a.test.js:4', '1 == 2'],
    ['tests/b.test.js: the file ran to its end', 'tests/b.test.js:1', 'test failed'],
  ]);
  assert.ok(r.passed.includes('tests/a.test.js > suite > inner ok'));
  assert.ok(r.passed.includes('packages/pdf/test/c.test.js > c works'));
  assert.ok(r.passed.includes('tests/a.test.js: the file ran to its end'));
  assert.equal(r.passed.some((n) => /skipped one|b\.test\.js: the file/.test(n)), false);
  // With no lines, TAP alone: the same failures.
  assert.deepEqual(readRun(TAP, '', FILES).failures.map((f) => f.name), r.failures.map((f) => f.name));
  // Lines that saw no failure TAP saw are short: TAP's failures stand.
  assert.equal(readRun(TAP, EVENTS.split('\n').filter((l) => !l.includes('"fail"')).join('\n'), FILES).failures.length, 2);
  const args = testArgs(['tests/a.test.js'], { events: '.gate/logs/x.jsonl', alone: true });
  assert.ok(args.includes('--test-reporter=tap') && args.includes('--test-reporter=./tools/gate-reporter.js') && args.includes('--test-reporter-destination=.gate/logs/x.jsonl') && args.includes('--test-concurrency=1'));
});

test('engine findings go into the failure list by file, flaky when they passed alone, and close by name', () => {
  const first = parseTap(TAP, FILES);
  const again = parseTap(`ok 1 - inner fails\n  ---\n  location: '/w/tests/a.test.js:4:38'\n  ...\n# tests 1\n# pass 1\n# fail 0\n# skipped 0\n`, FILES);
  // A run again names a subtest without its parent; the gate reruns by file, so the name it sees again is its full one.
  const againFull = { ...again, passed: ['tests/a.test.js > suite > inner fails'] };
  const f = engineFindings({ tests: [{ host: 'lab', code: 1, tap: first, again: againFull, againCode: 0 }] });
  assert.deepEqual(f.failures.map((x) => [x.part, x.flaky, x.app]), [['test#tests/a.test.js', true, 'engine'], ['test#tests/b.test.js', false, 'engine']]);
  assert.ok(f.clean.has('engine|test#packages/pdf/test/c.test.js'));
  assert.equal(f.clean.has('engine|test#tests/a.test.js'), false);
  assert.equal(f.clean.has('engine|test#*'), false);
  const ledger = emptyLedger();
  recordRun(ledger, { head: 'a1', clean: true }, f);
  assert.deepEqual(openParts(ledger).sort(), ['test#tests/a.test.js', 'test#tests/b.test.js']);
  // The next run passes both by name: each gains its pass.
  const next = parseTap(`ok 1 - x\n  ---\n  location: '/w/tests/b.test.js:2:1'\n  ...\n# Subtest: suite\n    ok 1 - inner fails\n      ---\n      location: '/w/tests/a.test.js:4:38'\n      ...\nok 2 - suite\n  ---\n  location: '/w/tests/a.test.js:4:1'\n  ...\n# tests 3\n# pass 3\n# fail 0\n# skipped 0\n`, FILES);
  recordRun(ledger, { head: 'a2', clean: true }, engineFindings({ tests: [{ host: 'lab', code: 0, tap: next }], wholeSuite: false }));
  for (const e of Object.values(ledger.entries)) assert.equal(e.passedAt.length, 1, e.name);
});

test('a run that never counted, or ended badly with no test named, is an entry of its own', () => {
  const none = engineFindings({ tests: [{ host: 'lab', code: 1, tap: { failures: [], passed: [], ran: [], totals: null } }] });
  assert.equal(none.failures.length, 1);
  assert.equal(none.failures[0].part, 'test#*');
  assert.equal(none.failures[0].synthetic, true);
  const odd = engineFindings({ tests: [{ host: 'here', code: 1, tap: { failures: [], passed: ['x'], ran: ['tests/a.test.js'], totals: { tests: 1, passed: 1, failed: 0, skipped: 0 } } }] });
  assert.match(odd.failures[0].name, /ended with exit 1, no test named/);
  const whole = engineFindings({ tests: [{ host: 'lab', code: 0, tap: { failures: [], passed: [], ran: [], totals: { tests: 1, passed: 1, failed: 0, skipped: 0 } } }], wholeSuite: true });
  assert.ok(whole.clean.has('engine|test#*'));
});

test('a fuzz pass that failed is an entry by its pass, flaky only when it passed alone', () => {
  const f = engineFindings({ fuzz: [{ name: 'fuzz:pdf', host: 'lab', ok: false, counted: { passed: 298, of: 300 }, again: { ok: true } }, { name: 'fuzz:open', host: 'lab', ok: true }] });
  assert.equal(f.failures.length, 1);
  assert.deepEqual([f.failures[0].part, f.failures[0].flaky, f.failures[0].synthetic], ['fuzz:pdf#*', true, true]);
  assert.match(f.failures[0].detail, /2 of 300/);
  assert.ok(f.clean.has('engine|fuzz:open#*'));
  const m = mergeFindings(f, { failures: [{ key: 'k' }], passed: new Set(['p']), clean: new Set(['c']) });
  assert.equal(m.failures.length, 2);
  assert.ok(m.clean.has('c') && m.clean.has('engine|fuzz:open#*'));
});

test('engine parts by id: a test file, the suite, a fuzz pass; anything else is named back', () => {
  assert.equal(isEngineId('test#tests/a.test.js'), true);
  assert.equal(isEngineId('fuzz:pdf#*'), true);
  assert.equal(isEngineId('mail#reading'), false);
  assert.deepEqual(engineParts(['test#tests/a.test.js', 'fuzz:pdf#*', 'test#tests/none.test.js', 'fuzz:nope#*', 'fuzz:pdf#seed'], FILES), { tests: ['tests/a.test.js'], fuzz: ['fuzz:pdf'], strange: ['test#tests/none.test.js', 'fuzz:nope#*', 'fuzz:pdf#seed'] });
  assert.deepEqual(engineParts(['test#*', 'test#tests/a.test.js'], FILES).tests, ['*']);
});

test("the record says what machine each ran on, by its system and cores, and never the machine's name, user or address", () => {
  // The host as RUTBA_GATE_VM_HOST might give it: none of it may reach the public record (PM, 11 Nov 2026).
  const HOST = 'secret-lab-77.example.internal';
  const s = hostsSentence([
    { pass: 'test', host: HOST, osName: 'Rocky Linux 9.5 (Blue Onyx)', cpus: 12, tests: 2100 },
    { pass: 'test', host: 'local', osName: 'Windows', tests: 300 },
    { pass: 'fuzz:open', host: HOST, osName: 'Rocky Linux 9.5 (Blue Onyx)', cpus: 12 },
    { pass: 'fuzz:pdf', host: HOST, osName: 'Rocky Linux 9.5 (Blue Onyx)', cpus: 12 },
  ]);
  assert.match(s, /2,100 engine tests on the Linux lab machine \(Rocky Linux 9\.5, 12 cores\)/);
  const section0 = gateSection({ test: { tests: 10, passed: 10, failed: 0, skipped: 0 } }, { version: '9.9.9', date: '2026-11-11', platform: 'win32', arch: 'x64', hosts: s });
  // The real host's name is checked too when it is set here, and so never written in this file.
  for (const bit of [HOST, 'secret', 'lab-77', 'example.internal', '192.168', '100.', ...(process.env.RUTBA_GATE_VM_HOST ? [process.env.RUTBA_GATE_VM_HOST] : [])]) {
    assert.ok(!s.includes(bit) && !section0.includes(bit), `the record says ${bit}`);
  }
  assert.match(s, /300 engine tests on this computer, Windows, those that touch Windows itself/);
  assert.match(s, /fuzz:open, fuzz:pdf on the Linux lab machine \(Rocky Linux 9\.5, 12 cores\)/);
  assert.equal(hostsSentence([{ pass: 'test', host: 'local', tests: 5 }]), null);
  const section = gateSection({ test: { tests: 10, passed: 10, failed: 0, skipped: 0 } }, { version: '9.9.9', date: '2026-11-11', platform: 'win32', arch: 'x64', hosts: s });
  assert.match(section, /Where they ran:/);
  assert.equal(/[\u2014]/.test(section), false);
});

test('every engine test file that touches the operating system is on one list or the other', () => {
  const files = engineFiles(ROOT);
  assert.ok(files.length > 400, `${files.length} files`);
  const hosts = readHosts(ROOT);
  assert.deepEqual(unlisted(ROOT, files, hosts), [], 'put each on windowsOnly, or on portable with why it runs the same on Linux (tools/engine-hosts.json)');
  for (const [list, entries] of Object.entries({ windowsOnly: hosts.windowsOnly, portable: hosts.portable })) {
    for (const [f, why] of Object.entries(entries)) {
      assert.ok(files.includes(f), `${list} names ${f}, which npm test does not run`);
      assert.ok(String(why).length >= 10, `${list}: ${f} needs a reason`);
    }
  }
  const both = Object.keys(hosts.windowsOnly).filter((f) => f in hosts.portable);
  assert.deepEqual(both, []);
  const s = splitEngine(files, hosts);
  assert.equal(s.vm.length + s.local.length, files.length);
  assert.ok(s.local.includes('tests/gate-again.test.js'));
});

test('the lab side refuses a pass it does not run, and a file that is not its own', () => {
  const run = (args) => spawnSync(process.execPath, [path.join(ROOT, 'tools', 'gate-vm.js'), ...args], { cwd: ROOT, encoding: 'utf8' });
  const a = run(['--passes', 'smoke']);
  assert.equal(a.status, 2);
  assert.match(a.stdout, /smoke/);
  const b = run(['--passes', 'test', '--files', 'tests/gate-again.test.js']);
  assert.equal(b.status, 2);
  assert.match(b.stdout, /not test files this machine runs/);
  assert.equal(fs.existsSync(path.join(ROOT, 'tools', 'gate-vm.js')), true);
});

test('a test file whose imports reach a module that touches Windows stays here, however far down, cycles and all (review of 11 Nov)', () => {
  const dir = fs.mkdtempSync(path.join(ROOT, '.gate', 'reach-'));
  try {
    const w = (f, t) => { fs.mkdirSync(path.dirname(path.join(dir, f)), { recursive: true }); fs.writeFileSync(path.join(dir, f), t); };
    w('lib/a.js', "import { b } from './b.js';\nexport const a = 1;\n");
    w('lib/b.js', "import { a } from './a.js';\nexport { c } from './deep/c.js';\nexport const b = 2;\n");
    w('lib/deep/c.js', "export const c = process.platform === 'win32' ? 'x' : 'y';\n");
    w('lib/plain.js', "const m = await import('./a.js');\nexport const p = 3;\n");
    w('lib/alone.js', 'export const z = 4;\n');
    w('tests/one.test.js', "import { a } from '../lib/a.js';\n");
    w('tests/two.test.js', "import { p } from '../lib/plain.js';\n");
    w('tests/three.test.js', "import { z } from '../lib/alone.js';\nimport os from 'node:os';\n");
    const cache = new Map();
    assert.equal(osReach(dir, 'tests/one.test.js', cache), 'lib/deep/c.js');
    assert.equal(osReach(dir, 'tests/two.test.js', cache), 'lib/deep/c.js', 'through a dynamic import');
    assert.equal(osReach(dir, 'tests/three.test.js', cache), null);
    const split = splitEngine(['tests/one.test.js', 'tests/two.test.js', 'tests/three.test.js', 'tests/listed.test.js'], { windowsOnly: { 'tests/listed.test.js': 'x' } }, (f) => osReach(dir, f, cache));
    assert.deepEqual(split, { vm: ['tests/three.test.js'], local: ['tests/one.test.js', 'tests/two.test.js', 'tests/listed.test.js'] });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test("the suite's own split keeps on the lab machine only files that reach nothing of Windows", () => {
  const files = engineFiles(ROOT);
  const cache = new Map();
  const split = splitEngine(files, readHosts(ROOT), (f) => osReach(ROOT, f, cache));
  assert.ok(split.vm.length > files.length / 2, `${split.vm.length} of ${files.length} on the lab machine`);
  for (const f of ['tests/mail.test.js', 'tests/documents.test.js']) if (files.includes(f)) assert.ok(split.local.includes(f), f);
});

test('a file that crashed on Windows, titled with backslashes, is its own failure and is not counted as run to its end', () => {
  const tap = ['TAP version 13', '# Subtest: tests\\crash.test.js', 'not ok 1 - tests\\crash.test.js', '  ---', '  duration_ms: 5', "  failureType: 'testCodeFailure'", "  error: 'boom'", '  ...', '1..1', '# tests 1', '# pass 0', '# fail 1'].join('\n');
  const r = parseTap(tap, ['tests/crash.test.js']);
  assert.deepEqual(r.failures.map((f) => f.name), ['tests/crash.test.js: the file ran to its end']);
  assert.ok(!r.passed.includes('tests/crash.test.js: the file ran to its end'));
});

test('an engine failure is proved passed only on the host it failed on (review of 11 Nov)', () => {
  const ledger = emptyLedger();
  const fail = { failures: [{ name: 'tests/x.test.js', file: 'tests/x.test.js', at: 'tests/x.test.js:3', detail: 'x' }], passed: [], ran: ['tests/x.test.js'], totals: { tests: 1, passed: 0, failed: 1, skipped: 0 } };
  const pass = { failures: [], passed: ['tests/x.test.js'], ran: ['tests/x.test.js'], totals: { tests: 1, passed: 1, failed: 0, skipped: 0 } };
  const failure = { ...fail, failures: [{ ...fail.failures[0], name: 'tests/x.test.js > t' }] };
  const passing = { ...pass, passed: ['tests/x.test.js > t'] };
  recordRun(ledger, { head: 'a1' }, engineFindings({ tests: [{ host: 'lab-box', code: 1, tap: failure }] }));
  const [entry] = Object.values(ledger.entries);
  assert.equal(entry.last.lane, 'lab-box');
  recordRun(ledger, { head: 'a2' }, engineFindings({ tests: [{ host: 'local', code: 0, tap: passing }] }));
  assert.equal(entry.passedAt.length, 0, 'a pass on Windows is no proof of a fix on Linux');
  recordRun(ledger, { head: 'a3' }, engineFindings({ tests: [{ host: 'lab-box', code: 0, tap: passing }] }));
  assert.equal(entry.passedAt.length, 1);
});

test('an engine entry whose file moved hosts to be fixed is proved by a pass where it runs now, and --local proves nothing of a lab file', () => {
  const ledger = emptyLedger();
  const failure = { failures: [{ name: 'tests/x.test.js > t', file: 'tests/x.test.js', at: 'tests/x.test.js:3', detail: 'x' }], passed: [], ran: ['tests/x.test.js'], totals: { tests: 1, passed: 0, failed: 1, skipped: 0 } };
  const passing = { failures: [], passed: ['tests/x.test.js > t'], ran: ['tests/x.test.js'], totals: { tests: 1, passed: 1, failed: 0, skipped: 0 } };
  recordRun(ledger, { head: 'a1' }, engineFindings({ tests: [{ host: 'lab-box', code: 1, tap: failure }], home: () => 'lab-box' }));
  const [entry] = Object.values(ledger.entries);
  // --local: the file still belongs to the lab machine, so a pass here is no proof.
  recordRun(ledger, { head: 'a2' }, engineFindings({ tests: [{ host: 'local', code: 0, tap: passing }], home: () => 'lab-box' }));
  assert.equal(entry.passedAt.length, 0);
  // Listed as Windows-only since: it belongs here now, and its pass here proves the fix.
  recordRun(ledger, { head: 'a3' }, engineFindings({ tests: [{ host: 'local', code: 0, tap: passing }], home: () => 'local' }));
  assert.equal(entry.passedAt.length, 1);
});

test('the import walk follows the workspace packages by name, through their exports', () => {
  const dir = fs.mkdtempSync(path.join(ROOT, '.gate', 'reach-'));
  try {
    const w = (f, t) => { fs.mkdirSync(path.dirname(path.join(dir, f)), { recursive: true }); fs.writeFileSync(path.join(dir, f), t); };
    w('packages/shell/package.json', JSON.stringify({ name: '@rutba/shell', exports: { '.': './src/index.js', './paths': './src/paths.js', './lib/*': './src/lib/*.js' } }));
    w('packages/shell/src/lib/fonts.js', "export const f = 'systemFonts';\n");
    w('tests/d.test.js', "import { f } from '@rutba/shell/lib/fonts';\n");
    w('packages/shell/src/index.js', "export const i = 1;\n");
    w('packages/shell/src/paths.js', "export const sep = process.platform === 'win32' ? '\\\\' : '/';\n");
    w('tests/a.test.js', "import { i } from '@rutba/shell';\n");
    w('tests/b.test.js', "import { sep } from '@rutba/shell/paths';\n");
    w('tests/c.test.js', "import x from '@rutba/missing';\n");
    const cache = new Map();
    assert.equal(osReach(dir, 'tests/a.test.js', cache), null);
    assert.equal(osReach(dir, 'tests/b.test.js', cache), 'packages/shell/src/paths.js');
    assert.equal(osReach(dir, 'tests/c.test.js', cache), null);
    assert.equal(osReach(dir, 'tests/d.test.js', cache), 'packages/shell/src/lib/fonts.js', 'through a pattern export');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test("a module touches the system by its code, not by the word in a comment or a message", () => {
  assert.equal(moduleTouches("/** The in-sheet clipboard mark. */\n// a clipboard we no longer own\nconst m = { 'There is nothing on the clipboard to paste.': 'x' };\n"), false);
  assert.equal(moduleTouches("const text = clipboard.readText();\n"), true);
  assert.equal(moduleTouches("// on win32 only\nexport const sep = '/';\n"), false);
  assert.equal(moduleTouches("export const sep = process.platform === 'x' ? 1 : 2;\n"), true);
});

test("every workspace package's exports are plain paths the import walk can follow: one star at most and no conditions", () => {
  for (const d of fs.readdirSync(path.join(ROOT, 'packages'))) {
    let pkg;
    try { pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages', d, 'package.json'), 'utf8')); } catch { continue; }
    if (!pkg.exports || typeof pkg.exports === 'string') continue;
    for (const [key, target] of Object.entries(pkg.exports)) {
      assert.ok(key.split('*').length <= 2, `${pkg.name} exports ${key}: osReach in tools/gate-hosts.js follows one star at most`);
      assert.equal(typeof target, 'string', `${pkg.name} exports ${key} by conditions: osReach in tools/gate-hosts.js reads a plain path`);
    }
  }
});
