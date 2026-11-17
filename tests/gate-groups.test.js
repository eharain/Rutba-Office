// The gate run unit by unit (tools/gate-groups.js): the units leave nothing
// out and run nothing twice, and each is a run of the gate itself.
import fs from 'node:fs';
import test from 'node:test';
import assert from 'node:assert/strict';
import { UNITS, unitArgs, chosen, readLog, summaryLine } from '../tools/gate-groups.js';
import { VERIFY_APPS, FUZZERS } from '../tools/gate.js';

const gateSource = fs.readFileSync(new URL('../tools/gate.js', import.meta.url), 'utf8');

test('every app\'s window checks are in exactly one unit', () => {
  const apps = UNITS.flatMap((u) => u.apps || []);
  assert.deepEqual(apps.slice().sort(), VERIFY_APPS.slice().sort());
  assert.equal(new Set(apps).size, apps.length, 'no app in two units');
});

test('every pass the gate knows is in exactly one unit', () => {
  // The names the gate asks for by `want('<pass>')`, and the fuzzers it runs from its list.
  const known = new Set([...gateSource.matchAll(/\bwant\('([a-z:]+)'\)/g)].map((m) => m[1]));
  for (const [name] of FUZZERS) known.add(name);
  known.delete('verify:apps');
  const passes = UNITS.flatMap((u) => u.passes || []);
  assert.deepEqual(passes.slice().sort(), [...known].sort());
  assert.equal(new Set(passes).size, passes.length, 'no pass in two units');
});

test('a unit is the gate\'s passes or the window checks of its apps, never both and never neither, under a name of its own', () => {
  for (const u of UNITS) {
    assert.ok(Boolean(u.apps?.length) !== Boolean(u.passes?.length), u.key);
    assert.ok(u.what.length > 3, u.key);
  }
  assert.equal(new Set(UNITS.map((u) => u.key)).size, UNITS.length);
});

test('a unit runs the gate itself, resumed on this commit, and never asks for the record', () => {
  assert.deepEqual(unitArgs({ key: 'mail', apps: ['mail'] }), ['--resume', '--jobs', '2', 'verify:apps', '--apps', 'mail']);
  assert.deepEqual(unitArgs({ key: 'x', apps: ['calendar', 'contacts'] }, { jobs: 3 }), ['--resume', '--jobs', '3', 'verify:apps', '--apps', 'calendar,contacts']);
  assert.deepEqual(unitArgs({ key: 'y', passes: ['smoke', 'verify:packs'] }), ['--resume', '--jobs', '2', 'smoke', 'verify:packs']);
  for (const u of UNITS) {
    const args = unitArgs(u);
    assert.ok(args.includes('--resume'), u.key);
    for (const never of ['--record', '--parts', '--failed', '--fewer-checks', '--no-settle']) assert.ok(!args.includes(never), `${u.key}: ${never}`);
  }
  // The record is asked of the gate once, with nothing left out, and only by the runner's --record.
  const runner = fs.readFileSync(new URL('../tools/gate-groups.js', import.meta.url), 'utf8');
  assert.match(runner, /gate\(\['--resume', '--record', '--jobs', String\(jobs\)\], logFile\)/);
  assert.equal((runner.match(/'--record'/g) || []).length, 2, 'named where the flag is read, and passed on to the gate once');
  assert.match(runner, /if \(record && units\.length < UNITS\.length\)/, 'and refused for a run of some units');
});

test('some units are chosen by name or from one on, in the units\' own order; a name that is no unit\'s is refused', () => {
  const keys = (list) => list.map((u) => u.key);
  assert.deepEqual(keys(chosen(UNITS)), keys(UNITS));
  assert.deepEqual(keys(chosen(UNITS, { only: ['mail', 'word'] })), ['word', 'mail']);
  assert.deepEqual(keys(chosen(UNITS, { from: 'damaged-pdf' })), ['damaged-pdf', 'harm']);
  assert.throws(() => chosen(UNITS, { only: ['words'] }), /No such unit: words/);
  assert.throws(() => chosen(UNITS, { from: 'nowhere' }), /No such unit: nowhere/);
});

test('a unit\'s line says what the gate said: its verdict, its counts, and the checks that failed', () => {
  const log = [
    '  [fuzz:open] ok   in 73.4 s, 9600/9600',
    '  [corpus] FAIL in 19.1 s, 70/71 (.gate\\logs\\corpus.log)',
    '         FAIL a file from outside opened',
    'the failure list: 1 failed in this run, 1 open in all',
    '\u001b[31mThe gate is closed.\u001b[0m Fix what failed and run it again',
  ].join('\r\n');
  const r = readLog(log);
  assert.equal(r.verdict, 'The gate is closed');
  assert.deepEqual(r.counted, ['[fuzz:open] ok   in 73.4 s, 9600/9600', '[corpus] FAIL in 19.1 s, 70/71']);
  assert.equal(r.fails, 1);
  assert.equal(r.listed, 'the failure list: 1 failed in this run, 1 open in all');
  const line = summaryLine({ key: 'outside' }, { code: 1, seconds: 96.2, log });
  assert.match(line, /^outside +exit 1 +96 s \| The gate is closed \| .*70\/71 \| FAIL lines 1 \| the failure list/);
  assert.equal(readLog('').verdict, 'no verdict');
});

test('the runner does not start while the screen is locked, before any unit', () => {
  const runner = fs.readFileSync(new URL('../tools/gate-groups.js', import.meta.url), 'utf8');
  const at = runner.indexOf('if (screenLocked())');
  assert.ok(at > 0 && /return 2;/.test(runner.slice(at, at + 200)), 'it stops with a code of its own');
  assert.ok(at < runner.indexOf('for (const unit of units)'), 'before the first unit');
  assert.ok(at > runner.indexOf("argv.includes('--list')"), 'the list is still shown');
});

test('npm run gate:groups is the runner', () => {
  const scripts = JSON.parse(fs.readFileSync(new URL('../package.json', import.meta.url), 'utf8')).scripts;
  assert.equal(scripts['gate:groups'], 'node tools/gate-groups.js');
});
