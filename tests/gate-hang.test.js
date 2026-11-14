// A window lane that gives no check result for CHECK_MINUTES is ended as
// hung, and the failure list says where it stopped (PM, 11 Nov 2026: a lane
// hung over an hour with only clock lines, holding the lock for everyone).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { run, CHECK_LINE, CHECK_MINUTES, treeOf, appOfCheck } from '../tools/gate.js';
import { runFindings, hangDetail, recordRun, emptyLedger, ledgerLines, openEntries } from '../tools/gate-ledger.js';
import { mergePieces } from '../tools/gate-plan.js';

test("a check's result line gives the check's name; a clock line is no result", () => {
  assert.equal(CHECK_LINE.exec('ok   [07.0221] word: a click ticks it : "- [x] Order milk"')[1], 'word: a click ticks it');
  assert.equal(CHECK_LINE.exec('FAIL [new] sheets: Find All lists Tea')[1], 'sheets: Find All lists Tea');
  assert.equal(CHECK_LINE.exec('     [clock] 302 s in, after "presentations: x"'), null);
  assert.ok(CHECK_MINUTES >= 2);
});

test('a lane that goes quiet is ended, and the last check it gave and those that failed are kept', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'gate-hang-'));
  try {
    const script = path.join(dir, 'lane.js');
    fs.writeFileSync(script, [
      "console.log('ok   [01.0001] shell: the window opens : yes');",
      "console.log('FAIL [01.0002] shell: the tab bar draws : no');",
      "console.log('     [clock] 5 s in');",
      'setTimeout(() => {}, 150000);',
      '',
    ].join('\n'));
    // Eight seconds of quiet, not one and a half: the clock starts at the spawn, and beside the rest of the
    // suite a node that took longer than that to print its first line read as hung before it began.
    // The lane holds on for two and a half minutes, so ending well inside that is still the proof, with the
    // exit code: a lane that ran out its time ends 0. Ending it lists the processes first, and on a busy
    // computer that listing can pass its 30 s before the gate ends the lane's own tree instead (unit run of
    // 14 Nov 2026 05:12: ended at 47 s, past the 45 s this allowed when the lane held for one minute).
    // 120 s: seen as hung at the second 8 s tick (16 s), a listing that runs out (30 s), and once more if
    // that first end failed (8 s and 30 s), with room to spare.
    const started = Date.now();
    const out = await run(process.execPath, [script], {}, { log: path.join(dir, 'lane.log'), stallMs: 8000, shell: false });
    assert.ok(Date.now() - started < 120_000, 'ended, not waited out');
    assert.deepEqual(out.hung, { minutes: 0, after: 'shell: the tab bar draws', failed: ['shell: the tab bar draws'] });
    assert.notEqual(out.code, 0);
    // A lane that gives its results in time is not touched.
    fs.writeFileSync(script, "console.log('ok   [01.0001] shell: the window opens : yes');\n");
    const fine = await run(process.execPath, [script], {}, { log: path.join(dir, 'lane.log'), stallMs: 8000, shell: false });
    assert.equal(fine.hung, null);
    assert.equal(fine.code, 0);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('a hung app and a hung machine lane are entries that say where they stopped, and keep the names of what failed before', () => {
  const hung = { minutes: 15, after: 'mail drafts: a draft sent leaves Drafts', failed: ['mail: a rule runs'] };
  // As windowRun reads a hung lane: no parts, the crash, and what it said failed before the hang.
  const piece = { ok: false, seconds: 900, seen: [], parts: null, sum: { passed: 40, of: 41, failed: ['the run ended before it counted', 'mail: a rule runs'] }, hung };
  const mail = mergePieces([piece], 'mail', () => []);
  assert.deepEqual(mail.hung, hung);
  assert.deepEqual(mail.failedParts, ['*']);
  const { failures } = runFindings({ apps: { mail }, machineCrashed: true, machineHung: { minutes: 15, after: null, failed: ['clipboard: a copy pastes'] }, machineScope: ['shell'] });
  const app = failures.find((f) => f.app === 'mail' && f.crash);
  assert.equal(app.hang, true);
  assert.equal(app.cause, undefined, 'a hang is no wait that gave up');
  assert.equal(app.detail, hangDetail(hung));
  assert.match(app.detail, /15 minutes after "mail drafts: a draft sent leaves Drafts"/);
  assert.ok(failures.some((f) => f.app === 'mail' && f.name === 'mail: a rule runs' && !f.crash), 'the check that failed before the hang keeps its name');
  const lane = failures.find((f) => f.app === 'machine' && f.crash);
  assert.match(lane.detail, /none at all/);
  assert.ok(failures.some((f) => f.app === 'machine' && f.name === 'clipboard: a copy pastes'));
  // Printed in the list: the detail, and no wait line made of nothing.
  const ledger = recordRun(emptyLedger(), { head: 'a1' }, { failures });
  const text = ledgerLines(openEntries(ledger)).join('\n');
  assert.match(text, /hung: no check result for 15 minutes/);
  assert.doesNotMatch(text, /undefined/);
});

test('what a run left behind is found by its parent, even once the parent has gone, but never through an id Windows has given to another', () => {
  const procs = [
    // The runner's bash (100, gone) left its ssh (101); ssh's own child (102).
    { pid: 101, ppid: 100, created: 1_010_000 },
    { pid: 102, ppid: 101, created: 1_020_000 },
    // Id 200 given again to another program later; its child is that program's, the older one ours.
    { pid: 200, ppid: 1, created: 5_000_000 },
    { pid: 201, ppid: 200, created: 2_000_000 },
    { pid: 202, ppid: 200, created: 5_100_000 },
    // Made before the run started: not the run's.
    { pid: 301, ppid: 300, created: 10 },
  ];
  // Our bash is gone: it is not listed, what it left is.
  assert.deepEqual(treeOf(100, 1_000_000, procs), [101, 102]);
  // Id 200 is another's now: neither it nor its child is ended, our older child is.
  assert.deepEqual(treeOf(200, 1_900_000, procs), [201]);
  assert.deepEqual(treeOf(300, 1_000_000, procs), []);
  // A live root is listed.
  assert.deepEqual(treeOf(101, 1_010_500, procs), [101, 102]);
  // The id held a while by another after ours ended, which left a child: that child is not ours.
  const reused = [
    { pid: 600, ppid: 500, created: 1_030_000 },
    { pid: 800, ppid: 500, created: 1_600_000 },
  ];
  assert.deepEqual(treeOf(500, 1_000_000, reused, 1_100_000), [600]);
});

test("a check that failed before the machine lane hung is its app's entry when its number says the app, else the lane's, run again by the lane", () => {
  const registry = { apps: { shell: 1, mail: 18 }, ids: { 'mail drafts: a draft sent leaves Drafts': '18.0185', 'ribbon: x': '02.0001' } };
  assert.equal(appOfCheck(registry, 'mail drafts: a draft sent leaves Drafts'), 'mail');
  assert.equal(appOfCheck(registry, 'a check with no number'), null);
  const { failures } = runFindings({ machineCrashed: true, machineHung: { minutes: 15, after: 'x', failed: [{ name: 'mail drafts: a draft sent leaves Drafts', app: 'mail' }, { name: 'new one', app: null }] }, machineScope: ['mail', 'shell'] });
  const mine = failures.find((f) => f.name === 'mail drafts: a draft sent leaves Drafts');
  assert.deepEqual([mine.app, mine.part, mine.lane, mine.synthetic], ['mail', 'mail#*', 'machine', false]);
  const lane = failures.find((f) => f.name === 'new one');
  assert.deepEqual([lane.app, lane.part, lane.synthetic, lane.rerun], ['machine', 'machine#lane', true, ['mail#*', 'shell#*']]);
});
