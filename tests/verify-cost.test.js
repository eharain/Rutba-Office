// A speed check passes when the clock or Office's own processor time is
// under its limit (D-2026-11-13-25, the owner: "Either one under the limit
// (Recommended)"), and says both, and how the computer ran: a fixed
// reference task, how busy, how much memory free. That last is information
// only.
import test from 'node:test';
import assert from 'node:assert/strict';
import { cpuSince, referenceMs, busyNow, costWords, within } from '../apps/desktop/main/verify-cost.js';

const metricsOf = (rows) => () => rows.map(([pid, s]) => ({ pid, cpu: { cumulativeCPUUsage: s } }));
// The main process of these made-up runs.
const mainPid = 1;

test('what each process spent since the start is added up: a process there before by what it added, one started since by all it used', () => {
  let rows = [[1, 10], [2, 4.5]];
  const metrics = () => metricsOf(rows)();
  const spent = cpuSince({ mainPid, metrics });
  // The main process spent 1.5 s, the window 0.25 s, and a new window 2 s since it started.
  rows = [[1, 11.5], [2, 4.75], [3, 2]];
  const r = spent.end();
  assert.equal(r.cpuMs, 3750);
  assert.equal(r.ended, 0);
  assert.ok(r.ms >= 0);
});

test('a process that ended in between is counted as ended, and nothing it spent is made up', () => {
  let rows = [[1, 10], [2, 4]];
  const spent = cpuSince({ mainPid, metrics: () => metricsOf(rows)() });
  rows = [[1, 10.2]];
  const r = spent.end();
  assert.equal(r.cpuMs, 200);
  assert.equal(r.ended, 1);
});

test('a process with no processor time reported, or a counter that went back, adds nothing', () => {
  let rows = [[1, 10]];
  const spent = cpuSince({ mainPid, metrics: () => [...metricsOf(rows)(), { pid: 9, cpu: {} }] });
  rows = [[1, 9]];
  assert.equal(spent.end().cpuMs, 0);
});

test('the reference task and the processors\' load are measured for real, and the words say the reference is information only', async () => {
  const ms = referenceMs();
  assert.ok(Number.isInteger(ms) && ms >= 0);
  const busy = await busyNow(100);
  assert.ok(busy === null || (busy >= 0 && busy <= 100));
  const words = costWords({ cpuMs: 41000, ms: 52000, ended: 1, measured: true }, { referenceMs: 180, busy: 63, freeMb: 4100 });
  assert.match(words, /^52000 ms by the clock, 41000 ms of Office's own processor time \(1 of its processes ended in between\)/);
  assert.match(words, /the reference task took 180 ms, the processors 63% busy, 4100 MB free \(information only\)$/);
});

test('a count that cannot be trusted says so, and leaves the clock alone to judge: nothing reported, the main process missing, or the working window missing', () => {
  assert.equal(cpuSince({ mainPid, metrics: () => [] }).end().measured, false, 'nothing reported: 0 ms would have passed');
  let rows = [[1, 10], [2, 4]];
  const spent = cpuSince({ mainPid, metrics: () => metricsOf(rows)() });
  rows = [[1, 10.5], [2, 4.1]];
  assert.equal(spent.end().measured, true);
  assert.equal(spent.end({ need: [2] }).measured, true);
  assert.equal(spent.end({ need: [7] }).measured, false, 'the window doing the work never reported');
  rows = [[2, 4.2]];
  assert.equal(spent.end().measured, false, 'the main process stopped reporting');
  // A process that gives no time at all is not counted as reporting.
  const blank = cpuSince({ mainPid, metrics: () => [{ pid: 1, cpu: {} }] });
  assert.equal(blank.end().measured, false);
  assert.match(costWords({ cpuMs: 0, ms: 900, measured: false }), /^900 ms by the clock, Office's own processor time could not be measured .*so the clock alone was judged$/);
});

test('a check passes when either figure is under its limit, and an unmeasured processor time leaves the clock to judge (D-2026-11-13-25)', () => {
  // The clock under, the processor time over: work spread over several cores.
  assert.equal(within({ cpuMs: 75000, ms: 41000, measured: true }, 60000), true);
  // The processor time under, the clock over: other work on the computer.
  assert.equal(within({ cpuMs: 41000, ms: 75000, measured: true }, 60000), true);
  // At the limit is not under it, for either figure.
  assert.equal(within({ cpuMs: 61000, ms: 60000, measured: true }, 60000), false, 'the clock at the limit, the processor time over');
  assert.equal(within({ cpuMs: 60000, ms: 75000, measured: true }, 60000), false, 'the processor time at the limit, the clock over');
  // Not measured: a 0 never passes on its own, and the clock still can.
  assert.equal(within({ cpuMs: 0, ms: 75000, measured: false }, 60000), false);
  assert.equal(within({ cpuMs: 0, ms: 60000, measured: false }, 60000), false, 'the clock at the limit');
  assert.equal(within({ cpuMs: 0, ms: 41000, measured: false }, 60000), true);
  // A reading that does not say it was measured is taken as not measured.
  assert.equal(within({ cpuMs: 0, ms: 75000 }, 60000), false);
});
