// A speed check judges Office's own processor time (D-2026-11-13-20, the
// owner: "Yes, own processor time (Recommended)"), and says beside it how the
// computer ran: a fixed reference task, how busy, how much memory free. That
// last is information only.
import test from 'node:test';
import assert from 'node:assert/strict';
import { cpuSince, referenceMs, busyNow, costWords } from '../apps/desktop/main/verify-cost.js';

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
  const words = costWords({ cpuMs: 41000, ms: 52000, ended: 1 }, { referenceMs: 180, busy: 63, freeMb: 4100 });
  assert.match(words, /^41000 ms of Office's own processor time \(52000 ms by the clock, 1 of its processes ended in between\)/);
  assert.match(words, /the reference task took 180 ms, the processors 63% busy, 4100 MB free \(information only\)$/);
});

test('a count that cannot be trusted says so, and a check judged on it does not pass: nothing reported, the main process missing, or the working window missing', () => {
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
  assert.match(costWords({ cpuMs: 0, ms: 900, measured: false }), /^Office's own processor time could not be measured/);
});
