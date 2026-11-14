// A wait that gives up says how busy the machine was over it and how much
// memory was free, beside its cause, in the log and the failure list: so a
// wait that ran out on a busy machine reads as one (a gate that co-exists,
// 13 Nov 2026: the load beside every timed result). The processors are read
// in a worker, never on the main thread (review of 14 Nov).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { until, gaveUp } from '../apps/desktop/main/verify-waits.js';
import { busyBetween, loadFrom, loadMark, loadSince, saidLoad } from '../apps/desktop/main/verify-load.js';
import { ledgerLines } from '../tools/gate-ledger.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const times = (idle, user) => ({ user, nice: 0, sys: 0, idle, irq: 0 });
const GB = 2 ** 30;

test('the busy share is what the processors did between two readings, all of them together', () => {
  const from = [times(1000, 1000), times(1000, 1000)];
  // One processor all busy, one half: 100 of 100 and 50 of 100, so 75%.
  assert.equal(busyBetween(from, [times(1000, 1100), times(1050, 1050)]), 0.75);
  assert.equal(busyBetween(from, [times(1100, 1000), times(1100, 1000)]), 0);
  assert.equal(busyBetween(from, from), null, 'no time passed: nothing measured, not 0% busy');
  assert.equal(busyBetween([], []), null, 'no processors listed: nothing measured');
});

test('the load is read from the reading before the wait began to the latest, and unsaid when there is nothing to compare', () => {
  const kept = [
    { at: 1000, times: [times(0, 0)] },
    { at: 3000, times: [times(100, 100)] },
    { at: 5000, times: [times(100, 300)] },
    { at: 7000, times: [times(150, 350)] },
  ];
  const free = { freemem: () => 3.25 * GB };
  // Begun at 4000: from the reading at 3000 to the one at 7000, 50 idle of 300.
  assert.deepEqual(loadFrom(4000, kept, free), { busy: 83, freeGb: 3.3 });
  // Begun before the first reading kept: from the oldest, 150 idle of 500.
  assert.deepEqual(loadFrom(500, kept, free), { busy: 70, freeGb: 3.3 });
  // Begun after the latest: nothing between, so nothing said of the processors.
  assert.deepEqual(loadFrom(8000, kept, free), { busy: null, freeGb: 3.3 });
  assert.deepEqual(loadFrom(0, kept.slice(0, 1), free), { busy: null, freeGb: 3.3 });
  assert.equal(saidLoad({ busy: 83, freeGb: 3.3 }), 'the processors 83% busy over the wait, 3.3 GB free');
  assert.equal(saidLoad({ busy: null, freeGb: 3.3 }), 'the processors not read over the wait, 3.3 GB free');
  assert.equal(saidLoad(null), '');
});

test('the processors are read in a worker, and the main thread never calls os.cpus() for a wait', async () => {
  const src = (f) => fs.readFileSync(path.join(here, '..', 'apps', 'desktop', 'main', f), 'utf8');
  // Only the sampler's own source, run in the worker, reads them.
  assert.doesNotMatch(src('verify-waits.js'), /os\.cpus\(/);
  // Its code without its comments: the one call is in the sampler's source.
  const load = src('verify-load.js').split('\n').filter((l) => !/^\s*\/\//.test(l)).join('\n');
  assert.equal((load.match(/os\.cpus\(\)/g) || []).length, 1);
  assert.match(load, /const SAMPLER = `[^`]*os\.cpus\(\)[^`]*`/);
  // And it does read them: two readings come within a few seconds of the first mark.
  const mark = loadMark();
  const end = Date.now() + 20_000;
  let load2 = loadSince(mark);
  while (load2.busy === null && Date.now() < end) {
    await new Promise((r) => setTimeout(r, 250));
    load2 = loadSince(mark - 60_000);
  }
  assert.ok(Number.isInteger(load2.busy) && load2.busy >= 0 && load2.busy <= 100, JSON.stringify(load2));
});

test('the sampler never keeps a run alive: a process that marked a wait ends on its own', async () => {
  const { spawnSync } = await import('node:child_process');
  const url = new URL('../apps/desktop/main/verify-load.js', import.meta.url).href;
  const out = spawnSync(process.execPath, ['--input-type=module', '-e', `import { loadMark } from ${JSON.stringify(url)}; loadMark(); setTimeout(() => {}, 2500);`], { timeout: 60_000, encoding: 'utf8' });
  // Ended by itself once its own 2.5 s timer was done, not by the 60 s limit.
  assert.equal(out.error, undefined, String(out.error));
  assert.equal(out.status, 0, out.stderr);
});

test('a wait that gives up keeps the load over it with its cause, and says it in the log', async () => {
  const said = [];
  gaveUp.last = null;
  await assert.rejects(until(() => false, 'the dialog', 200, { exact: true, log: (l) => said.push(l), tick: 20 }), /the dialog/);
  const { load, ms } = gaveUp.last;
  assert.ok(load.busy === null || (Number.isInteger(load.busy) && load.busy >= 0 && load.busy <= 100), JSON.stringify(load));
  assert.ok(load.freeGb >= 0, JSON.stringify(load));
  assert.match(said[0], new RegExp(`gave up on the dialog after ${ms} ms.*, the processors (\\d+% busy over the wait|not read over the wait), [\\d.]+ GB free$`));
  // A wait that held keeps nothing.
  gaveUp.last = null;
  assert.equal(await until(() => true, 'the menu', 200, { exact: true, log: () => {} }), true);
  assert.equal(gaveUp.last, null);
});

test('the failure list says the load beside the cause, and a cause recorded before it is said as before', () => {
  const entry = (cause) => ({ id: '01.0001', name: 'shell: the menu opens', at: 'verify-shell.js:10', part: 'shell#verifyShell', failures: 1, flaky: false, synthetic: false, reopened: 0, passedAt: [], last: { lane: 'parallel', cause, detail: null } });
  const now = ledgerLines([entry({ what: 'the menu', ms: 15004, at: 'verify-shell.js:9', load: { busy: 92, freeGb: 2.6 } })]).join('\n');
  assert.match(now, /after the wait for the menu gave up at 15004 ms \(verify-shell\.js:9\), the processors 92% busy over the wait, 2\.6 GB free/);
  const before = ledgerLines([entry({ what: 'the menu', ms: 15004, at: 'verify-shell.js:9' })]).join('\n');
  assert.match(before, /gave up at 15004 ms \(verify-shell\.js:9\)$/m);
  assert.doesNotMatch(before, /busy/);
});
