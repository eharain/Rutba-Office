// The gate's lanes and its run again at the end (owner, 2026-10-21: "Queue
// the failed ones to run at the end"): lanes started only while there is
// memory for them and added as it comes free; the parts to run again picked
// from what failed, a failure outside any part running its chunk whole; a run
// again that clears only when each check that failed is seen passing by its
// name; and the record naming each check that ran again.
import test from 'node:test';
import assert from 'node:assert/strict';
import { laneCap, poolByMemory, partsToRerun, rerunCleared, gateSection } from '../tools/gate.js';

test('the lanes the processors allow: one to each two and a half, one to four', () => {
  assert.equal(laneCap(1), 1);
  assert.equal(laneCap(8), 3);
  assert.equal(laneCap(32), 4);
});

test('a lane starts beside others only while memory is free for it, and more are added as it comes free', async () => {
  let free = 2.7;
  let clock = 0;
  let running = 0;
  let most = 0;
  const shorts = [];
  const gates = [];
  const tasks = Array.from({ length: 4 }, () => () => new Promise((resolve) => {
    running++;
    most = Math.max(most, running);
    gates.push(() => { running--; resolve(); });
  }));
  const done = poolByMemory(tasks, { cap: 3, freeGb: () => free, now: () => clock, tickMs: 5, onShort: (f, r) => shorts.push([f, r]) });
  await new Promise((r) => setTimeout(r, 20));
  assert.equal(running, 1, 'with 2.7 GB free, one lane: the second would leave less than 1.5 GB');
  assert.deepEqual(shorts[0], [2.7, 1]);
  free = 6; clock = 60_000; // memory came free, and the first lane has settled
  await new Promise((r) => setTimeout(r, 30));
  assert.equal(running, 3, 'the cap reached once there is room');
  while (gates.length) { gates.shift()(); await new Promise((r) => setTimeout(r, 10)); }
  await done;
  assert.equal(most, 3);
});

test('a lane just started counts as taking its share, so one reading of free memory does not start them all', async () => {
  let running = 0;
  const release = [];
  const tasks = Array.from({ length: 4 }, () => () => new Promise((resolve) => { running++; release.push(resolve); }));
  const done = poolByMemory(tasks, { cap: 4, freeGb: () => 3.5, now: () => 0, tickMs: 5 });
  await new Promise((r) => setTimeout(r, 20));
  assert.equal(running, 2, '3.5 GB: the first lane, and one more with 1.5 GB left for it after the first');
  for (;;) { const r = release.shift(); if (!r) { await new Promise((x) => setTimeout(x, 10)); if (!release.length) break; continue; } r(); await new Promise((x) => setTimeout(x, 10)); }
  await done;
});

test('the parts to run again: those that failed, by lane; a failure outside the parts runs its chunk whole', () => {
  const apps = {
    sheets: { ok: false, failed: ['sheets: a picture'], failedParts: ['sheets#12'] },
    word: { ok: true, failed: [] },
    access: { ok: false, failed: ['a11y-audit: mail'], failedParts: ['access#1'] },
    look: { ok: false, failed: ['look: set-up'], failedParts: ['*'] },
  };
  const machine = { ok: false, failed: ['machine: probe'], byApp: { machine: { failed: ['machine: probe'], failedParts: ['machine#1'] }, mail: { failed: [], failedParts: [] } } };
  const r = partsToRerun(apps, machine);
  assert.deepEqual(r.parallel.apps, ['sheets', 'access', 'look']);
  assert.deepEqual(r.parallel.ids, ['sheets#12', 'access#1']);
  assert.deepEqual(r.parallel.whole, ['look']);
  assert.deepEqual(r.parallel.names, ['sheets: a picture', 'a11y-audit: mail', 'look: set-up']);
  assert.deepEqual(r.machine.ids, ['machine#1']);
  assert.deepEqual(partsToRerun({ word: { ok: true, failed: [] } }, { ok: true }).parallel.names, []);
});

test('a run again clears only when each check that failed is seen passing by name, and nothing in it fails', () => {
  const names = ['sheets: a picture'];
  assert.equal(rerunCleared(names, { sheets: { failed: [], passedNames: ['sheets: a picture', 'sheets: saved'] } }), true);
  assert.equal(rerunCleared(names, { sheets: { failed: [], passedNames: ['sheets: something else'] } }), false, 'a part numbered otherwise cannot clear it');
  assert.equal(rerunCleared(names, { sheets: { failed: ['sheets: saved'], passedNames: ['sheets: a picture'] } }), false);
  assert.equal(rerunCleared(names, null), false);
});

test('a part that threw is cleared when its app\'s checks pass in the run again, since its catch is never seen passing', () => {
  const names = ['sheets: the error checking check ran'];
  assert.equal(rerunCleared(names, { sheets: { failed: [], passedNames: ['sheets: Error Checking lists D8'] } }), true);
  assert.equal(rerunCleared(names, { pdf: { failed: [], passedNames: ['pdf: Delete takes the page out'] } }), false, 'another app\'s passes do not clear it');
  assert.equal(rerunCleared(names, { sheets: { failed: ['sheets: the error checking check ran'], passedNames: ['sheets: Error Checking lists D8'] } }), false);
  assert.equal(rerunCleared(['sheets: a picture'], { sheets: { failed: [], passedNames: ['sheets: saved'] } }), false, 'only a catch is cleared this way');
});

test('the record names each check that ran again', () => {
  const counts = { 'verify:apps': { passed: 10, of: 10 } };
  const plain = gateSection(counts, { version: '9.9.9', date: '2026-10-21', platform: 'win32', arch: 'x64' });
  assert.doesNotMatch(plain, /run again/);
  const again = gateSection(counts, { version: '9.9.9', date: '2026-10-21', platform: 'win32', arch: 'x64', reran: ['sheets: a picture'] });
  assert.match(again, /One check failed beside the others and passed when run again alone/);
  assert.match(again, /\n- sheets: a picture\n$/);
});
