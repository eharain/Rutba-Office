// The gate runs above normal priority, itself and what it starts, so other
// work beside it does not starve its windows and timed checks (the owner,
// 13 Nov 2026: "you should learn to co exisit"). Its own processes only; no
// setting of the computer is changed, and one already higher is left alone.
import test from 'node:test';
import assert from 'node:assert/strict';
import os from 'node:os';
import { spawn } from 'node:child_process';
import { raiseOwn, raiseTrees, run, priorityNote, gateSection, listProcessesAsync, treeOf } from '../tools/gate.js';

const P = os.constants.priority;

test('each process below above-normal is raised to it; one already as high or higher is left as it is', () => {
  const now = new Map([[1, P.PRIORITY_NORMAL], [2, P.PRIORITY_BELOW_NORMAL], [3, P.PRIORITY_ABOVE_NORMAL], [4, P.PRIORITY_HIGH]]);
  const set = [];
  const n = raiseOwn([1, 2, 3, 4], { get: (pid) => now.get(pid), set: (pid) => set.push(pid) });
  assert.equal(n, 2);
  assert.deepEqual(set, [1, 2]);
});

test('a process that has ended, or is not ours to raise, is passed over and the rest are raised', () => {
  const set = [];
  const n = raiseOwn([1, 2, 3], {
    get: (pid) => { if (pid === 1) throw new Error('no such process'); return P.PRIORITY_NORMAL; },
    set: (pid) => { if (pid === 2) throw new Error('access denied'); set.push(pid); },
  });
  assert.equal(n, 1);
  assert.deepEqual(set, [3]);
});

test('on this computer a process the gate started is raised for real', { skip: process.platform !== 'win32' && 'raising a priority needs the right to on Linux; the gate raises on Windows' }, async () => {
  const child = spawn(process.execPath, ['-e', 'setTimeout(() => {}, 20000)'], { stdio: 'ignore' });
  try {
    await new Promise((r) => child.once('spawn', r));
    assert.equal(os.getPriority(child.pid), P.PRIORITY_NORMAL);
    assert.equal(raiseOwn([child.pid]), 1);
    assert.equal(os.getPriority(child.pid), P.PRIORITY_ABOVE_NORMAL);
    assert.equal(raiseOwn([child.pid]), 0, 'raised once');
  } finally {
    child.kill();
  }
});

test('what the gate\'s child starts is raised too: an Electron window three down, through cmd.exe, as npm starts it', { skip: process.platform !== 'win32' && 'the gate walks trees on Windows' }, async () => {
  // cmd.exe starts node, which starts node: the gate raises its child, and the walk raises the rest.
  // /s and the outer quotes: cmd.exe takes the line as it is, quotes and all.
  // Ten minutes each: longer than the wait and the listings below can take on a busy computer; ended in the finally.
  const inner = "require('child_process').spawn(process.execPath, ['-e', 'setTimeout(() => {}, 600000)'], { stdio: 'ignore' }); setTimeout(() => {}, 600000)";
  const child = spawn('cmd.exe', ['/d', '/s', '/c', `""${process.execPath}" -e "${inner}""`], { stdio: 'ignore', windowsVerbatimArguments: true });
  child.startedAt = Date.now();
  try {
    await new Promise((r) => child.once('spawn', r));
    // The gate's own listing, taken once it has answered. On a busy computer one listing can pass its 30 s and
    // answer nothing (the gate then walks again ten seconds later), and the old processes here held only 30 s,
    // so a slow listing could also see them gone: either gave the 0 raised in the unit run of 14 Nov 2026 05:12
    // (36 s). A listing that answered also keeps "not walked again" from passing on an empty one.
    let last = [];
    const listed = async () => {
      for (const end = Date.now() + 120_000; Date.now() < end; await new Promise((r) => setTimeout(r, 1000))) {
        const procs = await listProcessesAsync();
        if (procs.length) { last = procs; return procs; }
      }
      throw new Error('no listing of the processes in two minutes');
    };
    // Until the grandchild has started and the listing shows all three, with a long outer limit: beside the
    // rest of the suite a fixed ten seconds ran out before it did (D-2026-11-13-20: wait on the signal).
    let tree = [];
    const until = Date.now() + 60_000;
    while (tree.length < 3 && Date.now() < until) {
      await new Promise((r) => setTimeout(r, 250));
      // Walked, not raised: what the tree holds now.
      await raiseTrees([child], { list: listed, raise: (ids) => { tree = ids; return 0; } });
    }
    assert.equal(tree.length, 3, `cmd.exe, node and its node: ${tree.length}`);
    assert.ok(tree.every((id) => os.getPriority(id) === P.PRIORITY_NORMAL), 'normal until raised');
    const skip = new Set();
    assert.equal(await raiseTrees([child], { skip, list: listed }), 3);
    assert.ok(tree.every((id) => os.getPriority(id) === P.PRIORITY_ABOVE_NORMAL), 'each of them above normal');
    assert.equal(await raiseTrees([child], { skip, list: listed }), 0, 'those walked once are not walked again');
    // And not for want of seeing them: the listing that walk read holds all three.
    assert.deepEqual(treeOf(child.pid, child.startedAt, last).sort(), [...tree].sort(), 'the second walk saw the three');
  } finally {
    // The whole tree ended, as the gate ends its own.
    try { (await import('node:child_process')).execFileSync('taskkill', ['/pid', String(child.pid), '/T', '/F'], { stdio: 'ignore' }); } catch { /* ended already */ }
  }
});

test('an id Windows gave to a new process of the gate\'s is walked as new, not passed over as the one before it', async () => {
  // The gate's child 1 started 5 at 2000; that 5 ended, and a later child of 1 was given the id 5 at 9000.
  const first = [{ pid: 1, ppid: 0, created: 1000 }, { pid: 5, ppid: 1, created: 2000 }];
  const later = [{ pid: 1, ppid: 0, created: 1000 }, { pid: 5, ppid: 1, created: 9000 }];
  const child = { pid: 1, startedAt: 1000 };
  const skip = new Set();
  const asked = [];
  const raise = (ids) => { asked.push(ids); return ids.length; };
  assert.equal(await raiseTrees([child], { skip, list: async () => first, raise }), 2);
  assert.equal(await raiseTrees([child], { skip, list: async () => first, raise }), 0, 'the same two are not walked again');
  assert.equal(await raiseTrees([child], { skip, list: async () => later, raise }), 1);
  assert.deepEqual(asked.at(-1), [5]);
});

test('the record says the gate\'s processes ran above normal priority, as it says --no-settle, and nothing when they did not', () => {
  const counts = { 'verify:apps': { passed: 10, of: 10 } };
  const at = { version: '9.9.9', date: '2026-11-13', platform: 'win32', arch: 'x64' };
  assert.equal(priorityNote(null), '');
  assert.equal(priorityNote(['normal']), '');
  assert.doesNotMatch(gateSection(counts, { ...at, priority: ['normal'] }), /priority/);
  const raised = gateSection(counts, { ...at, priority: ['raised'] }).replace(/\s+/g, ' ');
  assert.match(raised, /The gate raised its own processes on this computer to above normal priority \(those that lived past its ten-second walk\), beside whatever else the computer was doing; every budget and limit is as before\./);
  const mixed = gateSection(counts, { ...at, priority: ['normal', 'raised'] }).replace(/\s+/g, ' ');
  assert.match(mixed, /For part of the run \(a resumed one\), the gate raised its own processes/);
});

test('the gate\'s run, imported by a test, raises nothing: only the gate run as a command does', async () => {
  const before = os.getPriority();
  const out = await run(process.execPath, ['-e', 'console.log(require("os").getPriority())'], {}, { shell: false, whole: true });
  assert.equal(out.code, 0);
  assert.equal(Number(out.tail.trim()), P.PRIORITY_NORMAL);
  assert.equal(os.getPriority(), before);
});
