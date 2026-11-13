// What the add-ons code removes, and the lock it holds while it does.
//
// A junction planted inside an add-on's folder must go as a link, never
// with what it points at. On Node 22 fs.rm does not follow it; on Node 24,
// the one Electron 44 runs (ELECTRON_RUN_AS_NODE=1), it does, and the shared
// removeTree (D-2026-11-13-37) that addons-folder.js removes through keeps
// it from that. Run under both.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { applyRemovals, markRemoval, rollBack, putInPlace, lockFolder, removeTree } from '../apps/desktop/main/addons-folder.js';

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-addons-links-'));
const planted = [];
const plant = (target, at) => {
  fs.symlinkSync(target, at, 'junction');
  planted.push(at);
};
test.after(() => {
  // The planted junctions first, so nothing is removed through them.
  for (const at of planted) {
    try { fs.unlinkSync(at); } catch { /* gone */ }
  }
  removeTree(tmp);
});
const windows = process.platform === 'win32' ? false : 'junctions are for Windows';

function elsewhere() {
  const at = path.join(tmp, 'real');
  fs.mkdirSync(path.join(at, 'inner'), { recursive: true });
  fs.writeFileSync(path.join(at, 'inner', 'keep.txt'), 'mine');
  return at;
}
const survived = (at) => fs.existsSync(path.join(at, 'inner', 'keep.txt'));

test('removing an add-on, or rolling one back, never deletes through a junction inside its folder', { skip: windows }, () => {
  const real = elsewhere();
  const dir = path.join(tmp, 'addons');
  for (const id of ['lang-xx', 'lang-yy']) {
    fs.mkdirSync(path.join(dir, id), { recursive: true });
    fs.writeFileSync(path.join(dir, `${id}.pack.json`), '{}');
    plant(real, path.join(dir, id, 'planted'));
  }
  markRemoval(dir, 'lang-xx');
  assert.deepEqual(applyRemovals(dir).removed, ['lang-xx']);
  assert.ok(!fs.existsSync(path.join(dir, 'lang-xx')));
  assert.ok(survived(real), 'removed: the planted junction went, not what it points at');
  fs.mkdirSync(path.join(dir, 'lang-yy.previous', 'lang-yy'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'lang-yy.previous', 'lang-yy.pack.json'), '{}');
  try {
    rollBack(dir, 'lang-yy');
  } catch {
    // whether it can roll back is not the point here
  }
  assert.ok(survived(real), 'rolled back: the same');
});

/** Whether anything, a link included, is at `p` (existsSync would follow a junction). */
const there = (p) => {
  try {
    fs.lstatSync(p);
    return true;
  } catch {
    return false;
  }
};

test('an envelope that is a junction goes as a link when an add-on is removed, rolled back or put in place, and what it points at stays whole', { skip: windows }, () => {
  // On Node 24, fs.rmSync of a junction as the path throws ERR_FS_EISDIR and leaves it.
  const real = elsewhere();
  const dir = path.join(tmp, 'envelopes');
  fs.mkdirSync(dir, { recursive: true });

  // Removed at the start.
  fs.mkdirSync(path.join(dir, 'lang-aa'));
  plant(real, path.join(dir, 'lang-aa.pack.json'));
  markRemoval(dir, 'lang-aa');
  assert.deepEqual(applyRemovals(dir), { removed: ['lang-aa'], held: [] });
  assert.ok(!there(path.join(dir, 'lang-aa.pack.json')), 'removed: the envelope link is gone');
  assert.ok(!there(path.join(dir, 'lang-aa')));
  assert.ok(survived(real), 'removed: what it pointed at stays whole');

  // Rolled back.
  fs.mkdirSync(path.join(dir, 'lang-bb'));
  plant(real, path.join(dir, 'lang-bb.pack.json'));
  fs.mkdirSync(path.join(dir, 'lang-bb.previous', 'lang-bb'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'lang-bb.previous', 'lang-bb.pack.json'), '{"was":1}');
  assert.equal(rollBack(dir, 'lang-bb'), true);
  assert.ok(fs.lstatSync(path.join(dir, 'lang-bb.pack.json')).isFile(), 'rolled back: the kept envelope is in place of the link');
  assert.equal(fs.readFileSync(path.join(dir, 'lang-bb.pack.json'), 'utf8'), '{"was":1}');
  assert.ok(!there(path.join(dir, 'lang-bb.previous')));
  assert.ok(survived(real), 'rolled back: what it pointed at stays whole');

  // Half an add-on (an envelope with no folder) replaced by a staged one.
  plant(real, path.join(dir, 'lang-cc.pack.json'));
  const from = path.join(tmp, 'staged');
  fs.mkdirSync(path.join(from, 'lang-cc'), { recursive: true });
  fs.writeFileSync(path.join(from, 'lang-cc', 'data.txt'), 'new');
  fs.writeFileSync(path.join(from, 'lang-cc.pack.json'), '{"new":1}');
  putInPlace(dir, from, 'lang-cc');
  assert.ok(fs.lstatSync(path.join(dir, 'lang-cc.pack.json')).isFile(), 'put in place: the staged envelope is in place of the link');
  assert.equal(fs.readFileSync(path.join(dir, 'lang-cc.pack.json'), 'utf8'), '{"new":1}');
  assert.ok(survived(real), 'put in place: what it pointed at stays whole');
});

test('the lock is written again between the steps of a long hold, so it is never stale while held', () => {
  const dir = path.join(tmp, 'locked');
  let clock = Date.parse('2026-11-13T12:00:00Z');
  const release = lockFolder(dir, { pid: 4242, now: () => clock });
  clock += 50000;
  release.refresh();
  assert.equal(Date.parse(JSON.parse(fs.readFileSync(path.join(dir, '.lock'), 'utf8')).at), clock);
  clock += 50000;
  assert.equal(lockFolder(dir, { pid: 4343, isAlive: () => true, now: () => clock }), null, '100 s held, written again 50 s ago: not stale');
  release();
  assert.ok(!fs.existsSync(path.join(dir, '.lock')));
  release.refresh();
  assert.ok(!fs.existsSync(path.join(dir, '.lock')), 'a lock no longer held is not made again');
});
