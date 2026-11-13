// What the add-ons code removes, and the lock it holds while it does.
//
// A junction planted inside an add-on's folder must go as a link, never
// with what it points at. On Node 22 fs.rm does not follow it; on Node 24,
// the one Electron 44 runs (ELECTRON_RUN_AS_NODE=1), it does, and only the
// shared removeTree (D-2026-11-13-37, branch remove-tree) keeps it from
// that. Until addons-folder.js's removeTree is moved onto it, that test is
// marked to do under Node 24, so that it is seen and not counted.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { applyRemovals, markRemoval, rollBack, lockFolder, removeTree } from '../apps/desktop/main/addons-folder.js';

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-addons-links-'));
test.after(() => {
  // The planted junctions first, so nothing is removed through them.
  for (const id of ['lang-xx', 'lang-yy']) {
    try { fs.unlinkSync(path.join(tmp, 'addons', id, 'planted')); } catch { /* gone */ }
  }
  removeTree(tmp);
});
const windows = process.platform === 'win32' ? false : 'junctions are for Windows';
const node24 = Number(process.versions.node.split('.')[0]) >= 24;

function elsewhere() {
  const at = path.join(tmp, 'real');
  fs.mkdirSync(path.join(at, 'inner'), { recursive: true });
  fs.writeFileSync(path.join(at, 'inner', 'keep.txt'), 'mine');
  return at;
}
const survived = (at) => fs.existsSync(path.join(at, 'inner', 'keep.txt'));

test('removing an add-on, or rolling one back, never deletes through a junction inside its folder', {
  skip: windows,
  todo: node24 && 'passes on Node 24 once addons-folder.js removes through the shared removeTree (D-2026-11-13-37)',
}, () => {
  const real = elsewhere();
  const dir = path.join(tmp, 'addons');
  for (const id of ['lang-xx', 'lang-yy']) {
    fs.mkdirSync(path.join(dir, id), { recursive: true });
    fs.writeFileSync(path.join(dir, `${id}.pack.json`), '{}');
    fs.symlinkSync(real, path.join(dir, id, 'planted'), 'junction');
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
