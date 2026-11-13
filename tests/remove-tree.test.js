// A folder removed with everything in it, never through a link
// (packages/office-shell/src/electron/remove-tree.js): from Node 24, the one
// Electron 44 runs, fs.rm with `recursive` follows a Windows junction and
// deletes what it points at. Run here, and under Electron's own Node
// (ELECTRON_RUN_AS_NODE=1), where rm would fail these.

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { removeTree, removeTreeAsync } from '../packages/office-shell/src/electron/remove-tree.js';
import { applyRemovals, markRemoval, rollBack, lockFolder } from '../apps/desktop/main/addons-folder.js';

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-remove-tree-'));
test.after(() => removeTree(tmp));
const windows = process.platform === 'win32' ? false : 'junctions are for Windows';

let n = 0;
/** A real folder elsewhere, with a file that must survive. */
function elsewhere() {
  const at = path.join(tmp, `real-${(n += 1)}`);
  fs.mkdirSync(path.join(at, 'inner'), { recursive: true });
  fs.writeFileSync(path.join(at, 'inner', 'keep.txt'), 'mine');
  return at;
}
const survived = (at) => fs.readFileSync(path.join(at, 'inner', 'keep.txt'), 'utf8') === 'mine';

for (const [name, remove] of [['removeTree', async (p) => removeTree(p)], ['removeTreeAsync', (p) => removeTreeAsync(p)]]) {
  test(`${name} removes a folder whole, a junction in it as a link and not what it points at`, { skip: windows }, async () => {
    const real = elsewhere();
    const box = path.join(tmp, `box-${(n += 1)}`);
    fs.mkdirSync(path.join(box, 'a', 'b'), { recursive: true });
    fs.writeFileSync(path.join(box, 'a', 'b', 'f.txt'), 'x');
    fs.writeFileSync(path.join(box, 'ro.txt'), 'x');
    fs.chmodSync(path.join(box, 'ro.txt'), 0o444);
    fs.symlinkSync(real, path.join(box, 'a', 'nested'), 'junction');
    await remove(box);
    assert.ok(!fs.existsSync(box));
    assert.ok(survived(real), 'the folder a junction inside pointed at is untouched');
  });

  test(`${name} given a junction itself removes the junction, not its folder; a path not there is no error`, { skip: windows }, async () => {
    const real = elsewhere();
    const link = path.join(tmp, `link-${(n += 1)}`);
    fs.symlinkSync(real, link, 'junction');
    await remove(link);
    assert.ok(!fs.existsSync(link));
    assert.ok(survived(real));
    await remove(path.join(tmp, 'never-made'));
  });
}

test('removeTree tries again after a busy file when asked, and says what it could not remove otherwise', () => {
  const box = path.join(tmp, 'busy');
  fs.mkdirSync(box);
  fs.writeFileSync(path.join(box, 'f.txt'), 'x');
  let tries = 0;
  const flaky = { ...fs, unlinkSync: (p) => { tries += 1; if (tries < 3) throw Object.assign(new Error('busy'), { code: 'EBUSY' }); return fs.unlinkSync(p); } };
  removeTree(box, { retries: 3, delayMs: 1, fs: flaky });
  assert.ok(!fs.existsSync(box));
  fs.mkdirSync(box);
  fs.writeFileSync(path.join(box, 'f.txt'), 'x');
  const stuck = { ...fs, unlinkSync: () => { throw Object.assign(new Error('busy'), { code: 'EBUSY' }); } };
  assert.throws(() => removeTree(box, { fs: stuck }), (err) => err.code === 'EBUSY');
  removeTree(box);
});

test('removing an add-on, or rolling one back, never deletes through a junction inside its folder', { skip: windows }, () => {
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
