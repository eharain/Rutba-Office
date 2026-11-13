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

test('the tries again after a busy file back off as rm\'s retryDelay does: delayMs * n before the n-th', async () => {
  const delayMs = 50;
  const busy = () => Object.assign(new Error('busy'), { code: 'EBUSY' });
  const check = (at, what) => {
    assert.equal(at.length, 4, `${what}: three busy tries, then the one that removes it`);
    for (let k = 1; k < at.length; k += 1) {
      // A little under, for the clock; fixed pacing would give 50 before the second and third.
      assert.ok(at[k] - at[k - 1] >= delayMs * k - 10, `${what}: ${at[k] - at[k - 1]} ms before try ${k + 1}, wanted ${delayMs * k}`);
    }
  };
  const box = path.join(tmp, 'backoff');
  fs.mkdirSync(box);
  fs.writeFileSync(path.join(box, 'f.txt'), 'x');
  const at = [];
  const flaky = { ...fs, unlinkSync: (p) => { at.push(performance.now()); if (at.length <= 3) throw busy(); return fs.unlinkSync(p); } };
  removeTree(box, { retries: 3, delayMs, fs: flaky });
  assert.ok(!fs.existsSync(box));
  check(at, 'removeTree');

  fs.mkdirSync(box);
  fs.writeFileSync(path.join(box, 'f.txt'), 'x');
  const atAsync = [];
  const flakyAsync = { ...fs, promises: { ...fs.promises, unlink: async (p) => { atAsync.push(performance.now()); if (atAsync.length <= 3) throw busy(); return fs.promises.unlink(p); } } };
  await removeTreeAsync(box, { retries: 3, delayMs, fs: flakyAsync });
  assert.ok(!fs.existsSync(box));
  check(atAsync, 'removeTreeAsync');
});

test('a folder whose .part a stopped download still holds open is removed without an error (Cancel), and gone once it closes', async () => {
  // fs.rm fails this on Node 24 with ENOTEMPTY; the lstat walk unlinks the file first.
  const parts = path.join(tmp, 'parts', 'lang-xx');
  fs.mkdirSync(parts, { recursive: true });
  const stream = fs.createWriteStream(path.join(parts, 'lang-xx.pack.part'));
  await new Promise((r) => stream.write(Buffer.alloc(64 * 1024), r));
  assert.doesNotThrow(() => removeTree(parts));
  await new Promise((r) => stream.end(r));
  removeTree(parts);
  assert.ok(!fs.existsSync(parts));
});

const both = [['removeTree', async (p, o) => removeTree(p, o)], ['removeTreeAsync', (p, o) => removeTreeAsync(p, o)]];
const enoent = (p) => Object.assign(new Error(`ENOENT: no such file or directory, scandir '${p}'`), { code: 'ENOENT' });
/** A tree to remove: a/b/f.txt and g.txt. */
function tree(name) {
  const box = path.join(tmp, `${name}-${(n += 1)}`);
  fs.mkdirSync(path.join(box, 'a', 'b'), { recursive: true });
  fs.writeFileSync(path.join(box, 'a', 'b', 'f.txt'), 'x');
  fs.writeFileSync(path.join(box, 'g.txt'), 'x');
  return box;
}

test('a folder that goes while the walk is under way counts as removed: readdir ENOENT mid-walk', async () => {
  // Another removal of the same tree took a/ between the lstat and the readdir.
  const box = tree('mid-walk');
  const sub = path.join(box, 'a');
  const takeIt = () => removeTree(sub);
  const flaky = { ...fs, readdirSync: (p, ...rest) => { if (p === sub) { takeIt(); throw enoent(p); } return fs.readdirSync(p, ...rest); } };
  removeTree(box, { fs: flaky });
  assert.ok(!fs.existsSync(box));

  const box2 = tree('mid-walk');
  const sub2 = path.join(box2, 'a');
  const flakyAsync = { ...fs, promises: { ...fs.promises, readdir: async (p, ...rest) => { if (p === sub2) { removeTree(sub2); throw enoent(p); } return fs.promises.readdir(p, ...rest); } } };
  await removeTreeAsync(box2, { fs: flakyAsync });
  assert.ok(!fs.existsSync(box2));
});

test('removals of the same tree at once all succeed', async () => {
  const box = tree('at-once');
  for (let i = 0; i < 40; i += 1) {
    fs.mkdirSync(path.join(box, `d${i}`, 'e'), { recursive: true });
    fs.writeFileSync(path.join(box, `d${i}`, 'e', 'f.txt'), 'x');
  }
  const results = await Promise.allSettled(Array.from({ length: 8 }, () => removeTreeAsync(box)));
  assert.deepEqual(results.map((r) => r.status === 'rejected' ? `${r.reason.code} ${r.reason.syscall}` : r.status), Array(8).fill('fulfilled'));
  assert.ok(!fs.existsSync(box));
});

test('no path, or a filesystem root, is refused with nothing touched', async () => {
  const touched = [];
  const record = (from, prefix) => new Proxy(from, { get: (t, k) => (typeof t[k] === 'function' ? (...a) => { touched.push(`${prefix}${String(k)}`); return t[k](...a); } : t[k]) });
  const spyPromises = record(fs.promises, 'promises.');
  const spy = new Proxy(record(fs, ''), { get: (t, k) => (k === 'promises' ? spyPromises : t[k]) });
  const roots = windows === false
    ? ['C:\\', 'C:', 'C:\\.', '\\\\server\\share\\', '\\\\server\\share', '\\\\?\\C:\\', '\\\\?\\UNC\\server\\share\\', '\\\\?\\UNC\\server\\share', '/', '\\']
    : ['/', '/.', '//'];
  for (const [name, remove] of both) {
    for (const bad of ['', undefined, null, 42, {}]) {
      await assert.rejects(async () => remove(bad, { fs: spy }), TypeError, `${name}: ${JSON.stringify(bad)}`);
    }
    for (const root of roots) {
      await assert.rejects(async () => remove(root, { fs: spy }), /is a filesystem root/, `${name}: ${root}`);
    }
  }
  assert.deepEqual(touched, []);
  // A drive given without a folder is that drive's working folder, resolved first: not refused.
  if (windows === false) {
    const here = process.cwd();
    const box = tree('drive-relative');
    process.chdir(box);
    try {
      removeTree(box.slice(0, 2), { fs: { ...fs, lstatSync: (p) => { touched.push(p); throw enoent(p); } } });
    } finally {
      process.chdir(here);
    }
    assert.deepEqual(touched, [box], 'D: is resolved to the working folder on D: before anything is done');
    removeTree(box);
  }
});

test('where the thread may not be held, removeTree throws the busy error itself, not the wait\'s', () => {
  const box = tree('no-wait');
  const busy = Object.assign(new Error('busy'), { code: 'EBUSY' });
  let tries = 0;
  const flaky = { ...fs, unlinkSync: (p) => { tries += 1; if (tries === 1) throw busy; return fs.unlinkSync(p); } };
  const wait = Atomics.wait;
  Atomics.wait = () => { throw new TypeError('Atomics.wait cannot be called in this context'); };
  try {
    assert.throws(() => removeTree(box, { retries: 2, delayMs: 1, fs: flaky }), (err) => err === busy);
  } finally {
    Atomics.wait = wait;
  }
  assert.equal(tries, 1);
  removeTree(box);
  assert.ok(!fs.existsSync(box));
});

/** A symbolic link, or a reason the test cannot make one here. */
function symlinkOr(target, at, type) {
  try {
    fs.symlinkSync(target, at, type);
    return null;
  } catch (err) {
    return `no symbolic links here (${err.code})`;
  }
}

for (const [name, remove] of both) {
  test(`${name} removes a symbolic link to a folder and one to a file as links, inside a folder and as the path`, async (t) => {
    const real = elsewhere();
    const file = path.join(real, 'inner', 'keep.txt');
    const box = tree('symlinks');
    const cannot = symlinkOr(real, path.join(box, 'a', 'dir-link'), 'dir');
    if (cannot) return t.skip(cannot);
    assert.equal(symlinkOr(file, path.join(box, 'a', 'b', 'file-link'), 'file'), null);
    await remove(box);
    assert.ok(!fs.existsSync(box));
    assert.ok(survived(real), 'what the links pointed at is untouched');

    const dirLink = path.join(tmp, `dir-link-${(n += 1)}`);
    const fileLink = path.join(tmp, `file-link-${(n += 1)}`);
    assert.equal(symlinkOr(real, dirLink, 'dir'), null);
    assert.equal(symlinkOr(file, fileLink, 'file'), null);
    await remove(dirLink);
    await remove(fileLink);
    for (const p of [dirLink, fileLink]) assert.throws(() => fs.lstatSync(p), { code: 'ENOENT' });
    assert.ok(survived(real));
  });

  test(`${name} removes a read-only folder with read-only files in it`, async () => {
    const box = tree('read-only');
    fs.chmodSync(path.join(box, 'a', 'b', 'f.txt'), 0o444);
    fs.chmodSync(path.join(box, 'a', 'b'), 0o555);
    fs.chmodSync(path.join(box, 'a'), 0o555);
    fs.chmodSync(box, 0o555);
    await remove(box);
    assert.ok(!fs.existsSync(box));
  });
}
