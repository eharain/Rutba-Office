// Faults a peer review found in the Locker, 10 Nov 2026, each reproduced on
// the code as it was before being fixed. The first of them lost every file in
// a folder without one error shown.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { FOLDER_KEYS, createKeys, newKdf, passwordKey } from '../apps/desktop/main/extras/locker/keys.js';
import { wipeFile, LARGE } from '../apps/desktop/main/extras/locker/folder.js';
import { newHeader, sealFrame, MODE_FOLDER } from '../apps/desktop/main/extras/locker/format.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
const tmp = (pre = 'rl-') => fs.mkdtempSync(path.join(os.tmpdir(), pre));

async function make(hooks = null, choices = {}) {
  const dir = tmp();
  const folder = tmp('rl-folder-');
  const locker = createLocker({ dir, cost: CHEAP, hooks });
  const { recovery } = await locker.setup('password one');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: false, keepVersions: 3, ...choices });
  return { locker, folder, id, recovery, dir };
}

const write = (folder, name, size = 2000) => {
  const bytes = crypto.randomBytes(size);
  fs.writeFileSync(path.join(folder, name), bytes);
  return bytes;
};

const otherLocker = async (password = 'someone else') => {
  const locker = createLocker({ dir: tmp(), cost: CHEAP });
  await locker.setup(password);
  return locker;
};

test('the locker closing during a first lock damages nothing: what was not reached stays plain', async () => {
  let reached = 0;
  let release;
  let arrived;
  const third = new Promise((r) => { arrived = r; });
  const gate = new Promise((r) => { release = r; });
  const hooks = { at: async (step) => {
    if (step !== 'before-delete') return;
    reached += 1;
    if (reached === 3) { arrived(); await gate; }
  } };
  const { locker, folder, id } = await make(hooks);
  const files = new Map();
  for (let i = 0; i < 12; i++) files.set(`f${i}.bin`, write(folder, `f${i}.bin`));

  const running = locker.lockWaiting(id);
  await third;
  let lockEnded = false;
  const closing = locker.lock({}).then((r) => { lockEnded = true; return r; });
  await new Promise((r) => setTimeout(r, 80));
  assert.equal(lockEnded, false, 'the keys are not wiped while a file is being locked');
  release();
  const out = await running;
  await closing;
  assert.equal(locker.isOpen(), false);
  assert.ok(out.locked.length >= 3 && out.locked.length < 12, `${out.locked.length} locked before it stopped`);
  assert.equal(out.locked.length + out.failed.length, 12);
  assert.ok(out.failed.every((f) => f.code === 'stopped'));

  await locker.unlock({ password: 'password one' });
  const list = await locker.scan(id);
  assert.equal(list.filter((e) => e.state === 'damaged').length, 0, 'no file is damaged');
  assert.equal(list.filter((e) => e.state === 'locked').length, out.locked.length);
  assert.equal(list.filter((e) => e.state === 'waiting').length, out.failed.length);
  const dest = tmp('rl-out-');
  for (const e of list.filter((x) => x.state === 'locked')) {
    await locker.exportCopy(id, e.rel, path.join(dest, e.name));
    assert.ok(fs.readFileSync(path.join(dest, e.name)).equals(files.get(e.name)), `${e.name} reads back`);
  }
});

test('new work is refused while the locker closes, and two closings are one', async () => {
  let release;
  let arrived;
  const there = new Promise((r) => { arrived = r; });
  const gate = new Promise((r) => { release = r; });
  const hooks = { at: async (step) => { if (step === 'before-delete') { arrived(); await gate; } } };
  const { locker, folder, id } = await make(hooks);
  write(folder, 'a.bin');
  const running = locker.lockWaiting(id);
  await there;
  const first = locker.lock({});
  const second = locker.lock({});
  assert.equal(first, second);
  await new Promise((r) => setTimeout(r, 30));
  await assert.rejects(locker.scan(id), (e) => e.code === 'locked');
  await assert.rejects(locker.addFolder(tmp('rl-another-')), (e) => e.code === 'locked');
  release();
  await running;
  await first;
});

test('a folder with an unreadable item is not stopped: nothing changes and the key file stays', async () => {
  const { locker, folder, id } = await make();
  const a = write(folder, 'a.bin');
  write(folder, 'b.bin');
  await locker.lockWaiting(id);
  fs.writeFileSync(path.join(folder, 'b.bin.rlock'), crypto.randomBytes(30));
  const out = await locker.removeFolder(id);
  assert.equal(out.removed, false);
  assert.equal(out.damaged, 1);
  assert.ok(fs.existsSync(path.join(folder, FOLDER_KEYS)), 'the key file stays');
  assert.ok(fs.existsSync(path.join(folder, 'a.bin.rlock')), 'nothing was unlocked');
  assert.equal(fs.existsSync(path.join(folder, 'a.bin')), false);

  const on = await locker.removeFolder(id, { leaveDamaged: true });
  assert.equal(on.removed, true);
  assert.ok(fs.readFileSync(path.join(folder, 'a.bin')).equals(a), 'the readable file is unlocked');
  assert.ok(fs.existsSync(path.join(folder, 'b.bin.rlock')), 'the unreadable one is left');
  assert.ok(fs.existsSync(path.join(folder, FOLDER_KEYS)), 'with what might open it');
  assert.equal((await locker.folders()).length, 0);
});

test('a replaced recovery key reaches the folders: the new one opens them and the old one does not', async () => {
  const { locker, folder, id, recovery: old } = await make();
  write(folder, 'a.bin');
  await locker.lockWaiting(id);
  const out = await locker.replaceRecovery('password one');
  assert.deepEqual(out.failed, []);
  const tryOld = await otherLocker();
  await assert.rejects(tryOld.addFolder(folder, {}, { recovery: old }), (e) => e.code === 'bad-password');
  const tryNew = await otherLocker();
  const joined = await tryNew.addFolder(folder, {}, { recovery: out.recovery });
  assert.equal(joined.joined, true);
  assert.equal((await tryNew.scan(joined.id))[0].state, 'locked');
});

test('a changed password reaches the folders, and the old one stops opening them', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.bin');
  await locker.lockWaiting(id);
  await locker.changePassword('password one', 'password two');
  const old = await otherLocker('x');
  await assert.rejects(old.addFolder(folder, {}, { password: 'password one' }), (e) => e.code === 'bad-password');
  const next = await otherLocker('x');
  assert.equal((await next.addFolder(folder, {}, { password: 'password two' })).joined, true);
});

test('a lost password is replaced with the recovery key', async () => {
  const { locker, dir, recovery } = await make();
  await locker.lock({});
  await locker.unlock({ recovery });
  await assert.rejects(locker.changePassword('', 'new password', { recovery: 'AAAA-BBBB' }), (e) => e.code === 'bad-recovery');
  await locker.changePassword('', 'new password', { recovery });
  await locker.lock({});
  const again = createLocker({ dir, cost: CHEAP });
  await assert.rejects(again.unlock({ password: 'password one' }), (e) => e.code === 'bad-password');
  await again.unlock({ password: 'new password' });
  assert.equal(again.isOpen(), true);
});

test('a folder opened by a second locker keeps the first locker\'s recovery key and password', async () => {
  const { locker: a, folder, id, recovery: ra } = await make();
  const bytes = write(folder, 'a.bin');
  await a.lockWaiting(id);
  const b = await otherLocker('password one');
  const joined = await b.addFolder(folder, {}, { password: 'password one' });
  await b.lock({});
  await b.unlock({ password: 'password one' });
  const file = JSON.parse(fs.readFileSync(path.join(folder, FOLDER_KEYS), 'utf8'));
  assert.equal(file.unlockers.filter((u) => u.kind === 'recovery').length, 2, 'both lockers\' recovery keys');
  assert.equal(file.unlockers.filter((u) => u.kind === 'password').length, 2, 'both lockers\' passwords');

  const third = await otherLocker();
  const opened = await third.addFolder(folder, {}, { recovery: ra });
  const out = path.join(tmp('rl-out-'), 'a.bin');
  await third.exportCopy(opened.id, 'a.bin', out);
  assert.ok(fs.readFileSync(out).equals(bytes));

  // And the second locker's unlocking again leaves the file as it is.
  const before = fs.readFileSync(path.join(folder, FOLDER_KEYS), 'utf8');
  await b.lock({});
  await b.unlock({ password: 'password one' });
  assert.equal(fs.readFileSync(path.join(folder, FOLDER_KEYS), 'utf8'), before);
  assert.equal((await b.scan(joined.id))[0].state, 'locked');
});

test('a folder that has gone is not made again when the locker opens', async () => {
  const { locker, folder, id } = await make();
  await locker.lockWaiting(id);
  fs.rmSync(folder, { recursive: true, force: true });
  const out = await locker.healFolderKeys();
  assert.equal(out.failed.length, 1);
  assert.equal(fs.existsSync(folder), false);
});

test('a state file that cannot be read is an error, never a locker with no folders', async () => {
  const { locker, dir } = await make();
  fs.writeFileSync(path.join(dir, 'state.json'), '{ cut sho');
  await assert.rejects(locker.folders(), (e) => e.code === 'format');
  assert.equal(fs.readFileSync(path.join(dir, 'state.json'), 'utf8'), '{ cut sho', 'left as it was');
});

test('a key file cannot ask for more memory than a person\'s computer can spare', async () => {
  await assert.rejects(passwordKey('x', { ...newKdf({ N: 2 ** 22, r: 8, p: 1 }) }), (e) => e.code === 'format');
  await assert.rejects(passwordKey('x', { ...newKdf({ N: 2 ** 20, r: 16, p: 1 }) }), (e) => e.code === 'format');
  assert.ok(createKeys({ dir: tmp(), cost: CHEAP }));
});

test('removing a plain file overwrites all of it, and never through a link', async () => {
  const dir = tmp();
  const big = path.join(dir, 'big.bin');
  fs.writeFileSync(big, crypto.randomBytes(12 * 1024 * 1024));
  const held = fs.openSync(big, 'r');
  await wipeFile(big);
  // Removal may wait for the handle on Windows; the bytes were overwritten either way.
  const after = Buffer.alloc(12 * 1024 * 1024);
  fs.readSync(held, after, 0, after.length, 0);
  fs.closeSync(held);
  assert.equal(after.every((b) => b === 0), true, 'every byte is gone');

  const target = path.join(dir, 'target.bin');
  const bytes = crypto.randomBytes(500);
  fs.writeFileSync(target, bytes);
  const link = path.join(dir, 'link.bin');
  try {
    fs.symlinkSync(target, link);
  } catch {
    return; // Links need a right this account may not have.
  }
  await wipeFile(link);
  assert.ok(fs.readFileSync(target).equals(bytes), 'what the link pointed at is untouched');
  assert.equal(fs.existsSync(link), false);
});

test('a large file that shrinks leaves one locked copy, not two', async () => {
  const { locker, folder, id } = await make();
  const name = 'big.bin';
  fs.writeFileSync(path.join(folder, name), crypto.randomBytes(LARGE + 1024));
  await locker.lockWaiting(id);
  assert.ok(fs.existsSync(path.join(folder, `${name}.rlock.d`)));
  const small = write(folder, name, 5000);
  await locker.lockWaiting(id);
  assert.ok(fs.existsSync(path.join(folder, `${name}.rlock.d`)), 'the pieces stay: the earlier version is in them');
  assert.equal(fs.existsSync(path.join(folder, `${name}.rlock`)), false, 'and there is no second copy beside them');
  const list = await locker.scan(id);
  assert.deepEqual(list.map((e) => e.state), ['locked']);
  assert.equal((await locker.versions(id, name)).length, 2, 'the version before it can still be had');
  const out = path.join(tmp('rl-out-'), name);
  await locker.exportCopy(id, name, out);
  assert.ok(fs.readFileSync(out).equals(small));
});

test('sealing the same frame twice never uses the same nonce: the epoch is random', () => {
  const key = crypto.randomBytes(32);
  const hash = crypto.randomBytes(32);
  const plain = Buffer.alloc(100, 7);
  const seen = new Set();
  for (let i = 0; i < 200; i++) seen.add(sealFrame(key, hash, 0, false, plain).subarray(0, 7).toString('hex'));
  assert.equal(seen.size, 200);
  assert.ok(newHeader({ mode: MODE_FOLDER, folderId: crypto.randomBytes(16) }));
});

test('a working folder another computer made in a shared folder is left alone', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.bin');
  await locker.lockWaiting(id);
  const theirs = path.join(folder, '.rutba-open', 'aaaaaaaaaaaa-0011223344556677');
  fs.mkdirSync(theirs, { recursive: true });
  fs.writeFileSync(path.join(theirs, 'in use there.txt'), 'somebody is working on this');
  await locker.lock({});
  await locker.unlock({ password: 'password one' });
  assert.equal(fs.readFileSync(path.join(theirs, 'in use there.txt'), 'utf8'), 'somebody is working on this');
});

test('a second program cannot unlock a profile another has open, and a crash does not stop it for good', async () => {
  const { spawn } = await import('node:child_process');
  const { dir, locker } = await make();
  const mark = path.join(dir, 'in-use');
  assert.ok(fs.existsSync(mark), 'marked while it is open');
  await locker.lock();
  assert.equal(fs.existsSync(mark), false, 'cleared when it locks');
  const other = spawn(process.execPath, ['-e', 'setTimeout(() => {}, 60000)'], { stdio: 'ignore' });
  try {
    fs.writeFileSync(mark, JSON.stringify({ pid: other.pid }));
    await assert.rejects(() => locker.unlock({ password: 'password one' }), (e) => e.code === 'busy' && /another program/.test(e.message));
    assert.equal(locker.isOpen(), false, 'refused before any key was opened');
    const old = new Date(Date.now() - 10 * 60000);
    fs.utimesSync(mark, old, old);
    await locker.unlock({ password: 'password one' });
    assert.equal(locker.isOpen(), true, 'a mark nobody has touched for minutes is not believed');
    await locker.lock();
  } finally {
    other.kill();
  }
  await new Promise((r) => { other.once('exit', r); });
  fs.writeFileSync(mark, JSON.stringify({ pid: other.pid }));
  await locker.unlock({ password: 'password one' });
  assert.equal(locker.isOpen(), true, 'a mark from a process that has ended is not believed');
  await locker.lock();
});
