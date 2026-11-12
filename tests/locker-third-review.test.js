// Faults the third review of the Locker found, 10 Nov 2026, each reproduced on the code as it stood.
import { test } from 'node:test';
import { tempDir } from './temp-dirs.js';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { wipeFile } from '../apps/desktop/main/extras/locker/folder.js';
import { sweepWorking, isWorkingFile } from '../apps/desktop/main/extras/locker/place.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
// The size over which these tests' lockers keep a file as pieces: small, so the pieces are made in seconds; the
// 50 MB the app uses runs in locker-large.test.js, on the lab machine (12 Nov 2026).
const SMALL = 4 * 1024 * 1024;
const tmp = (pre = 'rl3-') => tempDir(pre);

async function make(choices = {}) {
  const { large, ...folderChoices } = choices;
  const dir = tmp();
  const folder = tmp('rl3-folder-');
  const locker = createLocker({ dir, cost: CHEAP, large });
  await locker.setup('password one');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: false, keepVersions: 3, ...folderChoices });
  return { locker, folder, id, dir };
}
const write = (folder, name, size = 2000) => {
  const bytes = crypto.randomBytes(size);
  fs.writeFileSync(path.join(folder, name), bytes);
  return bytes;
};
const exported = async (locker, id, name) => {
  const out = path.join(tmp('rl3-out-'), name);
  await locker.exportCopy(id, name, out);
  return fs.readFileSync(out);
};

test('a wipe stopped after the zeros are written leaves no file under the real name', async (t) => {
  const dir = tmp();
  const file = path.join(dir, 'a.bin');
  fs.writeFileSync(file, crypto.randomBytes(3000));
  // What a power cut leaves: the zeros written, the removal never reached.
  t.mock.method(fsp, 'rm', async () => { throw new Error('the machine stopped'); });
  await wipeFile(file).catch(() => {});
  t.mock.restoreAll();
  assert.equal(fs.existsSync(file), false, 'not under its own name, to be locked as the next version');
  const left = fs.readdirSync(dir);
  assert.equal(left.length, 1);
  assert.ok(isWorkingFile(left[0]), `a working file the scan never lists: ${left[0]}`);
});

test('a locked file with one version kept survives a wipe that stopped: the zeros are not locked over it', async (t) => {
  const { locker, folder, id } = await make({ keepVersions: 1 });
  const bytes = write(folder, 'a.bin');
  await locker.lockWaiting(id);
  write(folder, 'a.bin', 2000);
  // The plain copy of the next version is locked, and its wipe stops short.
  const real = fsp.rm;
  t.mock.method(fsp, 'rm', async (p, o) => {
    if (path.dirname(String(p)) === folder && (isWorkingFile(path.basename(String(p))) || String(p).endsWith('a.bin'))) throw new Error('the machine stopped');
    return real(p, o);
  });
  await locker.lockWaiting(id);
  t.mock.restoreAll();
  const locked = await exported(locker, id, 'a.bin');
  assert.ok(!locked.equals(Buffer.alloc(locked.length)), 'what is locked is not zeros');
  await locker.lockWaiting(id);
  const again = await exported(locker, id, 'a.bin');
  assert.ok(!again.equals(Buffer.alloc(again.length)), 'nor after the next pass');
  assert.equal(again.length, 2000);
  void bytes;
});

test('working files left in a subfolder are overwritten and swept at the next lock', async () => {
  const { locker, folder, id } = await make();
  fs.mkdirSync(path.join(folder, 'sub', 'deep'), { recursive: true });
  const stray = path.join(folder, 'sub', 'deep', `.rl-${'a'.repeat(24)}.part`);
  fs.writeFileSync(stray, 'plain words left by a stop');
  await locker.lockWaiting(id);
  assert.equal(fs.existsSync(stray), false);
});

test('the sweep of the stage overwrites what it removes when told how', async () => {
  const dir = tmp();
  const stray = path.join(dir, `.rl-${'b'.repeat(24)}.part`);
  fs.writeFileSync(stray, 'plain');
  const seen = [];
  assert.equal(await sweepWorking(dir, async (p) => { seen.push(p); }), 1);
  assert.deepEqual(seen, [stray]);
});

test('a locked file that grows past the large size keeps its earlier versions', async () => {
  const { locker, folder, id } = await make({ keepVersions: 3, large: SMALL });
  const first = write(folder, 'grow.bin', 3000);
  await locker.lockWaiting(id);
  const second = write(folder, 'grow.bin', 4000);
  await locker.lockWaiting(id);
  const third = write(folder, 'grow.bin', SMALL + 4096);
  await locker.lockWaiting(id);
  assert.ok((await exported(locker, id, 'grow.bin')).equals(third), 'the newest is the file');
  const versions = await locker.versions(id, 'grow.bin');
  assert.equal(versions.length, 3, JSON.stringify(versions));
  const bytesOf = async (v) => {
    await locker.restore(id, 'grow.bin', v);
    return exported(locker, id, 'grow.bin');
  };
  // Each put back becomes the newest and the oldest kept then drops off, so the last of the list is the one to ask for each time.
  assert.ok((await bytesOf(2)).equals(first), 'the first');
  assert.ok((await bytesOf(2)).equals(second), 'and the one after it');
  assert.equal(fs.existsSync(path.join(folder, '.rlocker-old')) && fs.readdirSync(path.join(folder, '.rlocker-old')).length > 0, false, 'moved, not left behind as well');
});

// The copy is in the profile, so it is there when its folder is not: the save waits, and the folder is never made again at its old place.
test('a copy whose folder is not there (its drive away) is kept, with its change, for when it comes back', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.txt');
  await locker.lockWaiting(id);
  const { path: copy } = await locker.openPlain(id, 'a.txt');
  fs.writeFileSync(copy, 'changed while the drive was in');
  const away = `${folder}-away`;
  fs.renameSync(folder, away);
  await assert.rejects(() => locker.saveOpen(copy), (e) => e.code === 'away');
  assert.equal(fs.existsSync(folder), false, 'the folder is not made again to take the save');
  const locked = await locker.lock();
  assert.deepEqual(locked.left, [], 'not reported as still open in a program at each lock: it waits for its folder');
  assert.equal(fs.existsSync(folder), false);
  assert.equal(fs.readFileSync(copy, 'utf8'), 'changed while the drive was in', 'kept, change and all');
  await locker.unlock({ password: 'password one' });
  assert.equal(locker.isOpenCopy(copy), true, 'still on the list while its drive is away');
  fs.renameSync(away, folder);
  assert.equal(fs.readFileSync(copy, 'utf8'), 'changed while the drive was in', 'and nothing was done to it');
  await locker.lock({ keepCopies: true }).catch(() => {});
  await locker.unlock({ password: 'password one' });
  assert.equal((await exported(locker, id, 'a.txt')).toString(), 'changed while the drive was in', 'its changes reach the locked file once it is back');
});

test('a folder moved with a copy open: the copy stays where the program has it, and its save goes to the folder\'s new place', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.txt');
  await locker.lockWaiting(id);
  const { path: copy } = await locker.openPlain(id, 'a.txt');
  fs.writeFileSync(copy, 'edited');
  const to = path.join(tmp('rl3-new-'), 'moved');
  fs.renameSync(folder, to);
  const out = await locker.addFolder(to);
  assert.equal(out.moved, true);
  assert.equal(fs.readFileSync(copy, 'utf8'), 'edited', 'not wiped, and not moved: it was never in the folder');
  assert.equal(locker.isOpenCopy(copy), true);
  assert.ok((await locker.saveOpen(copy)).changed);
  assert.equal((await exported(locker, id, 'a.txt')).toString(), 'edited');
  assert.equal(fs.existsSync(folder), false, 'nothing made at the old place');
});

test('a working folder with no owner in its name is left while it may be another computer\'s', async () => {
  const { locker, folder, id, dir } = await make();
  void dir;
  const sub = path.join(folder, '.rutba-open', '0123456789abcdef');
  fs.mkdirSync(sub, { recursive: true });
  fs.writeFileSync(path.join(sub, 'a.txt'), 'open on the other computer');
  await locker.lock().catch(() => {});
  await locker.unlock({ password: 'password one' });
  assert.ok(fs.existsSync(path.join(sub, 'a.txt')), 'recent: left');
  const old = new Date(Date.now() - 3 * 24 * 3600 * 1000);
  fs.utimesSync(sub, old, old);
  await locker.lock().catch(() => {});
  await locker.unlock({ password: 'password one' });
  assert.equal(fs.existsSync(sub), false, 'days old: cleared');
  void id;
});

test('the key a piece of work held is overwritten when it is over', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.bin');
  await locker.lockWaiting(id);
  assert.deepEqual((await locker.scan(id)).map((e) => e.state), ['locked'], 'and the work after it still opens the folder');
  assert.equal((await exported(locker, id, 'a.bin')).length, 2000);
});

test('two folders saving the list at once do not collide', async () => {
  const { locker, id } = await make();
  const other = (await locker.addFolder(tmp('rl3-folder-'), { afterLock: 'delete', hidden: false, keepVersions: 3 })).id;
  await Promise.all(Array.from({ length: 40 }, (_, i) => locker.setChoices(i % 2 ? id : other, { keepVersions: 2 + (i % 3) })));
  assert.equal((await locker.scan(id)).length, 0);
});
