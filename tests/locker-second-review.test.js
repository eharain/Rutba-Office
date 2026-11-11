// Faults the second review of the Locker found, 10 Nov 2026, each reproduced on the code as it stood.
import { test } from 'node:test';
import { tempDir } from './temp-dirs.js';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { spawn } from 'node:child_process';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { FOLDER_KEYS } from '../apps/desktop/main/extras/locker/keys.js';
import { wipeFile } from '../apps/desktop/main/extras/locker/folder.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
// The size over which these tests' lockers keep a file as pieces: small, so the pieces are made in seconds; the
// 50 MB the app uses runs in locker-large.test.js, on the lab machine (12 Nov 2026).
const SMALL = 4 * 1024 * 1024;
const tmp = (pre = 'rl2-') => tempDir(pre);

async function make(choices = {}) {
  const { large, ...folderChoices } = choices;
  const dir = tmp();
  const folder = tmp('rl2-folder-');
  const locker = createLocker({ dir, cost: CHEAP, large });
  const { recovery } = await locker.setup('password one');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: false, keepVersions: 3, ...folderChoices });
  return { locker, folder, id, recovery, dir };
}
const write = (folder, name, size = 2000) => {
  const bytes = crypto.randomBytes(size);
  fs.writeFileSync(path.join(folder, name), bytes);
  return bytes;
};
const exported = async (locker, id, name) => {
  const out = path.join(tmp('rl2-out-'), name);
  await locker.exportCopy(id, name, out);
  return fs.readFileSync(out);
};

/** A file held by another program as C's fopen holds it: may be read and written, may not be deleted. */
function hold(file) {
  const script = `$f=[IO.File]::Open('${file.replace(/'/g, "''")}','Open','ReadWrite','ReadWrite');[Console]::Out.WriteLine('held');[Console]::In.ReadLine() | Out-Null;$f.Close()`;
  const child = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { stdio: ['pipe', 'pipe', 'inherit'] });
  return new Promise((resolve, reject) => {
    child.on('error', reject);
    child.stdout.once('data', () => resolve({
      release: () => new Promise((done) => { child.on('exit', done); child.stdin.end('\n'); }),
    }));
  });
}
const windowsOnly = { skip: process.platform !== 'win32' };

test('two computers saved a large file: deleting one version leaves the other, not an empty folder', async () => {
  const { locker, folder, id } = await make({ large: SMALL });
  const name = 'big.bin';
  const first = write(folder, name, SMALL + 2048);
  await locker.lockWaiting(id);
  const second = write(folder, name, SMALL + 4096);
  await locker.lockWaiting(id);
  const dir = path.join(folder, `${name}.rlock.d`);
  // What the second computer leaves: its list carries the same number as this one's.
  const lists = fs.readdirSync(dir).filter((f) => f.startsWith('list-')).sort();
  assert.equal(lists.length, 2);
  fs.renameSync(path.join(dir, lists[0]), path.join(dir, lists[1].replace(/-[0-9a-f]+\.rlist$/, '-aaaaaa.rlist')));

  const [row] = (await locker.scan(id)).filter((e) => e.state === 'conflict');
  assert.ok(row, 'shown as two versions');
  const shown = await exported(locker, id, name).catch(() => null);
  await locker.deleteFile(id, name, 'conflict', { lockedRel: row.lockedRel });

  assert.ok(fs.existsSync(dir), 'the folder is still there');
  const rows = await locker.scan(id);
  assert.deepEqual(rows.map((e) => e.state), ['locked'], 'one version, no longer in conflict');
  const left = await exported(locker, id, name);
  assert.ok(left.equals(first) || left.equals(second), 'and it opens whole');
  if (shown) assert.ok(!left.equals(shown), 'the one that stayed is the other one');
});

test('a file another program holds is not overwritten when it cannot be removed', windowsOnly, async () => {
  const dir = tmp();
  const file = path.join(dir, 'held.bin');
  const bytes = crypto.randomBytes(5000);
  fs.writeFileSync(file, bytes);
  const held = await hold(file);
  try {
    await wipeFile(file);
    assert.ok(fs.readFileSync(file).equals(bytes), 'every byte still there');
    assert.deepEqual(fs.readdirSync(dir), ['held.bin'], 'and nothing left beside it');
  } finally {
    await held.release();
  }
  await wipeFile(file);
  assert.equal(fs.existsSync(file), false, 'gone once it is let go of');
});

test('locking a file another program holds: its bytes stay, and the next pass does not lock zeros', windowsOnly, async () => {
  const { locker, folder, id } = await make();
  const bytes = write(folder, 'a.bin', 5000);
  const held = await hold(path.join(folder, 'a.bin'));
  try {
    await locker.lockWaiting(id);
    assert.ok(fs.readFileSync(path.join(folder, 'a.bin')).equals(bytes), 'the plain file is as it was');
  } finally {
    await held.release();
  }
  await locker.lockWaiting(id);
  assert.ok((await exported(locker, id, 'a.bin')).equals(bytes), 'what is locked is the real file');
});

test('wiping an ordinary file still overwrites and removes it, and leaves nothing behind', async () => {
  const dir = tmp();
  const file = path.join(dir, 'a.bin');
  fs.writeFileSync(file, crypto.randomBytes(3000));
  await wipeFile(file);
  assert.deepEqual(fs.readdirSync(dir), []);
});

test('removing a folder with an open copy whose changes cannot be saved: refused, and the copy kept', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.txt');
  await locker.lockWaiting(id);
  const { path: copy } = await locker.openPlain(id, 'a.txt');
  const edited = crypto.randomBytes(300);
  fs.writeFileSync(copy, edited);
  // The cloud makes a second version, so the save is refused.
  fs.copyFileSync(path.join(folder, 'a.txt.rlock'), path.join(folder, 'a (conflicted copy).txt.rlock'));

  const out = await locker.removeFolder(id);
  assert.equal(out.removed, false);
  assert.match(out.failed[0].message, /could not be saved/);
  assert.ok(fs.readFileSync(copy).equals(edited), 'the edits are still in the copy');
  assert.equal(locker.openCopyFor(copy), copy, 'and still known to be open');

  const shut = await locker.lock({});
  assert.deepEqual(shut.left, ['a.txt'], 'locking says it could not clear it');
  assert.ok(fs.readFileSync(copy).equals(edited));
  await locker.unlock({ password: 'password one' });
  assert.ok(fs.readFileSync(copy).equals(edited), 'and the next unlock does not wipe it');
});

test('unlocking for good with an open copy that cannot be saved is refused', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.txt');
  await locker.lockWaiting(id);
  const { path: copy } = await locker.openPlain(id, 'a.txt');
  fs.writeFileSync(copy, crypto.randomBytes(300));
  fs.copyFileSync(path.join(folder, 'a.txt.rlock'), path.join(folder, 'a (conflicted copy).txt.rlock'));
  await assert.rejects(locker.unlockForGood(id, 'a.txt'), /could not be saved/);
  assert.equal(locker.openCopyFor(copy), copy);
});

test('another folder\'s key file found here is left exactly as it is, and said', async () => {
  const { locker, folder, id } = await make();
  const file = path.join(folder, FOLDER_KEYS);
  const foreign = { v: 1, id: 'f'.repeat(32), unlockers: [] };
  fs.writeFileSync(file, JSON.stringify(foreign));
  const before = fs.readFileSync(file);
  const { failed } = await locker.healFolderKeys();
  assert.equal(failed.length, 1);
  assert.equal(failed[0].id, id);
  assert.match(failed[0].message, /different locked folder/);
  assert.ok(fs.readFileSync(file).equals(before));
});
