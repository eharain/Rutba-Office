// The Locker clears a folder with everything in it only through removeTree, so
// a junction a person put in one is unlinked where it is and what it points at
// is left alone. Electron 44's Node (24.21) follows a junction in a recursive
// fs.rmSync; its fs.promises.rm does not today, and removeTree does not depend
// on which. Run these under it (ELECTRON_RUN_AS_NODE) as well as under Node 22.

import { test } from 'node:test';
import { tempDir } from './temp-dirs.js';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { wipeWorking } from '../apps/desktop/main/extras/locker/working.js';
import { createMemoryFiles } from '../apps/desktop/main/memory-files.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const CHEAP = { N: 1024, r: 8, p: 1 };
const windowsOnly = { skip: process.platform === 'win32' ? false : 'junctions are for Windows' };
const random = (n) => crypto.randomBytes(n);

async function make(choices = {}) {
  const { large, ...folderChoices } = choices;
  const profile = tempDir('rl-rt-');
  const folder = tempDir('rl-rt-folder-');
  const locker = createLocker({ dir: profile, cost: CHEAP, large });
  await locker.setup('pw');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: false, keepVersions: 3, ...folderChoices });
  return { locker, folder, id, profile };
}

/** A folder outside the Locker, with a file in it, and whether that file is still there as it was. */
function bait() {
  const target = tempDir('rl-rt-bait-');
  fs.writeFileSync(path.join(target, 'keep.txt'), 'still here');
  const intact = () => fs.existsSync(path.join(target, 'keep.txt')) && fs.readFileSync(path.join(target, 'keep.txt'), 'utf8') === 'still here';
  return { target, intact };
}

/** A locked a.pdf with one earlier version, and the folder its versions are kept in. */
async function withVersion() {
  const m = await make();
  fs.writeFileSync(path.join(m.folder, 'a.pdf'), random(9000));
  await m.locker.lockWaiting(m.id);
  const next = path.join(tempDir('rl-rt-src-'), 'a.pdf');
  fs.writeFileSync(next, random(9000));
  await m.locker.importFile(m.id, next, 'a.pdf');
  const old = path.join(m.folder, '.rlocker-old');
  const kept = fs.readdirSync(old).map((n) => path.join(old, n)).filter((p) => fs.lstatSync(p).isDirectory());
  assert.equal(kept.length, 1, 'one file\'s versions folder');
  return { ...m, versions: kept[0] };
}

test('wiping a working folder unlinks a junction in it and leaves what it points at', windowsOnly, async () => {
  const { target, intact } = bait();
  const sub = path.join(tempDir('rl-rt-open-'), 'open', 'x');
  fs.mkdirSync(sub, { recursive: true });
  fs.writeFileSync(path.join(sub, 'copy.docx'), 'a working copy');
  fs.symlinkSync(target, path.join(sub, 'link'), 'junction');
  assert.equal(await wipeWorking(sub), true);
  assert.ok(!fs.existsSync(sub), 'the working folder has gone');
  assert.ok(intact(), 'what the junction pointed at is untouched');
});

test('deleting a locked file unlinks a junction in its versions folder and leaves what it points at', windowsOnly, async () => {
  const { target, intact } = bait();
  const m = await withVersion();
  fs.symlinkSync(target, path.join(m.versions, 'link'), 'junction');
  await m.locker.deleteFile(m.id, 'a.pdf', 'locked');
  assert.ok(!fs.existsSync(m.versions), 'its versions have gone');
  assert.ok(intact(), 'what the junction pointed at is untouched');
});

test('unlocking a file for good unlinks a junction in its versions folder and leaves what it points at', windowsOnly, async () => {
  const { target, intact } = bait();
  const m = await withVersion();
  fs.symlinkSync(target, path.join(m.versions, 'link'), 'junction');
  await m.locker.unlockForGood(m.id, 'a.pdf');
  assert.ok(fs.existsSync(path.join(m.folder, 'a.pdf')), 'the file is plain again');
  assert.ok(!fs.existsSync(m.versions), 'its versions have gone');
  assert.ok(intact(), 'what the junction pointed at is untouched');
});

test('taking a folder out of the Locker unlinks a junction in its old-copies folder and leaves what it points at', windowsOnly, async () => {
  const { target, intact } = bait();
  const m = await make();
  fs.writeFileSync(path.join(m.folder, 'a.pdf'), random(9000));
  await m.locker.lockWaiting(m.id);
  const old = path.join(m.folder, '.rlocker-old');
  fs.mkdirSync(old, { recursive: true });
  fs.symlinkSync(target, path.join(old, 'link'), 'junction');
  const out = await m.locker.removeFolder(m.id);
  assert.notEqual(out.removed, false);
  assert.ok(!fs.existsSync(old), 'the old-copies folder has gone');
  assert.ok(intact(), 'what the junction pointed at is untouched');
});

test('deleting a locked file kept in pieces unlinks a junction in its folder and leaves what it points at', windowsOnly, async () => {
  const { target, intact } = bait();
  const SMALL = 4 * 1024 * 1024;
  const m = await make({ large: SMALL });
  fs.writeFileSync(path.join(m.folder, 'big.pdf'), random(SMALL + 1024 * 1024));
  await m.locker.lockWaiting(m.id);
  const pieces = path.join(m.folder, 'big.pdf.rlock.d');
  assert.ok(fs.existsSync(pieces));
  fs.symlinkSync(target, path.join(pieces, 'link'), 'junction');
  await m.locker.deleteFile(m.id, 'big.pdf', 'locked');
  assert.ok(!fs.existsSync(pieces), 'its pieces have gone');
  assert.ok(intact(), 'what the junction pointed at is untouched');
});

test('markers swept from an earlier run: a junction in one, or one that is a junction, is unlinked and what it points at stays', windowsOnly, () => {
  const { target, intact } = bait();
  const root = path.join(tempDir('rl-rt-mem-'), 'mem');
  fs.mkdirSync(path.join(root, 'abcdef'), { recursive: true });
  fs.symlinkSync(target, path.join(root, 'abcdef', 'link'), 'junction');
  fs.symlinkSync(target, path.join(root, 'fedcba'), 'junction');
  const t = createMemoryFiles({ root });
  assert.equal(t.sweep(), 2);
  assert.deepEqual(fs.readdirSync(root), []);
  assert.ok(intact(), 'what the junctions pointed at is untouched');
});

test('no part of the Locker clears a folder with a recursive fs.rm: every one goes through removeTree', () => {
  const files = [
    ...fs.readdirSync(path.join(ROOT, 'apps/desktop/main/extras/locker')).filter((n) => n.endsWith('.js')).map((n) => `apps/desktop/main/extras/locker/${n}`),
    'apps/desktop/main/memory-files.js',
  ];
  const found = [];
  for (const rel of files) {
    const src = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    src.split(/\r?\n/).forEach((line, i) => {
      if (/\b(rm|rmSync|rmdir|rmdirSync)\s*\(/.test(line) && /recursive\s*:\s*true/.test(line)) found.push(`${rel}:${i + 1}`);
    });
  }
  assert.deepEqual(found, []);
  // The service's sweep of copies made for an email, which no test here builds.
  assert.match(fs.readFileSync(path.join(ROOT, 'apps/desktop/main/extras/locker/index.js'), 'utf8'), /removeTreeAsync\(at\)/);
});
