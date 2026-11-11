// The Locker where things go wrong: a cost a file should not be able to ask
// for, an earlier copy that cannot be reused, a file that changes or a copy
// that fails to read back while locking. In each, the plain file or the
// locked file a person already has must come through.

import { test } from 'node:test';
import { tempDir } from './temp-dirs.js';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { createKeys } from '../apps/desktop/main/extras/locker/keys.js';
import { sealFile } from '../apps/desktop/main/extras/locker/place.js';
import { sealPieces, openPieces, listsIn } from '../apps/desktop/main/extras/locker/pieces.js';
import { newHeader, openLocked } from '../apps/desktop/main/extras/locker/format.js';
import { createFolder } from '../apps/desktop/main/extras/locker/folder.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
const tmp = (p = 'rl-edge-') => tempDir(p);
const random = (n) => crypto.randomBytes(n);

async function make(choices = {}, hooks = null) {
  const folder = tmp('rl-edge-f-');
  const locker = createLocker({ dir: tmp(), cost: CHEAP, hooks });
  await locker.setup('pw');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: false, keepVersions: 3, ...choices });
  return { locker, folder, id };
}

async function read(file, key) {
  const o = await openLocked(file, key);
  const out = [];
  try {
    for await (const c of o.frames()) out.push(Buffer.from(c));
  } finally {
    await o.close();
  }
  return Buffer.concat(out);
}

test('a sent file cannot ask for a password cost that would stall the computer', async () => {
  const { locker, folder, id } = await make();
  fs.writeFileSync(path.join(folder, 'a.txt'), 'x');
  await locker.lockWaiting(id);
  const sent = path.join(tmp(), 'a.rlock');
  await locker.exportShare(id, 'a.txt', sent, 'for you');
  const good = fs.readFileSync(sent);
  for (const [at, value] of [[112, 2 ** 30], [112, 1000], [112, 1], [116, 1000], [116, 0], [120, 1000], [120, 0]]) {
    const bad = Buffer.from(good);
    bad.writeUInt32BE(value, at);
    fs.writeFileSync(sent, bad);
    const started = Date.now();
    await assert.rejects(() => locker.openShared(sent, path.join(tmp(), 'o'), 'for you'), (e) => e.code === 'format', `offset ${at} value ${value}`);
    assert.ok(Date.now() - started < 5000, 'refused before the work, not after');
  }
});

test('a file lock with an earlier copy that cannot be reused still comes out whole', async () => {
  const key = random(32);
  const dir = tmp();
  const srcA = path.join(dir, 'a.bin');
  fs.writeFileSync(srcA, random(200000));
  const first = path.join(dir, 'first.rlock');
  const headerBuf = newHeader();
  await sealFile({ source: srcA, target: first, rootKey: key, headerBuf, meta: { name: 'a' }, stageDir: dir });
  const cases = {
    missing: () => path.join(dir, 'never-there.rlock'),
    'another file\'s header': () => {
      const p = path.join(dir, 'other.rlock');
      const b = fs.readFileSync(first);
      b[20] ^= 1;
      fs.writeFileSync(p, b);
      return p;
    },
    'cut short after the header': () => {
      const p = path.join(dir, 'short.rlock');
      fs.writeFileSync(p, fs.readFileSync(first).subarray(0, 140));
      return p;
    },
    'damaged details': () => {
      const p = path.join(dir, 'details.rlock');
      const b = fs.readFileSync(first);
      b[300] ^= 1;
      fs.writeFileSync(p, b);
      return p;
    },
  };
  for (const [name, previous] of Object.entries(cases)) {
    const plain = random(150000);
    const src = path.join(dir, 'b.bin');
    fs.writeFileSync(src, plain);
    const target = path.join(dir, `${name.replace(/\W/g, '')}.rlock`);
    await sealFile({ source: src, target, rootKey: key, headerBuf, meta: { name: 'b' }, stageDir: dir, previous: previous() });
    assert.ok((await read(target, key)).equals(plain), name);
  }
});

test('a locked copy that fails its last read-back leaves the plain file where it was', async () => {
  let spoil = null;
  const { locker, folder, id } = await make({}, { at: async (step) => {
    if (step === 'before-delete' && spoil) {
      const f = spoil;
      spoil = null;
      const b = fs.readFileSync(f);
      b[b.length - 1] ^= 1;
      fs.writeFileSync(f, b);
    }
  } });
  const plain = path.join(folder, 'a.txt');
  fs.writeFileSync(plain, random(5000));
  const bytes = fs.readFileSync(plain);
  spoil = `${plain}.rlock`;
  const out = await locker.lockWaiting(id);
  assert.equal(out.failed.length, 1);
  assert.ok(fs.readFileSync(plain).equals(bytes), 'the plain file was not touched');
});

test('a file that changes while it is locked keeps the change: its plain copy stays, waiting', async () => {
  let plain = null;
  const { locker, folder, id } = await make({}, { at: async (step) => {
    if (step === 'swapped' && plain) {
      const f = plain;
      plain = null;
      fs.appendFileSync(f, 'typed while it was locking');
    }
  } });
  const file = path.join(folder, 'a.txt');
  fs.writeFileSync(file, 'first words ');
  plain = file;
  const out = await locker.lockWaiting(id);
  assert.equal(out.failed.length, 1);
  assert.match(fs.readFileSync(file, 'utf8'), /typed while it was locking/);
});

test('a folder\'s choices are checked: plain files kept need a place, and that place must be outside', async () => {
  const { locker, folder, id } = await make();
  await assert.rejects(() => locker.setChoices('nope', {}), (e) => e.code === 'state');
  await assert.rejects(() => locker.setChoices(id, { afterLock: 'keep', keepDir: null }), (e) => e.code === 'choice');
  await assert.rejects(() => locker.setChoices(id, { afterLock: 'keep', keepDir: path.join(folder, 'inside') }), (e) => e.code === 'choice');
  const outside = tmp();
  const after = await locker.setChoices(id, { afterLock: 'keep', keepDir: outside, keepVersions: 99, wanted: 'ignored' });
  assert.equal(after.afterLock, 'keep');
  assert.equal(after.keepDir, outside);
  assert.equal(after.keepVersions, 3, 'a value out of range is not taken');
  assert.equal(after.wanted, undefined);
});

test('a kept plain file is checked against what was locked before it moves, and lands where the person chose', async () => {
  const keep = tmp();
  const { locker, folder, id } = await make({ afterLock: 'keep', keepDir: keep });
  const bytes = random(4000);
  fs.writeFileSync(path.join(folder, 'a.txt'), bytes);
  await locker.lockWaiting(id);
  assert.ok(fs.existsSync(path.join(folder, 'a.txt.rlock')));
  assert.ok(!fs.existsSync(path.join(folder, 'a.txt')));
  const moved = fs.readdirSync(keep, { recursive: true }).map(String).find((n) => n.endsWith('a.txt'));
  assert.ok(moved);
  assert.ok(fs.readFileSync(path.join(keep, moved)).equals(bytes));
});

test('rules: written and read back whole, and they decide what waits', async () => {
  const { locker, folder, id } = await make();
  assert.equal(await locker.readRules(id), '');
  await locker.writeRules(id, '*.pdf\n');
  assert.equal(await locker.readRules(id), '*.pdf\n');
  fs.writeFileSync(path.join(folder, 'a.pdf'), 'p');
  fs.writeFileSync(path.join(folder, 'b.txt'), 't');
  const list = await locker.scan(id);
  assert.equal(list.find((e) => e.rel === 'a.pdf').state, 'waiting');
  assert.equal(list.find((e) => e.rel === 'b.txt').state, 'excluded');
  assert.equal((await locker.why(id, 'missing.pdf')).state, 'waiting');
  assert.equal((await locker.why(id, 'missing.txt')).state, 'excluded');
});

test('locking while keeping the temporary copies leaves them, and says how many', async () => {
  const { locker, folder, id } = await make();
  fs.writeFileSync(path.join(folder, 'a.txt'), 'words');
  await locker.lockWaiting(id);
  const { path: copy } = await locker.openPlain(id, 'a.txt');
  assert.ok(fs.existsSync(copy));
  const out = await locker.lock({ keepCopies: true });
  assert.equal(out.kept, 1);
  assert.ok(fs.existsSync(copy));
  await locker.unlock({ password: 'pw' });
  await locker.lock();
  assert.ok(!fs.existsSync(copy), 'gone at the next ordinary lock');
});

async function openDoc() {
  const made = await make();
  fs.writeFileSync(path.join(made.folder, 'a.txt'), 'first words');
  await made.locker.lockWaiting(made.id);
  const { path: copy } = await made.locker.openPlain(made.id, 'a.txt');
  return { ...made, copy };
}

async function lockedText({ locker, id }) {
  const out = path.join(tmp(), 'out.txt');
  await locker.exportCopy(id, 'a.txt', out);
  return fs.readFileSync(out, 'utf8');
}

test('locking under a document open in Office removes its file and keeps its folder; a Save made then is locked in at the next unlock', async () => {
  const d = await openDoc();
  const shown = (f) => f === d.copy;
  const out = await d.locker.lock({ shown });
  assert.deepEqual(out.open, ['a.txt']);
  assert.ok(!fs.existsSync(d.copy), 'nothing unlocked stays on the disk');
  assert.ok(fs.existsSync(path.dirname(d.copy)), 'the window\'s Save needs somewhere to go');
  fs.writeFileSync(d.copy, 'saved while locked');
  assert.deepEqual(await d.locker.reapClosed({ shown }), [], 'still open, nothing cleared');
  await d.locker.unlock({ password: 'pw' }, { shown });
  assert.equal(await lockedText(d), 'saved while locked');
  assert.ok(fs.existsSync(d.copy), 'unlocking does not take the file from under the window either');
  assert.ok(d.locker.isOpenCopy(d.copy));
  fs.writeFileSync(d.copy, 'saved after unlocking');
  await d.locker.saveOpen(d.copy);
  assert.equal(await lockedText(d), 'saved after unlocking');
  assert.deepEqual(await d.locker.reapClosed({ shown: () => false }), ['a.txt']);
  assert.ok(!fs.existsSync(d.copy));
});

test('a document closed while locked with a Save in it waits for the keys, then goes in and its copy goes', async () => {
  const d = await openDoc();
  await d.locker.lock({ shown: (f) => f === d.copy });
  fs.writeFileSync(d.copy, 'saved while locked');
  assert.deepEqual(await d.locker.reapClosed({ shown: () => false }), []);
  assert.ok(fs.existsSync(d.copy), 'the only place that work is');
  await d.locker.unlock({ password: 'pw' });
  assert.equal(await lockedText(d), 'saved while locked');
  assert.ok(!fs.existsSync(d.copy));
});

test('a document closed while locked with nothing saved has its copy cleared without the keys', async () => {
  const d = await openDoc();
  await d.locker.lock({ shown: (f) => f === d.copy });
  assert.deepEqual(await d.locker.reapClosed({ shown: () => false }), ['a.txt']);
  assert.ok(!fs.existsSync(path.dirname(d.copy)), 'its folder goes with it');
});

test('unlocking with nothing saved while locked puts the copy back for the window still showing it', async () => {
  const d = await openDoc();
  const shown = (f) => f === d.copy;
  await d.locker.lock({ shown });
  await d.locker.unlock({ password: 'pw' }, { shown });
  assert.equal(fs.readFileSync(d.copy, 'utf8'), 'first words');
  assert.deepEqual((await d.locker.saveOpen(d.copy)).changed, false);
  assert.equal(await lockedText(d), 'first words');
});

test('a lock with no Office window showing the file clears the copy as before', async () => {
  const d = await openDoc();
  const out = await d.locker.lock({ shown: () => false });
  assert.deepEqual(out.open, []);
  assert.ok(!fs.existsSync(d.copy));
});

test('quitting with a document open: at the next start an unchanged copy is cleared, and a changed one waits for the keys', async () => {
  const unchanged = await openDoc();
  await unchanged.locker.lock({ shown: (f) => f === unchanged.copy });
  const dirA = path.dirname(unchanged.locker.stateFile);
  const nextA = createLocker({ dir: dirA, cost: CHEAP });
  await nextA.loadOpenRecord();
  assert.ok(nextA.hasOpenCopies());
  assert.deepEqual(await nextA.reapClosed(), ['a.txt']);
  assert.ok(!fs.existsSync(unchanged.copy));

  const changed = await openDoc();
  await changed.locker.lock({ shown: (f) => f === changed.copy });
  fs.writeFileSync(changed.copy, 'saved as Office was closing');
  const dirB = path.dirname(changed.locker.stateFile);
  const nextB = createLocker({ dir: dirB, cost: CHEAP });
  await nextB.loadOpenRecord();
  assert.deepEqual(await nextB.reapClosed(), []);
  assert.ok(fs.existsSync(changed.copy));
  await nextB.unlock({ password: 'pw' });
  assert.equal(await lockedText({ locker: nextB, id: changed.id }), 'saved as Office was closing');
  assert.ok(!fs.existsSync(changed.copy));
});

test('keys: nothing to open it with is refused, an unknown folder is forgotten quietly, a folder file with no password route opens nothing', async () => {
  const k = createKeys({ dir: tmp(), cost: CHEAP });
  await k.setup('pw');
  await assert.rejects(() => k.unlock({}), (e) => e.code === 'bad-password');
  assert.doesNotThrow(() => k.forgetFolderKey('not-a-folder'));
  const fk = random(32);
  const file = k.wrapForFolder(fk, 'f');
  assert.equal(await k.openFolderKeys({ ...file, unlockers: [] }, { password: 'pw' }), null);
  assert.equal(await k.openFolderKeys({ ...file, unlockers: [] }, { recovery: 'AAAA' }), null);
  assert.equal(await k.openFolderKeys({ ...file, unlockers: [] }), null);
  k.lock();
  assert.equal(await k.openFolderKeys(file), null, 'nothing cached once locked');
});

test('a large file: a damaged piece is replaced when the same bytes are locked again', async () => {
  const dir = tmp();
  const key = random(32);
  const item = path.join(dir, 'x.rlock.d');
  const src = path.join(dir, 's.bin');
  const bytes = random(4 * 1024 * 1024);
  fs.writeFileSync(src, bytes);
  const args = { source: src, dir: item, folderKey: key, folderId: random(16), stageDir: tmp(), meta: { name: 'x' } };
  const first = await sealPieces(args);
  assert.ok(first.pieces >= 1);
  const again = await sealPieces(args);
  assert.equal(again.written, 0, 'good pieces are not written twice');
  const names = fs.readdirSync(path.join(item, 'p'));
  const victim = path.join(item, 'p', names[0]);
  const b = fs.readFileSync(victim);
  b[b.length - 1] ^= 1;
  fs.writeFileSync(victim, b);
  const repaired = await sealPieces(args);
  assert.equal(repaired.written, 1);
  const o = await openPieces({ dir: item, folderKey: key });
  const out = [];
  for await (const c of o.frames()) out.push(Buffer.from(c));
  assert.ok(Buffer.concat(out).equals(bytes));
});

test('a large file: a list that will not read stops the clean-up, so no piece is removed from under it', async () => {
  const dir = tmp();
  const key = random(32);
  const item = path.join(dir, 'x.rlock.d');
  const src = path.join(dir, 's.bin');
  const stageDir = tmp();
  const args = { source: src, dir: item, folderKey: key, folderId: random(16), stageDir, meta: { name: 'x' }, keep: 2 };
  fs.writeFileSync(src, random(1024 * 1024));
  await sealPieces(args);
  fs.writeFileSync(src, random(1024 * 1024));
  await sealPieces(args);
  const before = fs.readdirSync(path.join(item, 'p')).length;
  const [, older] = await listsIn(item);
  const f = path.join(item, older.file);
  const b = fs.readFileSync(f);
  b[b.length - 1] ^= 1;
  fs.writeFileSync(f, b);
  fs.writeFileSync(src, random(1024 * 1024));
  await sealPieces(args);
  assert.ok(fs.readdirSync(path.join(item, 'p')).length > before, 'the new pieces were added and none taken away');
});

test('a locked file that records a name leading out of its folder is shown as damaged, and nothing is written outside', async () => {
  const parent = tmp('rl-edge-p-');
  const root = path.join(parent, 'folder');
  fs.mkdirSync(root);
  const key = random(32);
  const folderId = random(16);
  const src = path.join(tmp(), 's.bin');
  const bytes = random(3000);
  fs.writeFileSync(src, bytes);
  const bad = ['../evil.pdf', '..\\evil.pdf', 'sub/evil.pdf', 'C:evil.pdf', '..', '.', 'nul\u0000.pdf'];
  for (const [n, name] of bad.entries()) {
    await sealFile({ source: src, target: path.join(root, `f${n}.pdf.rlock`), rootKey: key, headerBuf: newHeader({ folderId }), meta: { name }, stageDir: tmp() });
  }
  await sealPieces({ source: src, dir: path.join(root, 'big.bin.rlock.d'), folderKey: key, folderId, stageDir: tmp(), meta: { name: '../evil.bin' } });
  // And one good one, so the folder is read at all.
  await sealFile({ source: src, target: path.join(root, 'good.pdf.rlock'), rootKey: key, headerBuf: newHeader({ folderId }), meta: { name: 'good.pdf' }, stageDir: tmp() });
  const folder = createFolder({ root, key, folderId, stageDir: tmp(), choices: { hidden: false } });
  const list = await folder.scan();
  assert.deepEqual(list.filter((e) => e.state !== 'damaged').map((e) => e.rel), ['good.pdf']);
  assert.equal(list.filter((e) => e.state === 'damaged').length, bad.length + 1);
  const out = path.join(tmp(), 'good.pdf');
  await folder.exportCopy('good.pdf', out);
  assert.ok(fs.readFileSync(out).equals(bytes));
  assert.deepEqual(fs.readdirSync(parent), ['folder'], 'nothing beside the folder');
});
