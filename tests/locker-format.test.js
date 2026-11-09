// The version written in every Locker format, and what a reader does with it.
// The samples in tests/fixtures/locker-format-1 were made once and are never
// regenerated: if one of these tests fails, a later change has broken files
// people already have. Fix the change, not the samples.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { createKeys, FOLDER_KEYS } from '../apps/desktop/main/extras/locker/keys.js';
import { sealPieces } from '../apps/desktop/main/extras/locker/pieces.js';
import { FORMAT_VERSION, MAGIC, parseHeader, newHeader, versionIn, checkVersion } from '../apps/desktop/main/extras/locker/format.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const SAMPLES = path.join(here, 'fixtures', 'locker-format-1');
const expected = JSON.parse(fs.readFileSync(path.join(SAMPLES, 'expected.json'), 'utf8'));
const CHEAP = { N: 1024, r: 8, p: 1 };
const tmp = (p = 'rl-format-') => fs.mkdtempSync(path.join(os.tmpdir(), p));
const NEWER = (e) => e.code === 'newer' && /newer Rutba Office/.test(e.message);

function copySamples() {
  const to = tmp();
  fs.cpSync(SAMPLES, to, { recursive: true });
  return { profile: path.join(to, 'profile'), folder: path.join(to, 'folder'), sent: path.join(to, 'sent note.txt.rlock') };
}

async function joined(folder, how = { password: expected.password }) {
  const locker = createLocker({ dir: tmp(), cost: CHEAP });
  await locker.setup('another person');
  const { id } = await locker.addFolder(folder, {}, how);
  return { locker, id };
}

const stamp = (file, text) => {
  const b = fs.readFileSync(file);
  b.write(text, 0, 'latin1');
  fs.writeFileSync(file, b);
};
const editJson = (file, change) => fs.writeFileSync(file, JSON.stringify(change(JSON.parse(fs.readFileSync(file, 'utf8')))));

test('this Rutba Office writes format 1, and the samples of format 1 are the real thing', () => {
  assert.equal(FORMAT_VERSION, 1);
  assert.equal(MAGIC.toString('latin1'), 'RLOCKV1\n');
  const { folder, sent, profile } = copySamples();
  assert.equal(fs.readFileSync(path.join(folder, 'note.txt.rlock')).subarray(0, 8).toString('latin1'), 'RLOCKV1\n');
  assert.equal(fs.readFileSync(sent).subarray(0, 8).toString('latin1'), 'RLOCKV1\n');
  assert.equal(JSON.parse(fs.readFileSync(path.join(folder, FOLDER_KEYS), 'utf8')).v, 1);
  assert.equal(JSON.parse(fs.readFileSync(path.join(profile, 'keyring.json'), 'utf8')).v, 1);
  const big = path.join(folder, 'big.bin.rlock.d');
  assert.equal(JSON.parse(fs.readFileSync(path.join(big, 'item.json'), 'utf8')).v, 1);
  const pieces = fs.readdirSync(path.join(big, 'p'));
  assert.ok(pieces.length >= 1);
  for (const p of pieces) assert.equal(fs.readFileSync(path.join(big, 'p', p)).subarray(0, 8).toString('latin1'), 'RPIECEV1');
});

test('format 1 samples open: the keyring with its password and recovery key', async () => {
  const { profile } = copySamples();
  const keys = createKeys({ dir: profile, cost: CHEAP });
  await assert.rejects(() => keys.unlock({ password: 'wrong' }), (e) => e.code === 'bad-password');
  await keys.unlock({ password: expected.password });
  assert.ok(keys.isOpen());
  keys.lock();
  await keys.unlock({ recovery: expected.recovery });
  assert.ok(keys.isOpen());
  keys.lock();
});

for (const [how, cred] of [['password', { password: expected.password }], ['recovery key', { recovery: expected.recovery }]]) {
  test(`format 1 samples open: a folder on another computer with its ${how}, small and large files alike`, async () => {
    const { folder } = copySamples();
    const { locker, id } = await joined(folder, cred);
    const list = await locker.scan(id);
    assert.deepEqual(list.map((e) => [e.rel, e.state]), [['big.bin', 'locked'], ['note.txt', 'locked']]);
    const out = tmp();
    await locker.exportCopy(id, 'note.txt', path.join(out, 'note.txt'));
    assert.equal(fs.readFileSync(path.join(out, 'note.txt'), 'utf8'), expected.note);
    await locker.exportCopy(id, 'big.bin', path.join(out, 'big.bin'));
    const big = fs.readFileSync(path.join(out, 'big.bin'));
    assert.equal(big.length, expected.big.size);
    assert.equal(crypto.createHash('sha256').update(big).digest('hex'), expected.big.sha256);
  });
}

test('format 1 samples open: a sent file with its own password', async () => {
  const { sent } = copySamples();
  const locker = createLocker({ dir: tmp(), cost: CHEAP });
  const out = path.join(tmp(), 'n.txt');
  await assert.rejects(() => locker.openShared(sent, out, 'wrong'), (e) => e.code === 'bad-password');
  await locker.openShared(sent, out, expected.sentPassword);
  assert.equal(fs.readFileSync(out, 'utf8'), expected.note);
});

test('what is written today carries its version: files, details, key files, lists and the keyring', async () => {
  const dir = tmp();
  const folder = tmp('rl-format-f-');
  fs.writeFileSync(path.join(folder, 'a.txt'), 'hello');
  const locker = createLocker({ dir, cost: CHEAP });
  await locker.setup('pw');
  const { id } = await locker.addFolder(folder, { hidden: false, afterLock: 'delete', keepVersions: 3 });
  await locker.lockWaiting(id);
  const file = path.join(folder, 'a.txt.rlock');
  assert.equal((await locker.headerOf(file)).format, 1);
  assert.equal(parseHeader(fs.readFileSync(file).subarray(0, 128)).version, 1);
  assert.equal(JSON.parse(fs.readFileSync(path.join(folder, FOLDER_KEYS), 'utf8')).v, 1);
  assert.equal(JSON.parse(fs.readFileSync(path.join(dir, 'keyring.json'), 'utf8')).v, 1);
  const sent = path.join(tmp(), 'a.rlock');
  await locker.exportShare(id, 'a.txt', sent, 'for you');
  assert.equal((await locker.headerOf(sent)).format, 1);
  assert.equal((await locker.headerOf(sent)).mode, 'shared');
});

test('a locked file from a newer version is named as that, not as damage, and nothing touches it', async () => {
  const { folder } = copySamples();
  const file = path.join(folder, 'note.txt.rlock');
  stamp(file, 'RLOCKV2\n');
  const before = fs.readFileSync(file);
  const { locker, id } = await joined(folder);
  const row = (await locker.scan(id)).find((e) => e.lockedRel === 'note.txt.rlock');
  assert.equal(row.newer, true);
  assert.match(row.message, /newer Rutba Office/);
  await locker.adoptCopies(id);
  await locker.lockWaiting(id);
  await locker.healFolderKeys();
  assert.ok(fs.readFileSync(file).equals(before), 'left exactly as it was');
});

test('a sent file from a newer version says so and writes nothing', async () => {
  const { sent } = copySamples();
  stamp(sent, 'RLOCKV2\n');
  const locker = createLocker({ dir: tmp(), cost: CHEAP });
  const out = path.join(tmp(), 'n.txt');
  await assert.rejects(() => locker.openShared(sent, out, expected.sentPassword), NEWER);
  assert.ok(!fs.existsSync(out));
});

test('a two-digit version is read as a number, and something that is not a version is not a locked file', () => {
  assert.equal(versionIn(Buffer.from('RLOCKV12'), 'RLOCKV'), 12);
  assert.equal(versionIn(Buffer.from('RLOCKV1\n'), 'RLOCKV'), 1);
  assert.equal(versionIn(Buffer.from('RPIECEV3'), 'RPIECEV'), 3);
  assert.equal(versionIn(Buffer.from('RLOCKVx\n'), 'RLOCKV'), null);
  assert.equal(versionIn(Buffer.from('RLOCKV0\n'), 'RLOCKV'), null);
  assert.equal(versionIn(Buffer.from('RLO'), 'RLOCKV'), null);
  const h = newHeader();
  h.write('RLOCKV9\n', 0, 'latin1');
  assert.throws(() => parseHeader(h), NEWER);
  h.write('NOTLOCK\n', 0, 'latin1');
  assert.throws(() => parseHeader(h), (e) => e.code === 'format');
  assert.throws(() => parseHeader(Buffer.alloc(10)), (e) => e.code === 'format');
});

test('a header this version cannot read is newer: another chunk size, another mode', () => {
  const chunk = newHeader();
  chunk.writeUInt32BE(1024, 8);
  assert.throws(() => parseHeader(chunk), NEWER);
  const mode = newHeader();
  mode.writeUInt8(7, 12);
  assert.throws(() => parseHeader(mode), NEWER);
});

test('a version number in a record: this one passes, a higher one is newer, none at all is refused', () => {
  assert.doesNotThrow(() => checkVersion(1));
  assert.throws(() => checkVersion(2), NEWER);
  for (const bad of [undefined, null, 0, '1', 1.5]) assert.throws(() => checkVersion(bad), (e) => e.code === 'format');
});

test('a keyring from a newer version is refused, by every way of using it, and left as it is', async () => {
  const { profile } = copySamples();
  const file = path.join(profile, 'keyring.json');
  editJson(file, (r) => ({ ...r, v: 2 }));
  const before = fs.readFileSync(file);
  const keys = createKeys({ dir: profile, cost: CHEAP });
  assert.equal(await keys.exists(), true, 'the window can still say a locker is there');
  await assert.rejects(() => keys.unlock({ password: expected.password }), NEWER);
  await assert.rejects(() => keys.changePassword(expected.password, 'next one'), NEWER);
  await assert.rejects(() => keys.replaceRecovery(expected.password), NEWER);
  await assert.rejects(() => keys.disableHello(), NEWER);
  assert.ok(fs.readFileSync(file).equals(before));
});

test('a folder key file from a newer version is not joined, not repaired and not written over', async () => {
  const { folder } = copySamples();
  const file = path.join(folder, FOLDER_KEYS);
  const fresh = await joined(folder);
  editJson(file, (r) => ({ ...r, v: 2 }));
  const before = fs.readFileSync(file);
  const healed = await fresh.locker.healFolderKeys();
  assert.equal(healed.failed.length, 1);
  assert.match(healed.failed[0].message, /newer Rutba Office/);
  assert.ok(fs.readFileSync(file).equals(before), 'a repair must never put an older version over a newer one');
  await assert.rejects(() => joined(folder), NEWER);
  assert.ok(fs.readFileSync(file).equals(before));
});

test('a large file from a newer version says so: its piece, its item and its list', async () => {
  const piece = copySamples();
  const big = path.join(piece.folder, 'big.bin.rlock.d');
  const one = fs.readdirSync(path.join(big, 'p'))[0];
  stamp(path.join(big, 'p', one), 'RPIECEV2');
  const a = await joined(piece.folder);
  await assert.rejects(() => a.locker.exportCopy(a.id, 'big.bin', path.join(tmp(), 'b.bin')), NEWER);

  const item = copySamples();
  editJson(path.join(item.folder, 'big.bin.rlock.d', 'item.json'), (r) => ({ ...r, v: 2 }));
  const b = await joined(item.folder);
  assert.equal((await b.locker.scan(b.id)).find((e) => e.lockedRel === 'big.bin.rlock.d').newer, true);

  // A list and details sealed as a newer version would seal them.
  const key = crypto.randomBytes(32);
  const dir = path.join(tmp(), 'x.rlock.d');
  const src = path.join(tmp(), 'src.bin');
  fs.writeFileSync(src, crypto.randomBytes(1000));
  await sealPieces({ source: src, dir, folderKey: key, folderId: crypto.randomBytes(16), stageDir: tmp(), meta: { name: 'x', v: 2 } });
  const { openPieces } = await import('../apps/desktop/main/extras/locker/pieces.js');
  await assert.rejects(() => openPieces({ dir, folderKey: key }), NEWER);
});

test('a file whose sealed details are from a newer version says so', async () => {
  const { sealFile } = await import('../apps/desktop/main/extras/locker/place.js');
  const key = crypto.randomBytes(32);
  const src = path.join(tmp(), 's.bin');
  fs.writeFileSync(src, 'x');
  const target = path.join(tmp(), 't.rlock');
  // Its own read-back is the first reader to meet the number.
  await assert.rejects(() => sealFile({ source: src, target, rootKey: key, headerBuf: newHeader(), meta: { name: 't', v: 2 }, stageDir: tmp() }), NEWER);
});

test('the format page names every version this copy writes', () => {
  const page = fs.readFileSync(path.resolve(here, '..', 'apps', 'desktop', 'main', 'extras', 'locker', 'FORMAT.md'), 'utf8');
  assert.ok(page.includes(`## Version ${FORMAT_VERSION}`));
  assert.ok(page.includes('RLOCKV1') && page.includes('RPIECEV1'));
});
