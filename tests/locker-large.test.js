// The Locker with files over its large-file threshold (LARGE, 50 MB): kept
// as pieces, versions of them put back, renamed, and found inside a plain
// folder. Moved here from locker-edges and locker-folders (12 Nov 2026): each
// seals more than 50 MB, about 30 s alone, and on this computer's local lane,
// beside three window lanes, they ran over 120 s. They do not touch Windows
// (they never open a locked file in another program, the one place the
// Locker calls attrib and icacls), so they run on the lab machine
// (tools/engine-hosts.json, portable).
import { test } from 'node:test';
import { tempDir } from './temp-dirs.js';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { FOLDER_KEYS } from '../apps/desktop/main/extras/locker/keys.js';
import { LARGE } from '../apps/desktop/main/extras/locker/folder.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
const random = (n) => crypto.randomBytes(n);
const tmp = (pre = 'rl-') => tempDir(pre);

async function make(choices = {}, hooks = null) {
  const profile = tmp();
  const folder = tmp('rl-folder-');
  const keep = tmp('rl-keep-');
  const locker = createLocker({ dir: profile, cost: CHEAP, hooks });
  const { recovery } = await locker.setup('pw');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: false, keepVersions: 3, ...choices });
  return { locker, folder, keep, id, recovery, profile };
}

const write = (folder, rel, bytes) => {
  const p = path.join(folder, ...rel.split('/'));
  fs.mkdirSync(path.dirname(p), { recursive: true });
  fs.writeFileSync(p, bytes);
  return p;
};

const entry = async (locker, id, rel) => (await locker.scan(id)).find((e) => e.rel === rel);

const lockedBytes = async (locker, id, rel) => {
  const out = path.join(tmp(), 'out.bin');
  await locker.exportCopy(id, rel, out);
  return fs.readFileSync(out);
};

test('an earlier version of a large file can be put back', { timeout: 120000 }, async () => {
  const { locker, folder, id } = await make({ keepVersions: 3 });
  const base = random(LARGE + 1024 * 1024);
  fs.writeFileSync(path.join(folder, 'film.mp4'), base);
  await locker.lockWaiting(id);
  const edited = Buffer.from(base);
  edited.fill(7, 0, 4096);
  const src = path.join(tmp(), 'film.mp4');
  fs.writeFileSync(src, edited);
  await locker.importFile(id, src, 'film.mp4');
  const versions = await locker.versions(id, 'film.mp4');
  assert.equal(versions.length, 2);
  await locker.restore(id, 'film.mp4', 1);
  const out = path.join(tmp(), 'out.mp4');
  await locker.exportCopy(id, 'film.mp4', out);
  assert.ok(fs.readFileSync(out).equals(base));
});

test('large files are kept as pieces, change only what changed, and a missing or damaged piece is found', { timeout: 120000 }, async () => {
  const { locker, folder, id } = await make({ keepVersions: 2 });
  // A file over the threshold, with a repeating start so the cut points land the same.
  const base = random(LARGE + 3 * 1024 * 1024);
  const p = write(folder, 'film.mp4', base);
  await locker.lockWaiting(id);
  const dir = `${p}.rlock.d`;
  assert.ok(fs.statSync(dir).isDirectory());
  const piecesBefore = fs.readdirSync(path.join(dir, 'p'));
  assert.ok(piecesBefore.length > 5);
  const e = await entry(locker, id, 'film.mp4');
  assert.equal(e.pieces, piecesBefore.length);

  const copy = Buffer.from(base);
  copy.writeUInt8(copy[10 * 1024 * 1024] ^ 1, 10 * 1024 * 1024);
  const editSrc = path.join(tmp(), 'edit.mp4');
  fs.writeFileSync(editSrc, copy);
  await locker.importFile(id, editSrc, 'film.mp4');
  const piecesAfter = fs.readdirSync(path.join(dir, 'p'));
  const added = piecesAfter.filter((n) => !piecesBefore.includes(n));
  assert.ok(added.length >= 1 && added.length <= 3, `only the piece around the change (${added.length} new of ${piecesAfter.length})`);

  const out = path.join(tmp(), 'out.mp4');
  await locker.exportCopy(id, 'film.mp4', out);
  assert.ok(fs.readFileSync(out).equals(copy));

  // Only two versions kept: an old-only piece goes once a third version exists; a shared piece stays.
  const v3 = Buffer.from(copy);
  v3.writeUInt8(v3[30 * 1024 * 1024] ^ 1, 30 * 1024 * 1024);
  const editSrc2 = path.join(tmp(), 'edit2.mp4');
  fs.writeFileSync(editSrc2, v3);
  await locker.importFile(id, editSrc2, 'film.mp4');
  const piecesAtEnd = fs.readdirSync(path.join(dir, 'p'));
  assert.ok(piecesAtEnd.length <= piecesAfter.length + 3, 'old pieces no longer used went away');
  assert.equal((await locker.versions(id, 'film.mp4')).length, 2);

  // A missing piece is said plainly; a damaged one too.
  // The piece must be one the newest version reads, so it is one that version added:
  // the first by name could belong only to the version before, and be rightly unneeded.
  const onlyNewest = piecesAtEnd.find((n) => !piecesAfter.includes(n));
  assert.ok(onlyNewest, 'the last edit added a piece');
  const victim = path.join(dir, 'p', onlyNewest);
  fs.renameSync(victim, `${victim}.moved`);
  await assert.rejects(() => locker.exportCopy(id, 'film.mp4', path.join(tmp(), 'gone.mp4')), (err) => err.code === 'truncated');
  fs.renameSync(`${victim}.moved`, victim);
  const bytes = fs.readFileSync(victim);
  bytes[40] ^= 1;
  fs.writeFileSync(victim, bytes);
  await assert.rejects(() => locker.exportCopy(id, 'film.mp4', path.join(tmp(), 'bad.mp4')), (err) => err.code === 'tampered');
});

test('rename: a large file kept in pieces takes the new name and keeps its contents', { timeout: 300000 }, async () => {
  const { locker, folder, id } = await make({ hidden: false });
  const body = random(LARGE + 3 * 1024 * 1024);
  fs.writeFileSync(path.join(folder, 'big.pdf'), body);
  await locker.lockWaiting(id);
  assert.ok(fs.existsSync(path.join(folder, 'big.pdf.rlock.d')));
  assert.equal((await locker.renameFile(id, 'big.pdf', 'bigger.pdf')).rel, 'bigger.pdf');
  assert.ok(fs.existsSync(path.join(folder, 'bigger.pdf.rlock.d')));
  assert.ok(!fs.existsSync(path.join(folder, 'big.pdf.rlock.d')));
  assert.ok((await lockedBytes(locker, id, 'bigger.pdf')).equals(body));
  assert.ok((await locker.scan(id)).some((e) => e.rel === 'bigger.pdf' && e.state === 'locked' && e.pieces));
  await locker.lock();
  await locker.unlock({ password: 'pw' });
  assert.ok((await lockedBytes(locker, id, 'bigger.pdf')).equals(body), 'still reads after locking again');
});

test('a plain folder with a locked folder inside it says so, whichever kind of locked thing is there', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.pdf', random(3000));
  fs.writeFileSync(path.join(folder, 'big.bin'), random(LARGE + 1024));
  await locker.lockWaiting(id);
  const look = async (build) => {
    const parent = tmp('rl-parent-');
    write(parent, 'notes/plain.txt', 'nothing locked');
    build(parent);
    const before = fs.readdirSync(parent, { recursive: true }).sort();
    const { lockedBelow } = await locker.lookAtFolder(parent);
    assert.deepEqual(fs.readdirSync(parent, { recursive: true }).sort(), before, 'looking changes nothing');
    return lockedBelow;
  };
  assert.equal(await look((p) => fs.cpSync(folder, path.join(p, 'work', 'kept'), { recursive: true })), true, 'a whole locked folder two levels down');
  assert.equal(await look((p) => fs.cpSync(path.join(folder, FOLDER_KEYS), path.join(p, 'only-keys', FOLDER_KEYS), { recursive: true })), true, 'its key file alone');
  assert.equal(await look((p) => fs.cpSync(path.join(folder, 'a.pdf.rlock'), path.join(p, 'one', 'a.pdf.rlock'))), true, 'one locked file');
  assert.equal(await look((p) => fs.cpSync(path.join(folder, 'big.bin.rlock.d'), path.join(p, 'one', 'big.bin.rlock.d'), { recursive: true })), true, 'one large locked file');
  assert.equal(await look(() => {}), false, 'and none at all');
  assert.equal(await look((p) => fs.cpSync(path.join(folder, 'a.pdf.rlock'), path.join(p, '.git', 'a.pdf.rlock'))), false, 'a folder of the system\'s own is not looked in');
  assert.equal(await look((p) => fs.cpSync(path.join(folder, 'a.pdf.rlock'), path.join(p, '1', '2', '3', '4', '5', '6', '7', '8', 'a.pdf.rlock'))), false, 'nor further down than a few levels');
});
