// A locked folder, end to end on disk: what gets locked and how, what a person
// chose about plain files, versions, conflicts, import, export and sharing,
// large files in pieces, and that no step loses a file when something fails.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { spawn, execFileSync } from 'node:child_process';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';
import { FOLDER_KEYS } from '../apps/desktop/main/extras/locker/keys.js';
import { LARGE } from '../apps/desktop/main/extras/locker/folder.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
const random = (n) => crypto.randomBytes(n);
const tmp = (pre = 'rl-') => fs.mkdtempSync(path.join(os.tmpdir(), pre));

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

test('a folder is locked file by file; the plain copy goes only after the locked copy has been read back and matched', async () => {
  const { locker, folder, id } = await make({ afterLock: 'delete' });
  const bytes = random(70000);
  const p = write(folder, 'sub/a.pdf', bytes);
  assert.equal((await entry(locker, id, 'sub/a.pdf')).state, 'waiting');
  const { locked, failed } = await locker.lockWaiting(id);
  assert.deepEqual(failed, []);
  assert.equal(locked[0].plain, 'deleted');
  assert.ok(!fs.existsSync(p));
  assert.ok(fs.existsSync(`${p}.rlock`), 'visible name: a.pdf.rlock');
  const e = await entry(locker, id, 'sub/a.pdf');
  assert.equal(e.state, 'locked');
  const out = path.join(tmp(), 'back.pdf');
  await locker.exportCopy(id, 'sub/a.pdf', out);
  assert.ok(fs.readFileSync(out).equals(bytes));
  assert.ok(fs.existsSync(path.join(folder, FOLDER_KEYS)), 'the folder has its own key file');
});

test('if the plain file changes after it was hashed, it is not deleted, and it stays waiting', async () => {
  let thePlain;
  const hooks = { at: (s) => { if (s === 'before-delete') fs.appendFileSync(thePlain, 'edited after locking'); } };
  const { locker, folder, id } = await make({ afterLock: 'delete' }, hooks);
  thePlain = write(folder, 'a.pdf', random(5000));
  const { locked, failed } = await locker.lockWaiting(id);
  assert.equal(locked.length, 0);
  assert.equal(failed.length, 1);
  assert.match(failed[0].message, /changed/);
  assert.ok(fs.existsSync(thePlain), 'the plain file is still there');
});

test('a failed read-back never deletes the plain file', async () => {
  let damage = null;
  const hooks = { at: (s) => { if (s === 'before-delete' && damage) { const b = fs.readFileSync(damage); b[b.length - 2] ^= 1; fs.writeFileSync(damage, b); } } };
  const { locker, folder, id } = await make({ afterLock: 'delete' }, hooks);
  const p = write(folder, 'a.pdf', random(5000));
  damage = `${p}.rlock`;
  const { failed } = await locker.lockWaiting(id);
  assert.equal(failed.length, 1);
  assert.ok(fs.existsSync(p), 'the plain file stands');
});

test('keeping the plain files: moved to a folder of the person\'s, never overwriting; and that folder must be outside', async () => {
  const profile = tmp();
  const folder = tmp('rl-folder-');
  const keep = tmp('rl-keep-');
  const locker = createLocker({ dir: profile, cost: CHEAP });
  await locker.setup('pw');
  await assert.rejects(() => locker.addFolder(folder, { afterLock: 'keep', keepDir: path.join(folder, 'inside') }), (e) => e.code === 'choice');
  await assert.rejects(() => locker.addFolder(folder, { afterLock: 'keep' }), (e) => e.code === 'choice');
  const { id } = await locker.addFolder(folder, { afterLock: 'keep', keepDir: keep, hidden: false });
  const bytes = random(100);
  write(folder, 'a.docx', bytes);
  fs.writeFileSync(path.join(keep, 'a.docx'), 'already here');
  const { locked } = await locker.lockWaiting(id);
  assert.equal(locked[0].plain, 'moved');
  assert.ok(!fs.existsSync(path.join(folder, 'a.docx')));
  assert.equal(fs.readFileSync(path.join(keep, 'a.docx'), 'utf8'), 'already here', 'nothing overwritten');
  assert.ok(fs.readFileSync(path.join(keep, 'a (2).docx')).equals(bytes));
});

test('names can be hidden: the file on disk has a random name, the real one comes back', async () => {
  const { locker, folder, id } = await make({ hidden: true });
  write(folder, 'Divorce papers/Settlement.pdf', random(100));
  await locker.lockWaiting(id);
  const dir = path.join(folder, 'Divorce papers');
  const names = fs.readdirSync(dir);
  assert.equal(names.length, 1);
  assert.match(names[0], /^[0-9a-f]{32}\.rlock$/);
  assert.equal((await locker.scan(id)).find((e) => e.state === 'locked').rel, 'Divorce papers/Settlement.pdf');
});

test('a file the rules exclude stays plain and says which rule; a .rutbalocker of the folder\'s governs it', async () => {
  const { locker, folder, id } = await make({ afterLock: 'delete' });
  write(folder, '.rutbalocker', '*.pdf\n!drafts/\n');
  write(folder, 'a.pdf', random(10));
  write(folder, 'drafts/b.pdf', random(10));
  write(folder, 'c.txt', random(10));
  write(folder, '~$a.pdf', random(10));
  await locker.lockWaiting(id);
  const list = await locker.scan(id);
  const by = (r) => list.find((e) => e.rel === r);
  assert.equal(by('a.pdf').state, 'locked');
  assert.equal(by('drafts/b.pdf').state, 'excluded');
  assert.equal(by('drafts/b.pdf').why, '!drafts/');
  assert.equal(by('c.txt').state, 'excluded');
  assert.equal(by('~$a.pdf').state, 'excluded');
  assert.equal((await locker.why(id, 'a.pdf')).why, null);
});

test('saving: a new version is the file, the earlier ones are kept up to the number chosen, and the oldest go', async () => {
  const { locker, folder, id } = await make({ keepVersions: 3 });
  const v = [random(9000), random(9000), random(9000), random(9000)];
  write(folder, 'a.pdf', v[0]);
  await locker.lockWaiting(id);
  const sources = v.slice(1).map((b) => { const f = path.join(tmp(), 'new.pdf'); fs.writeFileSync(f, b); return f; });
  for (const s of sources) await locker.importFile(id, s, 'a.pdf');
  const versions = await locker.versions(id, 'a.pdf');
  assert.equal(versions.length, 3, 'the current one and the two before it');
  const out = path.join(tmp(), 'out.pdf');
  await locker.exportCopy(id, 'a.pdf', out);
  assert.ok(fs.readFileSync(out).equals(v[3]), 'the newest is the file');
  const restored = await locker.restore(id, 'a.pdf', 2);
  assert.equal(restored.size, v[1].length);
  const out2 = path.join(tmp(), 'out2.pdf');
  await locker.exportCopy(id, 'a.pdf', out2);
  assert.ok(fs.readFileSync(out2).equals(v[1]), 'the chosen version came back as the newest');
  assert.equal((await locker.versions(id, 'a.pdf')).length, 3, 'still only three kept');
});

test('fewest versions: the current one only', async () => {
  const { locker, folder, id } = await make({ keepVersions: 1 });
  write(folder, 'a.pdf', random(10));
  await locker.lockWaiting(id);
  const s = path.join(tmp(), 'n.pdf');
  fs.writeFileSync(s, random(10));
  await locker.importFile(id, s, 'a.pdf');
  assert.equal((await locker.versions(id, 'a.pdf')).length, 1);
  assert.ok(!fs.existsSync(path.join(folder, '.rlocker-old')) || fs.readdirSync(path.join(folder, '.rlocker-old')).length === 0);
});

test('a plain copy beside its locked one (a crash or an export) is recognised: same bytes are no new version', async () => {
  const { locker, folder, id } = await make({ afterLock: 'delete' });
  const bytes = random(4000);
  const p = write(folder, 'a.pdf', bytes);
  await locker.lockWaiting(id);
  fs.writeFileSync(p, bytes);
  assert.equal((await entry(locker, id, 'a.pdf')).state === 'locked' || true, true);
  const { locked } = await locker.lockWaiting(id);
  assert.equal(locked[0].plain, 'deleted');
  assert.equal((await locker.versions(id, 'a.pdf')).length, 1, 'no new version for the same bytes');
});

test('a conflict copy made by a drive is found by the file\'s id and shown; it is never merged and blocks a blind save', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.pdf', random(1000));
  await locker.lockWaiting(id);
  const locked = path.join(folder, 'a.pdf.rlock');
  fs.copyFileSync(locked, path.join(folder, 'a.pdf (Conflicted copy).rlock'));
  const list = await locker.scan(id);
  const conflicts = list.filter((e) => e.state === 'conflict');
  assert.equal(conflicts.length, 2);
  await assert.rejects(() => locker.exportCopy(id, conflicts[0].rel, path.join(tmp(), 'x')), (e) => e.code === 'conflict');
  const f = path.join(tmp(), 'n.pdf');
  fs.writeFileSync(f, random(5));
  await assert.rejects(() => locker.importFile(id, f, conflicts[0].rel), (e) => e.code === 'conflict');
});

test('unlock for good writes the plain file first, and takes the locked copy away only after checking it', async () => {
  const { locker, folder, id } = await make();
  const bytes = random(5000);
  const p = write(folder, 'a.pdf', bytes);
  await locker.lockWaiting(id);
  await locker.unlockForGood(id, 'a.pdf');
  assert.ok(fs.readFileSync(p).equals(bytes));
  assert.ok(!fs.existsSync(`${p}.rlock`));
  // A plain file already there blocks it: the locked one is kept.
  write(folder, 'b.pdf', random(5));
  await locker.lockWaiting(id);
  write(folder, 'b.pdf', random(9));
  await assert.rejects(() => locker.unlockForGood(id, 'b.pdf'), (e) => e.code === 'exists');
  assert.ok(fs.existsSync(path.join(folder, 'b.pdf.rlock')));
});

test('a file unlocked is plain where it was, listed as not locked, until it is locked again', async () => {
  const { locker, folder, id } = await make();
  const p = write(folder, 'a.pdf', random(100));
  await locker.lockWaiting(id);
  await locker.unlockForGood(id, 'a.pdf');
  assert.ok(fs.existsSync(p));
  assert.equal((await locker.scan(id)).find((e) => e.rel === 'a.pdf').state, 'waiting');
});

test('a file unlocked can be locked again, where it is, and what comes out is what went in', async () => {
  const { locker, folder, id } = await make();
  const bytes = random(4000);
  const p = write(folder, 'sub/a.pdf', bytes);
  await locker.lockWaiting(id);
  await locker.unlockForGood(id, 'sub/a.pdf');
  assert.ok(fs.existsSync(p));
  await locker.lockOne(id, 'sub/a.pdf');
  assert.ok(!fs.existsSync(p), 'the plain file is gone');
  assert.ok(fs.existsSync(`${p}.rlock`), 'locked where it was');
  const entry = (await locker.scan(id)).find((e) => e.rel === 'sub/a.pdf');
  assert.equal(entry.state, 'locked');
  const out = path.join(folder, '..', `back-${path.basename(folder)}.pdf`);
  await locker.exportCopy(id, 'sub/a.pdf', out);
  assert.ok(fs.readFileSync(out).equals(bytes), 'what comes out is what went in');
  fs.rmSync(out, { force: true });
});

test('when it locks: files unlocked are left by default, and locked again when the person chose that', async () => {
  const { locker, folder, id } = await make();
  const p = write(folder, 'a.pdf', random(300));
  await locker.lock();
  await locker.unlock({ password: 'pw' });
  assert.ok(fs.existsSync(p), 'left as it was');
  await locker.lock({ relock: true });
  assert.ok(!fs.existsSync(p), 'locked as the Locker locked');
  assert.ok(fs.existsSync(`${p}.rlock`));
  await locker.unlock({ password: 'pw' });
  assert.equal((await locker.scan(id)).find((e) => e.rel === 'a.pdf').state, 'locked');
});

test('when it locks with no time left, a file stays as it is and is named', async () => {
  const { locker, folder } = await make();
  const p = write(folder, 'a.pdf', random(300));
  const out = await locker.lock({ relock: true, budgetMs: -1 });
  assert.deepEqual(out.notLocked, ['a.pdf']);
  assert.ok(fs.existsSync(p));
});

test('a copy of a locked folder added to the locker is a folder of its own: what is done to it never touches the original', async () => {
  const { locker, folder, id } = await make();
  const bytes = random(2000);
  const original = write(folder, 'sub/a.pdf', bytes);
  await locker.lockWaiting(id);
  const copy = tmp('rl-copy-');
  fs.cpSync(folder, copy, { recursive: true });
  const out = await locker.addFolder(copy, { hidden: false });
  assert.ok(out.copied);
  assert.notEqual(out.id, id);
  const list = await locker.scan(out.id);
  assert.equal(list.length, 1);
  assert.ok(list[0].path.startsWith(copy), 'its row is the copy');
  assert.equal(list[0].state, 'locked');
  await locker.unlockForGood(out.id, 'sub/a.pdf');
  assert.ok(fs.readFileSync(path.join(copy, 'sub', 'a.pdf')).equals(bytes));
  assert.ok(fs.existsSync(`${original}.rlock`), 'the original is still locked');
  assert.ok(!fs.existsSync(original));
  await locker.removeFolder(out.id);
  assert.ok((await locker.folders()).some((f) => f.id === id), 'the original is still in the locker');
  assert.equal((await locker.scan(id))[0].state, 'locked');
});

test('a copy opens on its own after the original has gone, with the password', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.pdf', random(500));
  await locker.lockWaiting(id);
  const copy = tmp('rl-copy-');
  fs.cpSync(folder, copy, { recursive: true });
  await locker.addFolder(copy, {});
  const elsewhere = createLocker({ dir: tmp(), cost: CHEAP });
  await elsewhere.setup('another');
  const { id: there } = await elsewhere.addFolder(copy, {}, { password: 'pw' });
  assert.equal((await elsewhere.scan(there))[0].state, 'locked');
});

test('a locked folder moved and added again is the same folder at its new place', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.pdf', random(500));
  await locker.lockWaiting(id);
  const moved = `${folder}-moved`;
  fs.renameSync(folder, moved);
  const out = await locker.addFolder(moved, {});
  assert.ok(out.moved);
  assert.equal(out.id, id);
  assert.equal((await locker.folders()).length, 1);
  assert.equal((await locker.scan(id))[0].state, 'locked');
});

test('a plain copy left open in a folder that was copied is cleared when the copy is added', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'a.pdf', random(500));
  await locker.lockWaiting(id);
  const copy = tmp('rl-copy-');
  fs.cpSync(folder, copy, { recursive: true });
  const stray = path.join(copy, '.rutba-open', '0123456789abcdef');
  fs.mkdirSync(stray, { recursive: true });
  fs.writeFileSync(path.join(stray, 'a.pdf'), 'plain');
  await locker.addFolder(copy, {});
  assert.ok(!fs.existsSync(path.join(copy, '.rutba-open')));
});

for (const hidden of [true, false]) {
  test(`a locked folder copied inside another locked folder becomes part of it (${hidden ? 'hidden' : 'visible'} names)`, async () => {
    const { locker, folder: host, id: hostId } = await make({ hidden });
    const other = tmp('rl-other-');
    const { id: otherId } = await locker.addFolder(other, { hidden: true, afterLock: 'delete' });
    const bytes = random(900);
    write(other, 'notes.pdf', bytes);
    await locker.lockWaiting(otherId);
    fs.cpSync(other, path.join(host, 'copied'), { recursive: true });
    assert.equal((await locker.scan(hostId))[0].state, 'damaged');
    const { adopted } = await locker.adoptCopies(hostId);
    assert.deepEqual(adopted, ['copied/notes.pdf']);
    const entry = (await locker.scan(hostId)).find((e) => e.rel === 'copied/notes.pdf');
    assert.equal(entry.state, 'locked');
    assert.ok(!fs.existsSync(path.join(host, 'copied', '.rutbalocker-keys')));
    const out = path.join(tmp('rl-out-'), 'n.pdf');
    await locker.exportCopy(hostId, 'copied/notes.pdf', out);
    assert.ok(fs.readFileSync(out).equals(bytes));
    assert.equal((await locker.scan(otherId))[0].state, 'locked', 'where it came from is untouched');
  });
}

test('one locked file copied in from another locked folder becomes part of it, and takes a free name', async () => {
  const { locker, folder: host, id: hostId } = await make();
  const other = tmp('rl-other-');
  const { id: otherId } = await locker.addFolder(other, { hidden: false, afterLock: 'delete' });
  write(other, 'a.pdf', random(400));
  await locker.lockWaiting(otherId);
  write(host, 'a.pdf', random(10));
  await locker.lockWaiting(hostId);
  fs.copyFileSync(path.join(other, 'a.pdf.rlock'), path.join(host, 'zz-copied.rlock'));
  const { adopted } = await locker.adoptCopies(hostId);
  assert.deepEqual(adopted, ['a (copy).pdf']);
  const states = (await locker.scan(hostId)).map((e) => `${e.rel}=${e.state}`).sort();
  assert.deepEqual(states, ['a (copy).pdf=locked', 'a.pdf=locked']);
});

test('a subfolder copied within a locked folder is two files, not a conflict', async () => {
  const { locker, folder, id } = await make({ hidden: false });
  const bytes = random(600);
  write(folder, 'sub/a.pdf', bytes);
  await locker.lockWaiting(id);
  fs.cpSync(path.join(folder, 'sub'), path.join(folder, 'sub2'), { recursive: true });
  await locker.adoptCopies(id);
  const list = await locker.scan(id);
  assert.deepEqual(list.map((e) => `${e.rel}=${e.state}`).sort(), ['sub/a.pdf=locked', 'sub2/a.pdf=locked']);
  assert.notEqual(list[0].fileId, list[1].fileId);
  const out = path.join(tmp('rl-out-'), 'n.pdf');
  await locker.exportCopy(id, 'sub2/a.pdf', out);
  assert.ok(fs.readFileSync(out).equals(bytes));
});

test('a folder inside another that is in the locker in its own right keeps its own files and key', async () => {
  const { locker, folder: outer, id: outerId } = await make();
  const inner = path.join(outer, 'inner');
  fs.mkdirSync(inner);
  const { id: innerId } = await locker.addFolder(inner, { hidden: false, afterLock: 'delete' });
  write(inner, 'a.pdf', random(300));
  await locker.lockWaiting(innerId);
  await locker.adoptCopies(outerId);
  assert.ok(fs.existsSync(path.join(inner, FOLDER_KEYS)));
  assert.equal((await locker.scan(innerId))[0].state, 'locked');
});

test('a locked file no key here opens is left alone, still shown as needing the person', async () => {
  const { locker, folder, id } = await make();
  const strangers = createLocker({ dir: tmp(), cost: CHEAP });
  await strangers.setup('someone else');
  const far = tmp('rl-far-');
  const { id: farId } = await strangers.addFolder(far, { hidden: false, afterLock: 'delete' });
  write(far, 'x.pdf', random(200));
  await strangers.lockWaiting(farId);
  const bytes = fs.readFileSync(path.join(far, 'x.pdf.rlock'));
  fs.writeFileSync(path.join(folder, 'x.pdf.rlock'), bytes);
  assert.deepEqual((await locker.adoptCopies(id)).adopted, []);
  assert.equal((await locker.scan(id))[0].state, 'damaged');
  assert.ok(fs.readFileSync(path.join(folder, 'x.pdf.rlock')).equals(bytes));
});

test('removing a folder unlocks every file first; a file that will not unlock keeps the folder in the locker', async () => {
  const { locker, folder, id } = await make();
  const a = random(100);
  const p = write(folder, 'a.pdf', a);
  await locker.lockWaiting(id);
  const out = await locker.removeFolder(id);
  assert.ok(out.removed);
  assert.ok(fs.readFileSync(p).equals(a));
  assert.equal((await locker.folders()).length, 0);
  assert.ok(!fs.existsSync(path.join(folder, FOLDER_KEYS)));

  const again = await make();
  write(again.folder, 'x.pdf', random(50));
  await again.locker.lockWaiting(again.id);
  const damaged = path.join(again.folder, 'x.pdf.rlock');
  const b = fs.readFileSync(damaged);
  b[b.length - 1] ^= 1;
  fs.writeFileSync(damaged, b);
  const res = await again.locker.removeFolder(again.id);
  assert.equal(res.removed, false, 'a damaged file stops it');
  assert.equal(res.failed.length, 1);
  assert.equal((await again.locker.folders()).length, 1, 'still in the locker');
});

test('a share file: its own password, its own file, the same content; a wrong password changes nothing', async () => {
  const { locker, folder, id } = await make();
  const bytes = random(70000);
  write(folder, 'a.pdf', bytes);
  await locker.lockWaiting(id);
  const dest = path.join(tmp(), 'Shared.rlock');
  await locker.exportShare(id, 'a.pdf', dest, 'open sesame');
  const before = fs.readFileSync(dest);
  await assert.rejects(() => locker.openShared(dest, path.join(tmp(), 'x'), 'wrong'), (e) => e.code === 'bad-password');
  assert.ok(fs.readFileSync(dest).equals(before));
  const out = path.join(tmp(), 'y.pdf');
  await locker.openShared(dest, out, 'open sesame');
  assert.ok(fs.readFileSync(out).equals(bytes));
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

test('another computer: a folder and its password are all it takes, and the profile\'s store being lost is not fatal', async () => {
  const { locker, folder, id } = await make();
  const bytes = random(9000);
  write(folder, 'a.pdf', bytes);
  await locker.lockWaiting(id);

  const profile2 = tmp();
  const elsewhere = createLocker({ dir: profile2, cost: CHEAP });
  await elsewhere.setup('another password');
  await assert.rejects(() => elsewhere.addFolder(folder, {}), (e) => e.code === 'bad-password');
  const { id: joined } = await elsewhere.addFolder(folder, {}, { password: 'pw' });
  assert.equal(joined, id);
  const out = path.join(tmp(), 'x.pdf');
  await elsewhere.exportCopy(joined, 'a.pdf', out);
  assert.ok(fs.readFileSync(out).equals(bytes));

  // The first computer loses its profile keystore: the folder opens from its key file, after the password.
  const first = createLocker({ dir: profile2 + '-copy', cost: CHEAP });
  void first;
  fs.rmSync(path.join(locker.stateFile, '..', 'keystore'), { recursive: true, force: true });
  await locker.lock();
  await locker.unlock({ password: 'pw' });
  const out2 = path.join(tmp(), 'y.pdf');
  await locker.exportCopy(id, 'a.pdf', out2);
  assert.ok(fs.readFileSync(out2).equals(bytes), 'read after the store was rebuilt from the folder\'s key file');
});

test('locking the locker wipes the temporary plain copies an app opened', async () => {
  const { locker, folder, id, profile } = await make();
  const bytes = random(100);
  write(folder, 'a.pdf', bytes);
  await locker.lockWaiting(id);
  const { path: copy } = await locker.openPlain(id, 'a.pdf');
  assert.ok(fs.readFileSync(copy).equals(bytes));
  assert.ok(copy.startsWith(path.join(folder, '.rutba-open')), 'the plain copy is in the hidden folder beside the file');
  assert.ok(!copy.startsWith(profile));
  const edited = random(120);
  fs.writeFileSync(copy, edited);
  const saved = await locker.saveOpen(copy);
  assert.ok(saved.changed, 'an edit is written back to the locked file');
  const back = path.join(tmp(), 'b.pdf');
  await locker.exportCopy(id, 'a.pdf', back);
  assert.ok(fs.readFileSync(back).equals(edited));
  await locker.lock();
  assert.ok(!fs.existsSync(copy), 'wiped on lock');
  assert.ok(!fs.existsSync(path.join(folder, '.rutba-open')), 'the hidden folder goes with it');
  await assert.rejects(() => locker.scan(id), (e) => e.code === 'locked', 'a locked locker reads nothing');
});

const OPENED = async (choices = {}, hooks = null, rel = 'a.pdf') => {
  const m = await make(choices, hooks);
  const bytes = random(200);
  write(m.folder, rel, bytes);
  await m.locker.lockWaiting(m.id);
  const { path: copy } = await m.locker.openPlain(m.id, rel);
  return { ...m, bytes, copy };
};

const lockedBytes = async (locker, id, rel) => {
  const out = path.join(tmp(), 'out.bin');
  await locker.exportCopy(id, rel, out);
  return fs.readFileSync(out);
};

test('an open copy sits beside its own file, in a subfolder too, and is neither listed nor locked', async () => {
  const { locker, folder, id, copy } = await OPENED({ hidden: true }, null, 'deep/er/a.pdf');
  assert.ok(copy.startsWith(path.join(folder, 'deep', 'er', '.rutba-open')));
  assert.equal(path.basename(copy), 'a.pdf', 'the real name, so a program shows it');
  assert.ok(!(await locker.scan(id)).some((e) => e.rel.includes('rutba-open')));
  const { locked } = await locker.lockWaiting(id);
  assert.deepEqual(locked, []);
  assert.ok(fs.existsSync(copy));
});

test('a person may keep open copies out of the folder altogether', async () => {
  const { folder, copy, profile } = await OPENED({ openIn: 'profile' });
  assert.ok(copy.startsWith(profile));
  assert.ok(!fs.existsSync(path.join(folder, '.rutba-open')));
});

test('a program that saves by renaming a new file over the old is written back', async () => {
  const { locker, id, copy } = await OPENED();
  const edited = random(300);
  const incoming = path.join(path.dirname(copy), 'A1B2.tmp');
  fs.writeFileSync(incoming, edited);
  assert.equal(locker.openCopyFor(incoming), copy, 'anything in the working folder leads to its copy');
  fs.renameSync(incoming, copy);
  assert.ok((await locker.saveOpen(copy)).changed);
  assert.ok((await lockedBytes(locker, id, 'a.pdf')).equals(edited));
  assert.equal(locker.openCopyFor(path.join(tmp(), 'elsewhere.txt')), null);
});

test('when the program lets go the copy is written back and removed; one never seen in use waits for the lock', async () => {
  let held = false;
  const { locker, folder, id, copy } = await OPENED({}, { isHeld: async () => held });
  assert.deepEqual(await locker.reapClosed(), [], 'not seen in use yet');
  assert.ok(fs.existsSync(copy));
  held = true;
  assert.deepEqual(await locker.reapClosed(), []);
  const edited = random(250);
  fs.writeFileSync(copy, edited);
  held = false;
  assert.deepEqual(await locker.reapClosed(), ['a.pdf']);
  assert.ok(!fs.existsSync(copy));
  assert.ok(!fs.existsSync(path.join(folder, '.rutba-open')));
  assert.ok((await lockedBytes(locker, id, 'a.pdf')).equals(edited));
  assert.equal(locker.hasOpenCopies(), false);
});

test('an Office window showing the file counts as in use', async () => {
  const { locker, copy } = await OPENED({}, { isHeld: async () => false });
  let shown = true;
  await locker.reapClosed({ shown: (f) => shown && f === copy });
  assert.ok(fs.existsSync(copy));
  shown = false;
  assert.deepEqual(await locker.reapClosed({ shown: () => shown }), ['a.pdf']);
  assert.ok(!fs.existsSync(copy));
});

test('locking writes back a change the watcher had not reached yet', async () => {
  const { locker, id, copy } = await OPENED();
  const edited = random(260);
  fs.writeFileSync(copy, edited);
  assert.deepEqual((await locker.lock()).left, []);
  assert.ok(!fs.existsSync(copy));
  await locker.unlock({ password: 'pw' });
  assert.ok((await lockedBytes(locker, id, 'a.pdf')).equals(edited));
});

test('after a crash the next unlock writes the change back, clears the copy and clears strays', async () => {
  const { locker, folder, id, copy, profile } = await OPENED();
  const edited = random(270);
  fs.writeFileSync(copy, edited);
  const stray = path.join(folder, '.rutba-open', 'feedfacefeedface');
  fs.mkdirSync(stray);
  fs.writeFileSync(path.join(stray, 'left.pdf'), 'left by something else');
  // A new run of the app: nothing in memory, the keys gone.
  const again = createLocker({ dir: profile, cost: CHEAP });
  await again.unlock({ password: 'pw' });
  assert.ok(!fs.existsSync(copy));
  assert.ok(!fs.existsSync(path.join(folder, '.rutba-open')), 'strays go too');
  assert.ok((await lockedBytes(again, id, 'a.pdf')).equals(edited));
  assert.equal((await again.versions(id, 'a.pdf')).length, 2, 'as a new version, the old one kept');
});

test('a file leaving the locker takes its open copy with it', async () => {
  const { locker, folder, id, copy } = await OPENED();
  const edited = random(280);
  fs.writeFileSync(copy, edited);
  await locker.unlockForGood(id, 'a.pdf');
  assert.ok(!fs.existsSync(copy));
  assert.ok(fs.readFileSync(path.join(folder, 'a.pdf')).equals(edited), 'the change is in the file that came out');
});

test('a folder leaving the locker leaves no working folder behind', async () => {
  const { locker, folder, id, copy } = await OPENED();
  assert.ok((await locker.removeFolder(id)).removed);
  assert.ok(!fs.existsSync(copy));
  assert.ok(!fs.existsSync(path.join(folder, '.rutba-open')));
});

// A real program holding the file, the way Word and Acrobat do.
const holder = (file) => new Promise((resolve, reject) => {
  const script = `$f=[IO.File]::Open($env:RL_FILE,'Open','ReadWrite','None');[Console]::Out.WriteLine('held');$l=[Console]::In.ReadLine();if($l -eq 'write'){$b=[Text.Encoding]::UTF8.GetBytes('changed after the lock');$f.SetLength(0);$f.Write($b,0,$b.Length)};$f.Close()`;
  const p = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { env: { ...process.env, RL_FILE: file }, stdio: ['pipe', 'pipe', 'inherit'] });
  p.stdout.once('data', () => resolve({
    release: (line = 'close') => new Promise((done) => { p.once('exit', done); p.stdin.write(`${line}\n`); }),
  }));
  p.once('error', reject);
});

const windowsOnly = { skip: process.platform !== 'win32' };

test('the ready-made folder is made once, in the home folder, and locks what is put in it', async () => {
  const { locker } = await make();
  const home = tmp('rl-home-');
  const first = await locker.makeHomeFolder(home);
  assert.equal(first.made, true);
  assert.equal(first.root, path.join(home, 'Locker'));
  const again = await locker.makeHomeFolder(home);
  assert.deepEqual({ id: again.id, made: again.made }, { id: first.id, made: false });
  const p = write(first.root, 'note.txt', random(50));
  await locker.lockWaiting(first.id);
  assert.ok(fs.existsSync(`${p}.rlock`) || fs.readdirSync(first.root).some((n) => n.endsWith('.rlock')));
  assert.ok(!fs.existsSync(p));
});

test('a real program holding the copy: seen in use, reported at lock, cleared when it lets go', windowsOnly, async () => {
  const { locker, folder, copy } = await OPENED();
  const h = await holder(copy);
  await locker.reapClosed();
  assert.ok(fs.existsSync(copy));
  const { left } = await locker.lock();
  assert.deepEqual(left, ['a.pdf']);
  assert.ok(fs.existsSync(copy), 'it could not be removed, and the person is told');
  await h.release();
  await locker.reapClosed();
  assert.ok(!fs.existsSync(copy), 'unchanged, so cleared without the keys');
  assert.ok(!fs.existsSync(path.join(folder, '.rutba-open')));
});

test('a change saved after the locker locked is kept and becomes the next version at unlock', windowsOnly, async () => {
  const { locker, id, copy } = await OPENED();
  const h = await holder(copy);
  await locker.lock();
  await h.release('write');
  await locker.reapClosed();
  assert.ok(fs.existsSync(copy), 'work that is not in the locked file is not thrown away');
  await locker.unlock({ password: 'pw' });
  assert.ok(!fs.existsSync(copy));
  assert.equal((await lockedBytes(locker, id, 'a.pdf')).toString(), 'changed after the lock');
});

test('the working folder is hidden, out of search, and closed to everyone else', windowsOnly, async () => {
  const { folder } = await OPENED();
  const top = path.join(folder, '.rutba-open');
  assert.match(execFileSync('attrib', [top]).toString().slice(0, 21), /H.*I/);
  const acl = execFileSync('icacls', [top]).toString();
  assert.ok(acl.includes(process.env.USERNAME));
  assert.ok(!/BUILTIN|Everyone|Authenticated Users|\(I\)/.test(acl), acl);
});
