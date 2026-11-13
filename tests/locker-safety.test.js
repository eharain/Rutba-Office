// The safety copy: made before any plain file goes, shown with the locker shut,
// checked against the locked folder before it is removed, restored without the
// password into a new folder, and refused when the drive has not the room. Its
// key is sealed by the system's safe store (Electron's safeStorage; a stand-in
// here, and a stand-in safeStorage in the last tests, so nothing here needs
// Windows), and with the Office profile's password where it has one.
// What a stop midway leaves is cleared; a file that will not hold still can be
// left plain; a copy is asked about after 30 days and never removed unasked.

import { test } from 'node:test';
import { tempDir } from './temp-dirs.js';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { createLocker, plainRule, SAFETY_ASK_DAYS } from '../apps/desktop/main/extras/locker/locker.js';
import { createSafety, MANIFEST, RESERVE, spaceNeeded } from '../apps/desktop/main/extras/locker/safety.js';
import { systemSeal } from '../apps/desktop/main/extras/locker/index.js';
import { parseRules } from '../apps/desktop/main/extras/locker/rules.js';
import { createVault, isSealed } from '../packages/office-shell/src/profiles.js';
import { readHeader, LockerError } from '../apps/desktop/main/extras/locker/format.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
const random = (n) => crypto.randomBytes(n);
const tmp = (pre = 'rl-') => tempDir(pre);
const PLENTY = 1e15;

/**
 * The safe store's stand-in: seals with a key of its own, so a test can see
 * what was sealed. Another stand-in's seal does not open.
 */
function standIn(name = 'test') {
  const k = crypto.createHash('sha256').update(`stand-in:${name}`).digest();
  const sealed = [];
  return {
    name: 'system',
    sealed,
    available: async () => true,
    async seal(bytes) {
      sealed.push(Buffer.from(bytes));
      const iv = crypto.randomBytes(12);
      const c = crypto.createCipheriv('aes-256-gcm', k, iv);
      const body = Buffer.concat([c.update(bytes), c.final()]);
      return Buffer.concat([iv, c.getAuthTag(), body]);
    },
    async open(blob) {
      try {
        const d = crypto.createDecipheriv('aes-256-gcm', k, blob.subarray(0, 12));
        d.setAuthTag(blob.subarray(12, 28));
        return Buffer.concat([d.update(blob.subarray(28)), d.final()]);
      } catch {
        // As safeStorage.decryptString does: a plain Error, which the Locker says in its own words.
        throw new Error('Error while decrypting the ciphertext provided to safeStorage.decryptString.');
      }
    },
  };
}
const ACCOUNT = standIn();

/** A locker with the room `free` says. */
function lockerAt(profile, { free = () => PLENTY, at = null, large, protector = ACCOUNT, vault = null, staleMs } = {}) {
  return createLocker({ dir: profile, cost: CHEAP, large, vault, hooks: { safety: { freeSpace: async () => free(), at, protector, ...(staleMs === undefined ? {} : { staleMs }) } } });
}

/** The same profile's copies, read as the program would. */
const safetyAt = (profile, protector = ACCOUNT) => createSafety({ dir: profile, hooks: { protector } });

async function make({ safetyCopy = true, ...rest } = {}) {
  const profile = tmp();
  const folder = tmp('rl-folder-');
  const locker = lockerAt(profile, rest);
  await locker.setup('pw');
  const { id } = await locker.addFolder(folder, { afterLock: 'delete', hidden: true, safetyCopy });
  return { locker, folder, id, profile };
}

const write = (folder, rel, bytes) => {
  const p = path.join(folder, ...rel.split('/'));
  fs.mkdirSync(path.dirname(p), { recursive: true });
  fs.writeFileSync(p, bytes);
  return p;
};

const filesUnder = (dir) => {
  const out = [];
  const go = (d, rel) => {
    for (const e of fs.readdirSync(d, { withFileTypes: true })) {
      const r = rel ? `${rel}/${e.name}` : e.name;
      if (e.isDirectory()) go(path.join(d, e.name), r);
      else out.push(r);
    }
  };
  go(dir, '');
  return out.sort();
};

const safetyDir = (profile) => path.join(profile, 'safety');
const sealedCopies = (profile) => (fs.existsSync(safetyDir(profile)) ? fs.readdirSync(safetyDir(profile)).filter((n) => n.endsWith('.rsafe')) : []);

test('making: the copy is written and read back before the first plain file goes, then the folder locks as usual', async () => {
  const contents = { 'a.pdf': random(70000), 'sub/b.txt': random(10), 'sub/deep/c.bin': random(200000), 'empty.txt': Buffer.alloc(0) };
  let plainAtCopy = null;
  const { locker, folder, id, profile } = await make({ at: (step) => { if (step === 'safety-written') plainAtCopy = filesUnder(folder).filter((r) => !r.startsWith('.rutbalocker')); } });
  for (const [rel, bytes] of Object.entries(contents)) write(folder, rel, bytes);
  assert.equal((await locker.folders())[0].safetyWanted, true);
  const { locked, failed } = await locker.lockWaiting(id);
  assert.deepEqual(failed, []);
  assert.equal(locked.length, 4);
  assert.deepEqual(plainAtCopy, Object.keys(contents).sort(), 'every plain file was still there when the copy was sealed');
  for (const rel of Object.keys(contents)) assert.ok(!fs.existsSync(path.join(folder, ...rel.split('/'))), `${rel} plain copy removed after locking`);
  assert.equal((await locker.folders())[0].safetyWanted, undefined, 'the folder stops asking once the copy exists');
  const [copy] = await locker.safetyCopies();
  assert.deepEqual({ folderId: copy.folderId, files: copy.files, bytes: copy.bytes, inLocker: copy.inLocker, name: copy.name }, { folderId: id, files: 4, bytes: 70000 + 10 + 200000, inLocker: true, name: path.basename(folder) });
  assert.equal(sealedCopies(profile).length, 1);
  // An ordinary locked file of version 1 under the profile's own key, with the list of files first.
  const h = await readHeader(path.join(safetyDir(profile), sealedCopies(profile)[0]));
  assert.equal(h.version, 1);
  const manifest = await safetyAt(profile).manifest((await safetyAt(profile).list())[0]);
  assert.deepEqual(manifest.files.map((f) => f.rel).sort(), Object.keys(contents).sort());
  for (const f of manifest.files) assert.equal(f.sha256, crypto.createHash('sha256').update(contents[f.rel]).digest('hex'));
  // Nothing of the files' names or bytes is in the profile in the clear.
  const sealed = fs.readFileSync(path.join(safetyDir(profile), sealedCopies(profile)[0]));
  assert.ok(!sealed.includes(Buffer.from('sub/deep/c.bin')) && !sealed.includes(contents['a.pdf'].subarray(0, 64)));
  // Files locked later are not added to it: it is the folder as it was before the Locker touched it.
  write(folder, 'later.txt', random(5));
  await locker.lockWaiting(id);
  assert.equal((await locker.safetyCopies())[0].files, 4);
});

test('a file locked on its own still has the whole folder copied first', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'one.txt', random(100));
  write(folder, 'two.txt', random(100));
  await locker.lockOne(id, 'one.txt');
  const [copy] = await locker.safetyCopies();
  assert.equal(copy.files, 2);
});

test('showing: the copy is listed with the locker shut, and in a new run of the program', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(300));
  await locker.lockWaiting(id);
  await locker.lock();
  assert.equal(locker.isOpen(), false);
  const shut = await locker.safetyCopies();
  assert.equal(shut.length, 1);
  assert.equal(shut[0].files, 1);
  assert.ok(Date.now() - Date.parse(shut[0].made) < 60000);
  const later = lockerAt(profile);
  assert.equal((await later.safetyCopies()).length, 1);
});

test('no copy unless asked, and a folder that has one is not given a second', async () => {
  const { locker, folder, id, profile } = await make({ safetyCopy: false });
  write(folder, 'a.txt', random(30));
  await locker.lockWaiting(id);
  assert.deepEqual(await locker.safetyCopies(), []);
  assert.deepEqual(sealedCopies(profile), []);
  await locker.setChoices(id, { safetyCopy: true });
  write(folder, 'b.txt', random(30));
  await locker.lockWaiting(id);
  assert.equal((await locker.safetyCopies())[0].files, 1, 'asked for later: a copy of what was waiting then');
  await assert.rejects(() => locker.setChoices(id, { safetyCopy: true }), (e) => e.code === 'exists');
});

test('the first folder made for a person asks for a copy; later folders do not unless asked', async () => {
  const profile = tmp();
  const locker = lockerAt(profile);
  await locker.setup('pw');
  const home = await locker.makeHomeFolder(tmp('rl-home-'));
  const other = await locker.addFolder(tmp('rl-other-'), {});
  const list = await locker.folders();
  assert.equal(list.find((f) => f.id === home.id).safetyWanted, true);
  assert.equal(list.find((f) => f.id === other.id).safetyWanted, undefined);
  // An empty folder keeps asking until something is put in it to lock.
  write(home.root, 'x.txt', random(4));
  await locker.lockWaiting(other.id);
  assert.equal((await locker.folders()).find((f) => f.id === home.id).safetyWanted, true);
  await locker.lockWaiting(home.id);
  assert.equal((await locker.safetyCopies()).length, 1);
});

test('locking again as the locker closes leaves a folder still waiting for its copy as it is', async () => {
  const { locker, folder, id } = await make();
  const p = write(folder, 'a.txt', random(30));
  const out = await locker.lockWaiting(id, { deadline: Date.now() + 1000 });
  assert.deepEqual(out.locked, []);
  assert.equal(out.failed[0].code, 'safety');
  assert.ok(fs.existsSync(p));
  assert.deepEqual(await locker.safetyCopies(), []);
});

test('space: the room is worked out first, and a drive without it locks nothing and leaves nothing', async () => {
  let free = 10 * 1024 * 1024;
  const { locker, folder, id, profile } = await make({ free: () => free });
  const p = write(folder, 'big.bin', random(300000));
  const est = await locker.safetyEstimate(folder);
  assert.equal(est.files, 1);
  assert.equal(est.bytes, 300000);
  assert.ok(est.needed > 300000 && est.needed < 300000 + 8192, `needed ${est.needed}`);
  assert.equal(est.enough, false, 'the reserve is kept free too');
  await assert.rejects(() => locker.lockWaiting(id), (e) => e.code === 'space' && /needs about/.test(e.message) && /Nothing was locked/.test(e.message));
  await assert.rejects(() => locker.lockOne(id, 'big.bin'), (e) => e.code === 'space');
  assert.ok(fs.existsSync(p), 'the plain file is where it was');
  assert.equal((await locker.scan(id)).find((e) => e.rel === 'big.bin').state, 'waiting');
  assert.deepEqual(await locker.safetyCopies(), []);
  assert.deepEqual(fs.existsSync(safetyDir(profile)) ? fs.readdirSync(safetyDir(profile)).filter((n) => n !== 'key.json') : [], []);
  // With room, it goes ahead; and the estimate is what the copy came to, or a little over.
  free = PLENTY;
  assert.equal((await locker.safetyEstimate(folder)).enough, true);
  await locker.lockWaiting(id);
  const [copy] = await locker.safetyCopies();
  assert.ok(copy.size <= est.needed && copy.size > est.needed - 4096, `${copy.size} against ${est.needed}`);
  assert.equal(spaceNeeded([]) > 0, true);
  assert.ok(RESERVE >= 256 * 1024 * 1024);
});

test('space: a drive that fills while the copy is written leaves no part of it, and nothing is locked', async () => {
  const { locker, folder, id, profile } = await make({ at: (step) => { if (step === 'safety-written') throw Object.assign(new Error('no room'), { code: 'ENOSPC' }); } });
  const p = write(folder, 'a.bin', random(70000));
  await assert.rejects(() => locker.lockWaiting(id), (e) => e.code === 'space' && /ran out of room/.test(e.message));
  assert.ok(fs.existsSync(p));
  assert.deepEqual(fs.readdirSync(safetyDir(profile)).filter((n) => n !== 'key.json'), []);
  assert.deepEqual(await locker.safetyCopies(), []);
});

test('a file that changes while the copy is made stops it: nothing is locked and no part of the copy stays', async () => {
  let plain;
  const { locker, folder, id, profile } = await make({ at: (step) => { if (step === 'safety-measured') fs.appendFileSync(plain, 'saved meanwhile'); } });
  plain = write(folder, 'a.txt', random(1000));
  await assert.rejects(() => locker.lockWaiting(id), (e) => e.code === 'busy' && e.rel === 'a.txt' && /a\.txt changed while the safety copy/.test(e.message));
  assert.ok(fs.existsSync(plain));
  assert.deepEqual(fs.readdirSync(safetyDir(profile)).filter((n) => n !== 'key.json'), []);
  assert.deepEqual(await locker.safetyCopies(), []);
  const [rec] = await locker.folders();
  assert.equal(rec.safetyWanted, true, 'still asked for, to be made at the next lock');
  assert.equal(rec.safetyBusy, 'a.txt', 'the file is named on the folder, for its notice to offer leaving it plain');
});

test('saying no to the copy lets the folder lock without one', async () => {
  const { locker, folder, id } = await make({ free: () => 0 });
  const p = write(folder, 'a.txt', random(30));
  await assert.rejects(() => locker.lockWaiting(id), (e) => e.code === 'space');
  await locker.setChoices(id, { safetyCopy: false });
  await locker.lockWaiting(id);
  assert.ok(!fs.existsSync(p));
  assert.deepEqual(await locker.safetyCopies(), []);
});

test('restoring: with the locker shut and no password, into a new folder, never over anything', async () => {
  const contents = { 'Lease.pdf': random(140000), 'Tax/2025/return.xlsx': random(5000), 'Tax/note.txt': Buffer.from('keep this') };
  const { locker, folder, id, profile } = await make();
  for (const [rel, bytes] of Object.entries(contents)) write(folder, rel, bytes);
  const when = new Date('2025-03-04T05:06:07Z');
  fs.utimesSync(path.join(folder, 'Lease.pdf'), when, when);
  await locker.lockWaiting(id);
  await locker.lock();
  // A program started afresh, never unlocked: the copy opens without the password. That is what the window says.
  const fresh = lockerAt(profile);
  assert.equal(fresh.isOpen(), false);
  const [copy] = await fresh.safetyCopies();
  const dest = tmp('rl-restore-');
  write(dest, `${path.basename(folder)} (safety copy)/Lease.pdf`, Buffer.from('already here'));
  const out = await fresh.restoreSafetyCopy(copy.id, dest);
  assert.equal(out.dest, path.join(dest, `${path.basename(folder)} (safety copy 2)`));
  assert.equal(out.restored, 3);
  assert.deepEqual(out.failed, []);
  assert.deepEqual(filesUnder(out.dest), Object.keys(contents).sort());
  for (const [rel, bytes] of Object.entries(contents)) assert.ok(fs.readFileSync(path.join(out.dest, ...rel.split('/'))).equals(bytes), rel);
  assert.equal(Math.round(fs.statSync(path.join(out.dest, 'Lease.pdf')).mtimeMs / 1000), Math.round(when.getTime() / 1000));
  assert.equal(fs.readFileSync(path.join(dest, `${path.basename(folder)} (safety copy)`, 'Lease.pdf'), 'utf8'), 'already here', 'what was there is untouched');
  // A restore is a copy: the safety copy stays until the person says all is good.
  assert.equal((await fresh.safetyCopies()).length, 1);
  const again = await fresh.restoreSafetyCopy(copy.id, dest);
  assert.equal(path.basename(again.dest), `${path.basename(folder)} (safety copy 3)`);
});

test('restoring is refused into a locked folder, into Office\'s own, or to a folder that is not there', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(20));
  await locker.lockWaiting(id);
  const [copy] = await locker.safetyCopies();
  fs.mkdirSync(path.join(folder, 'inside'));
  await assert.rejects(() => locker.restoreSafetyCopy(copy.id, path.join(folder, 'inside')), (e) => e.code === 'choice');
  await assert.rejects(() => locker.restoreSafetyCopy(copy.id, profile), (e) => e.code === 'choice');
  await assert.rejects(() => locker.restoreSafetyCopy(copy.id, path.join(tmp(), 'not-there')), (e) => e.code === 'state');
  await assert.rejects(() => locker.restoreSafetyCopy('0'.repeat(32), tmp()), (e) => e.code === 'state');
});

test('a damaged safety copy restores nothing it cannot vouch for', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.bin', random(200000));
  await locker.lockWaiting(id);
  const [copy] = await locker.safetyCopies();
  const file = path.join(safetyDir(profile), sealedCopies(profile)[0]);
  const bytes = fs.readFileSync(file);
  bytes[bytes.length - 1000] ^= 0xFF;
  fs.writeFileSync(file, bytes);
  const dest = tmp('rl-restore-');
  await assert.rejects(() => locker.restoreSafetyCopy(copy.id, dest), (e) => e.code === 'tampered');
  const made = path.join(dest, `${path.basename(folder)} (safety copy)`);
  assert.deepEqual(fs.existsSync(made) ? filesUnder(made) : [], [], 'no half-written file is left');
});

test('all good: every file unlocks and matches, and only then is the copy wiped', async () => {
  const { locker, folder, id, profile } = await make({ large: 100000 });
  write(folder, 'small.txt', random(500));
  write(folder, 'big.bin', random(250000));
  await locker.lockWaiting(id);
  const check = await locker.checkSafetyCopy(id);
  assert.deepEqual(check, { files: 2, matched: 2, changed: [], missing: [], unreadable: [] });
  const sealedFile = path.join(safetyDir(profile), sealedCopies(profile)[0]);
  const out = await locker.removeSafetyCopy(id);
  assert.equal(out.removed, true);
  assert.ok(!fs.existsSync(sealedFile));
  assert.deepEqual(sealedCopies(profile), []);
  assert.deepEqual(await locker.safetyCopies(), []);
  assert.ok(fs.existsSync(path.join(safetyDir(profile), 'key.json')), 'the profile keeps its key for the next copy');
  await assert.rejects(() => locker.checkSafetyCopy(id), (e) => e.code === 'state');
});

test('all good: a file still waiting, changed, or gone is named and the copy kept unless the person accepts it', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'same.txt', random(40));
  write(folder, 'edited.txt', random(40));
  write(folder, 'gone.txt', random(40));
  await locker.lockWaiting(id);
  const edit = path.join(tmp(), 'edited.txt');
  fs.writeFileSync(edit, 'a newer version');
  await locker.importFile(id, edit, 'edited.txt');
  await locker.deleteFile(id, 'gone.txt');
  const out = await locker.removeSafetyCopy(id);
  assert.equal(out.removed, false);
  assert.deepEqual({ matched: out.check.matched, changed: out.check.changed, missing: out.check.missing }, { matched: 1, changed: ['edited.txt'], missing: ['gone.txt'] });
  assert.equal(sealedCopies(profile).length, 1);
  // The copy still holds what was there before.
  const back = await locker.restoreSafetyCopy((await locker.safetyCopies())[0].id, tmp());
  assert.equal(back.restored, 3);
  const accepted = await locker.removeSafetyCopy(id, { accept: true });
  assert.equal(accepted.removed, true);
  assert.deepEqual(sealedCopies(profile), []);
});

test('all good: a locked file that no longer reads back keeps the copy whatever is said', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.bin', random(150000));
  await locker.lockWaiting(id);
  const e = (await locker.scan(id)).find((x) => x.rel === 'a.bin');
  const bytes = fs.readFileSync(e.path);
  bytes[bytes.length - 500] ^= 0xFF;
  fs.writeFileSync(e.path, bytes);
  for (const accept of [false, true]) {
    const out = await locker.removeSafetyCopy(id, { accept });
    assert.equal(out.removed, false);
    assert.deepEqual(out.check.unreadable, ['a.bin']);
  }
  assert.equal(sealedCopies(profile).length, 1);
});

test('a folder taken out of the locker keeps its copy, which can then be removed on the person\'s word', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  const [copy] = await locker.safetyCopies();
  await assert.rejects(() => locker.discardSafetyCopy(copy.id), (e) => e.code === 'state');
  assert.equal((await locker.removeFolder(id)).removed, true);
  const [still] = await locker.safetyCopies();
  assert.equal(still.inLocker, false);
  await locker.lock();
  await assert.rejects(() => locker.discardSafetyCopy(copy.id), (e) => e.code === 'locked');
  await locker.unlock({ password: 'pw' });
  assert.deepEqual(await locker.discardSafetyCopy(copy.id), { removed: true });
  assert.deepEqual(sealedCopies(profile), []);
});

test('the profile key is made once and is not the password', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  const key = fs.readFileSync(path.join(safetyDir(profile), 'key.json'), 'utf8');
  const rec = JSON.parse(key);
  assert.deepEqual(Object.keys(rec).sort(), ['by', 'id', 'made', 'salt', 'sealed', 'v'], 'no plain key is kept');
  assert.equal(rec.by, 'system');
  await locker.removeSafetyCopy(id);
  await locker.setChoices(id, { safetyCopy: true });
  write(folder, 'b.txt', random(40));
  await locker.lockWaiting(id);
  assert.equal(fs.readFileSync(path.join(safetyDir(profile), 'key.json'), 'utf8'), key);
  await locker.changePassword('pw', 'another password');
  assert.equal(fs.readFileSync(path.join(safetyDir(profile), 'key.json'), 'utf8'), key, 'a new password does not make it again');
  await locker.lock();
  const back = await lockerAt(profile).restoreSafetyCopy((await locker.safetyCopies())[0].id, tmp());
  assert.equal(back.restored, 1, 'a new password changes nothing about the copy');
  // Neither password is kept in it, as it is, in base64 or in hex. Each field is compared whole: the two
  // letters "pw" turn up by chance in the file's base64 about once in 80 key files.
  const forms = (p) => [p, Buffer.from(p).toString('base64'), Buffer.from(p).toString('hex')];
  for (const [field, value] of Object.entries(rec)) {
    for (const p of ['pw', 'another password']) assert.ok(!forms(p).includes(String(value)), `${field} is not the password`);
  }
  assert.ok(!key.includes('another password'));
});

/* ── the key sealed for this user account, and for this Office profile ── */

test('the key is sealed by the system: the secret is never in the file, and another account cannot open it', async () => {
  const account = standIn('account-a');
  const { locker, folder, id, profile } = await make({ protector: account });
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  assert.equal(account.sealed.length, 1, 'sealed once, when the key was made');
  const secret = account.sealed[0];
  const file = fs.readFileSync(path.join(safetyDir(profile), 'key.json'));
  for (const form of [secret.toString('hex'), secret.toString('base64')]) assert.ok(!file.includes(form), 'the secret is not in the key file in any form');
  assert.ok(!file.includes(secret), 'nor as raw bytes');
  const [copy] = await locker.safetyCopies();
  // The same files read on another account (another safe store): nothing opens, nothing is written, and the Locker says so in its own words.
  const parent = tmp();
  await assert.rejects(() => lockerAt(profile, { protector: standIn('account-b') }).restoreSafetyCopy(copy.id, parent), (e) => e.code === 'protect' && /safe store would not open/.test(e.message));
  assert.deepEqual(fs.readdirSync(parent), []);
  assert.equal((await lockerAt(profile, { protector: account }).restoreSafetyCopy(copy.id, tmp())).restored, 1);
});

test('with nothing to seal the key with, no copy is offered or asked for', async () => {
  const profile = tmp();
  const folder = tmp('rl-folder-');
  const locker = lockerAt(profile, { protector: null });
  await locker.setup('pw');
  assert.deepEqual(await locker.safetyInfo(), { available: false, scope: 'account' });
  assert.equal((await locker.safetyEstimate(folder)).available, false);
  const { id } = await locker.addFolder(folder, { safetyCopy: true });
  assert.equal((await locker.folders())[0].safetyWanted, undefined, 'asked for, but not kept as asked: the folder locks as it would without one');
  await assert.rejects(() => locker.setChoices(id, { safetyCopy: true }), (e) => e.code === 'protect');
  const home = await locker.makeHomeFolder(tmp());
  assert.equal((await locker.folders()).find((f) => f.id === home.id).safetyWanted, undefined);
  write(folder, 'a.txt', random(10));
  assert.equal((await locker.lockWaiting(id)).locked.length, 1);
});

test('an Office profile with a password seals the key with it too: another profile on the same account cannot open the copy', async () => {
  const profileKey = random(32);
  const vault = createVault({ key: profileKey });
  const { locker, folder, id, profile } = await make({ vault });
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  assert.ok(isSealed(fs.readFileSync(path.join(safetyDir(profile), 'key.json'), 'utf8')), 'the key file is sealed by the profile\'s vault');
  const [copy] = await locker.safetyCopies();
  assert.equal(copy.scope, 'profile');
  assert.equal((await locker.safetyInfo()).scope, 'profile');
  // Another Office profile, same user account (same safe store), without this profile's password.
  const parent = tmp();
  await assert.rejects(() => lockerAt(profile).restoreSafetyCopy(copy.id, parent), (e) => e.code === 'protect');
  await assert.rejects(() => lockerAt(profile, { vault: createVault({ key: random(32) }) }).restoreSafetyCopy(copy.id, parent), (e) => e.code === 'protect');
  assert.deepEqual(fs.readdirSync(parent), []);
  // A key file it cannot read is never made again over the one the copy needs: a new copy is refused instead.
  const keyFile = path.join(safetyDir(profile), 'key.json');
  const before = fs.readFileSync(keyFile, 'utf8');
  const other = lockerAt(profile);
  await other.unlock({ password: 'pw' });
  const second = tmp('rl-folder-');
  write(second, 'x.txt', random(5));
  const { id: id2 } = await other.addFolder(second, { safetyCopy: true });
  await assert.rejects(() => other.lockWaiting(id2), (e) => e.code === 'protect');
  assert.equal(fs.readFileSync(keyFile, 'utf8'), before);
  assert.ok(fs.existsSync(path.join(second, 'x.txt')), 'and nothing of that folder was locked');
  // This profile opens it.
  assert.equal((await lockerAt(profile, { vault: createVault({ key: profileKey }) }).restoreSafetyCopy(copy.id, tmp())).restored, 1);
});

test('the profile\'s password changed or taken away: its vault seals the key again, and the copy still opens', async () => {
  const first = random(32);
  const vault = createVault({ key: first });
  const { locker, folder, id, profile } = await make({ vault });
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  const [copy] = await locker.safetyCopies();
  const next = random(32);
  await vault.rekey(next);
  assert.equal((await lockerAt(profile, { vault: createVault({ key: next }) }).restoreSafetyCopy(copy.id, tmp())).restored, 1);
  await assert.rejects(() => lockerAt(profile, { vault: createVault({ key: first }) }).restoreSafetyCopy(copy.id, tmp()), (e) => e.code === 'protect');
  await vault.rekey(null);
  const plainVault = createVault();
  const again = lockerAt(profile, { vault: plainVault });
  assert.equal((await again.safetyCopies())[0].scope, 'account', 'without a profile password the window says any profile on the account can open it');
  assert.equal((await again.restoreSafetyCopy(copy.id, tmp())).restored, 1);
});

/* ── a stop midway, and a file that will not hold still ─────────────────── */

test('a stop after the copy was listed but before the folder stopped asking: the next lock uses that copy', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  // The state as a stop at that moment leaves it.
  const stateFile = path.join(profile, 'state.json');
  const st = JSON.parse(fs.readFileSync(stateFile, 'utf8'));
  st.folders[0].safetyWanted = true;
  fs.writeFileSync(stateFile, JSON.stringify(st));
  const again = lockerAt(profile);
  await again.unlock({ password: 'pw' });
  write(folder, 'b.txt', random(40));
  const out = await again.lockWaiting(id);
  assert.deepEqual(out.failed, []);
  assert.equal(out.locked.length, 1, 'b.txt locks: "already has a safety copy" does not stop it');
  assert.equal((await again.folders())[0].safetyWanted, undefined);
  assert.equal((await again.safetyCopies()).length, 1);
  assert.equal(sealedCopies(profile).length, 1);
});

test('a stop while a copy is written, or before it is listed, leaves nothing the next run keeps', async () => {
  const { locker, folder, id, profile } = await make({ at: (step) => { if (step === 'safety-renamed') throw new Error('the power went'); } });
  const p = write(folder, 'a.txt', random(40));
  await assert.rejects(() => locker.lockWaiting(id), /power went/);
  assert.ok(fs.existsSync(p), 'nothing was locked');
  assert.equal(sealedCopies(profile).length, 0, 'a copy written but not listed is wiped at once by the run that wrote it');
  // A real power cut runs no clean-up: a whole copy, never listed, and a part are left as it leaves them.
  fs.writeFileSync(path.join(safetyDir(profile), `${crypto.randomBytes(16).toString('hex')}.rsafe`), random(5000));
  const part = path.join(safetyDir(profile), `.rl-${crypto.randomBytes(12).toString('hex')}.part`);
  fs.writeFileSync(part, random(5000));
  // Another program starting now leaves them: they could be its own copy, being written.
  assert.deepEqual(await lockerAt(profile).tidySafety(), { parts: 0, unlisted: 0 });
  assert.equal(sealedCopies(profile).length, 1);
  // The next run of the program, once they are old.
  const again = lockerAt(profile, { staleMs: 0 });
  await again.unlock({ password: 'pw' });
  assert.deepEqual(await again.tidySafety(), { parts: 1, unlisted: 1 });
  assert.ok(!fs.existsSync(part));
  assert.equal(sealedCopies(profile).length, 0);
  const out = await again.lockWaiting(id);
  assert.equal(out.locked.length, 1);
  const copies = await again.safetyCopies();
  assert.equal(copies.length, 1);
  assert.deepEqual(sealedCopies(profile), [copies[0].id + '.rsafe']);
  assert.deepEqual(fs.readdirSync(safetyDir(profile)).sort(), [copies[0].id + '.rsafe', 'copies.json', 'key.json'].sort());
});

test('a copy listed for a folder still asking, whose file is gone, is not taken for one: a new copy is made first', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(40));
  // The state a stop leaves (a copy listed, the folder still asking), then the copy's file taken away by hand or a scanner.
  const safety = safetyAt(profile);
  const [entry] = (await locker.scan(id)).filter((e) => e.state === 'waiting');
  const old = await safety.make({ folderId: id, root: folder, files: [{ path: entry.path, rel: entry.rel, size: entry.size, modified: entry.modified }] });
  fs.rmSync(path.join(safetyDir(profile), old.file));
  assert.equal((await locker.folders())[0].safetyWanted, true);
  const out = await locker.lockWaiting(id);
  assert.equal(out.locked.length, 1);
  const copies = await locker.safetyCopies();
  assert.equal(copies.length, 1);
  assert.notEqual(copies[0].id, old.id, 'a new copy, not the one that is gone');
  assert.equal((await locker.restoreSafetyCopy(copies[0].id, tmp())).restored, 1, 'and it opens');
});

test('a copy that no longer opens can be removed on the person\'s word, and only then', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  const [copy] = await locker.safetyCopies();
  // A password reset, or a lost Local State: the safe store no longer opens the key.
  const after = lockerAt(profile, { protector: standIn('after-reset') });
  await after.unlock({ password: 'pw' });
  const first = await after.removeSafetyCopy(id);
  assert.equal(first.removed, false);
  assert.ok(first.check.unopenable, 'said to be unopenable, not an error at every unlock');
  assert.equal(sealedCopies(profile).length, 1, 'kept until the person says');
  const second = await after.removeSafetyCopy(id, { accept: true });
  assert.equal(second.removed, true);
  assert.deepEqual(await after.safetyCopies(), []);
  assert.equal(sealedCopies(profile).length, 0);
  void copy;
});

test('a key written without the profile\'s vault, in a profile with a password, is sealed at the next read and said as it is', async () => {
  const profileKey = random(32);
  // The command line, which has no vault, makes the first copy.
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  const keyFile = path.join(safetyDir(profile), 'key.json');
  assert.ok(!isSealed(fs.readFileSync(keyFile, 'utf8')));
  const app = lockerAt(profile, { vault: createVault({ key: profileKey }) });
  assert.equal((await app.safetyCopies())[0].scope, 'account', 'not said to be the profile\'s until it is');
  assert.equal((await app.restoreSafetyCopy((await app.safetyCopies())[0].id, tmp())).restored, 1);
  assert.ok(isSealed(fs.readFileSync(keyFile, 'utf8')), 'sealed with the profile\'s password once read through its vault');
  assert.equal((await app.safetyCopies())[0].scope, 'profile');
});

test('a file whose name cannot be used here, or one of the same name, is passed over and the rest restored', async () => {
  const profile = tmp();
  const folder = tmp('rl-folder-');
  const safety = safetyAt(profile);
  const a = write(folder, 'a.txt', random(20));
  const b = write(folder, 'b.txt', random(30));
  const st = (p) => fs.statSync(p);
  const files = [
    { path: a, rel: 'bad:name.txt', size: st(a).size, modified: st(a).mtimeMs },
    { path: b, rel: 'ok.txt', size: st(b).size, modified: st(b).mtimeMs },
    { path: a, rel: 'sub/x.txt', size: st(a).size, modified: st(a).mtimeMs },
    { path: b, rel: 'sub/X.TXT', size: st(b).size, modified: st(b).mtimeMs },
  ];
  const copy = await safety.make({ folderId: 'f'.repeat(32), root: folder, files });
  const back = await safety.restore(copy, tmp());
  const failed = back.failed.map((x) => x.rel);
  assert.ok(failed.includes('bad:name.txt'));
  assert.ok(fs.existsSync(path.join(back.dest, 'ok.txt')), 'the files after it are restored');
  assert.ok(fs.existsSync(path.join(back.dest, 'sub', 'x.txt')));
  if (process.platform === 'win32') assert.ok(failed.includes('sub/X.TXT'), 'one of the same name in another case is not written over the first');
  assert.equal(back.restored + back.failed.length, 4);
});

test('the restored folder is named from the sealed list, not the plain one', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  const list = path.join(safetyDir(profile), 'copies.json');
  const rec = JSON.parse(fs.readFileSync(list, 'utf8'));
  rec.copies[0].name = '..\\..\\elsewhere';
  fs.writeFileSync(list, JSON.stringify(rec));
  const parent = tmp();
  const back = await lockerAt(profile).restoreSafetyCopy(rec.copies[0].id, parent);
  assert.equal(path.dirname(back.dest), parent);
  assert.equal(path.basename(back.dest), `${path.basename(folder)} (safety copy)`);
});

test('leaving a file plain is refused when it would change which other files lock', async () => {
  const { locker, folder, id } = await make();
  write(folder, 'top.log', random(10));
  write(folder, 'sub/report.pdf', random(10));
  write(folder, 'sub/notes.txt', random(10));
  fs.writeFileSync(path.join(folder, 'sub', '.rutbalocker'), '*.pdf\n');
  const before = Object.fromEntries((await locker.scan(id)).map((e) => [e.rel, e.state]));
  assert.equal(before['sub/notes.txt'], 'excluded');
  await assert.rejects(() => locker.leavePlain(id, 'top.log'), (e) => e.code === 'choice' && /other files/.test(e.message));
  assert.ok(!fs.existsSync(path.join(folder, '.rutbalocker')), 'nothing written');
  const after = Object.fromEntries((await locker.scan(id)).map((e) => [e.rel, e.state]));
  assert.deepEqual(after, before);
});

test('the sweep keeps every copy when the list of copies cannot be read', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  fs.writeFileSync(path.join(safetyDir(profile), 'copies.json'), '{ cut sh');
  assert.deepEqual(await safetyAt(profile).tidy(), { parts: 0, unlisted: 0 });
  assert.equal(sealedCopies(profile).length, 1);
});

test('lock the rest, leave this one plain: the file that kept changing stays plain and the rest is copied and locked', async () => {
  let plain;
  let changes = 0;
  const { locker, folder, id } = await make({ at: (step) => { if (step === 'safety-measured' && plain && changes < 1) { changes += 1; fs.appendFileSync(plain, 'more'); } } });
  write(folder, 'a.txt', random(30));
  write(folder, 'sub/b.txt', random(30));
  plain = write(folder, 'logs/run [2].log', random(100));
  await assert.rejects(() => locker.lockWaiting(id), (e) => e.code === 'busy' && e.rel === 'logs/run [2].log');
  assert.equal((await locker.folders())[0].safetyBusy, 'logs/run [2].log');
  assert.deepEqual(await locker.leavePlain(id, 'logs/run [2].log'), { rel: 'logs/run [2].log', state: 'excluded' });
  assert.equal((await locker.folders())[0].safetyBusy, undefined);
  const rules = fs.readFileSync(path.join(folder, '.rutbalocker'), 'utf8');
  assert.match(rules, /^\*\*$/m, 'every other file is still locked');
  assert.ok(rules.includes(plainRule('logs/run [2].log')));
  const out = await locker.lockWaiting(id);
  assert.deepEqual(out.failed, []);
  assert.deepEqual(out.locked.map((e) => e.rel ?? e).sort(), ['a.txt', 'sub/b.txt']);
  assert.ok(fs.existsSync(plain), 'left plain where it is');
  const states = Object.fromEntries((await locker.scan(id)).map((e) => [e.rel, e.state]));
  assert.equal(states['logs/run [2].log'], 'excluded');
  assert.equal(states['a.txt'], 'locked');
  const [copy] = await locker.safetyCopies();
  assert.equal(copy.files, 2, 'the copy holds the files that were locked');
  // Asked again, it changes nothing.
  await locker.leavePlain(id, 'logs/run [2].log');
  assert.equal(fs.readFileSync(path.join(folder, '.rutbalocker'), 'utf8'), rules);
});

test('leaving a file plain keeps the folder\'s own rules', async () => {
  const { locker, folder, id } = await make();
  await locker.writeRules(id, '# mine\n*.txt\n');
  write(folder, 'a.txt', random(30));
  write(folder, 'b.txt', random(30));
  await locker.leavePlain(id, 'b.txt');
  assert.equal(fs.readFileSync(path.join(folder, '.rutbalocker'), 'utf8'), `# mine\n*.txt\n# Left plain: it kept changing while the safety copy was being made.\n${plainRule('b.txt')}\n`);
  const out = await locker.lockWaiting(id);
  assert.equal(out.locked.length, 1);
});

test('a rule that leaves one file plain names that file and nothing else', () => {
  assert.equal(plainRule('a.txt'), '!/a.txt');
  assert.equal(plainRule('logs/run [2].log'), '!/logs/run \\[2\\].log');
  // And the rule, read back as the Locker reads rules, matches that path and no other.
  const [rule] = parseRules(plainRule('logs/run [2].log'));
  assert.ok(rule.negate && rule.anchored && rule.re.test('logs/run [2].log'));
  assert.ok(!rule.re.test('logs/run 2.log') && !rule.re.test('other/logs/run [2].log'));
  assert.equal(plainRule('odd\\name'), '!/odd\\\\name');
});

/* ── 30 days, and never removed unasked ──────────────────────────────────── */

test('after 30 days a copy is asked about at each unlock, and is never removed unasked', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  const [made] = await locker.safetyCopies();
  const at = (days) => Date.parse(made.made) + days * 86400000 + 1000;
  assert.equal(SAFETY_ASK_DAYS, 30);
  assert.equal((await locker.safetyCopies({ now: at(29) }))[0].ask, false);
  const late = (await locker.safetyCopies({ now: at(30) }))[0];
  assert.equal(late.ask, true);
  assert.equal(late.days, 30);
  assert.equal((await locker.safetyCopies({ now: at(400) }))[0].ask, true);
  for (let i = 0; i < 3; i += 1) {
    await locker.lock();
    await locker.unlock({ password: 'pw' });
    await locker.lockWaiting(id);
  }
  assert.equal((await locker.safetyCopies({ now: at(400) })).length, 1, 'still there');
  assert.equal(sealedCopies(profile).length, 1);
});

test('a manifest-named file in the folder is copied like any other', async () => {
  const { locker, folder, id } = await make();
  write(folder, MANIFEST, Buffer.from('{"not":"ours"}'));
  write(folder, 'a.txt', random(10));
  await locker.lockWaiting(id);
  const back = await locker.restoreSafetyCopy((await locker.safetyCopies())[0].id, tmp());
  assert.equal(fs.readFileSync(path.join(back.dest, MANIFEST), 'utf8'), '{"not":"ours"}');
  assert.equal(back.restored, 2);
});

test('two folders making their copies at once are both listed', async () => {
  const profile = tmp();
  const locker = lockerAt(profile);
  await locker.setup('pw');
  const ids = [];
  for (let i = 0; i < 3; i += 1) {
    const root = tmp(`rl-many-${i}-`);
    write(root, 'a.bin', random(150000));
    ids.push((await locker.addFolder(root, { safetyCopy: true })).id);
  }
  await Promise.all(ids.map((id) => locker.lockWaiting(id)));
  const copies = await locker.safetyCopies();
  assert.deepEqual(copies.map((c) => c.folderId).sort(), [...ids].sort());
  assert.equal(sealedCopies(profile).length, 3);
  assert.ok((await locker.folders()).every((f) => !f.safetyWanted), 'every folder stopped asking: no change to the list was lost');
  // One key for the profile, made once though three copies wanted it together: every copy opens.
  for (const c of copies) assert.equal((await locker.restoreSafetyCopy(c.id, tmp())).restored, 1);
});

/* ── the safe store, the same on every system (owner, 12 Nov 2026: rely less on Windows) ── */

/** Electron's safeStorage as index.js sees it: `encryptString` and `decryptString`, and what it says it has. */
function fakeSafeStorage({ available = true, backend } = {}) {
  const k = crypto.randomBytes(32);
  return {
    isEncryptionAvailable: () => available,
    ...(backend === undefined ? {} : { getSelectedStorageBackend: () => backend }),
    encryptString(text) {
      const iv = crypto.randomBytes(12);
      const c = crypto.createCipheriv('aes-256-gcm', k, iv);
      const body = Buffer.concat([c.update(text, 'utf8'), c.final()]);
      return Buffer.concat([iv, c.getAuthTag(), body]);
    },
    decryptString(blob) {
      const d = crypto.createDecipheriv('aes-256-gcm', k, blob.subarray(0, 12));
      d.setAuthTag(blob.subarray(12, 28));
      return Buffer.concat([d.update(blob.subarray(28)), d.final()]).toString('utf8');
    },
  };
}

test('given no safe store, the Locker makes no copy and never keeps the key plain', async () => {
  const profile = tmp();
  const folder = tmp('rl-folder-');
  // No protector at all: this module has no default of its own, on any system.
  const locker = createLocker({ dir: profile, cost: CHEAP, hooks: { safety: { freeSpace: async () => PLENTY } } });
  await locker.setup('pw');
  assert.equal((await locker.safetyInfo()).available, false);
  const { id } = await locker.addFolder(folder, { safetyCopy: true });
  await assert.rejects(() => locker.setChoices(id, { safetyCopy: true }), (e) => e.code === 'protect' && /no keyring to keep their key in/.test(e.message));
  write(folder, 'a.txt', random(10));
  assert.equal((await locker.lockWaiting(id)).locked.length, 1, 'the folder still locks, without a copy');
  assert.ok(!fs.existsSync(path.join(safetyDir(profile), 'key.json')), 'no key written, plain or otherwise');
});

test('the safe store through safeStorage: offered only where it is real, and the key round-trips', async () => {
  assert.equal(await systemSeal(fakeSafeStorage({ available: false })).available(), false, 'isEncryptionAvailable() false: none');
  const real = fakeSafeStorage({ backend: 'gnome_libsecret' });
  assert.equal(await systemSeal(real).available(), true);
  if (process.platform === 'linux') assert.equal(await systemSeal(fakeSafeStorage({ backend: 'basic_text' })).available(), false, 'no keyring on Linux: a fixed key is not taken for one');
  const broken = { isEncryptionAvailable: () => { throw new Error('not ready'); } };
  assert.equal(await systemSeal(broken).available(), false, 'a safe store that cannot say is taken for none');
  const store = systemSeal(real);
  assert.equal(store.name, 'system');
  const secret = random(32);
  const sealed = await store.seal(secret);
  assert.ok(!Buffer.from(sealed).includes(secret) && !Buffer.from(sealed).toString('latin1').includes(secret.toString('base64')));
  assert.deepEqual(Buffer.from(await store.open(sealed)), secret);

  // A whole copy made and restored through it, and a safe store that will not keep the key stops the copy and locks nothing.
  const { locker, folder, id, profile } = await make({ protector: store });
  write(folder, 'a.txt', random(40));
  await locker.lockWaiting(id);
  assert.equal(JSON.parse(fs.readFileSync(path.join(safetyDir(profile), 'key.json'), 'utf8')).by, 'system');
  assert.equal((await lockerAt(profile, { protector: store }).restoreSafetyCopy((await locker.safetyCopies())[0].id, tmp())).restored, 1);
  const refusing = systemSeal({ ...fakeSafeStorage(), encryptString: () => { throw new Error('Encryption is not available.'); } });
  const other = await make({ protector: refusing });
  write(other.folder, 'b.txt', random(10));
  await assert.rejects(() => other.locker.lockWaiting(other.id), (e) => e.code === 'protect' && /would not keep/.test(e.message));
  assert.ok(fs.existsSync(path.join(other.folder, 'b.txt')), 'nothing locked, nothing removed');
});

test('the safety folder is made with no system tool, and off Windows is this account\'s alone', async () => {
  const { locker, folder, id, profile } = await make();
  write(folder, 'a.txt', random(10));
  await locker.lockWaiting(id);
  assert.ok(fs.statSync(safetyDir(profile)).isDirectory());
  if (process.platform !== 'win32') assert.equal(fs.statSync(safetyDir(profile)).mode & 0o777, 0o700);
  const source = fs.readFileSync(new URL('../apps/desktop/main/extras/locker/safety.js', import.meta.url), 'utf8');
  for (const tool of [/\bicacls\b/i, /\battrib(\.exe)?\b/i, /\bwhoami\b/i, /powershell/i, /child_process/, /working\.js/, /dpapi\.js/]) assert.ok(!tool.test(source), `safety.js runs no ${tool}`);
});

test('each system\'s notice names what protects the copy there, and no other system (owner, 13 Nov 2026)', () => {
  // The window is JSX and cannot be loaded here, so its words are read from its source.
  const ui = fs.readFileSync(new URL('../apps/desktop/renderer/apps/locker.js', import.meta.url), 'utf8').replace(/\r\n/g, '\n');
  const table = /const SAFETY_WORDS = \{\n([\s\S]*?)\n\}\[SYSTEM\];/.exec(ui);
  assert.ok(table, 'one table of the notices, picked by system');
  const block = (name) => {
    const m = new RegExp(`^  ${name}: \\{\\n([\\s\\S]*?)\\n  \\},`, 'm').exec(table[1]);
    assert.ok(m, `${name} has its words`);
    for (const k of ['add', 'account', 'profile', 'others']) assert.match(m[1], new RegExp(`^    ${k}: \\(\\) => t\\('`, 'm'), `${name}: ${k}`);
    return m[1];
  };
  const [windows, mac, linux] = ['windows', 'mac', 'linux'].map(block);
  assert.match(windows, /Windows account/);
  assert.match(windows, /Your Windows password protects it/);
  assert.match(mac, /this Mac account/);
  assert.match(mac, /through the Keychain/);
  assert.match(linux, /this user account/);
  assert.match(linux, /keyring\\'s password protects it/);
  assert.match(linux, /If your keyring has no password, nothing does\./);
  for (const [name, words] of [['mac', mac], ['linux', linux]]) assert.ok(!/Windows/.test(words), `${name} never says Windows`);
  assert.ok(!/Keychain|keyring/.test(windows), 'Windows says neither Keychain nor keyring');
  // Linux with no keyring: the line where the box would be, and the same word when a copy is asked for anyway.
  assert.match(ui, /!canCopy && SYSTEM === 'linux' \? <div className="lk-safety-none">\{t\('Safety copies are not offered on this computer, because it has no keyring to keep their key in\. The folder still locks as usual\.'\)\}/);
  assert.ok(!/\bsafe store\b/.test(ui), 'the window never says "safe store"');
});
