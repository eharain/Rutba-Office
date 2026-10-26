// A profile's password set, changed or taken away, stopped at every step as
// a crash would stop it: no file is ever left sealed with a key the lock in
// force does not give, and the next open, with the old password or the new,
// finishes the change (profiles.js, rekeyProfile, unlockProfile, finishRekey).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  readProfiles, makeLock, unlock, isSealed, createVault, rekeyProfile, unlockProfile, finishRekey, needsPassword, setProfileLock, writeProfiles, defaultList, DEFAULT_ID,
} from '../packages/office-shell/src/profiles.js';

const scratch = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-rekey-'));
const gone = (dir) => fs.rmSync(dir, { recursive: true, force: true });
const STOP = new Error('stopped here, as a crash would');

/** A profile folder with four covered files and one sealed file no service covered this run. */
function profileWith(key) {
  const dir = scratch();
  writeProfiles(dir, defaultList());
  const files = {
    secrets: path.join(dir, 'secrets.json'),
    recent: path.join(dir, 'recent.json'),
    calendar: path.join(dir, 'calendar.json'),
    talk: path.join(dir, 'talk', 'c1.json'),
  };
  const vault = createVault({ key });
  vault.cover(files.secrets);
  vault.cover(files.recent);
  vault.cover(files.calendar);
  vault.cover(path.join(dir, 'talk'));
  vault.writeJson(files.secrets, { 'mail:1': 'kept' });
  vault.writeJson(files.recent, [{ path: 'C:\\a.docx' }]);
  vault.writeJson(files.calendar, { calendars: [{ name: 'Home' }] });
  vault.writeJson(files.talk, [{ text: 'see you at nine' }]);
  // Sealed by a service that did not start this run (its file is found in the folder).
  const stray = path.join(dir, 'later', 'stray.json');
  createVault({ key }).writeJson(stray, { stray: true });
  return { dir, files: { ...files, stray }, vault };
}

/** Each file read with `key` alone: its contents, or null when that key does not open it. */
function readAll(files, key) {
  const v = createVault({ key });
  return Object.fromEntries(Object.entries(files).map(([k, f]) => [k, v.readJson(f, null)]));
}
const EXPECTED = { secrets: { 'mail:1': 'kept' }, recent: [{ path: 'C:\\a.docx' }], calendar: { calendars: [{ name: 'Home' }] }, talk: [{ text: 'see you at nine' }], stray: { stray: true } };

/** The next open, as the shell does it: either password, the vault with both keys, the change finished. */
async function reopen(dir, password, covered) {
  const p = readProfiles(dir).profiles[0];
  assert.ok(needsPassword(p) || !p.rekey?.lock, 'a profile with a password, or one being set, asks for it');
  const opened = await unlockProfile(p, password);
  assert.ok(opened, `the password opens it (${password})`);
  const vault = createVault({ key: opened.key, also: opened.rekey ? [opened.rekey.from] : [] });
  for (const f of covered) vault.cover(f);
  // Every service reads its file before the change is finished, and finds it whole.
  for (const f of covered) assert.notEqual(vault.readJson(f, null), null, `${path.basename(f)} read before the change is finished`);
  if (opened.rekey) await finishRekey({ base: dir, id: DEFAULT_ID, vault, keys: opened.rekey, folder: dir });
  return { key: opened.key, entry: readProfiles(dir).profiles[0] };
}

async function stoppedAt(kind, at) {
  const old = kind === 'set' ? null : await makeLock('old password');
  const next = kind === 'remove' ? null : await makeLock('new password');
  const { dir, files, vault } = profileWith(old?.key || null);
  try {
    if (old) setProfileLock(dir, DEFAULT_ID, old.lock);
    let count = 0;
    const step = (name) => {
      count += 1;
      if (count === at) throw STOP;
      return name;
    };
    let stopped = false;
    try {
      await rekeyProfile({ base: dir, id: DEFAULT_ID, vault, from: old?.key || null, to: next?.key || null, lock: next?.lock || null, folder: dir, step });
    } catch (err) {
      if (err !== STOP) throw err;
      stopped = true;
    }
    const entry = readProfiles(dir).profiles[0];
    if (stopped) {
      assert.ok(entry.rekey, `${kind}, stopped at step ${at}: the change is noted`);
      assert.equal(entry.lock?.salt ?? null, old?.lock.salt ?? null, `${kind}, stopped at step ${at}: the old lock is still in force`);
    } else {
      assert.equal(entry.rekey, undefined);
      assert.equal(entry.lock?.salt ?? null, next?.lock.salt ?? null);
    }
    const covered = [files.secrets, files.recent, files.calendar, files.talk];
    // Opened again with the old password, then (from a fresh copy of the stop) with the new.
    const passwords = [old ? 'old password' : null, next ? 'new password' : null].filter(Boolean);
    const snapshot = scratch();
    fs.cpSync(dir, snapshot, { recursive: true });
    try {
      for (const [i, pw] of passwords.entries()) {
        if (i > 0) { gone(dir); fs.cpSync(snapshot, dir, { recursive: true }); }
        if (!stopped && pw === 'old password' && next) {
          assert.equal(await unlockProfile(readProfiles(dir).profiles[0], pw), null, 'once finished, the old password is not the profile\'s');
          continue;
        }
        if (!stopped && !next) continue;
        const { entry: after } = await reopen(dir, pw, covered);
        assert.equal(after.rekey, undefined, `${kind}, stopped at ${at}, opened with ${pw}: the change is finished`);
        assert.equal(after.lock?.salt ?? null, next?.lock.salt ?? null, 'the new lock in force');
        assert.deepEqual(readAll(files, next?.key || null), EXPECTED, `${kind}, stopped at ${at}, opened with ${pw}: every file opens with the new key`);
        // The covered files sealed only while a password is in force; the stray one, never covered, sealed again only if it was sealed.
        for (const f of covered) assert.equal(isSealed(fs.readFileSync(f, 'utf8')), Boolean(next), `${path.basename(f)} sealed only when a password is in force`);
        assert.equal(isSealed(fs.readFileSync(files.stray, 'utf8')), Boolean(next && old));
      }
      if (kind === 'remove') {
        // Taken away and stopped: the old password opens it, and it is plain after.
        assert.deepEqual(readAll(files, null), EXPECTED);
      }
    } finally {
      gone(snapshot);
    }
    return stopped;
  } finally {
    gone(dir);
  }
}

for (const kind of ['set', 'change', 'remove']) {
  test(`a password ${kind === 'set' ? 'set' : kind === 'change' ? 'changed' : 'taken away'}, stopped at every step, is finished at the next open and loses nothing`, async () => {
    // Steps: noted, then a file each (five), then sealed: seven, and the eighth is past the end.
    let at = 1;
    for (; at <= 12; at++) {
      const stopped = await stoppedAt(kind, at);
      if (!stopped) break;
    }
    assert.ok(at >= 7, `${kind}: every step was stopped at (${at - 1} of them)`);
  });
}

test('a note that is not one Rutba Office wrote leaves the lock in force as it was', async () => {
  const old = await makeLock('old password');
  const next = await makeLock('new password');
  const p = { id: DEFAULT_ID, lock: old.lock, rekey: { lock: next.lock, toSealed: 'nonsense', fromSealed: 'nonsense' } };
  const opened = await unlockProfile(p, 'old password');
  assert.ok(opened.key.equals(await unlock(old.lock, 'old password')));
  assert.equal(opened.rekey, null);
  assert.equal(await unlockProfile(p, 'new password'), null, 'the new password cannot give the old key, so it does not open it');
  assert.equal(await unlockProfile(p, 'neither'), null);
});
