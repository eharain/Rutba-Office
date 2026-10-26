// Each profile's saved passwords in a keystore of its own (D-2026-10-26-09):
// a long id made once, a random key linked to it and wrapped by the system's
// protection (a stand-in here, so another computer can be played), a store
// that refuses another computer and another profile and is put aside, not
// read, when it does, and the one-time move of passwords kept the old way.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import {
  profileUid, storeNameOf, openProfileKeystore, moveOldSecrets,
  ID_FILE, KEY_FILE, KEYSTORE_DIR, MOVE_FILE, OLD_SECRETS_FILE, UNREADABLE_FILE, NOT_KEPT, ITERATIONS,
} from '../packages/office-shell/src/profile-keystore.js';
import { readProfiles, makeProfile, renameProfile, profileFolder, rekeyProfile, makeLock, createVault, isSealed } from '../packages/office-shell/src/profiles.js';

const scratch = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-keystore-'));
const gone = (dir) => fs.rmSync(dir, { recursive: true, force: true });

/** The system's protection on one computer: what it wraps, only the same computer unwraps (as DPAPI does). */
function computer(on = true) {
  const machine = crypto.randomBytes(32);
  return {
    available: () => on,
    encrypt(text) {
      const iv = crypto.randomBytes(12);
      const c = crypto.createCipheriv('aes-256-gcm', machine, iv);
      const body = Buffer.concat([c.update(String(text), 'utf8'), c.final()]);
      return Buffer.concat([iv, c.getAuthTag(), body]);
    },
    decrypt(bytes) {
      const b = Buffer.from(bytes);
      const d = crypto.createDecipheriv('aes-256-gcm', machine, b.subarray(0, 12));
      d.setAuthTag(b.subarray(12, 28));
      return Buffer.concat([d.update(b.subarray(28)), d.final()]).toString('utf8');
    },
  };
}

/** Every file under a folder, as one text, to look for a secret in the clear. */
function allText(dir) {
  let out = '';
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    out += e.isDirectory() ? allText(p) : fs.readFileSync(p, 'latin1');
  }
  return out;
}

const uidOf = () => crypto.randomBytes(32).toString('hex');
const settingsStub = (start = {}) => {
  const s = { ...start };
  return { get: (k, f) => (k in s ? s[k] : f), set: (k, v) => { s[k] = v; }, delete: (k) => { delete s[k]; }, all: () => ({ ...s }) };
};

test('a profile gets a long id once: kept in its folder and in its entry, the same after a rename', () => {
  const base = scratch();
  try {
    // Default alone: no list is written, the id lives in the folder.
    const first = profileUid({ base, id: 'default', folder: base });
    assert.match(first, /^[0-9a-f]{64}$/, '32 random bytes');
    assert.equal(fs.existsSync(path.join(base, 'profiles.json')), false);
    assert.equal(JSON.parse(fs.readFileSync(path.join(base, ID_FILE), 'utf8')).uid, first);
    assert.equal(profileUid({ base, id: 'default', folder: base }), first);

    const { profile: p } = makeProfile(base, { name: 'Work' });
    const folder = profileFolder(base, p);
    const uid = profileUid({ base, id: p.id, folder });
    assert.notEqual(uid, first);
    assert.equal(readProfiles(base).profiles.find((x) => x.id === p.id).uid, uid, 'the entry keeps it');
    renameProfile(base, p.id, 'Office work');
    assert.equal(profileUid({ base, id: p.id, folder }), uid, 'a rename changes nothing');
    assert.equal(readProfiles(base).profiles.find((x) => x.id === p.id).uid, uid);

    // The entry is believed when the two differ, and the folder's copy follows it.
    fs.writeFileSync(path.join(folder, ID_FILE), JSON.stringify({ uid: uidOf() }));
    assert.equal(profileUid({ base, id: p.id, folder }), uid);
    assert.equal(JSON.parse(fs.readFileSync(path.join(folder, ID_FILE), 'utf8')).uid, uid);
  } finally {
    gone(base);
  }
});

test('a fresh profile gets a random key wrapped with its id, and a password saved is read back, never in the clear', () => {
  const folder = scratch();
  try {
    const here = computer();
    const uid = uidOf();
    const ks = openProfileKeystore({ folder, uid, protector: here });
    assert.equal(ks.available(), true);
    assert.equal(ks.status().status, 'open');
    const rec = JSON.parse(fs.readFileSync(path.join(folder, KEY_FILE), 'utf8'));
    assert.equal(rec.by, 'system');
    assert.equal(rec.uid, uid);
    assert.equal(rec.key, undefined, 'the key itself is not on disk');
    const inner = JSON.parse(here.decrypt(Buffer.from(rec.wrapped, 'base64')));
    assert.equal(inner.uid, uid, 'the id is wrapped with the key');
    assert.equal(Buffer.from(inner.key, 'base64').length, 32);
    assert.ok(fs.existsSync(path.join(folder, KEYSTORE_DIR, storeNameOf(uid), 'master.json')), 'the store is named from the id');

    ks.set('mail:ann@example.test', 'correct horse battery');
    ks.set('dav:42', 'second secret');
    assert.equal(ks.get('mail:ann@example.test'), 'correct horse battery');
    assert.equal(ks.get('nothing'), null);
    const text = allText(folder);
    assert.equal(text.includes('correct horse battery'), false, 'no password in the clear');
    assert.equal(text.includes('ann@example.test'), false, 'no address in the clear: names are hidden too');

    // Opened again, as at the next start.
    const again = openProfileKeystore({ folder, uid, protector: here });
    assert.equal(again.get('mail:ann@example.test'), 'correct horse battery');
    assert.deepEqual(again.keys().sort(), ['dav:42', 'mail:ann@example.test']);
    again.delete('dav:42');
    assert.equal(openProfileKeystore({ folder, uid, protector: here }).get('dav:42'), null);
    assert.equal(ITERATIONS, 10000);
  } finally {
    gone(folder);
  }
});

test('a keystore taken to another computer will not open: it is put aside, said once, and a fresh one is started', () => {
  const folder = scratch();
  try {
    const uid = uidOf();
    openProfileKeystore({ folder, uid, protector: computer() }).set('mail:ann@example.test', 'secret one');
    const told = [];
    const there = openProfileKeystore({ folder, uid, protector: computer(), onSetAside: (a) => told.push(a) });
    assert.equal(told.length, 1);
    assert.equal(told[0].why, 'other-computer');
    assert.equal(there.status().aside.why, 'other-computer');
    assert.equal(there.available(), true, 'a fresh keystore, ready for the passwords typed again');
    assert.equal(there.get('mail:ann@example.test'), null, 'needs its password');
    const names = fs.readdirSync(folder);
    assert.ok(names.some((n) => n.startsWith(`${KEYSTORE_DIR}-set-aside-`)), 'the old store is kept, renamed');
    assert.ok(names.some((n) => n.startsWith('keystore-key-set-aside-')), 'and its key file');
    there.set('mail:ann@example.test', 'typed again');
    assert.equal(openProfileKeystore({ folder, uid, protector: computer() }).status().aside?.why, 'other-computer', 'and a third computer refuses that one too');
  } finally {
    gone(folder);
  }
});

test('a keystore copied into another profile will not open there, whichever part was copied', () => {
  const a = scratch();
  const b = scratch();
  try {
    const here = computer();
    const uidA = uidOf();
    const uidB = uidOf();
    openProfileKeystore({ folder: a, uid: uidA, protector: here }).set('talk:token', 'A only');

    // The whole of it, key file and store.
    fs.cpSync(path.join(a, KEY_FILE), path.join(b, KEY_FILE));
    fs.cpSync(path.join(a, KEYSTORE_DIR), path.join(b, KEYSTORE_DIR), { recursive: true });
    const told = [];
    const ks = openProfileKeystore({ folder: b, uid: uidB, protector: here, onSetAside: (x) => told.push(x.why) });
    assert.deepEqual(told, ['other-profile']);
    assert.equal(ks.get('talk:token'), null);
    assert.equal(ks.available(), true);
    gone(b);
    fs.mkdirSync(b);

    // The key file's id changed by hand to B's: the wrapped id still says A.
    const rec = JSON.parse(fs.readFileSync(path.join(a, KEY_FILE), 'utf8'));
    fs.writeFileSync(path.join(b, KEY_FILE), JSON.stringify({ ...rec, uid: uidB }));
    fs.cpSync(path.join(a, KEYSTORE_DIR), path.join(b, KEYSTORE_DIR), { recursive: true });
    const told2 = [];
    assert.equal(openProfileKeystore({ folder: b, uid: uidB, protector: here, onSetAside: (x) => told2.push(x.why) }).get('talk:token'), null);
    assert.deepEqual(told2, ['other-profile']);
    gone(b);
    fs.mkdirSync(b);

    // B's own key, A's store laid beside it.
    openProfileKeystore({ folder: b, uid: uidB, protector: here });
    gone(path.join(b, KEYSTORE_DIR));
    fs.cpSync(path.join(a, KEYSTORE_DIR), path.join(b, KEYSTORE_DIR), { recursive: true });
    const told3 = [];
    assert.equal(openProfileKeystore({ folder: b, uid: uidB, protector: here, onSetAside: (x) => told3.push(x.why) }).get('talk:token'), null);
    assert.deepEqual(told3, ['other-profile']);

    // A itself still opens.
    assert.equal(openProfileKeystore({ folder: a, uid: uidA, protector: here }).get('talk:token'), 'A only');
  } finally {
    gone(a);
    gone(b);
  }
});

test('where nothing protects a key and the profile has no password, nothing is saved, and the place saving one is told', () => {
  const folder = scratch();
  try {
    const ks = openProfileKeystore({ folder, uid: uidOf(), protector: computer(false) });
    assert.equal(ks.available(), false);
    assert.equal(ks.status().why, 'no-protection');
    assert.throws(() => ks.set('mail:x', 'y'), (err) => err.message === NOT_KEPT);
    assert.equal(ks.get('mail:x'), null);
    assert.deepEqual(fs.readdirSync(folder), [], 'no key and no store were written');
  } finally {
    gone(folder);
  }
});

test('a profile password protects the key where the system cannot, and a locked profile reads and sets aside nothing', () => {
  const folder = scratch();
  try {
    const uid = uidOf();
    const key = crypto.randomBytes(32);
    const ks = openProfileKeystore({ folder, uid, vault: createVault({ key }), protector: computer(false) });
    assert.equal(ks.available(), true);
    ks.set('db:engine', 'engine secret');
    const raw = fs.readFileSync(path.join(folder, KEY_FILE), 'utf8');
    assert.ok(isSealed(raw), 'the key file is sealed with the profile password');
    assert.equal(allText(folder).includes('engine secret'), false);
    assert.equal(openProfileKeystore({ folder, uid, vault: createVault({ key }), protector: computer(false) }).get('db:engine'), 'engine secret');

    const told = [];
    const locked = openProfileKeystore({ folder, uid, vault: createVault(), protector: computer(false), onSetAside: (x) => told.push(x) });
    assert.equal(locked.status().why, 'locked');
    assert.deepEqual(told, [], 'not set aside: the password opens it again');
    assert.ok(fs.existsSync(path.join(folder, KEY_FILE)));

    // The system's protection here later: the key is wrapped by it as well, still sealed.
    const here = computer();
    assert.equal(openProfileKeystore({ folder, uid, vault: createVault({ key }), protector: here }).get('db:engine'), 'engine secret');
    const rec = JSON.parse(createVault({ key }).readText(path.join(folder, KEY_FILE)));
    assert.equal(rec.by, 'system');
  } finally {
    gone(folder);
  }
});

test('setting, changing and taking away a profile password seals the key again, and taking it away where nothing else protects the key puts the store aside', async () => {
  const base = scratch();
  try {
    const here = computer();
    const { profile: p } = makeProfile(base, { name: 'Home' });
    const folder = profileFolder(base, p);
    const uid = profileUid({ base, id: p.id, folder });
    let vault = createVault();
    const ks = openProfileKeystore({ folder, uid, vault, protector: here });
    ks.set('mail:home', 'home secret');

    const one = await makeLock('first pass');
    await rekeyProfile({ base, id: p.id, vault, to: one.key, lock: one.lock, folder });
    assert.ok(isSealed(fs.readFileSync(path.join(folder, KEY_FILE), 'utf8')));
    ks.refresh();
    assert.equal(ks.get('mail:home'), 'home secret');

    const two = await makeLock('second pass');
    await rekeyProfile({ base, id: p.id, vault, from: one.key, to: two.key, lock: two.lock, folder });
    assert.equal(openProfileKeystore({ folder, uid, vault: createVault({ key: two.key }), protector: here }).get('mail:home'), 'home secret');
    assert.equal(openProfileKeystore({ folder, uid, vault: createVault({ key: one.key }), protector: here }).status().why, 'locked', 'the old password opens nothing');

    // Taken away with the system's protection here: the key stays wrapped, the store opens.
    await rekeyProfile({ base, id: p.id, vault, from: two.key, to: null, lock: null, folder });
    ks.beforePasswordGoes();
    ks.refresh();
    assert.equal(ks.get('mail:home'), 'home secret');
    assert.equal(isSealed(fs.readFileSync(path.join(folder, KEY_FILE), 'utf8')), false);

    // Where nothing protects it, a password set, then taken away.
    const bare = scratch();
    try {
      vault = createVault();
      const uid2 = uidOf();
      const three = await makeLock('third pass');
      await vault.rekey(three.key, { folder: bare });
      const told = [];
      const k2 = openProfileKeystore({ folder: bare, uid: uid2, vault, protector: computer(false), onSetAside: (x) => told.push(x.why) });
      k2.set('talk:token', 'only the password keeps this');
      k2.beforePasswordGoes();
      await vault.rekey(null, { folder: bare });
      k2.refresh();
      assert.deepEqual(told, ['unprotected']);
      assert.equal(k2.available(), false, 'no password and no protection: nothing is kept from now on');
      assert.equal(k2.get('talk:token'), null);
      const aside = fs.readdirSync(bare).find((n) => n.startsWith('keystore-key-set-aside-'));
      const asideText = fs.readFileSync(path.join(bare, aside), 'utf8');
      assert.ok(isSealed(asideText) || JSON.parse(asideText).key === null, 'the key is not left in the clear beside its store');
    } finally {
      gone(bare);
    }
  } finally {
    gone(base);
  }
});

test('the one-time move: secrets.json and plain settings into the keystore, read back, then the old copies gone', () => {
  const folder = scratch();
  try {
    const here = computer();
    const elsewhere = computer();
    const old = {
      'mail:ann@example.test': here.encrypt('mail pass').toString('base64'),
      'oauth:ann': here.encrypt('refresh token').toString('base64'),
      'dav:7': elsewhere.encrypt('from another computer').toString('base64'),
    };
    fs.writeFileSync(path.join(folder, OLD_SECRETS_FILE), JSON.stringify(old));
    const settings = settingsStub({ 'ai.tools.remoteKey': 'plain remote key', theme: 'dark' });
    const ks = openProfileKeystore({ folder, uid: uidOf(), protector: here });
    const r = moveOldSecrets({ folder, keystore: ks, protector: here, settings, plain: [['ai.tools.remoteKey', 'ai.tools.remoteKey']] });
    assert.deepEqual(r, { moved: 3, unreadable: 1, waiting: false });
    assert.equal(ks.get('mail:ann@example.test'), 'mail pass');
    assert.equal(ks.get('oauth:ann'), 'refresh token');
    assert.equal(ks.get('ai.tools.remoteKey'), 'plain remote key');
    assert.equal(ks.get('dav:7'), null, 'one this computer cannot open is asked for again');
    assert.equal(fs.existsSync(path.join(folder, OLD_SECRETS_FILE)), false);
    assert.equal(fs.existsSync(path.join(folder, MOVE_FILE)), false);
    assert.equal(settings.get('ai.tools.remoteKey', null), null, 'the plain copy is gone');
    assert.equal(settings.get('theme'), 'dark');
    assert.deepEqual(Object.keys(JSON.parse(fs.readFileSync(path.join(folder, UNREADABLE_FILE), 'utf8'))), ['dav:7'], 'kept aside, never read');
    assert.deepEqual(moveOldSecrets({ folder, keystore: ks, protector: here, settings, plain: [['ai.tools.remoteKey', 'ai.tools.remoteKey']] }), { moved: 0, unreadable: 0, waiting: false }, 'once only');
  } finally {
    gone(folder);
  }
});

test('a move stopped midway is taken up at the next open, and nothing is lost', () => {
  const folder = scratch();
  try {
    const here = computer();
    const old = {};
    for (let i = 0; i < 5; i++) old[`mail:${i}`] = here.encrypt(`pass ${i}`).toString('base64');
    fs.writeFileSync(path.join(folder, OLD_SECRETS_FILE), JSON.stringify(old));
    const uid = uidOf();
    const ks = openProfileKeystore({ folder, uid, protector: here });
    // Office stops after the second password is written.
    let writes = 0;
    const stopping = { ...ks, set: (n, v) => { if (writes++ === 2) throw new Error('power cut'); ks.set(n, v); } };
    assert.throws(() => moveOldSecrets({ folder, keystore: stopping, protector: here, settings: settingsStub() }), /power cut/);
    assert.ok(fs.existsSync(path.join(folder, OLD_SECRETS_FILE)), 'the old copies stay until every one is read back');
    assert.ok(fs.existsSync(path.join(folder, MOVE_FILE)), 'the note says a move is under way');

    const next = openProfileKeystore({ folder, uid, protector: here });
    assert.deepEqual(moveOldSecrets({ folder, keystore: next, protector: here, settings: settingsStub() }), { moved: 5, unreadable: 0, waiting: false });
    for (let i = 0; i < 5; i++) assert.equal(next.get(`mail:${i}`), `pass ${i}`);
    assert.equal(fs.existsSync(path.join(folder, OLD_SECRETS_FILE)), false);
    assert.equal(fs.existsSync(path.join(folder, MOVE_FILE)), false);
  } finally {
    gone(folder);
  }
});

test('a move waits while nothing can keep the passwords, and while a sealed secrets.json cannot be opened', () => {
  const folder = scratch();
  try {
    const here = computer();
    fs.writeFileSync(path.join(folder, OLD_SECRETS_FILE), JSON.stringify({ 'mail:a': here.encrypt('a').toString('base64') }));
    const off = openProfileKeystore({ folder, uid: uidOf(), protector: computer(false) });
    assert.equal(moveOldSecrets({ folder, keystore: off, protector: computer(false), settings: settingsStub() }).waiting, true);
    assert.ok(fs.existsSync(path.join(folder, OLD_SECRETS_FILE)), 'left where it is');

    // Sealed with a profile password not yet given.
    const key = crypto.randomBytes(32);
    const sealedVault = createVault({ key });
    sealedVault.writeJson(path.join(folder, OLD_SECRETS_FILE), { 'mail:a': here.encrypt('a').toString('base64') });
    const ks = openProfileKeystore({ folder: scratch(), uid: uidOf(), protector: here });
    assert.equal(moveOldSecrets({ folder, keystore: ks, vault: createVault(), protector: here, settings: settingsStub() }).waiting, true);
    const opened = moveOldSecrets({ folder, keystore: ks, vault: createVault({ key }), protector: here, settings: settingsStub() });
    assert.equal(opened.moved, 1);
    assert.equal(ks.get('mail:a'), 'a');
  } finally {
    gone(folder);
  }
});

test('a store being made when Office stopped is tidied away, and a write cut short is mended from the last good copy', () => {
  const folder = scratch();
  try {
    const here = computer();
    const uid = uidOf();
    const ks = openProfileKeystore({ folder, uid, protector: here });
    ks.set('mail:a', 'kept');
    // A seed file left by a stop (it holds a key in the clear for a moment).
    fs.writeFileSync(path.join(folder, KEYSTORE_DIR, '.seed-0123abcd.json'), '{"master":"00"}');
    const again = openProfileKeystore({ folder, uid, protector: here });
    assert.equal(fs.readdirSync(path.join(folder, KEYSTORE_DIR)).some((n) => n.startsWith('.seed-')), false);
    assert.equal(again.available(), true);

    again.set('mail:b', 'second');
    const master = path.join(folder, KEYSTORE_DIR, storeNameOf(uid), 'master.json');
    fs.writeFileSync(master, '{"broken');
    const mended = openProfileKeystore({ folder, uid, protector: here });
    assert.equal(mended.available(), true);
    assert.equal(mended.get('mail:b'), 'second');
  } finally {
    gone(folder);
  }
});

test('the AI tools never open the keystore', () => {
  const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
  // The AI tools' own process (run as Node by another program) is apps/desktop/main/mcp.
  const dir = path.join(root, 'apps', 'desktop', 'main', 'mcp');
  const walk = (d) => fs.readdirSync(d, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]));
  const files = walk(dir).filter((f) => /\.(c|m)?js$/.test(f));
  assert.ok(files.length > 5, 'its files are found');
  for (const f of files) {
    const text = fs.readFileSync(f, 'utf8');
    assert.equal(/profile-keystore|secure-keystore|keystore-key\.json/.test(text), false, `${path.relative(root, f)} reads no keystore`);
  }
});
