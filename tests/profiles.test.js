// Profiles (D-2026-10-26-02): the list beside the folder people already have,
// which profile a run is, the folders, the rules for removing one, and a
// password that seals a profile's files with a key made from it (never the
// password, never the key, on disk).
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  readProfiles, writeProfiles, defaultList, profileFolder, findProfile, profileArg, chooseProfile, makeProfile,
  renameProfile, setProfileColour, setOpenFirst, markUsed, setProfileLock, removalRefused, removeProfile,
  isRunning, folderSize, makeLock, unlock, seal, unseal, isSealed, createVault, PROFILES_FILE, DEFAULT_ID,
} from '../packages/office-shell/src/profiles.js';

const scratch = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-profiles-'));
const gone = (dir) => fs.rmSync(dir, { recursive: true, force: true });

test('with no list, the folder Office has always used is the one profile, Default, and reading or choosing writes nothing', () => {
  const base = scratch();
  try {
    fs.writeFileSync(path.join(base, 'settings.json'), '{"theme":"dark"}');
    const list = readProfiles(base);
    assert.deepEqual(list.profiles.map((p) => p.id), [DEFAULT_ID]);
    assert.equal(list.profiles[0].name, 'Default');
    assert.equal(profileFolder(base, list.profiles[0]), base, 'Default is the base folder itself: nothing is moved');
    const { profile, how } = chooseProfile({ base, argv: ['office.exe'] });
    assert.equal(profile.id, DEFAULT_ID);
    assert.equal(how, 'opens first');
    assert.equal(markUsed(base, DEFAULT_ID), false);
    assert.equal(fs.existsSync(path.join(base, PROFILES_FILE)), false, 'no list is written for Default alone');
    assert.deepEqual(fs.readdirSync(base), ['settings.json']);
  } finally {
    gone(base);
  }
});

test('a new profile gets a folder of its own under Profiles, a colour not in use, and a name no other profile has', () => {
  const base = scratch();
  try {
    const { profile: work } = makeProfile(base, { name: '  Work  ' });
    assert.equal(work.name, 'Work');
    assert.match(work.id, /^[0-9a-f]{8}$/);
    assert.equal(profileFolder(base, work), path.join(base, 'Profiles', work.id));
    assert.ok(fs.statSync(path.join(base, 'Profiles', work.id)).isDirectory());
    assert.notEqual(work.colour, readProfiles(base).profiles[0].colour, 'a colour the first does not wear');
    assert.throws(() => makeProfile(base, { name: 'work' }), /already exists/);
    assert.throws(() => makeProfile(base, { name: '   ' }), /Enter a name/);
    assert.throws(() => makeProfile(base, { name: 'x'.repeat(41) }), /up to 40/);
    const { profile: home } = makeProfile(base, { name: 'Home', colour: '#1A9F7A' });
    assert.equal(home.colour, '#1A9F7A');
    const list = readProfiles(base);
    assert.deepEqual(list.profiles.map((p) => p.name), ['Default', 'Work', 'Home']);
    assert.equal(list.openFirst, DEFAULT_ID);
  } finally {
    gone(base);
  }
});

test('a run is the profile its command line names, by id or name, either way of writing it; else the one marked to open first', () => {
  const base = scratch();
  try {
    const { profile: work } = makeProfile(base, { name: 'Work' });
    makeProfile(base, { name: 'Home' });
    assert.equal(chooseProfile({ base, argv: ['office.exe', `--profile=${work.id}`] }).profile.id, work.id);
    assert.equal(chooseProfile({ base, argv: ['office.exe', '--profile', 'work'] }).profile.id, work.id, 'by name, any case');
    assert.equal(chooseProfile({ base, argv: ['office.exe', '--profile="Home"'] }).profile.name, 'Home');
    const unknown = chooseProfile({ base, argv: ['office.exe', '--profile=nobody'] });
    assert.equal(unknown.profile.id, DEFAULT_ID);
    assert.equal(unknown.how, 'not found');
    setOpenFirst(base, work.id);
    assert.equal(chooseProfile({ base, argv: ['office.exe'] }).profile.id, work.id);
    assert.throws(() => setOpenFirst(base, 'nobody'), /no longer there/);
    assert.deepEqual(profileArg(['a', '--profile', 'Work', 'b.docx']), { value: 'Work', at: 2 });
    assert.deepEqual(profileArg(['a', '--profile', '--app=word']), { value: null, at: -1 }, 'a switch after it is not its value');
  } finally {
    gone(base);
  }
});

test('a file handed over with no profile named goes to the profile used last', () => {
  const base = scratch();
  try {
    const { profile: work } = makeProfile(base, { name: 'Work' });
    assert.equal(markUsed(base, work.id), true);
    assert.equal(markUsed(base, work.id), false, 'written only when it changes');
    assert.equal(chooseProfile({ base, argv: ['office.exe', 'C:\\x.docx'], handed: true }).profile.id, work.id);
    assert.equal(chooseProfile({ base, argv: ['office.exe'] }).profile.id, DEFAULT_ID, 'with nothing handed over, the one that opens first');
    assert.equal(chooseProfile({ base, argv: ['office.exe', '--profile=default', 'C:\\x.docx'], handed: true }).profile.id, DEFAULT_ID, 'a profile named wins');
  } finally {
    gone(base);
  }
});

test('a list someone edited by hand is made whole: Default put back, a bad colour mended, a folder never outside Profiles', () => {
  const base = scratch();
  try {
    writeProfiles(base, { profiles: [{ id: 'abc', name: 'Odd', colour: 'red', folder: '..\\..\\Windows' }, { id: 'abc', name: 'Twice' }, { id: '../x', name: 'Bad id' }], openFirst: 'gone' });
    const list = readProfiles(base);
    assert.deepEqual(list.profiles.map((p) => p.id), [DEFAULT_ID, 'abc']);
    assert.equal(list.profiles[1].colour, defaultList().profiles[0].colour);
    assert.equal(profileFolder(base, list.profiles[1]), path.join(base, 'Profiles', 'abc'));
    assert.equal(list.openFirst, DEFAULT_ID);
    fs.writeFileSync(path.join(base, PROFILES_FILE), 'not json');
    assert.deepEqual(readProfiles(base).profiles.map((p) => p.id), [DEFAULT_ID]);
  } finally {
    gone(base);
  }
});

test('renaming and colouring keep to the rules', () => {
  const base = scratch();
  try {
    const { profile: work } = makeProfile(base, { name: 'Work' });
    renameProfile(base, DEFAULT_ID, 'Personal');
    assert.equal(findProfile(readProfiles(base), 'personal').id, DEFAULT_ID);
    assert.throws(() => renameProfile(base, work.id, 'PERSONAL'), /already exists/);
    renameProfile(base, work.id, 'work');
    assert.equal(findProfile(readProfiles(base), work.id).name, 'work', 'its own name in another case is no clash');
    assert.throws(() => setProfileColour(base, work.id, 'blue'), /not a colour/);
    setProfileColour(base, work.id, '#C2408F');
    assert.equal(findProfile(readProfiles(base), work.id).colour, '#c2408f');
  } finally {
    gone(base);
  }
});

test('removing: never Default, never the last, never this window\'s or one that is open; otherwise its folder goes and it leaves the list', async () => {
  const base = scratch();
  try {
    assert.match(removalRefused(base, DEFAULT_ID), /last profile/);
    const { profile: work } = makeProfile(base, { name: 'Work' });
    fs.writeFileSync(path.join(base, 'Profiles', work.id, 'settings.json'), '{"a":1}');
    setOpenFirst(base, work.id);
    markUsed(base, work.id);
    assert.match(removalRefused(base, DEFAULT_ID), /first profile keeps what every profile shares/);
    assert.match(removalRefused(base, work.id, { current: work.id }), /This window belongs to that profile/);
    assert.match(removalRefused(base, work.id, { running: () => true }), /Close its windows first/);
    await assert.rejects(removeProfile(base, work.id, { running: () => true, discard: async () => assert.fail('nothing goes') }), /Close its windows first/);
    const discarded = [];
    const list = await removeProfile(base, work.id, { running: () => false, discard: async (f) => { discarded.push(f); fs.rmSync(f, { recursive: true }); } });
    assert.deepEqual(discarded, [path.join(base, 'Profiles', work.id)]);
    assert.deepEqual(list.profiles.map((p) => p.id), [DEFAULT_ID]);
    assert.equal(list.openFirst, DEFAULT_ID);
    assert.equal(list.lastUsed, DEFAULT_ID);
    assert.ok(fs.existsSync(base), 'the base folder stays');
  } finally {
    gone(base);
  }
});

test('a profile is open while a running copy holds its lock file, and not when there is none', () => {
  const base = scratch();
  try {
    assert.equal(isRunning(base), false);
    // Left behind and held by nobody: not running.
    fs.writeFileSync(path.join(base, 'lockfile'), '');
    if (process.platform === 'win32') assert.equal(isRunning(base), false);
    fs.mkdirSync(path.join(base, 'sub'));
    fs.writeFileSync(path.join(base, 'sub', 'a.bin'), Buffer.alloc(1000));
    assert.equal(folderSize(base), 1000);
  } finally {
    gone(base);
  }
});

test('a password makes a lock that tells the right password from a wrong one, and the key it gives is never written', async () => {
  await assert.rejects(makeLock('abc'), /at least 4 characters/);
  const { lock, key } = await makeLock('correct horse');
  assert.equal(key.length, 32);
  assert.equal(lock.kdf, 'scrypt');
  const again = await unlock(lock, 'correct horse');
  assert.ok(again && again.equals(key), 'the same key from the same password');
  assert.equal(await unlock(lock, 'correct horse '), null);
  assert.equal(await unlock(lock, 'Correct horse'), null);
  assert.equal(await unlock(null, 'correct horse'), null);
  const base = scratch();
  try {
    setProfileLock(base, DEFAULT_ID, lock);
    const text = fs.readFileSync(path.join(base, PROFILES_FILE), 'utf8');
    assert.ok(!text.includes('correct horse'), 'never the password');
    assert.ok(!text.includes(key.toString('base64')) && !text.includes(key.toString('hex')), 'never the key');
    assert.ok(readProfiles(base).profiles[0].lock.verifier);
    setProfileLock(base, DEFAULT_ID, null);
    assert.equal(readProfiles(base).profiles[0].lock, null);
  } finally {
    gone(base);
  }
});

test('a sealed text opens only with its key, and shows nothing of what it holds', async () => {
  const { key } = await makeLock('first password');
  const { key: other } = await makeLock('second password');
  const sealed = seal(key, '{"mail:1:password":"hunter2"}');
  assert.ok(isSealed(sealed));
  assert.ok(!sealed.includes('hunter2') && !sealed.includes('mail:1'));
  assert.notEqual(seal(key, 'same'), seal(key, 'same'), 'a fresh nonce every time');
  assert.equal(unseal(key, sealed), '{"mail:1:password":"hunter2"}');
  assert.throws(() => unseal(other, sealed));
  assert.throws(() => unseal(null, sealed), /locked/);
});

test('saved passwords round trip through the vault: sealed on disk with a key, read back with it, plain again once the password goes', async () => {
  const dir = scratch();
  try {
    const secrets = path.join(dir, 'secrets.json');
    const talk = path.join(dir, 'talk', 'me');
    fs.mkdirSync(talk, { recursive: true });
    const plain = createVault();
    plain.cover(secrets);
    plain.cover(path.join(dir, 'talk'));
    plain.writeJson(secrets, { 'mail:1:password': 'AQAAANCMnd8BFdERjHoAwE' });
    fs.writeFileSync(path.join(talk, 'c1.json'), JSON.stringify([{ id: 1, text: 'see you at nine' }]));
    assert.equal(plain.sealed(), false);
    assert.ok(fs.readFileSync(secrets, 'utf8').includes('mail:1:password'));

    // A password set: every file it covers sealed, and still readable with the key.
    const { key } = await makeLock('open sesame');
    assert.deepEqual(await plain.rekey(key), { files: 2 });
    assert.ok(isSealed(fs.readFileSync(secrets, 'utf8')));
    assert.ok(!fs.readFileSync(path.join(talk, 'c1.json'), 'utf8').includes('nine'));
    const opened = createVault({ key });
    assert.deepEqual(opened.readJson(secrets, {}), { 'mail:1:password': 'AQAAANCMnd8BFdERjHoAwE' });
    opened.writeJson(secrets, { a: 'b' });
    assert.ok(isSealed(fs.readFileSync(secrets, 'utf8')), 'written sealed while it has a key');

    // Without the key, nothing is read and nothing is written over.
    const locked = createVault();
    assert.deepEqual(locked.readJson(secrets, { fallback: true }), { fallback: true });
    assert.throws(() => locked.writeJson(secrets, {}), /locked/);
    await assert.rejects(locked.writeJsonLater(secrets, {}), /locked/);
    assert.deepEqual(opened.readJson(secrets, {}), { a: 'b' }, 'still there');

    // Changed, then taken away.
    const { key: next } = await makeLock('new sesame');
    opened.cover(secrets);
    opened.cover(path.join(dir, 'talk'));
    await opened.rekey(next);
    assert.deepEqual(createVault({ key: next }).readJson(secrets, {}), { a: 'b' });
    assert.deepEqual(createVault({ key }).readJson(secrets, null), null, 'the old key no longer opens it');
    await opened.rekey(null);
    assert.deepEqual(JSON.parse(fs.readFileSync(secrets, 'utf8')), { a: 'b' });
    assert.deepEqual(JSON.parse(fs.readFileSync(path.join(talk, 'c1.json'), 'utf8')), [{ id: 1, text: 'see you at nine' }]);
  } finally {
    gone(dir);
  }
});

test('a write still on its way when the password changes lands before the files are sealed again, never after', async () => {
  const dir = scratch();
  try {
    const file = path.join(dir, 'calendar.json');
    const vault = createVault();
    vault.cover(file);
    vault.writeJson(file, { n: 0 });
    const { key } = await makeLock('timing test');
    const writes = [];
    for (let n = 1; n <= 20; n++) writes.push(vault.writeJsonLater(file, { n }));
    const rekeyed = vault.rekey(key);
    writes.push(vault.writeJsonLater(file, { n: 21 }));
    await Promise.all([...writes, rekeyed]);
    assert.ok(isSealed(fs.readFileSync(file, 'utf8')));
    assert.deepEqual(createVault({ key }).readJson(file, null), { n: 21 });
  } finally {
    gone(dir);
  }
});
