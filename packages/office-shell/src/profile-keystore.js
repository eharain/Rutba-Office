// Each profile's saved passwords, in a keystore of its own (D-2026-10-26-09:
// "for a profile we generate a long id in profile, and stores an id linked
// key in windows protection. if the keystore is stolen of moved it become
// practically unusable and the user should enter the passwords again").
//
// Three things, in the profile's folder:
//   profile-id.json     the profile's long id (32 random bytes), made once and
//                       never changed; profiles.json's entry keeps it too, and
//                       is the one believed when the two differ
//   keystore-key.json   the keystore's key: 32 random bytes, wrapped by the
//                       system's protection (DPAPI on Windows, the Keychain on
//                       macOS, the secret service on Linux) together with the
//                       id, and checked against the id when unwrapped. With a
//                       profile password the file is sealed with it as well
//                       (profiles.js, the vault), so a password change seals
//                       it again with the rest. Where the system gives no real
//                       protection, the key is kept only in a profile with a
//                       password, sealed by it; without one, nothing is saved.
//   keystore\           @tech-style/secure-keystore's directory: one store,
//                       named from the id, opened with a key made from the
//                       random key and the id (HKDF), so the right key with
//                       another profile's id opens nothing. Each entry's name
//                       is a keyed hash of what it is for, so the names on
//                       disk say nothing (an address, a server), and its value
//                       carries the real name. A marker entry holds the id.
//
// The store is read once as the profile opens and kept in memory; a password
// saved is written through at once. The library stretches its key once per
// call (PBKDF2); the key here is random, not a password, so the library's
// floor (10,000 rounds) is used: at its default, 600,000, every saved password
// would cost more than a second (measured: open 2.7 s and each write 2.4 s,
// against 135 ms and 46 ms at 10,000).
//
// A keystore that will not open here, because its key was wrapped on another
// computer or Windows account, or it belongs to another profile, or its key
// is gone, is put aside (renamed, never read again, never deleted at once),
// a new one is started, and the profile is told once (`setAside`) that its
// saved passwords must be typed again. Each place that needs one then asks.
//
// Passwords saved before this keystore (secrets.json, each under the
// system's protection, and the few kept plain in settings.json) are moved in
// once (moveOldSecrets): a note first, each copied, the store read back from
// disk and compared, and only then the old copies removed; a stop midway is
// taken up again at the next open.
//
// Nothing here needs Electron: the system's protection is handed in
// (`protector`: available, encrypt, decrypt), so the engine tests stand in
// for another computer. The AI tools' own process never opens this.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import keystoreLib from '@tech-style/secure-keystore';
import { readProfiles, writeProfiles, isSealed, PROFILES_FILE } from './profiles.js';

const { SecureKeystore, ErrorCodes } = keystoreLib;

export const ID_FILE = 'profile-id.json';
export const KEY_FILE = 'keystore-key.json';
export const KEYSTORE_DIR = 'keystore';
export const MOVE_FILE = 'keystore-move.json';
export const OLD_SECRETS_FILE = 'secrets.json';
export const UNREADABLE_FILE = 'secrets-unreadable.json';
/** The library's floor: the key is random, so stretching it buys nothing (see above). */
export const ITERATIONS = 10_000;
const MARKER = '__profile';
const GOOD = 'master.good.json';
const SEED = /^\.seed-[0-9a-f]+\.json$/;

/** What a place that saves a password is told when nothing here can keep it. */
export const NOT_KEPT = 'This computer has no protected store for saved passwords, so this one cannot be kept. Give this profile a password (Profiles) to keep passwords here, or type it each time it is asked for.';

const isUid = (s) => typeof s === 'string' && /^[0-9a-f]{64}$/.test(s);
const readPlain = (file) => {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return null;
  }
};
const writePlain = (file, value) => {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(value, null, 2));
  fs.renameSync(tmp, file);
};

/**
 * A profile's long id: profiles.json's entry, else the profile folder's
 * copy, else made now (32 random bytes). Kept in both places once made; the
 * entry is not written while profiles.json does not exist (Default alone).
 */
export function profileUid({ base, id, folder }) {
  let list = null;
  let entry = null;
  if (base && fs.existsSync(path.join(base, PROFILES_FILE))) {
    list = readProfiles(base);
    entry = list.profiles.find((p) => p.id === id) || null;
  }
  const fromEntry = isUid(entry?.uid) ? entry.uid : null;
  const kept = readPlain(path.join(folder, ID_FILE));
  const fromFolder = isUid(kept?.uid) ? kept.uid : null;
  const uid = fromEntry || fromFolder || crypto.randomBytes(32).toString('hex');
  if (fromFolder !== uid) writePlain(path.join(folder, ID_FILE), { uid, made: kept?.made && fromFolder === uid ? kept.made : new Date().toISOString() });
  if (entry && fromEntry !== uid) {
    entry.uid = uid;
    writeProfiles(base, list);
  }
  return uid;
}

/** The store's name: from the id, so another profile's store is not this one's. */
export const storeNameOf = (uid) => `profile-${crypto.createHash('sha256').update(String(uid)).digest('hex').slice(0, 32)}`;

/** The key the library is given, and the key the names are hashed with, both from the random key and the id. */
function derive(key, uid) {
  const out = Buffer.from(crypto.hkdfSync('sha256', key, Buffer.from(uid, 'hex'), 'rutba-office profile keystore 1', 64));
  return { userKey: out.subarray(0, 32).toString('hex'), names: out.subarray(32) };
}

const storesIn = (dir) => {
  try {
    return fs.readdirSync(dir).filter((n) => fs.existsSync(path.join(dir, n, 'master.json')));
  } catch {
    return [];
  }
};

/**
 * @param {object} o
 * @param {string} o.folder      the profile's folder
 * @param {string} o.uid         its long id (profileUid)
 * @param {object} [o.vault]     the profile's vault (profiles.js): seals the key file with the profile's password
 * @param {object} [o.protector] the system's protection: { available(), encrypt(text) -> Buffer, decrypt(Buffer) -> text }
 * @param {(aside: { at, why }) => void} [o.onSetAside]  told when a keystore that would not open is put aside
 */
export function openProfileKeystore({ folder, uid, vault = null, protector = null, onSetAside = null, now = () => Date.now(), iterations = ITERATIONS }) {
  if (!isUid(uid)) throw new Error('A profile keystore needs the profile\'s id.');
  const dir = path.join(folder, KEYSTORE_DIR);
  const keyFile = path.join(folder, KEY_FILE);
  const store = storeNameOf(uid);
  vault?.cover(keyFile);

  const cache = new Map();
  let lib = null;
  let userKey = null;
  let names = null;
  let state = { status: 'closed', why: null, aside: null };

  const protects = () => {
    try {
      return Boolean(protector?.available());
    } catch {
      return false;
    }
  };
  const passwordOn = () => Boolean(vault?.sealed());
  const readRecord = () => (vault ? vault.readJson(keyFile, null) : readPlain(keyFile));
  const writeRecord = (rec) => (vault ? vault.writeJson(keyFile, rec) : writePlain(keyFile, rec));
  const hidden = (name) => `k${crypto.createHmac('sha256', names).update(String(name)).digest('hex').slice(0, 40)}`;

  /** The key's record: wrapped by the system with the id, or, with nothing to wrap it, kept for the profile's password to seal. */
  const recordFor = (key) => (protects()
    ? { v: 1, uid, by: 'system', wrapped: protector.encrypt(JSON.stringify({ uid, key: key.toString('base64') })).toString('base64'), made: new Date(now()).toISOString() }
    : { v: 1, uid, by: 'password', key: key.toString('base64'), made: new Date(now()).toISOString() });

  /** The store as it stood after the last write that finished, for a write a crash cut short (the library writes in place). */
  const keepGood = () => {
    const master = path.join(dir, store, 'master.json');
    const tmp = path.join(dir, store, `${GOOD}.${process.pid}.tmp`);
    try {
      fs.copyFileSync(master, tmp);
      fs.renameSync(tmp, path.join(dir, store, GOOD));
    } catch {
      /* the next write keeps it */
    }
  };

  /** Put aside, as it is: renamed, never read again here, never deleted at once. */
  const setAside = (why) => {
    const at = new Date(now()).toISOString();
    const stamp = at.replace(/[:.]/g, '-');
    try {
      if (fs.existsSync(dir)) fs.renameSync(dir, `${dir}-set-aside-${stamp}`);
    } catch {
      /* left where it is; a new one is not started over it below */
    }
    try {
      if (fs.existsSync(keyFile)) {
        const aside = path.join(folder, `keystore-key-set-aside-${stamp}.json.aside`);
        fs.renameSync(keyFile, aside);
        // A key that lay plain (its password taken away where nothing protects it) is not kept beside its store.
        if (why === 'unprotected') {
          const raw = readPlain(aside);
          if (raw && raw.key) writePlain(aside, { ...raw, key: null });
        }
      }
    } catch {
      /* as above */
    }
    lib = null;
    cache.clear();
    state = { status: 'closed', why, aside: { at, why } };
    onSetAside?.({ at, why });
  };

  /** The store made new, with its marker: the library makes a store from a file, which goes at once. */
  const create = () => {
    fs.mkdirSync(dir, { recursive: true });
    const seed = path.join(dir, `.seed-${crypto.randomBytes(8).toString('hex')}.json`);
    fs.writeFileSync(seed, JSON.stringify({ master: crypto.randomBytes(32).toString('hex'), [hidden(MARKER)]: JSON.stringify({ n: MARKER, v: uid }) }), { flag: 'wx', mode: 0o600 });
    try {
      lib.secure(seed, userKey, store);
    } finally {
      fs.rmSync(seed, { force: true });
    }
    keepGood();
  };

  /** A store being made when Office stopped holds nothing but its marker: it goes, with the file it was made from. */
  const tidySeeds = () => {
    let left = [];
    try {
      left = fs.readdirSync(dir).filter((n) => SEED.test(n));
    } catch {
      return;
    }
    if (!left.length) return;
    for (const n of left) fs.rmSync(path.join(dir, n), { force: true });
    fs.rmSync(path.join(dir, store), { recursive: true, force: true });
  };

  const readAll = () => {
    try {
      return lib.getAll(userKey, store);
    } catch (err) {
      if (err?.code === ErrorCodes.STORE_CORRUPT && fs.existsSync(path.join(dir, store, GOOD))) {
        fs.copyFileSync(path.join(dir, store, GOOD), path.join(dir, store, 'master.json'));
        return readAll();
      }
      if (err?.code !== ErrorCodes.BAD_USER_KEY && err?.code !== ErrorCodes.STORE_CORRUPT) throw err;
      // One value a crash cut short: the rest read one by one, that one lost (its place asks for it again).
      let marker;
      try {
        marker = lib.get(hidden(MARKER), userKey, store);
      } catch {
        throw err;
      }
      const out = { [hidden(MARKER)]: marker };
      for (const k of lib.listKeys(store)) {
        if (k in out) continue;
        try {
          out[k] = lib.get(k, userKey, store);
        } catch {
          /* lost */
        }
      }
      return out;
    }
  };

  /** Opened with this key: read once, kept in memory. Answers whether it opened. */
  const attach = (key) => {
    ({ userKey, names } = derive(key, uid));
    lib = new SecureKeystore({ keystoreDir: dir, iterations, logger: null });
    tidySeeds();
    if (!fs.existsSync(path.join(dir, store, 'master.json'))) {
      if (storesIn(dir).length) {
        setAside('other-profile');
        return false;
      }
      create();
    }
    let values;
    try {
      values = readAll();
    } catch (err) {
      setAside(err?.code === ErrorCodes.BAD_USER_KEY ? 'other-profile' : 'unreadable');
      return false;
    }
    let mark = null;
    try {
      mark = JSON.parse(values[hidden(MARKER)] || 'null');
    } catch {
      mark = null;
    }
    if (mark?.v !== uid) {
      setAside('other-profile');
      return false;
    }
    cache.clear();
    for (const [k, raw] of Object.entries(values)) {
      if (k === hidden(MARKER)) continue;
      try {
        const e = JSON.parse(raw);
        if (e && typeof e.n === 'string' && hidden(e.n) === k) cache.set(e.n, String(e.v ?? ''));
      } catch {
        /* not one of ours */
      }
    }
    state = { status: 'open', why: null, aside: state.aside };
    return true;
  };

  /** A new keystore, if anything here can protect its key. */
  const start = () => {
    if (!protects() && !passwordOn()) {
      state = { status: 'closed', why: 'no-protection', aside: state.aside };
      return;
    }
    const key = crypto.randomBytes(32);
    writeRecord(recordFor(key));
    attach(key);
  };

  /** As the profile opens: its keystore opened, or put aside and a new one started. */
  const open = () => {
    lib = null;
    cache.clear();
    const rec = readRecord();
    if (!rec) {
      // Sealed with a key this run does not have: left as it is, nothing read or saved.
      let raw = null;
      try { raw = fs.readFileSync(keyFile, 'utf8'); } catch { raw = null; }
      if (raw !== null && isSealed(raw)) {
        state = { status: 'closed', why: 'locked', aside: state.aside };
        return undefined;
      }
      if (raw !== null || storesIn(dir).length) setAside(raw !== null ? 'unreadable' : 'no-key');
      return start();
    }
    if (rec.uid !== uid) {
      setAside('other-profile');
      return start();
    }
    let key = null;
    if (rec.wrapped) {
      // The system's protection away for now (a Linux keyring locked): nothing is put aside, nothing saved.
      if (!protects()) {
        state = { status: 'closed', why: 'no-protection', aside: state.aside };
        return undefined;
      }
      try {
        const inner = JSON.parse(protector.decrypt(Buffer.from(rec.wrapped, 'base64')));
        if (inner?.uid !== uid) {
          setAside('other-profile');
          return start();
        }
        key = Buffer.from(inner.key, 'base64');
      } catch {
        setAside('other-computer');
        return start();
      }
    } else if (rec.key) {
      if (!passwordOn()) {
        setAside('unprotected');
        return start();
      }
      key = Buffer.from(rec.key, 'base64');
      // The system's protection here now: the key wrapped by it as well.
      if (protects()) writeRecord(recordFor(key));
    } else {
      setAside('unreadable');
      return start();
    }
    if (!attach(key)) return start();
    return undefined;
  };

  open();

  const api = {
    /** Whether a password saved now is kept. */
    available: () => state.status === 'open',
    get: (name) => (cache.has(String(name)) ? cache.get(String(name)) : null),
    set: (name, value) => {
      if (state.status !== 'open') throw new Error(NOT_KEPT);
      const n = String(name);
      const v = String(value ?? '');
      lib.set(hidden(n), JSON.stringify({ n, v }), userKey, store);
      keepGood();
      cache.set(n, v);
    },
    delete: (name) => {
      const n = String(name);
      if (state.status === 'open' && cache.has(n)) {
        try {
          lib.deleteKey(hidden(n), store);
          keepGood();
        } catch (err) {
          if (err?.code !== ErrorCodes.KEY_NOT_FOUND) throw err;
        }
      }
      cache.delete(n);
    },
    keys: () => [...cache.keys()],
    /** { status: 'open' | 'closed', why, aside }: why it is closed, and when one was put aside this run. */
    status: () => ({ ...state }),
    /** Opened again, as at the next start: after the profile's password changes, and in the window checks. */
    refresh: () => {
      open();
      return api.status();
    },
    /**
     * The profile's password about to be taken away: a key that only it
     * protected (nothing else here can) would lie plain, so the keystore is
     * put aside first and passwords are not kept from then on.
     */
    beforePasswordGoes: () => {
      const rec = readRecord();
      if (rec && !rec.wrapped && rec.key && !protects()) setAside('unprotected');
    },
    /** These names, read again from disk, hold what is kept in memory: the check before old copies go. */
    matchesDisk(list) {
      if (state.status !== 'open') return false;
      const fresh = new SecureKeystore({ keystoreDir: dir, iterations, logger: null, createDir: false }).getAll(userKey, store);
      return list.every((n) => {
        try {
          const e = JSON.parse(fresh[hidden(n)] || 'null');
          return e?.n === n && e.v === cache.get(n);
        } catch {
          return false;
        }
      });
    },
  };
  return api;
}

/**
 * Passwords kept the old way, moved into the profile's keystore once: every
 * entry of secrets.json (each wrapped by the system's protection) and each
 * setting named in `plain` ([settingKey, name]), copied in, read back from
 * disk and compared, and only then taken from the old place. A note
 * (keystore-move.json) is written first; a stop midway is taken up at the
 * next open, and what is already in the keystore is not written again.
 * Entries this computer cannot unwrap (saved on another) are kept aside in
 * secrets-unreadable.json, never read again. Answers { moved, unreadable, waiting }.
 */
export function moveOldSecrets({ folder, keystore, vault = null, protector = null, settings = null, plain = [] }) {
  const file = path.join(folder, OLD_SECRETS_FILE);
  const note = path.join(folder, MOVE_FILE);
  // A file there that cannot be opened (sealed with a key this run does not have) is left alone.
  let text = null;
  try { text = vault ? vault.readText(file) : fs.readFileSync(file, 'utf8'); } catch { text = null; }
  if (text === null && fs.existsSync(file)) return { moved: 0, unreadable: 0, waiting: true };
  let old = {};
  try { old = text ? JSON.parse(text) || {} : {}; } catch { old = {}; }
  const plainOnes = plain.filter(([k]) => {
    const v = settings?.get(k, null);
    return typeof v === 'string' && v !== '';
  });
  const names = Object.keys(old);
  if (!names.length && !plainOnes.length) {
    fs.rmSync(note, { force: true });
    fs.rmSync(file, { force: true });
    return { moved: 0, unreadable: 0, waiting: false };
  }
  // Nothing here can keep them yet: they stay where they are until something can.
  if (!keystore.available()) return { moved: 0, unreadable: 0, waiting: true };
  writePlain(note, { started: new Date().toISOString(), count: names.length + plainOnes.length });
  const unreadable = {};
  const moved = [];
  for (const name of names) {
    if (keystore.get(name) !== null) {
      moved.push(name);
      continue;
    }
    let value = null;
    try {
      value = protector?.available() ? protector.decrypt(Buffer.from(String(old[name]), 'base64')) : null;
    } catch {
      value = null;
    }
    if (value === null) {
      unreadable[name] = old[name];
      continue;
    }
    keystore.set(name, value);
    moved.push(name);
  }
  for (const [k, name] of plainOnes) {
    if (keystore.get(name) === null) keystore.set(name, settings.get(k));
    moved.push(name);
  }
  if (!keystore.matchesDisk(moved)) throw new Error('The saved passwords could not be read back from the profile\'s keystore, so they were left where they were.');
  // Read back and the same: the old copies go.
  if (Object.keys(unreadable).length) {
    const kept = (vault ? vault.readJson(path.join(folder, UNREADABLE_FILE), {}) : readPlain(path.join(folder, UNREADABLE_FILE))) || {};
    const all = { ...kept, ...unreadable };
    if (vault) vault.writeJson(path.join(folder, UNREADABLE_FILE), all);
    else writePlain(path.join(folder, UNREADABLE_FILE), all);
  }
  fs.rmSync(file, { force: true });
  for (const [k] of plainOnes) settings.delete(k);
  fs.rmSync(note, { force: true });
  return { moved: moved.length, unreadable: Object.keys(unreadable).length, waiting: false };
}
