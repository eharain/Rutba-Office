// The Locker's keys, and what opens them.
//
// One random key (R) is the locker's root, kept wrapped in keyring.json in
// Office's own profile folder: one record per way of opening it. The password
// is the master and works on any computer; the recovery key works once a
// person has lost the password; Windows Hello, on this computer only, is
// another record that a sealed secret from the system opens. A wrong
// password opens nothing and changes nothing: the only write here is the
// keyring itself, swapped in whole, and only after every change has been
// worked out.
//
// Each locked folder has its own random key (its folder key). It is kept in
// two places, and either one is enough to open the folder:
//   - in the profile, in a @tech-style/secure-keystore store unlocked by a key
//     derived from R (this file; the store is the library's own, not ours);
//   - in the folder itself, in `.rutbalocker-keys`, wrapped by the password
//     and the recovery key, so the folder opens on another computer with
//     the password alone. That file holds nothing a thief can use without
//     them: it is encrypted key material only, as the synced folder must be.
//
// The algorithms are node:crypto's: scrypt for the password, HKDF-SHA256 and
// AES-256-GCM to wrap keys. A wrong password fails the GCM tag; nothing is
// compared by hand.

import crypto from 'node:crypto';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { promisify } from 'node:util';
import keystoreLib from '@rutba/office-shell/keystore-lib';
import { LockerError, checkVersion } from './format.js';

const { SecureKeystore, ErrorCodes } = keystoreLib;
const scryptAsync = promisify(crypto.scrypt);

export const KEY_LEN = 32;
/** The cost of a password, as OWASP suggests for scrypt: 128 MB. Recorded with what it derived, so it can rise. */
export const SCRYPT_DEFAULT = { N: 2 ** 17, r: 8, p: 1 };
const LIMITS = { N: 2 ** 20, r: 16, p: 4 };
/** At most this much memory for one password, whatever a file asks for: a file someone sent must not be able to take the computer's. */
const MOST_MEMORY = 2 ** 30;
/** The most entries of one kind tried in a folder's key file. */
const MOST_TRIED = 8;

/** What tells two entries of a key file apart: each locker's password and recovery key have their own salt. */
const saltOf = (u) => (u?.kind === 'password' ? u.kdf?.salt : u?.salt) || null;
const STORE = 'locker';
const CONTEXT = 'rutba-locker:v1:';
const KEYRING_FILE = 'keyring.json';
const KEYSTORE_DIR = 'keystore';

/** A folder's key file name, in the folder's root. */
export const FOLDER_KEYS = '.rutbalocker-keys';

const toHex = (b) => Buffer.from(b).toString('hex');
const fromHex = (s) => Buffer.from(String(s), 'hex');

/** A password as bytes: the same characters always the same bytes, however the keyboard typed them. */
export const passwordBytes = (password) => Buffer.from(String(password).normalize('NFKC'), 'utf8');

function scryptParams(kdf) {
  const N = Number(kdf?.N);
  const r = Number(kdf?.r);
  const p = Number(kdf?.p);
  if (!Number.isInteger(N) || !Number.isInteger(r) || !Number.isInteger(p) || N < 2 || N > LIMITS.N || r < 1 || r > LIMITS.r || p < 1 || p > LIMITS.p || !Number.isInteger(Math.log2(N)) || 128 * N * r > MOST_MEMORY) {
    throw new LockerError('format', 'This file asks for settings Rutba Office does not use.');
  }
  return { N, r, p, maxmem: 128 * N * r + 1024 * 1024 };
}

/** A password stretched with scrypt (off the main thread) into a 32-byte key. */
export async function passwordKey(password, kdf) {
  const { N, r, p, maxmem } = scryptParams(kdf);
  return Buffer.from(await scryptAsync(passwordBytes(password), fromHex(kdf.salt), KEY_LEN, { N, r, p, maxmem }));
}

/** The settings for a new password: a fresh salt, at today's cost. */
export const newKdf = (cost = SCRYPT_DEFAULT) => ({ alg: 'scrypt', ...cost, salt: crypto.randomBytes(32).toString('hex') });

/** A high-entropy secret (a recovery key, Hello's sealed secret) made into a wrapping key: HKDF, no stretching needed. */
export const secretKey = (secret, salt, info) => Buffer.from(crypto.hkdfSync('sha256', secret, fromHex(salt), Buffer.from(`${CONTEXT}${info}`, 'utf8'), KEY_LEN));

/** A key sealed by another key, bound to what it is for. */
export function wrap(kek, data, context) {
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv('aes-256-gcm', kek, iv);
  c.setAAD(Buffer.from(`${CONTEXT}wrap:${context}`, 'utf8'));
  const body = Buffer.concat([c.update(data), c.final()]);
  return { iv: toHex(iv), tag: toHex(c.getAuthTag()), data: toHex(body) };
}

/** A sealed key opened, or null when this key does not open it. */
export function unwrap(kek, rec, context) {
  try {
    const d = crypto.createDecipheriv('aes-256-gcm', kek, fromHex(rec.iv));
    d.setAAD(Buffer.from(`${CONTEXT}wrap:${context}`, 'utf8'));
    d.setAuthTag(fromHex(rec.tag));
    return Buffer.concat([d.update(fromHex(rec.data)), d.final()]);
  } catch {
    return null;
  }
}

/** A recovery key: 32 random bytes, written in groups a person can copy down. */
export function newRecoveryKey() {
  const bytes = crypto.randomBytes(32);
  return { bytes, text: base32(bytes).match(/.{1,4}/g).join('-') };
}

const BASE32 = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

function base32(bytes) {
  let bits = 0;
  let val = 0;
  let out = '';
  for (const b of bytes) {
    val = (val << 8) | b;
    bits += 8;
    while (bits >= 5) {
      out += BASE32[(val >>> (bits - 5)) & 31];
      bits -= 5;
    }
  }
  if (bits) out += BASE32[(val << (5 - bits)) & 31];
  return out;
}

/** A recovery key as a person typed it: separators and case do not matter. */
export function recoveryBytes(text) {
  const clean = String(text).toUpperCase().replace(/[^A-Z2-9]/g, '');
  let bits = 0;
  let val = 0;
  const out = [];
  for (const ch of clean) {
    const i = BASE32.indexOf(ch);
    if (i < 0) throw new LockerError('bad-recovery', 'That is not a recovery key.');
    val = (val << 5) | i;
    bits += 5;
    if (bits >= 8) {
      out.push((val >>> (bits - 8)) & 255);
      bits -= 8;
    }
  }
  if (out.length !== 32) throw new LockerError('bad-recovery', 'That is not a recovery key.');
  return Buffer.from(out);
}

/** A small JSON file written whole: to a temporary name, flushed, then renamed over the old one. */
export async function writeJsonAtomic(file, value) {
  await fsp.mkdir(path.dirname(file), { recursive: true });
  const tmp = `${file}.${process.pid}.${crypto.randomBytes(4).toString('hex')}.tmp`;
  const fh = await fsp.open(tmp, 'wx', 0o600);
  try {
    await fh.writeFile(JSON.stringify(value, null, 2));
    await fh.sync();
  } catch (err) {
    await fh.close().catch(() => {});
    await fsp.rm(tmp, { force: true }).catch(() => {});
    throw err;
  }
  await fh.close();
  for (let tries = 1; ; tries += 1) {
    try {
      await fsp.rename(tmp, file);
      return;
    } catch (err) {
      // Windows refuses a rename over a file that another read has open for that moment: tried again shortly.
      if (tries >= 10 || !['EPERM', 'EACCES', 'EBUSY'].includes(err.code)) {
        await fsp.rm(tmp, { force: true }).catch(() => {});
        throw err;
      }
      await new Promise((r) => { setTimeout(r, 15 * tries); });
    }
  }
}

async function readJson(file) {
  try {
    return JSON.parse(await fsp.readFile(file, 'utf8'));
  } catch (err) {
    if (err.code === 'ENOENT') return null;
    throw new LockerError('format', 'The locker\'s key file could not be read.');
  }
}

/**
 * The locker, in the profile folder `dir`.
 *
 * Nothing here is written while the locker is opened: `unlock` only reads. A
 * wrong password answers 'bad-password' and leaves every file exactly as it
 * was.
 *
 * @param {object} o
 * @param {string} o.dir    the locker's folder in Office's profile
 * @param {object} [o.cost] scrypt cost for a new password (a test passes a cheap one)
 */
export function createKeys({ dir, cost = SCRYPT_DEFAULT }) {
  const ringFile = path.join(dir, KEYRING_FILE);
  let root = null;
  let store = null;
  let userKey = null;
  let kdf = null;
  /** What a folder key file is sealed with, cached while unlocked: { wrapKey, recoveryWrap }. */
  let folderWrap = null;
  const folderKeys = new Map();

  const ring = () => readJson(ringFile);
  /** The keyring, for anything that opens or rewrites it: one from a newer Rutba Office is refused and left as it is. */
  const ringToUse = async () => {
    const rec = await ring();
    if (rec) checkVersion(rec.v);
    return rec;
  };
  const wipe = (b) => { if (Buffer.isBuffer(b)) b.fill(0); };

  function attachStore(key) {
    userKey = toHex(Buffer.from(crypto.hkdfSync('sha256', key, Buffer.alloc(32), Buffer.from(`${CONTEXT}keystore`, 'utf8'), 32)));
    store = new SecureKeystore({ keystoreDir: path.join(dir, KEYSTORE_DIR), iterations: 10_000, logger: null });
  }

  function storeHas() {
    return fs.existsSync(path.join(dir, KEYSTORE_DIR, STORE, 'master.json'));
  }

  function createStore() {
    fs.mkdirSync(path.join(dir, KEYSTORE_DIR), { recursive: true });
    const seed = path.join(dir, `.seed-${crypto.randomBytes(8).toString('hex')}.json`);
    // The keystore makes a store from a file, gone at once.
    fs.writeFileSync(seed, JSON.stringify({ master: crypto.randomBytes(32).toString('hex'), __marker: 'locker' }), { flag: 'wx', mode: 0o600 });
    try {
      store.secure(seed, userKey, STORE);
    } finally {
      fs.rmSync(seed, { force: true });
    }
  }

  function loadFolderKeys() {
    folderKeys.clear();
    let names = [];
    try {
      names = store.listKeys(STORE);
    } catch {
      return;
    }
    for (const name of names) {
      if (!name.startsWith('folder:')) continue;
      try {
        folderKeys.set(name.slice(7), fromHex(store.get(name, userKey, STORE)));
      } catch {
        // One entry a crash cut short: the folder's own key file still opens it.
      }
    }
  }

  function storeSecret(name, value) {
    store.set(name, value, userKey, STORE);
  }

  function readSecret(name) {
    try {
      return store.get(name, userKey, STORE);
    } catch {
      return null;
    }
  }

  function afterOpen(k, { folderWrapKey = null } = {}) {
    root = k;
    attachStore(k);
    if (!storeHas()) createStore();
    loadFolderKeys();
    const wrapHex = folderWrapKey ? toHex(folderWrapKey) : readSecret('folder-wrap');
    if (folderWrapKey) storeSecret('folder-wrap', wrapHex);
    const recoveryRec = readSecret('recovery-wrap');
    folderWrap = {
      wrapKey: wrapHex ? fromHex(wrapHex) : null,
      recoveryWrap: recoveryRec ? { salt: recoveryRec.split(':')[0], key: fromHex(recoveryRec.split(':')[1]) } : null,
    };
  }

  const api = {
    /** Whether this profile has a locker yet. */
    exists: async () => Boolean(await ring()),
    /** Whether it is open now. */
    isOpen: () => root !== null,
    /** The ways of opening it this profile has: 'password', 'recovery', 'hello'. */
    ways: async () => ((await ring())?.unlockers || []).map((u) => u.kind),

    /** A new locker: answers the recovery key, which is never stored in a form that shows it again. */
    async setup(password) {
      if (await ring()) throw new LockerError('exists', 'This profile already has a locker.');
      const r = crypto.randomBytes(KEY_LEN);
      kdf = newKdf(cost);
      const pk = await passwordKey(password, kdf);
      const recovery = newRecoveryKey();
      const salt = crypto.randomBytes(32).toString('hex');
      const rWrap = secretKey(recovery.bytes, salt, 'recovery');
      const record = {
        v: 1,
        made: new Date().toISOString(),
        unlockers: [
          { kind: 'password', kdf, ...wrap(pk, r, 'password') },
          { kind: 'recovery', salt, ...wrap(rWrap, r, 'recovery') },
        ],
      };
      await writeJsonAtomic(ringFile, record);
      afterOpen(r, { folderWrapKey: pk });
      storeSecret('recovery-wrap', `${salt}:${toHex(rWrap)}`);
      folderWrap.recoveryWrap = { salt, key: rWrap };
      return { recovery: recovery.text };
    },

    /**
     * Opened with one of the ways. Nothing is written to any folder or file by
     * this; a wrong password throws 'bad-password' and that is all that happens.
     *
     * @param {{ password?: string, recovery?: string, secret?: Buffer }} how
     */
    async unlock({ password = null, recovery = null, secret = null } = {}) {
      const rec = await ringToUse();
      if (!rec) throw new LockerError('none', 'This profile has no locker yet.');
      if (password !== null && password !== undefined) {
        const u = rec.unlockers.find((x) => x.kind === 'password');
        if (!u) throw new LockerError('bad-password', 'That password does not open this locker.');
        const pk = await passwordKey(password, u.kdf);
        const r = unwrap(pk, u, 'password');
        if (!r) throw new LockerError('bad-password', 'That password does not open this locker.');
        kdf = u.kdf;
        afterOpen(r, { folderWrapKey: pk });
        return;
      }
      if (recovery !== null && recovery !== undefined) {
        const u = rec.unlockers.find((x) => x.kind === 'recovery');
        const bytes = recoveryBytes(recovery);
        const r = u ? unwrap(secretKey(bytes, u.salt, 'recovery'), u, 'recovery') : null;
        if (!r) throw new LockerError('bad-recovery', 'That recovery key does not open this locker.');
        kdf = rec.unlockers.find((x) => x.kind === 'password')?.kdf || kdf;
        afterOpen(r);
        return;
      }
      if (secret) {
        const u = rec.unlockers.find((x) => x.kind === 'hello');
        const r = u ? unwrap(secretKey(secret, u.salt, 'hello'), u, 'hello') : null;
        if (!r) throw new LockerError('bad-hello', 'Windows Hello could not open this locker.');
        kdf = rec.unlockers.find((x) => x.kind === 'password')?.kdf || kdf;
        afterOpen(r);
        return;
      }
      throw new LockerError('bad-password', 'Enter the password.');
    },

    /** Closed: the keys forgotten. */
    lock() {
      wipe(root);
      root = null;
      store = null;
      userKey = null;
      folderKeys.forEach(wipe);
      folderKeys.clear();
      if (folderWrap) {
        wipe(folderWrap.wrapKey);
        wipe(folderWrap.recoveryWrap?.key);
      }
      folderWrap = null;
    },

    /** A new password. The old one or the recovery key must be right; on any failure the keyring is exactly as it was. */
    async changePassword(current, next, { recovery = null } = {}) {
      const rec = await ringToUse();
      const u = rec.unlockers.find((x) => x.kind === 'password');
      let r;
      if (recovery !== null) {
        const way = rec.unlockers.find((x) => x.kind === 'recovery');
        r = way ? unwrap(secretKey(recoveryBytes(recovery), way.salt, 'recovery'), way, 'recovery') : null;
        if (!r) throw new LockerError('bad-recovery', 'That recovery key does not open this locker.');
      } else {
        r = unwrap(await passwordKey(current, u.kdf), u, 'password');
        if (!r) throw new LockerError('bad-password', 'That password does not open this locker.');
      }
      const nextKdf = newKdf(cost);
      const nextPk = await passwordKey(next, nextKdf);
      const others = rec.unlockers.filter((x) => x.kind !== 'password');
      await writeJsonAtomic(ringFile, { ...rec, retired: [...(rec.retired || []), u.kdf.salt], unlockers: [{ kind: 'password', kdf: nextKdf, ...wrap(nextPk, r, 'password') }, ...others] });
      kdf = nextKdf;
      if (root) {
        storeSecret('folder-wrap', toHex(nextPk));
        folderWrap.wrapKey = nextPk;
      }
      return { kdf: nextKdf };
    },

    /** A new recovery key, replacing the last one. The password must be right. */
    async replaceRecovery(current) {
      const rec = await ringToUse();
      const u = rec.unlockers.find((x) => x.kind === 'password');
      const pk = await passwordKey(current, u.kdf);
      const r = unwrap(pk, u, 'password');
      if (!r) throw new LockerError('bad-password', 'That password does not open this locker.');
      const recovery = newRecoveryKey();
      const salt = crypto.randomBytes(32).toString('hex');
      const rWrap = secretKey(recovery.bytes, salt, 'recovery');
      const others = rec.unlockers.filter((x) => x.kind !== 'recovery');
      const before = rec.unlockers.find((x) => x.kind === 'recovery')?.salt;
      await writeJsonAtomic(ringFile, { ...rec, retired: [...(rec.retired || []), ...(before ? [before] : [])], unlockers: [...others, { kind: 'recovery', salt, ...wrap(rWrap, r, 'recovery') }] });
      if (root) {
        storeSecret('recovery-wrap', `${salt}:${toHex(rWrap)}`);
        folderWrap.recoveryWrap = { salt, key: rWrap };
      }
      return { recovery: recovery.text };
    },

    /** The name of this locker's key in Windows, once Windows Hello opens it. */
    helloId: async () => ((await ring())?.unlockers || []).find((u) => u.kind === 'hello')?.id || null,

    /** Windows Hello made a way of opening this locker. The locker must be open; the secret comes from the system's helper. */
    async enableHello(secret, id = null) {
      if (!root) throw new LockerError('locked', 'Unlock the locker first.');
      const rec = await ringToUse();
      const salt = crypto.randomBytes(32).toString('hex');
      const entry = { kind: 'hello', id, salt, ...wrap(secretKey(secret, salt, 'hello'), root, 'hello') };
      await writeJsonAtomic(ringFile, { ...rec, unlockers: [...rec.unlockers.filter((x) => x.kind !== 'hello'), entry] });
    },

    /** Windows Hello no longer opens this locker. */
    async disableHello() {
      const rec = await ringToUse();
      await writeJsonAtomic(ringFile, { ...rec, unlockers: rec.unlockers.filter((x) => x.kind !== 'hello') });
    },

    /** The folder key for an id, from the profile's store, while unlocked. */
    folderKey: (id) => folderKeys.get(id) || null,

    /** Every folder key held while unlocked, to try on a locked file copied in from another folder. */
    knownFolderKeys: () => [...folderKeys.values()].map((k) => Buffer.from(k)),

    /** A new or found folder key kept in the profile's store. */
    rememberFolderKey(id, key) {
      if (!root) throw new LockerError('locked', 'Unlock the locker first.');
      storeSecret(`folder:${id}`, toHex(key));
      folderKeys.set(id, Buffer.from(key));
    },

    /** A folder forgotten by this profile (its files are untouched). */
    forgetFolderKey(id) {
      if (!root) throw new LockerError('locked', 'Unlock the locker first.');
      try {
        store.deleteKey(`folder:${id}`, STORE);
      } catch (err) {
        if (err?.code !== ErrorCodes.KEY_NOT_FOUND) throw err;
      }
      folderKeys.delete(id);
    },

    /** What goes in a folder's own key file, so it opens on any computer: the folder key sealed by the password's key and the recovery key's. */
    wrapForFolder(fk, id) {
      if (!folderWrap?.wrapKey || !kdf) throw new LockerError('locked', 'Unlock the locker with its password first.');
      const unlockers = [{ kind: 'password', kdf, ...wrap(folderWrap.wrapKey, fk, `folder:${id}`) }];
      if (folderWrap.recoveryWrap) unlockers.push({ kind: 'recovery', salt: folderWrap.recoveryWrap.salt, ...wrap(folderWrap.recoveryWrap.key, fk, `folder:${id}`) });
      return { v: 1, id, unlockers };
    },

    /** The folder key inside a key file, with the password or recovery key, or with what is cached while unlocked; null when none of them opens it. */
    async openFolderKeys(file, { password = null, recovery = null } = {}) {
      checkVersion(file?.v);
      const id = file.id;
      // A folder opened by more than one locker has an entry for each.
      const entries = (kind) => (file.unlockers || []).filter((x) => x.kind === kind).slice(0, MOST_TRIED);
      if (password !== null) {
        for (const u of entries('password')) {
          const opened = unwrap(await passwordKey(password, u.kdf), u, `folder:${id}`);
          if (opened) return opened;
        }
        return null;
      }
      if (recovery !== null) {
        if (!entries('recovery').length) return null;
        const bytes = recoveryBytes(recovery);
        for (const u of entries('recovery')) {
          const opened = unwrap(secretKey(bytes, u.salt, 'recovery'), u, `folder:${id}`);
          if (opened) return opened;
        }
        return null;
      }
      if (folderWrap?.wrapKey) {
        for (const u of entries('password')) {
          const opened = unwrap(folderWrap.wrapKey, u, `folder:${id}`);
          if (opened) return opened;
        }
      }
      return null;
    },

    /** Whether a folder's key file already has this locker's present password and recovery key, and nothing retired. */
    async folderKeysCurrent(file) {
      const retired = new Set((await ring())?.retired || []);
      const salts = (file.unlockers || []).map(saltOf);
      if (salts.some((s) => retired.has(s))) return false;
      return salts.includes(kdf?.salt) && (!folderWrap?.recoveryWrap || salts.includes(folderWrap.recoveryWrap.salt));
    },

    /**
     * A folder's key file brought up to date for this locker: its own entries
     * made afresh and those it has retired removed. Entries belonging to another
     * locker that opened the same folder stay: this locker cannot make them
     * again, and they may be the only way that person has in.
     */
    async mergedFolderKeys(existing, fk, id) {
      const retired = new Set((await ring())?.retired || []);
      const mine = api.wrapForFolder(fk, id);
      const own = new Set(mine.unlockers.map(saltOf));
      const others = (existing?.unlockers || []).filter((u) => !retired.has(saltOf(u)) && !own.has(saltOf(u)));
      return { ...mine, unlockers: [...mine.unlockers, ...others] };
    },

    /** Everything this module would write as a folder's key file. */
    readFolderKeysFile: (p) => readJson(p),
    writeFolderKeysFile: (p, value) => writeJsonAtomic(p, value),
    kdfNow: () => kdf,
  };
  return api;
}
