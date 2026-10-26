// Profiles (D-2026-10-26-02): each person, or each part of a person's work,
// in a folder of its own, two or more open side by side, as in Chrome.
//
// What lives in a profile's folder, and so is kept apart for each profile:
//   settings.json      every setting: language, theme, apps chosen, mail and
//                      calendar accounts (addresses and servers, never their
//                      passwords), AI Tools' switch and folders, Data Files'
//                      allowed sources, the personal dictionary
//   secrets.json       saved passwords and sign-in tokens (store.js)
//   recent.json        the recent files
//   mail\              every account's folders, messages and attachments
//                      (@rutba/mailbox store; its search index is in memory)
//   calendar.json      calendars and events;  contacts.json  the address book
//   outgoing\          invitation replies on their way out
//   talk\              Talk's messages, per contact
//   recovery\          unsaved work a crash would take
//   thumbnails\        the Pictures and file-list thumbnails
//   ai-tools-log.jsonl, ai-tools-connections.json   what the AI tools did, who is connected
//   errors.log         the main process's faults
//   and Chromium's own (Local Storage, Cache, ...), which follow the folder.
// What every profile shares, so it stays in the base folder or the program:
//   the installed program and its updates, the downloaded template sets
//   (<base>\templates, template-set-choice.js), the spelling dictionaries
//   (in the program), and this list, profiles.json.
//
// The base folder is the one Office has always used (%APPDATA%\Rutba Office,
// or the folder --user-data-dir names) and it stays the first profile,
// "Default", so nothing is moved. A profile made later lives in
// <base>\Profiles\<id>. The list names each profile, its colour, its folder,
// the one that opens first and the one used last, and, for a profile with a
// password, the salt and a verifier: never the password, never the key
// (while a password change is under way, each key sealed with the other).
//
// Nothing here needs Electron, so the engine tests and the AI tools (run as
// Node) read the same list the windows do.

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';

export const PROFILES_FILE = 'profiles.json';
export const PROFILES_DIR = 'Profiles';
export const DEFAULT_ID = 'default';
/** The colours offered for a profile, the first the Default's. */
export const COLOURS = ['#2b5fd9', '#1a9f7a', '#d9534f', '#e08b2b', '#7b5cd6', '#c2408f', '#3aa0b5', '#8c6d1f'];
export const NAME_MAX = 40;
export const PASSWORD_MIN = 4;

const isColour = (c) => typeof c === 'string' && /^#[0-9a-f]{6}$/i.test(c);
const cleanName = (name) => String(name ?? '').replace(/\s+/g, ' ').trim();

function defaultProfile() {
  return { id: DEFAULT_ID, name: 'Default', colour: COLOURS[0], folder: '.', picture: null, lock: null };
}

/** The list as it is when nobody has made a second profile: Default alone, and no file written. */
export function defaultList() {
  return { version: 1, openFirst: DEFAULT_ID, lastUsed: DEFAULT_ID, profiles: [defaultProfile()] };
}

/** One entry made whole: a bad colour, folder or name put right, never thrown over. */
function normalise(p) {
  const id = String(p?.id || '').trim();
  if (!/^[a-z0-9-]{1,40}$/i.test(id)) return null;
  const folder = id === DEFAULT_ID ? '.' : `${PROFILES_DIR}/${id}`;
  const lock = p.lock && typeof p.lock.salt === 'string' && typeof p.lock.verifier === 'string' ? p.lock : null;
  return { ...p, id, name: cleanName(p.name).slice(0, NAME_MAX) || (id === DEFAULT_ID ? 'Default' : id), colour: isColour(p.colour) ? p.colour : COLOURS[0], folder, picture: typeof p.picture === 'string' ? p.picture : null, lock };
}

/** The list in a base folder. A list that is missing or unreadable is Default alone. */
export function readProfiles(base) {
  let raw = null;
  try {
    raw = JSON.parse(fs.readFileSync(path.join(base, PROFILES_FILE), 'utf8'));
  } catch {
    raw = null;
  }
  if (!raw || !Array.isArray(raw.profiles)) return defaultList();
  const seen = new Set();
  const profiles = [];
  for (const p of raw.profiles) {
    const n = normalise(p);
    if (!n || seen.has(n.id)) continue;
    seen.add(n.id);
    profiles.push(n);
  }
  // Default is always there: its folder is the base folder, which holds the list itself.
  if (!seen.has(DEFAULT_ID)) profiles.unshift(defaultProfile());
  const has = (id) => profiles.some((p) => p.id === id);
  return { ...raw, version: 1, profiles, openFirst: has(raw.openFirst) ? raw.openFirst : DEFAULT_ID, lastUsed: has(raw.lastUsed) ? raw.lastUsed : DEFAULT_ID };
}

/** Written whole, beside itself first, so a crash mid-write leaves the old list. */
export function writeProfiles(base, list) {
  fs.mkdirSync(base, { recursive: true });
  const file = path.join(base, PROFILES_FILE);
  const tmp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(list, null, 2));
  fs.renameSync(tmp, file);
  return list;
}

/** A profile's folder: the base folder for Default, <base>\Profiles\<id> for the rest. */
export function profileFolder(base, profile) {
  if (!profile || profile.id === DEFAULT_ID || profile.folder === '.') return base;
  return path.join(base, PROFILES_DIR, profile.id);
}

/** A profile named by its id, or else by its name, whatever its case. */
export function findProfile(list, key) {
  if (!key) return null;
  const k = String(key).trim();
  return list.profiles.find((p) => p.id === k) || list.profiles.find((p) => p.name.toLowerCase() === k.toLowerCase()) || null;
}

/**
 * `--profile <id or name>` or `--profile=<id or name>` on a command line:
 * the profile asked for, and the index of its value when it came as a word
 * of its own (so it is not taken for a file to open).
 */
export function profileArg(argv = []) {
  for (let i = 0; i < argv.length; i++) {
    const a = String(argv[i]);
    if (a.startsWith('--profile=')) return { value: a.slice('--profile='.length).replace(/^"|"$/g, '') || null, at: -1 };
    if (a === '--profile' && i + 1 < argv.length && !String(argv[i + 1]).startsWith('-')) return { value: String(argv[i + 1]), at: i + 1 };
  }
  return { value: null, at: -1 };
}

/**
 * The profile this run is: the one the command line names; else, for a file
 * or a link handed to Office (from Explorer, an email), the one used last,
 * so the file opens where the person was working; else the one marked to
 * open first.
 */
export function chooseProfile({ base, argv = [], handed = false }) {
  const list = readProfiles(base);
  const asked = profileArg(argv).value;
  const named = asked ? findProfile(list, asked) : null;
  if (named) return { profile: named, list, how: 'asked' };
  if (handed) {
    const last = findProfile(list, list.lastUsed);
    if (last) return { profile: last, list, how: 'last used' };
  }
  return { profile: findProfile(list, list.openFirst) || list.profiles[0], list, how: asked ? 'not found' : 'opens first' };
}

/** A new profile's id: short, and never one in use. */
function newId(list) {
  for (;;) {
    const id = crypto.randomBytes(4).toString('hex');
    if (!list.profiles.some((p) => p.id === id)) return id;
  }
}

function checkName(list, name, except = null) {
  const n = cleanName(name);
  if (!n) throw new Error('Enter a name for the profile.');
  if (n.length > NAME_MAX) throw new Error(`A profile's name can have up to ${NAME_MAX} characters.`);
  if (list.profiles.some((p) => p.id !== except && p.name.toLowerCase() === n.toLowerCase())) throw new Error('A profile of that name already exists.');
  return n;
}

function entry(list, id) {
  const p = list.profiles.find((x) => x.id === id);
  if (!p) throw new Error('That profile is no longer there.');
  return p;
}

/** A new profile, with a folder of its own. */
export function makeProfile(base, { name, colour } = {}) {
  const list = readProfiles(base);
  const n = checkName(list, name);
  const id = newId(list);
  const used = new Set(list.profiles.map((p) => p.colour));
  const p = { id, name: n, colour: isColour(colour) ? colour : COLOURS.find((c) => !used.has(c)) || COLOURS[list.profiles.length % COLOURS.length], folder: `${PROFILES_DIR}/${id}`, picture: null, lock: null, made: new Date().toISOString() };
  fs.mkdirSync(profileFolder(base, p), { recursive: true });
  list.profiles.push(p);
  writeProfiles(base, list);
  return { list, profile: p };
}

export function renameProfile(base, id, name) {
  const list = readProfiles(base);
  entry(list, id).name = checkName(list, name, id);
  return writeProfiles(base, list);
}

export function setProfileColour(base, id, colour) {
  if (!isColour(colour)) throw new Error('That is not a colour.');
  const list = readProfiles(base);
  entry(list, id).colour = colour.toLowerCase();
  return writeProfiles(base, list);
}

export function setOpenFirst(base, id) {
  const list = readProfiles(base);
  entry(list, id);
  list.openFirst = id;
  return writeProfiles(base, list);
}

/** The profile in use now, for a file opened from Explorer. Nothing is written while Default is the only profile. */
export function markUsed(base, id) {
  const list = readProfiles(base);
  if (list.lastUsed === id || !list.profiles.some((p) => p.id === id)) return false;
  if (!fs.existsSync(path.join(base, PROFILES_FILE))) return false;
  list.lastUsed = id;
  writeProfiles(base, list);
  return true;
}

/** A profile's lock, written or taken away (the caller re-seals its files). */
export function setProfileLock(base, id, lock) {
  const list = readProfiles(base);
  entry(list, id).lock = lock || null;
  return writeProfiles(base, list);
}

/**
 * Whether a profile is open in a running Rutba Office. Chromium holds a
 * `lockfile` in the folder while it runs (Windows deletes it when the
 * program ends, even after a crash); elsewhere a `SingletonLock` link names
 * the process, which is looked for.
 */
export function isRunning(folder) {
  if (process.platform === 'win32') {
    const lock = path.join(folder, 'lockfile');
    if (!fs.existsSync(lock)) return false;
    try {
      // Held open by the running copy with no sharing for writing: opening it fails while it runs.
      fs.closeSync(fs.openSync(lock, 'r+'));
      return false;
    } catch (err) {
      return err?.code === 'EBUSY' || err?.code === 'EPERM' || err?.code === 'EACCES';
    }
  }
  try {
    const target = fs.readlinkSync(path.join(folder, 'SingletonLock'));
    const pid = Number(String(target).split('-').pop());
    if (!pid) return false;
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

/** How much a folder holds, for the question before a profile is removed. */
export function folderSize(folder) {
  let total = 0;
  const walk = (dir) => {
    let entries = [];
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const e of entries) {
      const p = path.join(dir, e.name);
      if (e.isDirectory()) walk(p);
      else {
        try {
          total += fs.statSync(p).size;
        } catch {
          /* gone while counted */
        }
      }
    }
  };
  walk(folder);
  return total;
}

/**
 * Whether a profile may be removed, and if not, why, in a sentence: never
 * Default (its folder is the base folder, which holds what every profile
 * shares), never the last one, never the one this window belongs to, never
 * one open in another running Rutba Office.
 */
export function removalRefused(base, id, { current = null, running = isRunning } = {}) {
  const list = readProfiles(base);
  const p = list.profiles.find((x) => x.id === id);
  if (!p) return 'That profile is no longer there.';
  if (list.profiles.length <= 1) return 'The last profile cannot be removed.';
  if (id === DEFAULT_ID) return 'The first profile keeps what every profile shares, so it cannot be removed. You can rename it.';
  if (id === current) return 'This window belongs to that profile. Open another profile and remove it from there.';
  if (running(profileFolder(base, p))) return 'That profile is open. Close its windows first.';
  return null;
}

/** A profile taken off the list, its folder handed to `discard` (the Recycle Bin, in the app). */
export async function removeProfile(base, id, { current = null, running = isRunning, discard }) {
  const refused = removalRefused(base, id, { current, running });
  if (refused) throw new Error(refused);
  const list = readProfiles(base);
  const p = entry(list, id);
  const folder = profileFolder(base, p);
  // The folder goes first: a profile still listed with its folder gone opens empty, never the other way round.
  if (fs.existsSync(folder)) await discard(folder);
  const now = readProfiles(base);
  now.profiles = now.profiles.filter((x) => x.id !== id);
  if (now.openFirst === id) now.openFirst = DEFAULT_ID;
  if (now.lastUsed === id) now.lastUsed = DEFAULT_ID;
  return writeProfiles(base, now);
}

/* ── an optional password ───────────────────────────────────────────────── */

// scrypt at 2^15: about a tenth of a second on a laptop, 32 MB of memory,
// for every guess. The 64 bytes it makes are two keys: the first seals the
// files, the second only makes the verifier, so the verifier on disk says
// nothing about the first.
export const SCRYPT = { N: 1 << 15, r: 8, p: 1 };
const MAXMEM = 128 * 1024 * 1024;
const VERIFIER = 'rutba-office-profile-verifier';

function derive(password, salt, params = SCRYPT) {
  return new Promise((resolve, reject) => {
    crypto.scrypt(String(password).normalize('NFC'), salt, 64, { N: params.N, r: params.r, p: params.p, maxmem: MAXMEM }, (err, out) => (err ? reject(err) : resolve(out)));
  });
}

const verifierOf = (check) => crypto.createHmac('sha256', check).update(VERIFIER).digest('base64');

/** A new lock for a password: what profiles.json keeps, and the key, which it never does. */
export async function makeLock(password) {
  if (String(password ?? '').length < PASSWORD_MIN) throw new Error(`A password needs at least ${PASSWORD_MIN} characters.`);
  const salt = crypto.randomBytes(16);
  const out = await derive(password, salt);
  const lock = { kdf: 'scrypt', N: SCRYPT.N, r: SCRYPT.r, p: SCRYPT.p, salt: salt.toString('base64'), verifier: verifierOf(out.subarray(32)), set: new Date().toISOString() };
  return { lock, key: Buffer.from(out.subarray(0, 32)) };
}

/** The key for a password, or null when it is not this lock's. */
export async function unlock(lock, password) {
  if (!lock?.salt || !lock?.verifier) return null;
  const out = await derive(password, Buffer.from(lock.salt, 'base64'), { N: lock.N || SCRYPT.N, r: lock.r || SCRYPT.r, p: lock.p || SCRYPT.p });
  const a = Buffer.from(verifierOf(out.subarray(32)), 'base64');
  const b = Buffer.from(lock.verifier, 'base64');
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  return Buffer.from(out.subarray(0, 32));
}

/* ── files sealed with the key ──────────────────────────────────────────── */

const AAD = Buffer.from('rutba-office-profile');

/** A text sealed: AES-256-GCM, a fresh nonce each time, written as a small JSON envelope. */
export function seal(key, text) {
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv('aes-256-gcm', key, iv);
  c.setAAD(AAD);
  const data = Buffer.concat([c.update(Buffer.from(String(text), 'utf8')), c.final()]);
  return JSON.stringify({ rutbaSealed: 1, iv: iv.toString('base64'), tag: c.getAuthTag().toString('base64'), data: data.toString('base64') });
}

export const isSealed = (text) => typeof text === 'string' && text.startsWith('{"rutbaSealed":');

// Chromium's own folders and those of other profiles (under Default's) hold nothing sealed.
const NOT_SEALED_HERE = new Set([PROFILES_DIR, 'Cache', 'Code Cache', 'GPUCache', 'DawnGraphiteCache', 'DawnWebGPUCache', 'Local Storage', 'Session Storage', 'Service Worker', 'IndexedDB', 'blob_storage', 'Crashpad', 'Network', 'Shared Dictionary', 'shared_proto_db', 'WebStorage', 'Partitions', 'thumbnails', 'recovery']);

/** Every .json file in a profile's folder that is sealed, whoever covers it. */
export function sealedFilesIn(folder) {
  const out = [];
  const head = Buffer.alloc(16);
  const walk = (dir, top) => {
    let entries = [];
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const e of entries) {
      const p = path.join(dir, e.name);
      if (e.isDirectory()) {
        if (!(top && NOT_SEALED_HERE.has(e.name))) walk(p, false);
      } else if (e.isFile() && e.name.endsWith('.json')) {
        try {
          const fd = fs.openSync(p, 'r');
          let n = 0;
          try { n = fs.readSync(fd, head, 0, 16, 0); } finally { fs.closeSync(fd); }
          if (isSealed(head.subarray(0, n).toString('utf8'))) out.push(p);
        } catch {
          /* gone while looked at */
        }
      }
    }
  };
  walk(folder, true);
  return out;
}

/** A sealed text opened, or an error when the key is not the one it was sealed with. */
export function unseal(key, text) {
  if (!key) throw new Error('This profile is locked.');
  const env = JSON.parse(text);
  const d = crypto.createDecipheriv('aes-256-gcm', key, Buffer.from(env.iv, 'base64'));
  d.setAAD(AAD);
  d.setAuthTag(Buffer.from(env.tag, 'base64'));
  return Buffer.concat([d.update(Buffer.from(env.data, 'base64')), d.final()]).toString('utf8');
}

/**
 * The files a profile's password protects, read and written through here:
 * sealed while the profile has a key, plain while it has none, and every
 * one of them sealed again when the key changes (`rekey`). Writes to a file
 * are queued one after another, so a write still under way when the key
 * changes cannot land after the file was sealed again. A sealed file met
 * without its key is never written over: its reads give the fallback and
 * its writes are refused, so nothing in it is lost.
 *
 * `also` names keys a file may still be sealed with while a password change
 * stopped midway is finished (finishRekey): read with any of them, written
 * with `key` alone, and forgotten once every file is sealed again.
 */
export function createVault({ key = null, also = [] } = {}) {
  let current = key ? Buffer.from(key) : null;
  let older = (also || []).filter(Boolean).map((k) => Buffer.from(k));
  const covered = new Set();
  // Stores that seal their own files (the mail store): { useKey, prepare, rekey }.
  const attached = new Set();
  const held = new Set();
  const chains = new Map();

  const queue = (file, step) => {
    const run = (chains.get(file) || Promise.resolve()).catch(() => {}).then(step);
    chains.set(file, run);
    return run;
  };
  const sealed = (text) => (current ? seal(current, text) : text);
  const openWith = (raw, keys) => {
    for (const k of keys) {
      if (!k) continue;
      try {
        return unseal(k, raw);
      } catch {
        /* the next key */
      }
    }
    return null;
  };
  const writeRaw = (file, text) => {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    const tmp = `${file}.${process.pid}.tmp`;
    fs.writeFileSync(tmp, text);
    fs.renameSync(tmp, file);
  };
  const refuseHeld = (file) => {
    if (held.has(file)) throw new Error('This profile’s protected files are locked.');
  };

  const vault = {
    /** Whether files written now are sealed. */
    sealed: () => Boolean(current),
    /** A file this vault keeps (sealed again when the key changes); a folder covers every .json in it. */
    cover: (p) => covered.add(p),
    /** A store that seals its own files: given the key now, told before the lock changes, and sealing its files again in rekey. */
    attach(store) {
      attached.add(store);
      store.useKey?.(current);
      return () => attached.delete(store);
    },
    /** Called before the profile's lock changes to `next`, so an attached store can write down what a crash would need. */
    async prepare(next) {
      for (const s of attached) await s.prepare?.(current, next ? Buffer.from(next) : null);
    },
    readText(file) {
      let raw;
      try {
        raw = fs.readFileSync(file, 'utf8');
      } catch {
        return null;
      }
      if (!isSealed(raw)) return raw;
      const plain = openWith(raw, [current, ...older]);
      if (plain === null) held.add(file);
      return plain;
    },
    readJson(file, fallback) {
      const text = vault.readText(file);
      if (text === null) return fallback;
      try {
        return JSON.parse(text);
      } catch {
        return fallback;
      }
    },
    writeJson(file, value) {
      refuseHeld(file);
      writeRaw(file, sealed(JSON.stringify(value)));
    },
    /** Written off the caller's way, after any write to the same file still under way. */
    writeJsonLater(file, value) {
      if (held.has(file)) return Promise.reject(new Error('This profile’s protected files are locked.'));
      return queue(file, async () => {
        const text = sealed(JSON.stringify(value));
        await fs.promises.mkdir(path.dirname(file), { recursive: true });
        const tmp = `${file}.${process.pid}.later.tmp`;
        await fs.promises.writeFile(tmp, text);
        await fs.promises.rename(tmp, file);
      });
    },
    /**
     * Every covered file sealed with a new key (or left plain, with none).
     * `folder`: any other file found sealed in the profile's folder is sealed
     * again too, so a file no service covered this run is not left with a
     * key about to be forgotten. `step('file')` is told after each file.
     */
    async rekey(next, { folder = null, step = null } = {}) {
      const keys = [current, ...older];
      current = next ? Buffer.from(next) : null;
      const files = [];
      for (const p of covered) {
        let st = null;
        try {
          st = fs.statSync(p);
        } catch {
          continue;
        }
        if (st.isFile()) files.push(p);
        else if (st.isDirectory()) {
          const walk = (dir) => {
            for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
              const f = path.join(dir, e.name);
              if (e.isDirectory()) walk(f);
              else if (e.name.endsWith('.json')) files.push(f);
            }
          };
          walk(p);
        }
      }
      if (folder) for (const f of sealedFilesIn(folder)) if (!files.includes(f)) files.push(f);
      let done = 0;
      for (const file of files) {
        await queue(file, async () => {
          let raw;
          try {
            raw = fs.readFileSync(file, 'utf8');
          } catch {
            return;
          }
          // Sealed with the old key, or by a write queued just before with the new one.
          const plain = isSealed(raw) ? openWith(raw, [...keys, current]) : raw;
          if (plain === null) return;
          writeRaw(file, sealed(plain));
          held.delete(file);
          done += 1;
        });
        await step?.('file');
      }
      for (const s of attached) await s.rekey?.(keys[0], current);
      older = [];
      return { files: done };
    },
  };
  return vault;
}

/* ── a password changed safely ──────────────────────────────────────────── */

// Setting, changing or taking away a password seals every covered file
// again. A stop midway (a crash, the power going) must never leave files
// sealed with a key the lock in force no longer gives. So the change is
// written down in the profile's entry before any file is touched (`rekey`:
// the new lock, and each key sealed with the other), the lock in force
// stays the old one until every file is sealed again, and the new lock goes
// in, and the note goes, in the one last write. A change found half done at
// the next open is finished then: either password opens the profile, the
// note gives the other key, and the files are sealed again before anything
// of the profile is used (finishRekey).

const wrapKey = (key, other) => seal(key, Buffer.from(other).toString('base64'));
const unwrapKey = (key, text) => {
  try {
    return text ? Buffer.from(unseal(key, text), 'base64') : null;
  } catch {
    return null;
  }
};

/** Whether opening a profile asks for a password: it has one, or one is being set. */
export const needsPassword = (p) => Boolean(p?.lock || p?.rekey?.lock);

/**
 * What a password opens of a profile's entry: { key, rekey }, `key` the one
 * its files are written with from now on and `rekey` ({ from, to }, the old
 * key and the new) when a change was stopped midway, for finishRekey; null
 * when the password is not the profile's, old or new.
 */
export async function unlockProfile(p, password) {
  const r = p?.rekey || null;
  if (!r) {
    const key = await unlock(p?.lock, password);
    return key ? { key, rekey: null } : null;
  }
  let from = p.lock ? await unlock(p.lock, password) : null;
  let to = null;
  if (from) to = r.lock ? unwrapKey(from, r.toSealed) : null;
  else {
    to = r.lock ? await unlock(r.lock, password) : null;
    if (!to) return null;
    from = p.lock ? unwrapKey(to, r.fromSealed) : null;
  }
  // A note whose keys do not open each other is not one this code wrote: the lock in force stays.
  if ((p.lock && !from) || (r.lock && !to)) return from ? { key: from, rekey: null } : null;
  return { key: to, rekey: { from, to } };
}

/** The new lock in force and the note gone, in one write. */
function commitRekey(base, id) {
  const list = readProfiles(base);
  const p = entry(list, id);
  if (p.rekey) p.lock = p.rekey.lock || null;
  delete p.rekey;
  writeProfiles(base, list);
  return p.lock;
}

/**
 * A profile's password set (`from` null), changed, or taken away (`to` and
 * `lock` null): the note, every file `vault` covers and every sealed file in
 * `folder` sealed again with `to`, then the new lock. Answers the lock in
 * force. `step(name)` hears each step reached ('noted', 'file' after each
 * file, 'sealed'): the tests stop the change at each, as a crash would.
 */
export async function rekeyProfile({ base, id, vault, from = null, to = null, lock = null, folder = null, step = null }) {
  const list = readProfiles(base);
  const p = entry(list, id);
  p.rekey = {
    lock: lock || null,
    ...(from && to ? { toSealed: wrapKey(from, to), fromSealed: wrapKey(to, from) } : {}),
    started: new Date().toISOString(),
  };
  writeProfiles(base, list);
  await step?.('noted');
  await vault.rekey(to, { folder, step });
  await step?.('sealed');
  return commitRekey(base, id);
}

/**
 * A change found half done, finished: the vault was opened with the new key
 * and the old one beside it (createVault({ key: to, also: [from] })); every
 * file is sealed with the new key, then the new lock goes in. Answers the
 * lock in force.
 */
export async function finishRekey({ base, id, vault, keys, folder = null }) {
  await vault.rekey(keys?.to || null, { folder });
  return commitRekey(base, id);
}
