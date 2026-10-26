// Settings, saved passwords and the recent list, in the profile's folder.
//
// Saved passwords, sign-in tokens and keys (mail, calendar and contacts
// servers, Talk, Data Files' engines and sources, stream keys, certificates'
// keys, the AI tools' network key) are kept in the profile's own keystore
// (profile-keystore.js, D-2026-10-26-09): its key is random, linked to the
// profile's long id and wrapped by the system's protection (Electron's
// safeStorage: DPAPI on Windows, the Keychain on macOS, the secret service on
// Linux). Where the system gives no real protection (Linux with no keyring:
// Electron then says encryption is available but uses a fixed key,
// basic_text), passwords are kept only in a profile with a password; without
// one, a place that would save one is told so and asks each time instead.
// Every caller goes through `secrets` here, as before.
//
// A profile with a password (profiles.js) seals the keystore's key file and
// recent.json once more with the key made from it, through the vault handed
// in here, so another person at the same Windows account, whom the system's
// protection lets in, still cannot read them. settings.json stays plain: the
// AI tools, run as Node by another program, read it for their switch and
// folders, and never the keystore.

import { app, safeStorage } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { removeEntry, renameEntry, presentEntries } from '../recent-store.js';
import { createVault, DEFAULT_ID } from '../profiles.js';
import { profileUid, openProfileKeystore, moveOldSecrets, OLD_SECRETS_FILE } from '../profile-keystore.js';

/** Settings that once held a key in the clear, moved into the keystore once: [setting, name in the keystore]. */
export const PLAIN_SECRETS = [['ai.tools.remoteKey', 'ai.tools.remoteKey']];
/** Where a keystore put aside is noted, to be said once. */
export const SET_ASIDE_KEY = 'keystore.setAside';

function readJson(file, fallback) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return fallback;
  }
}

function writeJson(file, value) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = `${file}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(value, null, 2));
  fs.renameSync(tmp, file); // atomic-enough: a crash mid-write leaves the old file
}

/** The system's protection, as the keystore asks for it. */
export const systemProtection = {
  available: () => {
    try {
      // On a Linux desktop with no keyring, Electron still reports encryption
      // as available but falls back to a fixed built-in key (basic_text):
      // that is no protection, so it is not taken for one.
      if (process.platform === 'linux' && safeStorage.getSelectedStorageBackend?.() === 'basic_text') return false;
      return safeStorage.isEncryptionAvailable();
    } catch {
      return false;
    }
  },
  encrypt: (text) => safeStorage.encryptString(String(text)),
  decrypt: (bytes) => safeStorage.decryptString(Buffer.from(bytes)),
};

/**
 * @param {object} [o]
 * @param {object} [o.vault]     the profile's vault (profiles.js)
 * @param {object} [o.profile]   the profile this run is ({ base, id }), whose long id the keystore is linked to
 * @param {object} [o.protector] the system's protection (systemProtection)
 */
export function createStores({ vault = createVault(), profile = null, protector = systemProtection } = {}) {
  const dir = app.getPath('userData');
  const settingsFile = path.join(dir, 'settings.json');
  const recentFile = path.join(dir, 'recent.json');
  // Read for the one-time move only: passwords are not written here any more.
  vault.cover(path.join(dir, OLD_SECRETS_FILE));
  vault.cover(recentFile);

  let settings = readJson(settingsFile, {});
  let recent = vault.readJson(recentFile, []);

  const settingsApi = {
    get: (key, fallback) => (key in settings ? settings[key] : fallback),
    set: (key, value) => {
      settings[key] = value;
      writeJson(settingsFile, settings);
    },
    delete: (key) => {
      delete settings[key];
      writeJson(settingsFile, settings);
    },
    all: () => ({ ...settings }),
  };

  // The profile's keystore, opened once its password (if it has one) has been given.
  const uid = profileUid({ base: profile?.base || null, id: profile?.id || DEFAULT_ID, folder: dir });
  const keystore = openProfileKeystore({
    folder: dir,
    uid,
    vault,
    protector,
    onSetAside: (aside) => settingsApi.set(SET_ASIDE_KEY, { ...aside, seen: false }),
  });
  const move = () => {
    try {
      return moveOldSecrets({ folder: dir, keystore, vault, protector, settings: settingsApi, plain: PLAIN_SECRETS });
    } catch (err) {
      console.error('saved passwords, moving into the keystore:', err?.message || err);
      return null;
    }
  };
  move();

  const secretsApi = {
    /** Whether a password saved now is kept. */
    available: () => keystore.available(),
    get: (key) => keystore.get(key),
    set: (key, value) => keystore.set(key, value),
    delete: (key) => keystore.delete(key),
    keys: () => keystore.keys(),
    status: () => keystore.status(),
    /** The keystore opened again, as at the next start (after the profile's password changes), and anything left to move moved. */
    refresh: () => {
      const s = keystore.refresh();
      move();
      return s;
    },
    beforePasswordGoes: () => keystore.beforePasswordGoes(),
    /** A keystore put aside and not yet said: { at, why }, taken as it is given, so only one window says it. */
    notice: () => {
      const n = settingsApi.get(SET_ASIDE_KEY, null);
      if (!n || n.seen) return null;
      settingsApi.set(SET_ASIDE_KEY, { ...n, seen: true });
      return { at: n.at, why: n.why };
    },
  };
  // A window check opens the keystore again as the next start would (verify-keystore.js).
  if (process.env.RUTBA_OFFICE_VERIFY_APPS) globalThis.__rutbaCheckSecrets = { ...secretsApi, forgetNotice: () => settingsApi.delete(SET_ASIDE_KEY) };

  const recentApi = {
    // Looked for together, without waiting on a drive that has gone (see presentEntries).
    list: () => presentEntries(recent, { exists: (p) => fs.promises.access(p).then(() => true, () => false) }),
    add: ({ path: p, app: which }) => {
      if (!p) return recentApi.list();
      recent = [
        { path: p, name: path.basename(p), app: which || null, at: Date.now() },
        ...recent.filter((r) => r.path !== p),
      ].slice(0, 60);
      vault.writeJson(recentFile, recent);
      try {
        app.addRecentDocument(p);
      } catch { /* jump list is a nicety, not a requirement */ }
      return recentApi.list();
    },
    remove: ({ path: p }) => {
      recent = removeEntry(recent, p);
      vault.writeJson(recentFile, recent);
      return recentApi.list();
    },
    // The file on disk has already been renamed by the caller; this just
    // points the entry at where it went.
    rename: ({ path: p, to }) => {
      recent = renameEntry(recent, p, to);
      vault.writeJson(recentFile, recent);
      return recentApi.list();
    },
    clear: () => {
      recent = [];
      vault.writeJson(recentFile, recent);
      try { app.clearRecentDocuments(); } catch { /* as above */ }
      return [];
    },
  };

  return { settings: settingsApi, secrets: secretsApi, recent: recentApi, dir, vault };
}
