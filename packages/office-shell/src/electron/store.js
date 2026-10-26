// Settings, secrets and the recent list — three small JSON files in userData.
//
// Secrets go through Electron's safeStorage, which uses the OS keychain
// (DPAPI on Windows, Keychain on macOS, libsecret on Linux). When the OS has no
// keystore available, we refuse to write rather than quietly saving a password
// in clear text: the caller is told, and can decide to ask every time instead.
//
// A profile with a password (profiles.js) seals secrets.json and recent.json
// once more with the key made from it, through the vault handed in here, so
// another person at the same Windows account, whom the keychain lets in,
// still cannot read them. settings.json stays plain: the AI tools, run as
// Node by another program, read it for their switch and folders.

import { app, safeStorage } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import { removeEntry, renameEntry, presentEntries } from '../recent-store.js';
import { createVault } from '../profiles.js';

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

export function createStores({ vault = createVault() } = {}) {
  const dir = app.getPath('userData');
  const settingsFile = path.join(dir, 'settings.json');
  const secretsFile = path.join(dir, 'secrets.json');
  const recentFile = path.join(dir, 'recent.json');
  vault.cover(secretsFile);
  vault.cover(recentFile);

  let settings = readJson(settingsFile, {});
  let secrets = vault.readJson(secretsFile, {});
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

  const secretsApi = {
    available: () => {
      try {
        // On a Linux desktop with no keyring, Electron still reports encryption
        // as available but falls back to a fixed built-in key (basic_text):
        // that is not a keystore, so nothing is saved there either.
        if (process.platform === 'linux' && safeStorage.getSelectedStorageBackend?.() === 'basic_text') return false;
        return safeStorage.isEncryptionAvailable();
      } catch {
        return false;
      }
    },
    get: (key) => {
      const blob = secrets[key];
      if (!blob) return null;
      try {
        return safeStorage.decryptString(Buffer.from(blob, 'base64'));
      } catch {
        return null; // keychain rotated, or the file moved between machines
      }
    },
    set: (key, value) => {
      if (!secretsApi.available()) {
        throw new Error('No OS keystore is available, so this password cannot be stored safely.');
      }
      secrets[key] = safeStorage.encryptString(String(value)).toString('base64');
      vault.writeJson(secretsFile, secrets);
    },
    delete: (key) => {
      delete secrets[key];
      vault.writeJson(secretsFile, secrets);
    },
    keys: () => Object.keys(secrets),
  };

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
