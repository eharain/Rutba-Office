// Settings, secrets and the recent list — three small JSON files in userData.
//
// Secrets go through Electron's safeStorage, which uses the OS keychain
// (DPAPI on Windows, Keychain on macOS, libsecret on Linux). When the OS has no
// keystore available, we refuse to write rather than quietly saving a password
// in clear text: the caller is told, and can decide to ask every time instead.

import { app, safeStorage } from 'electron';
import fs from 'node:fs';
import path from 'node:path';

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

export function createStores() {
  const dir = app.getPath('userData');
  const settingsFile = path.join(dir, 'settings.json');
  const secretsFile = path.join(dir, 'secrets.json');
  const recentFile = path.join(dir, 'recent.json');

  let settings = readJson(settingsFile, {});
  let secrets = readJson(secretsFile, {});
  let recent = readJson(recentFile, []);

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
      writeJson(secretsFile, secrets);
    },
    delete: (key) => {
      delete secrets[key];
      writeJson(secretsFile, secrets);
    },
    keys: () => Object.keys(secrets),
  };

  const recentApi = {
    list: () => recent.filter((r) => fs.existsSync(r.path)).slice(0, 30),
    add: ({ path: p, app: which }) => {
      if (!p) return recentApi.list();
      recent = [
        { path: p, name: path.basename(p), app: which || null, at: Date.now() },
        ...recent.filter((r) => r.path !== p),
      ].slice(0, 60);
      writeJson(recentFile, recent);
      try {
        app.addRecentDocument(p);
      } catch { /* jump list is a nicety, not a requirement */ }
      return recentApi.list();
    },
    clear: () => {
      recent = [];
      writeJson(recentFile, recent);
      try { app.clearRecentDocuments(); } catch { /* as above */ }
      return [];
    },
  };

  return { settings: settingsApi, secrets: secretsApi, recent: recentApi, dir };
}
