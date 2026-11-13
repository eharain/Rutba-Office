// Touch ID on a Mac, for "this computer only" unlocking (owner, 13 Nov,
// D-2026-11-13-22: Touch ID on a Mac, off by default; Linux has the password
// and the recovery key only; Windows Hello stays as it is).
//
// It answers the same calls as Windows Hello (hello.js), so the keyring keeps
// one quick way in, of the same kind, on each system. Turning it on makes a
// random 32-byte secret, kept sealed by the Keychain (Electron's safeStorage),
// that wraps the locker's key as Windows Hello's secret does. To unlock, Touch
// ID is asked first, and only then is the secret opened.
//
// This is a check that the person is at the Mac, on top of the Keychain, not a
// key held by the fingerprint sensor itself: anything that can read this Mac
// account's Keychain without Touch ID could open the secret. The setting says
// so. The password remains the master.

import crypto from 'node:crypto';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { LockerError } from './format.js';

const ID = /^[0-9a-f]{16}$/;

/** What went wrong, said for the person. A cancelled prompt has the same code as Windows Hello's, so the window stays quiet about it. */
export function touchIdFailure(err) {
  const said = String(err?.message || err || '');
  if (/cancel/i.test(said)) return new LockerError('hello-cancelled', 'Touch ID was cancelled.');
  if (/lockout|locked/i.test(said)) return new LockerError('hello-locked', 'Touch ID is locked after too many tries. Use your password, or try again after unlocking your Mac with its password.');
  return new LockerError('hello-failed', 'Touch ID did not finish. Use your password instead.');
}

const GONE = 'This Mac no longer has the Touch ID key for this locker. Use your password, then turn Touch ID on again in Locker settings.';

/**
 * @param {object} o
 * @param {object} o.systemPreferences  Electron's: canPromptTouchID, promptTouchID
 * @param {object} o.safeStorage        Electron's: isEncryptionAvailable, encryptString, decryptString
 * @param {string} o.dir                the Locker's profile folder
 * @param {string} [o.platform]
 */
export function touchIdProvider({ systemPreferences, safeStorage, dir, platform = process.platform }) {
  const fileOf = (id) => path.join(dir, `touchid-${id}.key`);

  const available = async () => {
    if (platform !== 'darwin') return false;
    try {
      return Boolean(systemPreferences?.canPromptTouchID?.()) && Boolean(safeStorage?.isEncryptionAvailable?.());
    } catch {
      return false;
    }
  };

  async function ask(reason) {
    try {
      await systemPreferences.promptTouchID(reason);
    } catch (err) {
      throw touchIdFailure(err);
    }
  }

  return {
    available,
    /** Touch ID turned on: asked once, so the person knows it works, then a new secret sealed by the Keychain. */
    enroll: async (id) => {
      if (!ID.test(String(id)) || !(await available())) throw new LockerError('hello-failed', 'Touch ID is not available on this Mac.');
      await ask('turn on Touch ID for your Locker');
      const secret = crypto.randomBytes(32);
      const sealed = safeStorage.encryptString(secret.toString('base64'));
      await fsp.mkdir(dir, { recursive: true });
      await fsp.writeFile(fileOf(id), sealed, { mode: 0o600 });
      return secret;
    },
    /** The secret, once Touch ID says the person is here. */
    secret: async (id) => {
      if (!ID.test(String(id))) throw new LockerError('hello-gone', GONE);
      if (!(await available())) throw new LockerError('hello-failed', 'Touch ID is not available on this Mac now. Use your password instead.');
      await ask('unlock your Locker');
      let sealed;
      try {
        sealed = await fsp.readFile(fileOf(id));
      } catch {
        throw new LockerError('hello-gone', GONE);
      }
      try {
        return Buffer.from(safeStorage.decryptString(sealed), 'base64');
      } catch {
        throw new LockerError('hello-gone', GONE);
      }
    },
    forget: async (id) => {
      if (ID.test(String(id))) await fsp.rm(fileOf(id), { force: true }).catch(() => {});
    },
  };
}

/** A system with no quick way in (Linux): the password and the recovery key only. */
export const noQuickUnlock = () => ({
  available: async () => false,
  enroll: async () => { throw new LockerError('hello-failed', 'This system has no quick unlock. Use your password.'); },
  secret: async () => { throw new LockerError('hello-failed', 'This system has no quick unlock. Use your password.'); },
  forget: async () => {},
});

/** Which quick way in this system offers, by name: 'hello' on Windows, 'touchid' on a Mac, null on Linux. */
export const quickUnlockName = (platform = process.platform) => (platform === 'win32' ? 'hello' : platform === 'darwin' ? 'touchid' : null);
