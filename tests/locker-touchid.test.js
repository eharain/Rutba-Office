// Touch ID on a Mac as the Locker's quick way in (owner, 13 Nov,
// D-2026-11-13-22): asked before the secret is opened, kept sealed by the
// Keychain, off until turned on; and no quick way in on Linux. Electron's
// systemPreferences and safeStorage are stood in for: these run on any system.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { tempDir } from './temp-dirs.js';
import { touchIdProvider, noQuickUnlock, quickUnlockName } from '../apps/desktop/main/extras/locker/touchid.js';
import { createLocker } from '../apps/desktop/main/extras/locker/locker.js';

const CHEAP = { N: 1024, r: 8, p: 1 };

function fakes({ can = true, answer = () => Promise.resolve() } = {}) {
  const asked = [];
  const keychain = crypto.randomBytes(32);
  return {
    asked,
    systemPreferences: {
      canPromptTouchID: () => can,
      promptTouchID: (reason) => {
        asked.push(reason);
        return answer(reason);
      },
    },
    safeStorage: {
      isEncryptionAvailable: () => true,
      // A stand-in seal: what it writes is not the secret's text.
      encryptString: (text) => Buffer.from(Buffer.from(text).map((b, i) => b ^ keychain[i % 32])),
      decryptString: (buf) => Buffer.from(Buffer.from(buf).map((b, i) => b ^ keychain[i % 32])).toString(),
    },
  };
}

const ID = '0123456789abcdef';

test('each system has its own quick way in: Windows Hello, Touch ID, and none on Linux', async () => {
  assert.equal(quickUnlockName('win32'), 'hello');
  assert.equal(quickUnlockName('darwin'), 'touchid');
  assert.equal(quickUnlockName('linux'), null);
  assert.equal(await noQuickUnlock().available(), false);
  await assert.rejects(noQuickUnlock().secret(ID), (err) => err.code === 'hello-failed');
});

test('Touch ID is offered only on a Mac that can ask for it', async () => {
  const dir = tempDir('rl-touch-');
  assert.equal(await touchIdProvider({ ...fakes(), dir, platform: 'linux' }).available(), false);
  assert.equal(await touchIdProvider({ ...fakes({ can: false }), dir, platform: 'darwin' }).available(), false);
  assert.equal(await touchIdProvider({ ...fakes(), dir, platform: 'darwin' }).available(), true);
});

test('turning Touch ID on asks for it, and keeps the secret sealed; unlocking asks before the secret is opened', async () => {
  const dir = tempDir('rl-touch-');
  const f = fakes();
  const t = touchIdProvider({ ...f, dir, platform: 'darwin' });
  const secret = await t.enroll(ID);
  assert.equal(secret.length, 32);
  assert.equal(f.asked.length, 1);
  const kept = fs.readFileSync(path.join(dir, `touchid-${ID}.key`));
  assert.ok(!kept.includes(Buffer.from(secret.toString('base64'))), 'the file is not the secret');
  const again = await t.secret(ID);
  assert.deepEqual(again, secret);
  assert.equal(f.asked.length, 2);
  await t.forget(ID);
  await assert.rejects(t.secret(ID), (err) => err.code === 'hello-gone');
});

test('a cancelled Touch ID is quiet, and a refused one opens nothing', async () => {
  const dir = tempDir('rl-touch-');
  let say = null;
  const f = fakes({ answer: () => (say ? Promise.reject(new Error(say)) : Promise.resolve()) });
  const t = touchIdProvider({ ...f, dir, platform: 'darwin' });
  await t.enroll(ID);
  say = 'User canceled the authentication';
  await assert.rejects(t.secret(ID), (err) => err.code === 'hello-cancelled');
  say = 'Biometry is locked out';
  await assert.rejects(t.secret(ID), (err) => err.code === 'hello-locked');
  say = 'something else';
  await assert.rejects(t.secret(ID), (err) => err.code === 'hello-failed');
});

test('a locker opened by Touch ID: the secret it keeps unlocks, and the password still does', async () => {
  const dir = tempDir('rl-touch-');
  const locker = createLocker({ dir, cost: CHEAP });
  await locker.setup('pw');
  const t = touchIdProvider({ ...fakes(), dir, platform: 'darwin' });
  const secret = await t.enroll(ID);
  await locker.enableHello(secret, ID);
  assert.ok((await locker.ways()).includes('hello'));
  await locker.lock();
  await locker.unlock({ secret: await t.secret(ID) });
  assert.ok(locker.isOpen());
  await locker.lock();
  await locker.unlock({ password: 'pw' });
  assert.ok(locker.isOpen());
});
