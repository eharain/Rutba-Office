// Windows Hello for the Locker, as far as a check can take it without a person
// at the keyboard: what each Windows answer becomes, that a key is named for one
// locker, the real system's answers up to the point of a prompt, and that a
// locker opened this way can do what one opened with the password can.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { helloProvider } from '../apps/desktop/main/extras/locker/hello.js';
import { createKeys } from '../apps/desktop/main/extras/locker/keys.js';

const CHEAP = { N: 1024, r: 8, p: 1 };
const ID = '0123456789abcdef';
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'rl-hello-'));
const answering = (out) => helloProvider({ run: async () => out });

test('the secret is the hash of what Windows signed, the same each time', async () => {
  const signature = crypto.randomBytes(256);
  const h = answering(signature.toString('base64'));
  const a = await h.secret(ID);
  assert.deepEqual(a, crypto.createHash('sha256').update(signature).digest());
  assert.deepEqual(await h.secret(ID), a);
});

for (const [status, code] of [
  ['UserCanceled', 'hello-cancelled'],
  ['UserPrefersPassword', 'hello-cancelled'],
  ['NotFound', 'hello-gone'],
  ['SecurityDeviceLocked', 'hello-locked'],
  ['UnknownError', 'hello-failed'],
]) {
  test(`Windows answering ${status} is ${code}, in words without the status`, async () => {
    await assert.rejects(() => answering(`status:${status}`).secret(ID), (e) => e.code === code && !e.message.includes(status));
  });
}

test('PowerShell failing or answering nothing is a failure, never an empty secret', async () => {
  await assert.rejects(() => answering('').secret(ID), (e) => e.code === 'hello-failed');
  const broken = helloProvider({ run: async () => { throw new Error('blocked by policy'); } });
  await assert.rejects(() => broken.enroll(ID), (e) => e.code === 'hello-failed');
});

test('a key is named for its locker, and an id that is not one never reaches the script', async () => {
  const bodies = [];
  const h = helloProvider({ run: async (body) => { bodies.push(body); return 'AAAA'; } });
  await h.enroll(ID);
  await h.secret('fedcba9876543210');
  assert.match(bodies[0], /RequestCreateAsync\('RutbaOfficeLocker-0123456789abcdef'/);
  assert.match(bodies[1], /OpenAsync\('RutbaOfficeLocker-fedcba9876543210'/);
  const before = bodies.length;
  for (const bad of [null, undefined, '', "x'); Remove-Item C:\\ #", 'ABCDEF0123456789']) {
    await assert.rejects(() => h.secret(bad), (e) => e.code === 'hello-gone');
    await h.forget(bad);
  }
  assert.equal(bodies.length, before);
});

test('asking whether Windows Hello is set up: a yes is kept, a no is asked again later', { skip: process.platform !== 'win32' }, async (t) => {
  let asked = 0;
  let answer = 'False';
  const h = helloProvider({ run: async () => { asked += 1; return answer; } });
  t.mock.timers.enable({ apis: ['Date'], now: 1000000 });
  assert.equal(await h.available(), false);
  assert.equal(await h.available(), false);
  assert.equal(asked, 1);
  answer = 'True';
  t.mock.timers.tick(61000);
  assert.equal(await h.available(), true);
  t.mock.timers.tick(10000000);
  assert.equal(await h.available(), true);
  assert.equal(asked, 2);
});

test('the real system: its answers up to a prompt, with no prompt shown', { skip: process.platform !== 'win32' }, async (t) => {
  const h = helloProvider();
  if (!(await h.available())) return t.skip('Windows Hello is not set up on this computer');
  // A key that was never made: the whole script runs, the window helper compiles, and Windows says NotFound before any prompt.
  const never = crypto.randomBytes(8).toString('hex');
  await assert.rejects(() => h.secret(never), (e) => e.code === 'hello-gone');
  await h.forget(never);
});

test('a locker opened with Windows Hello can do what one opened with the password can, and a new password keeps it', async () => {
  const dir = tmp();
  const k = createKeys({ dir, cost: CHEAP });
  await k.setup('first password');
  const secret = crypto.randomBytes(32);
  await k.enableHello(secret, ID);
  assert.equal(await k.helloId(), ID);

  k.lock();
  await k.unlock({ secret });
  const fk = crypto.randomBytes(32);
  const file = k.wrapForFolder(fk, 'folder-a');
  assert.deepEqual(await k.openFolderKeys(file, { password: 'first password' }), fk);

  await k.changePassword('first password', 'second password');
  k.lock();
  await k.unlock({ secret });
  const later = k.wrapForFolder(fk, 'folder-b');
  assert.deepEqual(await k.openFolderKeys(later, { password: 'second password' }), fk);
  assert.equal(await k.helloId(), ID);

  await k.disableHello();
  assert.equal(await k.helloId(), null);
});
