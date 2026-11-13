// The keys a release's add-ons contract trusts (tools/check-contract-keys.mjs,
// run first by the release build): a packages key there, and every key a
// public Ed25519 key that is not the shared cases' test key.

import test from 'node:test';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import { contractKeyFaults, TEST_KEY_FILE } from '../tools/check-contract-keys.mjs';

const pub = (k) => k.publicKey.export({ type: 'spki', format: 'pem' }).toString();
const real = crypto.generateKeyPairSync('ed25519');
const testKey = crypto.createPublicKey(fs.readFileSync(TEST_KEY_FILE));

test('a release contract trusts a packages public key, and only public keys', () => {
  assert.deepEqual(contractKeyFaults({ keys: { packages: [pub(real)], release: [] } }, { testKey }), []);
  assert.match(contractKeyFaults({ keys: { packages: [], release: [] } }).join(), /keys\.packages is empty/);
  assert.match(contractKeyFaults({}).join(), /no keys/);
  const priv = real.privateKey.export({ type: 'pkcs8', format: 'pem' }).toString();
  assert.match(contractKeyFaults({ keys: { packages: [priv] } }).join(), /private key/);
  assert.match(contractKeyFaults({ keys: { packages: [pub(crypto.generateKeyPairSync('rsa', { modulusLength: 1024 }))] } }).join(), /rsa key, not Ed25519/);
  assert.match(contractKeyFaults({ keys: { packages: ['hello'] } }).join(), /not a public key PEM/);
  assert.match(contractKeyFaults({ keys: { packages: [7] } }).join(), /not a PEM string/);
  assert.match(contractKeyFaults({ keys: { packages: [pub(real), pub(real)] } }).join(), /twice/);
  assert.match(contractKeyFaults({ keys: { packages: [testKey.export({ type: 'spki', format: 'pem' }).toString()] } }, { testKey }).join(), /test key/);
  assert.match(contractKeyFaults({ keys: { packages: [pub(real)], release: ['x'] } }).join(), /keys\.release\[0\]/);
});
