// The keys a release's add-ons contract trusts, checked before it is built
// (apps/desktop/build/dist.js runs this first; D-2026-11-13-28).
//
// A release must trust a packages key, or no add-on, feed or English copy can
// ever be verified and every update that counts one waits for good. Each key
// in it must be a public key and nothing else: an Ed25519 SPKI PEM, never a
// private half, never anything that merely parses, and never the test key
// the shared cases are signed with. The packages key is pinned by its id:
// another key in its place, the right shape or not, refuses the build. A new
// packages key is a decision, made here as well as in the contract.
//
//   node tools/check-contract-keys.mjs [contract.json]

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const CONTRACT_FILE = path.join(ROOT, 'packages', 'office-shell', 'src', 'addons-contract.json');
export const TEST_KEY_FILE = path.join(ROOT, 'packages', 'office-shell', 'test-vectors', 'addons', 'test-key.pub');

/** The key ids each role must list, exactly: keyIdOf (packages/office-shell/src/electron/packs.js). */
export const PINNED_KEYS = Object.freeze({ packages: Object.freeze(['c693adde39776bf8']) });

/** A key's short name, as keyIdOf answers it: the first 16 hex digits of the SHA-256 of its SPKI. */
export const keyIdOfPem = (pem) => crypto.createHash('sha256').update(crypto.createPublicKey(pem).export({ type: 'spki', format: 'der' })).digest('hex').slice(0, 16);

const PUBLIC_PEM = /^-----BEGIN PUBLIC KEY-----\r?\n[A-Za-z0-9+/=\r\n]+-----END PUBLIC KEY-----\r?\n?$/;

/** What is wrong with one key a contract lists, or null. */
function keyFault(key, { testKey = null } = {}) {
  if (typeof key !== 'string') return 'is not a PEM string';
  if (/PRIVATE/i.test(key)) return 'is a private key: only the public half goes in the contract';
  if (!PUBLIC_PEM.test(key.trim() + '\n')) return 'is not a public key PEM (-----BEGIN PUBLIC KEY-----)';
  let k;
  try {
    k = crypto.createPublicKey(key);
  } catch (err) {
    return `cannot be read as a public key (${err?.message || err})`;
  }
  if (k.type !== 'public' || k.asymmetricKeyType !== 'ed25519') return `is a ${k.asymmetricKeyType || k.type} key, not Ed25519`;
  if (testKey && k.export({ type: 'spki', format: 'der' }).equals(testKey.export({ type: 'spki', format: 'der' }))) return 'is the test key the shared cases are signed with';
  return null;
}

/**
 * The faults of a contract's keys for a release: none when it may be built.
 * @param {object} contract
 * @param {{ testKey?: crypto.KeyObject|null, pinned?: Record<string, string[]>|null }} [o]
 *   `pinned`: the key ids a role must list, exactly (PINNED_KEYS for a release)
 * @returns {string[]}
 */
export function contractKeyFaults(contract, { testKey = null, pinned = null } = {}) {
  const faults = [];
  const keys = contract?.keys;
  if (!keys || typeof keys !== 'object') return ['the contract has no keys'];
  if (!Array.isArray(keys.packages) || !keys.packages.length) faults.push('keys.packages is empty: a release must trust the packages public key');
  for (const [role, list] of Object.entries(keys)) {
    if (!Array.isArray(list)) { faults.push(`keys.${role} is not a list`); continue; }
    list.forEach((key, i) => {
      const f = keyFault(key, { testKey });
      if (f) faults.push(`keys.${role}[${i}] ${f}`);
    });
    if (new Set(list).size !== list.length) faults.push(`keys.${role} lists a key twice`);
  }
  for (const [role, want] of Object.entries(pinned || {})) {
    const list = Array.isArray(keys[role]) ? keys[role] : [];
    const ids = list.map((k) => {
      try {
        return keyIdOfPem(k);
      } catch {
        return null;
      }
    });
    const missing = want.filter((id) => !ids.includes(id));
    const other = ids.filter((id) => id && !want.includes(id));
    if (missing.length || other.length) faults.push(`keys.${role} must be exactly ${want.join(', ')}${other.length ? `; it lists ${other.join(', ')}` : ''}${missing.length ? `; ${missing.join(', ')} is not there` : ''}`);
  }
  return faults;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const file = path.resolve(process.argv[2] || CONTRACT_FILE);
  const testKey = fs.existsSync(TEST_KEY_FILE) ? crypto.createPublicKey(fs.readFileSync(TEST_KEY_FILE)) : null;
  const faults = contractKeyFaults(JSON.parse(fs.readFileSync(file, 'utf8')), { testKey, pinned: PINNED_KEYS });
  if (faults.length) {
    for (const f of faults) console.error(`${path.relative(ROOT, file)}: ${f}`);
    process.exit(1);
  }
  console.log(`${path.relative(ROOT, file)}: its keys may ship`);
}
