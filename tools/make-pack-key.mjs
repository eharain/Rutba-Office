// A key pair for signing app packs (docs/APP-PACKS.md, section 3).
//
//   node tools/make-pack-key.mjs <folder>
//
// Writes <folder>/pack-key.pem, the private half (PKCS #8), and
// <folder>/pack-key.pub.pem, the public half the core carries (SPKI). The
// folder must be outside this repository: the real key is made once and
// kept with the other release secrets, and a test makes a throwaway one in
// a temporary folder. Nothing is overwritten. Only the public half and its
// id are printed.

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { keyIdOf } from '../packages/office-shell/src/electron/packs.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/** Whether a path is inside this repository, where no private key may be. */
export function insideRepository(p) {
  const r = path.relative(ROOT, path.resolve(p));
  return r === '' || (!r.startsWith('..') && !path.isAbsolute(r));
}

/** A new Ed25519 key pair, as PEM text, with the public half's id. */
export function makeKeyPair() {
  const { privateKey, publicKey } = crypto.generateKeyPairSync('ed25519');
  return {
    privatePem: privateKey.export({ format: 'pem', type: 'pkcs8' }),
    publicPem: publicKey.export({ format: 'pem', type: 'spki' }),
    keyId: keyIdOf(publicKey),
  };
}

/** Writes a new key pair into a folder outside the repository. */
export function writeKeyPair(folder) {
  if (insideRepository(folder)) throw new Error('A pack key is never kept inside the repository: choose a folder outside it.');
  const priv = path.join(folder, 'pack-key.pem');
  const pub = path.join(folder, 'pack-key.pub.pem');
  for (const f of [priv, pub]) if (fs.existsSync(f)) throw new Error(`${f} already exists; a key is not overwritten.`);
  fs.mkdirSync(folder, { recursive: true });
  const pair = makeKeyPair();
  fs.writeFileSync(priv, pair.privatePem, { mode: 0o600, flag: 'wx' });
  fs.writeFileSync(pub, pair.publicPem, { flag: 'wx' });
  return { privateFile: priv, publicFile: pub, publicPem: pair.publicPem, keyId: pair.keyId };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const folder = process.argv[2];
  if (!folder) {
    console.error('Usage: node tools/make-pack-key.mjs <folder outside the repository>');
    process.exit(2);
  }
  try {
    const made = writeKeyPair(folder);
    console.log(`The private key is in ${made.privateFile}. Keep it with the release secrets; it is not printed.`);
    console.log(`The public key, id ${made.keyId}, is in ${made.publicFile}:\n${made.publicPem}`);
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }
}
