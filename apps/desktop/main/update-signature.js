// Whether a downloaded update is one the makers of this program signed.
//
// The installer's own SHA-512 comes from latest.yml, which sits in the same
// GitHub release as the installer, so anyone who gets hold of the repository's
// token can change both and the check still passes. Windows code signing would
// answer that and this project has no certificate yet. So each installer is
// signed with a key that is never on GitHub and never in Actions, and the
// public halves are fixed here, in the copy a person already has.
//
// What is signed names the product, the version and the file as well as the
// bytes, so an old signed installer cannot be offered again as a new version.
//
// There are two keys so that the loss of one does not end updates for everyone:
// each copy accepts either.

import crypto from 'node:crypto';
import fs from 'node:fs';

export const UPDATE_KEYS = [
  `-----BEGIN PUBLIC KEY-----
MCowBQYDK2VwAyEAmrWY1Hy8mJmgCEYIJ/JpRb2qO98h3Lcb2l14XJkytsk=
-----END PUBLIC KEY-----
`,
  `-----BEGIN PUBLIC KEY-----
MCowBQYDK2VwAyEAon9gQlBB+rusUwVXhIXcwzZc3hiqBGjtiQx3HNc4WnA=
-----END PUBLIC KEY-----
`,
];

/** An Ed25519 signature is 64 bytes; base64 and a line ending stay well inside this. */
const MOST_SIGNATURE_TEXT = 256;
const ED25519_BYTES = 64;

/** The words a signature is over. */
export function signedWords({ version, file, sha512 }) {
  return `Rutba Office update\nversion: ${version}\nfile: ${file}\nsha512: ${sha512}\n`;
}

export function sha512OfFile(filePath) {
  return new Promise((resolve, reject) => {
    const hash = crypto.createHash('sha512');
    fs.createReadStream(filePath).on('error', reject).on('data', (chunk) => hash.update(chunk)).on('end', () => resolve(hash.digest('base64')));
  });
}

/** Whether this signature, as the text of a .sig file, is good for these words under one of the keys. */
export function signatureHolds(words, signatureText, keys = UPDATE_KEYS) {
  const text = String(signatureText ?? '').trim();
  if (!text || text.length > MOST_SIGNATURE_TEXT || !/^[A-Za-z0-9+/]+={0,2}$/.test(text)) return false;
  const signature = Buffer.from(text, 'base64');
  if (signature.length !== ED25519_BYTES) return false;
  return keys.some((pem) => {
    try {
      return crypto.verify(null, Buffer.from(words, 'utf8'), pem, signature);
    } catch {
      return false;
    }
  });
}

/** Where a version's signature is published, beside the installer. */
export function signatureAddress({ owner, repo, version, file }) {
  return `https://github.com/${owner}/${repo}/releases/download/v${encodeURIComponent(version)}/${encodeURIComponent(file)}.sig`;
}

/**
 * Why this downloaded file may not be installed, as a sentence, or null when
 * it may. Anything that stops the check stops the update: it comes round
 * again at the next check, and an update that waits does no harm.
 */
export async function whyNotOurs({ filePath, version, file, fetchSignature, keys = UPDATE_KEYS }) {
  if (!version || !file) return 'This update does not say what it is, so it was not installed.';
  let text;
  try {
    text = await fetchSignature({ version, file });
  } catch {
    return 'The signature of this update could not be fetched, so it was not installed.';
  }
  let sha512;
  try {
    sha512 = await sha512OfFile(filePath);
  } catch {
    return 'The downloaded update could not be read, so it was not installed.';
  }
  return signatureHolds(signedWords({ version, file, sha512 }), text, keys) ? null : 'This update is not signed by the makers of Rutba Office, so it was not installed.';
}

/** The installer a Windows feed names, by its file name. */
export function installerIn(info) {
  const url = (info?.files || []).map((f) => f?.url).find((u) => /\.exe$/i.test(String(u || ''))) ?? info?.path;
  return url ? decodeURIComponent(String(url)).split('/').pop() : null;
}

/**
 * Puts the check where electron-updater looks for one. Its Windows updater
 * calls `verifySignature` with the finished download and refuses it when the
 * answer is not null.
 */
export function guardDownloads(updater, { current, fetchSignature, keys = UPDATE_KEYS }) {
  updater.verifySignature = async (filePath) => {
    const info = current();
    return whyNotOurs({ filePath, version: info?.version, file: installerIn(info), fetchSignature, keys });
  };
}

/**
 * electron-updater sends a lasting random id with each request, for staged
 * rollouts, which are not used here. Headers set on the updater are written
 * over it, so every copy sends the same word.
 */
export function keepAnonymous(updater) {
  updater.requestHeaders = { ...updater.requestHeaders, 'x-user-staging-id': 'rutba-office' };
}
