// The mail store's files sealed with a profile's key (D-2026-10-26-02).
//
// AES-256-GCM, a fresh 12-byte nonce for every file and every line, the key
// the profile's password makes (office-shell profiles.js makeLock/unlock).
// The same cipher the profile's other files are sealed with, written in a
// form that suits mail: binary for a whole file (a message, an attachment,
// a list), so a large attachment is not made a third larger by base64; and
// one line at a time for a folder's index, so a message arriving is still
// one line appended, not the whole index sealed again.
//
//   a whole file   MAGIC, nonce (12), tag (16), the sealed bytes
//   an index line  "~" and the base64 of nonce, tag and the sealed bytes
//
// A plain file never starts with MAGIC (JSON starts with { or [, and a NUL
// first is what no stored text has), and a plain index line starts with {,
// so plain and sealed sit side by side while a store is being sealed and
// each is read as what it is.

import crypto from 'node:crypto';

export const MAGIC = Buffer.from('\u0000RSM1', 'latin1');
const AAD = Buffer.from('rutba-office-mail');
const IV = 12;
const TAG = 16;

/** The error a sealed file gives when no key held opens it. */
export class LockedError extends Error {
  constructor() {
    super('This mail is sealed with the profile’s password.');
    this.code = 'MAIL_LOCKED';
  }
}

export const isSealed = (buf) => Buffer.isBuffer(buf) && buf.length >= MAGIC.length + IV + TAG && buf.subarray(0, MAGIC.length).equals(MAGIC);
export const isSealedLine = (line) => typeof line === 'string' && line.charCodeAt(0) === 126; // "~"

function encrypt(key, bytes) {
  const iv = crypto.randomBytes(IV);
  const c = crypto.createCipheriv('aes-256-gcm', key, iv);
  c.setAAD(AAD);
  const data = Buffer.concat([c.update(bytes), c.final()]);
  return [iv, c.getAuthTag(), data];
}

function decrypt(key, iv, tag, data) {
  const d = crypto.createDecipheriv('aes-256-gcm', key, iv);
  d.setAAD(AAD);
  d.setAuthTag(tag);
  return Buffer.concat([d.update(data), d.final()]);
}

/** Bytes sealed whole. */
export function sealBytes(key, bytes) {
  const [iv, tag, data] = encrypt(key, Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes));
  return Buffer.concat([MAGIC, iv, tag, data]);
}

/** Sealed bytes opened with the first of `keys` that is theirs; LockedError when none is. */
export function openBytes(keys, sealed) {
  const iv = sealed.subarray(MAGIC.length, MAGIC.length + IV);
  const tag = sealed.subarray(MAGIC.length + IV, MAGIC.length + IV + TAG);
  const data = sealed.subarray(MAGIC.length + IV + TAG);
  for (const key of keys) {
    if (!key) continue;
    try {
      return decrypt(key, iv, tag, data);
    } catch {
      /* the next key */
    }
  }
  throw new LockedError();
}

/** One line of text sealed, as an index keeps it. */
export function sealLine(key, text) {
  const [iv, tag, data] = encrypt(key, Buffer.from(text, 'utf8'));
  return `~${Buffer.concat([iv, tag, data]).toString('base64')}`;
}

/** A sealed line opened; LockedError when no key is its. */
export function openLine(keys, line) {
  const raw = Buffer.from(line.slice(1), 'base64');
  const iv = raw.subarray(0, IV);
  const tag = raw.subarray(IV, IV + TAG);
  const data = raw.subarray(IV + TAG);
  for (const key of keys) {
    if (!key) continue;
    try {
      return decrypt(key, iv, tag, data).toString('utf8');
    } catch {
      /* the next key */
    }
  }
  throw new LockedError();
}

/** Which of `keys` a sealed file or line is sealed with (its index), or -1; for a plain one, null. */
export function sealedWith(keys, content) {
  const line = typeof content === 'string';
  if (line ? !isSealedLine(content) : !isSealed(content)) return null;
  for (let i = 0; i < keys.length; i++) {
    if (!keys[i]) continue;
    try {
      if (line) openLine([keys[i]], content);
      else openBytes([keys[i]], content);
      return i;
    } catch {
      /* not this one */
    }
  }
  return -1;
}
