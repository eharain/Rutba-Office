/**
 * The passwords a workbook keeps: how Excel turns one into what it writes,
 * and how a password typed later is checked against it.
 *
 * Excel writes two kinds, and a file from any year may carry either.
 *
 *   - The modern hash (ECMA-376 Part 1, 18.2.29 and 18.3.1.85): SHA-512 of
 *     the salt followed by the password as UTF-16LE, then `spinCount` more
 *     rounds, each the hash of the previous one followed by the round's
 *     number as four little-endian bytes. Written base64 as `hashValue`,
 *     beside `saltValue`, `spinCount` and `algorithmName` — the workbook's
 *     own element prefixes each with `workbook`.
 *
 *   - The legacy 16-bit hash (`password`, `workbookPassword`): each
 *     character's code rotated left within 15 bits by its position, the
 *     results XORed together with the length and 0xCE4B, written as four
 *     hex digits. Weak, and read so a file from Excel 2007 can still be
 *     unlocked with its password; never written.
 *
 * `node:crypto` is used, so this is for the engine's side, not a window's.
 */
import { createHash, randomBytes } from 'node:crypto';

/** Excel's round count for a new password — the value Excel 2013 onward writes. */
export const SPIN_COUNT = 100000;

const ALGORITHMS = { 'SHA-512': 'sha512', 'SHA-384': 'sha384', 'SHA-256': 'sha256', 'SHA-1': 'sha1', MD5: 'md5' };

/**
 * The modern hash of a password: base64 of the last round.
 *
 * @param {string} password
 * @param {{ salt: string, spinCount?: number, algorithm?: string }} spec the salt as base64
 */
export function hashPassword(password, { salt, spinCount = SPIN_COUNT, algorithm = 'SHA-512' }) {
  const name = ALGORITHMS[algorithm] || ALGORITHMS[String(algorithm).toUpperCase()];
  if (!name) throw new Error('a password hashed with ' + algorithm + ' cannot be checked here');
  const iterator = Buffer.alloc(4);
  let h = createHash(name).update(Buffer.from(salt, 'base64')).update(Buffer.from(String(password), 'utf16le')).digest();
  for (let i = 0; i < spinCount; i++) {
    iterator.writeUInt32LE(i, 0);
    h = createHash(name).update(h).update(iterator).digest();
  }
  return h.toString('base64');
}

/** The legacy 16-bit hash, as the four upper-case hex digits Excel writes. */
export function legacyPasswordHash(password) {
  const text = String(password);
  let hash = 0;
  for (let i = 0; i < text.length; i++) {
    let v = text.charCodeAt(i) & 0x7fff;
    // Rotate left within 15 bits, once per position (the first character once).
    for (let k = 0; k <= i; k++) v = ((v << 1) & 0x7fff) | (v >> 14);
    hash ^= v;
  }
  hash ^= text.length;
  hash ^= 0xce4b;
  return hash.toString(16).toUpperCase().padStart(4, '0');
}

/**
 * The attributes a new password is written with: a fresh 16-byte salt,
 * SHA-512 and Excel's round count, under the names the element uses
 * (`prefix` is 'workbook' for `<workbookProtection>`, '' elsewhere).
 */
export function passwordAttrs(password, { prefix = '', salt = randomBytes(16).toString('base64'), spinCount = SPIN_COUNT } = {}) {
  const key = (name) => (prefix ? prefix + name[0].toUpperCase() + name.slice(1) : name);
  return {
    [key('algorithmName')]: 'SHA-512',
    [key('hashValue')]: hashPassword(password, { salt, spinCount }),
    [key('saltValue')]: salt,
    [key('spinCount')]: String(spinCount),
  };
}

/**
 * Whether an element's attributes carry a password at all — modern or legacy.
 * `legacy` names the legacy attribute (`password`, `workbookPassword`).
 */
export function hasPassword(a, { prefix = '', legacy = 'password' } = {}) {
  if (!a) return false;
  const key = (name) => (prefix ? prefix + name[0].toUpperCase() + name.slice(1) : name);
  return Boolean(a[key('hashValue')] || a[legacy]);
}

/**
 * Does `password` open what these attributes lock? True when they carry no
 * password. The modern hash is checked when present, else the legacy one.
 */
export function checkPassword(password, a, { prefix = '', legacy = 'password' } = {}) {
  if (!hasPassword(a, { prefix, legacy })) return true;
  const key = (name) => (prefix ? prefix + name[0].toUpperCase() + name.slice(1) : name);
  const hash = a[key('hashValue')];
  if (hash) {
    const spin = Number(a[key('spinCount')] ?? 0) || 0;
    try {
      return hashPassword(String(password ?? ''), { salt: a[key('saltValue')] || '', spinCount: spin, algorithm: a[key('algorithmName')] || 'SHA-512' }) === hash;
    } catch {
      return false;
    }
  }
  return legacyPasswordHash(String(password ?? '')) === String(a[legacy]).toUpperCase().padStart(4, '0');
}
