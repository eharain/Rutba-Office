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

/* ── Word: w:documentProtection ─────────────────────────────────────────────
 *
 * Word hashes a Restrict Editing password with the same salted, spun hash —
 * but not of the password itself. It first takes the 32-bit key Word 97
 * made from a password ([MS-OFFCRYPTO] 2.3.7.4, "Method 2": a high word from
 * the initial code table and the encryption matrix, the low word the 16-bit
 * verifier of 2.3.7.1), writes its four bytes low byte first as eight hex
 * digits, and hashes THAT string as the password ([MS-OI29500] on
 * documentProtection). The attributes are Word's own: cryptProviderType,
 * cryptAlgorithmClass/Type/Sid (14 is SHA-512), cryptSpinCount, hash, salt.
 */

const WORD_INITIAL_CODE = [0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9, 0x280c, 0xa96a, 0x4ec3];

/**
 * The encryption matrix: fifteen rows of seven. Each entry is the one before
 * it shifted left once in sixteen bits, XORed with 0x1021 when a bit falls
 * off the top — so the table is its first column and that rule, which is
 * how it is written here rather than as 105 numbers to mistype.
 */
const WORD_MATRIX = [0xaefc, 0x7b61, 0x4563, 0x0375, 0xd849, 0x6f45, 0xeb23, 0x47d3, 0xb861, 0x45a0, 0xaa51, 0x76b4, 0x3730, 0x3331, 0x1021].map((head) => {
  const row = [head];
  for (let i = 1; i < 7; i++) {
    let v = row[i - 1] << 1;
    if (v & 0x10000) v = (v & 0xffff) ^ 0x1021;
    row.push(v);
  }
  return row;
});

export { WORD_MATRIX, WORD_INITIAL_CODE };

/** A password's characters as Word 97 took them: one byte each, the low byte or, if that is 0, the high one. */
function ansiBytes(password) {
  const text = String(password ?? '').slice(0, 15);
  const out = [];
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    out.push(c & 0xff || (c >> 8) & 0xff);
  }
  return out;
}

/** The 32-bit Word 97 key of a password: high word from the matrix, low word the 16-bit verifier. */
export function wordLegacyKey(password) {
  const bytes = ansiBytes(password);
  if (!bytes.length) return 0;
  let high = WORD_INITIAL_CODE[bytes.length - 1];
  bytes.forEach((b, i) => {
    const row = WORD_MATRIX[15 - bytes.length + i];
    for (let bit = 0; bit < 7; bit++) if (b & (1 << bit)) high ^= row[bit];
  });
  const rotate = (v) => ((v >> 14) & 1) | ((v << 1) & 0x7fff);
  let low = 0;
  for (let i = bytes.length - 1; i >= 0; i--) low = rotate(low) ^ bytes[i];
  low = (rotate(low) ^ bytes.length) ^ 0xce4b;
  return ((high << 16) | low) >>> 0;
}

/** The key as the string Word hashes: its four bytes low byte first, as upper-case hex. */
export function wordLegacyHex(password) {
  const key = wordLegacyKey(password);
  const hex = (n) => (n & 0xff).toString(16).toUpperCase().padStart(2, '0');
  return hex(key) + hex(key >>> 8) + hex(key >>> 16) + hex(key >>> 24);
}

/** cryptAlgorithmSid → the hash it names. */
export const WORD_ALGORITHM_SID = { 1: 'MD2', 2: 'MD4', 3: 'MD5', 4: 'SHA-1', 12: 'SHA-256', 13: 'SHA-384', 14: 'SHA-512' };

/**
 * The attributes Word 2013 onwards writes on w:documentProtection for a
 * password: SHA-512 (sid 14) under the AES provider, 100,000 rounds and a
 * fresh 16-byte salt. Keys carry the `w:` prefix.
 */
export function wordProtectionAttrs(password, { salt = randomBytes(16).toString('base64'), spinCount = SPIN_COUNT } = {}) {
  return {
    'w:cryptProviderType': 'rsaAES',
    'w:cryptAlgorithmClass': 'hash',
    'w:cryptAlgorithmType': 'typeAny',
    'w:cryptAlgorithmSid': '14',
    'w:cryptSpinCount': String(spinCount),
    'w:hash': hashPassword(wordLegacyHex(password), { salt, spinCount, algorithm: 'SHA-512' }),
    'w:salt': salt,
  };
}

/** Does a w:documentProtection (attributes without their prefix) carry a password? */
export function wordHasPassword(a) {
  return Boolean(a && (a.hash || a.hashValue));
}

/**
 * Does `password` take off the protection these attributes describe? True
 * when there is no password. Word's form (cryptAlgorithmSid, hash, salt,
 * cryptSpinCount) and the ISO form (algorithmName, hashValue, saltValue,
 * spinCount) are both read; each is tried on Word's pre-hashed key first,
 * then on the password itself, which is what some other writers hash.
 */
export function checkWordPassword(password, a) {
  if (!wordHasPassword(a)) return true;
  const hash = a.hash || a.hashValue;
  const salt = a.salt || a.saltValue || '';
  const spinCount = Number(a.cryptSpinCount ?? a.spinCount ?? 0) || 0;
  const algorithm = a.algorithmName || WORD_ALGORITHM_SID[Number(a.cryptAlgorithmSid)] || 'SHA-1';
  for (const input of [wordLegacyHex(password), String(password ?? '')]) {
    try {
      if (hashPassword(input, { salt, spinCount, algorithm }) === hash) return true;
    } catch {
      return false;
    }
  }
  return false;
}
