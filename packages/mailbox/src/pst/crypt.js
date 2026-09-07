// The two ciphers a .pst / .ost can use — [MS-PST] §5.1.
//
// Neither is encryption in the security sense; both are obfuscation, and
// Microsoft documents them in full precisely so that other software can read
// the files. What matters here is that they are byte-exact: one wrong entry in
// a substitution table does not fail loudly, it silently returns a different
// character in the middle of somebody's mail.
//
// So the tables are validated on load. A substitution cipher's table must be a
// permutation of 0..255 — every value present exactly once — and if the table
// we hold is not one, we refuse to use it rather than hand back plausible
// nonsense. `available()` says which methods this build can actually decode,
// and the reader turns that into an honest message instead of a corrupt inbox.

import { mpbbR, mpbbS, mpbbI, verifyTables } from './tables.js';

export const CRYPT = { NONE: 0, PERMUTE: 1, CYCLIC: 2 };

export const CRYPT_NAME = {
  0: 'none',
  1: 'compressible (permute)',
  2: 'high (cyclic)',
};

/** @type {{ encode: Uint8Array, decode: Uint8Array } | null} */
let permute = null;
/** @type {[Uint8Array, Uint8Array, Uint8Array] | null} */
let cyclic = null;

/** A substitution table is only usable if it is a bijection on 0..255. */
export function isPermutation(table) {
  if (!table || table.length !== 256) return false;
  const seen = new Uint8Array(256);
  for (let i = 0; i < 256; i++) {
    const v = table[i];
    if (v < 0 || v > 255 || seen[v]) return false;
    seen[v] = 1;
  }
  return true;
}

function invert(table) {
  const out = new Uint8Array(256);
  for (let i = 0; i < 256; i++) out[table[i]] = i;
  return out;
}

/**
 * Install the substitution tables. Called by whichever module holds the data,
 * so this file stays code and the tables stay data.
 * @param {{ permute?: ArrayLike<number>, cyclic?: [ArrayLike<number>, ArrayLike<number>, ArrayLike<number>] }} tables
 */
export function installTables(tables = {}) {
  if (tables.permute) {
    const encode = Uint8Array.from(tables.permute);
    if (!isPermutation(encode)) throw new Error('permute table is not a permutation of 0..255');
    const decode = tables.permuteDecode ? Uint8Array.from(tables.permuteDecode) : invert(encode);
    permute = { encode, decode };
  }
  if (tables.cyclic) {
    const three = tables.cyclic.map((t) => Uint8Array.from(t));
    if (three.length !== 3 || !three.every(isPermutation)) {
      throw new Error('cyclic tables must be three permutations of 0..255');
    }
    cyclic = three;
  }
}

// The published tables are installed on load, having first been checked against
// themselves. A build that somehow shipped a damaged table decodes nothing
// rather than decoding wrongly.
const CHECK = verifyTables();
if (CHECK.ok) {
  installTables({ permute: mpbbR, permuteDecode: mpbbI, cyclic: [mpbbR, mpbbS, mpbbI] });
}

/** Why a method is unavailable, when it is — surfaced in the reader's error. */
export const tableStatus = () => CHECK;

/** Which crypt methods this build can decode. */
export function available() {
  return {
    [CRYPT.NONE]: true,
    [CRYPT.PERMUTE]: Boolean(permute),
    [CRYPT.CYCLIC]: Boolean(cyclic),
  };
}

export function supports(method) {
  return Boolean(available()[method]);
}

/** NDB_CRYPT_PERMUTE: one byte-for-byte substitution, applied in reverse. */
export function decodePermute(bytes) {
  if (!permute) throw new Error('permute table not installed');
  const table = permute.decode;
  const out = new Uint8Array(bytes.length);
  for (let i = 0; i < bytes.length; i++) out[i] = table[bytes[i]];
  return out;
}

/**
 * NDB_CRYPT_CYCLIC: three substitutions with a rolling key derived from the
 * block id. Symmetric — the same routine encodes and decodes.
 */
export function codeCyclic(bytes, key) {
  if (!cyclic) throw new Error('cyclic tables not installed');
  const [t1, t2, t3] = cyclic;
  const out = Uint8Array.from(bytes);
  let w = ((key & 0xffff) ^ ((key >>> 16) & 0xffff)) & 0xffff;
  for (let i = 0; i < out.length; i++) {
    let b = out[i];
    b = (b + (w & 0xff)) & 0xff;
    b = t1[b];
    b = (b + ((w >>> 8) & 0xff)) & 0xff;
    b = t2[b];
    b = (b - ((w >>> 8) & 0xff)) & 0xff;
    b = t3[b];
    out[i] = (b - (w & 0xff)) & 0xff;
    w = (w + 1) & 0xffff;
  }
  return out;
}

/**
 * Decode a block.
 * @throws when the file uses a method this build has no table for — the caller
 *   turns that into a message naming the method, so the user knows what to do.
 */
export function decodeBlock(bytes, method, bid) {
  if (method === CRYPT.NONE) return bytes;
  if (method === CRYPT.PERMUTE) return decodePermute(bytes);
  if (method === CRYPT.CYCLIC) return codeCyclic(bytes, bid);
  throw new Error(`unknown Outlook crypt method ${method}`);
}
