// RC2 (RFC 2268), in CBC mode.
//
// An old cipher, and the TLS libraries Node and Electron are built on no
// longer offer it, but a .pfx exported by Windows or older OpenSSL keeps its
// certificates encrypted with 40-bit RC2, and some old encrypted mail used
// it: so it is here to read them (and, for the tests, to write such a file).
// tests/mail-smime.test.js checks it against the RFC's test vectors.

// RFC 2268 section 2, PITABLE: a permutation of the bytes, from the digits of pi.
const PITABLE = Buffer.from(
  [
    'd978f9c419ddb5ed28e9fd794aa0d89d', 'c67e37832b76538e624c6488448bfba2', '179a59f587b34f1361456d8d09817d32', 'bd8f40eb86b77b0bf09521225c6b4e82',
    '54d66593ce60b21c7356c014a78cf1dc', '1275ca1f3bbee4d1423dd430a33cb626', '6fbf0eda4669075727f21d9bbc944303', 'f811c7f690ef3ee706c3d52fc8661ed7',
    '08e8eade8052eef784aa72ac354d6a2a', '961ad2715a1549744b9fd05e0418a4ec', 'c2e0416e0f51cbcc2491af50a1f47039', '997c3a8523b8b47afc02365b25559731',
    '2d5dfa98e38a92ae05df2910676cbac9', 'd300e6cfe19ea82c6316013f58e289a9', '0d38341bab33ffb0bb480c5fb9b1cd2e', 'c5f3db47e5a59c770aa62068fe7fc1ad',
  ].join(''),
  'hex'
);

/** The 64 16-bit subkeys for `key` at `bits` effective key bits. */
function expand(key, bits) {
  const T = key.length;
  if (T < 1 || T > 128) throw new Error('RC2: a key of 1 to 128 bytes');
  const L = Buffer.alloc(128);
  Buffer.from(key).copy(L);
  for (let i = T; i < 128; i++) L[i] = PITABLE[(L[i - 1] + L[i - T]) & 0xff];
  const T8 = (bits + 7) >> 3;
  const TM = 0xff >> (8 * T8 - bits);
  L[128 - T8] = PITABLE[L[128 - T8] & TM];
  for (let i = 127 - T8; i >= 0; i--) L[i] = PITABLE[L[i + 1] ^ L[i + T8]];
  const K = new Uint16Array(64);
  for (let i = 0; i < 64; i++) K[i] = L[2 * i] | (L[2 * i + 1] << 8);
  return K;
}

const S = [1, 2, 3, 5];

function encryptBlock(K, b, at) {
  const R = [b[at] | (b[at + 1] << 8), b[at + 2] | (b[at + 3] << 8), b[at + 4] | (b[at + 5] << 8), b[at + 6] | (b[at + 7] << 8)];
  let j = 0;
  const mix = () => {
    for (let i = 0; i < 4; i++) {
      const r = (R[i] + K[j++] + (R[(i + 3) & 3] & R[(i + 2) & 3]) + (~R[(i + 3) & 3] & R[(i + 1) & 3])) & 0xffff;
      R[i] = ((r << S[i]) | (r >>> (16 - S[i]))) & 0xffff;
    }
  };
  const mash = () => {
    for (let i = 0; i < 4; i++) R[i] = (R[i] + K[R[(i + 3) & 3] & 63]) & 0xffff;
  };
  for (let r = 0; r < 5; r++) mix();
  mash();
  for (let r = 0; r < 6; r++) mix();
  mash();
  for (let r = 0; r < 5; r++) mix();
  for (let i = 0; i < 4; i++) {
    b[at + 2 * i] = R[i] & 0xff;
    b[at + 2 * i + 1] = R[i] >> 8;
  }
}

function decryptBlock(K, b, at) {
  const R = [b[at] | (b[at + 1] << 8), b[at + 2] | (b[at + 3] << 8), b[at + 4] | (b[at + 5] << 8), b[at + 6] | (b[at + 7] << 8)];
  let j = 63;
  const rmix = () => {
    for (let i = 3; i >= 0; i--) {
      const r = ((R[i] >>> S[i]) | (R[i] << (16 - S[i]))) & 0xffff;
      R[i] = (r - K[j--] - (R[(i + 3) & 3] & R[(i + 2) & 3]) - (~R[(i + 3) & 3] & R[(i + 1) & 3])) & 0xffff;
    }
  };
  const rmash = () => {
    for (let i = 3; i >= 0; i--) R[i] = (R[i] - K[R[(i + 3) & 3] & 63]) & 0xffff;
  };
  for (let r = 0; r < 5; r++) rmix();
  rmash();
  for (let r = 0; r < 6; r++) rmix();
  rmash();
  for (let r = 0; r < 5; r++) rmix();
  for (let i = 0; i < 4; i++) {
    b[at + 2 * i] = R[i] & 0xff;
    b[at + 2 * i + 1] = R[i] >> 8;
  }
}

/** The table, for the test that it is a permutation. */
export const RC2_PITABLE = PITABLE;

/** One block, for the RFC's test vectors. */
export function rc2Block(key, bits, block, decrypt = false) {
  const K = expand(Buffer.from(key), bits);
  const b = Buffer.from(block);
  (decrypt ? decryptBlock : encryptBlock)(K, b, 0);
  return b;
}

/** Decrypt RC2-CBC with PKCS #7 padding. */
export function rc2Decrypt(data, key, iv, bits = key.length * 8) {
  if (data.length % 8) throw new Error('RC2: the data is not a whole number of blocks');
  const K = expand(Buffer.from(key), bits);
  const out = Buffer.from(data);
  let prev = Buffer.from(iv);
  for (let at = 0; at < out.length; at += 8) {
    const c = Buffer.from(out.subarray(at, at + 8));
    decryptBlock(K, out, at);
    for (let i = 0; i < 8; i++) out[at + i] ^= prev[i];
    prev = c;
  }
  const pad = out[out.length - 1];
  const bad = () => new Error('RC2: bad padding (a wrong password, or a damaged file)');
  if (!pad || pad > 8 || out.length < pad) throw bad();
  for (let i = out.length - pad; i < out.length; i++) if (out[i] !== pad) throw bad();
  return out.subarray(0, out.length - pad);
}

/** Encrypt RC2-CBC with PKCS #7 padding. */
export function rc2Encrypt(data, key, iv, bits = key.length * 8) {
  const K = expand(Buffer.from(key), bits);
  const pad = 8 - (data.length % 8);
  const out = Buffer.concat([Buffer.from(data), Buffer.alloc(pad, pad)]);
  let prev = Buffer.from(iv);
  for (let at = 0; at < out.length; at += 8) {
    for (let i = 0; i < 8; i++) out[at + i] ^= prev[i];
    encryptBlock(K, out, at);
    prev = out.subarray(at, at + 8);
  }
  return out;
}
