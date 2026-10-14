// A .p12 / .pfx file (PKCS #12, RFC 7292): a certificate, its private key
// and the certificates that vouch for it, under a password.
//
// Windows, macOS and Outlook export an identity this way. Two generations of
// protection are read: PBES2 (PBKDF2 with AES, what Windows writes when AES
// is chosen and OpenSSL 3 writes by default) and the PKCS #12 password-based
// ciphers before it (Triple DES for the key, 40-bit RC2 for the
// certificates, what older Windows and OpenSSL wrote). The file's MAC is
// checked first, so a wrong password is said to be one.

import crypto from 'node:crypto';
import { parse, oidOf, smallInt, isContext, seq, set, int, oid, octets, explicit, tlv, algorithm, nul } from './der.js';
import { OID } from './oids.js';
import { octetsOf } from './cms.js';
import { certificate } from './cert.js';
import { rc2Decrypt, rc2Encrypt } from './rc2.js';

export class Pkcs12Error extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'Pkcs12Error';
    this.code = code;
  }
}

/**
 * The most work a file may ask for. Windows and OpenSSL write 2,000 to
 * 600,000 iterations; ten million is already minutes. A file asking for more,
 * or a key longer than any cipher takes, is refused rather than ground through.
 */
export const MAX_ITERATIONS = 10_000_000;
export const MAX_KEY_LENGTH = 64;

/** An iteration count, refused past the most any real file asks for. */
function iterationsOf(node) {
  const n = node ? smallInt(node, 4) : 1;
  if (n < 1 || n > MAX_ITERATIONS) throw new Pkcs12Error('algorithm', 'That file asks for more work to open than any real one does, so it is not opened.');
  return n;
}

const HMAC = { [OID.sha1]: 'sha1', [OID.sha256]: 'sha256', [OID.sha384]: 'sha384', [OID.sha512]: 'sha512' };
const PRF = { [OID.hmacSHA1]: 'sha1', [OID.hmacSHA256]: 'sha256', [OID.hmacSHA384]: 'sha384', [OID.hmacSHA512]: 'sha512' };

/** The password as PKCS #12's own derivation wants it: UTF-16 big-endian, two zero bytes after. */
function bmpPassword(password) {
  const s = String(password ?? '');
  const out = Buffer.alloc((s.length + 1) * 2);
  for (let i = 0; i < s.length; i++) out.writeUInt16BE(s.charCodeAt(i), i * 2);
  return out;
}

/** RFC 7292 appendix B.2: key (id 1), IV (2) or MAC key (3) material from a password and salt. */
export function pkcs12Kdf(password, salt, iterations, id, n, hash = 'sha1') {
  const u = crypto.createHash(hash).digest().length;
  const v = hash === 'sha384' || hash === 'sha512' ? 128 : 64;
  const P = bmpPassword(password);
  const fill = (src) => {
    if (!src.length) return Buffer.alloc(0);
    const out = Buffer.alloc(v * Math.ceil(src.length / v));
    for (let i = 0; i < out.length; i++) out[i] = src[i % src.length];
    return out;
  };
  const D = Buffer.alloc(v, id);
  const I = Buffer.concat([fill(Buffer.from(salt)), fill(P)]);
  const out = [];
  let have = 0;
  while (have < n) {
    let A = crypto.createHash(hash).update(D).update(I).digest();
    for (let r = 1; r < iterations; r++) A = crypto.createHash(hash).update(A).digest();
    out.push(A);
    have += A.length;
    if (have >= n) break;
    const B = Buffer.alloc(v);
    for (let i = 0; i < v; i++) B[i] = A[i % u];
    for (let j = 0; j < I.length; j += v) {
      let carry = 1;
      for (let k = v - 1; k >= 0; k--) {
        const sum = I[j + k] + B[k] + carry;
        I[j + k] = sum & 0xff;
        carry = sum >> 8;
      }
    }
  }
  return Buffer.concat(out).subarray(0, n);
}

/** Decrypt by a password-based algorithm identifier. */
function decryptWith(algNode, data, password) {
  const alg = oidOf(algNode.children[0]);
  const params = algNode.children[1];
  if (alg === OID.pbes2) {
    const [kdf, scheme] = params.children;
    if (oidOf(kdf.children[0]) !== OID.pbkdf2) throw new Pkcs12Error('algorithm', 'The file is protected in a way this cannot open.');
    const p = kdf.children[1].children;
    const salt = octetsOf(p[0]);
    const iterations = iterationsOf(p[1]);
    let k = 2;
    let keyLength = null;
    if (p[k] && p[k].tag === 0x02) keyLength = smallInt(p[k++], 2);
    if (keyLength != null && (keyLength < 1 || keyLength > MAX_KEY_LENGTH)) throw new Pkcs12Error('algorithm', 'That file asks for a key no cipher takes, so it is not opened.');
    const prf = p[k] ? PRF[oidOf(p[k].children[0])] || 'sha1' : 'sha1';
    const cipherOid = oidOf(scheme.children[0]);
    const cipher = { [OID.aes128CBC]: ['aes-128-cbc', 16], [OID.aes192CBC]: ['aes-192-cbc', 24], [OID.aes256CBC]: ['aes-256-cbc', 32], [OID.desEde3CBC]: ['des-ede3-cbc', 24] }[cipherOid];
    if (!cipher) throw new Pkcs12Error('algorithm', 'The file is protected with a cipher this cannot open.');
    const iv = octetsOf(scheme.children[1]);
    const key = crypto.pbkdf2Sync(Buffer.from(String(password ?? ''), 'utf8'), salt, iterations, keyLength || cipher[1], prf);
    const d = crypto.createDecipheriv(cipher[0], key, iv);
    return Buffer.concat([d.update(data), d.final()]);
  }
  const salt = octetsOf(params.children[0]);
  const iterations = iterationsOf(params.children[1]);
  if (alg === OID.pbeSHA1And3DES || alg === OID.pbeSHA1And2DES) {
    let key = pkcs12Kdf(password, salt, iterations, 1, 24);
    if (alg === OID.pbeSHA1And2DES) key = Buffer.concat([key.subarray(0, 16), key.subarray(0, 8)]);
    const iv = pkcs12Kdf(password, salt, iterations, 2, 8);
    const d = crypto.createDecipheriv('des-ede3-cbc', key, iv);
    return Buffer.concat([d.update(data), d.final()]);
  }
  if (alg === OID.pbeSHA1And40RC2 || alg === OID.pbeSHA1And128RC2) {
    const n = alg === OID.pbeSHA1And40RC2 ? 5 : 16;
    const key = pkcs12Kdf(password, salt, iterations, 1, n);
    const iv = pkcs12Kdf(password, salt, iterations, 2, 8);
    return rc2Decrypt(data, key, iv, n * 8);
  }
  throw new Pkcs12Error('algorithm', `The file is protected in a way this cannot open (${alg}).`);
}

/** The bags of a SafeContents. */
function bagsOf(safeContents, password, out) {
  for (const bag of parse(safeContents).children) {
    const type = oidOf(bag.children[0]);
    const value = bag.children[1].children[0];
    const attrs = new Map();
    for (const a of bag.children[2]?.children || []) attrs.set(oidOf(a.children[0]), a.children[1].children[0]);
    const localKeyId = attrs.get(OID.localKeyId) ? octetsOf(attrs.get(OID.localKeyId)).toString('hex') : null;
    if (type === OID.certBag) {
      if (oidOf(value.children[0]) !== OID.x509Certificate) continue;
      out.certs.push({ der: octetsOf(value.children[1].children[0]), localKeyId });
    } else if (type === OID.keyBag) {
      out.keys.push({ der: Buffer.from(value.raw), localKeyId });
    } else if (type === OID.pkcs8ShroudedKeyBag) {
      const der = decryptWith(value.children[0], octetsOf(value.children[1]), password);
      out.keys.push({ der, localKeyId });
    }
  }
}

/**
 * Read a .p12/.pfx.
 * @returns {{ key: crypto.KeyObject, cert: object, chain: object[] }} the key, the certificate it belongs to, and the others
 */
export function readPkcs12(bytes, password) {
  let pfx;
  try {
    pfx = parse(Buffer.from(bytes));
    if (pfx.children.length < 2 || smallInt(pfx.children[0], 1) !== 3) throw new Error('version');
  } catch {
    throw new Pkcs12Error('not-pkcs12', 'That file is not a certificate with its key (.p12 or .pfx).');
  }
  const authSafe = pfx.children[1];
  if (oidOf(authSafe.children[0]) !== OID.data) throw new Pkcs12Error('algorithm', 'That file is signed rather than password protected, which this cannot open.');
  const content = octetsOf(authSafe.children[1].children[0]);
  const macData = pfx.children[2];
  if (macData) {
    const [digestInfo, saltNode, iterNode] = macData.children;
    const hash = HMAC[oidOf(digestInfo.children[0].children[0])] || 'sha1';
    const expected = octetsOf(digestInfo.children[1]);
    const salt = octetsOf(saltNode);
    const iterations = iterationsOf(iterNode);
    const macKey = pkcs12Kdf(password, salt, iterations, 3, crypto.createHash(hash).digest().length, hash);
    const mac = crypto.createHmac(hash, macKey).update(content).digest();
    if (!crypto.timingSafeEqual(mac, expected.length === mac.length ? expected : Buffer.alloc(mac.length))) {
      throw new Pkcs12Error('password', 'That password is not right for this certificate.');
    }
  }
  const found = { certs: [], keys: [] };
  try {
    for (const ci of parse(content).children) {
      const type = oidOf(ci.children[0]);
      const inner = ci.children[1].children[0];
      if (type === OID.data) bagsOf(octetsOf(inner), password, found);
      else if (type === OID.encryptedData) {
        const eci = inner.children[1];
        bagsOf(decryptWith(eci.children[1], octetsOf(eci.children[2]), password), password, found);
      }
    }
  } catch (err) {
    if (err instanceof Pkcs12Error) throw err;
    throw new Pkcs12Error(macData ? 'damaged' : 'password', macData ? 'That file is damaged and cannot be read.' : 'That password is not right for this certificate.');
  }
  if (!found.keys.length) throw new Pkcs12Error('no-key', 'That file holds certificates but no private key: export it again with the private key included.');
  const key = crypto.createPrivateKey({ key: found.keys[0].der, format: 'der', type: 'pkcs8' });
  const certs = found.certs.map((c) => ({ ...c, cert: certificate(c.der) }));
  // The key's own certificate: the one its local key id names, or the one whose public key is the key's.
  const publicDer = crypto.createPublicKey(key).export({ type: 'spki', format: 'der' });
  const own =
    certs.find((c) => found.keys[0].localKeyId && c.localKeyId === found.keys[0].localKeyId) ||
    certs.find((c) => c.cert.x509.publicKey.export({ type: 'spki', format: 'der' }).equals(publicDer));
  if (!own) throw new Pkcs12Error('no-certificate', 'That file holds a private key without its certificate.');
  return { key, cert: own.cert, chain: certs.filter((c) => c !== own).map((c) => c.cert) };
}

// ── Writing, for the tests and nothing else in the suite ─────────────────

const bag = (type, value, localKeyId, name) =>
  seq(oid(type), explicit(0, value), set(seq(oid(OID.localKeyId), set(octets(localKeyId))), name ? seq(oid(OID.friendlyName), set(tlv(0x1e, bmpPassword(name).subarray(0, -2)))) : null));

/**
 * A .p12 of a key and its certificates. `legacy` writes them as older
 * Windows and OpenSSL did (Triple DES for the key, 40-bit RC2 for the
 * certificates, a SHA-1 MAC); otherwise as now (PBES2 with AES-256 and
 * PBKDF2-SHA256, a SHA-256 MAC).
 */
export function writePkcs12({ key, cert, chain = [], password, legacy = false, name = null, iterations = 2048 }) {
  const keyDer = (typeof key === 'string' ? crypto.createPrivateKey(key) : key).export({ type: 'pkcs8', format: 'der' });
  const localKeyId = crypto.createHash('sha1').update(certificate(cert).der).digest();
  const encrypt = (data, kind) => {
    const salt = crypto.randomBytes(8);
    if (legacy) {
      if (kind === 'key') {
        const k = pkcs12Kdf(password, salt, iterations, 1, 24);
        const iv = pkcs12Kdf(password, salt, iterations, 2, 8);
        const c = crypto.createCipheriv('des-ede3-cbc', k, iv);
        return { alg: algorithm(OID.pbeSHA1And3DES, seq(octets(salt), int(iterations))), data: Buffer.concat([c.update(data), c.final()]) };
      }
      const k = pkcs12Kdf(password, salt, iterations, 1, 5);
      const iv = pkcs12Kdf(password, salt, iterations, 2, 8);
      return { alg: algorithm(OID.pbeSHA1And40RC2, seq(octets(salt), int(iterations))), data: rc2Encrypt(data, k, iv, 40) };
    }
    const iv = crypto.randomBytes(16);
    const k = crypto.pbkdf2Sync(Buffer.from(String(password), 'utf8'), salt, iterations, 32, 'sha256');
    const c = crypto.createCipheriv('aes-256-cbc', k, iv);
    const params = seq(algorithm(OID.pbkdf2, seq(octets(salt), int(iterations), algorithm(OID.hmacSHA256))), algorithm(OID.aes256CBC, octets(iv)));
    return { alg: algorithm(OID.pbes2, params), data: Buffer.concat([c.update(data), c.final()]) };
  };
  const shrouded = encrypt(keyDer, 'key');
  const keyBags = seq(bag(OID.pkcs8ShroudedKeyBag, seq(shrouded.alg, octets(shrouded.data)), localKeyId, name));
  const certBags = seq(
    ...[cert, ...chain].map((c, i) => {
      const der = certificate(c).der;
      const value = seq(oid(OID.x509Certificate), explicit(0, octets(der)));
      return i === 0 ? bag(OID.certBag, value, localKeyId, name) : seq(oid(OID.certBag), explicit(0, value));
    })
  );
  const sealed = encrypt(certBags, 'certs');
  const authSafe = seq(
    seq(oid(OID.encryptedData), explicit(0, seq(int(0), seq(oid(OID.data), sealed.alg, tlv(0x80, sealed.data))))),
    seq(oid(OID.data), explicit(0, octets(keyBags)))
  );
  const hash = legacy ? 'sha1' : 'sha256';
  const macSalt = crypto.randomBytes(8);
  const macKey = pkcs12Kdf(password, macSalt, iterations, 3, crypto.createHash(hash).digest().length, hash);
  const mac = crypto.createHmac(hash, macKey).update(authSafe).digest();
  return seq(int(3), seq(oid(OID.data), explicit(0, octets(authSafe))), seq(seq(algorithm(legacy ? OID.sha1 : OID.sha256), octets(mac)), octets(macSalt), int(iterations)));
}

export { nul, isContext };
