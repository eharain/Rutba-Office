// Cryptographic Message Syntax (RFC 5652): the signatures and the encryption
// inside S/MIME mail.
//
// Read: a SignedData is checked - the digest of the content against the one
// the signer signed, and the signature against the signer's certificate -
// and an EnvelopedData (or AuthEnvelopedData, AES-GCM, RFC 5083) is opened
// with a certificate and key the person holds. Written: a detached
// SignedData over a MIME entity, and an EnvelopedData for the recipients'
// certificates. Everything is Node's crypto; the structures are read and
// written by der.js.

import crypto from 'node:crypto';
import { parse, oidOf, smallInt, timeOf, isContext, seq, set, int, oid, octets, explicit, tlv, time, algorithm, nul } from './der.js';
import { OID, DIGEST, DIGEST_OID, SIGNATURE_DIGEST, CIPHER } from './oids.js';
import { certificate } from './cert.js';
import { rc2Decrypt } from './rc2.js';

export class CmsError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'CmsError';
    this.code = code;
  }
}

/** An OCTET STRING's bytes, put together when BER split it into pieces. */
export function octetsOf(node) {
  if (!node.constructed) return Buffer.from(node.value);
  return Buffer.concat(node.children.map(octetsOf));
}

/** A ContentInfo: its type and its content node. */
export function contentInfo(bytes) {
  const root = parse(bytes);
  if (!root.children?.length) throw new CmsError('not-cms', 'This is not a signed or encrypted message part.');
  const type = oidOf(root.children[0]);
  const content = root.children[1] && isContext(root.children[1], 0) ? root.children[1].children[0] : null;
  return { type, content };
}

const attributeMap = (setNode) => {
  const out = new Map();
  for (const attr of setNode?.children || []) out.set(oidOf(attr.children[0]), attr.children[1].children);
  return out;
};

/** Who signed or is a recipient: an issuer and serial number, or a subject key identifier. */
function identifierOf(node) {
  if (isContext(node, 0)) return { subjectKeyId: Buffer.from(node.value).toString('hex') };
  return { issuerRaw: Buffer.from(node.children[0].raw), serialRaw: Buffer.from(node.children[1].value) };
}

const sameSerial = (a, b) => {
  const strip = (x) => {
    let i = 0;
    while (i < x.length - 1 && x[i] === 0) i++;
    return x.subarray(i);
  };
  return strip(a).equals(strip(b));
};

/** Is `cert` the one `id` names. */
export function names(id, cert) {
  if (id.subjectKeyId) return cert.subjectKeyId === id.subjectKeyId;
  return cert.issuerRaw.equals(id.issuerRaw) && sameSerial(cert.serialRaw, id.serialRaw);
}

// ── SignedData ────────────────────────────────────────────────────────────

/**
 * Read a SignedData.
 * @returns {{ content: Buffer|null, contentType, certificates: object[], signers: object[] }}
 */
export function readSignedData(node) {
  const c = node.children;
  let i = 2;
  const encap = c[i++];
  const contentType = oidOf(encap.children[0]);
  const content = encap.children[1] ? octetsOf(encap.children[1].children[0]) : null;
  const certificates = [];
  if (c[i] && isContext(c[i], 0)) {
    for (const certNode of c[i].children) {
      // Only X.509 certificates (a SEQUENCE); attribute certificates and the like are left.
      if (certNode.tag !== 0x30) continue;
      try {
        certificates.push(certificate(Buffer.from(certNode.raw)));
      } catch {
        /* a certificate Node cannot read is not one to trust */
      }
    }
    i++;
  }
  if (c[i] && isContext(c[i], 1)) i++;
  const signers = (c[i]?.children || []).map((si) => {
    const s = si.children;
    let k = 1;
    const id = identifierOf(s[k++]);
    const digest = oidOf(s[k++].children[0]);
    let signedAttrs = null;
    if (isContext(s[k], 0)) signedAttrs = s[k++];
    const sigAlg = s[k++];
    const signature = octetsOf(s[k++]);
    return { id, digest, signedAttrs, sigAlg: { oid: oidOf(sigAlg.children[0]), params: sigAlg.children[1] || null }, signature };
  });
  return { content, contentType, certificates, signers };
}

/**
 * Verify a signature with a certificate's key, by the algorithm the signer
 * named. A signature algorithm that names a hash other than the digest the
 * signer gave is refused (a signature over one hash claiming another).
 */
function verifySignature({ oid: alg, params }, digestName, data, signature, cert) {
  const key = cert.x509.publicKey;
  const mismatch = () => new CmsError('mismatch', 'The signature names a different hash from the one the message was summed with.');
  if (alg === OID.rsassaPss) {
    // RSASSA-PSS-params: [0] hash, [1] mask generation, [2] salt length (default SHA-1 and 20).
    let hash = 'sha1';
    let saltLength = 20;
    for (const p of params?.children || []) {
      if (isContext(p, 0)) hash = DIGEST[oidOf(p.children[0].children[0])] || hash;
      if (isContext(p, 2)) saltLength = smallInt(p.children[0], 2);
    }
    if (hash !== digestName) throw mismatch();
    return crypto.verify(hash, data, { key, padding: crypto.constants.RSA_PKCS1_PSS_PADDING, saltLength }, signature);
  }
  if (SIGNATURE_DIGEST[alg] && SIGNATURE_DIGEST[alg] !== digestName) throw mismatch();
  if (alg === OID.rsaEncryption || SIGNATURE_DIGEST[alg] && alg.startsWith('1.2.840.113549')) {
    return crypto.verify(digestName, data, { key, padding: crypto.constants.RSA_PKCS1_PADDING }, signature);
  }
  if (alg === OID.ecPublicKey || alg.startsWith('1.2.840.10045.4')) {
    return crypto.verify(digestName, data, { key, dsaEncoding: 'der' }, signature);
  }
  throw new CmsError('algorithm', `The message is signed with an algorithm this cannot check (${alg}).`);
}

/**
 * Check each signer of a SignedData over `content` (the detached content, or
 * the SignedData's own).
 * @param {object} signed from readSignedData
 * @param {Buffer|null} detached
 * @param {object[]} more certificates found elsewhere (a contact's), to find a signer not included
 * @returns {Array<{ cert, intact: boolean, valid: boolean, signedAt: Date|null, digest: string, error?: string }>}
 */
export function verifySigners(signed, detached = null, more = []) {
  const content = detached ?? signed.content;
  if (!content) throw new CmsError('no-content', 'The signed content is missing.');
  return signed.signers.map((s) => {
    const digestName = DIGEST[s.digest];
    const cert = [...signed.certificates, ...more].find((c) => names(s.id, c)) || null;
    if (!digestName) return { cert, intact: false, valid: false, signedAt: null, digest: s.digest, error: 'algorithm' };
    const digest = crypto.createHash(digestName).update(content).digest();
    let data = content;
    let intact = true;
    let signedAt = null;
    if (s.signedAttrs) {
      const attrs = attributeMap(s.signedAttrs);
      const md = attrs.get(OID.messageDigest)?.[0];
      intact = Boolean(md) && octetsOf(md).equals(digest);
      const ct = attrs.get(OID.contentType)?.[0];
      if (ct && oidOf(ct) !== signed.contentType) intact = false;
      const when = attrs.get(OID.signingTime)?.[0];
      signedAt = when ? timeOf(when) : null;
      // The signature is over the attributes as a SET (tag 0x31), not as the [0] they are sent under.
      data = Buffer.from(s.signedAttrs.raw);
      data[0] = 0x31;
    }
    if (!cert) return { cert: null, intact, valid: false, signedAt, digest: digestName, error: 'no-certificate' };
    let valid = false;
    let error;
    try {
      valid = verifySignature(s.sigAlg, digestName, data, s.signature, cert);
    } catch (err) {
      error = err.code || 'algorithm';
    }
    return { cert, intact, valid, signedAt, digest: digestName, ...(error ? { error } : {}) };
  });
}

/**
 * A detached SignedData over `content` (or, with `encapsulate`, one that carries it), signed with `key` and `cert`, the
 * chain beside it (S/MIME 4.0, RFC 8551: SHA-256, the signing time and the
 * ciphers this can read listed in the signed attributes).
 */
export function signDetached(content, { key, cert, chain = [], digest = 'sha256', now = new Date(), encapsulate = false }) {
  const signer = certificate(cert);
  const digestOid = DIGEST_OID[digest];
  const md = crypto.createHash(digest).update(content).digest();
  // Only AES is offered: Triple DES is too weak to ask anyone to encrypt with.
  const caps = seq(seq(oid(OID.aes256CBC)), seq(oid(OID.aes128CBC)));
  const attrs = [
    seq(oid(OID.contentType), set(oid(OID.data))),
    seq(oid(OID.signingTime), set(time(now))),
    seq(oid(OID.messageDigest), set(octets(md))),
    seq(oid(OID.smimeCapabilities), set(caps)),
  ];
  const signedSet = set(...attrs);
  const privateKey = typeof key === 'string' || Buffer.isBuffer(key) ? crypto.createPrivateKey(key) : key;
  const ec = privateKey.asymmetricKeyType === 'ec';
  const signature = ec ? crypto.sign(digest, signedSet, { key: privateKey, dsaEncoding: 'der' }) : crypto.sign(digest, signedSet, { key: privateKey, padding: crypto.constants.RSA_PKCS1_PADDING });
  const sigAlg = ec ? algorithm({ sha256: OID.ecdsaWithSHA256, sha384: OID.ecdsaWithSHA384, sha512: OID.ecdsaWithSHA512 }[digest], false) : algorithm(OID.rsaEncryption);
  const signerInfo = seq(
    int(1),
    seq(signer.issuerRaw, tlv(0x02, signer.serialRaw)),
    algorithm(digestOid),
    tlv(0xa0, parse(signedSet).value),
    sigAlg,
    octets(signature)
  );
  const certs = [signer.der, ...chain.map((c) => certificate(c).der)];
  const encap = encapsulate ? seq(oid(OID.data), explicit(0, octets(content))) : seq(oid(OID.data));
  const signedData = seq(int(1), set(algorithm(digestOid)), encap, tlv(0xa0, certs.map((d) => Buffer.from(d))), set(signerInfo));
  return seq(oid(OID.signedData), explicit(0, signedData));
}

// ── EnvelopedData ─────────────────────────────────────────────────────────

/** Read an EnvelopedData or AuthEnvelopedData. */
export function readEnveloped(node, auth = false) {
  const c = node.children;
  let i = 1;
  if (isContext(c[i], 0)) i++;
  const recipients = c[i++].children.map((ri) => {
    if (ri.tag !== 0x30) return { kind: ri.cls === 2 ? ['', 'agree', 'kek', 'password', 'other'][ri.number] || 'other' : 'other' };
    const r = ri.children;
    return { kind: 'transport', id: identifierOf(r[1]), alg: { oid: oidOf(r[2].children[0]), params: r[2].children[1] || null }, encryptedKey: octetsOf(r[3]) };
  });
  const eci = c[i++];
  const cipherAlg = { oid: oidOf(eci.children[1].children[0]), params: eci.children[1].children[1] || null };
  const encrypted = eci.children[2] ? octetsOf(eci.children[2]) : null;
  let authAttrs = null;
  let mac = null;
  if (auth) {
    if (c[i] && isContext(c[i], 1)) authAttrs = c[i++];
    mac = c[i] ? octetsOf(c[i]) : null;
  }
  return { recipients, cipherAlg, encrypted, authAttrs, mac, auth };
}

/** The content key, opened with the private key: RSA PKCS #1 v1.5 or RSAES-OAEP. */
function unwrapKey(recipient, privateKey) {
  if (recipient.alg.oid === OID.rsaesOaep) {
    let hash = 'sha1';
    for (const p of recipient.alg.params?.children || []) if (isContext(p, 0)) hash = DIGEST[oidOf(p.children[0].children[0])] || hash;
    return crypto.privateDecrypt({ key: privateKey, padding: crypto.constants.RSA_PKCS1_OAEP_PADDING, oaepHash: hash }, recipient.encryptedKey);
  }
  if (recipient.alg.oid === OID.rsaEncryption) return crypto.privateDecrypt({ key: privateKey, padding: crypto.constants.RSA_PKCS1_PADDING }, recipient.encryptedKey);
  throw new CmsError('algorithm', `The message's key is wrapped in a way this cannot open (${recipient.alg.oid}).`);
}

/**
 * Open an enveloped message with the first identity it was encrypted for.
 * @param {object} env from readEnveloped
 * @param {Array<{ cert: object, key: crypto.KeyObject }>} identities
 * @returns {{ content: Buffer, identity, cipher: string }}
 */
export function decryptEnveloped(env, identities) {
  let identity = null;
  let recipient = null;
  for (const r of env.recipients) {
    if (r.kind !== 'transport') continue;
    identity = identities.find((x) => names(r.id, x.cert));
    if (identity) {
      recipient = r;
      break;
    }
  }
  if (!recipient) {
    const others = env.recipients.filter((r) => r.kind !== 'transport');
    throw new CmsError(others.length && !env.recipients.some((r) => r.kind === 'transport') ? 'recipient-kind' : 'no-identity', 'This message was not encrypted for a certificate on this computer.');
  }
  let cek;
  try {
    cek = unwrapKey(recipient, identity.key);
  } catch (err) {
    if (err instanceof CmsError) throw err;
    throw new CmsError('key', 'The message could not be opened with your certificate\'s key.');
  }
  const alg = env.cipherAlg;
  let content;
  if (alg.oid === OID.rc2CBC) {
    // RC2-CBC parameters: a version (the effective key bits, encoded) and the IV.
    const p = alg.params;
    const iv = octetsOf(p.children[p.children.length - 1]);
    const version = p.children.length > 1 ? smallInt(p.children[0], 3) : 32;
    const bits = { 160: 40, 120: 64, 58: 128 }[version] || version;
    content = rc2Decrypt(env.encrypted, cek, iv, bits);
    return { content, identity, cipher: `RC2 (${bits}-bit)` };
  }
  const cipher = CIPHER[alg.oid];
  if (!cipher) throw new CmsError('algorithm', `The message is encrypted with a cipher this cannot open (${alg.oid}).`);
  try {
    if (cipher.gcm) {
      const nonce = octetsOf(alg.params.children[0]);
      const tagLength = alg.params.children[1] ? smallInt(alg.params.children[1], 1) : 12;
      if (tagLength < 12 || tagLength > 16) throw new CmsError('changed', 'The encrypted message names an authentication tag no sender uses.');
      const d = crypto.createDecipheriv(cipher.name, cek, nonce, { authTagLength: tagLength });
      if (env.authAttrs) {
        const aad = Buffer.from(env.authAttrs.raw);
        aad[0] = 0x31;
        d.setAAD(aad);
      }
      d.setAuthTag(env.mac);
      content = Buffer.concat([d.update(env.encrypted), d.final()]);
    } else {
      const iv = octetsOf(alg.params);
      const d = crypto.createDecipheriv(cipher.name, cek, iv);
      content = Buffer.concat([d.update(env.encrypted), d.final()]);
    }
  } catch {
    throw new CmsError('changed', 'The encrypted message has been changed or damaged, so it cannot be opened.');
  }
  return { content, identity, cipher: cipher.label };
}

/**
 * An EnvelopedData of `content` for each certificate given (RSA keys), AES-256-CBC,
 * the content key wrapped with RSA PKCS #1 v1.5, which every mail program
 * that reads S/MIME opens.
 */
export function encryptEnveloped(content, recipients) {
  const cek = crypto.randomBytes(32);
  const iv = crypto.randomBytes(16);
  const c = crypto.createCipheriv('aes-256-cbc', cek, iv);
  const encrypted = Buffer.concat([c.update(content), c.final()]);
  const infos = recipients.map((r) => {
    const cert = certificate(r);
    if (cert.x509.publicKey.asymmetricKeyType !== 'rsa') throw new CmsError('recipient-kind', `${cert.email || cert.name}'s certificate is not an RSA one, which this cannot encrypt for yet.`);
    const wrapped = crypto.publicEncrypt({ key: cert.x509.publicKey, padding: crypto.constants.RSA_PKCS1_PADDING }, cek);
    return seq(int(0), seq(cert.issuerRaw, tlv(0x02, cert.serialRaw)), algorithm(OID.rsaEncryption), octets(wrapped));
  });
  const eci = seq(oid(OID.data), algorithm(OID.aes256CBC, octets(iv)), tlv(0x80, encrypted));
  return seq(oid(OID.envelopedData), explicit(0, seq(int(0), set(...infos), eci)));
}

export { nul };
