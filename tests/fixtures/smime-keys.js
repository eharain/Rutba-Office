// Certificates and keys for the S/MIME tests and checks, made when they run.
//
// A small certificate authority, people's certificates it issues (for mail:
// signing and encryption, their address in the certificate), and the
// awkward cases a reader has to say something true about: a certificate
// past its date, one no authority on this computer vouches for, one for
// another address. Nothing here is a real key; each run makes its own.

import crypto from 'node:crypto';
import { seq, set, oid, int, utf8, ia5, printable, bool, bits, octets, explicit, time, algorithm, tlv } from '../../packages/mailbox/src/smime/der.js';
import { OID } from '../../packages/mailbox/src/smime/oids.js';
import { writePkcs12 } from '../../packages/mailbox/src/smime/pkcs12.js';
import { pem } from '../../packages/mailbox/src/smime/cert.js';

const DAY = 86_400_000;

function name({ cn, o, email }) {
  return seq(
    ...[o ? set(seq(oid('2.5.4.10'), utf8(o))) : null, cn ? set(seq(oid(OID.commonName), utf8(cn))) : null, email ? set(seq(oid(OID.emailAddress), ia5(email))) : null].filter(Boolean)
  );
}

const ext = (id, critical, value) => seq(oid(id), critical ? bool(true) : null, octets(value));

/**
 * A certificate.
 * The awkward ones are made with: `pathLen` on an authority, `caUsage` (an
 * authority's own extended key usage), `permitted` (mail name constraints),
 * `keyUsage` (the byte, to replace the usual one), `noBasicConstraints`,
 * and `unknownCritical` (an extension nobody knows, marked critical).
 * @param {{ subject: {cn, o?, email?}, issuer?: {cn, o?, email?}, publicKey, signingKey, serial?, from?, to?, ca?, email?: string, usage?: string[], pathLen?: number, caUsage?: string[], permitted?: string[], keyUsage?: number, noBasicConstraints?: boolean, unknownCritical?: boolean }} o
 * @returns {string} PEM
 */
export function makeCertificate({ subject, issuer = subject, publicKey, signingKey, serial = crypto.randomBytes(9), from = new Date(Date.now() - DAY), to = new Date(Date.now() + 365 * DAY), ca = false, email = null, usage = [OID.emailProtection], pathLen = null, caUsage = null, permitted = null, keyUsage = null, noBasicConstraints = false, unknownCritical = false }) {
  const spki = publicKey.export({ type: 'spki', format: 'der' });
  const ski = crypto.createHash('sha1').update(spki).digest();
  const ec = signingKey.asymmetricKeyType === 'ec';
  const sigAlg = ec ? algorithm(OID.ecdsaWithSHA256, false) : algorithm(OID.sha256WithRSA);
  const extensions = [
    ...(noBasicConstraints ? [] : [ext(OID.basicConstraints, ca, ca ? seq(bool(true), pathLen == null ? null : int(pathLen)) : seq())]),
    ext(OID.keyUsage, true, bits(Buffer.from([keyUsage ?? (ca ? 0x06 : 0xa0)]))),
    ext(OID.subjectKeyIdentifier, false, octets(ski)),
    ...(email ? [ext(OID.subjectAltName, false, seq(tlv(0x81, Buffer.from(email, 'latin1'))))] : []),
    ...(!ca && usage ? [ext(OID.extKeyUsage, false, seq(...usage.map((u) => oid(u))))] : []),
    ...(ca && caUsage ? [ext(OID.extKeyUsage, false, seq(...caUsage.map((u) => oid(u))))] : []),
    ...(permitted ? [ext('2.5.29.30', true, seq(explicit(0, ...permitted.map((p) => seq(tlv(0x81, Buffer.from(p, 'latin1')))))))] : []),
    ...(unknownCritical ? [ext('1.3.6.1.4.1.55555.1', true, seq())] : []),
  ];
  const tbs = seq(explicit(0, int(2)), int(serial), sigAlg, name(issuer), seq(time(from), time(to)), name(subject), Buffer.from(spki), explicit(3, seq(...extensions)));
  const signature = ec ? crypto.sign('sha256', tbs, { key: signingKey, dsaEncoding: 'der' }) : crypto.sign('sha256', tbs, signingKey);
  return pem(seq(tbs, sigAlg, bits(signature)));
}

const rsa = () => crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });

/**
 * An authority and the people the tests write as, each with a key, a
 * certificate and the PEM of both.
 */
export function makeIdentities({ org = 'Rutba Test Authority' } = {}) {
  const caKeys = rsa();
  const caCert = makeCertificate({ subject: { cn: `${org} Root`, o: org }, publicKey: caKeys.publicKey, signingKey: caKeys.privateKey, ca: true, to: new Date(Date.now() + 3650 * DAY) });
  const person = (cn, email, opts = {}) => {
    const keys = opts.ec ? crypto.generateKeyPairSync('ec', { namedCurve: 'P-256' }) : rsa();
    const cert = makeCertificate({ subject: { cn, email }, issuer: { cn: `${org} Root`, o: org }, publicKey: keys.publicKey, signingKey: opts.issuerKey || caKeys.privateKey, email, ...opts });
    return { key: keys.privateKey, publicKey: keys.publicKey, cert, keyPem: keys.privateKey.export({ type: 'pkcs8', format: 'pem' }), email, name: cn };
  };
  const strangerCa = rsa();
  return {
    ca: { key: caKeys.privateKey, cert: caCert },
    amina: person('Amina Rahman', 'amina@northwind.example'),
    owner: person('Archive Owner', 'owner@example.com'),
    bo: person('Bo Lindqvist', 'bo@contoso.example'),
    // Past its date: issued two years ago for a year.
    expired: person('Old Certificate', 'old@northwind.example', { from: new Date(Date.now() - 730 * DAY), to: new Date(Date.now() - 365 * DAY) }),
    // Issued by an authority this computer has never heard of.
    stranger: person('Unknown Signer', 'stranger@elsewhere.example', { issuerKey: strangerCa.privateKey, issuer: { cn: 'Elsewhere Authority Root', o: 'Elsewhere' } }),
    // An elliptic-curve signer.
    ecSigner: person('Curve Signer', 'curve@northwind.example', { ec: true }),
  };
}

/** A person's .p12, as Windows exports one. */
export function makePkcs12(person, password, { legacy = false, chain = [] } = {}) {
  return writePkcs12({ key: person.key, cert: person.cert, chain, password, legacy, name: person.name });
}

export { printable };
