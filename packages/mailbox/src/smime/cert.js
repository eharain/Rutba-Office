// A certificate, as S/MIME needs to know it.
//
// Node's X509Certificate checks signatures and dates; what it does not give
// - the issuer's name and the serial number exactly as encoded, which name a
// certificate inside a signature or an encrypted message, the subject key
// identifier, the addresses it is for and what it may be used for - is read
// here from its DER.

import crypto from 'node:crypto';
import { parse, oidOf, textOf, timeOf, isContext } from './der.js';
import { OID } from './oids.js';

/** PEM blocks of a kind in a text, as DER. */
export function pemBlocks(text, label = 'CERTIFICATE') {
  const out = [];
  const re = new RegExp(`-----BEGIN ${label}-----([\\s\\S]*?)-----END ${label}-----`, 'g');
  for (const m of String(text).matchAll(re)) out.push(Buffer.from(m[1].replace(/[^A-Za-z0-9+/=]/g, ''), 'base64'));
  return out;
}

/** DER as PEM. */
export const pem = (der, label = 'CERTIFICATE') => `-----BEGIN ${label}-----\n${Buffer.from(der).toString('base64').replace(/.{64}/g, '$&\n').replace(/\n$/, '')}\n-----END ${label}-----\n`;

function nameParts(node) {
  const parts = [];
  for (const rdn of node.children || []) {
    for (const atv of rdn.children || []) {
      const [type, value] = atv.children || [];
      if (type && value) parts.push({ oid: oidOf(type), value: textOf(value) });
    }
  }
  return parts;
}

/**
 * Read a certificate (DER, PEM text, or an X509Certificate).
 * @returns {{ x509, der, issuerRaw, serialRaw, serial, subjectKeyId, name, email, emails, issuerName, organisation, validFrom, validTo, thumbprint, ca, usage }}
 */
export function certificate(input) {
  let x509;
  if (input instanceof crypto.X509Certificate) x509 = input;
  else if (typeof input === 'string') x509 = new crypto.X509Certificate(input);
  else x509 = new crypto.X509Certificate(Buffer.from(input));
  const der = x509.raw;
  const root = parse(der);
  const tbs = root.children[0];
  let i = 0;
  if (isContext(tbs.children[0], 0)) i++;
  const serial = tbs.children[i];
  const issuer = tbs.children[i + 2];
  const subject = tbs.children[i + 4];
  const subjectParts = nameParts(subject);
  const issuerParts = nameParts(issuer);
  const cn = (parts) => parts.filter((p) => p.oid === OID.commonName).pop()?.value || null;

  let subjectKeyId = null;
  const emails = [];
  let usage = null;
  let ca = false;
  const extensions = tbs.children.find((c) => isContext(c, 3));
  for (const ext of extensions?.children?.[0]?.children || []) {
    const id = oidOf(ext.children[0]);
    const value = parse(ext.children[ext.children.length - 1].value);
    if (id === OID.subjectKeyIdentifier) subjectKeyId = Buffer.from(value.value).toString('hex');
    else if (id === OID.subjectAltName) {
      for (const g of value.children || []) if (g.cls === 2 && g.number === 1) emails.push(Buffer.from(g.value).toString('latin1'));
    } else if (id === OID.extKeyUsage) usage = (value.children || []).map(oidOf);
    else if (id === OID.basicConstraints) ca = Boolean(value.children?.[0]?.number === 1 && value.children[0].value[0]);
  }
  for (const p of subjectParts) if (p.oid === OID.emailAddress && !emails.includes(p.value)) emails.push(p.value);
  return {
    x509,
    der,
    issuerRaw: Buffer.from(issuer.raw),
    serialRaw: Buffer.from(serial.value),
    serial: Buffer.from(serial.value).toString('hex').replace(/^(00)+(?=.)/, '').toUpperCase(),
    subjectKeyId,
    name: cn(subjectParts) || emails[0] || x509.subject,
    email: emails[0]?.toLowerCase() || null,
    emails: emails.map((e) => e.toLowerCase()),
    issuerName: cn(issuerParts) || x509.issuer,
    organisation: subjectParts.find((p) => p.oid === '2.5.4.10')?.value || null,
    validFrom: new Date(x509.validFrom),
    validTo: new Date(x509.validTo),
    thumbprint: x509.fingerprint.replace(/:/g, ''),
    ca: ca || x509.ca,
    usage,
    selfIssued: issuer.raw.length === subject.raw.length && Buffer.from(issuer.raw).equals(Buffer.from(subject.raw)),
  };
}

/** May the certificate be used for mail (no extended key usage at all, or one that includes mail)? */
export const forMail = (c) => !c.usage || c.usage.includes(OID.emailProtection) || c.usage.includes(OID.anyExtendedKeyUsage);

/** The facts about a certificate a window shows, with nothing that cannot be sent to it. */
export const describe = (c) => ({
  name: c.name,
  email: c.email,
  emails: c.emails,
  issuer: c.issuerName,
  organisation: c.organisation,
  serial: c.serial,
  thumbprint: c.thumbprint,
  validFrom: c.validFrom.toISOString(),
  validTo: c.validTo.toISOString(),
});

export { timeOf };
