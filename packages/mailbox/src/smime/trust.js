// Whose certificate is it, and does anyone this computer trusts vouch for it.
//
// A signer's certificate is followed to an authority: through the
// certificates the message carries and any kept from earlier mail, each link
// checked (the issuer's name, and its signature on the certificate below),
// until one is an authority this computer trusts - the Windows certificate
// store's trusted roots, read from the computer itself, as Node reads them
// (and on a computer without one, the roots Node carries). Nothing is fetched
// from anywhere: an issuer not on the computer or in the message is not
// found, and a certificate's revocation, which needs its authority asked
// over the network, is not checked; the window says so.

import tls from 'node:tls';
import { certificate } from './cert.js';

let system = null;

/** The trusted roots of this computer, read once. */
export function systemAnchors() {
  if (system) return system;
  let pems = [];
  try {
    pems = tls.getCACertificates?.('system') || [];
  } catch {
    pems = [];
  }
  if (!pems.length) pems = tls.rootCertificates || [];
  system = [];
  for (const p of pems) {
    try {
      system.push(certificate(p));
    } catch {
      /* a root Node cannot read is one this cannot use */
    }
  }
  return system;
}

/** For the tests: forget the roots read, or use these instead. */
export function setSystemAnchors(list) {
  system = list;
}

const issued = (child, parent) => {
  try {
    return child.x509.checkIssued(parent.x509) && child.x509.verify(parent.x509.publicKey);
  } catch {
    return false;
  }
};

/**
 * The chain from `cert` up to a trusted authority.
 * @param {object} cert from certificate()
 * @param {object[]} pool other certificates (the message's, ones kept)
 * @param {object[]} anchors trusted authorities
 * @returns {{ trusted: boolean, chain: object[], anchor: object|null }}
 */
export function chainOf(cert, pool = [], anchors = systemAnchors()) {
  const chain = [cert];
  const isAnchor = (c) => anchors.some((a) => a.thumbprint === c.thumbprint);
  let current = cert;
  for (let step = 0; step < 8; step++) {
    if (isAnchor(current)) return { trusted: true, chain, anchor: current };
    const anchor = anchors.find((a) => issued(current, a));
    if (anchor) {
      chain.push(anchor);
      return { trusted: true, chain, anchor };
    }
    if (current.selfIssued) break;
    const next = pool.find((c) => c.thumbprint !== current.thumbprint && !chain.includes(c) && issued(current, c));
    if (!next) break;
    chain.push(next);
    current = next;
  }
  return { trusted: false, chain, anchor: null };
}

/** Was every certificate in the chain in its dates at `at`. */
export const inDate = (chain, at) => chain.every((c) => c.validFrom <= at && at <= c.validTo);
