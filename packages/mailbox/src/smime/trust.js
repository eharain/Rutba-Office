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
//
// A chain that links up is not yet trusted. Every certificate above the
// signer's must be an authority (basicConstraints CA:TRUE; an old root with
// no such extension is taken as one, as Windows takes it), allowed to sign
// certificates (its key usage, when it has one), within its path length, for
// mail (its extended key usage, when it has one: e-mail protection or any
// purpose), and its name constraints on mail addresses kept; a certificate
// marked with a critical extension this does not understand is not trusted
// at all. Windows' own chain engine is not used: it is reached only through
// a native call, and left to itself it fetches issuers and revocation lists
// over the network, which this never does. Nor can the purposes a person set
// on a root in the Windows store be read without one; a root's own extended
// key usage is honoured.

import tls from 'node:tls';
import { certificate, forMail } from './cert.js';

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

/** Does a mail address fall under a name constraint (an address, a host, or ".domain" for any under it). */
function underConstraint(email, constraint) {
  const domain = email.split('@')[1] || '';
  if (constraint.includes('@')) return email === constraint;
  if (constraint.startsWith('.')) return domain.endsWith(constraint);
  return domain === constraint;
}

/**
 * Why a chain that links up is still not to be trusted, or null.
 * chain[0] is the signer's certificate, the last the authority it ends at.
 */
export function chainProblem(chain) {
  const leaf = chain[0];
  for (let i = 0; i < chain.length; i++) {
    const c = chain[i];
    if (c.unknownCritical?.length) return 'unknown-critical';
    if (i === 0) continue;
    const isRoot = i === chain.length - 1;
    // An authority: CA:TRUE, or an old root with no basicConstraints at all.
    if (!c.ca && !(isRoot && !c.basicConstraints)) return 'not-a-ca';
    if (c.keyUsage && !c.keyUsage.keyCertSign) return 'not-a-ca';
    // pathLen: how many authorities may stand between it and the signer's certificate.
    if (c.pathLen != null && i - 1 > c.pathLen) return 'path-length';
    if (!forMail(c)) return 'not-for-mail';
    const nc = c.nameConstraints;
    if (nc) {
      if (nc.other) return 'name-constraints';
      for (const email of leaf.emails || []) {
        if (nc.permitted.length && !nc.permitted.some((p) => underConstraint(email, p))) return 'name-constraints';
        if (nc.excluded.some((p) => underConstraint(email, p))) return 'name-constraints';
      }
    }
  }
  return null;
}

/**
 * The chain from `cert` up to a trusted authority.
 * @param {object} cert from certificate()
 * @param {object[]} pool other certificates (the message's, ones kept)
 * @param {object[]} anchors trusted authorities
 * @returns {{ trusted: boolean, chain: object[], anchor: object|null, reason?: string }}
 *   `reason`: 'no-authority' when no trusted authority was reached, or what chainProblem found
 */
export function chainOf(cert, pool = [], anchors = systemAnchors()) {
  const chain = [cert];
  const isAnchor = (c) => anchors.some((a) => a.thumbprint === c.thumbprint);
  const judged = (anchor) => {
    const reason = chainProblem(chain);
    return reason ? { trusted: false, chain, anchor: null, reason } : { trusted: true, chain, anchor };
  };
  let current = cert;
  for (let step = 0; step < 8; step++) {
    if (isAnchor(current)) return judged(current);
    const anchor = anchors.find((a) => issued(current, a));
    if (anchor) {
      chain.push(anchor);
      return judged(anchor);
    }
    if (current.selfIssued) break;
    const next = pool.find((c) => c.thumbprint !== current.thumbprint && !chain.includes(c) && issued(current, c));
    if (!next) break;
    chain.push(next);
    current = next;
  }
  return { trusted: false, chain, anchor: null, reason: 'no-authority' };
}

/** Was every certificate in the chain in its dates at `at`. */
export const inDate = (chain, at) => chain.every((c) => c.validFrom <= at && at <= c.validTo);
