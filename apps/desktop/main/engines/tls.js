// How a linked database's encrypted connection is made. "Encrypted" with the
// certificate unchecked keeps out someone listening but not someone in the
// middle, so the certificate is checked, and a server with a certificate of
// its own, which is common on a person's own network, is the person's to accept.

/** The driver's TLS options for a link's settings. */
export function tlsOptions({ ssl = false, sslAnyCertificate = false } = {}) {
  return ssl ? { rejectUnauthorized: !sslAnyCertificate } : false;
}

const UNCHECKED = /self.signed|unable to verify|certificate|altnames|CERT_|DEPTH_ZERO/i;

/** A failed certificate check in words a person can act on, or null if it was something else. */
export function certificateProblem(err, host) {
  const text = `${err?.code || ''} ${err?.message || err}`;
  if (!UNCHECKED.test(text)) return null;
  return `The certificate of ${host} could not be checked, so nothing was sent to it. If this is your own server with its own certificate, turn on "Accept a certificate that cannot be checked" for this link.`;
}
