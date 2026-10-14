// S/MIME (RFC 8551): signed and encrypted mail, read and written.
//
// Reading takes a message as it arrived, every byte of it - a signature is
// over the bytes as sent, so a message parsed and put back together would
// not verify - and peels it: an encrypted layer is opened with a certificate
// and key the person imported, a signed layer is checked (the content
// unchanged since it was signed, the signature the signer's, the signer's
// certificate one an authority on this computer vouches for, in its dates,
// for mail, for the address the message is from). What is inside is parsed
// as the message to show, with what was found said beside it.
//
// Writing takes a message as the mail program would send it and signs it
// (multipart/signed, so a reader without S/MIME still reads it) and/or
// encrypts it for its recipients' certificates and the sender's own, so the
// sent copy can be read again.
//
// Nothing here goes to the network. Revocation is not checked; the window
// says so.

import crypto from 'node:crypto';
import { parseMessage, parseHeaders, parseParameters, decodeBase64 } from '../mime.js';
import { contentInfo, readSignedData, verifySigners, signDetached, readEnveloped, decryptEnveloped, encryptEnveloped, CmsError } from './cms.js';
import { certificate, describe, forMail, forSigning, forEncrypting, pemBlocks, pem } from './cert.js';
import { chainOf, inDate, systemAnchors, setSystemAnchors } from './trust.js';
import { readPkcs12, writePkcs12, Pkcs12Error } from './pkcs12.js';
import { OID } from './oids.js';

const latin1 = (b) => Buffer.from(b).toString('latin1');

/** The largest message or CMS structure read: past this a message is said to be too large. */
export const MAX_SMIME_BYTES = 50 * 1048576;

const CR = 13;
const LF = 10;

/**
 * Where the headers end: the index of the blank line's start, and where the
 * body begins. A blank line however its ends were written (CRLF, LF, or the
 * CR CR LF a program writing CRLF in text mode on Windows leaves), found in
 * one pass over the bytes: every byte is looked at a fixed number of times,
 * so a run of carriage returns cannot make it slow.
 */
export function headEnd(bytes) {
  const n = Math.min(bytes.length, 262144);
  for (let i = 0; i < n; i++) {
    if (bytes[i] !== LF) continue;
    let j = i + 1;
    while (j < n && bytes[j] === CR) j++;
    if (j < n && bytes[j] === LF) {
      let k = i;
      while (k > 0 && bytes[k - 1] === CR) k--;
      return { head: k, body: j + 1 };
    }
    // The carriage returns just passed are not a blank line's: carry on after them.
    i = j - 1;
  }
  return { head: bytes.length, body: bytes.length };
}

/** A message or entity split into its header text and body bytes. */
export function splitEntity(bytes) {
  const b = Buffer.from(bytes);
  const { head, body } = headEnd(b);
  return { headText: latin1(b.subarray(0, head)), body: b.subarray(body) };
}

/**
 * Every line ending CRLF, as MIME is signed (stray carriage returns before a
 * newline are one), in one pass over the bytes.
 */
export function canonical(bytes) {
  const b = Buffer.from(bytes);
  const out = Buffer.allocUnsafe(b.length * 2);
  let o = 0;
  for (let i = 0; i < b.length; i++) {
    const c = b[i];
    if (c === CR) {
      let j = i;
      while (j < b.length && b[j] === CR) j++;
      if (j < b.length && b[j] === LF) {
        out[o++] = CR;
        out[o++] = LF;
        i = j;
      } else {
        // Carriage returns not before a newline stay as they were.
        for (let k = i; k < j; k++) out[o++] = CR;
        i = j - 1;
      }
    } else if (c === LF) {
      out[o++] = CR;
      out[o++] = LF;
    } else out[o++] = c;
  }
  return Buffer.from(out.subarray(0, o));
}

/** Does the content have a line end that is not CRLF (a bare LF, or CR CR LF), by one pass. */
function oddLineEnds(bytes) {
  for (let i = 0; i < bytes.length; i++) {
    if (bytes[i] !== LF) continue;
    if (i === 0 || bytes[i - 1] !== CR) return true;
    if (i > 1 && bytes[i - 2] === CR) return true;
  }
  return false;
}

/** The header lines of a header text, each with its continuation lines. */
function headerLines(headText) {
  const out = [];
  for (const line of headText.split(/\r?\n/)) {
    if (/^[ \t]/.test(line) && out.length) out[out.length - 1] += `\r\n${line}`;
    else if (line) out.push(line);
  }
  return out;
}

const isContentHeader = (line) => /^content-/i.test(line);

/** What kind of S/MIME, if any, a message or entity is, from its Content-Type. */
export function smimeKind(bytesOrHeaders) {
  const headers = typeof bytesOrHeaders?.get === 'function' ? bytesOrHeaders : parseHeaders(splitEntity(bytesOrHeaders).headText);
  const { value, params } = parseParameters(headers.get('content-type') || 'text/plain');
  const type = value.toLowerCase();
  if (type === 'multipart/signed' && /pkcs7-signature/i.test(params.protocol || '')) return 'signed';
  if (type === 'application/pkcs7-mime' || type === 'application/x-pkcs7-mime') {
    const st = String(params['smime-type'] || '').toLowerCase();
    if (st === 'signed-data') return 'opaque-signed';
    if (st === 'enveloped-data' || st === 'authenveloped-data') return 'enveloped';
    return 'pkcs7';
  }
  return null;
}

/** The two parts of a multipart/signed body, exactly as sent (RFC 1847: the CRLF before a boundary belongs to it). */
function signedParts(body, boundary) {
  const s = latin1(body);
  const marker = `--${boundary}`;
  const starts = [];
  let at = 0;
  while ((at = s.indexOf(marker, at)) >= 0) {
    if (at === 0 || s[at - 1] === '\n') starts.push(at);
    at += marker.length;
  }
  if (starts.length < 3) throw new CmsError('damaged', 'The signed message is incomplete.');
  const lineEnd = (p) => {
    const nl = s.indexOf('\n', p);
    return nl < 0 ? s.length : nl + 1;
  };
  const partOf = (k) => {
    const begin = lineEnd(starts[k]);
    let end = starts[k + 1];
    // The line end before the boundary: CRLF, LF, or CR CR LF from a text-mode writer.
    if (s[end - 1] === '\n') {
      end--;
      while (s[end - 1] === '\r') end--;
    }
    return body.subarray(begin, end);
  };
  return [partOf(0), partOf(1)];
}

/** An entity's body, decoded by its Content-Transfer-Encoding. */
function decodedBody(entity) {
  const { headText, body } = splitEntity(entity);
  const enc = (parseHeaders(headText).get('content-transfer-encoding') || '').toLowerCase();
  if (enc === 'base64') return Buffer.from(decodeBase64(latin1(body)));
  return Buffer.from(body);
}

/** A message made of the outer headers (less their Content- ones) and an inner entity. */
function joinOuter(outerHeadText, inner) {
  const kept = headerLines(outerHeadText).filter((l) => !isContentHeader(l) && !/^mime-version:/i.test(l));
  return Buffer.concat([Buffer.from(`${kept.join('\r\n')}\r\nMIME-Version: 1.0\r\n`, 'latin1'), Buffer.from(inner)]);
}

/**
 * The statuses a signature can have, least worrying first. Only 'valid' is
 * shown as good: signed by someone this computer trusts, unchanged, in date
 * now, and for the address the message is from. 'other-sender' (a good
 * signature by someone other than the From address) and 'no-from' (no From
 * to compare) are said, not passed over.
 */
const SEVERITY = ['valid', 'other-sender', 'no-from', 'expired', 'untrusted', 'no-certificate', 'not-for-mail', 'unsupported', 'bad-signature', 'changed', 'damaged'];

/** One signer's result as the window says it. */
function judgeSigner(r, { pool, anchors, now, from }) {
  if (r.error === 'mismatch') return { status: 'bad-signature', signer: r.cert ? describe(r.cert) : null, digest: r.digest };
  if (r.error === 'algorithm') return { status: 'unsupported', signer: r.cert ? describe(r.cert) : null };
  if (!r.cert) return { status: r.intact ? 'no-certificate' : 'changed', signer: null, signedAt: r.signedAt?.toISOString() || null };
  const chain = chainOf(r.cert, pool, anchors);
  const address = String(from || '').toLowerCase();
  const fromMatches = address ? r.cert.emails.includes(address) : null;
  // In date now, and when it says it was signed (a date the signer chose, so never the only test).
  const dated = inDate(chain.chain, now) && (!r.signedAt || r.signedAt > now || inDate(chain.chain, r.signedAt));
  let status = 'valid';
  if (!r.intact) status = 'changed';
  else if (!r.valid) status = 'bad-signature';
  else if (!forSigning(r.cert)) status = 'not-for-mail';
  else if (!chain.trusted) status = 'untrusted';
  else if (!dated) status = 'expired';
  else if (!address) status = 'no-from';
  else if (!fromMatches) status = 'other-sender';
  // Weak hashes are said, whatever the status.
  const cautions = [];
  if (r.digest === 'sha1') cautions.push('sha1');
  return {
    status,
    signer: describe(r.cert),
    signedAt: r.signedAt?.toISOString() || null,
    digest: r.digest,
    trustedBy: chain.anchor ? chain.anchor.name : null,
    chain: chain.chain.map((c) => c.name),
    untrustedBecause: chain.trusted ? null : chain.reason || null,
    fromMatches,
    cautions,
    revocation: 'not-checked',
  };
}

/**
 * Read an S/MIME message (or any message: one that is not S/MIME comes back
 * parsed, with `signed` and `encrypted` null).
 * @param {Uint8Array} raw the message as it arrived
 * @param {{ identities?: Array<{cert, key}>, certificates?: object[], anchors?: object[], now?: Date }} [opts]
 * @returns {{ message: object, signed: object|null, encrypted: object|null, certificates: object[] }}
 *   `certificates` the signers' certificates found, to keep for encrypting to them later
 */
export function readSmime(raw, { identities = [], certificates = [], anchors = systemAnchors(), now = new Date() } = {}) {
  // A message larger than any real one is not read: said so, and the rest of the window carries on.
  if (raw.length > MAX_SMIME_BYTES) {
    const head = Buffer.from(raw).subarray(0, headEnd(raw).head);
    const m = parseMessage(Buffer.concat([head, Buffer.from('\r\n\r\n', 'latin1')]));
    delete m.structure;
    return { message: { ...m, text: null, html: null, attachments: [], hasAttachments: false }, signed: null, encrypted: null, error: 'too-large', certificates: [] };
  }
  let current = Buffer.from(raw);
  let outerHead = splitEntity(current).headText;
  const from = parseMessage(Buffer.from(`${outerHead}\r\n\r\n`, 'latin1')).from?.[0]?.address || null;
  let signed = null;
  let encrypted = null;
  const found = [];
  const shell = () => {
    const m = parseMessage(Buffer.from(`${headerLines(outerHead).filter((l) => !isContentHeader(l)).join('\r\n')}\r\nContent-Type: text/plain\r\n\r\n`, 'latin1'));
    delete m.structure;
    return { ...m, text: null, html: null, attachments: [], hasAttachments: false };
  };
  let kind = null;
  try {
    return peel();
  } catch (err) {
    // A layer that cannot even be read - cut short, or not what its header says.
    if (kind === 'signed' || kind === 'opaque-signed') {
      signed = { status: 'damaged', error: err.message, signers: [] };
      const message = parseMessage(current);
      delete message.structure;
      return { message, signed, encrypted, certificates: found };
    }
    encrypted = { status: 'damaged', error: err.message, recipients: 0 };
    return { message: shell(), signed, encrypted, certificates: found };
  }

  function peel() {
  for (let depth = 0; depth < 4; depth++) {
    const { headText, body } = splitEntity(current);
    const headers = parseHeaders(headText);
    kind = smimeKind(headers);
    if (!kind) break;
    if (kind === 'signed') {
      const { params } = parseParameters(headers.get('content-type'));
      const [content, sigPart] = signedParts(body, params.boundary);
      const sd = readSignedData(contentInfo(decodedBody(sigPart)).content);
      let results = verifySigners(sd, content, certificates);
      let shown = content;
      // A message whose line endings were changed on the way (stored with bare
      // newlines, or with a carriage return too many) is checked as it would
      // have been sent, and shown so.
      if (results.some((r) => !r.intact) && oddLineEnds(content)) {
        const fixed = canonical(content);
        const again = verifySigners(sd, fixed, certificates);
        if (again.every((r) => r.intact)) {
          results = again;
          shown = fixed;
        }
      }
      signed = summarise(results, { pool: [...sd.certificates, ...certificates], anchors, now, from, layer: signed });
      found.push(...sd.certificates);
      current = joinOuter(outerHead, shown);
      continue;
    }
    const der = decodedBody(current);
    const ci = contentInfo(der);
    if (ci.type === OID.signedData) {
      const sd = readSignedData(ci.content);
      const results = verifySigners(sd, null, certificates);
      signed = summarise(results, { pool: [...sd.certificates, ...certificates], anchors, now, from, layer: signed });
      found.push(...sd.certificates);
      current = joinOuter(outerHead, sd.content || Buffer.alloc(0));
      continue;
    }
    if (ci.type === OID.envelopedData || ci.type === OID.authEnvelopedData) {
      const env = readEnveloped(ci.content, ci.type === OID.authEnvelopedData);
      try {
        const opened = decryptEnveloped(env, identities);
        // RC2 and Triple DES open, with a caution: they are too weak to keep a message private now.
        encrypted = { status: 'decrypted', cipher: opened.cipher, recipient: describe(opened.identity.cert), recipients: env.recipients.length, cautions: /RC2|Triple DES/.test(opened.cipher) ? ['weak-cipher'] : [] };
        current = joinOuter(outerHead, opened.content);
        outerHead = splitEntity(current).headText;
        continue;
      } catch (err) {
        encrypted = { status: err.code === 'no-identity' ? 'no-identity' : err.code === 'changed' ? 'damaged' : err.code === 'recipient-kind' || err.code === 'algorithm' ? 'unsupported' : 'failed', error: err.message, recipients: env.recipients.length };
        return { message: shell(), signed, encrypted, certificates: found };
      }
    }
    break;
  }
  const message = parseMessage(current);
  delete message.structure;
  return { message, signed, encrypted, certificates: found };
  }
}

/** The signers' results as one: the worst of them, and each. */
function summarise(results, ctx) {
  const each = results.map((r) => judgeSigner(r, ctx));
  if (!each.length) return { status: 'no-certificate', signers: [] };
  const worst = each.reduce((a, b) => (SEVERITY.indexOf(b.status) > SEVERITY.indexOf(a.status) ? b : a));
  return { ...worst, signers: each };
}

/**
 * Sign and/or encrypt a whole message as it would be sent.
 * @param {Uint8Array|string} raw the message (headers and body)
 * @param {{ sign?: { key, cert, chain? }, encryptFor?: Array<string|Uint8Array>, now?: Date }} opts
 * @returns {Buffer} the message to send
 */
export function protect(raw, { sign = null, encryptFor = null, now = new Date() } = {}) {
  const { headText, body } = splitEntity(Buffer.from(raw));
  const lines = headerLines(headText);
  const contentHeaders = lines.filter(isContentHeader);
  let entity = canonical(Buffer.concat([Buffer.from(`${contentHeaders.join('\r\n') || 'Content-Type: text/plain; charset=utf-8'}\r\n\r\n`, 'latin1'), Buffer.from(body)]));
  if (sign) {
    const boundary = `----=_Signed_${crypto.randomBytes(12).toString('hex')}`;
    const signature = signDetached(entity, { key: sign.key, cert: sign.cert, chain: sign.chain || [], now });
    entity = Buffer.concat([
      Buffer.from(`Content-Type: multipart/signed; protocol="application/pkcs7-signature"; micalg=sha-256;\r\n\tboundary="${boundary}"\r\n\r\nThis is a cryptographically signed message in MIME format.\r\n\r\n--${boundary}\r\n`, 'latin1'),
      entity,
      Buffer.from(`\r\n--${boundary}\r\nContent-Type: application/pkcs7-signature; name="smime.p7s"\r\nContent-Transfer-Encoding: base64\r\nContent-Disposition: attachment; filename="smime.p7s"\r\nContent-Description: S/MIME Cryptographic Signature\r\n\r\n${signature.toString('base64').replace(/.{76}/g, '$&\r\n')}\r\n--${boundary}--\r\n`, 'latin1'),
    ]);
  }
  if (encryptFor?.length) {
    const cms = encryptEnveloped(entity, encryptFor);
    entity = Buffer.from(
      `Content-Type: application/pkcs7-mime; smime-type=enveloped-data; name="smime.p7m"\r\nContent-Transfer-Encoding: base64\r\nContent-Disposition: attachment; filename="smime.p7m"\r\nContent-Description: S/MIME Encrypted Message\r\n\r\n${cms.toString('base64').replace(/.{76}/g, '$&\r\n')}\r\n`,
      'latin1'
    );
  }
  const outer = lines.filter((l) => !isContentHeader(l) && !/^mime-version:/i.test(l));
  return Buffer.concat([Buffer.from(`${outer.join('\r\n')}\r\nMIME-Version: 1.0\r\n`, 'latin1'), entity]);
}

/**
 * The bytes of an S/MIME message an Outlook data file keeps as an
 * attachment (smime.p7m): the whole MIME entity of a signed message, or the
 * encrypted CMS itself. Null when the attachment is not one.
 */
export function entityFromAttachment(a) {
  if (!a?.bytes || !/^smime\.p7m$/i.test(a.filename || '')) return null;
  const bytes = Buffer.from(a.bytes);
  const head = latin1(bytes.subarray(0, 512));
  // An entity with its headers: its first line is a header.
  if (/^[A-Za-z][A-Za-z0-9-]*:/.test(head)) return bytes;
  // A multipart body without its headers: its first line names the boundary.
  const first = /^--([^\r\n]+)\r?\n/.exec(head);
  if (first) return Buffer.concat([Buffer.from(`Content-Type: multipart/signed; protocol="application/pkcs7-signature"; boundary="${first[1].trim()}"\r\n\r\n`, 'latin1'), bytes]);
  // CMS bytes as they are.
  if (bytes[0] === 0x30) return Buffer.concat([Buffer.from('Content-Type: application/pkcs7-mime\r\nContent-Transfer-Encoding: base64\r\n\r\n', 'latin1'), Buffer.from(bytes.toString('base64').replace(/.{76}/g, '$&\r\n'), 'latin1')]);
  return null;
}

export { readPkcs12, writePkcs12, Pkcs12Error, CmsError, certificate, describe, pem, pemBlocks, forMail, forEncrypting, systemAnchors, setSystemAnchors, chainOf };
export { forSigning };
