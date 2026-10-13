// A message written back out as RFC 5322.
//
// An export for legal hold wants each message as a mail program would read
// it, in an mbox anything can open. A message read from an mbox is exported
// as the archive holds it, byte for byte; one read from a .pst, an .olm or
// the local store is held as its parts, and is written here: its headers
// (non-ASCII words encoded as RFC 2047 says), its text and HTML as a
// multipart/alternative, and its attachments in base64 under a
// multipart/mixed. The same message always gives the same bytes, so its hash
// can be checked again later.

import crypto from 'node:crypto';

const CRLF = '\r\n';
const ascii = (s) => /^[\x20-\x7e]*$/.test(s);

/** A header value, as an encoded word when it is not plain ASCII. */
function word(value) {
  const s = String(value ?? '');
  if (ascii(s)) return s;
  return `=?UTF-8?B?${Buffer.from(s, 'utf8').toString('base64')}?=`;
}

function address(p) {
  if (!p) return '';
  const name = p.name && p.name !== p.address ? p.name : '';
  if (!name) return `<${p.address || 'unknown'}>`;
  const shown = ascii(name) ? (/[(),.:;<>@[\]"\\]/.test(name) ? `"${name.replace(/["\\]/g, '\\$&')}"` : name) : word(name);
  return p.address ? `${shown} <${p.address}>` : shown;
}

const list = (people) => (people || []).map(address).filter(Boolean).join(`,${CRLF} `);

/** Fold base64 at 76 characters. */
const b64 = (bytes) => (Buffer.from(bytes).toString('base64').match(/.{1,76}/g) || ['']).join(CRLF);

/** A text part: plain 7-bit when it can be, else base64 UTF-8. */
function textPart(type, text) {
  const s = String(text ?? '').replace(/\r?\n/g, () => CRLF);
  if (ascii(s.replace(/\r\n/g, '')) && s.split(CRLF).every((l) => l.length <= 998)) {
    return `Content-Type: ${type}; charset=us-ascii${CRLF}Content-Transfer-Encoding: 7bit${CRLF}${CRLF}${s}${CRLF}`;
  }
  return `Content-Type: ${type}; charset=utf-8${CRLF}Content-Transfer-Encoding: base64${CRLF}${CRLF}${b64(Buffer.from(s, 'utf8'))}${CRLF}`;
}

/** A filename parameter, RFC 2231 when it is not ASCII. */
function filenameParam(name) {
  const s = String(name || 'attachment');
  if (ascii(s) && !/["\\]/.test(s)) return `filename="${s}"`;
  return `filename*=UTF-8''${encodeURIComponent(s)}`;
}

function boundary(seed, n) {
  return `=_rutba_${crypto.createHash('sha1').update(`${seed}|${n}`).digest('hex').slice(0, 24)}`;
}

/**
 * The message as bytes.
 * @param {object} m a parsed message: subject, from, to, cc, bcc, date, messageId, inReplyTo, references, text, html,
 *   attachments with `bytes`
 * @returns {Buffer}
 */
export function toRfc822(m) {
  const from = Array.isArray(m.from) ? m.from[0] : m.from;
  const head = [];
  const add = (name, value) => value && head.push(`${name}: ${value}`);
  add('From', address(from));
  add('To', list(m.to));
  add('Cc', list(m.cc));
  add('Bcc', list(m.bcc));
  add('Subject', word(m.subject || ''));
  if (m.date && !Number.isNaN(new Date(m.date).getTime())) add('Date', new Date(m.date).toUTCString().replace('GMT', '+0000'));
  const id = (v) => (v ? `<${String(v).replace(/^<|>$/g, '')}>` : '');
  add('Message-ID', id(m.messageId));
  add('In-Reply-To', id(m.inReplyTo));
  if (m.references?.length) add('References', m.references.map(id).join(' '));
  head.push('MIME-Version: 1.0');

  const seed = `${m.messageId || ''}|${m.subject || ''}|${m.date || ''}`;
  const text = m.text ?? null;
  const html = m.html ?? null;
  let body;
  const alternative = () => {
    if (text != null && html != null) {
      const b = boundary(seed, 1);
      return `Content-Type: multipart/alternative; boundary="${b}"${CRLF}${CRLF}--${b}${CRLF}${textPart('text/plain', text)}--${b}${CRLF}${textPart('text/html', html)}--${b}--${CRLF}`;
    }
    if (html != null) return textPart('text/html', html);
    return textPart('text/plain', text ?? '');
  };
  const files = (m.attachments || []).filter((a) => a && a.bytes);
  if (files.length) {
    const b = boundary(seed, 0);
    const parts = [alternative()];
    for (const a of files) {
      const type = a.type || 'application/octet-stream';
      const disposition = a.inline && a.contentId ? 'inline' : 'attachment';
      parts.push(
        `Content-Type: ${type}${CRLF}Content-Disposition: ${disposition}; ${filenameParam(a.filename)}${CRLF}` +
          (a.contentId ? `Content-ID: <${String(a.contentId).replace(/^<|>$/g, '')}>${CRLF}` : '') +
          `Content-Transfer-Encoding: base64${CRLF}${CRLF}${b64(a.bytes)}${CRLF}`
      );
    }
    body = `Content-Type: multipart/mixed; boundary="${b}"${CRLF}${CRLF}${parts.map((p) => `--${b}${CRLF}${p}`).join('')}--${b}--${CRLF}`;
  } else {
    body = alternative();
  }
  // The body's own headers follow the message's.
  return Buffer.from(`${head.join(CRLF)}${CRLF}${body}`, 'utf8');
}
