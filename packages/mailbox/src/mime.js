// Reading a message.
//
// RFC 5322 headers, RFC 2045 bodies, RFC 2047 encoded words and RFC 2231
// parameter continuations — the four specifications between a byte stream and
// something a person can read. All of it here, with no dependencies, because a
// mail client that cannot open a message without a package manager is not a
// mail client anyone can rely on in ten years.
//
// The parse is deliberately forgiving. Real mail is malformed constantly: bare
// line feeds where CRLF is required, a Content-Type with no boundary, a base64
// body with a stray space, headers in an encoding nobody declared. Every one of
// those has a sensible reading, and refusing the message is never it.

const CRLF = /\r\n|\r|\n/;

function toBytes(input) {
  if (typeof input === 'string') return new TextEncoder().encode(input);
  if (input instanceof Uint8Array) return input;
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  return new Uint8Array(input);
}

const latin1 = new TextDecoder('latin1');

/** Split a message into its raw header block and body bytes. */
export function splitMessage(input) {
  const bytes = toBytes(input);
  // The blank line that ends the headers, tolerating CRLF, LF and CR.
  for (let i = 0; i < bytes.length - 1; i++) {
    if (bytes[i] === 10 && bytes[i + 1] === 10) return { head: latin1.decode(bytes.subarray(0, i)), body: bytes.subarray(i + 2) };
    if (bytes[i] === 13 && bytes[i + 1] === 10 && bytes[i + 2] === 13 && bytes[i + 3] === 10) {
      return { head: latin1.decode(bytes.subarray(0, i)), body: bytes.subarray(i + 4) };
    }
  }
  return { head: latin1.decode(bytes), body: new Uint8Array(0) };
}

/** Unfold and split header lines into ordered name/value pairs. */
export function parseHeaders(head) {
  const lines = head.split(CRLF);
  const out = [];
  for (const line of lines) {
    if (!line) continue;
    if (/^[ \t]/.test(line) && out.length) {
      out[out.length - 1].value += ` ${line.trim()}`;
      continue;
    }
    const at = line.indexOf(':');
    if (at < 0) continue;
    out.push({ name: line.slice(0, at).trim(), key: line.slice(0, at).trim().toLowerCase(), value: line.slice(at + 1).trim() });
  }
  const map = new Map();
  for (const h of out) {
    if (!map.has(h.key)) map.set(h.key, []);
    map.get(h.key).push(h.value);
  }
  return {
    list: out,
    get: (name) => map.get(String(name).toLowerCase())?.[0] ?? null,
    all: (name) => map.get(String(name).toLowerCase()) ?? [],
    has: (name) => map.has(String(name).toLowerCase()),
    map,
  };
}

function decodeCharset(bytes, charset) {
  const label = (charset || 'utf-8').toLowerCase().replace(/^["']|["']$/g, '');
  const aliases = { 'ks_c_5601-1987': 'euc-kr', 'gb2312': 'gbk', 'cp1252': 'windows-1252', 'unicode-1-1-utf-7': 'utf-8', 'iso-8859-8-i': 'iso-8859-8' };
  try {
    return new TextDecoder(aliases[label] || label, { fatal: false }).decode(bytes);
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

export function decodeBase64(text) {
  const clean = String(text).replace(/[^A-Za-z0-9+/=]/g, '');
  const pad = clean.length % 4;
  const padded = pad ? clean + '='.repeat(4 - pad) : clean;
  try {
    const bin = typeof atob === 'function' ? atob(padded) : Buffer.from(padded, 'base64').toString('latin1');
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  } catch {
    return new Uint8Array(0);
  }
}

export function decodeQuotedPrintable(text) {
  const s = String(text).replace(/=(?:\r\n|\n|\r)/g, ''); // soft line breaks
  const out = [];
  for (let i = 0; i < s.length; i++) {
    if (s[i] === '=' && i + 2 < s.length) {
      const hex = s.slice(i + 1, i + 3);
      if (/^[0-9a-fA-F]{2}$/.test(hex)) {
        out.push(parseInt(hex, 16));
        i += 2;
        continue;
      }
    }
    const code = s.charCodeAt(i);
    if (code > 255) {
      // A QP body that already contains decoded text — push its UTF-8 bytes.
      for (const b of new TextEncoder().encode(s[i])) out.push(b);
    } else out.push(code);
  }
  return new Uint8Array(out);
}

/** RFC 2047: `=?utf-8?B?…?=` inside a header value. */
export function decodeWords(value) {
  if (!value || !value.includes('=?')) return value || '';
  return String(value)
    // Adjacent encoded words are one string; the space between them is not data.
    .replace(/(=\?[^?]+\?[bBqQ]\?[^?]*\?=)\s+(?==\?)/g, '$1')
    .replace(/=\?([^?]+)\?([bBqQ])\?([^?]*)\?=/g, (m, charset, enc, text) => {
      try {
        const bytes = enc.toLowerCase() === 'b'
          ? decodeBase64(text)
          : decodeQuotedPrintable(text.replace(/_/g, ' '));
        return decodeCharset(bytes, charset);
      } catch {
        return m;
      }
    });
}

/**
 * Parse a structured header value with parameters, handling RFC 2231
 * continuations and percent-encoding: `name*0*=utf-8''%E2%82%AC; name*1*=uro`.
 */
export function parseParameters(value) {
  const [head, ...rest] = String(value || '').split(';');
  const params = {};
  const continued = {};
  for (const chunk of rest) {
    const at = chunk.indexOf('=');
    if (at < 0) continue;
    let key = chunk.slice(0, at).trim().toLowerCase();
    let val = chunk.slice(at + 1).trim().replace(/^"(.*)"$/s, '$1');
    const m = /^([^*]+)\*(\d+)?(\*)?$/.exec(key);
    if (m) {
      const base = m[1];
      const idx = m[2] ? Number(m[2]) : 0;
      continued[base] = continued[base] || { parts: [], extended: false };
      continued[base].parts[idx] = val;
      if (m[3]) continued[base].extended = true;
      continue;
    }
    params[key] = val;
  }
  for (const [key, { parts, extended }] of Object.entries(continued)) {
    let joined = parts.filter((p) => p != null).join('');
    if (extended) {
      const m = /^([^']*)'([^']*)'(.*)$/s.exec(joined);
      const charset = m ? m[1] : 'utf-8';
      const raw = m ? m[3] : joined;
      const bytes = new Uint8Array(
        [...raw.matchAll(/%([0-9a-fA-F]{2})|(.)/g)].map((x) => (x[1] ? parseInt(x[1], 16) : x[2].charCodeAt(0)))
      );
      joined = decodeCharset(bytes, charset);
    }
    params[key] = joined;
  }
  return { value: head.trim(), params };
}

/** `"A Person" <a@b.c>, other@d.e` → addresses. */
export function parseAddresses(value) {
  if (!value) return [];
  const out = [];
  let depth = 0;
  let quote = false;
  let current = '';
  const flush = () => {
    const raw = current.trim();
    current = '';
    if (!raw) return;
    const angle = /<([^>]*)>/.exec(raw);
    let address = angle ? angle[1].trim() : raw;
    let name = angle ? raw.slice(0, angle.index).trim() : '';
    if (!angle) {
      // `a@b.c (A Person)` is the other legal shape.
      const paren = /\(([^)]*)\)/.exec(raw);
      if (paren) {
        name = paren[1].trim();
        address = raw.replace(/\([^)]*\)/, '').trim();
      }
    }
    name = decodeWords(name).replace(/^"(.*)"$/s, '$1').trim();
    address = address.replace(/^mailto:/i, '').trim();
    if (!address && !name) return;
    out.push({ name: name || null, address, display: name ? `${name} <${address}>` : address });
  };
  for (const ch of String(value)) {
    if (ch === '"') quote = !quote;
    if (!quote) {
      if (ch === '<' || ch === '(') depth++;
      else if (ch === '>' || ch === ')') depth--;
      else if ((ch === ',' || ch === ';') && depth <= 0) {
        flush();
        continue;
      }
    }
    current += ch;
  }
  flush();
  return out;
}

function decodeBody(bytes, encoding) {
  const enc = (encoding || '7bit').toLowerCase().trim();
  if (enc === 'base64') return decodeBase64(latin1.decode(bytes));
  if (enc === 'quoted-printable') return decodeQuotedPrintable(latin1.decode(bytes));
  return bytes;
}

/** Find `--boundary` lines and slice the parts between them. */
function splitMultipart(bytes, boundary) {
  const marker = new TextEncoder().encode(`--${boundary}`);
  const positions = [];
  outer: for (let i = 0; i <= bytes.length - marker.length; i++) {
    // A boundary is only a boundary at the start of a line.
    if (i > 0 && bytes[i - 1] !== 10 && bytes[i - 1] !== 13) continue;
    for (let j = 0; j < marker.length; j++) if (bytes[i + j] !== marker[j]) continue outer;
    positions.push(i);
  }
  const parts = [];
  for (let k = 0; k < positions.length - 1; k++) {
    let start = positions[k] + marker.length;
    // Skip the CRLF that ends the boundary line.
    while (start < bytes.length && (bytes[start] === 13 || bytes[start] === 10)) start++;
    let end = positions[k + 1];
    while (end > start && (bytes[end - 1] === 13 || bytes[end - 1] === 10)) end--;
    parts.push(bytes.subarray(start, end));
  }
  return parts;
}

let partSeq = 0;

function parsePart(bytes, depth = 0) {
  const { head, body } = splitMessage(bytes);
  const headers = parseHeaders(head);
  const ctRaw = headers.get('content-type') || 'text/plain';
  const { value: mime, params } = parseParameters(ctRaw);
  const type = mime.toLowerCase();
  const encoding = headers.get('content-transfer-encoding') || '7bit';
  const dispRaw = headers.get('content-disposition') || '';
  const { value: disposition, params: dispParams } = parseParameters(dispRaw);
  const id = `p${++partSeq}`;

  const node = {
    id,
    type,
    charset: params.charset || null,
    encoding,
    disposition: (disposition || '').toLowerCase() || null,
    filename: decodeWords(dispParams.filename || params.name || '') || null,
    contentId: (headers.get('content-id') || '').replace(/^<|>$/g, '') || null,
    headers,
    size: body.length,
    children: [],
  };

  if (type.startsWith('multipart/') && depth < 24) {
    const boundary = params.boundary;
    if (boundary) {
      node.children = splitMultipart(body, boundary).map((p) => parsePart(p, depth + 1));
      return node;
    }
    // A multipart with no boundary: treat what is there as one text part rather
    // than showing the reader nothing.
    node.type = 'text/plain';
  }

  if (type === 'message/rfc822' && depth < 24) {
    node.children = [parsePart(decodeBody(body, encoding), depth + 1)];
    return node;
  }

  node.bytes = decodeBody(body, encoding);
  if (node.type.startsWith('text/')) node.text = decodeCharset(node.bytes, node.charset || 'utf-8');
  return node;
}

function collectText(node, out) {
  if (node.type === 'text/plain' && node.disposition !== 'attachment') out.plain.push(node.text || '');
  else if (node.type === 'text/html' && node.disposition !== 'attachment') out.html.push(node.text || '');
  for (const c of node.children) collectText(c, out);
  return out;
}

function collectAttachments(node, out = []) {
  const isAttachment =
    node.disposition === 'attachment' ||
    (node.filename && !node.type.startsWith('multipart/')) ||
    (node.contentId && !node.type.startsWith('text/') && !node.type.startsWith('multipart/'));
  if (isAttachment && node.bytes) {
    out.push({
      id: node.id,
      filename: node.filename || `attachment-${out.length + 1}${extensionFor(node.type)}`,
      type: node.type,
      size: node.bytes.length,
      contentId: node.contentId,
      inline: node.disposition === 'inline' || (!!node.contentId && node.disposition !== 'attachment'),
      bytes: node.bytes,
    });
  }
  for (const c of node.children) collectAttachments(c, out);
  return out;
}

function extensionFor(type) {
  const map = { 'application/pdf': '.pdf', 'image/png': '.png', 'image/jpeg': '.jpg', 'text/calendar': '.ics', 'message/rfc822': '.eml' };
  return map[type] || '';
}

/**
 * Parse a whole message.
 * @param {Uint8Array|string|Buffer} input raw RFC 5322 bytes
 * @returns {object} a message the client can render without further work
 */
export function parseMessage(input) {
  const root = parsePart(toBytes(input));
  const h = root.headers;
  const texts = collectText(root, { plain: [], html: [] });
  const attachments = collectAttachments(root);
  const date = h.get('date');
  const parsedDate = date ? new Date(date.replace(/\s*\([^)]*\)\s*$/, '')) : null;

  return {
    subject: decodeWords(h.get('subject') || ''),
    from: parseAddresses(h.get('from')),
    to: parseAddresses(h.get('to')),
    cc: parseAddresses(h.get('cc')),
    bcc: parseAddresses(h.get('bcc')),
    replyTo: parseAddresses(h.get('reply-to')),
    date: parsedDate && !Number.isNaN(parsedDate.getTime()) ? parsedDate.toISOString() : null,
    messageId: (h.get('message-id') || '').replace(/^<|>$/g, '') || null,
    inReplyTo: (h.get('in-reply-to') || '').replace(/^<|>$/g, '') || null,
    references: (h.get('references') || '').split(/\s+/).map((r) => r.replace(/^<|>$/g, '')).filter(Boolean),
    listId: h.get('list-id'),
    priority: h.get('x-priority') || h.get('importance') || null,
    text: texts.plain.join('\n').trim() || null,
    html: texts.html.join('\n').trim() || null,
    attachments: attachments.filter((a) => !a.inline || a.filename),
    inlineImages: attachments.filter((a) => a.inline && a.contentId),
    headers: h.list,
    structure: root,
    hasAttachments: attachments.some((a) => !a.inline),
  };
}

/** Everything but the bytes — for a message list that must stay light. */
export function summarize(message, extra = {}) {
  return {
    subject: message.subject || '(no subject)',
    from: message.from[0] || null,
    to: message.to,
    date: message.date,
    messageId: message.messageId,
    preview: (message.text || stripHtml(message.html || '')).replace(/\s+/g, ' ').slice(0, 220),
    hasAttachments: message.hasAttachments,
    attachmentCount: message.attachments.length,
    ...extra,
  };
}

/** A plain-text shadow of an HTML body, for previews and search. */
export function stripHtml(html) {
  return String(html || '')
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|li|h[1-6])>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (m, d) => String.fromCharCode(Number(d)))
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export { decodeCharset, toBytes };
