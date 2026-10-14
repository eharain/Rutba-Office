// ASN.1 in DER, the encoding certificates, signatures and encrypted mail are
// written in: just enough to read and write what S/MIME needs.
//
// Reading gives a tree of nodes, each with its tag, its contents and the
// exact bytes it was read from (`raw`), because a signature is over bytes as
// they were sent and a re-encoding that differs by one byte does not verify.
// Writing builds bytes from small functions (seq, set, oid, int...), so a
// structure reads in the code as it does in the RFC.

/** A node: `{ tag, cls, constructed, number, header, length, raw, value, children }`. */
export function read(bytes, at = 0, end = bytes.length) {
  const start = at;
  if (at + 2 > end) throw new Error('ASN.1: the data ends inside a tag');
  const tag = bytes[at++];
  let number = tag & 0x1f;
  if (number === 0x1f) {
    number = 0;
    let b;
    do {
      if (at >= end) throw new Error('ASN.1: the data ends inside a tag');
      b = bytes[at++];
      number = number * 128 + (b & 0x7f);
    } while (b & 0x80);
  }
  if (at >= end) throw new Error('ASN.1: the data ends inside a length');
  let length = bytes[at++];
  let indefinite = false;
  if (length === 0x80) indefinite = true;
  else if (length & 0x80) {
    const n = length & 0x7f;
    if (n > 6 || at + n > end) throw new Error('ASN.1: a length too long to be real');
    length = 0;
    for (let i = 0; i < n; i++) length = length * 256 + bytes[at++];
  }
  const header = at - start;
  const constructed = Boolean(tag & 0x20);
  const node = { tag, cls: tag >> 6, constructed, number, header, length: 0, raw: null, value: null, children: null };
  if (indefinite) {
    // BER, which some mail programs still write: read children up to the end-of-contents marker.
    if (!constructed) throw new Error('ASN.1: an indefinite length on a primitive value');
    node.children = [];
    while (true) {
      if (at + 2 > end) throw new Error('ASN.1: the data ends before its end-of-contents');
      if (bytes[at] === 0 && bytes[at + 1] === 0) {
        at += 2;
        break;
      }
      const child = read(bytes, at, end);
      node.children.push(child);
      at += child.raw.length;
    }
    node.raw = bytes.subarray(start, at);
    node.value = bytes.subarray(start + header, at - 2);
    node.length = node.value.length;
    return node;
  }
  if (at + length > end) throw new Error('ASN.1: a value runs past the end of the data');
  node.length = length;
  node.value = bytes.subarray(at, at + length);
  node.raw = bytes.subarray(start, at + length);
  if (constructed) {
    node.children = [];
    let p = at;
    while (p < at + length) {
      const child = read(bytes, p, at + length);
      node.children.push(child);
      p += child.raw.length;
    }
  }
  return node;
}

/** Parse bytes that hold exactly one value. */
export function parse(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  return read(b, 0, b.length);
}

/** Is the node `[n]` in the context class. */
export const isContext = (node, n) => node && node.cls === 2 && node.number === n;

/** An OBJECT IDENTIFIER's dotted text. */
export function oidOf(node) {
  const v = node.value;
  const parts = [];
  let n = 0;
  for (let i = 0; i < v.length; i++) {
    n = n * 128 + (v[i] & 0x7f);
    if (!(v[i] & 0x80)) {
      if (!parts.length) parts.push(n < 80 ? Math.floor(n / 40) : 2, n < 80 ? n % 40 : n - 80);
      else parts.push(n);
      n = 0;
    }
  }
  return parts.join('.');
}

/** An INTEGER as a BigInt. */
export function intOf(node) {
  let n = 0n;
  for (const b of node.value) n = (n << 8n) | BigInt(b);
  if (node.value.length && node.value[0] & 0x80) n -= 1n << BigInt(node.value.length * 8);
  return n;
}

/** A string type's text (UTF8String, PrintableString, IA5String, BMPString...). */
export function textOf(node) {
  if (node.number === 30) {
    let s = '';
    for (let i = 0; i + 1 < node.value.length; i += 2) s += String.fromCharCode((node.value[i] << 8) | node.value[i + 1]);
    return s;
  }
  return Buffer.from(node.value).toString(node.number === 12 ? 'utf8' : 'latin1');
}

/** A UTCTime or GeneralizedTime as a Date. */
export function timeOf(node) {
  const s = Buffer.from(node.value).toString('latin1');
  // UTCTime (tag 23) has a two-digit year, GeneralizedTime (24) four.
  const utc = node.number === 23;
  const m = (utc ? /^(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})?Z$/ : /^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})?(?:\.\d+)?Z$/).exec(s);
  if (!m) return null;
  let year = Number(m[1]);
  if (utc) year += year < 50 ? 2000 : 1900;
  return new Date(Date.UTC(year, Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6] || 0)));
}

// ── Writing ───────────────────────────────────────────────────────────────

function lengthBytes(n) {
  if (n < 0x80) return [n];
  const out = [];
  while (n > 0) {
    out.unshift(n % 256);
    n = Math.floor(n / 256);
  }
  return [0x80 | out.length, ...out];
}

const concat = (parts) => Buffer.concat(parts.map((p) => (Buffer.isBuffer(p) ? p : Buffer.from(p))));

/** A value with its tag. */
export function tlv(tag, contents) {
  // An array of numbers is bytes; an array of anything else is parts, one after another.
  const body = Array.isArray(contents) ? (contents.every((x) => typeof x === 'number') ? Buffer.from(contents) : concat(contents)) : Buffer.from(contents);
  return Buffer.concat([Buffer.from([tag, ...lengthBytes(body.length)]), body]);
}

export const seq = (...items) => tlv(0x30, items.filter(Boolean));
/** A SET OF, its members in DER's order (sorted by their encodings). */
export const set = (...items) => tlv(0x31, items.filter(Boolean).sort(Buffer.compare));
export const nul = () => Buffer.from([0x05, 0x00]);
export const octets = (bytes) => tlv(0x04, Buffer.from(bytes));
export const utf8 = (s) => tlv(0x0c, Buffer.from(String(s), 'utf8'));
export const printable = (s) => tlv(0x13, Buffer.from(String(s), 'latin1'));
export const ia5 = (s) => tlv(0x16, Buffer.from(String(s), 'latin1'));
export const bool = (v) => tlv(0x01, [v ? 0xff : 0]);
/** BIT STRING of whole bytes. */
export const bits = (bytes) => tlv(0x03, [Buffer.from([0]), Buffer.from(bytes)]);
/** `[n]` EXPLICIT around what is given (constructed). */
export const explicit = (n, ...items) => tlv(0xa0 | n, items.filter(Boolean));
/** `[n]` IMPLICIT: a constructed value's contents under a context tag. */
export const implicit = (n, encoded, constructed = true) => {
  const node = parse(encoded);
  return tlv((constructed ? 0xa0 : 0x80) | n, Buffer.from(node.value));
};
/** `[n]` IMPLICIT for primitive contents given as bytes. */
export const contextPrimitive = (n, bytes) => tlv(0x80 | n, Buffer.from(bytes));

export function oid(dotted) {
  const parts = String(dotted).split('.').map(Number);
  const out = [parts[0] * 40 + parts[1]];
  for (const n of parts.slice(2)) {
    const b = [n & 0x7f];
    let v = Math.floor(n / 128);
    while (v > 0) {
      b.unshift(0x80 | (v & 0x7f));
      v = Math.floor(v / 128);
    }
    out.push(...b);
  }
  return tlv(0x06, out);
}

/** An INTEGER from a number, a BigInt or its bytes (unsigned, big-endian). */
export function int(value) {
  let bytes;
  if (typeof value === 'number' || typeof value === 'bigint') {
    let n = BigInt(value);
    if (n === 0n) bytes = [0];
    else {
      bytes = [];
      while (n > 0n) {
        bytes.unshift(Number(n & 0xffn));
        n >>= 8n;
      }
    }
  } else bytes = [...value];
  while (bytes.length > 1 && bytes[0] === 0 && !(bytes[1] & 0x80)) bytes.shift();
  if (bytes[0] & 0x80) bytes.unshift(0);
  return tlv(0x02, bytes);
}

/** A time as UTCTime before 2050 and GeneralizedTime from it, as RFC 5280 says. */
export function time(date) {
  const d = new Date(date);
  const p = (n, w = 2) => String(n).padStart(w, '0');
  const rest = `${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}Z`;
  const y = d.getUTCFullYear();
  if (y >= 1950 && y < 2050) return tlv(0x17, Buffer.from(`${p(y % 100)}${rest}`, 'latin1'));
  return tlv(0x18, Buffer.from(`${p(y, 4)}${rest}`, 'latin1'));
}

/** An AlgorithmIdentifier, with NULL parameters unless others are given (or `false` for none). */
export const algorithm = (dotted, params = nul()) => seq(oid(dotted), params || null);
