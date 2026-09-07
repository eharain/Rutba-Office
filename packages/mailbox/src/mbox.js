// mbox, and its relatives.
//
// The oldest mail archive format still in daily use: messages concatenated,
// each introduced by a line starting `From ` — the "From_ line", which is not a
// header and is not part of the message. Because that separator is just text, a
// message body containing a line starting with "From " has to be escaped, and
// the three families of mbox disagree about how. `mboxrd` prefixes such lines
// with `>`, which is the only variant that round-trips, and the one Thunderbird
// and most exporters write.
//
// Reading is therefore a matter of splitting on the separator and unescaping
// what the writer escaped, without corrupting an archive written by something
// that escaped nothing.

import { parseMessage, summarize, toBytes } from './mime.js';

const FROM_ = new TextEncoder().encode('From ');

function isFromLine(bytes, at) {
  if (at !== 0 && bytes[at - 1] !== 10) return false;
  for (let i = 0; i < FROM_.length; i++) if (bytes[at + i] !== FROM_[i]) return false;
  return true;
}

/** Byte ranges of each message in an mbox, without decoding any of them. */
export function scanMbox(input) {
  const bytes = toBytes(input);
  const starts = [];
  for (let i = 0; i <= bytes.length - FROM_.length; i++) {
    if (isFromLine(bytes, i)) starts.push(i);
  }
  const ranges = [];
  for (let i = 0; i < starts.length; i++) {
    const from = starts[i];
    const to = i + 1 < starts.length ? starts[i + 1] : bytes.length;
    // Step past the From_ line itself — it belongs to the archive, not the mail.
    let bodyStart = from;
    while (bodyStart < to && bytes[bodyStart] !== 10) bodyStart++;
    bodyStart++;
    ranges.push({ start: bodyStart, end: to, separator: new TextDecoder('latin1').decode(bytes.subarray(from, bodyStart - 1)) });
  }
  return { bytes, ranges };
}

/** Undo `>From ` quoting, and any run of `>` before it (mboxrd). */
function unescape(bytes) {
  const text = new TextDecoder('latin1').decode(bytes);
  if (!/^>+From /m.test(text)) return bytes;
  return new TextEncoder().encode(text.replace(/^(>+)(From )/gm, (m, gt, rest) => gt.slice(1) + rest));
}

/**
 * Read an mbox lazily: the index is cheap, and a message is only parsed when
 * something asks for it. A 4 GB archive opens instantly this way.
 */
export class Mbox {
  constructor(input) {
    const { bytes, ranges } = scanMbox(input);
    this.bytes = bytes;
    this.ranges = ranges;
  }

  get length() {
    return this.ranges.length;
  }

  raw(index) {
    const r = this.ranges[index];
    if (!r) return null;
    return unescape(this.bytes.subarray(r.start, r.end));
  }

  message(index) {
    const raw = this.raw(index);
    return raw ? parseMessage(raw) : null;
  }

  /** Headers only, for a message list — no bodies decoded. */
  summaries({ limit = Infinity, offset = 0 } = {}) {
    const out = [];
    for (let i = offset; i < Math.min(this.length, offset + limit); i++) {
      const msg = this.message(i);
      if (msg) out.push(summarize(msg, { index: i }));
    }
    return out;
  }

  *[Symbol.iterator]() {
    for (let i = 0; i < this.length; i++) yield this.message(i);
  }
}

/**
 * One byte per character, which is what a message body is.
 *
 * The escaping pass below works on text, and message bytes are not text — they
 * are whatever encoding the sender used. Decoding them as latin1 makes each
 * byte one character; encoding them back the same way returns exactly the bytes
 * that came in. Going out through UTF-8 instead would turn every byte above 127
 * into two, which is how an em-dash becomes three mojibake characters.
 */
function latin1Bytes(text) {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) out[i] = text.charCodeAt(i) & 0xff;
  return out;
}

/** Write messages back out as mboxrd. */
export function writeMbox(messages) {
  const enc = new TextEncoder();
  const chunks = [];
  for (const m of messages) {
    // A string given as `raw` is text and becomes UTF-8; bytes stay bytes.
    const raw = typeof m.raw === 'string' ? enc.encode(m.raw) : toBytes(m.raw ?? m);
    const from = m.from?.address || 'unknown@localhost';
    const when = m.date ? new Date(m.date) : new Date();
    chunks.push(enc.encode(`From ${from} ${when.toUTCString()}\n`));
    const text = new TextDecoder('latin1').decode(raw).replace(/^(>*From )/gm, '>$1');
    chunks.push(latin1Bytes(text));
    chunks.push(enc.encode('\n\n'));
  }
  const total = chunks.reduce((n, c) => n + c.length, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const c of chunks) {
    out.set(c, at);
    at += c.length;
  }
  return out;
}
