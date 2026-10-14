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

import fs from 'node:fs';
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
    ranges.push({ start: bodyStart, end: to, separator: latin1Text(bytes.subarray(from, bodyStart - 1)) });
  }
  return { bytes, ranges };
}

/** Undo `>From ` quoting, and any run of `>` before it (mboxrd). */
function unescape(bytes) {
  const text = latin1Text(bytes);
  if (!/^>+From /m.test(text)) return bytes;
  return latin1Bytes(text.replace(/^(>+)(From )/gm, (m, gt, rest) => gt.slice(1) + rest));
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

  /** @param {{ light?: boolean }} [opts] see parseMessage */
  message(index, opts) {
    const raw = this.raw(index);
    return raw ? parseMessage(raw, opts) : null;
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
 * are whatever encoding the sender used. Reading each byte as one character
 * and writing each character back as one byte returns exactly the bytes that
 * came in. Going out through UTF-8 instead would turn every byte above 127
 * into two, which is how an em-dash becomes three mojibake characters. The
 * pair is written by hand rather than with `TextDecoder('latin1')`, which is
 * Windows-1252 by the encoding standard: it reads the bytes 0x80–0x9F as
 * the euro sign, the curly quotes and their neighbours, whose code points
 * do not come back to the byte they were — an em-dash in a UTF-8 body
 * (E2 80 94) went out as E2 AC 1D.
 */
function latin1Text(bytes) {
  let out = '';
  for (let i = 0; i < bytes.length; i += 8192) out += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
  return out;
}

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
    const text = latin1Text(raw).replace(/^(>*From )/gm, '>$1');
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

/**
 * An mbox read from its file, never whole.
 *
 * A Thunderbird or Gmail export of a few years' mail is one file of several
 * gigabytes: past what a Node buffer holds (2 GB), and far past what a mail
 * program should take into memory to show a folder. So the file is read
 * once, in blocks of a few megabytes, only to find where each message
 * begins - a From_ line, at the start of a line - and those places are all
 * that is kept (eight bytes a message). A message is read from the file
 * when it is asked for. Nothing is ever written.
 */
export class MboxFile {
  /**
   * @param {string} filePath
   * @param {{ block?: number }} [opts] how much is read at a time while finding the messages
   */
  constructor(filePath, { block = 8 * 1048576 } = {}) {
    this.path = filePath;
    this.fd = fs.openSync(filePath, 'r');
    this.bytes = fs.fstatSync(this.fd).size;
    try {
      this.starts = findStarts(this.fd, this.bytes, block);
    } catch (err) {
      this.close();
      throw err;
    }
  }

  get length() {
    return this.starts.length;
  }

  /** Where message `index` lies in the file, its From_ line included. */
  range(index) {
    if (index < 0 || index >= this.starts.length) return null;
    return { start: this.starts[index], end: index + 1 < this.starts.length ? this.starts[index + 1] : this.bytes };
  }

  /** Its size in the file. */
  size(index) {
    const r = this.range(index);
    return r ? r.end - r.start : 0;
  }

  /** One buffer read into again and again, for messages read only to be indexed. */
  #scratch = null;

  #read(at, n, reuse = false) {
    let buf;
    if (reuse) {
      if (!this.#scratch || this.#scratch.length < n) this.#scratch = Buffer.allocUnsafe(Math.max(n, 1048576));
      buf = this.#scratch;
    } else buf = Buffer.allocUnsafe(n);
    let got = 0;
    while (got < n) {
      const r = fs.readSync(this.fd, buf, got, n - got, at + got);
      if (!r) break;
      got += r;
    }
    return new Uint8Array(buf.buffer, buf.byteOffset, got);
  }

  /**
   * The message exactly as the archive holds it, its From_ line left out and
   * its quoting undone. `reuse`: read into a buffer the next read writes
   * over (for a message parsed at once and let go, as indexing does).
   */
  raw(index, { reuse = false } = {}) {
    const r = this.range(index);
    if (!r) return null;
    const bytes = this.#read(r.start, r.end - r.start, reuse);
    let body = 0;
    while (body < bytes.length && bytes[body] !== 10) body++;
    return unescape(bytes.subarray(Math.min(bytes.length, body + 1)));
  }

  /**
   * @param {{ light?: boolean }} [opts] see parseMessage; a light message is
   *   read into the reused buffer, and keeps nothing of it
   */
  message(index, opts) {
    const raw = this.raw(index, { reuse: Boolean(opts?.light) });
    if (!raw) return null;
    const m = parseMessage(raw, opts);
    if (opts?.light) delete m.structure;
    return m;
  }

  /**
   * Its headers alone, parsed - who, when, about what - read from the first
   * part of the message rather than the whole of it, for a list of a large
   * archive drawn before its bodies are read.
   */
  headers(index, { most = 65536 } = {}) {
    const r = this.range(index);
    if (!r) return null;
    const bytes = this.#read(r.start, Math.min(most, r.end - r.start));
    let at = 0;
    while (at < bytes.length && bytes[at] !== 10) at++;
    let end = bytes.length;
    for (let i = at + 1; i < bytes.length - 1; i++) {
      if (bytes[i] === 10 && (bytes[i + 1] === 10 || (bytes[i + 1] === 13 && bytes[i + 2] === 10))) {
        end = i + 1;
        break;
      }
    }
    const head = parseMessage(unescape(bytes.subarray(Math.min(bytes.length, at + 1), end)));
    const type = head.headers.find((x) => x.key === 'content-type')?.value || '';
    return { ...head, hasAttachments: /multipart\/(mixed|related)/i.test(type) };
  }

  *[Symbol.iterator]() {
    for (let i = 0; i < this.length; i++) yield this.message(i);
  }

  close() {
    if (this.fd != null) fs.closeSync(this.fd);
    this.fd = null;
  }
}

const FROM_LINE = Buffer.from('\nFrom ');

/** Where every message begins: a From_ line at the start of the file or of a line. */
function findStarts(fd, size, block) {
  let starts = new Float64Array(1024);
  let n = 0;
  const push = (v) => {
    if (n === starts.length) {
      const more = new Float64Array(n * 2);
      more.set(starts);
      starts = more;
    }
    starts[n++] = v;
  };
  // Each block is read with a few bytes of the next, so a From_ line across
  // the join is found once, in the block its newline is in.
  const buf = Buffer.allocUnsafe(block + FROM_LINE.length);
  for (let off = 0; off < size; off += block) {
    const want = Math.min(block + FROM_LINE.length, size - off);
    let got = 0;
    while (got < want) {
      const r = fs.readSync(fd, buf, got, want - got, off + got);
      if (!r) break;
      got += r;
    }
    const view = buf.subarray(0, got);
    if (off === 0 && got >= 5 && view.toString('latin1', 0, 5) === 'From ') push(0);
    const limit = Math.min(block, got);
    for (let j = view.indexOf(FROM_LINE); j >= 0 && j < limit; j = view.indexOf(FROM_LINE, j + 1)) push(off + j + 1);
  }
  return starts.slice(0, n);
}

/**
 * Open an mbox where it lies: from its file, read a message at a time.
 * @param {string} filePath
 */
export function openMbox(filePath, opts) {
  return new MboxFile(filePath, opts);
}
