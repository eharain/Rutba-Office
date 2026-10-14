// A zip read from its file a piece at a time.
//
// An Outlook for Mac archive is a zip, and one of a few years' mail is
// several gigabytes: more than a Node buffer can hold (2 GB), and far more
// than a mail program should take into memory to show a folder. So the zip
// is not read whole. Its central directory - a few bytes a message - is read
// from the end of the file, and an entry's bytes are read from where the
// directory says they are, when they are asked for, and let go after.
// Zip64 records, which an archive past 4 GB or 65,535 entries needs, are
// read as well. Nothing is ever written.

import fs from 'node:fs';
import zlib from 'node:zlib';

const EOCD = 0x06054b50;
const EOCD64 = 0x06064b50;
const EOCD64_LOCATOR = 0x07064b50;
const CENTRAL = 0x02014b50;
const LOCAL = 0x04034b50;

/** Does the file begin as a zip does. */
export function looksLikeZip(head) {
  return head.length >= 4 && head[0] === 0x50 && head[1] === 0x4b && ((head[2] === 3 && head[3] === 4) || (head[2] === 5 && head[3] === 6));
}

export class ZipFile {
  /** @param {string} filePath */
  constructor(filePath) {
    this.path = filePath;
    this.fd = fs.openSync(filePath, 'r');
    try {
      this.size = fs.fstatSync(this.fd).size;
      this.entries = this.#directory();
    } catch (err) {
      fs.closeSync(this.fd);
      this.fd = null;
      throw err;
    }
  }

  /** One buffer read into again and again (`data(entry, { reuse: true })`). */
  #scratch = null;

  /** `n` bytes at `at`, all of them, or a sentence saying the file is cut short. */
  #read(at, n, what, reuse = false) {
    if (!Number.isSafeInteger(at) || !Number.isSafeInteger(n) || at < 0 || n < 0 || at + n > this.size) {
      throw new Error(`not a complete zip archive: ${what} points past the end of the file`);
    }
    let buf;
    if (reuse) {
      if (!this.#scratch || this.#scratch.length < n) this.#scratch = Buffer.allocUnsafe(Math.max(n, 1048576));
      buf = this.#scratch.subarray(0, n);
    } else buf = Buffer.allocUnsafe(n);
    let got = 0;
    while (got < n) {
      const r = fs.readSync(this.fd, buf, got, n - got, at + got);
      if (!r) throw new Error(`not a complete zip archive: the file ends inside ${what}`);
      got += r;
    }
    return buf;
  }

  #directory() {
    const tailLength = Math.min(this.size, 22 + 0xffff + 20);
    const tailAt = this.size - tailLength;
    const tail = this.#read(tailAt, tailLength, 'the end of the archive');
    let e = -1;
    for (let i = tail.length - 22; i >= 0; i--) {
      if (tail.readUInt32LE(i) === EOCD) {
        e = i;
        break;
      }
    }
    if (e < 0) throw new Error('not a zip archive: no end-of-central-directory record');
    let count = tail.readUInt16LE(e + 10);
    let cdSize = tail.readUInt32LE(e + 12);
    let cdOffset = tail.readUInt32LE(e + 16);
    if (count === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) {
      // The zip64 locator sits just before the end record and says where the zip64 end record is.
      const l = e - 20;
      if (l >= 0 && tail.readUInt32LE(l) === EOCD64_LOCATOR) {
        const at = Number(tail.readBigUInt64LE(l + 8));
        const rec = this.#read(at, 56, 'the zip64 directory');
        if (rec.readUInt32LE(0) !== EOCD64) throw new Error('malformed zip64 end-of-central-directory');
        count = Number(rec.readBigUInt64LE(32));
        cdSize = Number(rec.readBigUInt64LE(40));
        cdOffset = Number(rec.readBigUInt64LE(48));
      }
    }
    const cd = this.#read(cdOffset, cdSize, 'the central directory');
    const entries = [];
    let p = 0;
    for (let i = 0; i < count; i++) {
      if (p + 46 > cd.length || cd.readUInt32LE(p) !== CENTRAL) throw new Error(`malformed central directory at entry ${i + 1}`);
      const flags = cd.readUInt16LE(p + 8);
      const method = cd.readUInt16LE(p + 10);
      const crc = cd.readUInt32LE(p + 16);
      let compressedSize = cd.readUInt32LE(p + 20);
      let size = cd.readUInt32LE(p + 24);
      const nameLength = cd.readUInt16LE(p + 28);
      const extraLength = cd.readUInt16LE(p + 30);
      const commentLength = cd.readUInt16LE(p + 32);
      let localOffset = cd.readUInt32LE(p + 42);
      if (p + 46 + nameLength + extraLength + commentLength > cd.length) throw new Error(`malformed central directory at entry ${i + 1}`);
      // Bit 11: the name is UTF-8; otherwise it is the old DOS code page, which for the names an archive writes is ASCII.
      const name = cd.toString(flags & 0x800 ? 'utf8' : 'latin1', p + 46, p + 46 + nameLength);
      if (size === 0xffffffff || compressedSize === 0xffffffff || localOffset === 0xffffffff) {
        const extra = cd.subarray(p + 46 + nameLength, p + 46 + nameLength + extraLength);
        for (let x = 0; x + 4 <= extra.length; ) {
          const id = extra.readUInt16LE(x);
          const len = extra.readUInt16LE(x + 2);
          if (id === 0x0001) {
            let o = x + 4;
            const take = () => {
              const v = o + 8 <= x + 4 + len ? Number(extra.readBigUInt64LE(o)) : null;
              o += 8;
              return v;
            };
            if (size === 0xffffffff) size = take() ?? size;
            if (compressedSize === 0xffffffff) compressedSize = take() ?? compressedSize;
            if (localOffset === 0xffffffff) localOffset = take() ?? localOffset;
            break;
          }
          x += 4 + len;
        }
      }
      entries.push({ name, method, crc, compressedSize, size, localOffset });
      p += 46 + nameLength + extraLength + commentLength;
    }
    return entries;
  }

  /**
   * An entry's bytes, read from the file now and not kept. `reuse`: into a
   * buffer the next such read writes over, for bytes used at once and let go.
   */
  data(entry, { reuse = false } = {}) {
    const head = this.#read(entry.localOffset, 30, `the entry ${entry.name}`);
    if (head.readUInt32LE(0) !== LOCAL) throw new Error(`malformed local header for ${entry.name}`);
    const start = entry.localOffset + 30 + head.readUInt16LE(26) + head.readUInt16LE(28);
    const compressed = this.#read(start, entry.compressedSize, `the entry ${entry.name}`, reuse);
    if (entry.method === 0) return compressed;
    if (entry.method === 8) return zlib.inflateRawSync(compressed);
    throw new Error(`${entry.name} is compressed in a way this cannot read (method ${entry.method}).`);
  }

  close() {
    if (this.fd != null) fs.closeSync(this.fd);
    this.fd = null;
  }
}
