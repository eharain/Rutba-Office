/**
 * ZIP layer, written for PRESERVATION rather than for convenience.
 *
 * Fidelity starts here and most implementations lose it here. A library that
 * decompresses everything on read and recompresses everything on write has
 * already changed every byte of a file it was only asked to edit one cell of.
 * Then a customer opens it in Excel, something subtle is different, and nobody
 * can tell you which layer did it.
 *
 * So every entry keeps BOTH forms:
 *   - `stored`  the original compressed bytes, method, CRC and sizes
 *   - `raw`     lazily decompressed, only if something asks for it
 *
 * An entry nobody modified is written back from `stored`, verbatim. That is what
 * makes "we preserve what we do not understand" a guarantee rather than a hope:
 * charts, pivot caches, conditional formatting, VML, embedded objects and OLE
 * survive because their bytes are never regenerated.
 *
 * Entry ORDER is preserved too. It carries no semantics in the spec, but some
 * consumers are order-sensitive in practice, and preserving it costs nothing.
 */
import zlib from 'node:zlib';

const LOCAL_SIG = 0x04034b50;
const CENTRAL_SIG = 0x02014b50;
const EOCD_SIG = 0x06054b50;
const EOCD64_LOCATOR_SIG = 0x07064b50;
const EOCD64_SIG = 0x06064b50;

/**
 * Table-based CRC-32 in plain JS. This used to live inside the feature test
 * below, constructed only on Node versions lacking `zlib.crc32` — which made it
 * unreachable for testing on versions that have it. It is now built
 * unconditionally (256 words, once per process) and exported, because it is the
 * implementation a browser runs and its answers must be provably identical to
 * zlib's.
 */
const crc32Js = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c >>> 0;
  }
  return (buf) => {
    let c = 0xffffffff;
    for (let i = 0; i < buf.length; i++) c = table[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
})();

const crc32 = typeof zlib.crc32 === 'function' ? (buf) => zlib.crc32(buf) >>> 0 : crc32Js;

export class ZipEntry {
  constructor({ name, method, crc, compressedSize, uncompressedSize, compressed, dosTime, dosDate, flags, externalAttrs, comment }) {
    this.name = name;
    this.method = method;
    this.crc = crc;
    this.compressedSize = compressedSize;
    this.uncompressedSize = uncompressedSize;
    this.dosTime = dosTime;
    this.dosDate = dosDate;
    this.flags = flags;
    this.externalAttrs = externalAttrs;
    this.comment = comment;
    /** original compressed bytes — the thing we write back when untouched */
    this.compressed = compressed;
    this._raw = null;
    this.modified = false;
  }

  /** Decompressed bytes. Decompressing does NOT mark the entry modified. */
  get data() {
    if (this._raw === null) {
      this._raw = this.method === 8 ? zlib.inflateRawSync(this.compressed) : Buffer.from(this.compressed);
    }
    return this._raw;
  }

  /** Replacing the bytes is the only thing that marks an entry modified. */
  set data(buf) {
    const bytes = Buffer.isBuffer(buf) ? buf : Buffer.from(buf);
    this._raw = bytes;
    this.modified = true;
    this.uncompressedSize = bytes.length;
    this.crc = crc32(bytes);
  }

  /** Bytes to write: the untouched original, or a fresh compression. */
  serialize() {
    if (!this.modified) {
      return { method: this.method, payload: this.compressed, crc: this.crc, uncompressedSize: this.uncompressedSize };
    }
    // Recompress with the method the entry already used, so an edited part does
    // not silently change storage strategy relative to its siblings.
    if (this.method === 0) {
      return { method: 0, payload: this._raw, crc: this.crc, uncompressedSize: this._raw.length };
    }
    const deflated = zlib.deflateRawSync(this._raw, { level: 9 });
    return { method: 8, payload: deflated, crc: this.crc, uncompressedSize: this._raw.length };
  }
}

/**
 * Inflate raw-deflate bytes WITHOUT `node:zlib`, via the WHATWG
 * `DecompressionStream('deflate-raw')` — a global in every modern browser and
 * in Node >= 18. Raw deflate has exactly one decoding, so the output is
 * byte-for-byte what `zlib.inflateRawSync` produces; only the shape of the API
 * differs (async, Uint8Array chunks).
 */
export async function inflateRawStream(compressed) {
  const ds = new DecompressionStream('deflate-raw');
  // Copy before writing: entry payloads are subarrays of the whole archive
  // buffer, and a stream implementation is allowed to detach what it is handed.
  const chunk = new Uint8Array(compressed.byteLength);
  chunk.set(compressed);
  const writer = ds.writable.getWriter();
  const writing = writer.write(chunk).then(() => writer.close());
  writing.catch(() => {}); // corrupt input surfaces via the read loop, not as an unhandled rejection
  const reader = ds.readable.getReader();
  const chunks = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.length;
  }
  await writing;
  if (chunks.length === 1) return chunks[0];
  const out = new Uint8Array(total);
  let offset = 0;
  for (const c of chunks) { out.set(c, offset); offset += c.length; }
  return out;
}

const hasZlibInflate = typeof zlib?.inflateRawSync === 'function';

/**
 * Downstream code calls `.toString('utf8')` and friends on entry data, so on a
 * host that has Buffer the stream inflater's output is re-viewed as one
 * (zero-copy). On a host without Buffer the Uint8Array passes through — those
 * call sites are the Buffer audit's problem, not this function's.
 */
const asHostBytes = (u8) => (typeof Buffer === 'undefined' ? u8 : Buffer.from(u8.buffer, u8.byteOffset, u8.byteLength));

/**
 * Pre-inflate every deflated entry so that everything DOWNSTREAM of an async
 * read stays synchronous and untouched: once `_raw` is populated, the lazy
 * `data` getter never reaches for zlib. Populating the cache is a read, not an
 * edit — `modified` stays false and untouched entries still round-trip from
 * their original compressed bytes.
 *
 * @param {ZipEntry[]} entries
 * @param {{inflate?: 'stream'}} [options] internal — `'stream'` forces the
 *   DecompressionStream path even where zlib exists, which is how tests
 *   exercise the browser path on Node. Default: zlib when present (fast,
 *   byte-identical), the stream otherwise.
 */
export async function inflateEntries(entries, { inflate } = {}) {
  const useStream = inflate === 'stream' || !hasZlibInflate;
  await Promise.all(entries.map(async (entry) => {
    if (entry.method !== 8 || entry._raw !== null) return;
    entry._raw = useStream
      ? asHostBytes(await inflateRawStream(entry.compressed))
      : zlib.inflateRawSync(entry.compressed);
  }));
}

/** Locate the end-of-central-directory, tolerating a trailing comment. */
function findEocd(buf) {
  const min = Math.max(0, buf.length - 22 - 0xffff);
  for (let i = buf.length - 22; i >= min; i--) {
    if (buf.readUInt32LE(i) === EOCD_SIG) return i;
  }
  throw new Error('not a zip archive: no end-of-central-directory record');
}

/**
 * @param {Buffer} buf
 * @returns {{entries: ZipEntry[], comment: Buffer}}
 */
export function readZip(buf) {
  const eocd = findEocd(buf);
  let count = buf.readUInt16LE(eocd + 10);
  let cdOffset = buf.readUInt32LE(eocd + 16);
  const commentLen = buf.readUInt16LE(eocd + 20);
  const comment = buf.subarray(eocd + 22, eocd + 22 + commentLen);

  // ZIP64: a real-world large workbook can exceed the 32-bit fields.
  if (count === 0xffff || cdOffset === 0xffffffff) {
    for (let i = eocd - 20; i >= 0; i--) {
      if (buf.readUInt32LE(i) === EOCD64_LOCATOR_SIG) {
        const z64 = Number(buf.readBigUInt64LE(i + 8));
        if (buf.readUInt32LE(z64) !== EOCD64_SIG) throw new Error('malformed zip64 end-of-central-directory');
        count = Number(buf.readBigUInt64LE(z64 + 32));
        cdOffset = Number(buf.readBigUInt64LE(z64 + 48));
        break;
      }
    }
  }

  const entries = [];
  let p = cdOffset;
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(p) !== CENTRAL_SIG) throw new Error('malformed central directory at entry ' + i);
    const flags = buf.readUInt16LE(p + 8);
    const method = buf.readUInt16LE(p + 10);
    const dosTime = buf.readUInt16LE(p + 12);
    const dosDate = buf.readUInt16LE(p + 14);
    const crc = buf.readUInt32LE(p + 16);
    let compressedSize = buf.readUInt32LE(p + 20);
    let uncompressedSize = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLength = buf.readUInt16LE(p + 32);
    const externalAttrs = buf.readUInt32LE(p + 38);
    let localOffset = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen);
    const extra = buf.subarray(p + 46 + nameLen, p + 46 + nameLen + extraLen);

    // ZIP64 extended information extra field
    if (uncompressedSize === 0xffffffff || compressedSize === 0xffffffff || localOffset === 0xffffffff) {
      let e = 0;
      while (e + 4 <= extra.length) {
        const headerId = extra.readUInt16LE(e);
        const size = extra.readUInt16LE(e + 2);
        if (headerId === 0x0001) {
          let o = e + 4;
          if (uncompressedSize === 0xffffffff) { uncompressedSize = Number(extra.readBigUInt64LE(o)); o += 8; }
          if (compressedSize === 0xffffffff) { compressedSize = Number(extra.readBigUInt64LE(o)); o += 8; }
          if (localOffset === 0xffffffff) { localOffset = Number(extra.readBigUInt64LE(o)); o += 8; }
          break;
        }
        e += 4 + size;
      }
    }

    // The local header's name/extra lengths can differ from the central ones.
    if (buf.readUInt32LE(localOffset) !== LOCAL_SIG) throw new Error('malformed local header for ' + name);
    const lNameLen = buf.readUInt16LE(localOffset + 26);
    const lExtraLen = buf.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + lNameLen + lExtraLen;

    entries.push(new ZipEntry({
      name, method, crc, compressedSize, uncompressedSize, flags, dosTime, dosDate, externalAttrs,
      comment: buf.subarray(p + 46 + nameLen + extraLen, p + 46 + nameLen + extraLen + commentLength),
      compressed: buf.subarray(dataStart, dataStart + compressedSize),
    }));
    p += 46 + nameLen + extraLen + commentLength;
  }
  return { entries, comment };
}

/**
 * @param {ZipEntry[]} entries in the order they should appear
 * @param {Buffer} [archiveComment]
 */
export function writeZip(entries, archiveComment = Buffer.alloc(0)) {
  const locals = [];
  const centrals = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBuf = Buffer.from(entry.name, 'utf8');
    const { method, payload, crc, uncompressedSize } = entry.serialize();
    // Data descriptors would move the sizes after the payload; we always write
    // them in the header, so clear the flag rather than lie about it.
    const flags = (entry.flags ?? 0) & ~0x0008;

    const local = Buffer.alloc(30);
    local.writeUInt32LE(LOCAL_SIG, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(flags, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(entry.dosTime ?? 0, 10);
    local.writeUInt16LE(entry.dosDate ?? 0x2821, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(payload.length, 18);
    local.writeUInt32LE(uncompressedSize, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    locals.push(local, nameBuf, payload);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(CENTRAL_SIG, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(flags, 8);
    central.writeUInt16LE(method, 10);
    central.writeUInt16LE(entry.dosTime ?? 0, 12);
    central.writeUInt16LE(entry.dosDate ?? 0x2821, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(payload.length, 20);
    central.writeUInt32LE(uncompressedSize, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(entry.externalAttrs ?? 0, 38);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);

    offset += local.length + nameBuf.length + payload.length;
  }

  const centralBuf = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22 + archiveComment.length);
  eocd.writeUInt32LE(EOCD_SIG, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(archiveComment.length, 20);
  archiveComment.copy(eocd, 22);
  return Buffer.concat([...locals, centralBuf, eocd]);
}

export { crc32, crc32Js };
