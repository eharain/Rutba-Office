// The node database layer of a .pst / .ost — [MS-PST] §2.2.2.
//
// A PST is a paged store with two B-trees over it. The node B-tree (NBT) maps a
// node id to the blocks holding its data; the block B-tree (BBT) maps a block
// id to a byte offset and a length. Everything above this file — folders,
// messages, attachments — is built out of nodes, so these two trees and the
// block reader are the whole foundation.
//
// Three details cause most of the difficulty:
//
//   - Data larger than one block is stored in a tree of blocks (XBLOCK, and
//     XXBLOCK above it), so "read this node" means "walk a tree and
//     concatenate", not "read one range".
//   - Data blocks are obfuscated; internal blocks are not. Which is which is
//     carried in a bit of the block id.
//   - Sub-nodes hang off a node with their own little index. A message's
//     recipients and attachments live there.
//
// The file arrives through a reader interface, so the same code serves a
// Buffer in a test and a 950 MB file on disk that must never be loaded whole.

import { createRequire } from 'node:module';
import { CRYPT, decodeBlock, supports, CRYPT_NAME } from './crypt.js';

export { CRYPT, CRYPT_NAME };

// Outlook 2013 and later compress the larger data blocks of an .ost with zlib.
// It is not in the published structure — the block simply begins with a zlib
// header where a heap header should be — and a reader that does not notice sees
// every message in the file as empty.
//
// Detection is safe rather than clever. A heap block's first two bytes are the
// offset of its allocation map, which cannot exceed the block size, so the high
// byte is always small; a zlib header's second byte never is. And inflate
// verifies an Adler-32 checksum over the whole stream, so a block that is not
// really compressed fails instead of decoding to something plausible.
let inflateSync = null;
try {
  inflateSync = createRequire(import.meta.url)('node:zlib').inflateSync;
} catch {
  inflateSync = null; // a browser build simply does not meet these files
}

function maybeInflate(bytes) {
  if (!inflateSync || bytes.length < 6) return bytes;
  const cmf = bytes[0];
  const flg = bytes[1];
  // zlib: low nibble 8 (deflate), and the two header bytes are a multiple of 31
  if ((cmf & 0x0f) !== 8 || ((cmf << 8) | flg) % 31 !== 0) return bytes;
  try {
    const out = inflateSync(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength));
    return new Uint8Array(out.buffer, out.byteOffset, out.byteLength);
  } catch {
    return bytes;
  }
}

export const NID_TYPE = {
  HID: 0x00,
  INTERNAL: 0x01,
  NORMAL_FOLDER: 0x02,
  SEARCH_FOLDER: 0x03,
  NORMAL_MESSAGE: 0x04,
  ATTACHMENT: 0x05,
  SEARCH_UPDATE_QUEUE: 0x06,
  SEARCH_CRITERIA_OBJECT: 0x07,
  ASSOC_MESSAGE: 0x08,
  CONTENTS_TABLE_INDEX: 0x0a,
  RECEIVE_FOLDER_TABLE: 0x0b,
  OUTGOING_QUEUE_TABLE: 0x0c,
  HIERARCHY_TABLE: 0x0d,
  CONTENTS_TABLE: 0x0e,
  ASSOC_CONTENTS_TABLE: 0x0f,
  SEARCH_CONTENTS_TABLE: 0x10,
  ATTACHMENT_TABLE: 0x11,
  RECIPIENT_TABLE: 0x12,
  SEARCH_TABLE_INDEX: 0x13,
  LTP: 0x1f,
};

export const NID = {
  MESSAGE_STORE: 0x21,
  NAME_TO_ID_MAP: 0x61,
  ROOT_FOLDER: 0x122,
};

export const nidType = (nid) => nid & 0x1f;
export const nidIndex = (nid) => nid >>> 5;
export const makeNid = (index, type) => ((index << 5) | type) >>> 0;
/** A folder's tables share its index and differ only in type. */
export const siblingNid = (nid, type) => (((nid & ~0x1f) | type) >>> 0);

export class NdbError extends Error {
  constructor(message, code) {
    super(message);
    this.name = 'NdbError';
    this.code = code;
  }
}

/** A 64-bit little-endian read kept as a Number — PST offsets fit in 2^53. */
function readU64(dv, at) {
  return dv.getUint32(at + 4, true) * 4294967296 + dv.getUint32(at, true);
}

/** Read a whole file into memory. Fine for a .msg, wrong for a big .ost. */
export function bufferReader(bytes) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  return {
    size: b.length,
    read: (offset, length) => b.subarray(offset, Math.min(offset + length, b.length)),
    close() {},
  };
}

export class Ndb {
  /** @param {{ size: number, read: (offset: number, length: number) => Uint8Array }} reader */
  constructor(reader) {
    this.reader = reader;
    this.#readHeader();
  }

  #view(offset, length) {
    const bytes = this.reader.read(offset, length);
    return { bytes, dv: new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength) };
  }

  #readHeader() {
    const { bytes, dv } = this.#view(0, 1024);
    if (!(bytes[0] === 0x21 && bytes[1] === 0x42 && bytes[2] === 0x44 && bytes[3] === 0x4e)) {
      throw new NdbError('not an Outlook data file: the !BDN signature is missing', 'not-pst');
    }
    this.magicClient = dv.getUint16(8, true);
    this.version = dv.getUint16(10, true);
    this.isOst = this.magicClient === 0x4f53;
    this.unicode = this.version >= 23;
    this.fourK = this.version >= 36;
    this.pageSize = this.fourK ? 4096 : 512;

    if (!this.unicode && this.version !== 14 && this.version !== 15) {
      throw new NdbError(`unsupported Outlook file version ${this.version}`, 'version');
    }

    const rootAt = this.unicode ? 0xb4 : 0xa4;
    if (this.unicode) {
      this.fileEof = readU64(dv, rootAt + 4);
      this.nbtRoot = { bid: readU64(dv, rootAt + 36), ib: readU64(dv, rootAt + 44) };
      this.bbtRoot = { bid: readU64(dv, rootAt + 52), ib: readU64(dv, rootAt + 60) };
      this.cryptMethod = bytes[0x201];
    } else {
      this.fileEof = dv.getUint32(rootAt + 4, true);
      this.nbtRoot = { bid: dv.getUint32(rootAt + 20, true), ib: dv.getUint32(rootAt + 24, true) };
      this.bbtRoot = { bid: dv.getUint32(rootAt + 28, true), ib: dv.getUint32(rootAt + 32, true) };
      this.cryptMethod = bytes[0x1cd];
    }
    this.entrySize = this.unicode ? 24 : 12;
  }

  /** True when this build holds the tables needed to read this file's blocks. */
  get readable() {
    return supports(this.cryptMethod);
  }

  assertReadable() {
    if (this.readable) return;
    throw new NdbError(
      `This file uses Outlook's ${CRYPT_NAME[this.cryptMethod] || this.cryptMethod} encoding, ` +
        'which this build cannot decode.',
      'crypt-unsupported'
    );
  }

  /**
   * The counters at the end of a B-tree page. Their offsets differ between the
   * 512-byte format and the 4 KB format introduced with Outlook 2013, so both
   * are measured back from the end of the page rather than forward from its
   * start — the trailer is what is actually fixed.
   */
  #pageCounters(bytes, dv) {
    if (this.fourK) {
      return {
        cEnt: dv.getUint16(this.pageSize - 40, true),
        cEntMax: dv.getUint16(this.pageSize - 38, true),
        cbEnt: bytes[this.pageSize - 36],
        cLevel: bytes[this.pageSize - 35],
      };
    }
    if (this.unicode) {
      return {
        cEnt: bytes[this.pageSize - 24],
        cEntMax: bytes[this.pageSize - 23],
        cbEnt: bytes[this.pageSize - 22],
        cLevel: bytes[this.pageSize - 21],
      };
    }
    return {
      cEnt: bytes[this.pageSize - 16],
      cEntMax: bytes[this.pageSize - 15],
      cbEnt: bytes[this.pageSize - 14],
      cLevel: bytes[this.pageSize - 13],
    };
  }

  page(ib) {
    const { bytes, dv } = this.#view(ib, this.pageSize);
    return { bytes, dv, ...this.#pageCounters(bytes, dv) };
  }

  /** Walk a B-tree depth-first, handing each leaf entry to `onLeaf`. */
  #walk(ib, onLeaf, seen = new Set(), depth = 0) {
    if (depth > 40 || seen.has(ib) || ib <= 0 || ib >= this.reader.size) return;
    seen.add(ib);
    const page = this.page(ib);
    const size = page.cbEnt || this.entrySize;
    if (!page.cEnt || size < 4 || page.cEnt * size > this.pageSize) return;

    for (let i = 0; i < page.cEnt; i++) {
      const at = i * size;
      if (at + size > this.pageSize) break;
      if (page.cLevel > 0) {
        // BTENTRY: the key, then a BREF naming the child page.
        const childIb = this.unicode ? readU64(page.dv, at + 16) : page.dv.getUint32(at + 8, true);
        this.#walk(childIb, onLeaf, seen, depth + 1);
      } else {
        onLeaf(page.dv, at, page.bytes);
      }
    }
  }

  /** block id → { ib, cb }. Built once; a PST has one of these, not one per read. */
  blockIndex() {
    if (this._bbt) return this._bbt;
    const map = new Map();
    this.#walk(this.bbtRoot.ib, (dv, at) => {
      if (this.unicode) {
        const bid = readU64(dv, at);
        map.set(bid, { bid, ib: readU64(dv, at + 8), cb: dv.getUint16(at + 16, true) });
      } else {
        const bid = dv.getUint32(at, true);
        map.set(bid, { bid, ib: dv.getUint32(at + 4, true), cb: dv.getUint16(at + 8, true) });
      }
    });
    this._bbt = map;
    return map;
  }

  /** node id → { dataBid, subBid, parentNid }. */
  nodeIndex() {
    if (this._nbt) return this._nbt;
    const map = new Map();
    this.#walk(this.nbtRoot.ib, (dv, at) => {
      const nid = dv.getUint32(at, true);
      if (this.unicode) {
        map.set(nid, {
          nid,
          dataBid: readU64(dv, at + 8),
          subBid: readU64(dv, at + 16),
          parentNid: dv.getUint32(at + 24, true),
        });
      } else {
        map.set(nid, {
          nid,
          dataBid: dv.getUint32(at + 4, true),
          subBid: dv.getUint32(at + 8, true),
          parentNid: dv.getUint32(at + 12, true),
        });
      }
    });
    this._nbt = map;
    return map;
  }

  /** Bytes of one block, decoded. Internal blocks are never obfuscated. */
  #rawBlock(bid) {
    const entry = this.blockIndex().get(bid);
    if (!entry) return null;
    const { bytes } = this.#view(entry.ib, entry.cb);
    const internal = (bid & 0x02) !== 0;
    if (internal) return bytes;
    this.assertReadable();
    return maybeInflate(decodeBlock(bytes, this.cryptMethod, bid));
  }

  /**
   * A node's data, following XBLOCK/XXBLOCK trees so anything over one block
   * arrives as one contiguous run of bytes.
   */
  blockData(bid, depth = 0) {
    if (!bid || depth > 8) return new Uint8Array(0);
    const raw = this.#rawBlock(bid);
    if (!raw) return new Uint8Array(0);
    if ((bid & 0x02) === 0) return raw;

    const dv = new DataView(raw.buffer, raw.byteOffset, raw.byteLength);
    if (raw[0] !== 0x01) return raw; // a sub-node block, not a data tree
    const cEnt = dv.getUint16(2, true);
    const idSize = this.unicode ? 8 : 4;
    const parts = [];
    for (let i = 0; i < cEnt; i++) {
      const at = 8 + i * idSize;
      if (at + idSize > raw.length) break;
      parts.push(this.blockData(this.unicode ? readU64(dv, at) : dv.getUint32(at, true), depth + 1));
    }
    const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    let off = 0;
    for (const p of parts) {
      out.set(p, off);
      off += p.length;
    }
    return out;
  }

  /**
   * A node's data as the list of blocks it is stored in, rather than one run.
   *
   * The heap-on-node layer needs this: a heap id names a block and an index
   * within it, so concatenating first and splitting later would put every item
   * after the first block at the wrong offset.
   */
  blockParts(bid, out = [], depth = 0) {
    if (!bid || depth > 8) return out;
    const raw = this.#rawBlock(bid);
    if (!raw) return out;
    if ((bid & 0x02) === 0) {
      out.push(raw);
      return out;
    }
    const dv = new DataView(raw.buffer, raw.byteOffset, raw.byteLength);
    if (raw[0] !== 0x01) {
      out.push(raw);
      return out;
    }
    const cEnt = dv.getUint16(2, true);
    const idSize = this.unicode ? 8 : 4;
    for (let i = 0; i < cEnt; i++) {
      const at = 8 + i * idSize;
      if (at + idSize > raw.length) break;
      this.blockParts(this.unicode ? readU64(dv, at) : dv.getUint32(at, true), out, depth + 1);
    }
    return out;
  }

  /** The sub-node index hanging off a node. */
  subnodes(subBid, depth = 0) {
    const map = new Map();
    if (!subBid || depth > 8) return map;
    const raw = this.#rawBlock(subBid);
    if (!raw || raw.length < 8 || raw[0] !== 0x02) return map;
    const dv = new DataView(raw.buffer, raw.byteOffset, raw.byteLength);
    const cLevel = raw[1];
    const cEnt = dv.getUint16(2, true);

    if (cLevel === 0) {
      const size = this.unicode ? 24 : 12;
      for (let i = 0; i < cEnt; i++) {
        const at = 8 + i * size;
        if (at + size > raw.length) break;
        const nid = dv.getUint32(at, true);
        map.set(nid, {
          nid,
          dataBid: this.unicode ? readU64(dv, at + 8) : dv.getUint32(at + 4, true),
          subBid: this.unicode ? readU64(dv, at + 16) : dv.getUint32(at + 8, true),
        });
      }
      return map;
    }

    const size = this.unicode ? 16 : 8;
    for (let i = 0; i < cEnt; i++) {
      const at = 8 + i * size;
      if (at + size > raw.length) break;
      const child = this.unicode ? readU64(dv, at + 8) : dv.getUint32(at + 4, true);
      for (const [k, v] of this.subnodes(child, depth + 1)) map.set(k, v);
    }
    return map;
  }

  /** Everything a node holds. */
  node(nid) {
    const entry = this.nodeIndex().get(nid >>> 0);
    if (!entry) return null;
    return {
      nid: entry.nid,
      parentNid: entry.parentNid,
      data: this.blockData(entry.dataBid),
      parts: this.blockParts(entry.dataBid),
      subnodes: this.subnodes(entry.subBid),
    };
  }

  /** A sub-node addresses blocks exactly as a node does. */
  subnode(sub) {
    if (!sub) return null;
    return {
      nid: sub.nid,
      data: this.blockData(sub.dataBid),
      parts: this.blockParts(sub.dataBid),
      subnodes: this.subnodes(sub.subBid),
    };
  }

  /** Node ids of one type, e.g. every message. */
  nodesOfType(type) {
    const out = [];
    for (const nid of this.nodeIndex().keys()) if (nidType(nid) === type) out.push(nid);
    return out;
  }

  stats() {
    return {
      version: this.version,
      unicode: this.unicode,
      fourK: this.fourK,
      pageSize: this.pageSize,
      isOst: this.isOst,
      crypt: CRYPT_NAME[this.cryptMethod] ?? this.cryptMethod,
      readable: this.readable,
      nodes: this.nodeIndex().size,
      blocks: this.blockIndex().size,
      fileEof: this.fileEof,
    };
  }
}

export { readU64 };
