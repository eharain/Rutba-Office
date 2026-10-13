// An Outlook data file written from scratch, for tests and window checks.
//
// The suite reads .pst and .ost files and has never needed to write one; the
// archive search, its performance budget and the legal-hold export do need
// .pst files to read, of any size, made on the spot rather than checked in.
// So this writes the Unicode format of [MS-PST] (version 23, 512-byte pages):
// the header, data blocks with their trailers, data trees (XBLOCK and XXBLOCK)
// for anything over one block, sub-node blocks, heaps on node with property
// contexts and table contexts, a folder's hierarchy and contents tables, a
// message's recipient table and its attachments, and the two B-trees.
//
// It writes as it goes, so a file of several gigabytes needs only its
// B-tree entries in memory. It does not write the allocation maps (AMap,
// PMap) or the name-to-id map Outlook keeps for named properties: the
// suite's reader does not use them, and nothing here claims Outlook would
// open the result. Blocks are encoded with the permute cipher by default, as
// Outlook's own "compressible" files are, so a reader test pays the decode.

import fs from 'node:fs';
import { mpbbR } from '../../packages/mailbox/src/pst/tables.js';

const PAGE = 512;
const BLOCK_MAX = 8176; // 8192 less the 16-byte trailer
const HEAP_ITEM_MAX = 3580;
const NID_TYPE = { FOLDER: 0x02, MESSAGE: 0x04, ATTACHMENT: 0x05, HIERARCHY: 0x0d, CONTENTS: 0x0e, ASSOC: 0x0f, LTP: 0x1f };
const PT = { I2: 0x02, LONG: 0x03, BOOLEAN: 0x0b, I8: 0x14, UNICODE: 0x1f, SYSTIME: 0x40, BINARY: 0x102 };
const NID_RECIPIENTS = 0x692;
const NID_ATTACHMENTS = 0x671;

const makeNid = (index, type) => ((index << 5) | type) >>> 0;

/* ── CRC and signatures, [MS-PST] §5.3 and §5.5 ─────────────────────────── */

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc(bytes, start = 0, end = bytes.length) {
  let c = 0;
  for (let i = start; i < end; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return c >>> 0;
}

function sig(ib, bid) {
  const x = ((ib % 4294967296) ^ (bid % 4294967296)) >>> 0;
  return ((x >>> 16) ^ x) & 0xffff;
}

const u64 = (buf, at, value) => {
  buf.writeUInt32LE(value % 4294967296 >>> 0, at);
  buf.writeUInt32LE(Math.floor(value / 4294967296) >>> 0, at + 4);
};

/* ── property values ────────────────────────────────────────────────────── */

const utf16 = (s) => Buffer.from(String(s ?? ''), 'utf16le');

function filetime(iso) {
  const ms = new Date(iso).getTime();
  const ticks = BigInt(Math.round(ms)) * 10000n + 116444736000000000n;
  const b = Buffer.alloc(8);
  b.writeBigUInt64LE(ticks);
  return b;
}

/** The bytes of a value of a type stored outside the property record. */
function valueBytes(type, value) {
  if (type === PT.UNICODE) return utf16(value);
  if (type === PT.SYSTIME) return filetime(value);
  if (type === PT.I8) {
    const b = Buffer.alloc(8);
    b.writeBigInt64LE(BigInt(value));
    return b;
  }
  return Buffer.isBuffer(value) ? value : Buffer.from(value);
}

const INLINE = new Set([PT.I2, PT.LONG, PT.BOOLEAN]);

/* ── heap on node, [MS-PST] §2.3.1 ──────────────────────────────────────── */

/**
 * A heap spread over as many blocks as it needs. Item ids name the block
 * and the index in that block's allocation map, as the format does.
 */
class HeapBuilder {
  constructor(clientSig) {
    this.clientSig = clientSig;
    this.blocks = [{ items: [], used: 12 }]; // the first block carries the 12-byte heap header
    this.userRoot = 0;
  }

  alloc(bytes) {
    const data = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
    if (data.length > HEAP_ITEM_MAX) throw new Error(`a heap item of ${data.length} bytes is over the format's limit`);
    let block = this.blocks[this.blocks.length - 1];
    const mapAfter = (b, extra) => 4 + 2 * (b.items.length + extra + 1);
    if (block.used + data.length + mapAfter(block, 1) + 1 > BLOCK_MAX) {
      const index = this.blocks.length;
      // Block 8, and every 128th after it, carries a fill-level bitmap.
      const header = index === 8 || (index > 8 && (index - 8) % 128 === 0) ? 66 : 2;
      block = { items: [], used: header };
      this.blocks.push(block);
    }
    block.items.push(data);
    block.used += data.length;
    const blockIndex = this.blocks.length - 1;
    return ((blockIndex << 16) | (block.items.length << 5)) >>> 0;
  }

  /** How many bytes the heap will take if one more item of `n` bytes is added. */
  fits(n) {
    const block = this.blocks[this.blocks.length - 1];
    return this.blocks.length === 1 && block.used + n + 4 + 2 * (block.items.length + 2) + 1 <= BLOCK_MAX;
  }

  build() {
    return this.blocks.map((block, index) => {
      const headerSize = index === 0 ? 12 : index === 8 || (index > 8 && (index - 8) % 128 === 0) ? 66 : 2;
      let at = headerSize;
      const offsets = [at];
      for (const item of block.items) {
        at += item.length;
        offsets.push(at);
      }
      const mapAt = at + (at & 1);
      const out = Buffer.alloc(mapAt + 4 + 2 * offsets.length);
      out.writeUInt16LE(mapAt, 0);
      if (index === 0) {
        out[2] = 0xec;
        out[3] = this.clientSig;
        out.writeUInt32LE(this.userRoot, 4);
      }
      let p = headerSize;
      for (const item of block.items) {
        item.copy(out, p);
        p += item.length;
      }
      out.writeUInt16LE(block.items.length, mapAt);
      out.writeUInt16LE(0, mapAt + 2);
      offsets.forEach((o, i) => out.writeUInt16LE(o, mapAt + 4 + i * 2));
      return out;
    });
  }
}

/* ── the writer ─────────────────────────────────────────────────────────── */

export class PstWriter {
  /**
   * @param {string} filePath
   * @param {{ name?: string, crypt?: 'none'|'permute' }} [opts]
   */
  constructor(filePath, { name = 'Personal Folders', crypt = 'permute' } = {}) {
    this.filePath = filePath;
    this.fd = fs.openSync(filePath, 'w');
    this.crypt = crypt === 'none' ? 0 : 1;
    this.name = name;
    this.pos = 0x4400; // where the first allocation map would sit; data follows it
    this.buffer = Buffer.alloc(8 << 20);
    this.buffered = 0;
    this.bufferAt = this.pos;
    this.nextBid = 4;
    // B-tree entries, kept compact: a big file has hundreds of thousands.
    this.bbt = { bid: new Float64Array(1 << 16), ib: new Float64Array(1 << 16), cb: new Uint16Array(1 << 16), n: 0 };
    this.nbt = [];
    this.folders = [];
    this.nextFolderIndex = 0x400;
    this.nextMessageIndex = 0x10000;
    this.messages = 0;
    this.root = this.#folderRecord('', 0x122, 0x122);
    this.top = this.folder('Top of Personal Folders', this.root);
  }

  #folderRecord(name, nid, parentNid) {
    const record = { name, nid, parentNid, children: [], rows: [], unread: 0 };
    this.folders.push(record);
    return record;
  }

  /** A folder under `parent` (the top of the visible tree when omitted). */
  folder(name, parent = this.top) {
    const nid = makeNid(this.nextFolderIndex++, NID_TYPE.FOLDER);
    const record = this.#folderRecord(name, nid, parent.nid);
    parent.children.push(record);
    return record;
  }

  /* bytes out */

  #write(bytes) {
    if (this.buffered + bytes.length > this.buffer.length) this.#flush();
    if (bytes.length > this.buffer.length) {
      fs.writeSync(this.fd, bytes, 0, bytes.length, this.pos);
      this.bufferAt = this.pos + bytes.length;
    } else {
      bytes.copy(this.buffer, this.buffered);
      this.buffered += bytes.length;
    }
    this.pos += bytes.length;
  }

  #flush() {
    if (this.buffered) fs.writeSync(this.fd, this.buffer, 0, this.buffered, this.bufferAt);
    this.bufferAt += this.buffered;
    this.buffered = 0;
  }

  #pad(to) {
    const gap = (to - (this.pos % to)) % to;
    if (gap) this.#write(Buffer.alloc(gap));
  }

  #recordBlock(bid, ib, cb) {
    const t = this.bbt;
    if (t.n === t.bid.length) {
      const grow = (a, C) => {
        const b = new C(a.length * 2);
        b.set(a);
        return b;
      };
      t.bid = grow(t.bid, Float64Array);
      t.ib = grow(t.ib, Float64Array);
      t.cb = grow(t.cb, Uint16Array);
    }
    t.bid[t.n] = bid;
    t.ib[t.n] = ib;
    t.cb[t.n] = cb;
    t.n++;
  }

  /** One block; data blocks are encoded, internal ones are not. */
  #block(data, internal = false) {
    if (data.length > BLOCK_MAX) throw new Error('a block holds at most 8176 bytes');
    const bid = this.nextBid + (internal ? 2 : 0);
    this.nextBid += 4;
    const total = Math.ceil((data.length + 16) / 64) * 64;
    const out = Buffer.alloc(total);
    if (!internal && this.crypt === 1) for (let i = 0; i < data.length; i++) out[i] = mpbbR[data[i]];
    else data.copy ? data.copy(out, 0) : out.set(data, 0);
    const ib = this.pos;
    out.writeUInt16LE(data.length, total - 16);
    out.writeUInt16LE(sig(ib, bid), total - 14);
    out.writeUInt32LE(crc(out, 0, data.length), total - 12);
    u64(out, total - 8, bid);
    this.#write(out);
    this.#recordBlock(bid, ib, data.length);
    return bid;
  }

  /** Data of any size: one block, or a tree of them. */
  #data(bytes) {
    const data = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
    if (data.length <= BLOCK_MAX) return this.#block(data);
    const leaves = [];
    for (let at = 0; at < data.length; at += BLOCK_MAX) leaves.push(this.#block(data.subarray(at, at + BLOCK_MAX)));
    return this.#tree(leaves, data.length);
  }

  /** Blocks already written, as one node's data: an XBLOCK, or an XXBLOCK of them. */
  #tree(bids, total, sizes = null) {
    if (bids.length === 1) return bids[0];
    const PER = (BLOCK_MAX - 8) / 8; // 1021
    const xblock = (list, level, lcb) => {
      const out = Buffer.alloc(8 + list.length * 8);
      out[0] = 0x01;
      out[1] = level;
      out.writeUInt16LE(list.length, 2);
      out.writeUInt32LE(lcb >>> 0, 4);
      list.forEach((b, i) => u64(out, 8 + i * 8, b));
      return this.#block(out, true);
    };
    if (bids.length <= PER) return xblock(bids, 1, total);
    const mids = [];
    for (let i = 0; i < bids.length; i += PER) {
      const part = bids.slice(i, i + PER);
      const lcb = sizes ? sizes.slice(i, i + PER).reduce((n, s) => n + s, 0) : Math.min(BLOCK_MAX * part.length, total - i * BLOCK_MAX);
      mids.push(xblock(part, 1, lcb));
    }
    if (mids.length > PER) throw new Error('data larger than an XXBLOCK holds');
    return xblock(mids, 2, total);
  }

  /** A heap's blocks as one node's data. */
  #heapData(heap) {
    const blocks = heap.build();
    if (blocks.length === 1) return this.#block(blocks[0]);
    const bids = blocks.map((b) => this.#block(b));
    return this.#tree(bids, blocks.reduce((n, b) => n + b.length, 0), blocks.map((b) => b.length));
  }

  /** A sub-node block over entries `{ nid, dataBid, subBid }`. */
  #subnodeBlock(entries) {
    if (!entries.length) return 0;
    const sorted = entries.slice().sort((a, b) => a.nid - b.nid);
    if (sorted.length > 340) throw new Error('more sub-nodes than one block holds');
    const out = Buffer.alloc(8 + sorted.length * 24);
    out[0] = 0x02;
    out[1] = 0;
    out.writeUInt16LE(sorted.length, 2);
    sorted.forEach((e, i) => {
      u64(out, 8 + i * 24, e.nid);
      u64(out, 16 + i * 24, e.dataBid);
      u64(out, 24 + i * 24, e.subBid || 0);
    });
    return this.#block(out, true);
  }

  /* contexts */

  /**
   * A property context: `props` is `[[id, type, value], …]`. Values too big
   * for the heap go to sub-nodes, which are added to `subnodes`.
   */
  #pc(props, subnodes, counter) {
    const heap = new HeapBuilder(0xbc);
    const records = [];
    const list = props.filter(([, , v]) => v !== undefined && v !== null).sort((a, b) => a[0] - b[0]);
    // Largest values first to the sub-nodes, until the heap fits one block.
    const sized = list.map(([id, type, value]) => ({ id, type, value, bytes: INLINE.has(type) ? null : valueBytes(type, value) }));
    let heapBytes = 12 + 8 + sized.length * 8 + 8;
    const out = new Set();
    for (const p of sized) if (p.bytes && p.bytes.length > HEAP_ITEM_MAX) out.add(p);
    for (const p of sized) if (p.bytes && !out.has(p)) heapBytes += p.bytes.length + 2;
    for (const p of sized.filter((x) => x.bytes && !out.has(x)).sort((a, b) => b.bytes.length - a.bytes.length)) {
      if (heapBytes + 4 + 2 * (sized.length + 3) < BLOCK_MAX - 64) break;
      out.add(p);
      heapBytes -= p.bytes.length + 2;
    }
    for (const p of sized) {
      let ref;
      if (INLINE.has(p.type)) {
        ref = p.type === PT.BOOLEAN ? (p.value ? 1 : 0) : p.type === PT.I2 ? p.value & 0xffff : p.value >>> 0;
      } else if (out.has(p)) {
        const nid = makeNid(counter.next++, NID_TYPE.LTP);
        subnodes.push({ nid, dataBid: this.#data(p.bytes), subBid: 0 });
        ref = nid;
      } else {
        ref = p.bytes.length ? heap.alloc(p.bytes) : 0;
      }
      const rec = Buffer.alloc(8);
      rec.writeUInt16LE(p.id, 0);
      rec.writeUInt16LE(p.type, 2);
      rec.writeUInt32LE(ref >>> 0, 4);
      records.push(rec);
    }
    const recordsHid = records.length ? heap.alloc(Buffer.concat(records)) : 0;
    const bth = Buffer.alloc(8);
    bth[0] = 0xb5;
    bth[1] = 2;
    bth[2] = 6;
    bth[3] = 0;
    bth.writeUInt32LE(recordsHid, 4);
    heap.userRoot = heap.alloc(bth);
    return this.#heapData(heap);
  }

  /**
   * A table context. `columns` is `[[id, type], …]`; each row is a Map or
   * object from id to value, plus `rowId`.
   */
  #tc(columns, rows, subnodes, counter) {
    const heap = new HeapBuilder(0x7c);
    const size = (type) => (type === PT.SYSTIME || type === PT.I8 ? 8 : type === PT.I2 ? 2 : type === PT.BOOLEAN ? 1 : 4);
    const cols = [[0x67f2, PT.LONG], [0x67f3, PT.LONG], ...columns];
    // The row id first, then the cells by size: eights and fours, twos, ones.
    const ordered = [cols[0], ...cols.slice(1).sort((a, b) => size(b[1]) - size(a[1]))];
    let at = 0;
    const layout = ordered.map(([id, type], bit) => {
      const cell = { id, type, size: size(type), offset: at, bit };
      at += cell.size;
      return cell;
    });
    const end4 = layout.filter((c) => c.size >= 4).reduce((n, c) => n + c.size, 0);
    const end2 = end4 + layout.filter((c) => c.size === 2).length * 2;
    const end1 = end2 + layout.filter((c) => c.size === 1).length;
    const rowSize = end1 + Math.ceil(layout.length / 8);

    const strings = (value, type) => {
      const bytes = valueBytes(type, value);
      if (!bytes.length) return 0;
      if (bytes.length > HEAP_ITEM_MAX) {
        const nid = makeNid(counter.next++, NID_TYPE.LTP);
        subnodes.push({ nid, dataBid: this.#data(bytes), subBid: 0 });
        return nid;
      }
      return heap.alloc(bytes);
    };

    const matrix = Buffer.alloc(rows.length * rowSize);
    rows.forEach((row, r) => {
      const base = r * rowSize;
      const get = (id) => (id === 0x67f2 ? row.rowId : id === 0x67f3 ? 1 : row instanceof Map ? row.get(id) : row[id]);
      for (const cell of layout) {
        const value = get(cell.id);
        if (value === undefined || value === null) continue;
        const o = base + cell.offset;
        if (cell.type === PT.LONG) matrix.writeUInt32LE(value >>> 0, o);
        else if (cell.type === PT.I2) matrix.writeUInt16LE(value & 0xffff, o);
        else if (cell.type === PT.BOOLEAN) matrix[o] = value ? 1 : 0;
        else if (cell.size === 8) valueBytes(cell.type, value).copy(matrix, o);
        else matrix.writeUInt32LE(strings(value, cell.type) >>> 0, o);
        matrix[base + end1 + (cell.bit >> 3)] |= 1 << (7 - (cell.bit & 7));
      }
    });

    // Rows in the heap when they fit, else a sub-node of whole rows per block.
    let hnidRows = 0;
    if (rows.length) {
      if (matrix.length <= HEAP_ITEM_MAX) hnidRows = heap.alloc(matrix);
      else {
        const per = Math.floor(BLOCK_MAX / rowSize);
        const bids = [];
        const sizes = [];
        for (let r = 0; r < rows.length; r += per) {
          const part = matrix.subarray(r * rowSize, Math.min(rows.length, r + per) * rowSize);
          bids.push(this.#block(part));
          sizes.push(part.length);
        }
        const nid = makeNid(counter.next++, NID_TYPE.LTP);
        subnodes.push({ nid, dataBid: this.#tree(bids, matrix.length, sizes), subBid: 0 });
        hnidRows = nid;
      }
    }

    // The row index: a BTH from row id to row number, as the format keeps it.
    let rowIndexHid = 0;
    {
      const recs = rows.map((row, i) => [row.rowId >>> 0, i]).sort((a, b) => a[0] - b[0]);
      const per = Math.floor(HEAP_ITEM_MAX / 8);
      const leaves = [];
      for (let i = 0; i < recs.length; i += per) {
        const part = recs.slice(i, i + per);
        const b = Buffer.alloc(part.length * 8);
        part.forEach(([k, v], j) => {
          b.writeUInt32LE(k, j * 8);
          b.writeUInt32LE(v, j * 8 + 4);
        });
        leaves.push({ key: part[0][0], hid: heap.alloc(b) });
      }
      let level = 0;
      let nodes = leaves;
      while (nodes.length > 1) {
        const up = [];
        for (let i = 0; i < nodes.length; i += per) {
          const part = nodes.slice(i, i + per);
          const b = Buffer.alloc(part.length * 8);
          part.forEach((n, j) => {
            b.writeUInt32LE(n.key, j * 8);
            b.writeUInt32LE(n.hid, j * 8 + 4);
          });
          up.push({ key: part[0].key, hid: heap.alloc(b) });
        }
        nodes = up;
        level++;
      }
      const bth = Buffer.alloc(8);
      bth[0] = 0xb5;
      bth[1] = 4;
      bth[2] = 4;
      bth[3] = level;
      bth.writeUInt32LE(nodes[0]?.hid || 0, 4);
      rowIndexHid = heap.alloc(bth);
    }

    const info = Buffer.alloc(22 + layout.length * 8);
    info[0] = 0x7c;
    info[1] = layout.length;
    info.writeUInt16LE(end4, 2);
    info.writeUInt16LE(end2, 4);
    info.writeUInt16LE(end1, 6);
    info.writeUInt16LE(rowSize, 8);
    info.writeUInt32LE(rowIndexHid, 10);
    info.writeUInt32LE(hnidRows >>> 0, 14);
    info.writeUInt32LE(0, 18);
    // Column descriptors sorted by tag, as the format asks.
    layout
      .slice()
      .sort((a, b) => ((a.id << 16) | a.type) - ((b.id << 16) | b.type))
      .forEach((c, i) => {
        info.writeUInt32LE((((c.id << 16) | c.type) >>> 0), 22 + i * 8);
        info.writeUInt16LE(c.offset, 26 + i * 8);
        info[28 + i * 8] = c.size;
        info[29 + i * 8] = c.bit;
      });
    heap.userRoot = heap.alloc(info);
    return this.#heapData(heap);
  }

  /* messages */

  /**
   * One message into `folder`.
   * @param {object} folder from folder()
   * @param {{ subject: string, from: { name?: string, address: string }, to?: object[], cc?: object[], bcc?: object[],
   *   date: string, body?: string, html?: string, messageId?: string, inReplyTo?: string, unread?: boolean,
   *   attachments?: Array<{ filename: string, type?: string, bytes: Uint8Array }> }} m
   */
  message(folder, m) {
    const nid = makeNid(this.nextMessageIndex++, NID_TYPE.MESSAGE);
    const subnodes = [];
    const counter = { next: 1 };
    const attachments = m.attachments || [];
    const to = m.to || [];
    const cc = m.cc || [];
    const bcc = m.bcc || [];
    const size = Buffer.byteLength(m.body || '') + Buffer.byteLength(m.html || '') + attachments.reduce((n, a) => n + a.bytes.length, 0) + 512;
    const flags = (m.unread ? 0 : 0x01) | (attachments.length ? 0x10 : 0);
    const show = (list) => list.map((r) => r.name || r.address).join('; ');

    // Recipients: a table in a sub-node of its own.
    const recipientRows = [
      ...to.map((r) => ({ ...r, kind: 1 })),
      ...cc.map((r) => ({ ...r, kind: 2 })),
      ...bcc.map((r) => ({ ...r, kind: 3 })),
    ].map((r, i) => ({ rowId: i, 0x0c15: r.kind, 0x3001: r.name || r.address, 0x3002: 'SMTP', 0x3003: r.address, 0x39fe: r.address }));
    const recipientSubs = [];
    const recipientCounter = { next: 1 };
    const recipientData = this.#tc(
      [[0x0c15, PT.LONG], [0x3001, PT.UNICODE], [0x3002, PT.UNICODE], [0x3003, PT.UNICODE], [0x39fe, PT.UNICODE]],
      recipientRows,
      recipientSubs,
      recipientCounter
    );
    subnodes.push({ nid: NID_RECIPIENTS, dataBid: recipientData, subBid: this.#subnodeBlock(recipientSubs) });

    // Attachments: a table naming them, and each one a sub-node of its own.
    if (attachments.length) {
      const rows = [];
      attachments.forEach((a, i) => {
        const attachNid = makeNid(0x8000 + i, NID_TYPE.ATTACHMENT);
        const own = [];
        const ownCounter = { next: 1 };
        const name = a.filename || `attachment-${i + 1}`;
        const data = this.#pc(
          [
            [0x0e20, PT.LONG, a.bytes.length],
            [0x3701, PT.BINARY, Buffer.from(a.bytes)],
            [0x3704, PT.UNICODE, name.slice(0, 12)],
            [0x3705, PT.LONG, 1],
            [0x3707, PT.UNICODE, name],
            [0x370b, PT.LONG, -1],
            [0x370e, PT.UNICODE, a.type || 'application/octet-stream'],
          ],
          own,
          ownCounter
        );
        subnodes.push({ nid: attachNid, dataBid: data, subBid: this.#subnodeBlock(own) });
        rows.push({ rowId: attachNid, 0x0e20: a.bytes.length, 0x3707: name, 0x3705: 1, 0x370b: -1 });
      });
      const tableSubs = [];
      const data = this.#tc([[0x0e20, PT.LONG], [0x3705, PT.LONG], [0x3707, PT.UNICODE], [0x370b, PT.LONG]], rows, tableSubs, { next: 1 });
      subnodes.push({ nid: NID_ATTACHMENTS, dataBid: data, subBid: this.#subnodeBlock(tableSubs) });
    }

    const props = [
      [0x001a, PT.UNICODE, 'IPM.Note'],
      [0x0037, PT.UNICODE, m.subject || ''],
      [0x0039, PT.SYSTIME, m.date],
      [0x0042, PT.UNICODE, m.from?.name || m.from?.address],
      [0x0065, PT.UNICODE, m.from?.address],
      [0x0070, PT.UNICODE, m.subject || ''],
      [0x0c1a, PT.UNICODE, m.from?.name || m.from?.address],
      [0x0c1e, PT.UNICODE, 'SMTP'],
      [0x0c1f, PT.UNICODE, m.from?.address],
      [0x0e03, PT.UNICODE, show(cc)],
      [0x0e04, PT.UNICODE, show(to)],
      [0x0e06, PT.SYSTIME, m.date],
      [0x0e07, PT.LONG, flags],
      [0x0e08, PT.LONG, size],
      [0x0e1d, PT.UNICODE, m.subject || ''],
      [0x1000, PT.UNICODE, m.body || null],
      [0x1013, PT.BINARY, m.html ? Buffer.from(m.html, 'utf8') : null],
      [0x1035, PT.UNICODE, m.messageId ? `<${m.messageId}>` : null],
      [0x1042, PT.UNICODE, m.inReplyTo ? `<${m.inReplyTo}>` : null],
      [0x3007, PT.SYSTIME, m.date],
      [0x5d01, PT.UNICODE, m.from?.address],
    ];
    const dataBid = this.#pc(props, subnodes, counter);
    const subBid = this.#subnodeBlock(subnodes);
    this.nbt.push({ nid, dataBid, subBid, parent: folder.nid });
    folder.rows.push({
      rowId: nid,
      0x0037: m.subject || '',
      0x0c1a: m.from?.name || m.from?.address || '',
      0x0e04: show(to),
      0x0e06: m.date,
      0x0e07: flags,
      0x0e08: size,
      0x001a: 'IPM.Note',
    });
    if (m.unread) folder.unread++;
    this.messages++;
    return nid;
  }

  /* closing */

  #folderNodes(folder) {
    const isRoot = folder.nid === 0x122;
    const children = folder.children;
    // The folder itself.
    const subs = [];
    const counter = { next: 1 };
    const pc = this.#pc(
      [
        [0x3001, PT.UNICODE, folder.name],
        [0x3602, PT.LONG, folder.rows.length],
        [0x3603, PT.LONG, folder.unread],
        [0x360a, PT.BOOLEAN, children.length > 0],
        [0x3613, PT.UNICODE, isRoot ? null : 'IPF.Note'],
      ],
      subs,
      counter
    );
    this.nbt.push({ nid: folder.nid, dataBid: pc, subBid: this.#subnodeBlock(subs), parent: folder.parentNid });

    // Its subfolders, in a hierarchy table.
    const hSubs = [];
    const hierarchy = this.#tc(
      [[0x3001, PT.UNICODE], [0x3602, PT.LONG], [0x3603, PT.LONG], [0x360a, PT.BOOLEAN]],
      children.map((c) => ({ rowId: c.nid, 0x3001: c.name, 0x3602: c.rows.length, 0x3603: c.unread, 0x360a: c.children.length > 0 })),
      hSubs,
      { next: 1 }
    );
    this.nbt.push({ nid: ((folder.nid & ~0x1f) | NID_TYPE.HIERARCHY) >>> 0, dataBid: hierarchy, subBid: this.#subnodeBlock(hSubs), parent: 0 });

    // Its messages, in a contents table with the columns a list shows.
    const cSubs = [];
    const contents = this.#tc(
      [[0x001a, PT.UNICODE], [0x0037, PT.UNICODE], [0x0c1a, PT.UNICODE], [0x0e04, PT.UNICODE], [0x0e06, PT.SYSTIME], [0x0e07, PT.LONG], [0x0e08, PT.LONG]],
      folder.rows,
      cSubs,
      { next: 1 }
    );
    this.nbt.push({ nid: ((folder.nid & ~0x1f) | NID_TYPE.CONTENTS) >>> 0, dataBid: contents, subBid: this.#subnodeBlock(cSubs), parent: 0 });

    const aSubs = [];
    const assoc = this.#tc([[0x001a, PT.UNICODE]], [], aSubs, { next: 1 });
    this.nbt.push({ nid: ((folder.nid & ~0x1f) | NID_TYPE.ASSOC) >>> 0, dataBid: assoc, subBid: 0, parent: 0 });
    // A long folder's rows are not needed once written.
    folder.count = folder.rows.length;
    folder.rows = [];
  }

  /** Pages of a B-tree, leaves first; returns the root's BREF. */
  #btree(entries, leafSize, writeLeaf, ptype) {
    const writePage = (body, cEnt, cbEnt, level) => {
      this.#pad(PAGE);
      const page = Buffer.alloc(PAGE);
      body.copy(page, 0);
      page[488] = cEnt;
      page[489] = Math.floor(488 / cbEnt);
      page[490] = cbEnt;
      page[491] = level;
      const bid = this.nextBid;
      this.nextBid += 4;
      const ib = this.pos;
      page[496] = ptype;
      page[497] = ptype;
      page.writeUInt16LE(sig(ib, bid), 498);
      page.writeUInt32LE(crc(page, 0, 496), 500);
      u64(page, 504, bid);
      this.#write(page);
      return { bid, ib };
    };
    const perLeaf = Math.floor(488 / leafSize);
    let level = [];
    for (let i = 0; i < Math.max(1, entries.length); i += perLeaf) {
      const part = entries.slice(i, i + perLeaf);
      const body = Buffer.alloc(488);
      part.forEach((e, j) => writeLeaf(body, j * leafSize, e));
      const ref = writePage(body, part.length, leafSize, 0);
      level.push({ key: part[0]?.key ?? 0, ...ref });
    }
    let depth = 1;
    while (level.length > 1) {
      const up = [];
      for (let i = 0; i < level.length; i += 20) {
        const part = level.slice(i, i + 20);
        const body = Buffer.alloc(488);
        part.forEach((e, j) => {
          u64(body, j * 24, e.key);
          u64(body, j * 24 + 8, e.bid);
          u64(body, j * 24 + 16, e.ib);
        });
        up.push({ key: part[0].key, ...writePage(body, part.length, 24, depth) });
      }
      level = up;
      depth++;
    }
    return level[0];
  }

  /** Finish the file: folders, the store, both B-trees and the header. */
  close() {
    // Folders, deepest last so every parent is known.
    for (const folder of this.folders) this.#folderNodes(folder);

    // The message store: its name, and where the visible tree starts.
    const entryId = Buffer.alloc(24);
    entryId.writeUInt32LE(this.top.nid, 20);
    const storeSubs = [];
    const store = this.#pc(
      [
        [0x0ff9, PT.BINARY, Buffer.alloc(16, 0x5a)],
        [0x3001, PT.UNICODE, this.name],
        [0x35e0, PT.BINARY, entryId],
      ],
      storeSubs,
      { next: 1 }
    );
    this.nbt.push({ nid: 0x21, dataBid: store, subBid: this.#subnodeBlock(storeSubs), parent: 0 });

    // The block B-tree, sorted by block id.
    const t = this.bbt;
    const order = Array.from({ length: t.n }, (_, i) => i).sort((a, b) => t.bid[a] - t.bid[b]);
    const bbtEntries = order.map((i) => ({ key: t.bid[i], bid: t.bid[i], ib: t.ib[i], cb: t.cb[i] }));
    const bbtRoot = this.#btree(bbtEntries, 24, (body, at, e) => {
      u64(body, at, e.bid);
      u64(body, at + 8, e.ib);
      body.writeUInt16LE(e.cb, at + 16);
      body.writeUInt16LE(2, at + 18);
    }, 0x80);

    // The node B-tree, sorted by node id.
    const nbtEntries = this.nbt.sort((a, b) => a.nid - b.nid).map((e) => ({ ...e, key: e.nid }));
    const nbtRoot = this.#btree(nbtEntries, 32, (body, at, e) => {
      u64(body, at, e.nid);
      u64(body, at + 8, e.dataBid);
      u64(body, at + 16, e.subBid || 0);
      body.writeUInt32LE(e.parent >>> 0, at + 24);
    }, 0x81);

    this.#pad(PAGE);
    this.#flush();
    const eof = this.pos;

    // The header, [MS-PST] §2.2.2.6.
    const h = Buffer.alloc(564);
    h.write('!BDN', 0, 'latin1');
    h.writeUInt16LE(0x4d53, 8); // "SM", a .pst
    h.writeUInt16LE(23, 10);
    h.writeUInt16LE(19, 12);
    h[14] = 1;
    h[15] = 1;
    u64(h, 32, this.nextBid); // bidNextP
    h.writeUInt32LE(1, 40);
    const root = 0xb4;
    u64(h, root + 4, eof);
    u64(h, root + 12, 0x4400);
    u64(h, root + 36, nbtRoot.bid);
    u64(h, root + 44, nbtRoot.ib);
    u64(h, root + 52, bbtRoot.bid);
    u64(h, root + 60, bbtRoot.ib);
    h[root + 68] = 0; // the allocation maps are not written
    h.fill(0xff, 256, 512);
    h[0x200] = 0x80;
    h[0x201] = this.crypt;
    u64(h, 0x204, this.nextBid); // bidNextB
    h.writeUInt32LE(crc(h, 8, 8 + 471), 4);
    h.writeUInt32LE(crc(h, 8, 8 + 516), 0x20c);
    fs.writeSync(this.fd, h, 0, h.length, 0);
    fs.ftruncateSync(this.fd, eof);
    fs.closeSync(this.fd);
    return { path: this.filePath, bytes: eof, messages: this.messages, folders: this.folders.length - 2 };
  }
}

/** Write a .pst of `folders`: `[{ name, messages: [...], children?: [...] }]`. */
export function writePst(filePath, folders, opts = {}) {
  const w = new PstWriter(filePath, opts);
  const add = (spec, parent) => {
    const folder = w.folder(spec.name, parent);
    for (const m of spec.messages || []) w.message(folder, m);
    for (const child of spec.children || []) add(child, folder);
  };
  for (const spec of folders) add(spec);
  return w.close();
}
