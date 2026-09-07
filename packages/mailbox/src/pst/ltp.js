// The lists, tables and properties layer — [MS-PST] §2.3.
//
// Above the raw nodes sits a tiny allocator and two data structures built on
// it. Every folder, message and attachment in a PST is one of the two:
//
//   Heap-on-Node (HN)   a node's bytes carved into numbered items. An item is
//                       addressed by a heap id, which names a block and an
//                       index inside that block's allocation map.
//   Property Context    a B-tree over the heap, keyed by property id. This is
//   (PC)                what a folder or a message *is*.
//   Table Context (TC)  a row matrix with a column descriptor per property.
//                       Folder contents and recipient lists are tables.
//
// Values too large for the heap live in sub-nodes, so reading a property means
// following one of two kinds of reference. That indirection is the single most
// common source of bugs in PST readers, so it is handled in exactly one place
// here: `resolveHnid`.

import { PT, decodeValue, describeTag } from '../props.js';

const HN_SIG = 0xec;

export const CLIENT_SIG = {
  RESERVED_1: 0x6c,
  TABLE_CONTEXT: 0x7c,
  RESERVED_2: 0x8c,
  RESERVED_3: 0x9c,
  RESERVED_4: 0xa5,
  RESERVED_5: 0xac,
  BTH: 0xb5,
  PROPERTY_CONTEXT: 0xbc,
  RESERVED_6: 0xcc,
};

export const hidIsHid = (hnid) => (hnid & 0x1f) === 0;
export const hidIndex = (hid) => (hid >>> 5) & 0x7ff;
export const hidBlockIndex = (hid) => (hid >>> 16) & 0xffff;

/**
 * A heap spread over a node's blocks.
 *
 * Every block carries an allocation map at the offset its header names; item N
 * of block B runs between entries N-1 and N of that map. Page headers differ by
 * block: the first has the heap header, every 128th after the eighth has a
 * fill-level bitmap, and the rest have a two-byte header — but all three end
 * with the same map offset, which is why only the map is read here.
 */
export class Heap {
  /** @param {Uint8Array[]} parts the node's blocks, in order */
  constructor(parts) {
    this.parts = (parts || []).filter((p) => p && p.length >= 4);
    const first = this.parts[0];
    if (!first || first.length < 12) {
      this.valid = false;
      return;
    }
    const dv = new DataView(first.buffer, first.byteOffset, first.byteLength);
    this.valid = first[2] === HN_SIG;
    this.clientSig = first[3];
    this.userRoot = dv.getUint32(4, true);
  }

  /** Bytes of one heap item, or an empty array when the id does not resolve. */
  get(hid) {
    if (!hid || !hidIsHid(hid)) return new Uint8Array(0);
    const block = this.parts[hidBlockIndex(hid)];
    if (!block || block.length < 4) return new Uint8Array(0);
    const index = hidIndex(hid);
    if (index < 1) return new Uint8Array(0);
    const dv = new DataView(block.buffer, block.byteOffset, block.byteLength);
    const mapAt = dv.getUint16(0, true);
    if (mapAt + 4 > block.length) return new Uint8Array(0);
    const cAlloc = dv.getUint16(mapAt, true);
    if (index > cAlloc) return new Uint8Array(0);
    const startAt = mapAt + 4 + (index - 1) * 2;
    if (startAt + 4 > block.length) return new Uint8Array(0);
    const start = dv.getUint16(startAt, true);
    const end = dv.getUint16(startAt + 2, true);
    if (end < start || end > block.length) return new Uint8Array(0);
    return block.subarray(start, end);
  }
}

/**
 * A B-tree inside the heap. Two shapes matter: a property context's index
 * (2-byte key, 6-byte record) and a table's row index (4 and 4).
 */
export function readBth(heap, hidRoot) {
  const header = heap.get(hidRoot);
  if (header.length < 8 || header[0] !== CLIENT_SIG.BTH) return { entries: [], keySize: 0, entrySize: 0 };
  const cbKey = header[1];
  const cbEnt = header[2];
  const levels = header[3];
  const dv = new DataView(header.buffer, header.byteOffset, header.byteLength);
  const root = dv.getUint32(4, true);

  const entries = [];
  const walk = (hid, level) => {
    const bytes = heap.get(hid);
    if (!bytes.length) return;
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (level > 0) {
      const size = cbKey + 4;
      for (let at = 0; at + size <= bytes.length; at += size) {
        walk(view.getUint32(at + cbKey, true), level - 1);
      }
      return;
    }
    const size = cbKey + cbEnt;
    for (let at = 0; at + size <= bytes.length; at += size) {
      entries.push({
        key: cbKey === 2 ? view.getUint16(at, true) : view.getUint32(at, true),
        data: bytes.subarray(at + cbKey, at + size),
      });
    }
  };
  walk(root, levels);
  return { entries, keySize: cbKey, entrySize: cbEnt };
}

/**
 * Follow a value reference: either into the heap, or out to a sub-node.
 * The low five bits decide which — zero means a heap id.
 */
export function resolveHnid(hnid, heap, subnodes, ndb) {
  if (!hnid) return new Uint8Array(0);
  if (hidIsHid(hnid)) return heap.get(hnid);
  const sub = subnodes?.get(hnid >>> 0);
  if (!sub || !ndb) return new Uint8Array(0);
  return ndb.blockData(sub.dataBid);
}

/**
 * A property context: the properties of a folder, a message, an attachment or
 * a recipient.
 * @returns {Map<number, { id: number, type: number, name: string, value: any }>}
 */
export function readPc(heap, subnodes, ndb) {
  const out = new Map();
  if (!heap?.valid || heap.clientSig !== CLIENT_SIG.PROPERTY_CONTEXT) return out;
  const { entries } = readBth(heap, heap.userRoot);
  for (const entry of entries) {
    if (entry.data.length < 6) continue;
    const dv = new DataView(entry.data.buffer, entry.data.byteOffset, entry.data.byteLength);
    const type = dv.getUint16(0, true);
    const raw = dv.getUint32(2, true);
    const id = entry.key;
    let value;
    if (type === PT.I2 || type === PT.LONG || type === PT.R4 || type === PT.ERROR) {
      // Small fixed values sit in the reference itself rather than the heap.
      const inline = new Uint8Array(4);
      new DataView(inline.buffer).setUint32(0, raw, true);
      value = decodeValue(type, inline);
    } else if (type === PT.BOOLEAN) {
      value = (raw & 0xff) !== 0;
    } else if (type === PT.NULL || type === PT.UNSPECIFIED) {
      value = null;
    } else {
      const bytes = resolveHnid(raw, heap, subnodes, ndb);
      value = decodeValue(type, bytes);
    }
    out.set(id, { id, type, name: describeTag((id << 16) | type).name, value });
  }
  return out;
}

/**
 * A table context: the rows of a folder's contents, a message's recipients or
 * its attachments.
 * @returns {{ columns: Array, rows: Array<Map<number, any>> }}
 */
export function readTc(node, ndb) {
  const heap = new Heap(node.parts);
  const empty = { columns: [], rows: [] };
  if (!heap.valid || heap.clientSig !== CLIENT_SIG.TABLE_CONTEXT) return empty;

  const info = heap.get(heap.userRoot);
  if (info.length < 22 || info[0] !== CLIENT_SIG.TABLE_CONTEXT) return empty;
  const dv = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const cCols = info[1];
  const endOf4 = dv.getUint16(2, true);
  const endOf2 = dv.getUint16(4, true);
  const endOf1 = dv.getUint16(6, true); // where the cell-existence bitmap starts
  const rowSize = dv.getUint16(8, true);
  const hnidRows = dv.getUint32(14, true);

  const columns = [];
  for (let i = 0; i < cCols; i++) {
    const at = 22 + i * 8;
    if (at + 8 > info.length) break;
    const tag = dv.getUint32(at, true);
    const { id, type, name } = describeTag(tag);
    columns.push({ id, type, name, offset: dv.getUint16(at + 4, true), size: info[at + 6], bit: info[at + 7] });
  }
  if (!rowSize || !columns.length) return { columns, rows: [] };

  // Row data is either a heap item or a sub-node; either way rows never span a
  // block, so each block holds floor(size / rowSize) of them.
  let blocks;
  if (hidIsHid(hnidRows)) {
    const bytes = heap.get(hnidRows);
    blocks = bytes.length ? [bytes] : [];
  } else {
    const sub = node.subnodes?.get(hnidRows >>> 0);
    blocks = sub ? ndb.blockParts(sub.dataBid) : [];
  }

  const rows = [];
  for (const block of blocks) {
    const perBlock = Math.floor(block.length / rowSize);
    for (let r = 0; r < perBlock; r++) {
      const at = r * rowSize;
      const row = block.subarray(at, at + rowSize);
      const rdv = new DataView(row.buffer, row.byteOffset, row.byteLength);
      const ceb = row.subarray(endOf1, rowSize);
      const values = new Map();
      values.set('rowId', rdv.getUint32(0, true));
      for (const col of columns) {
        const byte = ceb[col.bit >> 3];
        const present = byte != null && (byte & (1 << (7 - (col.bit & 7)))) !== 0;
        if (!present) continue;
        if (col.offset + col.size > rowSize) continue;
        const cell = row.subarray(col.offset, col.offset + col.size);
        let value;
        if (col.type === PT.I2 || col.type === PT.LONG || col.type === PT.R4 || col.type === PT.ERROR) {
          value = decodeValue(col.type, cell);
        } else if (col.type === PT.BOOLEAN) {
          value = cell[0] !== 0;
        } else if (col.size === 4) {
          const hnid = new DataView(cell.buffer, cell.byteOffset, 4).getUint32(0, true);
          value = decodeValue(col.type, resolveHnid(hnid, heap, node.subnodes, ndb));
        } else {
          value = decodeValue(col.type, cell);
        }
        values.set(col.id, value);
      }
      rows.push(values);
    }
  }
  void endOf4;
  void endOf2;
  return { columns, rows };
}

/** A property context read straight from a node. */
export function nodePc(node, ndb) {
  const heap = new Heap(node.parts);
  return readPc(heap, node.subnodes, ndb);
}

/** Plain values from a property map, by name. */
export function values(pc) {
  const out = {};
  for (const p of pc.values()) if (p.name && !/^0x/.test(p.name)) out[p.name] = p.value;
  return out;
}
