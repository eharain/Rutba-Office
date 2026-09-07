// Compound File Binary Format — [MS-CFB].
//
// The container under .msg, and under Word/Excel/PowerPoint 97-2003. It is a
// FAT filesystem in a file: a header, a sector allocation table, a red-black
// tree of directory entries, and a second smaller allocation table for streams
// under 4 KB so that a 12-byte property does not cost a whole sector.
//
// This is a reader. Nothing here writes, because everything we write is OOXML.

const SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;
const DIFSECT = 0xfffffffc;

export class CfbError extends Error {
  constructor(message) {
    super(message);
    this.name = 'CfbError';
  }
}

/** @typedef {{ name: string, type: number, size: number, start: number, children: number[], clsid: string, mtime: number }} Entry */

export class CompoundFile {
  /** @param {Uint8Array} bytes */
  constructor(bytes) {
    this.b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    this.view = new DataView(this.b.buffer, this.b.byteOffset, this.b.byteLength);
    this.#readHeader();
    this.#readFat();
    this.#readDirectory();
    this.#readMiniFat();
  }

  static is(bytes) {
    if (!bytes || bytes.length < 8) return false;
    for (let i = 0; i < 8; i++) if (bytes[i] !== SIG[i]) return false;
    return true;
  }

  u32(at) {
    return this.view.getUint32(at, true);
  }

  u16(at) {
    return this.view.getUint16(at, true);
  }

  #readHeader() {
    if (!CompoundFile.is(this.b)) throw new CfbError('not a compound file');
    this.major = this.u16(26);
    this.sectorShift = this.u16(30);
    this.miniSectorShift = this.u16(32);
    this.sectorSize = 1 << this.sectorShift;
    this.miniSectorSize = 1 << this.miniSectorShift;
    if (this.sectorSize !== 512 && this.sectorSize !== 4096) {
      throw new CfbError(`unsupported sector size ${this.sectorSize}`);
    }
    this.fatCount = this.u32(44);
    this.dirStart = this.u32(48);
    this.miniCutoff = this.u32(56);
    this.miniFatStart = this.u32(60);
    this.miniFatCount = this.u32(64);
    this.difatStart = this.u32(68);
    this.difatCount = this.u32(72);
  }

  /** Byte offset of a sector in the file. Sector 0 starts after the header. */
  #offset(sector) {
    return (sector + 1) * this.sectorSize;
  }

  #sector(n) {
    const at = this.#offset(n);
    if (at + this.sectorSize > this.b.length) {
      // Truncated file: hand back what exists rather than throwing, so a
      // partially copied .pst or .msg still shows the messages it does have.
      return this.b.subarray(Math.min(at, this.b.length));
    }
    return this.b.subarray(at, at + this.sectorSize);
  }

  #readFat() {
    // The DIFAT lists the sectors that hold the FAT. The first 109 entries live
    // in the header; the rest are chained through their own sectors.
    const difat = [];
    for (let i = 0; i < 109; i++) {
      const s = this.u32(76 + i * 4);
      if (s === FREESECT || s === ENDOFCHAIN) break;
      difat.push(s);
    }
    let next = this.difatStart;
    let guard = 0;
    while (next !== ENDOFCHAIN && next !== FREESECT && guard++ < 1 << 20) {
      const sec = this.#sector(next);
      const per = this.sectorSize / 4 - 1;
      const dv = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
      for (let i = 0; i < per; i++) {
        const s = dv.getUint32(i * 4, true);
        if (s !== FREESECT && s !== ENDOFCHAIN) difat.push(s);
      }
      next = dv.getUint32(per * 4, true);
    }

    const fat = new Uint32Array(difat.length * (this.sectorSize / 4));
    let at = 0;
    for (const s of difat) {
      const sec = this.#sector(s);
      const dv = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
      const n = Math.floor(sec.byteLength / 4);
      for (let i = 0; i < n; i++) fat[at++] = dv.getUint32(i * 4, true);
    }
    this.fat = fat;
  }

  #chain(start, fat) {
    const out = [];
    let s = start;
    let guard = 0;
    while (s !== ENDOFCHAIN && s !== FREESECT && s !== FATSECT && s !== DIFSECT && guard++ < 1 << 22) {
      out.push(s);
      s = fat[s];
      if (s === undefined) break;
    }
    return out;
  }

  #readDirectory() {
    const sectors = this.#chain(this.dirStart, this.fat);
    /** @type {Entry[]} */
    const entries = [];
    const perSector = this.sectorSize / 128;
    for (const s of sectors) {
      const sec = this.#sector(s);
      const dv = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
      for (let i = 0; i < perSector && (i + 1) * 128 <= sec.byteLength; i++) {
        const at = i * 128;
        const nameLen = dv.getUint16(at + 64, true);
        let name = '';
        for (let c = 0; c + 1 < Math.max(0, nameLen - 2); c += 2) {
          name += String.fromCharCode(dv.getUint16(at + c, true));
        }
        const type = sec[at + 66];
        const left = dv.getUint32(at + 68, true);
        const right = dv.getUint32(at + 72, true);
        const child = dv.getUint32(at + 76, true);
        const clsidBytes = sec.subarray(at + 80, at + 96);
        const start = dv.getUint32(at + 116, true);
        const lo = dv.getUint32(at + 120, true);
        const hi = dv.getUint32(at + 124, true);
        // 32-bit sizes only in v3; the high word is a stream size for v4.
        const size = this.major >= 4 ? lo + hi * 2 ** 32 : lo;
        const ftLow = dv.getUint32(at + 108, true);
        const ftHigh = dv.getUint32(at + 112, true);
        entries.push({
          index: entries.length,
          name,
          type,
          left,
          right,
          child,
          start,
          size,
          clsid: [...clsidBytes].map((x) => x.toString(16).padStart(2, '0')).join(''),
          mtime: filetimeToMs(ftLow, ftHigh),
          children: [],
        });
      }
    }
    this.entries = entries;

    // Rebuild the tree: each storage names one child, and siblings hang off it
    // in a red-black tree we only need to walk, never to balance.
    const walk = (id, into) => {
      if (id === FREESECT || id >= entries.length) return;
      const e = entries[id];
      if (!e) return;
      walk(e.left, into);
      into.push(id);
      walk(e.right, into);
    };
    for (const e of entries) {
      if (e.type === 1 || e.type === 5) {
        const kids = [];
        walk(e.child, kids);
        e.children = kids;
      }
    }
    this.root = entries[0] || null;
  }

  #readMiniFat() {
    if (!this.root) {
      this.miniFat = new Uint32Array(0);
      this.miniStream = new Uint8Array(0);
      return;
    }
    const sectors = this.#chain(this.miniFatStart, this.fat);
    const mini = new Uint32Array(sectors.length * (this.sectorSize / 4));
    let at = 0;
    for (const s of sectors) {
      const sec = this.#sector(s);
      const dv = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
      const n = Math.floor(sec.byteLength / 4);
      for (let i = 0; i < n; i++) mini[at++] = dv.getUint32(i * 4, true);
    }
    this.miniFat = mini;
    this.miniStream = this.#readChainBytes(this.root.start, this.root.size, this.fat, this.sectorSize);
  }

  #readChainBytes(start, size, fat, unit) {
    const chain = this.#chain(start, fat);
    const out = new Uint8Array(Math.min(size, chain.length * unit));
    let at = 0;
    for (const s of chain) {
      if (at >= out.length) break;
      const sec = unit === this.sectorSize ? this.#sector(s) : null;
      const src = sec ?? this.miniStream.subarray(s * unit, (s + 1) * unit);
      const take = Math.min(unit, out.length - at, src.length);
      out.set(src.subarray(0, take), at);
      at += take;
    }
    return out.subarray(0, Math.min(size, at));
  }

  /** Bytes of one directory entry's stream. */
  read(entry) {
    const e = typeof entry === 'number' ? this.entries[entry] : entry;
    if (!e) return new Uint8Array(0);
    if (e.size === 0) return new Uint8Array(0);
    if (e.size < this.miniCutoff && e.index !== 0) {
      return this.#readChainBytes(e.start, e.size, this.miniFat, this.miniSectorSize);
    }
    return this.#readChainBytes(e.start, e.size, this.fat, this.sectorSize);
  }

  /** Direct children of a storage, by name. */
  childrenOf(entry) {
    const e = typeof entry === 'number' ? this.entries[entry] : entry;
    return (e?.children || []).map((i) => this.entries[i]).filter(Boolean);
  }

  /** Find one entry by path, e.g. ['__attach_version1.0_#00000000', '__substg1.0_37010102']. */
  find(pathParts, from = this.root) {
    let node = from;
    for (const part of pathParts) {
      const kid = this.childrenOf(node).find((c) => c.name === part);
      if (!kid) return null;
      node = kid;
    }
    return node;
  }

  /** Every stream name in the file, for identification. */
  names() {
    return this.entries.map((e) => e.name);
  }

  /**
   * Which 97-2003 application wrote this: the root storage's stream names are
   * unambiguous, where the file extension is not.
   */
  application() {
    const names = new Set(this.childrenOf(this.root).map((c) => c.name));
    if (names.has('WordDocument')) return 'doc';
    if (names.has('Workbook') || names.has('Book')) return 'xls';
    if (names.has('PowerPoint Document')) return 'ppt';
    if ([...names].some((n) => n.startsWith('__substg1.0_') || n === '__properties_version1.0')) return 'msg';
    return null;
  }
}

function filetimeToMs(low, high) {
  const ft = high * 2 ** 32 + low;
  if (!ft) return 0;
  // FILETIME is 100 ns ticks since 1601-01-01.
  return Math.round(ft / 10000 - 11644473600000);
}

export { filetimeToMs };
