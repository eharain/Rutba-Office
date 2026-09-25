// Compound File Binary Format — the writing half of [MS-CFB].
//
// The container a password-protected .docx, .xlsx or .pptx is written in:
// the encrypted package, the description of how it was encrypted, and the
// \x06DataSpaces storage that names the transform. Version 3 (512-byte
// sectors), which is what Word, Excel and PowerPoint write for these files.
//
// Everything a reader needs and nothing else: the header, the FAT (with DIFAT
// sectors once a file outgrows the 109 the header can list), the directory
// with each storage's children in a red-black tree, and the mini stream with
// its own allocation table for the streams under 4096 bytes. The spec says a
// stream under the cutoff MUST live in the mini stream, so there is no
// "everything in regular sectors" shortcut: a reader decides where to look
// by the size alone.

const SECTOR = 512;
const MINI = 64;
const CUTOFF = 4096;
const PER_FAT = SECTOR / 4; // 128 sector numbers in a FAT sector
const PER_DIFAT = PER_FAT - 1; // 127, the last slot chains to the next

const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const FATSECT = 0xfffffffd;
const DIFSECT = 0xfffffffc;
const NOSTREAM = 0xffffffff;

const STORAGE = 1;
const STREAM = 2;
const ROOT = 5;
const RED = 0;
const BLACK = 1;

/**
 * [MS-CFB] 2.6.4: siblings are ordered by name length first, then by the
 * names compared character by character in upper case.
 */
export function compareNames(a, b) {
  if (a.length !== b.length) return a.length - b.length;
  const A = a.toUpperCase();
  const B = b.toUpperCase();
  for (let i = 0; i < A.length; i++) {
    const d = A.charCodeAt(i) - B.charCodeAt(i);
    if (d) return d;
  }
  return 0;
}

/**
 * Write a compound file.
 *
 * @param {{ path: string[], data: Uint8Array }[]} streams each stream's path
 *   from the root, e.g. ['\x06DataSpaces', 'Version']; the storages on the
 *   way are made as they are named.
 * @returns {Buffer}
 */
export function writeCompoundFile(streams) {
  /* ── the tree ─────────────────────────────────────────────────────────── */
  const root = { name: 'Root Entry', type: ROOT, kids: new Map(), data: null };
  for (const { path, data } of streams) {
    if (!Array.isArray(path) || !path.length) throw new Error('a stream needs a path');
    let node = root;
    path.forEach((part, i) => {
      const name = String(part);
      if (!name || name.length > 31 || /[/\\:!]/.test(name)) throw new Error(`"${name}" cannot name a compound-file entry`);
      const last = i === path.length - 1;
      let kid = node.kids.get(name.toUpperCase());
      if (!kid) {
        kid = last ? { name, type: STREAM, data: toBytes(data) } : { name, type: STORAGE, kids: new Map() };
        node.kids.set(name.toUpperCase(), kid);
      } else if (last || kid.type !== STORAGE) {
        throw new Error(`"${path.join('/')}" is named twice`);
      }
      node = kid;
    });
  }

  // Number the entries: the root first, then each storage's children in
  // order, depth first.
  const entries = [];
  const number = (node) => {
    node.index = entries.length;
    entries.push(node);
    if (node.kids) {
      node.sorted = [...node.kids.values()].sort((x, y) => compareNames(x.name, y.name));
      for (const kid of node.sorted) number(kid);
    }
  };
  number(root);

  // Each storage's children as a red-black tree: split the sorted list at
  // its middle, recursively. Every null is then at the tree's last level or
  // the one above it, so colouring the nodes on the last level red and every
  // other node black gives each path the same number of black nodes, and no
  // red node a red child — a valid tree, never rebalanced because nothing
  // is ever inserted into it.
  for (const node of entries) {
    node.left = NOSTREAM;
    node.right = NOSTREAM;
    node.child = NOSTREAM;
    node.color = BLACK;
  }
  for (const node of entries) {
    if (!node.sorted?.length) continue;
    const depthOf = new Map();
    const build = (lo, hi, depth) => {
      if (lo > hi) return NOSTREAM;
      const mid = (lo + hi) >> 1;
      const kid = node.sorted[mid];
      depthOf.set(kid, depth);
      kid.left = build(lo, mid - 1, depth + 1);
      kid.right = build(mid + 1, hi, depth + 1);
      return kid.index;
    };
    node.child = build(0, node.sorted.length - 1, 0);
    const deepest = Math.max(...depthOf.values());
    for (const [kid, depth] of depthOf) kid.color = depth === deepest && deepest > 0 ? RED : BLACK;
  }

  /* ── where each stream goes ───────────────────────────────────────────── */
  const small = [];
  const large = [];
  for (const node of entries) {
    if (node.type !== STREAM) continue;
    if (!node.data.length) {
      node.start = ENDOFCHAIN;
      continue;
    }
    (node.data.length < CUTOFF ? small : large).push(node);
  }

  // The mini stream: each small stream in 64-byte mini sectors, chained in
  // the mini FAT.
  let miniSectors = 0;
  for (const node of small) {
    node.start = miniSectors;
    miniSectors += Math.ceil(node.data.length / MINI);
  }
  const miniStreamBytes = miniSectors * MINI;
  const miniFatSectors = Math.ceil((miniSectors * 4) / SECTOR);
  const miniStreamSectors = Math.ceil(miniStreamBytes / SECTOR);
  const dirSectors = Math.ceil((entries.length * 128) / SECTOR);
  const largeSectors = large.reduce((n, node) => n + Math.ceil(node.data.length / SECTOR), 0);
  const content = dirSectors + miniFatSectors + miniStreamSectors + largeSectors;

  // The FAT covers every sector, its own included, and past 109 FAT sectors
  // the DIFAT needs sectors of its own, which the FAT covers too.
  let fatSectors = Math.max(1, Math.ceil(content / PER_FAT));
  let difatSectors = 0;
  for (;;) {
    difatSectors = fatSectors > 109 ? Math.ceil((fatSectors - 109) / PER_DIFAT) : 0;
    const need = Math.ceil((content + fatSectors + difatSectors) / PER_FAT);
    if (need <= fatSectors) break;
    fatSectors = need;
  }
  const total = fatSectors + difatSectors + content;

  /* ── the allocation ───────────────────────────────────────────────────── */
  const fat = new Uint32Array(fatSectors * PER_FAT).fill(FREESECT);
  let next = 0;
  const fatAt = [];
  for (let i = 0; i < fatSectors; i++) {
    fatAt.push(next);
    fat[next++] = FATSECT;
  }
  const difatAt = [];
  for (let i = 0; i < difatSectors; i++) {
    difatAt.push(next);
    fat[next++] = DIFSECT;
  }
  const chain = (count) => {
    if (!count) return ENDOFCHAIN;
    const start = next;
    for (let i = 0; i < count; i++, next++) fat[next] = i === count - 1 ? ENDOFCHAIN : next + 1;
    return start;
  };
  const dirStart = chain(dirSectors);
  const miniFatStart = chain(miniFatSectors);
  const miniStreamStart = chain(miniStreamSectors);
  for (const node of large) node.start = chain(Math.ceil(node.data.length / SECTOR));
  root.start = miniStreamSectors ? miniStreamStart : ENDOFCHAIN;
  root.data = { length: miniStreamBytes };

  /* ── the bytes ────────────────────────────────────────────────────────── */
  const out = Buffer.alloc(SECTOR + total * SECTOR);
  const sectorAt = (n) => SECTOR + n * SECTOR;

  // Header.
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(out, 0);
  out.writeUInt16LE(0x003e, 24); // minor version
  out.writeUInt16LE(0x0003, 26); // major version 3
  out.writeUInt16LE(0xfffe, 28); // byte order
  out.writeUInt16LE(9, 30); // 512-byte sectors
  out.writeUInt16LE(6, 32); // 64-byte mini sectors
  out.writeUInt32LE(0, 40); // directory sectors: always 0 in version 3
  out.writeUInt32LE(fatSectors, 44);
  out.writeUInt32LE(dirStart, 48);
  out.writeUInt32LE(0, 52); // transaction signature
  out.writeUInt32LE(CUTOFF, 56);
  out.writeUInt32LE(miniFatSectors ? miniFatStart : ENDOFCHAIN, 60);
  out.writeUInt32LE(miniFatSectors, 64);
  out.writeUInt32LE(difatSectors ? difatAt[0] : ENDOFCHAIN, 68);
  out.writeUInt32LE(difatSectors, 72);
  for (let i = 0; i < 109; i++) out.writeUInt32LE(i < fatAt.length ? fatAt[i] : FREESECT, 76 + i * 4);

  // DIFAT sectors: the FAT sectors after the first 109, chained.
  for (let d = 0; d < difatSectors; d++) {
    const at = sectorAt(difatAt[d]);
    for (let i = 0; i < PER_DIFAT; i++) {
      const k = 109 + d * PER_DIFAT + i;
      out.writeUInt32LE(k < fatAt.length ? fatAt[k] : FREESECT, at + i * 4);
    }
    out.writeUInt32LE(d + 1 < difatSectors ? difatAt[d + 1] : ENDOFCHAIN, at + PER_DIFAT * 4);
  }

  // FAT.
  fatAt.forEach((s, i) => {
    const at = sectorAt(s);
    for (let k = 0; k < PER_FAT; k++) out.writeUInt32LE(fat[i * PER_FAT + k], at + k * 4);
  });

  // Directory: unused slots carry no name and point nowhere.
  const dirBase = sectorAt(dirStart);
  for (let i = 0; i < dirSectors * (SECTOR / 128); i++) {
    const at = dirBase + i * 128;
    const node = entries[i];
    if (!node) {
      out.writeUInt32LE(NOSTREAM, at + 68);
      out.writeUInt32LE(NOSTREAM, at + 72);
      out.writeUInt32LE(NOSTREAM, at + 76);
      continue;
    }
    const name = Buffer.from(node.name + '\0', 'utf16le');
    name.copy(out, at);
    out.writeUInt16LE(name.length, at + 64);
    out[at + 66] = node.type;
    out[at + 67] = node.index === 0 ? BLACK : node.color;
    out.writeUInt32LE(node.left, at + 68);
    out.writeUInt32LE(node.right, at + 72);
    out.writeUInt32LE(node.child, at + 76);
    if (node.type === STORAGE) {
      out.writeUInt32LE(0, at + 116);
      continue;
    }
    const size = node.type === ROOT ? miniStreamBytes : node.data.length;
    out.writeUInt32LE(size ? node.start : ENDOFCHAIN, at + 116);
    out.writeUInt32LE(size, at + 120);
  }

  // Mini FAT and the mini stream.
  if (miniSectors) {
    const miniFat = new Uint32Array(miniFatSectors * PER_FAT).fill(FREESECT);
    const base = sectorAt(miniStreamStart);
    for (const node of small) {
      const count = Math.ceil(node.data.length / MINI);
      for (let i = 0; i < count; i++) miniFat[node.start + i] = i === count - 1 ? ENDOFCHAIN : node.start + i + 1;
      Buffer.from(node.data.buffer, node.data.byteOffset, node.data.byteLength).copy(out, base + node.start * MINI);
    }
    const mfBase = sectorAt(miniFatStart);
    for (let k = 0; k < miniFat.length; k++) out.writeUInt32LE(miniFat[k], mfBase + k * 4);
  }

  // The large streams, each in its run of sectors.
  for (const node of large) {
    Buffer.from(node.data.buffer, node.data.byteOffset, node.data.byteLength).copy(out, sectorAt(node.start));
  }
  return out;
}

function toBytes(data) {
  if (data == null) return new Uint8Array(0);
  if (data instanceof Uint8Array) return data;
  if (typeof data === 'string') return Buffer.from(data, 'utf8');
  return new Uint8Array(data);
}
