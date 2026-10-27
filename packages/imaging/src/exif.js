// EXIF, and the one field that changes what you see.
//
// Orientation is the reason this exists. A phone writes every photograph in the
// sensor's own orientation and records how to turn it; a viewer that ignores
// the tag shows a third of a holiday sideways. The rest — camera, lens,
// exposure, when and where — is what the Details panel shows.
//
// The EXIF block is the same TIFF directory in every file that carries one;
// only where it sits differs: an APP1 segment in a JPEG, the file itself in a
// TIFF (and a DNG and most camera RAW), an eXIf chunk in a PNG, an EXIF chunk
// in a WebP's RIFF, and the Exif item of a HEIC or AVIF's meta box. The XMP
// packet beside it holds the rating and keywords Windows Photos and
// Lightroom write. Every file read here is untrusted: each walk is bounded
// (directories, entries, boxes, chunks, text), and a damaged block is a
// sentence (`problem`), never a hang or a throw.

const TAGS = {
  0x010e: 'description',
  0x010f: 'make',
  0x0110: 'model',
  0x0112: 'orientation',
  0x011a: 'xResolution',
  0x011b: 'yResolution',
  0x0131: 'software',
  0x0132: 'dateTime',
  0x013b: 'artist',
  0x4746: 'rating',
  0x4749: 'ratingPercent',
  0x8298: 'copyright',
  0x829a: 'exposureTime',
  0x829d: 'fNumber',
  0x8827: 'iso',
  0x9003: 'dateTimeOriginal',
  0x9004: 'dateTimeDigitized',
  0x920a: 'focalLength',
  0x9209: 'flash',
  0x9c9e: 'xpKeywords',
  0xa002: 'pixelWidth',
  0xa003: 'pixelHeight',
  0xa405: 'focalLength35mm',
  0xa434: 'lens',
  0x8825: '@gps',
  0x8769: '@exif',
};

const GPS_TAGS = {
  0x0001: 'latRef',
  0x0002: 'lat',
  0x0003: 'lonRef',
  0x0004: 'lon',
  0x0006: 'altitude',
};

/** How the image should be turned, as the operations a canvas would apply. */
export const ORIENTATION = {
  1: { rotate: 0, flip: false, label: 'as shot' },
  2: { rotate: 0, flip: true, label: 'mirrored' },
  3: { rotate: 180, flip: false, label: 'upside down' },
  4: { rotate: 180, flip: true, label: 'mirrored, upside down' },
  5: { rotate: 90, flip: true, label: 'mirrored, quarter turn' },
  6: { rotate: 90, flip: false, label: 'quarter turn clockwise' },
  7: { rotate: 270, flip: true, label: 'mirrored, three-quarter turn' },
  8: { rotate: 270, flip: false, label: 'quarter turn anticlockwise' },
};

// ── limits: what a file may ask of the reader ───────────────────────────────

/** Entries in one directory; a camera writes under a hundred. */
const MAX_ENTRIES = 1000;
/** Directories followed from IFD0 (Exif, GPS, Interop): never round in a loop. */
const MAX_DEPTH = 3;
/** One text value. */
const MAX_TEXT = 64 * 1024;
/** Boxes, chunks or segments walked in one file. */
const MAX_PARTS = 4096;
/** Items listed in a HEIF meta box. */
const MAX_ITEMS = 10000;
/** An XMP packet. */
const MAX_XMP = 2 * 1024 * 1024;
/** An EXIF block (a JPEG's APP1 is at most 64 KB; a TIFF's can be larger). */
const MAX_BLOCK = 16 * 1024 * 1024;
/** Keywords kept from one file, and the length of each. */
const MAX_KEYWORDS = 200;
const MAX_KEYWORD = 200;

const TYPE_SIZE = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8, 13: 4 };

/** More of the file is wanted than the reader was given. */
class Need {
  constructor(offset, length) {
    this.offset = offset;
    this.length = length;
  }
}

/** The file says something it cannot mean; `message` is the sentence. */
class Damage extends Error {}

/**
 * The parts of a file the reader holds: one whole file, or the pieces read
 * from it so far. A read inside the file but outside the pieces asks for
 * more (Need); a read past the file's end is damage.
 */
class Pieces {
  constructor(pieces, size, end = 'the file') {
    this.pieces = pieces;
    this.size = size;
    this.end = end;
    this.fetch = false;
  }
  bytes(o, n, what = 'a value') {
    if (!(o >= 0) || !(n >= 0) || o + n > this.size) throw new Damage(`${what} runs past the end of ${this.end}`);
    for (const p of this.pieces) {
      if (o >= p.at && o + n <= p.at + p.bytes.length) return p.bytes.subarray(o - p.at, o - p.at + n);
    }
    throw new Need(o, n);
  }
  u8(o) { return this.bytes(o, 1)[0]; }
  u16(o, le = false) { const b = this.bytes(o, 2); return le ? b[0] | (b[1] << 8) : (b[0] << 8) | b[1]; }
  u32(o, le = false) {
    const b = this.bytes(o, 4);
    return le ? (b[0] | (b[1] << 8) | (b[2] << 16)) + b[3] * 0x1000000 : b[0] * 0x1000000 + ((b[1] << 16) | (b[2] << 8) | b[3]);
  }
  u64(o) { return this.u32(o) * 0x100000000 + this.u32(o + 4); }
  sized(o, n) { return n === 0 ? 0 : n === 1 ? this.u8(o) : n === 2 ? this.u16(o) : n === 4 ? this.u32(o) : n === 8 ? this.u64(o) : 0; }
  str(o, n) { return String.fromCharCode(...this.bytes(o, n)); }
}

const asBytes = (bytes) => (bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes || []));
const decodeText = (bytes, label = 'utf-8') => new TextDecoder(label, { fatal: false }).decode(bytes);

// ── where the EXIF and the XMP sit ──────────────────────────────────────────

const XMP_JPEG = 'http://ns.adobe.com/xap/1.0/\0';
const isTiffHeader = (src, at) => {
  const a = src.u8(at);
  const b = src.u8(at + 1);
  if (a === 0x49 && b === 0x49) return src.u16(at + 2, true) === 42;
  if (a === 0x4d && b === 0x4d) return src.u16(at + 2, false) === 42;
  return false;
};

const HEIF_BRANDS = new Set(['heic', 'heix', 'heim', 'heis', 'hevc', 'hevx', 'mif1', 'mif2', 'msf1', 'avif', 'avis', 'avio', 'miaf']);

/** What kind of file this is, from its first bytes: 'jpeg', 'tiff', 'png', 'webp', 'heif' or null. */
export function containerOf(bytes) {
  const b = asBytes(bytes);
  if (b[0] === 0xff && b[1] === 0xd8) return 'jpeg';
  if ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 42 && b[3] === 0) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0 && b[3] === 42)) return 'tiff';
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png';
  if (String.fromCharCode(...b.subarray(0, 4)) === 'RIFF' && String.fromCharCode(...b.subarray(8, 12)) === 'WEBP') return 'webp';
  if (String.fromCharCode(...b.subarray(4, 8)) === 'ftyp') {
    // A video is in the same boxes: only a still picture's brands make it HEIF.
    const end = Math.min(b.length, ((b[0] << 24) | (b[1] << 16) | (b[2] << 8) | b[3]) >>> 0);
    for (let at = 8; at + 4 <= end; at += at === 8 ? 8 : 4) {
      if (HEIF_BRANDS.has(String.fromCharCode(...b.subarray(at, at + 4)))) return 'heif';
    }
  }
  return null;
}

function locateJpeg(src) {
  const out = { exif: null, xmp: null };
  let at = 2;
  for (let n = 0; n < MAX_PARTS && at + 4 <= src.size; n++) {
    if (src.u8(at) !== 0xff) break; // not a marker: the walk is lost, and what was found stands
    const marker = src.u8(at + 1);
    if (marker === 0xff) { at += 1; continue; } // fill
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { at += 2; continue; }
    if (marker === 0xda || marker === 0xd9) break; // the picture itself: nothing above it is left
    const length = src.u16(at + 2);
    if (length < 2) throw new Damage('a JPEG segment has an impossible length');
    if (marker === 0xe1 && length >= 8) {
      const head = src.bytes(at + 4, Math.min(length - 2, 29), 'a JPEG segment');
      const sig = String.fromCharCode(...head);
      if (!out.exif && sig.startsWith('Exif\0')) out.exif = { at: at + 10, length: length - 8 };
      else if (!out.xmp && sig === XMP_JPEG && length > 31) out.xmp = { at: at + 4 + 29, length: length - 2 - 29 };
    }
    at += 2 + length;
  }
  return out;
}

function locatePng(src) {
  const out = { exif: null, xmp: null };
  let at = 8;
  for (let n = 0; n < MAX_PARTS * 4 && at + 12 <= src.size; n++) {
    const length = src.u32(at);
    const type = src.str(at + 4, 4);
    if (at + 12 + length > src.size) throw new Damage(`the PNG chunk ${type.replace(/[^\x20-\x7e]/g, '?')} runs past the end of the file`);
    if (type === 'eXIf' && !out.exif) {
      // Some writers lead with the JPEG's "Exif\0\0"; the chunk is the TIFF block itself.
      const lead = length >= 6 && src.str(at + 8, 6) === 'Exif\0\0' ? 6 : 0;
      out.exif = { at: at + 8 + lead, length: length - lead };
    } else if (type === 'iTXt' && !out.xmp && length > 22) {
      const head = src.bytes(at + 8, Math.min(length, 200));
      if (String.fromCharCode(...head.subarray(0, 18)) === 'XML:com.adobe.xmp\0' && head[18] === 0) {
        // Not compressed: language tag and translated keyword, each ending in a nul, then the text.
        let i = 20;
        let nuls = 0;
        while (i < head.length && nuls < 2) { if (head[i] === 0) nuls += 1; i += 1; }
        if (nuls === 2) out.xmp = { at: at + 8 + i, length: length - i };
      }
    } else if (type === 'IEND') break;
    at += 12 + length;
  }
  return out;
}

function locateWebp(src) {
  const out = { exif: null, xmp: null };
  const claimed = 8 + src.u32(4, true);
  if (claimed > src.size) out.short = true;
  const end = Math.min(src.size, claimed);
  let at = 12;
  for (let n = 0; n < MAX_PARTS && at + 8 <= end; n++) {
    const type = src.str(at, 4);
    const length = src.u32(at + 4, true);
    if (at + 8 + length > end) throw new Damage(`the WebP chunk ${type.replace(/[^\x20-\x7e]/g, '?').trim()} runs past the end of the file`);
    if (type === 'EXIF' && !out.exif) {
      const lead = length >= 6 && src.str(at + 8, 6) === 'Exif\0\0' ? 6 : 0;
      out.exif = { at: at + 8 + lead, length: length - lead };
    } else if (type === 'XMP ' && !out.xmp) {
      out.xmp = { at: at + 8, length };
    }
    at += 8 + length + (length & 1);
  }
  return out;
}

/** The boxes inside [start, end): { type, at (the box), body (its contents), end }. */
function boxes(src, start, end) {
  const list = [];
  let at = start;
  while (at + 8 <= end) {
    if (list.length >= MAX_PARTS) throw new Damage('the file has more boxes than any picture needs');
    let size = src.u32(at);
    const type = src.str(at + 4, 4);
    let body = at + 8;
    if (size === 1) { size = src.u64(at + 8); body = at + 16; } else if (size === 0) size = end - at;
    if (size < body - at || at + size > end) throw new Damage(`the box ${type.replace(/[^\x20-\x7e]/g, '?')} runs past the end of its parent`);
    list.push({ type, at, body, end: at + size });
    at += size;
  }
  return list;
}

function locateHeif(src) {
  const out = { exif: null, xmp: null };
  const meta = boxes(src, 0, src.size).find((b) => b.type === 'meta');
  if (!meta) return out;
  const kids = boxes(src, meta.body + 4, meta.end); // a full box: version and flags first
  const iinf = kids.find((b) => b.type === 'iinf');
  const iloc = kids.find((b) => b.type === 'iloc');
  const idat = kids.find((b) => b.type === 'idat');
  if (!iinf || !iloc) return out;

  // The items: which is the Exif, which the XMP.
  let exifId = null;
  let xmpId = null;
  const infeVersion = src.u8(iinf.body);
  const infes = boxes(src, iinf.body + 4 + (infeVersion === 0 ? 2 : 4), iinf.end);
  for (const infe of infes) {
    if (infe.type !== 'infe') continue;
    const version = src.u8(infe.body);
    if (version < 2) continue;
    let p = infe.body + 4;
    const id = version === 2 ? src.u16(p) : src.u32(p);
    p += version === 2 ? 2 : 4;
    p += 2; // protection index
    const type = src.str(p, 4);
    p += 4;
    if (type === 'Exif' && exifId == null) exifId = id;
    if (type === 'mime' && xmpId == null) {
      // The content type, nul-terminated, after the item's name.
      const rest = src.bytes(p, Math.min(256, infe.end - p));
      const nameEnd = rest.indexOf(0);
      const typeEnd = nameEnd < 0 ? -1 : rest.indexOf(0, nameEnd + 1);
      if (typeEnd > nameEnd && String.fromCharCode(...rest.subarray(nameEnd + 1, typeEnd)) === 'application/rdf+xml') xmpId = id;
    }
  }
  if (exifId == null && xmpId == null) return out;

  // Where each item is.
  const version = src.u8(iloc.body);
  let p = iloc.body + 4;
  const sizes = src.u8(p);
  const sizes2 = src.u8(p + 1);
  p += 2;
  const offsetSize = sizes >> 4;
  const lengthSize = sizes & 15;
  const baseSize = sizes2 >> 4;
  const indexSize = version === 1 || version === 2 ? sizes2 & 15 : 0;
  if (![0, 4, 8].includes(offsetSize) || ![0, 4, 8].includes(lengthSize) || ![0, 4, 8].includes(baseSize)) throw new Damage('the item locations use field sizes no file uses');
  const count = version < 2 ? src.u16(p) : src.u32(p);
  p += version < 2 ? 2 : 4;
  if (count > MAX_ITEMS) throw new Damage('the file lists more items than any picture needs');
  for (let i = 0; i < count; i++) {
    const id = version < 2 ? src.u16(p) : src.u32(p);
    p += version < 2 ? 2 : 4;
    let method = 0;
    if (version === 1 || version === 2) { method = src.u16(p) & 15; p += 2; }
    p += 2; // data reference index
    const base = src.sized(p, baseSize);
    p += baseSize;
    const extents = src.u16(p);
    p += 2;
    let first = null;
    for (let e = 0; e < extents; e++) {
      p += indexSize;
      const offset = src.sized(p, offsetSize);
      p += offsetSize;
      const length = src.sized(p, lengthSize);
      p += lengthSize;
      if (e === 0) first = { offset, length };
    }
    if (p > iloc.end) throw new Damage('the item locations run past their box');
    if ((id !== exifId && id !== xmpId) || extents !== 1 || !first) continue;
    let at;
    if (method === 0) at = base + first.offset;
    else if (method === 1 && idat) at = idat.body + base + first.offset;
    else continue;
    const length = first.length || src.size - at;
    if (at + length > src.size) throw new Damage('an item runs past the end of the file');
    if (id === exifId) {
      // The item opens with how far past the next four bytes the TIFF header is.
      const skip = src.u32(at);
      if (skip + 4 > length) throw new Damage('the Exif item points past its own end');
      out.exif = { at: at + 4 + skip, length: length - 4 - skip };
    } else {
      out.xmp = { at, length };
    }
  }
  return out;
}

/**
 * Where a file keeps its EXIF block and its XMP packet, without reading
 * either: `{ container, exif: { at, length } | null, xmp: { at, length } | null }`,
 * each `at` the file offset of the TIFF header (of the packet's text). A
 * TIFF's XMP is found by readMetadata, in its first directory.
 */
function locate(src) {
  const head = src.bytes(0, Math.min(src.size, 64));
  const container = containerOf(head);
  let found = { exif: null, xmp: null };
  if (container === 'jpeg') found = locateJpeg(src);
  else if (container === 'tiff') found = { exif: { at: 0, length: src.size }, xmp: null };
  else if (container === 'png') found = locatePng(src);
  else if (container === 'webp') found = locateWebp(src);
  else if (container === 'heif') found = locateHeif(src);
  return { container, ...found };
}

// ── the TIFF directory ──────────────────────────────────────────────────────

/** One value of an entry, as JavaScript: text, a number, or a list of numbers. */
function valueOf(src, type, count, offset, little) {
  if (type === 2) {
    const bytes = src.bytes(offset, Math.min(count, MAX_TEXT));
    return decodeText(bytes).replace(/\0+$/, '').trim();
  }
  if (type === 1 || type === 7) return src.bytes(offset, Math.min(count, MAX_TEXT));
  if (type === 3 || type === 8) {
    const v = src.u16(offset, little);
    return type === 8 && v > 0x7fff ? v - 0x10000 : v;
  }
  if (type === 4 || type === 9) {
    const v = src.u32(offset, little);
    return type === 9 && v > 0x7fffffff ? v - 0x100000000 : v;
  }
  if (type === 5 || type === 10) {
    const parts = [];
    for (let n = 0; n < Math.min(count, 3); n++) {
      let numerator = src.u32(offset + n * 8, little);
      let denominator = src.u32(offset + n * 8 + 4, little);
      if (type === 10) {
        if (numerator > 0x7fffffff) numerator -= 0x100000000;
        if (denominator > 0x7fffffff) denominator -= 0x100000000;
      }
      parts.push(denominator ? numerator / denominator : 0);
    }
    return count > 1 ? parts : parts[0];
  }
  return null;
}

/** A directory's entries: `[{ tag, type, count, offset (of the value, in the file), entry (its 12 bytes) }]` and where the next directory is. */
function entriesOf(src, base, ifd, little, what = 'a directory') {
  const at = base + ifd;
  const count = src.u16(at, little);
  if (count > MAX_ENTRIES) throw new Damage(`${what} claims ${count} entries`);
  src.bytes(at, 2 + count * 12, what);
  const list = [];
  for (let i = 0; i < count; i++) {
    const entry = at + 2 + i * 12;
    const tag = src.u16(entry, little);
    const type = src.u16(entry + 2, little);
    const n = src.u32(entry + 4, little);
    const size = (TYPE_SIZE[type] || 0) * n;
    const offset = size > 4 ? base + src.u32(entry + 8, little) : entry + 8;
    list.push({ tag, type, count: n, size, offset, entry, pointer: src.u32(entry + 8, little) });
  }
  // Some writers leave off the last directory's next-pointer.
  const tail = at + 2 + count * 12;
  return { list, next: tail + 4 <= src.size ? src.u32(tail, little) : 0, at };
}

function readIfd(src, base, ifd, little, into, tagMap, depth, seen) {
  if (depth > MAX_DEPTH || seen.has(ifd)) return into;
  seen.add(ifd);
  const { list } = entriesOf(src, base, ifd, little, depth ? 'an inner EXIF directory' : 'the first EXIF directory');
  for (const e of list) {
    if (e.tag === 0x02bc && depth === 0 && !into['@xmp'] && (e.type === 1 || e.type === 7) && e.count > 0) {
      // A TIFF's XMP packet, in its first directory.
      into['@xmp'] = { at: e.offset, length: e.count };
      continue;
    }
    const name = tagMap[e.tag];
    if (!name || !TYPE_SIZE[e.type]) continue;
    if (name === '@exif' || name === '@gps') {
      try {
        readIfd(src, base, e.pointer, little, into, name === '@gps' ? GPS_TAGS : TAGS, depth + 1, seen);
      } catch (err) {
        if (err instanceof Need) throw err;
        // One broken inner directory: what IFD0 said still stands.
        into.problem = into.problem || sentence(err);
      }
      continue;
    }
    let value;
    try {
      value = valueOf(src, e.type, e.count, e.offset, little);
    } catch (err) {
      if (err instanceof Need) throw err;
      continue; // one value past the end: the others still read
    }
    if (name === 'xpKeywords') value = value instanceof Uint8Array ? decodeText(value, 'utf-16le').replace(/\0+$/, '').trim() : null;
    else if (value instanceof Uint8Array) value = null;
    if (value !== null && value !== '' && value !== undefined) into[name] = value;
  }
  return into;
}

const sentence = (err) => `The EXIF block of this file is damaged: ${err?.message || 'it cannot be read'}.`;

/** The TIFF block at `at` read into fields. */
function readTiff(src, at, length) {
  if (length < 8) throw new Damage('it is too short to hold a TIFF header');
  if (!isTiffHeader(src, at)) throw new Damage('it does not begin with a TIFF header');
  const little = src.u8(at) === 0x49;
  const ifd = src.u32(at + 4, little);
  if (ifd < 8 || ifd + 2 > length) throw new Damage('its first directory lies outside it');
  const window = new Pieces(src.pieces, Math.min(src.size, at + length), 'the block');
  return readIfd(window, at, ifd, little, {}, TAGS, 0, new Set());
}

function finish(out) {
  if (out.lat != null && out.latRef) {
    const toDegrees = (parts) => (Array.isArray(parts) ? parts[0] + (parts[1] || 0) / 60 + (parts[2] || 0) / 3600 : parts);
    const lat = toDegrees(out.lat) * (String(out.latRef).toUpperCase() === 'S' ? -1 : 1);
    const lon = toDegrees(out.lon) * (String(out.lonRef).toUpperCase() === 'W' ? -1 : 1);
    if (Number.isFinite(lat) && Number.isFinite(lon) && Math.abs(lat) <= 90 && Math.abs(lon) <= 180) {
      out.latitude = lat;
      out.longitude = lon;
    }
  }
  return out;
}

// ── XMP: the rating and the keywords ────────────────────────────────────────

const NS = {
  xmp: 'http://ns.adobe.com/xap/1.0/',
  dc: 'http://purl.org/dc/elements/1.1/',
  ms: 'http://ns.microsoft.com/photo/1.0/',
};

const unescapeXml = (s) => s
  .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
  .replace(/&#x([0-9a-f]+);/gi, (_m, h) => { const c = parseInt(h, 16); return c > 0 && c < 0x110000 ? String.fromCodePoint(c) : ''; })
  .replace(/&#(\d+);/g, (_m, d) => { const c = Number(d); return c > 0 && c < 0x110000 ? String.fromCodePoint(c) : ''; })
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** The prefixes a packet gives a namespace (its usual one first). */
function prefixes(text, uri, usual) {
  const out = new Set([usual]);
  const re = new RegExp(`xmlns:([A-Za-z_][\\w.-]*)\\s*=\\s*["']${escapeRe(uri)}["']`, 'g');
  let m;
  for (let n = 0; n < 20 && (m = re.exec(text)); n++) out.add(m[1]);
  return [...out];
}

/** A simple property: as an attribute (xmp:Rating="3") or an element (<xmp:Rating>3</xmp:Rating>). */
function simple(text, prefix, name) {
  const p = escapeRe(prefix);
  const attr = new RegExp(`\\s${p}:${name}\\s*=\\s*(["'])([^"']*)\\1`).exec(text);
  if (attr) return unescapeXml(attr[2]).trim();
  const el = new RegExp(`<${p}:${name}(?:\\s[^>]*)?>([^<]*)</${p}:${name}>`).exec(text);
  return el ? unescapeXml(el[1]).trim() : null;
}

/** The items of a list property (<dc:subject><rdf:Bag><rdf:li>…). */
function list(text, prefix, name) {
  const p = escapeRe(prefix);
  const block = new RegExp(`<${p}:${name}(?:\\s[^>]*)?>([\\s\\S]*?)</${p}:${name}>`).exec(text);
  if (!block) return [];
  const out = [];
  const re = /<rdf:li(?:\s[^>]*)?>([\s\S]*?)<\/rdf:li>/g;
  let m;
  while ((m = re.exec(block[1])) && out.length < MAX_KEYWORDS) out.push(unescapeXml(m[1]).trim());
  return out;
}

/**
 * The rating and keywords of an XMP packet, as Windows Photos, Explorer and
 * Lightroom write them: `{ rating, rejected, ratingPercent, keywords }`.
 * A rating is 0 to 5 (Lightroom's -1 is a rejected picture); keywords are
 * dc:subject's, with Windows' own copy of them beside it.
 */
export function readXmp(input) {
  const raw = typeof input === 'string' ? input : decodeText(asBytes(input).subarray(0, MAX_XMP));
  const text = raw.slice(0, MAX_XMP);
  const out = {};
  for (const prefix of prefixes(text, NS.xmp, 'xmp')) {
    const value = simple(text, prefix, 'Rating');
    if (value == null || value === '') continue;
    const n = Number(value);
    if (n === -1) out.rejected = true;
    else if (Number.isFinite(n) && n >= 0 && n <= 5) out.rating = Math.round(n);
    break;
  }
  for (const prefix of prefixes(text, NS.ms, 'MicrosoftPhoto')) {
    const value = simple(text, prefix, 'Rating');
    if (value != null && Number.isFinite(Number(value))) { out.ratingPercent = Number(value); break; }
  }
  const keywords = [];
  for (const prefix of prefixes(text, NS.dc, 'dc')) keywords.push(...list(text, prefix, 'subject'));
  for (const prefix of prefixes(text, NS.ms, 'MicrosoftPhoto')) keywords.push(...list(text, prefix, 'LastKeywordXMP'));
  const kept = cleanKeywords(keywords);
  if (kept.length) out.keywords = kept;
  return out;
}

function cleanKeywords(words) {
  const seen = new Set();
  const out = [];
  for (const w of words) {
    const word = String(w || '').replace(/[\u0000-\u001f]/g, '').trim().slice(0, MAX_KEYWORD);
    if (!word || seen.has(word.toLowerCase())) continue;
    seen.add(word.toLowerCase());
    out.push(word);
    if (out.length >= MAX_KEYWORDS) break;
  }
  return out;
}

/** Windows' percent rating as stars, the way Explorer turns one into the other. */
export function starsFromPercent(p) {
  const n = Number(p);
  if (!Number.isFinite(n) || n <= 0) return 0;
  if (n < 13) return 1;
  if (n < 38) return 2;
  if (n < 63) return 3;
  if (n < 88) return 4;
  return 5;
}

// ── reading ─────────────────────────────────────────────────────────────────

function readFrom(src) {
  const place = locate(src);
  let exif = {};
  let problem = null;
  let partial = false;
  if (place.exif) {
    if (place.exif.length > MAX_BLOCK) problem = 'The EXIF block of this file is larger than any camera writes, and was not read.';
    else {
      try {
        exif = finish(readTiff(src, place.exif.at, place.exif.length));
      } catch (err) {
        if (err instanceof Need) {
          if (src.fetch) throw err;
          partial = true;
        } else problem = sentence(err);
        exif = {};
      }
    }
  }
  if (exif.problem) { problem = problem || exif.problem; delete exif.problem; }
  if (!place.exif && place.short) problem = 'The file is cut short: it ends before its header says it does, and its EXIF may be in the part that is missing.';
  const xmpAt = place.xmp || exif['@xmp'] || null;
  delete exif['@xmp'];
  let xmp = {};
  if (xmpAt && xmpAt.length > 0) {
    try {
      xmp = readXmp(src.bytes(xmpAt.at, Math.min(xmpAt.length, MAX_XMP), 'the XMP packet'));
    } catch (err) {
      if (err instanceof Need) {
        if (src.fetch) throw err;
        partial = true;
      } else problem = problem || 'The XMP packet of this file runs past its end, and was not read.';
    }
  }
  return { container: place.container, exif, xmp, problem, ...(partial ? { partial } : {}) };
}

/**
 * Everything this reads from a file: `{ container, exif, xmp, marks, problem }`.
 * `bytes` is the whole file, or its start when `size` says the file is
 * longer: what lies past the bytes given is then simply not read (use
 * readMetadataFrom to fetch it). `problem` is a sentence when the EXIF is
 * damaged; the fields read before the damage still stand.
 */
export function readMetadata(bytes, { size } = {}) {
  const b = asBytes(bytes);
  const src = new Pieces([{ at: 0, bytes: b }], Math.max(b.length, size || 0));
  try {
    return withMarks(readFrom(src));
  } catch (err) {
    if (err instanceof Need) return withMarks({ container: containerOf(b), exif: {}, xmp: {}, problem: null, partial: true });
    return withMarks({ container: containerOf(b), exif: {}, xmp: {}, problem: err instanceof Damage ? sentence(err) : 'The EXIF block of this file could not be read.' });
  }
}

/**
 * The same, reading only the parts it needs: `read(offset, length)` answers
 * a promise of those bytes (fewer at the file's end), `size` is the file's.
 * A WebP keeps its EXIF after the picture and a TIFF its directory where it
 * likes; this reads the start, then each place the file points to, a few
 * reads in all and never the picture itself.
 */
export async function readMetadataFrom(read, size, { first = 256 * 1024, chunk = 256 * 1024, rounds = 24 } = {}) {
  const pieces = [{ at: 0, bytes: asBytes(await read(0, Math.min(size, first))) }];
  const src = new Pieces(pieces, size);
  src.fetch = true;
  for (let i = 0; i < rounds; i++) {
    try {
      return withMarks(readFrom(src));
    } catch (err) {
      if (err instanceof Need) {
        const length = Math.min(size - err.offset, Math.max(err.length, chunk));
        const bytes = asBytes(await read(err.offset, length));
        if (bytes.length < err.length) return withMarks({ container: containerOf(pieces[0].bytes), exif: {}, xmp: {}, problem: 'The file ended before its EXIF block did.' });
        // Fewer bytes than the file's size promised: it is shorter than it says, and a read past here is damage.
        if (bytes.length < length) src.size = err.offset + bytes.length;
        pieces.push({ at: err.offset, bytes });
        continue;
      }
      return withMarks({ container: containerOf(pieces[0].bytes), exif: {}, xmp: {}, problem: err instanceof Damage ? sentence(err) : 'The EXIF block of this file could not be read.' });
    }
  }
  return withMarks({ container: containerOf(pieces[0].bytes), exif: {}, xmp: {}, problem: null, partial: true });
}

/**
 * The stars and tags the file itself carries, from whichever of its places
 * has them: XMP's rating first (what Lightroom and Windows Photos write),
 * then EXIF's Rating, then its RatingPercent; the keywords of XMP and of
 * Windows' XPKeywords together. `{ rating: 0..5 | null, rejected, tags }`.
 */
export function fileMarks({ exif = {}, xmp = {} } = {}) {
  let rating = null;
  if (xmp.rating != null) rating = xmp.rating;
  else if (Number.isFinite(exif.rating) && exif.rating >= 0 && exif.rating <= 5) rating = exif.rating;
  else if (exif.ratingPercent != null) rating = starsFromPercent(exif.ratingPercent);
  else if (xmp.ratingPercent != null) rating = starsFromPercent(xmp.ratingPercent);
  const tags = cleanKeywords([...(xmp.keywords || []), ...String(exif.xpKeywords || '').split(';')]);
  return { rating, rejected: Boolean(xmp.rejected), tags };
}

function withMarks(meta) {
  return { ...meta, marks: fileMarks(meta) };
}

/**
 * Read the EXIF block of a JPEG, TIFF, PNG, WebP, HEIC or AVIF.
 * @returns {object} the fields that were found; `{}` when there are none, and
 *   `problem` (a sentence) beside whatever was read when the block is damaged
 */
export function readExif(bytes, options) {
  const meta = readMetadata(bytes, options);
  return meta.problem ? { ...meta.exif, problem: meta.problem } : meta.exif;
}

/** The details panel's rows: only what is there, phrased the way a camera would. */
export function describeExif(exif) {
  const rows = [];
  const add = (label, value) => value != null && value !== '' && rows.push({ label, value: String(value) });

  add('Camera', [exif.make, exif.model].filter(Boolean).join(' ').trim());
  add('Lens', exif.lens);
  if (exif.exposureTime) {
    const t = exif.exposureTime;
    add('Exposure', t >= 1 ? `${t}s` : `1/${Math.round(1 / t)}s`);
  }
  if (exif.fNumber) add('Aperture', `f/${Math.round(exif.fNumber * 10) / 10}`);
  add('ISO', exif.iso);
  if (exif.focalLength) add('Focal length', `${Math.round(exif.focalLength)}mm${exif.focalLength35mm ? ` (${exif.focalLength35mm}mm equivalent)` : ''}`);
  const taken = exif.dateTimeOriginal || exif.dateTime;
  if (taken) add('Taken', String(taken).replace(/^(\d{4}):(\d{2}):(\d{2})/, '$1-$2-$3'));
  if (exif.latitude != null) add('Location', `${exif.latitude.toFixed(5)}, ${exif.longitude.toFixed(5)}`);
  add('Software', exif.software);
  add('Artist', exif.artist);
  add('Copyright', exif.copyright);
  if (exif.orientation && exif.orientation !== 1) add('Orientation', ORIENTATION[exif.orientation]?.label);
  add('Metadata', exif.problem);
  return rows;
}

/** The orientation a quarter turn clockwise more makes, whatever turn and mirror one says already. */
export const TURN_CLOCKWISE = { 1: 6, 6: 3, 3: 8, 8: 1, 2: 7, 7: 4, 4: 5, 5: 2 };

/**
 * A JPEG with its EXIF orientation set to `orientation`, its pixels left
 * exactly as they were — the lossless turn Windows Photos makes. The tag is
 * rewritten where the file has one; a file with no EXIF gets a small block
 * holding only it, after its JFIF header. Null where the file is not a JPEG
 * or has EXIF without an orientation tag, which cannot be added without
 * moving everything after it.
 */
export function withOrientation(bytes, orientation) {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (!(orientation >= 1 && orientation <= 8)) throw new Error('an orientation is 1 to 8');
  if (b.length < 4 || b[0] !== 0xff || b[1] !== 0xd8) return null;
  let at = 2;
  let afterHeader = 2;
  while (at + 4 <= b.length && b[at] === 0xff) {
    const marker = b[at + 1];
    const length = (b[at + 2] << 8) | b[at + 3];
    if (length < 2) return null;
    if (marker === 0xe0 && at === afterHeader) afterHeader = at + 2 + length;
    if (marker === 0xe1 && at + 10 <= b.length && String.fromCharCode(b[at + 4], b[at + 5], b[at + 6], b[at + 7]) === 'Exif') {
      const tiff = at + 10;
      const end = Math.min(b.length, at + 2 + length);
      if (tiff + 8 > end) return null;
      const view = new DataView(b.buffer, b.byteOffset, b.byteLength);
      const little = b[tiff] === 0x49;
      const ifd = tiff + view.getUint32(tiff + 4, little);
      if (ifd + 2 > end) return null;
      const count = view.getUint16(ifd, little);
      for (let i = 0; i < count; i++) {
        const entry = ifd + 2 + i * 12;
        if (entry + 12 > end) return null;
        if (view.getUint16(entry, little) !== 0x0112) continue;
        if (view.getUint16(entry + 2, little) !== 3) return null;
        const out = new Uint8Array(b);
        new DataView(out.buffer).setUint16(entry + 8, orientation, little);
        return out;
      }
      return null;
    }
    if (marker === 0xda) break;
    at += 2 + length;
  }
  // No EXIF at all: one holding the orientation and nothing else.
  const tiff = new Uint8Array(26);
  const tv = new DataView(tiff.buffer);
  tiff.set([0x49, 0x49], 0);
  tv.setUint16(2, 42, true);
  tv.setUint32(4, 8, true);
  tv.setUint16(8, 1, true);
  tv.setUint16(10, 0x0112, true);
  tv.setUint16(12, 3, true);
  tv.setUint32(14, 1, true);
  tv.setUint16(18, orientation, true);
  tv.setUint32(22, 0, true);
  const payload = new Uint8Array(6 + tiff.length);
  payload.set([0x45, 0x78, 0x69, 0x66, 0, 0], 0);
  payload.set(tiff, 6);
  const app1 = new Uint8Array(4 + payload.length);
  app1.set([0xff, 0xe1, (payload.length + 2) >> 8, (payload.length + 2) & 0xff], 0);
  app1.set(payload, 4);
  const out = new Uint8Array(b.length + app1.length);
  out.set(b.subarray(0, afterHeader), 0);
  out.set(app1, afterHeader);
  out.set(b.subarray(afterHeader), afterHeader + app1.length);
  return out;
}

/** What the viewer must do to show the image the right way up. */
export function orientationOf(exif) {
  return ORIENTATION[exif?.orientation] || ORIENTATION[1];
}
