'use strict';

/**
 * TrueType fonts, read far enough to embed one.
 *
 * The base-14 fonts know Latin-1 and nothing else, so a document in Arabic,
 * Hebrew, Greek or Russian printed as rows of '?'. The answer is the one
 * every PDF writer reaches in the end: a real font, embedded, each glyph
 * written by its number. This file reads what that needs from a .ttf (or one
 * face of a .ttc) — the metrics, the character map, the glyph outlines —
 * and writes a SUBSET: the same font with every glyph the page does not use
 * emptied, so a page of Arabic carries tens of kilobytes of font rather than
 * the megabyte the system keeps.
 *
 * Glyph numbers are kept as they were. The PDF names glyphs by number
 * (`/CIDToGIDMap /Identity`), so an emptied glyph is cheaper than a
 * renumbered one: nothing that points at a glyph has to change.
 *
 * Fonts with CFF outlines (OpenType 'OTTO') are not read: their glyphs are
 * a different machine, and every system this runs on has a TrueType font
 * that covers the scripts in question.
 */

const u16 = (b, o) => b.readUInt16BE(o);
const i16 = (b, o) => b.readInt16BE(o);
const u32 = (b, o) => b.readUInt32BE(o);

/** The table directory of the face starting at `base`. */
function tablesAt(buf, base) {
  const count = u16(buf, base + 4);
  const tables = {};
  for (let i = 0; i < count; i++) {
    const at = base + 12 + 16 * i;
    tables[buf.toString('latin1', at, at + 4)] = { offset: u32(buf, at + 8), length: u32(buf, at + 12) };
  }
  return tables;
}

/** A name from the 'name' table — the family (1), the full name (4) or the PostScript name (6). */
function nameOf(buf, tables, id) {
  const t = tables.name;
  if (!t) return null;
  const count = u16(buf, t.offset + 2);
  const strings = t.offset + u16(buf, t.offset + 4);
  let fallback = null;
  for (let i = 0; i < count; i++) {
    const at = t.offset + 6 + 12 * i;
    const platform = u16(buf, at);
    const nameId = u16(buf, at + 6);
    if (nameId !== id) continue;
    const length = u16(buf, at + 8);
    const start = strings + u16(buf, at + 10);
    if (platform === 3 || platform === 0) {
      let s = '';
      for (let j = 0; j + 1 < length; j += 2) s += String.fromCharCode(u16(buf, start + j));
      return s;
    }
    if (platform === 1 && fallback === null) fallback = buf.toString('latin1', start, start + length);
  }
  return fallback;
}

/** A lookup from code point to glyph number, from the best Unicode subtable the font has. */
function readCmap(buf, tables) {
  const t = tables.cmap;
  if (!t) return () => 0;
  const count = u16(buf, t.offset + 2);
  let format4 = null;
  let format12 = null;
  for (let i = 0; i < count; i++) {
    const at = t.offset + 4 + 8 * i;
    const platform = u16(buf, at);
    const encoding = u16(buf, at + 2);
    const sub = t.offset + u32(buf, at + 4);
    const format = u16(buf, sub);
    if (format === 12 && ((platform === 3 && encoding === 10) || platform === 0)) format12 = sub;
    if (format === 4 && ((platform === 3 && (encoding === 1 || encoding === 0)) || platform === 0)) format4 = format4 ?? sub;
  }
  if (format12 !== null) {
    const groups = u32(buf, format12 + 12);
    return (cp) => {
      let lo = 0;
      let hi = groups - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        const at = format12 + 16 + 12 * mid;
        const start = u32(buf, at);
        const end = u32(buf, at + 4);
        if (cp < start) hi = mid - 1;
        else if (cp > end) lo = mid + 1;
        else return u32(buf, at + 8) + (cp - start);
      }
      return 0;
    };
  }
  if (format4 !== null) {
    const segs = u16(buf, format4 + 6) / 2;
    const ends = format4 + 14;
    const starts = ends + 2 * segs + 2;
    const deltas = starts + 2 * segs;
    const ranges = deltas + 2 * segs;
    return (cp) => {
      if (cp > 0xFFFF) return 0;
      let lo = 0;
      let hi = segs - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        const end = u16(buf, ends + 2 * mid);
        const start = u16(buf, starts + 2 * mid);
        if (cp > end) lo = mid + 1;
        else if (cp < start) hi = mid - 1;
        else {
          const range = u16(buf, ranges + 2 * mid);
          const delta = i16(buf, deltas + 2 * mid);
          if (range === 0) return (cp + delta) & 0xFFFF;
          const g = u16(buf, ranges + 2 * mid + range + 2 * (cp - start));
          return g === 0 ? 0 : (g + delta) & 0xFFFF;
        }
      }
      return 0;
    };
  }
  return () => 0;
}

/**
 * Read a font. `face` picks one face of a collection by its family or full
 * name ('Segoe UI', 'Arial Bold'); without one, the first.
 */
function readTrueType(bytes, { face = null } = {}) {
  const buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
  let base = 0;
  if (buf.toString('latin1', 0, 4) === 'ttcf') {
    const count = u32(buf, 8);
    const bases = Array.from({ length: count }, (_, i) => u32(buf, 12 + 4 * i));
    base = bases[0];
    if (face) {
      const want = String(face).toLowerCase();
      const hit = bases.find((b) => {
        const t = tablesAt(buf, b);
        return [1, 4].some((id) => String(nameOf(buf, t, id) || '').toLowerCase() === want);
      });
      if (hit !== undefined) base = hit;
    }
  }
  const tag = buf.toString('latin1', base, base + 4);
  if (tag === 'OTTO') throw new Error('this font keeps CFF outlines, which this writer does not embed');
  if (u32(buf, base) !== 0x00010000 && tag !== 'true') throw new Error('not a TrueType font');
  const tables = tablesAt(buf, base);
  for (const need of ['head', 'hhea', 'maxp', 'hmtx', 'loca', 'glyf', 'cmap']) {
    if (!tables[need]) throw new Error(`the font has no ${need} table`);
  }
  const head = tables.head.offset;
  const hhea = tables.hhea.offset;
  const unitsPerEm = u16(buf, head + 18);
  const numGlyphs = u16(buf, tables.maxp.offset + 4);
  const numberOfHMetrics = u16(buf, hhea + 34);
  const longLoca = i16(buf, head + 50) === 1;
  const os2 = tables['OS/2'] ? tables['OS/2'].offset : null;
  const os2Version = os2 !== null ? u16(buf, os2) : 0;
  const font = {
    buf,
    tables,
    unitsPerEm,
    numGlyphs,
    postscriptName: (nameOf(buf, tables, 6) || nameOf(buf, tables, 4) || 'Font').replace(/[^A-Za-z0-9+-]/g, ''),
    family: nameOf(buf, tables, 1) || '',
    bbox: [i16(buf, head + 36), i16(buf, head + 38), i16(buf, head + 40), i16(buf, head + 42)],
    ascent: os2 !== null ? i16(buf, os2 + 68) : i16(buf, hhea + 4),
    descent: os2 !== null ? i16(buf, os2 + 70) : i16(buf, hhea + 6),
    capHeight: os2 !== null && os2Version >= 2 ? i16(buf, os2 + 88) : Math.round(i16(buf, hhea + 4) * 0.7),
    italicAngle: tables.post ? buf.readInt32BE(tables.post.offset + 4) / 65536 : 0,
    // OS/2 fsType: 0x0002 is "restricted licence embedding" — the font may not be embedded at all.
    fsType: os2 !== null ? u16(buf, os2 + 8) : 0,
    glyphOf: readCmap(buf, tables),
    advance(gid) {
      const i = Math.min(gid, numberOfHMetrics - 1);
      return u16(buf, tables.hmtx.offset + 4 * i);
    },
    glyphRange(gid) {
      const loca = tables.loca.offset;
      const at = longLoca ? u32(buf, loca + 4 * gid) : u16(buf, loca + 2 * gid) * 2;
      const next = longLoca ? u32(buf, loca + 4 * gid + 4) : u16(buf, loca + 2 * gid + 2) * 2;
      return [tables.glyf.offset + at, tables.glyf.offset + next];
    },
  };
  font.embeddable = (font.fsType & 0x000F) !== 0x0002;
  return font;
}

/** The glyphs a composite glyph is built from. */
function componentsOf(font, gid) {
  const [start, end] = font.glyphRange(gid);
  if (end - start < 10 || i16(font.buf, start) >= 0) return [];
  const out = [];
  let at = start + 10;
  for (;;) {
    const flags = u16(font.buf, at);
    out.push(u16(font.buf, at + 2));
    at += 4 + (flags & 0x0001 ? 4 : 2);
    if (flags & 0x0008) at += 2;
    else if (flags & 0x0040) at += 4;
    else if (flags & 0x0080) at += 8;
    if (!(flags & 0x0020)) break;
  }
  return out;
}

function checksum(buf) {
  let sum = 0;
  for (let i = 0; i < buf.length; i += 4) sum = (sum + buf.readUInt32BE(i)) >>> 0;
  return sum;
}

const pad4 = (b) => (b.length % 4 ? Buffer.concat([b, Buffer.alloc(4 - (b.length % 4))]) : b);

/**
 * The font with only the glyphs in `used` kept (and glyph 0, and the parts
 * of composite glyphs), every other glyph emptied in place, as a .ttf.
 */
function subsetTrueType(font, used) {
  const keep = new Set([0]);
  const stack = [...used].filter((g) => g >= 0 && g < font.numGlyphs);
  while (stack.length) {
    const g = stack.pop();
    if (keep.has(g) && g !== 0) continue;
    keep.add(g);
    for (const c of componentsOf(font, g)) if (!keep.has(c)) stack.push(c);
  }
  const pieces = [];
  const loca = Buffer.alloc(4 * (font.numGlyphs + 1));
  let offset = 0;
  for (let g = 0; g < font.numGlyphs; g++) {
    loca.writeUInt32BE(offset, 4 * g);
    if (!keep.has(g)) continue;
    const [start, end] = font.glyphRange(g);
    if (end <= start) continue;
    const piece = pad4(font.buf.subarray(start, end));
    pieces.push(piece);
    offset += piece.length;
  }
  loca.writeUInt32BE(offset, 4 * font.numGlyphs);
  const glyf = Buffer.concat(pieces.length ? pieces : [Buffer.alloc(4)]);
  const copy = (tag) => (font.tables[tag] ? Buffer.from(font.buf.subarray(font.tables[tag].offset, font.tables[tag].offset + font.tables[tag].length)) : null);
  const head = copy('head');
  head.writeUInt32BE(0, 8); // checkSumAdjustment, worked out below
  head.writeInt16BE(1, 50); // long loca
  const tables = { head, hhea: copy('hhea'), maxp: copy('maxp'), hmtx: copy('hmtx'), loca, glyf, cmap: copy('cmap') };
  for (const tag of ['cvt ', 'fpgm', 'prep', 'OS/2', 'post']) { const t = copy(tag); if (t) tables[tag] = t; }
  const tags = Object.keys(tables).sort();
  const n = tags.length;
  const power = 2 ** Math.floor(Math.log2(n));
  const header = Buffer.alloc(12 + 16 * n);
  header.writeUInt32BE(0x00010000, 0);
  header.writeUInt16BE(n, 4);
  header.writeUInt16BE(power * 16, 6);
  header.writeUInt16BE(Math.log2(power), 8);
  header.writeUInt16BE(n * 16 - power * 16, 10);
  let at = header.length;
  const bodies = [];
  tags.forEach((tag, i) => {
    const body = pad4(tables[tag]);
    const dir = 12 + 16 * i;
    header.write(tag, dir, 4, 'latin1');
    header.writeUInt32BE(checksum(body), dir + 4);
    header.writeUInt32BE(at, dir + 8);
    header.writeUInt32BE(tables[tag].length, dir + 12);
    bodies.push(body);
    at += body.length;
  });
  const file = Buffer.concat([header, ...bodies]);
  const headAt = header.readUInt32BE(12 + 16 * tags.indexOf('head') + 8);
  file.writeUInt32BE((0xB1B0AFBA - checksum(file)) >>> 0, headAt + 8);
  return file;
}

module.exports = { readTrueType, subsetTrueType };
