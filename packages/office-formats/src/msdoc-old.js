// The Word formats before Word 97, read into msdoc.js's model.
//
// Word 6.0 and 95 write a compound file like Word 97's, but everything —
// text, properties, style sheet, fonts, pictures — sits in the one
// WordDocument stream, the FIB is laid out differently, the text is 8-bit
// in the document's code page, and properties are lists of one-byte sprms
// with a numbering of their own. Word 2.0 and 1.x for Windows write no
// compound file at all: the FIB opens the file, its pairs are shorter,
// character properties are a fixed structure rather than sprms, and the
// style sheet is a different shape again. Each is read here and its sprms
// translated into Word 97's, so msdoc.js builds the model and msdoc-docx.js
// writes the .docx as for any Word 97 document.
//
// Windows Write and Word for DOS (the two share a layout: a 128-byte
// header, the text, then 128-byte pages of character and paragraph
// properties) have no sprms, styles or piece table, and are read straight
// into the model.
//
// The layouts follow LibreOffice's Word import (sw/source/filter/ww8) and,
// for Write and Word for DOS, libwps.

import { CompoundFile } from './cfb.js';
import {
  DocError, buildModel, readStyles, resolveStyles, groupTables, readPicture,
  ICO, u8, u16, i16, u32, i32, FC, entriesIn,
} from './msdoc.js';
import { decodeIn, encode1252 } from './codepage.js';

const EMPTY = new Uint8Array(0);

/* ── which Word ───────────────────────────────────────────────────────── */

/** 'word2', 'word1', 'write' or 'worddos' for a file no compound file holds, or null. */
export function oldWordKind(bytes) {
  if (!bytes || bytes.length < 128) return null;
  const ident = u16(bytes, 0);
  if (ident === 0xa5db) return 'word2';
  if (ident === 0xa59b || ident === 0xa59c) return 'word1';
  // Write and Word for DOS: the same six bytes; only Write counts its pages at 96.
  if ((bytes[0] === 0x31 || bytes[0] === 0x32) && bytes[1] === 0xbe && bytes[2] === 0 && bytes[3] === 0 && bytes[4] === 0 && bytes[5] === 0xab) {
    return u16(bytes, 96) ? 'write' : 'worddos';
  }
  return null;
}

/** Read an older Word document (`kind` as wordKind says) into the model. */
export function readOldWord(bytes, kind) {
  if (kind === 'word6') return readWord6(bytes);
  if (kind === 'word2' || kind === 'word1') return readOldStreams(bytes, readOldFib(bytes, kind), kind);
  if (kind === 'write' || kind === 'worddos') return readWriteOrDos(bytes, kind === 'worddos');
  throw new DocError('not a Word document this suite knows');
}

function readWord6(bytes) {
  const cfb = new CompoundFile(bytes);
  const entry = cfb.find(['WordDocument']);
  if (!entry) throw new DocError('not a Word document (no WordDocument stream)');
  const wd = cfb.read(entry);
  return readOldStreams(wd, readOldFib(wd, 'word6'), 'word6');
}

/* ── code pages ───────────────────────────────────────────────────────── */

// A Windows character set, as the code page its 8-bit text is in.
const CHARSET_CODEPAGE = {
  0: 'windows-1252', 1: 'windows-1252', 77: 'macintosh', 128: 'shift_jis', 129: 'euc-kr', 130: 'euc-kr', 134: 'gbk', 136: 'big5',
  161: 'windows-1253', 162: 'windows-1254', 163: 'windows-1258', 177: 'windows-1255', 178: 'windows-1256', 186: 'windows-1257',
  204: 'windows-1251', 222: 'windows-874', 238: 'windows-1250',
};

// A language (its primary id), as the code page Windows wrote it in.
const LANGUAGE_CODEPAGE = {
  0x01: 'windows-1256', 0x02: 'windows-1251', 0x05: 'windows-1250', 0x08: 'windows-1253', 0x0d: 'windows-1255', 0x0e: 'windows-1250',
  0x11: 'shift_jis', 0x12: 'euc-kr', 0x15: 'windows-1250', 0x18: 'windows-1250', 0x19: 'windows-1251', 0x1a: 'windows-1250',
  0x1b: 'windows-1250', 0x1c: 'windows-1250', 0x1e: 'windows-874', 0x1f: 'windows-1254', 0x20: 'windows-1256', 0x22: 'windows-1251',
  0x23: 'windows-1251', 0x24: 'windows-1250', 0x25: 'windows-1257', 0x26: 'windows-1257', 0x27: 'windows-1257', 0x29: 'windows-1256',
  0x2a: 'windows-1258', 0x2f: 'windows-1251',
};

/**
 * The code page of a document's 8-bit text, as Word chose it: the
 * character set the FIB names, or — when it names none — the language the
 * document was written in; Mac Roman for a Word for the Macintosh file.
 */
function documentCodepage(chse, lid) {
  if (chse === 0x100) return 'macintosh';
  if (chse === 0 && lid >= 999) {
    const primary = lid & 0x3ff;
    if (primary === 0x04) return lid === 0x0804 || lid === 0x1004 ? 'gbk' : 'big5';
    return LANGUAGE_CODEPAGE[primary] || 'windows-1252';
  }
  return CHARSET_CODEPAGE[chse] || 'windows-1252';
}

/** A DOS code page number as a decoder's name. */
function dosCodepage(n) {
  if (n === 850) return 'cp850';
  if (n === 866) return 'ibm866';
  return 'cp437';
}

/* ── sprms ────────────────────────────────────────────────────────────── */

// Each one-byte sprm's size, as LibreOffice's tables give it: the operand's
// length and whether a length byte (v) or two (w) come first. An id the
// table does not know has a length byte.
function sizeTable(spec) {
  const t = [];
  for (const item of spec.split(' ')) {
    const [ids, rest] = item.split(':');
    const vari = rest.endsWith('w') ? 2 : rest.endsWith('v') ? 1 : 0;
    const len = parseInt(rest, 10);
    const [a, b = a] = ids.split('-').map(Number);
    for (let id = a; id <= b; id++) t[id] = [len, vari];
  }
  return t;
}
const WW6_SIZES = sizeTable('0:0 2:2 3:3v 4-11:1 12:0v 13-14:1 15:0v 16-19:2 20:4 21-22:2 23:0v 24-25:1 26-28:2 29:1 30-36:2 37:1 38-43:2 44:1 45-49:2 50-51:1 52:0 '
  + '64:0v 65-67:1 68:0v 69:2 70:4 71:1 72:2 73:3 74:0v 75:1 77:0v 79:0v 80:2 81-82:0v 83:0 85-92:1 93:2 94:1 95:3 96-97:2 98:1 99:2 100:1 101:2 102:1 103:0v 104:1 '
  + '105-106:0v 107:2 108:0v 109-110:2 111-116:0v 117-119:1 120:12v 121-124:2 131-132:1 133:0v 136-137:3 138-139:1 140-141:2 142-143:1 144-145:2 146-147:1 148-149:2 '
  + '150-153:1 154-157:2 158-159:1 160-161:2 162:1 163:0 164-171:2 179:0v 181:0v 182-184:2 185-186:1 187:12 188:0w 189:2 190:0w 191:1v 192:4 193:5 194:4 195:2 196:4 197-198:2 199:5 200:4 207:0v');
const WW2_SIZES = sizeTable('0:0 2:1 3:0v 4-14:1 15:0v 16-22:2 23:0v 24-25:1 26-28:2 29:1 30-36:2 37:1 38-43:2 44:1 45-49:2 50-51:1 52:0 53-55:1 57:0v 58:0 60-67:1 68:2 69:1 70:3 '
  + '71-72:2 73-77:1 78:0v 80-81:1 82-83:2 84-87:1 94:1 95:12v 96-99:2 112:1 114:1 115-116:2 117-118:1 119-120:2 121-122:1 123-124:2 125-128:1 129-132:2 133-134:1 135-136:2 '
  + '137-138:1 139-148:2 149:1 152:0w 153:2 154:0w 155:1v 157:5 158:4 159:2 160:4 161-162:2 163:5 164:4');

// Section sprms whose operand passes over as it is: the old id, Word 97's opcode.
const WW6_SECTION = { 142: 0x3009, 143: 0x300a, 144: 0x500b, 145: 0x900c, 147: 0x300e, 153: 0x3014, 156: 0xb017, 157: 0xb018, 161: 0x501c, 162: 0x301d, 164: 0xb01f, 165: 0xb020, 166: 0xb021, 167: 0xb022, 168: 0x9023, 169: 0x9024, 170: 0xb025 };
const WW2_SECTION = { 117: 0x3009, 118: 0x300a, 119: 0x500b, 120: 0x900c, 122: 0x300e, 128: 0x3014, 131: 0xb017, 132: 0xb018, 136: 0x501c, 137: 0x301d, 139: 0xb01f, 140: 0xb020, 141: 0xb021, 142: 0xb022, 143: 0x9023, 144: 0x9024, 145: 0xb025 };

/** A Word 6 border (two bytes: width, kind, shadow, colour, spacing) as Word 97's four. */
function brc80(v) {
  let width = v & 7;
  let type = (v >> 3) & 3;
  if (width > 5) { type = width; width = 1; } // 6 dotted, 7 dashed
  if (!type) return [0, 0, 0, 0];
  return [width * 6, type, (v >> 6) & 0x1f, ((v >> 11) & 0x1f) | ((v >> 5) & 1) << 5];
}

/** A palette colour as a COLORREF's four bytes (automatic for none). */
function colorrefBytes(ico) {
  const hex = ICO[ico];
  if (!hex) return [0, 0, 0, 0xff];
  return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16), 0];
}

/**
 * The function that turns a version's sprms into Word 97's: Word 6/95's
 * (kind 'word6') or Word 2/1's. The text a sprm carries (round a list
 * number) is decoded here, in the document's code page — or as the symbol
 * font's own characters, when the number's font is a symbol font.
 */
function translator(kind, fonts, codepage) {
  const w2 = kind !== 'word6';
  const sizes = w2 ? WW2_SIZES : WW6_SIZES;
  const sectionOps = w2 ? WW2_SECTION : WW6_SECTION;
  const text = (bytes, font) => (fonts[font]?.charset === 2 ? [...bytes].map((b) => 0xf000 | b) : [...decodeIn(codepage, bytes)].map((c) => c.charCodeAt(0)));

  return function translate(g) {
    if (!g || !g.length) return EMPTY;
    const out = [];
    const put = (code, bytes) => { out.push(code & 0xff, code >> 8); for (const x of bytes) out.push(x & 0xff); };
    const w = (o) => [o[0] ?? 0, o[1] ?? 0];

    // A table row's cells: their edges, and each cell's merging and borders in Word 97's 20-byte form.
    const defTable = (o, borders = true) => {
      const n = u8(o, 0);
      const bytes = [n];
      for (let i = 0; i <= n; i++) bytes.push(...w(o.subarray(1 + i * 2)));
      const tcAt = 1 + (n + 1) * 2;
      for (let i = 0; i < n; i++) {
        const at = tcAt + i * 10;
        const has = at + 10 <= o.length;
        bytes.push(has ? u8(o, at) & 3 : 0, 0, 0, 0);
        for (let k = 0; k < 4; k++) bytes.push(...(has && borders ? brc80(u16(o, at + 2 + k * 2)) : [0, 0, 0, 0]));
      }
      const len = bytes.length + 1;
      out.push(0x08, 0xd6, len & 0xff, len >> 8, ...bytes);
    };

    // Word 6's numbering, as Word 97 still records it: ANLD (one level) and OLST (nine).
    const anld = (o) => {
      const bytes = [...o.subarray(0, 20)];
      while (bytes.length < 20) bytes.push(0);
      const cch = Math.min(32, u8(o, 1) + u8(o, 2));
      const chars = text(o.subarray(20, 20 + cch), u16(o, 6));
      for (let i = 0; i < 32; i++) bytes.push((chars[i] ?? 0) & 0xff, (chars[i] ?? 0) >> 8);
      put(0xc63e, [84, ...bytes]);
    };
    const olst = (o) => {
      const bytes = [...o.subarray(0, 148)];
      while (bytes.length < 148) bytes.push(0);
      const chars = [];
      let at = 148;
      for (let lv = 0; lv < 9; lv++) {
        const cch = u8(o, lv * 16 + 1) + u8(o, lv * 16 + 2);
        chars.push(...text(o.subarray(at, Math.min(o.length, at + cch)), u16(o, lv * 16 + 6)));
        at += cch;
      }
      for (let i = 0; i < 32; i++) bytes.push((chars[i] ?? 0) & 0xff, (chars[i] ?? 0) >> 8);
      put(0xd202, [212, ...bytes]);
    };

    const word6 = (id, o, raw) => {
      switch (id) {
        case 2: put(0x4600, w(o)); break;
        case 5: put(0x2403, [o[0]]); break;
        case 7: put(0x2405, [o[0]]); break;
        case 8: put(0x2406, [o[0]]); break;
        case 9: put(0x2407, [o[0]]); break;
        case 12: anld(o); break;
        case 13: put(0x240d, [o[0]]); break;
        case 15: put(0xc60d, raw); break;
        case 16: put(0x840e, w(o)); break;
        case 17: put(0x840f, w(o)); break;
        case 18: put(0x4610, w(o)); break;
        case 19: put(0x8411, w(o)); break;
        case 20: put(0x6412, [...w(o), ...w(o.subarray(2))]); break;
        case 21: put(0xa413, w(o)); break;
        case 22: put(0xa414, w(o)); break;
        case 23: put(0xc615, raw); break;
        case 24: put(0x2416, [o[0]]); break;
        case 25: put(0x2417, [o[0]]); break;
        case 38: case 39: case 40: case 41: put(0x6424 + (id - 38), brc80(u16(o, 0))); break;
        case 47: put(0x442d, w(o)); break;
        case 51: put(0x2431, [o[0]]); break;
        case 65: put(0x0800, [o[0]]); break;
        case 66: put(0x0801, [o[0]]); break;
        case 67: put(0x0802, [o[0]]); break;
        case 68: put(0x6a03, [...o.subarray(0, 4)]); break;
        case 69: put(0x4804, w(o)); break;
        case 71: put(0x0806, [o[0]]); break;
        case 74: { // a symbol: its font, and its character in that font
          const ch = 0xf000 | u8(o, 2);
          put(0x6a09, [...w(o), ch & 0xff, ch >> 8]);
          break;
        }
        case 75: put(0x080a, [o[0]]); break;
        case 80: put(0x4a30, w(o)); break;
        case 83: put(0x2a33, [0]); break;
        case 85: case 86: case 87: case 88: case 89: case 90: case 91: case 92: put(0x0835 + (id - 85), [o[0]]); break;
        case 93: put(0x4a4f, w(o)); break;
        case 94: put(0x2a3e, [o[0]]); break;
        case 95: if (o[0]) put(0x4a43, [o[0], 0]); break;
        case 96: put(0x8840, w(o)); break;
        case 98: put(0x2a42, [o[0]]); break;
        case 99: put(0x4a43, w(o)); break;
        case 104: put(0x2a48, [o[0]]); break;
        case 117: put(0x0855, [o[0]]); break;
        case 118: put(0x0856, [o[0]]); break;
        case 133: olst(o); break;
        case 182: put(0x5400, w(o)); break;
        case 183: put(0x9601, w(o)); break;
        case 184: put(0x9602, w(o)); break;
        case 186: put(0x3404, [o[0]]); break;
        case 187: { const b = [24]; for (let k = 0; k < 6; k++) b.push(...brc80(u16(o, k * 2))); put(0xd605, b); break; }
        case 189: put(0x9407, w(o)); break;
        case 190: defTable(o); break;
        case 191: { // each cell's shading: Word 6's two bytes as Word 97's ten
          const b = [];
          for (let at = 0; at + 2 <= o.length && b.length < 250; at += 2) {
            const v = u16(o, at);
            b.push(...colorrefBytes(v & 0x1f), ...colorrefBytes((v >> 5) & 0x1f), (v >> 10) & 0x3f, 0);
          }
          put(0xd612, [b.length, ...b]);
          break;
        }
        default:
          if (sectionOps[id]) put(sectionOps[id], [...o.subarray(0, sizes[id][0])]);
          break;
      }
    };

    const word2 = (id, o, raw) => {
      switch (id) {
        case 2: put(0x4600, [o[0], 0]); break;
        case 5: put(0x2403, [o[0]]); break;
        case 7: put(0x2405, [o[0]]); break;
        case 8: put(0x2406, [o[0]]); break;
        case 9: put(0x2407, [o[0]]); break;
        case 15: put(0xc60d, raw); break;
        case 16: put(0x840e, w(o)); break;
        case 17: put(0x840f, w(o)); break;
        case 18: put(0x4610, w(o)); break;
        case 19: put(0x8411, w(o)); break;
        case 20: { // line spacing: in twelfths of a line when positive, exact when negative
          const dya = i16(o, 0);
          if (dya) put(0x6412, [...w(o), dya > 0 ? 1 : 0, 0]);
          break;
        }
        case 21: put(0xa413, w(o)); break;
        case 22: put(0xa414, w(o)); break;
        case 23: put(0xc615, raw); break;
        case 24: put(0x2416, [o[0]]); break;
        case 25: put(0x2417, [o[0]]); break;
        case 38: case 39: case 40: case 41: put(0x6424 + (id - 38), brc80(u16(o, 0))); break;
        case 47: put(0x442d, w(o)); break;
        case 51: put(0x2431, [o[0]]); break;
        case 58: put(0x2a33, [0]); break;
        case 60: case 61: case 62: case 63: case 64: case 65: case 66: case 67: put(0x0835 + (id - 60), [o[0]]); break;
        case 68: put(0x4a4f, w(o)); break;
        case 69: put(0x2a3e, [o[0]]); break;
        case 70: if (o[0]) put(0x4a43, [o[0], 0]); break;
        case 71: put(0x8840, w(o)); break;
        case 73: put(0x2a42, [o[0]]); break;
        case 74: put(0x4a43, [o[0], 0]); break;
        case 76: { const v = (o[0] << 24) >> 24; put(0x2a48, [v > 0 ? 1 : v < 0 ? 2 : 0]); break; }
        case 80: put(0x085c, [o[0]]); break;
        case 81: put(0x085d, [o[0]]); break;
        case 85: put(0x4a61, [o[0], 0]); break;
        case 86: put(0x085a, [o[0]]); break;
        case 146: put(0x5400, w(o)); break;
        case 147: put(0x9601, w(o)); break;
        case 148: put(0x9602, w(o)); break;
        case 152: defTable(o, false); break; // the oldest cell borders are of another shape: left out
        case 153: put(0x9407, w(o)); break;
        case 154: defTable(o); break;
        default:
          if (sectionOps[id]) put(sectionOps[id], [...o.subarray(0, sizes[id][0])]);
          break;
      }
    };

    let i = 0;
    while (i < g.length) {
      const id = g[i];
      const [len, vari] = sizes[id] ?? [0, 1];
      let size;
      if (id === 23) { // tab stops: a length byte, or 255 and the counts
        if (u8(g, i + 1) !== 255) size = 2 + u8(g, i + 1);
        else {
          const del = u8(g, i + 2);
          size = 2 + 2 + 4 * del + 3 * u8(g, i + 3 + 4 * del);
        }
      } else if (vari === 0) size = 1 + len;
      else if (vari === 1) size = 2 + u8(g, i + 1) + len;
      else size = 3 + Math.max(0, u16(g, i + 1) - 1) + len;
      if (i + size > g.length) break;
      const raw = g.subarray(i + 1, i + size);
      (w2 ? word2 : word6)(id, raw.subarray(vari), raw);
      i += size;
    }
    return Uint8Array.from(out);
  };
}

/**
 * Word 2's character properties: a fixed structure, cut short after the
 * last field that says anything — flags that flip the style's bold,
 * italic and the rest, then the font, size, colour, underline and raise
 * the run sets itself, and where its picture is. As Word 97 sprms.
 */
function word2Chpx(b) {
  const out = [];
  const put = (code, ...bytes) => out.push(code & 0xff, code >> 8, ...bytes);
  const f0 = u8(b, 0);
  const f1 = u8(b, 1);
  const set = u8(b, 2);
  if (f0 & 0x01) put(0x0835, 0x81);
  if (f0 & 0x02) put(0x0836, 0x81);
  if (f0 & 0x04) put(0x0800, 1);
  if (f0 & 0x08) put(0x0838, 0x81);
  if (f0 & 0x10) put(0x0802, 1);
  if (f0 & 0x20) put(0x083a, 0x81);
  if (f0 & 0x40) put(0x083b, 0x81);
  if (f0 & 0x80) put(0x083c, 0x81);
  if (f1 & 0x01) put(0x0801, 1);
  if (f1 & 0x02) put(0x0855, 1);
  if (f1 & 0x04) put(0x0837, 0x81);
  if (f1 & 0x08) put(0x0856, 1);
  if (f1 & 0x10) put(0x085c, 0x81);
  if (f1 & 0x20) put(0x085d, 0x81);
  if (f1 & 0x40) put(0x085a, 1);
  if (set & 0x02 && b.length >= 6) put(0x4a4f, b[4], b[5]);
  if (set & 0x04 && b.length >= 8) put(0x4a43, b[6], b[7]);
  if (set & 0x01 && b.length >= 10) put(0x2a42, b[9] & 0x1f);
  if (set & 0x08 && b.length >= 10) put(0x2a3e, b[9] >> 5);
  if (set & 0x10 && b.length >= 11) { const v = (b[10] << 24) >> 24; put(0x2a48, v > 0 ? 1 : v < 0 ? 2 : 0); }
  if (f1 & 0x02 && b.length >= 24) { const fc = u32(b, 20); put(0x6a03, fc & 0xff, (fc >> 8) & 0xff, (fc >> 16) & 0xff, fc >>> 24); }
  return Uint8Array.from(out);
}

/* ── the FIB, the pieces, the property pages ──────────────────────────── */

/**
 * An older FIB: the counts of each story's characters, then the pairs
 * saying where each structure is — eight bytes each in Word 6, six in
 * Word 2 and 1 (whose list lacks two printer entries) — in Word 97's order,
 * so FC's indexes reach them; then the property pages' first page numbers
 * and counts, which a file saved without its full tables relies on.
 */
function readOldFib(wd, kind) {
  const w2 = kind !== 'word6';
  const w1 = kind === 'word1';
  const flags = u16(wd, 10);
  const pairSize = w2 ? 6 : 8;
  const slots = w1 ? 36 : 38;
  const slotOf = (i) => {
    if (!w1) return i <= 37 ? i : -1;
    if (i <= 26) return i;
    if (i === 28) return 27;
    return i >= 30 && i <= 37 ? i - 2 : -1;
  };
  // Word 6 keeps the endnotes' pairs after the page counts.
  const extra = w2 ? {} : { [FC.PlcfendRef]: 0x1d2, [FC.PlcfendTxt]: 0x1da };
  const pair = (i) => {
    const s = slotOf(i);
    if (s >= 0) {
      const at = 0x58 + s * pairSize;
      return { fc: u32(wd, at), lcb: w2 ? u16(wd, at + 4) : u32(wd, at + 4) };
    }
    if (extra[i] != null) return { fc: u32(wd, extra[i]), lcb: u32(wd, extra[i] + 4) };
    return { fc: 0, lcb: 0 };
  };
  const tail = 0x58 + slots * pairSize + (w1 ? 6 : 2);
  return {
    nFib: u16(wd, 2),
    lid: u16(wd, 6),
    chse: u16(wd, 20),
    fDot: Boolean(flags & 0x0001),
    fComplex: Boolean(flags & 0x0004),
    fEncrypted: Boolean(flags & 0x0100),
    fExtChar: Boolean(flags & 0x1000),
    fcMin: u32(wd, 0x18),
    fcMac: u32(wd, 0x1c),
    ccpText: i32(wd, 0x34),
    ccpFtn: i32(wd, 0x38),
    ccpHdd: i32(wd, 0x3c),
    ccpMcr: i32(wd, 0x40),
    ccpAtn: i32(wd, 0x44),
    ccpEdn: i32(wd, 0x48),
    ccpTxbx: i32(wd, 0x4c),
    ccpHdrTxbx: i32(wd, 0x50),
    pnChpFirst: w1 ? 0 : u16(wd, tail),
    pnPapFirst: w1 ? 0 : u16(wd, tail + 2),
    cpnBteChp: u16(wd, w1 ? tail : tail + 4),
    cpnBtePap: u16(wd, w1 ? tail + 2 : tail + 6),
    pair,
  };
}

/** A fast-saved file's piece table: positions raw, text 8-bit unless the FIB says otherwise. */
function readOldPieces(table, clx, w2, unicode, translate) {
  const grpprls = [];
  let at = clx.fc;
  const end = clx.fc + clx.lcb;
  while (at < end && table[at] === 0x01) {
    const cb = u16(table, at + 1);
    grpprls.push(translate(table.subarray(at + 3, at + 3 + cb)));
    at += 3 + cb;
  }
  if (table[at] !== 0x02) throw new DocError('the document\'s piece table is damaged');
  const lcb = w2 ? u16(table, at + 1) : u32(table, at + 1);
  const plc = at + (w2 ? 3 : 5);
  const n = entriesIn(lcb, 12, plc, table.length);
  const pieces = [];
  for (let i = 0; i < n; i++) {
    const pcd = plc + (n + 1) * 4 + i * 8;
    pieces.push({ cpStart: u32(table, plc + i * 4), cpEnd: u32(table, plc + (i + 1) * 4), compressed: !unicode, fc: u32(table, pcd + 2), prm: u16(table, pcd + 6) });
  }
  return { pieces, grpprls };
}

/**
 * Every character or paragraph property run, from the pages the bin table
 * lists — or, when the FIB counts more pages than the table names (a file
 * saved without its full table), from the pages it counts. Each run's
 * properties as Word 97 sprms.
 */
function readOldFkps(wd, table, plc, cpn, pnFirst, which, kind, translate) {
  const w2 = kind !== 'word6';
  const n = plc.lcb >= 4 ? entriesIn(plc.lcb, 6, plc.fc, table.length) : 0;
  let pns = [];
  for (let i = 0; i < n; i++) pns.push(u16(table, plc.fc + (n + 1) * 4 + i * 2));
  if (cpn > n && pnFirst) pns = Array.from({ length: cpn }, (_, i) => pnFirst + i);
  const bx = which === 'pap' && !w2 ? 7 : 1;
  const out = [];
  for (const pn of pns) {
    const page = pn * 512;
    if (page + 512 > wd.length) continue;
    const count = wd[page + 511];
    for (let r = 0; r < count; r++) {
      const offAt = page + (count + 1) * 4 + r * bx;
      if (offAt >= page + 511) break;
      const fcStart = u32(wd, page + r * 4);
      const fcEnd = u32(wd, page + (r + 1) * 4);
      const off = wd[offAt] * 2;
      const lim = page + 512;
      if (which === 'chp') {
        if (!off) { out.push({ fcStart, fcEnd, grpprl: EMPTY }); continue; }
        const raw = wd.subarray(page + off + 1, Math.min(lim, page + off + 1 + wd[page + off]));
        out.push({ fcStart, fcEnd, grpprl: w2 ? word2Chpx(raw) : translate(raw) });
      } else {
        if (!off) { out.push({ fcStart, fcEnd, istd: 0, grpprl: EMPTY }); continue; }
        const len = wd[page + off] * 2;
        // Word 2: a one-byte style code and six bytes of layout cache before the sprms; Word 6: a two-byte style.
        const istd = w2 ? wd[page + off + 1] : u16(wd, page + off + 1);
        const start = page + off + (w2 ? 8 : 3);
        const grpprl = len >= (w2 ? 7 : 2) ? wd.subarray(start, Math.min(lim, page + off + 1 + len)) : EMPTY;
        out.push({ fcStart, fcEnd, istd, grpprl: translate(grpprl) });
      }
    }
  }
  return out.sort((a, b) => a.fcStart - b.fcStart);
}

/* ── fonts and styles ─────────────────────────────────────────────────── */

/** A font name, in the code page of its own character set. */
function fontName(bytes, charset, lid) {
  let end = 0;
  while (end < bytes.length && bytes[end]) end++;
  const cp = charset === 2 ? 'windows-1252' : documentCodepage(charset, lid);
  return decodeIn(cp, bytes.subarray(0, end));
}

/**
 * The font table: a total length, then each font — its size less one, its
 * family, (Word 6) its weight, its character set, and its name — the name
 * six bytes in for Word 6, three for Word 2.
 */
function readOldFonts(table, plc, lid, w2) {
  const fonts = [];
  if (plc.lcb <= 2) return fonts;
  const end = plc.fc + plc.lcb;
  let at = plc.fc + 2;
  while (at < end) {
    const cb = u8(table, at) + 1;
    if (cb <= 1 || at + cb > end + 1) break;
    const charset = u8(table, at + (w2 ? 2 : 4));
    fonts.push({ name: fontName(table.subarray(at + (w2 ? 3 : 6), Math.min(end, at + cb)), charset, lid), family: (u8(table, at + 1) >> 4) & 7, charset });
    at += cb;
  }
  return fonts;
}

/** A Word 6 style name: a length byte, the characters and a terminating zero. */
const pascalNames = (codepage) => (b, at) => {
  const n = u8(b, at);
  return { name: decodeIn(codepage, b.subarray(at + 1, at + 1 + n)), end: at + 1 + n + 1 };
};

// Word 2's style codes for the built-in styles, from 222 up, as Word 97's style identifiers.
const STC_STI = [4095, 39, 30, 26, 25, 24, 23, 22, 21, 20, 19, 16, 15, 14, 13, 12, 11, 10, 40, 33, 32, 31, 38, 29, 9, 8, 7, 6, 5, 4, 3, 2, 1, 28];
const STI_NAME = { 0: 'Normal', 28: 'Normal Indent', 29: 'footnote text', 30: 'annotation text', 31: 'header', 32: 'footer', 33: 'index heading', 38: 'footnote reference', 39: 'annotation reference', 40: 'line number' };
const stiName = (sti) => STI_NAME[sti] || (sti >= 1 && sti <= 9 ? `heading ${sti}` : sti >= 10 && sti <= 18 ? `index ${sti - 9}` : sti >= 19 && sti <= 27 ? `toc ${sti - 18}` : null);

/**
 * Word 2's style sheet: how many built-in styles come first, then the
 * names (a length byte each — 0 for a built-in style's own name, 255 for
 * none), each style's character properties (a Word 2 CHPX), its paragraph
 * properties (a style code, six bytes of layout cache, then sprms), and its
 * next and base styles. Styles are named by code: the built-in ones from
 * 222 up, Normal 0, the user's own after.
 */
function readWord2Styles(table, plc, codepage, translate) {
  const styles = [];
  if (!plc.lcb) return { styles, defaultFonts: [0, 0, 0] };
  let at = plc.fc;
  const cstcStd = u16(table, at);
  at += 2;
  const stcOf = (stcp) => (stcp - cstcStd) & 255;

  const entries = [];
  const cbName = u16(table, at);
  let p = at + 2;
  const namesEnd = at + cbName;
  while (p < namesEnd) {
    const n = u8(table, p++);
    const entry = { defined: n !== 0xff, name: '' };
    if (n !== 0xff && n) { entry.name = decodeIn(codepage, table.subarray(p, p + n)); p += n; }
    entries.push(entry);
  }
  at = namesEnd;

  const cbChpx = u16(table, at);
  p = at + 2;
  for (let i = 0; p < at + cbChpx && i < entries.length; i++) {
    const cb = u8(table, p++);
    if (cb !== 0xff) { entries[i].chpx = word2Chpx(table.subarray(p, p + cb)); p += cb; }
  }
  at += cbChpx;

  const cbPapx = u16(table, at);
  p = at + 2;
  for (let i = 0; p < at + cbPapx && i < entries.length; i++) {
    const cb = u8(table, p++);
    if (cb !== 0xff) {
      if (cb >= 7) entries[i].papx = translate(table.subarray(p + 7, p + cb));
      p += cb;
    }
  }
  at += cbPapx;

  const iMac = Math.min(u16(table, at), entries.length);
  at += 2;
  entries.forEach((e, stcp) => {
    const stc = stcOf(stcp);
    const sti = stc === 0 ? 0 : stc < 222 ? 4094 : STC_STI[stc - 222];
    if (!e.defined || sti === 4095) return;
    let next = stcp < iMac ? u8(table, at + stcp * 2) : stc;
    let base = stcp < iMac ? u8(table, at + stcp * 2 + 1) : 222;
    if (base === stc || base === 222) base = null;
    if (next === 222) next = null;
    const isChar = (sti === 38 || sti === 39 || sti === 40) && !e.papx?.length;
    styles[stc] = {
      istd: stc, sti, name: e.name || stiName(sti) || `Style ${stc}`, kind: isChar ? 'character' : 'paragraph',
      base, next, papx: isChar ? null : e.papx || EMPTY, chpx: e.chpx || EMPTY,
    };
  });
  resolveStyles(styles, [0, 0, 0]);
  return { styles, defaultFonts: [0, 0, 0] };
}

/* ── pictures ─────────────────────────────────────────────────────────── */

/**
 * A picture before Word 97: the PIC header (the same first 46 bytes as
 * Word 97's) and then the picture itself, a Windows metafile, in the main
 * stream. Given a placeable header so a reader knows its size.
 */
function readOldPicture(stream, at) {
  if (at + 68 > stream.length) return null;
  const mm = i16(stream, at + 6);
  if (mm === 0x64 || mm === 0x66) return readPicture(stream, at);
  if (mm === 94 || mm === 98 || mm === 99) return null; // a linked file, not in the document
  const lcb = u32(stream, at);
  const cbHeader = u16(stream, at + 4);
  const widthTwips = Math.round((i16(stream, at + 28) * (u16(stream, at + 32) || 1000)) / 1000);
  const heightTwips = Math.round((i16(stream, at + 30) * (u16(stream, at + 34) || 1000)) / 1000);
  const wmf = stream.subarray(at + cbHeader, Math.min(stream.length, at + lcb));
  if (wmf.length < 18) return null;
  return { contentType: 'image/x-wmf', ext: 'wmf', bytes: placeableWmf(wmf, widthTwips, heightTwips), widthTwips, heightTwips };
}

/**
 * A metafile with the 22-byte placeable header in front: its bounds from
 * the metafile's own window origin and extent, and units per inch that
 * make those bounds the size the document gives the picture.
 */
export function placeableWmf(wmf, widthTwips, heightTwips) {
  if (u32(wmf, 0) === 0x9ac6cdd7) return wmf;
  let org = [0, 0];
  let ext = null;
  let p = u16(wmf, 2) * 2;
  while (p + 6 <= wmf.length) {
    const size = u32(wmf, p) * 2;
    const fn = u16(wmf, p + 4);
    if (fn === 0x020b) org = [i16(wmf, p + 8), i16(wmf, p + 6)];
    if (fn === 0x020c) ext = [i16(wmf, p + 8), i16(wmf, p + 6)];
    if (!size || fn === 0) break;
    p += size;
  }
  const [ex, ey] = ext || [widthTwips || 1440, heightTwips || 1440];
  const inch = Math.max(1, Math.min(0xffff, Math.round((Math.abs(ex) * 1440) / (widthTwips || Math.abs(ex)))));
  const head = new Uint8Array(22);
  const dv = new DataView(head.buffer);
  dv.setUint32(0, 0x9ac6cdd7, true);
  dv.setInt16(6, org[0], true);
  dv.setInt16(8, org[1], true);
  dv.setInt16(10, org[0] + ex, true);
  dv.setInt16(12, org[1] + ey, true);
  dv.setUint16(14, inch, true);
  let sum = 0;
  for (let i = 0; i < 10; i++) sum ^= dv.getUint16(i * 2, true);
  dv.setUint16(20, sum, true);
  const out = new Uint8Array(22 + wmf.length);
  out.set(head);
  out.set(wmf, 22);
  return out;
}

/* ── Word 6/95 and Word 2/1 ───────────────────────────────────────────── */

function readOldStreams(wd, fib, kind) {
  if (fib.fEncrypted) {
    const err = new DocError('This Word document is password-protected.');
    err.encrypted = true;
    throw err;
  }
  const w2 = kind !== 'word6';
  const codepage = documentCodepage(fib.chse, fib.lid);
  const fonts = readOldFonts(wd, fib.pair(FC.SttbfFfn), fib.lid, w2);
  const translate = translator(kind, fonts, codepage);
  const { pieces, grpprls } = fib.fComplex
    ? readOldPieces(wd, fib.pair(FC.Clx), w2, fib.fExtChar, translate)
    : { pieces: [{ cpStart: 0, cpEnd: Math.max(0, Math.floor((fib.fcMac - fib.fcMin) / (fib.fExtChar ? 2 : 1))), compressed: !fib.fExtChar, fc: fib.fcMin, prm: 0 }], grpprls: [] };

  // The header and footer stories run on from section to section: each
  // section has those its flags name, after the notes' separators the
  // document's own flags name.
  const dop = fib.pair(FC.Dop);
  let story = popcount(dop.lcb ? u8(wd, dop.fc + 1) & 0x3f : 0);
  const headerSlots = (s, sep) => [0, 1, 2, 3, 4, 5].map((bit) => ((sep.grpfIhdt || 0) & (1 << bit) ? story++ : null));

  const model = buildModel({
    wd, table: wd, data: wd, fib, pieces, grpprls,
    chpxList: readOldFkps(wd, wd, fib.pair(FC.PlcfBteChpx), fib.cpnBteChp, fib.pnChpFirst, 'chp', kind, translate),
    papxList: readOldFkps(wd, wd, fib.pair(FC.PlcfBtePapx), fib.cpnBtePap, fib.pnPapFirst, 'pap', kind, translate),
    readStylesWith: w2 ? (t, plc) => readWord2Styles(t, plc, codepage, translate) : (t, plc) => readStyles(t, plc, pascalNames(codepage), translate),
    readListsWith: () => ({ lists: [], lfos: [] }),
    readFontsWith: () => fonts,
    readPictureWith: (at) => readOldPicture(wd, at),
    sedSize: w2 ? 6 : 12,
    readSepx: (fc) => translate(w2 ? wd.subarray(fc + 1, fc + 1 + u8(wd, fc)) : wd.subarray(fc + 2, fc + 2 + u16(wd, fc))),
    headerSlots,
    format: kind,
  });
  if (!w2) numberOldLists(model);
  if (!fib.fExtChar) recodeText(model, codepage);
  return model;
}

function popcount(v) {
  let n = 0;
  for (let x = v; x; x >>= 1) n += x & 1;
  return n;
}

/** Each paragraph of a story, table cells' too, in order. */
function eachParagraph(blocks, fn) {
  for (const b of blocks || []) {
    if (b.type === 'table') for (const row of b.rows) for (const cell of row.cells) eachParagraph(cell, fn);
    else fn(b);
  }
}
/** Every story of the model: the body, the notes, the headers and footers. */
function eachStory(model, fn) {
  fn(model.body, 'body');
  for (const kind of ['footnote', 'endnote']) for (const n of model.notes[kind]) fn(n.blocks, kind);
  const seen = new Set();
  for (const hf of model.headers) {
    for (const blocks of [...Object.values(hf.headers), ...Object.values(hf.footers)]) {
      if (!seen.has(blocks)) { seen.add(blocks); fn(blocks, 'header'); }
    }
  }
}

/**
 * The text, read as Windows-1252, in the code page it was written in: the
 * document's, or the run's font's own character set — and a symbol font's
 * characters where Word keeps them, in the private-use area.
 */
function recodeText(model, codepage) {
  eachStory(model, (blocks) => eachParagraph(blocks, (p) => {
    for (const run of p.runs) {
      if (run.text == null || !run.text) continue;
      const charset = model.fonts[run.chp?.font]?.charset ?? 0;
      if (charset === 2) {
        run.text = String.fromCharCode(...[...encode1252(run.text)].map((b) => (b >= 0x20 ? 0xf000 | b : b)));
        continue;
      }
      const cp = charset === 0 || charset === 1 ? codepage : CHARSET_CODEPAGE[charset] || codepage;
      if (cp !== 'windows-1252') run.text = decodeIn(cp, encode1252(run.text));
    }
  }));
}

/**
 * Word 6's numbering as lists. A paragraph says what kind of number it
 * has — 10 a numbered list, 11 bullets, 1 to 9 a heading level, 12 a pause
 * — and carries the number's look (an ANLD). A numbered or bulleted list
 * runs while the paragraphs after one another keep its kind, and starts
 * again after a paragraph with none; heading numbers run through the
 * document, their looks from the section's OLST.
 */
function numberOldLists(model) {
  const { lists, lfos } = model.lists;
  let lsid = 1;
  const add = (levels) => {
    lists.push({ lsid, levels });
    lfos.push({ lsid });
    lsid += 1;
    return lfos.length;
  };
  let outline = null;
  let section = 0;
  eachStory(model, (blocks, story) => {
    let current = null;
    const visit = (list) => {
      for (const b of list) {
        if (b.type === 'table') {
          for (const row of b.rows) for (const cell of row.cells) visit(cell);
        } else {
          const lvl = b.pap.nLvlAnm ?? 0;
          const anld = b.pap.anld;
          if ((lvl === 10 || lvl === 11) && anld) {
            if (!current || current.lvl !== lvl) current = { lvl, ilfo: add([anlvLevel(anld, anld.text, 0, lvl === 11)]) };
            b.pap.ilfo = current.ilfo;
            b.pap.ilvl = 0;
          } else if (lvl >= 1 && lvl <= 9) {
            const olst = model.sections[section]?.sep.olst;
            if (!outline) {
              const levels = [];
              for (let k = 0; k < 9; k++) levels.push(olst ? olstLevel(olst, k) : { start: 1, nfc: 0, jc: 0, follow: 0, pap: {}, chp: {}, codes: [k, 0x2e] });
              outline = { ilfo: add(levels), levels, fromOlst: Boolean(olst) };
            }
            if (!outline.fromOlst && anld) outline.levels[lvl - 1] = anlvLevel(anld, anld.text, lvl - 1, false);
            b.pap.ilfo = outline.ilfo;
            b.pap.ilvl = lvl - 1;
            current = null;
          } else if (lvl !== 12) current = null;
        }
        if (story === 'body' && b.sectionEnd != null) section = b.sectionEnd + 1;
      }
    };
    visit(blocks);
  });
}

/** One level of a list, from an ANLV and the text round its number. */
function anlvLevel(anlv, text, ilvl, bullet) {
  const before = [...text.slice(0, anlv.before)].map((c) => c.charCodeAt(0));
  const after = [...text.slice(anlv.before, anlv.before + anlv.after)].map((c) => c.charCodeAt(0));
  let codes;
  if (bullet) codes = [before[0] ?? after[0] ?? 0x2022];
  else {
    codes = [];
    if (anlv.prev) for (let k = 0; k < ilvl; k++) codes.push(k, 0x2e);
    codes.push(...before, ilvl, ...after);
  }
  return {
    start: anlv.start, nfc: bullet ? 23 : anlv.nfc <= 5 ? anlv.nfc : 0, jc: anlv.jc === 1 || anlv.jc === 2 ? anlv.jc : 0,
    follow: anlv.hang ? 0 : 1, pap: {}, chp: bullet ? { font: anlv.font } : {}, codes,
  };
}

/** One heading level from the section's OLST, whose levels share one run of text in turn. */
function olstLevel(olst, k) {
  let at = 0;
  for (let i = 0; i < k; i++) at += olst.levels[i].before + olst.levels[i].after;
  const lv = olst.levels[k];
  return anlvLevel(lv, olst.text.slice(at), k, false);
}

/* ── Windows Write and Word for DOS ───────────────────────────────────── */

const DOS_COLOURS = ['000000', 'FF0000', '00FF00', '0000FF', '7F00FF', 'FF00FF', '00FFFF', 'FFFFFF'];

/**
 * Write or Word for DOS. After the 128-byte header comes the text, then
 * pages of property runs — character runs first, from the page after the
 * text, paragraph runs from the page the header names — each run its end
 * and a few bytes that override the defaults. Then the section, the fonts
 * (Write) and the footnotes (Word for DOS). Running heads are paragraphs
 * of the text marked as such; Write's pictures are paragraphs whose bytes
 * are the picture.
 */
function readWriteOrDos(bytes, dos) {
  const fcMac = u32(bytes, 14);
  if (fcMac <= 0x80 || fcMac > bytes.length) throw new DocError('the document is damaged (its text runs past the end of the file)');
  const codepage = dos ? dosCodepage(u16(bytes, 126)) : 'windows-1252';
  const pages = Math.ceil(bytes.length / 128);

  // The property runs of a kind, from page `pn` on, until the text's end.
  const fods = (pn) => {
    const out = [];
    let fc = 0x80;
    for (let page = pn; page < pages && fc < fcMac; page++) {
      const base = page * 128;
      const cfod = Math.min(20, u8(bytes, base + 127));
      if (!cfod) break;
      for (let f = 0; f < cfod; f++) {
        const fcLim = u32(bytes, base + 4 + f * 6);
        const bfProp = u16(bytes, base + 4 + f * 6 + 4);
        let prop = EMPTY;
        if (bfProp < 0x7f - 4) {
          const cch = u8(bytes, base + 4 + bfProp);
          if (bfProp + cch + 4 < 0x80) prop = bytes.subarray(base + 5 + bfProp, base + 5 + bfProp + cch);
        }
        out.push({ fcFirst: fc, fcLim, prop });
        if (fcLim >= fcMac || fcLim <= fc) return out;
        fc = fcLim;
      }
    }
    return out;
  };
  const overlay = (defaults, prop) => {
    const b = Uint8Array.from(defaults);
    b.set(prop.subarray(0, b.length));
    return b;
  };

  // Fonts: Write's table, or Word for DOS's printer fonts by family.
  const fonts = [];
  const fontIndex = new Map();
  const fontOf = (name) => {
    if (!fontIndex.has(name)) { fontIndex.set(name, fonts.length); fonts.push({ name, family: 0, charset: /^(Symbol|Wingdings|Webdings)/i.test(name) ? 2 : 0 }); }
    return fontIndex.get(name);
  };
  const writeFonts = [];
  if (!dos) {
    let pn = u16(bytes, 28);
    if (pn && pn !== u16(bytes, 96) && u16(bytes, pn * 128)) {
      let at = pn * 128 + 2;
      for (let guard = 0; guard < 4096 && at + 2 <= bytes.length; guard++) {
        const cb = u16(bytes, at);
        if (!cb) break;
        if (cb === 0xffff) { pn += 1; at = pn * 128; continue; }
        writeFonts.push(fontOf(decodeIn(codepage, bytes.subarray(at + 3, at + 2 + cb)).replace(/\0+$/, '')));
        at += 2 + cb;
      }
    }
    if (!writeFonts.length) writeFonts.push(fontOf('Arial'));
  }
  const dosFont = (ftc) => fontOf(ftc <= 15 ? 'Courier New' : ftc <= 31 ? 'Times New Roman' : ftc <= 39 ? 'Script' : ftc >= 56 ? 'Symbol' : 'Courier New');

  // Character runs.
  const chps = fods(Math.floor((fcMac + 127) / 128)).map((f) => {
    const b = overlay([dos ? 0 : 1, 0, 24, 0, 0, 0, 0], f.prop);
    const ftc = dos ? b[1] >> 2 : (b[1] >> 2) | ((b[4] & 7) << 6);
    const chp = { size: b[2] || 24, font: dos ? dosFont(ftc) : writeFonts[ftc] ?? writeFonts[0] };
    if (b[1] & 1) chp.bold = true;
    if (b[1] & 2) chp.italic = true;
    if (b[3] & 1) chp.underline = 1;
    if (dos) {
      if (b[3] & 2) chp.strike = true;
      if (b[3] & 4) chp.underline = 3;
      if ((b[3] & 0x30) === 0x10) chp.caps = true;
      if ((b[3] & 0x30) === 0x30) chp.smallCaps = true;
      if (b[3] & 0x80) chp.vanish = true;
      if (b[6] & 7) { chp.colour = DOS_COLOURS[b[6] & 7]; chp.colourSet = true; }
    }
    if (b[5]) chp.iss = b[5] & 0x80 ? 2 : 1;
    const special = Boolean(b[3] & 0x40);
    const styled = dos && b[0] & 1 ? b[0] >> 1 : null;
    return { ...f, chp, special, note: styled === 13 || styled === 26 };
  });

  // Paragraph runs.
  const paps = fods(u16(bytes, 18)).map((f) => {
    const size = dos ? 102 : 78;
    const def = new Uint8Array(size);
    def[10] = 240;
    const b = overlay(def, f.prop);
    const pap = { istd: 0, jc: b[1] & 3 };
    const left = i16(b, 6);
    const right = i16(b, 4);
    const first = i16(b, 8);
    if (left) pap.left = left;
    if (right) pap.right = right;
    if (first) pap.firstLine = first;
    const line = i16(b, 10);
    if (line > 0 && line !== 240) pap.line = { dya: line, mult: true };
    if (u16(b, 12)) pap.before = u16(b, 12);
    if (u16(b, 14)) pap.after = u16(b, 14);
    const tabs = [];
    for (let t = 0; t < (dos ? 20 : 14); t++) {
      const pos = u16(b, 22 + t * 4);
      if (!pos) break;
      const jcTab = b[22 + t * 4 + 2];
      tabs.push({ pos, jc: dos ? jcTab & 3 : (jcTab & 3) === 3 ? 3 : 0, leader: dos ? (jcTab >> 3) & 3 : 0 });
    }
    if (tabs.length) pap.tabs = tabs;
    const rhc = b[16];
    const out = { ...f, pap };
    if (dos) {
      if (b[1] & 4) pap.keep = true;
      if (b[1] & 8) pap.keepNext = true;
      if (rhc & 0x0e) out.runningHead = { footer: Boolean(rhc & 1), first: Boolean(rhc & 8), every: ((rhc >> 1) & 3) !== 0 };
      if (rhc & 0x30) {
        const sides = (rhc & 0x30) === 0x10 ? 15 : b[17] & 15;
        const kind = rhc & 0xc0;
        const brc = { width: kind === 0x40 ? 12 : kind === 0xc0 ? 24 : 4, type: kind === 0x80 ? 3 : 1, colour: DOS_COLOURS[(b[17] >> 4) & 7], space: 0 };
        pap.borders = { left: sides & 1 ? brc : null, right: sides & 2 ? brc : null, top: sides & 4 ? brc : null, bottom: sides & 8 ? brc : null };
      }
      const shade = Math.min(100, b[18] & 0x7f);
      if (shade) {
        const c = DOS_COLOURS[(b[19] >> 4) & 7];
        const mix = (h) => Math.min(255, parseInt(h, 16) + Math.round((255 * (100 - shade)) / 100)).toString(16).padStart(2, '0');
        pap.shading = (mix(c.slice(0, 2)) + mix(c.slice(2, 4)) + mix(c.slice(4, 6))).toUpperCase();
      }
      if (b[0] & 1 && (b[0] >> 1 === 39 || b[0] >> 1 === 87)) out.note = true;
    } else {
      if (rhc & 0x10) out.picture = true;
      else if (rhc & 6) out.runningHead = { footer: Boolean(rhc & 1), first: Boolean(rhc & 8), every: true };
    }
    return out;
  });
  if (!paps.length) throw new DocError('the document is damaged (it has no paragraphs)');

  // Footnotes (Word for DOS): where each reference is, and where its text starts.
  const notesAt = [];
  if (dos) {
    const pn = u16(bytes, 20);
    if (pn && pn !== u16(bytes, 22)) {
      const n = u16(bytes, pn * 128);
      for (let i = 0; i < n && i < 1000; i++) notesAt.push({ ref: u32(bytes, pn * 128 + 4 + i * 8) + 0x80, text: u32(bytes, pn * 128 + 8 + i * 8) + 0x80 });
    }
  }

  // The section: the page and where the text sits on it, in twips.
  const sep = { width: 12240, height: 15840, left: 1800, right: 1440, top: 1440, bottom: 1440, header: 1080, footer: 1080, landscape: false, titlePage: false, columns: 1, break: 2 };
  const pnSetb = u16(bytes, 24);
  if (pnSetb && pnSetb !== u16(bytes, 26) && pnSetb * 128 + 16 <= bytes.length) {
    const fcSep = u32(bytes, pnSetb * 128 + 4 + 6);
    const cch = u8(bytes, fcSep);
    if (fcSep !== 0xffffffff && fcSep + 23 <= bytes.length && cch >= 22) {
      const v = (o) => u16(bytes, fcSep + o);
      const [yaMac, xaMac, yaTop, dyaText, xaLeft, dxaText, yaHeader, yaFooter] = [v(3), v(5), v(9), v(11), v(13), v(15), v(19), v(21)];
      if (yaMac && xaMac && dyaText && dxaText) {
        Object.assign(sep, {
          height: yaMac, width: xaMac, top: yaTop, bottom: Math.max(0, yaMac - yaTop - dyaText), left: xaLeft, right: Math.max(0, xaMac - xaLeft - dxaText),
          header: yaHeader, footer: Math.max(0, yaMac - yaFooter), landscape: xaMac > yaMac,
        });
      }
      if (dos && cch >= 18 && u8(bytes, fcSep + 18) > 1) sep.columns = u8(bytes, fcSep + 18);
    }
  }

  // The paragraphs, each its runs of text.
  const images = [];
  let chpIndex = 0;
  const chpAt = (fc) => {
    while (chpIndex < chps.length - 1 && fc >= chps[chpIndex].fcLim) chpIndex++;
    while (chpIndex > 0 && fc < chps[chpIndex].fcFirst) chpIndex--;
    return chps[chpIndex] || { chp: { size: 24, font: 0 }, special: false, note: false };
  };
  const noteRefs = new Map(notesAt.slice(0, -1).map((n, i) => [n.ref, i + 1]));
  const paragraph = (pp) => {
    const runs = [];
    let cur = null;
    let key = '';
    const end = Math.min(pp.fcLim, fcMac);
    let fc = pp.fcFirst;
    // A paragraph ends with a carriage return and a line feed; they are its mark.
    let stop = end;
    if (stop > fc && bytes[stop - 1] === 10) stop--;
    if (stop > fc && bytes[stop - 1] === 13) stop--;
    for (; fc < stop; fc++) {
      const c = chpAt(fc);
      const byte = bytes[fc];
      const push = (run) => { runs.push({ ...run, chp: c.chp }); cur = null; key = ''; };
      if (c.special) {
        if (byte === 1 || byte === 9) push({ kind: 'field', instr: 'PAGE' });
        else if (byte === 2) push({ kind: 'field', instr: 'DATE' });
        else if (byte === 3) push({ kind: 'field', instr: 'TIME' });
        else if ((byte === 4 || byte === 5) && noteRefs.has(fc)) push({ kind: 'noteRef', note: { kind: 'footnote', id: noteRefs.get(fc) } });
        continue;
      }
      if (c.note && noteRefs.has(fc)) { push({ kind: 'noteRef', note: { kind: 'footnote', id: noteRefs.get(fc) } }); continue; }
      if (byte < 0x20) {
        if (byte === 9) push({ kind: 'tab' });
        else if (byte === 10 || byte === 11) push({ kind: 'break', type: 'line' });
        else if (byte === 12) push({ kind: 'break', type: 'page' });
        else if (byte === 14 && dos) push({ kind: 'break', type: 'column' });
        else if (byte === 15 && dos) push({ text: '—' });
        else if (byte === 31) push({ kind: 'softHyphen' });
        else if (byte === 30) push({ text: ' ' });
        continue;
      }
      const k = JSON.stringify(c.chp);
      if (!cur || k !== key) { cur = { text: '', chp: c.chp, bytes: [] }; key = k; runs.push(cur); }
      cur.bytes.push(byte);
    }
    for (const r of runs) {
      if (!r.bytes) continue;
      const symbol = fonts[r.chp.font]?.charset === 2;
      r.text = symbol ? String.fromCharCode(...r.bytes.map((b) => 0xf000 | b)) : decodeIn(codepage, Uint8Array.from(r.bytes));
      delete r.bytes;
    }
    // A page number, date or time as a field: begin, its instruction, end.
    const expanded = [];
    for (const r of runs) {
      if (r.kind === 'field') expanded.push({ kind: 'fieldBegin', chp: r.chp }, { text: ` ${r.instr} `, chp: r.chp }, { kind: 'fieldEnd', chp: r.chp });
      else expanded.push(r);
    }
    return { pap: pp.pap, runs: expanded, mark: 'para', cpEnd: pp.fcLim, markChp: chpAt(Math.max(pp.fcFirst, end - 1)).chp };
  };

  // Write's pictures: a header naming the kind and size, then a metafile or a bitmap.
  const pictureParagraph = (pp) => {
    const at = pp.fcFirst;
    const end = Math.min(pp.fcLim, fcMac);
    if (end - at < 40) return null;
    const mm = u16(bytes, at);
    const mx = u16(bytes, at + 36);
    const my = u16(bytes, at + 38);
    const scale = mx > 10 && my > 10 && !(mx === 1000 && my === 1000);
    const widthTwips = Math.round(u16(bytes, at + 10) * (scale ? mx / 1000 : 1));
    const heightTwips = Math.round(u16(bytes, at + 12) * (scale ? my / 1000 : 1));
    const cb = u32(bytes, at + 32);
    const data = bytes.subarray(at + 40, Math.min(end, at + 40 + cb));
    let image = null;
    if (mm === 0x88 && data.length > 18) image = { contentType: 'image/x-wmf', ext: 'wmf', bytes: placeableWmf(data, widthTwips, heightTwips) };
    else if (mm === 0xe3) {
      const bmp = monochromeBmp(data, u16(bytes, at + 18), u16(bytes, at + 20), u16(bytes, at + 22), u8(bytes, at + 24), u8(bytes, at + 25));
      if (bmp) image = { contentType: 'image/bmp', ext: 'bmp', bytes: bmp };
    }
    if (!image) return null;
    const index = images.push({ ...image, widthTwips, heightTwips }) - 1;
    return { pap: pp.pap, runs: [{ kind: 'picture', image: index, chp: chpAt(at).chp }], mark: 'para', cpEnd: pp.fcLim };
  };

  const inNote = (pp) => notesAt.length > 1 && pp.fcFirst >= notesAt[0].text && pp.fcFirst < notesAt[notesAt.length - 1].text;
  const body = [];
  const heads = { header: { default: null, first: null }, footer: { default: null, first: null } };
  let firstPage = false;
  for (const pp of paps) {
    if (pp.fcFirst >= fcMac) break;
    if (pp.note || inNote(pp)) continue;
    if (pp.runningHead) {
      const which = pp.runningHead.footer ? 'footer' : 'header';
      const para = paragraph(pp);
      if (pp.runningHead.every) (heads[which].default ||= []).push(para);
      if (pp.runningHead.first) { (heads[which].first ||= []).push(para); firstPage = true; }
      continue;
    }
    const para = pp.picture ? pictureParagraph(pp) : paragraph(pp);
    if (para) body.push(para);
  }
  // Write prints its running heads on the first page only when told to.
  const hasHeads = heads.header.default || heads.footer.default;
  if (hasHeads && !firstPage) sep.titlePage = !dos;
  if (heads.header.first || heads.footer.first) sep.titlePage = true;

  const footnote = notesAt.slice(0, -1).map((n, i) => {
    const next = notesAt[i + 1].text;
    const paras = paps.filter((pp) => pp.fcFirst >= n.text && pp.fcFirst < next).map(paragraph);
    return { id: i + 1, blocks: groupTables(paras) };
  });

  const toBlocks = (paras) => (paras ? groupTables(paras) : null);
  const normal = { istd: 0, sti: 0, name: 'Normal', kind: 'paragraph', base: null, next: null, pap: {}, chp: { size: 24, font: dos ? dosFont(0) : writeFonts[0] } };
  const hf = {
    headers: Object.fromEntries(Object.entries(heads.header).filter(([, v]) => v).map(([k, v]) => [k, toBlocks(v)])),
    footers: Object.fromEntries(Object.entries(heads.footer).filter(([, v]) => v).map(([k, v]) => [k, toBlocks(v)])),
  };
  return {
    format: dos ? 'worddos' : 'write',
    nFib: 0,
    defaultFonts: [normal.chp.font, normal.chp.font, normal.chp.font],
    evenAndOdd: false,
    styles: [normal],
    fonts,
    lists: { lists: [], lfos: [] },
    sections: [{ cpEnd: fcMac, sep, headers: hf.headers, footers: hf.footers }],
    body: groupTables(body.map((p) => ({ ...p, pap: { ...p.pap } }))),
    notes: { footnote, endnote: [] },
    headers: [hf],
    images,
  };
}

/**
 * A Windows 2/3 device-dependent bitmap, as Write stored it, as a BMP: one
 * plane of one bit per pixel, its rows top to bottom and padded to a
 * word; a BMP's rows run bottom to top, padded to four bytes.
 */
function monochromeBmp(data, width, height, rowBytes, planes, bits) {
  if (planes !== 1 || bits !== 1 || !width || !height || data.length < rowBytes * height) return null;
  const stride = Math.ceil(width / 32) * 4;
  const header = 14 + 40 + 8;
  const out = new Uint8Array(header + stride * height);
  const dv = new DataView(out.buffer);
  out[0] = 0x42; out[1] = 0x4d;
  dv.setUint32(2, out.length, true);
  dv.setUint32(10, header, true);
  dv.setUint32(14, 40, true);
  dv.setInt32(18, width, true);
  dv.setInt32(22, height, true);
  dv.setUint16(26, 1, true);
  dv.setUint16(28, 1, true);
  dv.setUint32(46, 2, true);
  out.set([0, 0, 0, 0, 255, 255, 255, 0], 54);
  for (let y = 0; y < height; y++) out.set(data.subarray(y * rowBytes, y * rowBytes + Math.min(rowBytes, stride)), header + (height - 1 - y) * stride);
  return out;
}

