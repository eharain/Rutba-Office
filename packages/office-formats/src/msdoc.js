// Word 97-2003 documents (.doc, .dot) — [MS-DOC].
//
// A .doc is a compound file. Its WordDocument stream opens with the File
// Information Block, which says where everything else is: the text (found
// through a piece table, because a fast-saved file keeps its text in
// pieces scattered through the stream, some in 8-bit and some in 16-bit
// characters), the character and paragraph properties (in 512-byte pages
// of "formatted disk pages", keyed by file position), the style sheet, the
// fonts, the lists, the sections, the headers and footers, the notes, and
// pictures in the Data stream. Most of those live in the table stream —
// 1Table or 0Table, as the FIB says.
//
// Properties are lists of "sprms" — an opcode and an operand — applied on
// top of a style, which is itself sprms on top of its base. Read here into
// plain objects: a paragraph's alignment, indents, spacing, list and table
// membership; a run's bold, italic, size, font, colour and the rest.
//
// The older Word formats — 6.0/95 (also compound files, with a different
// FIB and one-byte sprms), 2.0, 1.x, DOS and Write — are read by
// msdoc-old.js into the same model.
//
// Pure: bytes in, a document model out. Nothing here writes.

import { CompoundFile } from './cfb.js';
import { header as artHeader, children as artChildren, child as artChild, artColour, readFopt, PRESETS, LINES, freeformPath } from './officeart.js';
import { decode1252 } from './codepage.js';

export class DocError extends Error {
  constructor(message) {
    super(message);
    this.name = 'DocError';
  }
}

const u8 = (b, at) => b[at] ?? 0;
const u16 = (b, at) => (b[at] ?? 0) | ((b[at + 1] ?? 0) << 8);
const i16 = (b, at) => { const v = u16(b, at); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (b, at) => (u16(b, at) | (u16(b, at + 2) << 16)) >>> 0;
const i32 = (b, at) => u32(b, at) | 0;

// The FibRgFcLcb97 pairs this reader uses, by index — [MS-DOC] 2.5.6.
const FC = {
  Stshf: 1, PlcffndRef: 2, PlcffndTxt: 3, PlcfandRef: 4, PlcfandTxt: 5, PlcfSed: 6, PlcfHdd: 11,
  PlcfBteChpx: 12, PlcfBtePapx: 13, SttbfFfn: 15, PlcfFldMom: 16, SttbfBkmk: 21, PlcfBkf: 22, PlcfBkl: 23,
  Dop: 31, Clx: 33, GrpXstAtnOwners: 36, PlcSpaMom: 40, PlcfendRef: 46, PlcfendTxt: 47, DggInfo: 50,
  PlcSpaHdr: 41, PlcftxbxTxt: 56, PlcfHdrtxbxTxt: 57, PlfLst: 73, PlfLfo: 74,
};

/* ── the FIB and the text ─────────────────────────────────────────────── */

/**
 * The File Information Block: the version, the flags, the character counts
 * of each part of the text, and where each structure is.
 */
export function readFib(wd) {
  const wIdent = u16(wd, 0);
  if (wIdent !== 0xa5ec) throw new DocError('not a Word 97-2003 document (its WordDocument stream does not begin with a Word header)');
  const nFib = u16(wd, 2);
  const flags = u16(wd, 10);
  const csw = u16(wd, 32);
  const lwAt = 34 + csw * 2;
  const cslw = u16(wd, lwAt);
  const lw = lwAt + 2;
  const fcLcbAt = lw + cslw * 4;
  const cbRgFcLcb = u16(wd, fcLcbAt);
  const pairsAt = fcLcbAt + 2;
  const pair = (i) => (i < cbRgFcLcb ? { fc: u32(wd, pairsAt + i * 8), lcb: u32(wd, pairsAt + i * 8 + 4) } : { fc: 0, lcb: 0 });
  // A FIB newer than Word 97's names its real version in FibRgCswNew, after the pairs.
  const cswNewAt = pairsAt + cbRgFcLcb * 8;
  const cswNew = u16(wd, cswNewAt);
  const nFibNew = cswNew ? u16(wd, cswNewAt + 2) : 0;
  return {
    nFib: nFibNew || nFib,
    lid: u16(wd, 6),
    fDot: Boolean(flags & 0x0001),
    fComplex: Boolean(flags & 0x0004),
    fEncrypted: Boolean(flags & 0x0100),
    fWhichTblStm: (flags & 0x0200) ? 1 : 0,
    fObfuscated: Boolean(flags & 0x8000),
    lKey: u32(wd, 14),
    ccpText: i32(wd, lw + 12),
    ccpFtn: i32(wd, lw + 16),
    ccpHdd: i32(wd, lw + 20),
    ccpMcr: i32(wd, lw + 24),
    ccpAtn: i32(wd, lw + 28),
    ccpEdn: i32(wd, lw + 32),
    ccpTxbx: i32(wd, lw + 36),
    ccpHdrTxbx: i32(wd, lw + 40),
    pair,
  };
}

/**
 * The piece table, from the Clx: the text as pieces, each a run of
 * character positions (CPs) stored from a file position (FC), in 8-bit
 * Windows-1252 ("compressed") or UTF-16. And the property lists the Clx
 * carries for pieces a fast save changed.
 */
export function readPieces(table, clx) {
  const grpprls = [];
  let at = clx.fc;
  const end = clx.fc + clx.lcb;
  while (at < end && table[at] === 0x01) {
    const cb = u16(table, at + 1);
    grpprls.push(table.subarray(at + 3, at + 3 + cb));
    at += 3 + cb;
  }
  if (table[at] !== 0x02) throw new DocError('the document\'s piece table is damaged');
  const lcb = u32(table, at + 1);
  const plc = at + 5;
  const n = entriesIn(lcb, 12, plc, table.length);
  const pieces = [];
  for (let i = 0; i < n; i++) {
    const cpStart = u32(table, plc + i * 4);
    const cpEnd = u32(table, plc + (i + 1) * 4);
    const pcd = plc + (n + 1) * 4 + i * 8;
    const fcRaw = u32(table, pcd + 2);
    const compressed = Boolean(fcRaw & 0x40000000);
    const fc = fcRaw & 0x3fffffff;
    const prm = u16(table, pcd + 6);
    pieces.push({ cpStart, cpEnd, compressed, fc: compressed ? fc / 2 : fc, prm });
  }
  return { pieces, grpprls };
}

/** Characters [cpStart, cpEnd) of the whole text, through the piece table. */
export function textOf(wd, pieces, cpStart, cpEnd) {
  let out = '';
  for (const p of pieces) {
    const a = Math.max(cpStart, p.cpStart);
    const b = Math.min(cpEnd, p.cpEnd);
    if (a >= b) continue;
    const off = a - p.cpStart;
    if (p.compressed) out += decode1252(wd.subarray(p.fc + off, p.fc + off + (b - a)));
    else {
      const start = p.fc + off * 2;
      const codes = new Array(b - a);
      for (let i = 0; i < b - a; i++) codes[i] = u16(wd, start + i * 2);
      for (let i = 0; i < codes.length; i += 8192) out += String.fromCharCode(...codes.slice(i, i + 8192));
    }
  }
  return out;
}

/* ── sprms ────────────────────────────────────────────────────────────── */

/** The operand length of a sprm whose size its opcode does not fix (spra 6). */
function variableLength(op, b, at) {
  if (op === 0xd608 || op === 0xd606) return { skip: 2, len: Math.max(0, u16(b, at) - 1) };
  if (op === 0xc615) {
    const cb = u8(b, at);
    if (cb !== 255) return { skip: 1, len: cb };
    // A PChgTabs whose own length overflowed a byte: count its parts.
    const del = u8(b, at + 1);
    const addAt = at + 2 + del * 4;
    const add = u8(b, addAt);
    return { skip: 1, len: 1 + del * 4 + 1 + add * 3 };
  }
  return { skip: 1, len: u8(b, at) };
}

/** Each sprm in a property list: its opcode and its operand bytes. */
export function* sprms(grpprl) {
  let i = 0;
  while (i + 2 <= grpprl.length) {
    const op = u16(grpprl, i);
    i += 2;
    const spra = op >> 13;
    let skip = 0;
    let len;
    if (spra === 0 || spra === 1) len = 1;
    else if (spra === 2 || spra === 4 || spra === 5) len = 2;
    else if (spra === 3) len = 4;
    else if (spra === 7) len = 3;
    else ({ skip, len } = variableLength(op, grpprl, i));
    if (i + skip + len > grpprl.length + 1) return;
    yield { op, operand: grpprl.subarray(i + skip, i + skip + len) };
    i += skip + len;
  }
}

// A toggle's operand: 0 and 1 say, 0x80 and 0x81 say "as the style" and "not as the style".
const toggle = (v, base) => (v === 0 ? false : v === 1 ? true : v === 0x80 ? Boolean(base) : v === 0x81 ? !base : Boolean(v));

/** A COLORREF: three bytes of colour and one saying "automatic". */
const colorref = (b, at) => (u8(b, at + 3) === 0xff ? null : [u8(b, at), u8(b, at + 1), u8(b, at + 2)].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase());

/** Word's sixteen-colour palette, by index (Ico). */
export const ICO = [null, '000000', '0000FF', '00FFFF', '00FF00', 'FF00FF', 'FF0000', 'FFFF00', 'FFFFFF', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0'];

/** A Brc80 (4 bytes) or a Brc (8 bytes) as { widthEighths, type, colour, space }. */
function brc(b, at, wide) {
  if (wide) {
    const type = u8(b, at + 5);
    if (!type || type === 0xff) return null;
    return { width: u8(b, at + 4), type, colour: colorref(b, at), space: u8(b, at + 6) & 0x1f };
  }
  if (u32(b, at) === 0xffffffff) return null;
  const type = u8(b, at + 1);
  if (!type) return null;
  return { width: u8(b, at), type, colour: ICO[u8(b, at + 2)] ?? null, space: u8(b, at + 3) & 0x1f };
}

/** A Shd80 (two bytes, palette colours) or an Shd (ten bytes, COLORREFs) as a fill colour or null. */
function shd(b, at, wide) {
  if (wide) {
    const fore = colorref(b, at);
    const back = colorref(b, at + 4);
    const ipat = u16(b, at + 8);
    if (ipat === 0xffff) return null;
    return ipat === 1 ? fore : ipat === 0 ? back : back ?? fore;
  }
  const v = u16(b, at);
  const fore = ICO[v & 0x1f] ?? null;
  const back = ICO[(v >> 5) & 0x1f] ?? null;
  const ipat = (v >> 10) & 0x3f;
  return ipat === 1 ? fore : ipat === 0 ? back : back ?? fore;
}

/** The character properties a run starts from, before any style. */
export const DEFAULT_CHP = Object.freeze({ size: 20, font: 0, fontFE: 0, fontOther: 0 });

/** Apply character sprms to a CHP, on top of `base` (what the style says, for the toggles). */
export function applyChp(chp, grpprl, base = chp) {
  for (const { op, operand: o } of sprms(grpprl)) {
    switch (op) {
      case 0x0835: chp.bold = toggle(o[0], base.bold); break;
      case 0x0836: chp.italic = toggle(o[0], base.italic); break;
      case 0x0837: chp.strike = toggle(o[0], base.strike); break;
      case 0x0838: chp.outline = toggle(o[0], base.outline); break;
      case 0x0839: chp.shadow = toggle(o[0], base.shadow); break;
      case 0x083a: chp.smallCaps = toggle(o[0], base.smallCaps); break;
      case 0x083b: chp.caps = toggle(o[0], base.caps); break;
      case 0x083c: chp.vanish = toggle(o[0], base.vanish); break;
      case 0x2a53: chp.dstrike = toggle(o[0], base.dstrike); break;
      case 0x085c: chp.boldBi = toggle(o[0], base.boldBi); break;
      case 0x085d: chp.italicBi = toggle(o[0], base.italicBi); break;
      case 0x085a: chp.rtl = toggle(o[0], base.rtl); break;
      case 0x2a3e: chp.underline = o[0]; break;
      case 0x2a42: chp.colour = ICO[o[0]] ?? null; chp.colourSet = true; break;
      case 0x6870: chp.colour = colorref(o, 0); chp.colourSet = true; break;
      case 0x4a43: chp.size = u16(o, 0); break;
      case 0x4a61: chp.sizeBi = u16(o, 0); break;
      case 0x2a48: chp.iss = o[0]; break;
      case 0x4a4f: chp.font = u16(o, 0); break;
      case 0x4a50: chp.fontFE = u16(o, 0); break;
      case 0x4a51: chp.fontOther = u16(o, 0); break;
      case 0x4a30: chp.istd = u16(o, 0); break;
      case 0x2a0c: chp.highlight = ICO[o[0]] ?? null; break;
      case 0x0855: chp.spec = o[0] !== 0; break;
      case 0x6a03: chp.picLocation = u32(o, 0); break;
      case 0x0806: chp.data = o[0] !== 0; break;
      case 0x080a: chp.ole2 = o[0] !== 0; break;
      case 0x0801: chp.rmarkIns = o[0] !== 0; break;
      case 0x0800: chp.rmarkDel = o[0] !== 0; break;
      case 0x4804: chp.ibstRMark = u16(o, 0); break;
      case 0x6a09: chp.symbol = { font: u16(o, 0), char: u16(o, 2) }; break;
      case 0x8840: chp.spacing = i16(o, 0); break;
      case 0xca71: chp.shading = shd(o, 1, true); break;
      case 0x4866: chp.shading = shd(o, 0, false); break;
      case 0x2a33: { // sprmCPlain: back to the style's own character properties
        const keep = { istd: chp.istd, spec: chp.spec, picLocation: chp.picLocation };
        for (const k of Object.keys(chp)) delete chp[k];
        Object.assign(chp, base, keep);
        break;
      }
      default: break;
    }
  }
  return chp;
}

/** Apply paragraph sprms to a PAP. `data` is the Data stream, for a PAPX too large for its page. */
export function applyPap(pap, grpprl, data = null) {
  for (const { op, operand: o } of sprms(grpprl)) {
    switch (op) {
      case 0x4600: pap.istd = u16(o, 0); break;
      case 0x2403: case 0x2461: pap.jc = o[0]; break;
      case 0x2405: pap.keep = o[0] !== 0; break;
      case 0x2406: pap.keepNext = o[0] !== 0; break;
      case 0x2407: pap.pageBreakBefore = o[0] !== 0; break;
      case 0x2431: pap.widowControl = o[0] !== 0; break;
      // Word 2007's "don't add space between paragraphs of the same style".
      case 0x246d: pap.contextualSpacing = o[0] !== 0; break;
      // Word 6's numbering, which Word 97 still reads: the level kind, and how the number looks.
      case 0x240d: pap.nLvlAnm = o[0]; break;
      case 0xc63e: pap.anld = readAnld(o); break;
      case 0x260a: pap.ilvl = o[0]; break;
      case 0x460b: pap.ilfo = u16(o, 0); break;
      case 0x840e: case 0x845d: pap.right = i16(o, 0); break;
      case 0x840f: case 0x845e: pap.left = i16(o, 0); break;
      case 0x8411: case 0x8460: pap.firstLine = i16(o, 0); break;
      case 0x4610: pap.left = (pap.left || 0) + i16(o, 0); break;
      case 0x6412: pap.line = { dya: i16(o, 0), mult: i16(o, 2) !== 0 }; break;
      case 0xa413: pap.before = u16(o, 0); break;
      case 0xa414: pap.after = u16(o, 0); break;
      case 0x2416: pap.inTable = o[0] !== 0; break;
      case 0x2417: pap.ttp = o[0] !== 0; break;
      case 0x6649: pap.itap = i32(o, 0); break;
      case 0x244b: pap.innerCell = o[0] !== 0; break;
      case 0x244c: pap.innerTtp = o[0] !== 0; break;
      case 0x2441: pap.bidi = o[0] !== 0; break;
      case 0x2640: pap.outLvl = o[0]; break;
      case 0x442d: pap.shading = shd(o, 0, false); break;
      case 0xc64d: pap.shading = shd(o, 1, true); break;
      case 0x6424: (pap.borders ||= {}).top = brc(o, 0, false); break;
      case 0x6425: (pap.borders ||= {}).left = brc(o, 0, false); break;
      case 0x6426: (pap.borders ||= {}).bottom = brc(o, 0, false); break;
      case 0x6427: (pap.borders ||= {}).right = brc(o, 0, false); break;
      case 0xc64e: (pap.borders ||= {}).top = brc(o, 1, true); break;
      case 0xc64f: (pap.borders ||= {}).left = brc(o, 1, true); break;
      case 0xc650: (pap.borders ||= {}).bottom = brc(o, 1, true); break;
      case 0xc651: (pap.borders ||= {}).right = brc(o, 1, true); break;
      case 0xc60d: case 0xc615: applyTabs(pap, o); break;
      case 0x6646: // sprmPHugePapx: the real list is in the Data stream
        if (data) {
          const at = u32(o, 0);
          const cb = u16(data, at);
          applyPap(pap, data.subarray(at + 2, at + 2 + cb), null);
        }
        break;
      // Table properties ride the row's last paragraph.
      case 0xd608: pap.tap = { ...(pap.tap || {}), ...tdefTable(o) }; break;
      case 0xd605: (pap.tap ||= {}).borders = tableBorders(o, false); break;
      case 0xd613: (pap.tap ||= {}).borders = tableBorders(o, true); break;
      case 0x9407: (pap.tap ||= {}).rowHeight = i16(o, 0); break;
      case 0x3404: (pap.tap ||= {}).header = o[0] !== 0; break;
      case 0x9601: (pap.tap ||= {}).left = i16(o, 0); break;
      case 0x9602: (pap.tap ||= {}).gapHalf = i16(o, 0); break;
      case 0x5400: case 0x548a: (pap.tap ||= {}).jc = u16(o, 0); break;
      case 0xd612: (pap.tap ||= {}).cellShading = cellShading(o); break;
      default: break;
    }
  }
  return pap;
}

/**
 * An ANLV — one level of Word 6 numbering: the number's format, how many
 * characters of text go before and after it, its alignment, whether it
 * shows the levels above and hangs, its font, start and indent.
 */
function readAnlv(b, at) {
  const bits1 = u8(b, at + 3);
  const bits2 = u8(b, at + 4);
  return {
    nfc: u8(b, at), before: u8(b, at + 1), after: u8(b, at + 2),
    jc: bits1 & 3, prev: Boolean(bits1 & 4), hang: Boolean(bits1 & 8),
    bold: bits1 & 0x10 ? Boolean(bits2 & 0x08) : undefined,
    italic: bits1 & 0x20 ? Boolean(bits2 & 0x10) : undefined,
    font: u16(b, at + 6), size: u16(b, at + 8), start: u16(b, at + 10), indent: i16(b, at + 12), space: u16(b, at + 14),
  };
}

/** sprmPAnld: an ANLV, three flags, and the text round the number (32 UTF-16 characters). */
function readAnld(o) {
  const level = readAnlv(o, 0);
  let text = '';
  for (let i = 0; i < 32 && 20 + i * 2 + 1 < o.length; i++) text += String.fromCharCode(u16(o, 20 + i * 2));
  return { ...level, text };
}

/** sprmSOlstAnm: the section's heading numbering — nine ANLVs and the text they share. */
function readOlst(o) {
  const levels = [];
  for (let i = 0; i < 9; i++) levels.push(readAnlv(o, i * 16));
  let text = '';
  for (let i = 0; i < 32 && 148 + i * 2 + 1 < o.length; i++) text += String.fromCharCode(u16(o, 148 + i * 2));
  return { levels, restart: u8(o, 144) !== 0, text };
}

/** Tab stops added and taken away. */
function applyTabs(pap, o) {
  const tabs = new Map((pap.tabs || []).map((t) => [t.pos, t]));
  let at = 0;
  const del = u8(o, at++);
  for (let i = 0; i < del; i++) tabs.delete(i16(o, at + i * 2));
  at += del * 2;
  // PChgTabs carries a "close" tolerance per deleted tab; PChgTabsPapx does not.
  if (o.length > at && o.length >= at + del * 2 + 1 && (o.length - (at + del * 2 + 1)) % 3 === 0 && del) at += del * 2;
  const add = u8(o, at++);
  for (let i = 0; i < add; i++) {
    const pos = i16(o, at + i * 2);
    const tbd = u8(o, at + add * 2 + i);
    tabs.set(pos, { pos, jc: tbd & 7, leader: (tbd >> 3) & 7 });
  }
  pap.tabs = [...tabs.values()].sort((a, b) => a.pos - b.pos);
}

/** sprmTDefTable: the cells' edges, and each cell's merging, width and borders. */
function tdefTable(o) {
  const n = u8(o, 0);
  const centers = [];
  for (let i = 0; i <= n; i++) centers.push(i16(o, 1 + i * 2));
  const cells = [];
  const tcAt = 1 + (n + 1) * 2;
  for (let i = 0; i < n; i++) {
    const at = tcAt + i * 20;
    const has = at + 20 <= o.length;
    const grf = has ? u16(o, at) : 0;
    cells.push({
      width: centers[i + 1] - centers[i],
      hMerge: grf & 3,
      vMerge: (grf >> 5) & 3,
      vAlign: (grf >> 7) & 3,
      borders: has ? { top: brc(o, at + 4, false), left: brc(o, at + 8, false), bottom: brc(o, at + 12, false), right: brc(o, at + 16, false) } : null,
    });
  }
  return { cells, centers };
}

/** The table's own borders: top, left, bottom, right, and the inside lines. */
function tableBorders(o, wide) {
  const size = wide ? 8 : 4;
  const names = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
  const out = {};
  names.forEach((name, i) => { out[name] = brc(o, i * size, wide); });
  return out;
}

/** sprmTDefTableShd: one Shd per cell. */
function cellShading(o) {
  const out = [];
  for (let at = 0; at + 10 <= o.length; at += 10) out.push(shd(o, at, true));
  return out;
}

/* ── formatted disk pages ─────────────────────────────────────────────── */

/** Every CHPX run in the document, by file position: [{ fcStart, fcEnd, grpprl }]. */
function readChpx(wd, table, plc) {
  const out = [];
  if (!plc.lcb) return out;
  const n = entriesIn(plc.lcb, 8, plc.fc, table.length);
  for (let i = 0; i < n; i++) {
    const pn = u32(table, plc.fc + (n + 1) * 4 + i * 4) & 0x3fffff;
    const page = pn * 512;
    const crun = u8(wd, page + 511);
    for (let r = 0; r < crun; r++) {
      const fcStart = u32(wd, page + r * 4);
      const fcEnd = u32(wd, page + (r + 1) * 4);
      const off = u8(wd, page + (crun + 1) * 4 + r) * 2;
      const grpprl = off ? wd.subarray(page + off + 1, page + off + 1 + u8(wd, page + off)) : new Uint8Array(0);
      out.push({ fcStart, fcEnd, grpprl });
    }
  }
  return out.sort((a, b) => a.fcStart - b.fcStart);
}

/** Every PAPX in the document, by file position: [{ fcStart, fcEnd, istd, grpprl }]. */
function readPapx(wd, table, plc) {
  const out = [];
  if (!plc.lcb) return out;
  const n = entriesIn(plc.lcb, 8, plc.fc, table.length);
  for (let i = 0; i < n; i++) {
    const pn = u32(table, plc.fc + (n + 1) * 4 + i * 4) & 0x3fffff;
    const page = pn * 512;
    const cpara = u8(wd, page + 511);
    for (let r = 0; r < cpara; r++) {
      const fcStart = u32(wd, page + r * 4);
      const fcEnd = u32(wd, page + (r + 1) * 4);
      const off = u8(wd, page + (cpara + 1) * 4 + r * 13) * 2;
      if (!off) { out.push({ fcStart, fcEnd, istd: 0, grpprl: new Uint8Array(0) }); continue; }
      let cb = u8(wd, page + off);
      let start = page + off + 1;
      let len;
      if (cb === 0) { cb = u8(wd, page + off + 1); start = page + off + 2; len = cb * 2; } else len = cb * 2 - 1;
      out.push({ fcStart, fcEnd, istd: u16(wd, start), grpprl: wd.subarray(start + 2, start + len) });
    }
  }
  return out.sort((a, b) => a.fcStart - b.fcStart);
}

/** The entry whose [fcStart, fcEnd) holds `fc`, by binary search. */
function findFc(list, fc) {
  let lo = 0;
  let hi = list.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const e = list[mid];
    if (fc < e.fcStart) hi = mid - 1;
    else if (fc >= e.fcEnd) lo = mid + 1;
    else return e;
  }
  return null;
}

/* ── styles, fonts, lists ─────────────────────────────────────────────── */

/**
 * The style sheet: each style's name, kind, base and next style, and its
 * paragraph and character sprms — resolved down its chain of bases, so a
 * style's pap and chp are what it means on its own. Word 6/95's style
 * sheet has the same shape with 8-bit names and its own sprms: `readName`
 * reads its names and `translate` turns its sprms into Word 97's.
 */
export function readStyles(table, plc, readName = readXstz, translate = null) {
  const styles = [];
  if (!plc.lcb) return { styles, defaultFonts: [0, 0, 0] };
  const at0 = plc.fc;
  const cbStshi = u16(table, at0);
  const shi = at0 + 2;
  const cstd = u16(table, shi);
  const cbBase = u16(table, shi + 2);
  // Word 97 names three default fonts; Word 6 one, for every script.
  const defaultFonts = cbStshi >= 18 ? [u16(table, shi + 12), u16(table, shi + 14), u16(table, shi + 16)]
    : cbStshi >= 14 ? Array(3).fill(u16(table, shi + 12)) : [0, 0, 0];
  let at = shi + cbStshi;
  for (let istd = 0; istd < cstd; istd++) {
    const cb = u16(table, at);
    const std = at + 2;
    at += 2 + cb;
    if (!cb) { styles.push(null); continue; }
    const sti = u16(table, std) & 0x0fff;
    const stk = u16(table, std + 2) & 0x000f;
    const istdBase = u16(table, std + 2) >> 4;
    const cupx = u16(table, std + 4) & 0x000f;
    const istdNext = u16(table, std + 4) >> 4;
    let p = std + cbBase;
    const { name, end } = readName(table, p);
    p = end;
    if ((p - std) % 2) p++;
    const upx = [];
    for (let i = 0; i < cupx && p + 2 <= std + cb; i++) {
      const cbUpx = u16(table, p);
      upx.push(table.subarray(p + 2, p + 2 + cbUpx));
      p += 2 + cbUpx;
      if (cbUpx % 2) p++;
    }
    const kind = stk === 1 ? 'paragraph' : stk === 2 ? 'character' : stk === 3 ? 'table' : stk === 4 ? 'numbering' : 'paragraph';
    const tr = (g) => (g && translate ? translate(g) : g);
    styles.push({
      istd, sti, name, kind,
      base: istdBase === 0x0fff ? null : istdBase,
      next: istdNext === 0x0fff ? null : istdNext,
      papx: tr(kind === 'paragraph' && upx[0] ? upx[0].subarray(2) : null),
      chpx: tr(kind === 'paragraph' ? upx[1] ?? null : kind === 'character' ? upx[0] ?? null : null),
    });
  }
  resolveStyles(styles, defaultFonts);
  return { styles, defaultFonts };
}

/** An Xstz: a 16-bit count, that many UTF-16 characters, and a terminating zero. */
function readXstz(b, at) {
  const n = u16(b, at);
  let name = '';
  for (let i = 0; i < n; i++) name += String.fromCharCode(u16(b, at + 2 + i * 2));
  return { name, end: at + 2 + n * 2 + 2 };
}

/** Each style's pap and chp, down its chain of bases. */
export function resolveStyles(styles, defaultFonts) {
  const done = new Set();
  const resolve = (st, depth = 0) => {
    if (!st || done.has(st.istd)) return;
    const base = st.base != null && depth < 16 ? styles[st.base] : null;
    if (base) resolve(base, depth + 1);
    const basePap = base?.pap ?? {};
    // A character style is only what it changes over the paragraph's look: it starts from nothing.
    const baseChp = base?.chp ?? (st.kind === 'character' ? {} : { ...DEFAULT_CHP, font: defaultFonts[0], fontFE: defaultFonts[1], fontOther: defaultFonts[2] });
    st.pap = applyPap({ ...basePap }, st.papx ?? new Uint8Array(0));
    st.chp = applyChp({ ...baseChp }, st.chpx ?? new Uint8Array(0), baseChp);
    delete st.chp.istd;
    done.add(st.istd);
  };
  for (const st of styles) resolve(st);
}

/** The font table: each font's name, by index (ftc). */
export function readFonts(table, plc) {
  const fonts = [];
  if (!plc.lcb) return fonts;
  let at = plc.fc;
  const count = u16(table, at);
  const extended = count === 0xffff;
  const n = extended ? u16(table, at + 2) : count;
  at += extended ? 6 : 4;
  for (let i = 0; i < n; i++) {
    const cb = u8(table, at) + 1;
    let name = '';
    for (let p = at + 40; p + 1 < at + cb; p += 2) {
      const c = u16(table, p);
      if (!c) break;
      name += String.fromCharCode(c);
    }
    fonts.push({ name, family: (u8(table, at + 1) >> 4) & 7, charset: u8(table, at + 4) });
    at += cb;
  }
  return fonts;
}

/**
 * The lists: each list's levels (start, number format, text with
 * placeholders for the level numbers, indents) and the list overrides
 * paragraphs point at by number.
 */
export function readLists(table, plfLst, plfLfo) {
  const lists = [];
  if (plfLst.lcb) {
    let at = plfLst.fc;
    const n = i16(table, at);
    at += 2;
    const lstf = [];
    for (let i = 0; i < n; i++) {
      const p = at + i * 28;
      lstf.push({ lsid: i32(table, p), simple: Boolean(u8(table, p + 26) & 1) });
    }
    at += n * 28;
    for (const l of lstf) {
      const levels = [];
      for (let lv = 0; lv < (l.simple ? 1 : 9); lv++) {
        const start = i32(table, at);
        const nfc = u8(table, at + 4);
        const jc = u8(table, at + 5) & 3;
        const ixchFollow = u8(table, at + 15);
        const cbChpx = u8(table, at + 24);
        const cbPapx = u8(table, at + 25);
        const papx = table.subarray(at + 28, at + 28 + cbPapx);
        const chpx = table.subarray(at + 28 + cbPapx, at + 28 + cbPapx + cbChpx);
        let p = at + 28 + cbPapx + cbChpx;
        const cch = u16(table, p);
        const codes = [];
        for (let c = 0; c < cch; c++) codes.push(u16(table, p + 2 + c * 2));
        p += 2 + cch * 2;
        levels.push({ start, nfc, jc, follow: ixchFollow, pap: applyPap({}, papx), chp: applyChp({}, chpx), codes });
        at = p;
      }
      lists.push({ lsid: l.lsid, levels });
    }
  }
  const lfos = [];
  if (plfLfo.lcb) {
    let at = plfLfo.fc;
    const n = u32(table, at);
    at += 4;
    for (let i = 0; i < n; i++) lfos.push({ lsid: i32(table, at + i * 16), overrides: u8(table, at + i * 16 + 12) });
  }
  return { lists, lfos };
}

/* ── sections, headers, notes, pictures ───────────────────────────────── */

/** The page a section is laid on, with Word's defaults where its SEPX says nothing. */
function readSep(grpprl) {
  const sep = { width: 12240, height: 15840, left: 1800, right: 1800, top: 1440, bottom: 1440, header: 720, footer: 720, landscape: false, titlePage: false, columns: 1, break: 2 };
  for (const { op, operand: o } of sprms(grpprl)) {
    switch (op) {
      case 0xb01f: sep.width = u16(o, 0); break;
      case 0xb020: sep.height = u16(o, 0); break;
      case 0xb021: sep.left = u16(o, 0); break;
      case 0xb022: sep.right = u16(o, 0); break;
      case 0x9023: sep.top = i16(o, 0); break;
      case 0x9024: sep.bottom = i16(o, 0); break;
      case 0xb017: sep.header = u16(o, 0); break;
      case 0xb018: sep.footer = u16(o, 0); break;
      case 0x301d: sep.landscape = o[0] === 2; break;
      case 0x300a: sep.titlePage = o[0] !== 0; break;
      case 0x500b: sep.columns = u16(o, 0) + 1; break;
      case 0x3009: sep.break = o[0]; break;
      // Word 6 and older: which headers and footers this section has, and its heading numbering.
      case 0x3014: sep.grpfIhdt = o[0]; break;
      case 0xd202: sep.olst = readOlst(o); break;
      default: break;
    }
  }
  return sep;
}

/**
 * How many entries a table of `per` bytes each, `lcb` long at `fc`, can hold
 * in a stream of `length` bytes. The count in the file is a 32-bit field
 * a damaged file sets to billions; it is never allowed past the stream.
 */
export function entriesIn(lcb, per, fc, length) {
  const claimed = Math.floor((lcb - 4) / per);
  const room = Math.floor((length - fc - 4) / per);
  return Math.max(0, Math.min(claimed, room));
}

/** A PLC: n+1 positions, then n entries of `size` bytes. */
function readPlc(table, plc, size) {
  if (!plc.lcb) return { cps: [], entries: [] };
  const n = entriesIn(plc.lcb, 4 + size, plc.fc, table.length);
  const cps = [];
  for (let i = 0; i <= n; i++) cps.push(i32(table, plc.fc + i * 4));
  const entries = [];
  for (let i = 0; i < n; i++) entries.push(plc.fc + (n + 1) * 4 + i * size);
  return { cps, entries };
}

/**
 * A picture held in the Data stream at `at`: the PICF says its size on the
 * page, and the Office Art records after it hold the picture itself — a
 * PNG, JPEG or DIB as it is, an EMF or WMF as the metafile.
 */
export function readPicture(data, at) {
  if (!data.length || at + 68 > data.length) return null;
  const lcb = u32(data, at);
  const cbHeader = u16(data, at + 4);
  const mm = i16(data, at + 6);
  const dxaGoal = i16(data, at + 28);
  const dyaGoal = i16(data, at + 30);
  const mx = u16(data, at + 32) || 1000;
  const my = u16(data, at + 34) || 1000;
  const end = Math.min(data.length, at + lcb);
  let p = at + cbHeader;
  if (mm === 0x66) p += 1 + u8(data, p);
  const blip = findBlip(data, p, end);
  if (!blip) return null;
  return { ...blip, widthTwips: Math.round((dxaGoal * mx) / 1000), heightTwips: Math.round((dyaGoal * my) / 1000) };
}

const BLIP_TYPES = {
  0xf01a: { type: 'image/x-emf', ext: 'emf', meta: true }, 0xf01b: { type: 'image/x-wmf', ext: 'wmf', meta: true },
  0xf01c: { type: 'image/pict', ext: 'pict', meta: true }, 0xf01d: { type: 'image/jpeg', ext: 'jpeg' },
  0xf01e: { type: 'image/png', ext: 'png' }, 0xf01f: { type: 'image/bmp', ext: 'bmp', dib: true },
  0xf029: { type: 'image/tiff', ext: 'tiff' }, 0xf02a: { type: 'image/jpeg', ext: 'jpeg' },
};
// Instances that carry a second 16-byte id after the first.
const TWO_UIDS = new Set([0x3d5, 0x217, 0x543, 0x46b, 0x6e3, 0x6e1, 0x7a9, 0x6e5]);

/**
 * The first picture among Office Art records [from, to): walked as
 * records, a container's children entered, a blip's bytes returned.
 */
export function findBlip(b, from, to) {
  let p = from;
  while (p + 8 <= to) {
    const verInst = u16(b, p);
    const type = u16(b, p + 2);
    const len = u32(b, p + 4);
    const ver = verInst & 0xf;
    const inst = verInst >> 4;
    const body = p + 8;
    if (BLIP_TYPES[type]) return blipBytes(b, body, Math.min(to, body + len), type, inst);
    if (ver === 0xf) {
      const inner = findBlip(b, body, Math.min(to, body + len));
      if (inner) return inner;
    } else if (type === 0xf007) {
      // An FBSE: 36 bytes of header, a name, then the blip itself.
      const cbName = u8(b, body + 33);
      const inner = findBlip(b, body + 36 + cbName, Math.min(to, body + len));
      if (inner) return inner;
    }
    if (len <= 0 && ver !== 0xf) p = body;
    else p = body + len;
  }
  return null;
}

/** A blip record's own bytes, made a file a browser or Word can open. */
function blipBytes(b, body, end, type, inst) {
  const kind = BLIP_TYPES[type];
  let p = body + 16 + (TWO_UIDS.has(inst) ? 16 : 0);
  if (kind.meta) {
    // A metafile header: its size, bounds, size in EMU, saved size, compression.
    const cbSize = u32(b, p);
    const cbSave = u32(b, p + 28);
    const compression = u8(b, p + 32);
    p += 34;
    const raw = b.subarray(p, Math.min(end, p + cbSave));
    return { contentType: kind.type, ext: kind.ext, bytes: raw, deflated: compression === 0, size: cbSize };
  }
  p += 1;
  const bytes = b.subarray(p, end);
  return { contentType: kind.type, ext: kind.ext, bytes: kind.dib ? dibFile(bytes) : bytes };
}

// The DOP's compatibility options, bit by bit, as the w:compat elements Word
// writes for them (LibreOffice's layout of the bits, ww8scan.cxx). The first
// word is Word 6's and later; the second, from Word 2000 on, sits after it.
const COMPAT_1 = {
  0x1: 'noTabHangInd', 0x2: 'noSpaceRaiseLower', 0x4: 'suppressSpBfAfterPgBrk', 0x8: 'wrapTrailSpaces', 0x10: 'printColBlack',
  0x20: 'noColumnBalance', 0x40: 'convMailMergeEsc', 0x80: 'suppressTopSpacing', 0x100: 'useSingleBorderforContiguousCells',
  0x400: 'showBreaksInFrames', 0x800: 'swapBordersFacingPages', 0x10000: 'suppressTopSpacingWP', 0x20000: 'spacingInWholePoints',
  0x40000: 'printBodyTextBeforeHeader', 0x80000: 'noLeading', 0x200000: 'mwSmallCaps', 0x80000000: 'usePrinterMetrics',
};
const COMPAT_2 = {
  0x1: 'shapeLayoutLikeWW8', 0x2: 'footnoteLayoutLikeWW8', 0x4: 'doNotUseHTMLParagraphAutoSpacing', 0x10: 'forgetLastTabAlignment',
  0x20: 'autoSpaceLikeWord95', 0x40: 'alignTablesRowByRow', 0x80: 'layoutRawTableWidth', 0x100: 'layoutTableRowsApart',
  0x200: 'useWord97LineBreakRules', 0x400: 'doNotBreakWrappedTables', 0x800: 'doNotSnapToGridInCell', 0x2000: 'applyBreakingRules',
  0x4000: 'doNotWrapTextWithPunct', 0x8000: 'doNotUseEastAsianBreakRules', 0x10000: 'useWord2002TableStyleRules', 0x20000: 'growAutofit',
  0x40000: 'useNormalStyleForList', 0x80000: 'doNotUseIndentAsNumberingTabStop', 0x100000: 'useAltKinsokuLineBreakRules',
  0x200000: 'allowSpaceOfSameStyleInTable', 0x400000: 'doNotSuppressIndentation', 0x800000: 'doNotAutofitConstrainedTables',
  0x1000000: 'autofitToFirstFixedWidthCell', 0x2000000: 'underlineTabInNumList', 0x4000000: 'displayHangulFixedWidth',
  0x8000000: 'splitPgBreakAndParaMark', 0x10000000: 'doNotVertAlignCellWithSp', 0x20000000: 'doNotBreakConstrainedForcedTable',
  0x40000000: 'doNotVertAlignInTxbx', 0x80000000: 'useAnsiKerningPairs',
};

/** The document's layout rules from its DOP, as w:compat element names. */
function compatOf(table, dop) {
  const bits = (word, map) => Object.entries(map).filter(([bit]) => (word & Number(bit)) >>> 0).map(([, name]) => name);
  const end = dop.fc + dop.lcb;
  if (dop.lcb >= 516 && end <= table.length) return [...bits(u32(table, dop.fc + 508), COMPAT_1), ...bits(u32(table, dop.fc + 512), COMPAT_2)];
  if (dop.lcb >= 88 && end <= table.length) return bits(u32(table, dop.fc + 84), COMPAT_1);
  return [];
}

/** A DIB is a BMP without its 14-byte file header: given one, a browser draws it. */
export function dibFile(bytes) {
  if (u32(bytes, 0) === 12) bytes = windowsDib(bytes);
  const headerSize = u32(bytes, 0);
  const bitCount = u16(bytes, 14);
  const colors = u32(bytes, 32) || (bitCount <= 8 ? 1 << bitCount : 0);
  const off = 14 + headerSize + colors * 4;
  const file = new Uint8Array(14 + bytes.length);
  file[0] = 0x42; file[1] = 0x4d;
  const dv = new DataView(file.buffer);
  dv.setUint32(2, file.length, true);
  dv.setUint32(10, off, true);
  file.set(bytes, 14);
  return file;
}

/**
 * An OS/2 DIB — a 12-byte header, three bytes to a palette colour — as the
 * Windows one every reader of a .bmp knows: Excel 95 and older keep their
 * pictures so.
 */
function windowsDib(core) {
  const bitCount = u16(core, 10);
  const colors = bitCount <= 8 ? 1 << bitCount : 0;
  const pixels = core.subarray(12 + colors * 3);
  const out = new Uint8Array(40 + colors * 4 + pixels.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, 40, true);
  dv.setInt32(4, u16(core, 4), true);
  dv.setInt32(8, u16(core, 6), true);
  dv.setUint16(12, 1, true);
  dv.setUint16(14, bitCount, true);
  dv.setUint32(32, colors, true);
  for (let i = 0; i < colors; i++) out.set(core.subarray(12 + i * 3, 15 + i * 3), 40 + i * 4);
  out.set(pixels, 40 + colors * 4);
  return out;
}

/**
 * The Office Art behind a document's floating drawings (DggInfo): the
 * drawing group's pictures — kept in the WordDocument stream, or in the
 * record itself — then a drawing for the main text and one for the
 * headers, every shape in them by its id: a group with its members, their
 * boxes in the group's own coordinates.
 */
function readOfficeArt(table, pair, wd) {
  const out = { shapes: new Map(), blips: [] };
  if (!pair.lcb || pair.fc + 8 > table.length) return out;
  const end = Math.min(table.length, pair.fc + pair.lcb);
  const dgg = artHeader(table, pair.fc);
  if (dgg.type !== 0xf000) return out;
  for (const c of artChildren(table, artChild(table, dgg, 0xf001))) {
    if (c.type !== 0xf007) { out.blips.push(null); continue; }
    const foDelay = u32(table, c.body + 28);
    const cbName = u8(table, c.body + 33);
    out.blips.push(foDelay !== 0xffffffff && foDelay + 8 <= wd.length
      ? findBlip(wd, foDelay, Math.min(wd.length, foDelay + 8 + u32(wd, foDelay + 4)))
      : findBlip(table, c.body + 36 + cbName, c.end));
  }
  const rect = (h) => ({ left: i32(table, h.body), top: i32(table, h.body + 4), right: i32(table, h.body + 8), bottom: i32(table, h.body + 12) });
  const shapeOf = (sp) => {
    const fsp = artChild(table, sp, 0xf00a);
    if (!fsp) return null;
    const props = readFopt(table, artChild(table, sp, 0xf00b));
    for (const [k, v] of readFopt(table, artChild(table, sp, 0xf122))) if (!props.has(k)) props.set(k, v);
    const anchor = artChild(table, sp, 0xf00f);
    const group = artChild(table, sp, 0xf009);
    return { spid: u32(table, fsp.body), type: fsp.inst, flags: u32(table, fsp.body + 4), props, anchor: anchor ? rect(anchor) : null, group: group ? rect(group) : null };
  };
  const readGroup = (h) => {
    const kids = artChildren(table, h);
    const own = kids[0]?.type === 0xf004 ? shapeOf(kids[0]) : null;
    if (!own) return null;
    own.children = [];
    for (const k of kids.slice(1)) {
      const s = k.type === 0xf003 ? readGroup(k) : k.type === 0xf004 ? shapeOf(k) : null;
      if (s) { own.children.push(s); out.shapes.set(s.spid, s); }
    }
    out.shapes.set(own.spid, own);
    return own;
  };
  // Each drawing after its one-byte label: 0 the main text's, 1 the headers'.
  let p = dgg.end;
  while (p + 9 <= end) {
    const dg = artHeader(table, p + 1);
    if (dg.type !== 0xf002) break;
    for (const k of artChildren(table, dg)) if (k.type === 0xf003) readGroup(k);
    p = dg.end;
  }
  return out;
}

/** A UTF-16 string up to its terminating zero. */
const utf16z = (b) => { let s = ''; for (let i = 0; i + 1 < b.length; i += 2) { const c = u16(b, i); if (!c) break; s += String.fromCharCode(c); } return s; };

/* ── the document ─────────────────────────────────────────────────────── */

/**
 * Read a Word 97-2003 document into a model: styles, fonts, lists, sections
 * with their headers and footers, the body as paragraphs and tables, notes
 * and pictures.
 */
export function readDoc(bytes) {
  const cfb = new CompoundFile(bytes);
  const wdEntry = cfb.find(['WordDocument']);
  if (!wdEntry) throw new DocError('not a Word document (no WordDocument stream)');
  const wd = cfb.read(wdEntry);
  if (u16(wd, 0) === 0xa5dc || (u16(wd, 0) === 0xa5ec && u16(wd, 2) < 0xc1)) {
    const err = new DocError('a Word 6.0/95 document');
    err.older = true;
    throw err;
  }
  const fib = readFib(wd);
  if (fib.fEncrypted) {
    const err = new DocError('This Word document is password-protected.');
    err.encrypted = true;
    throw err;
  }
  const tableEntry = cfb.find([fib.fWhichTblStm ? '1Table' : '0Table']);
  const table = tableEntry ? cfb.read(tableEntry) : new Uint8Array(0);
  const dataEntry = cfb.find(['Data']);
  const data = dataEntry ? cfb.read(dataEntry) : new Uint8Array(0);
  return buildModel({ wd, table, data, fib, cfb });
}

/**
 * The model, from the streams — shared with the readers of the older
 * versions (msdoc-old.js), which hand in what they lay out differently:
 * the FIB, the pieces, the property runs already read from their pages
 * (and their sprms already Word 97's), the style sheet, the fonts, the
 * size of a section entry and how a section's properties are stored, which
 * header and footer stories a section has, and where pictures are.
 */
export function buildModel({
  wd, table, data, fib, pieces = null, grpprls = null, chpxList = null, papxList = null,
  readStylesWith = null, readListsWith = null, readFontsWith = null, readPictureWith = null,
  sedSize = 12, readSepx = null, headerSlots = null, format = 'word97',
}) {
  let pcs = pieces;
  let prls = grpprls;
  if (!pcs) ({ pieces: pcs, grpprls: prls } = readPieces(table, fib.pair(FC.Clx)));
  const chpx = chpxList || readChpx(wd, table, fib.pair(FC.PlcfBteChpx));
  const papx = papxList || readPapx(wd, table, fib.pair(FC.PlcfBtePapx));
  const { styles, defaultFonts } = (readStylesWith || readStyles)(table, fib.pair(FC.Stshf));
  const fonts = (readFontsWith || readFonts)(table, fib.pair(FC.SttbfFfn));
  const lists = (readListsWith || readLists)(table, fib.pair(FC.PlfLst), fib.pair(FC.PlfLfo));
  const images = [];
  const imageAt = new Map();
  const notes = { footnote: [], endnote: [] };

  const styleOf = (istd) => styles[istd] || styles[0] || { pap: {}, chp: { ...DEFAULT_CHP, font: defaultFonts[0] } };

  // The property list a fast save left on a piece, when it is one of the
  // Clx's lists. (A Prm0 — one sprm named by a seven-bit index into a
  // table of its own — is not read: that piece keeps the rest of its
  // formatting.)
  const pieceSprms = (piece) => (piece && piece.prm & 1 ? prls?.[piece.prm >> 1] ?? null : null);

  // Sections first: a section's last character is its section mark, which
  // ends a paragraph as a paragraph mark does.
  const sed = readPlc(table, fib.pair(FC.PlcfSed), sedSize);
  const sectionEnds = new Set(sed.cps.slice(1));

  // Each character's file position, piece by piece.
  const locate = (cp) => {
    for (const p of pcs) if (cp >= p.cpStart && cp < p.cpEnd) return { fc: p.fc + (cp - p.cpStart) * (p.compressed ? 1 : 2), piece: p };
    return null;
  };

  /** The paragraphs of a story [cpStart, cpEnd): each its PAP and its runs of characters with their CHP. */
  const readParagraphs = (cpStart, cpEnd) => {
    const text = textOf(wd, pcs, cpStart, cpEnd);
    const out = [];
    let paraStart = 0;
    for (let i = 0; i < text.length; i++) {
      const ch = text.charCodeAt(i);
      if (ch !== 13 && ch !== 7 && !(ch === 12 && sectionEnds.has(cpStart + i + 1))) continue;
      const markCp = cpStart + i;
      const loc = locate(markCp);
      const px = loc ? findFc(papx, loc.fc) : null;
      const st = styleOf(px?.istd ?? 0);
      const pap = applyPap({ ...st.pap, istd: px?.istd ?? 0 }, px?.grpprl ?? new Uint8Array(0), data);
      const pieceMods = pieceSprms(loc?.piece);
      if (pieceMods) applyPap(pap, pieceMods, data);
      const runs = readRuns(cpStart + paraStart, cpStart + i, text.slice(paraStart, i), styleOf(pap.istd));
      out.push({ pap, runs, mark: ch === 7 ? 'cell' : ch === 12 ? 'section' : 'para', cpEnd: markCp + 1, markChp: chpAt(markCp, styleOf(pap.istd)) });
      paraStart = i + 1;
    }
    if (paraStart < text.length) {
      const st = styleOf(0);
      out.push({ pap: { ...st.pap, istd: 0 }, runs: readRuns(cpStart + paraStart, cpEnd, text.slice(paraStart), st), mark: 'para', cpEnd });
    }
    return out;
  };

  // A run's CHP depends only on its CHPX, its piece and its paragraph's
  // style: worked out once for each such trio, not once per character.
  const chpCache = new Map();
  /** The CHP of the character at `cp`, in a paragraph of style `st`. */
  const chpAt = (cp, st) => {
    const loc = locate(cp);
    const cx = loc ? findFc(chpx, loc.fc) : null;
    const key = (cx ? cx.fcStart : -1) + ':' + (st.istd ?? 'x') + ':' + (loc?.piece?.cpStart ?? -1);
    const hit = chpCache.get(key);
    if (hit) return hit;
    const chp = computeChp(cx, st, loc);
    chpCache.set(key, chp);
    return chp;
  };
  const computeChp = (cx, st, loc) => {
    const base = st.chp ?? { ...DEFAULT_CHP };
    const chp = { ...base };
    // A character style first: the run's sprmCIstd, applied before the rest.
    const grp = cx?.grpprl;
    if (grp && grp.length) {
      for (const { op, operand } of sprms(grp)) {
        if (op === 0x4a30) {
          const cs = styles[u16(operand, 0)];
          if (cs?.kind === 'character' && cs.chpx) applyChp(chp, cs.chpx, base);
          chp.istd = u16(operand, 0);
        }
      }
      applyChp(chp, grp, base);
    }
    const mods = pieceSprms(loc?.piece);
    if (mods) applyChp(chp, mods, base);
    return chp;
  };

  /** The runs of one paragraph's characters: text runs split where the CHP changes, and the special characters as their own runs. */
  const readRuns = (cpStart, cpEnd, text, st) => {
    const runs = [];
    let cur = null;
    let curKey = '';
    for (let i = 0; i < text.length; i++) {
      const cp = cpStart + i;
      const ch = text.charCodeAt(i);
      const chp = chpAt(cp, st);
      const special = specialRun(ch, chp, cp);
      if (special) {
        cur = null;
        curKey = '';
        if (special !== 'skip') runs.push({ ...special, chp });
        continue;
      }
      const key = chpKey(chp);
      if (!cur || key !== curKey) {
        cur = { text: '', chp };
        curKey = key;
        runs.push(cur);
      }
      cur.text += text[i];
    }
    return runs;
  };

  const specialRun = (ch, chp, cp) => {
    switch (ch) {
      case 0x13: return { kind: 'fieldBegin' };
      case 0x14: return { kind: 'fieldSep' };
      case 0x15: return { kind: 'fieldEnd' };
      case 0x09: return { kind: 'tab' };
      case 0x0b: return { kind: 'break', type: 'line' };
      case 0x0c: return { kind: 'break', type: 'page' };
      case 0x0e: return { kind: 'break', type: 'column' };
      case 0x1e: return { kind: 'noBreakHyphen' };
      case 0x1f: return { kind: 'softHyphen' };
      case 0x01:
        if (!chp.spec) return null;
        if (chp.data || chp.ole2) return 'skip';
        return pictureRun(chp);
      case 0x02:
        if (!chp.spec) return null;
        return { kind: 'noteMark', cp };
      case 0x03: case 0x04: return chp.spec ? 'skip' : null;
      case 0x05: return chp.spec ? 'skip' : null;
      case 0x08: { // a floating drawing's anchor
        if (!chp.spec) return null;
        const spa = floats.get(cp);
        const float = spa ? floatOf(spa) : null;
        return float ? { kind: 'float', float } : 'skip';
      }
      case 0x28: case 0xf000:
        if (chp.symbol) return { kind: 'symbol', font: chp.symbol.font, char: chp.symbol.char };
        return null;
      default:
        if (chp.symbol && chp.spec) return { kind: 'symbol', font: chp.symbol.font, char: chp.symbol.char };
        return null;
    }
  };

  const pictureRun = (chp) => {
    const at = chp.picLocation;
    if (at == null) return 'skip';
    if (!imageAt.has(at)) {
      const pic = readPictureWith ? readPictureWith(at) : readPicture(data, at);
      imageAt.set(at, pic ? images.push(pic) - 1 : -1);
    }
    const index = imageAt.get(at);
    return index < 0 ? 'skip' : { kind: 'picture', image: index };
  };

  // The stories, by where each starts in the one stream of characters.
  const base = {
    main: 0,
    footnote: fib.ccpText,
    header: fib.ccpText + fib.ccpFtn,
    annotation: fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpMcr,
    endnote: fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpMcr + fib.ccpAtn,
    textbox: fib.ccpText + fib.ccpFtn + fib.ccpHdd + fib.ccpMcr + fib.ccpAtn + fib.ccpEdn,
  };
  base.headerTextbox = base.textbox + (fib.ccpTxbx || 0);

  // Floating drawings: each anchor's place in the text (the main text's,
  // and the headers', whose places count from their story), its box, wrap
  // and stacking, and the Office Art shape it is — read once, drawn where
  // its anchor character is met.
  const art = readOfficeArt(table, fib.pair(FC.DggInfo), wd);
  const floats = new Map();
  for (const [pair, from] of [[fib.pair(FC.PlcSpaMom), base.main], [fib.pair(FC.PlcSpaHdr), base.header]]) {
    const plc = readPlc(table, pair, 26);
    plc.entries.forEach((e, i) => {
      floats.set(from + plc.cps[i], { spid: u32(table, e), left: i32(table, e + 4), top: i32(table, e + 8), right: i32(table, e + 12), bottom: i32(table, e + 16), flags: u16(table, e + 20) });
    });
  }
  // The text boxes' stories: each box's paragraphs, found by the shape's id.
  const boxStories = [];
  for (const [pair, from] of [[fib.pair(FC.PlcftxbxTxt), base.textbox], [fib.pair(FC.PlcfHdrtxbxTxt), base.headerTextbox]]) {
    const plc = readPlc(table, pair, 22);
    plc.entries.forEach((e, i) => boxStories.push({ lid: i32(table, e + 14), index: i, from: from + plc.cps[i], to: from + plc.cps[i + 1], header: from !== base.textbox }));
  }
  const boxBlocks = (shape) => {
    const txid = shape.props.get(0x0080)?.op;
    if (txid == null) return null;
    const story = boxStories.find((s) => s.lid === shape.spid) || boxStories.filter((s) => !s.header)[(txid >>> 16) - 1];
    if (!story || story.to <= story.from) return null;
    const paras = readParagraphs(story.from, story.to);
    if (paras.length > 1 && !paras[paras.length - 1].runs.length) paras.pop();
    return groupTables(paras);
  };
  const floatOf = (spa) => {
    const shape = art.shapes.get(spa.spid);
    if (!shape) return null;
    const px = (tw) => tw / 15;
    const box = { x: px(spa.left), y: px(spa.top), w: Math.max(1, px(spa.right - spa.left)), h: Math.max(1, px(spa.bottom - spa.top)) };
    const wr = (spa.flags >> 5) & 0x0f;
    const float = {
      relH: ['margin', 'page', 'column'][(spa.flags >> 1) & 3] || 'column',
      relV: ['margin', 'page', 'paragraph'][(spa.flags >> 3) & 3] || 'paragraph',
      wrap: { 0: 'square', 1: 'topAndBottom', 2: 'square', 3: 'none', 4: 'tight', 5: 'through' }[wr] || 'square',
      side: ['bothSides', 'left', 'right', 'largest'][(spa.flags >> 9) & 0x0f] || 'bothSides',
      behind: Boolean(spa.flags & 0x4000),
      items: [],
    };
    // A group's members, each carried from the group's own coordinates into the anchor's box.
    const place = (s, at) => {
      if (s.children?.length && s.group) {
        const g = s.group;
        const gw = g.right - g.left || 1;
        const gh = g.bottom - g.top || 1;
        for (const c of s.children) {
          const a = c.anchor;
          if (!a) continue;
          place(c, { x: at.x + ((a.left - g.left) * at.w) / gw, y: at.y + ((a.top - g.top) * at.h) / gh, w: ((a.right - a.left) * at.w) / gw, h: ((a.bottom - a.top) * at.h) / gh });
        }
        return;
      }
      const item = drawingItem(s, at);
      if (item) float.items.push(item);
    };
    place(shape, box);
    return float.items.length ? float : null;
  };
  // One shape as what a .docx draws: a picture, a text box, a preset or freeform, a line.
  const drawingItem = (s, at) => {
    const p = s.props;
    const name = p.get(0x0380)?.complex ? utf16z(p.get(0x0380).complex) : null;
    const pib = p.get(0x0104)?.op;
    if (pib && art.blips[pib - 1]) {
      const blip = art.blips[pib - 1];
      const key = 'art:' + pib;
      if (!imageAt.has(key)) imageAt.set(key, images.push({ ...blip, widthTwips: at.w * 15, heightTwips: at.h * 15 }) - 1);
      return { kind: 'picture', ...at, name, image: imageAt.get(key) };
    }
    const bools = (id) => p.get(id)?.op;
    const filled = bools(0x01bf) != null && bools(0x01bf) & 0x100000 ? Boolean(bools(0x01bf) & 0x10) : true;
    const lined = bools(0x01ff) != null && bools(0x01ff) & 0x80000 ? Boolean(bools(0x01ff) & 0x08) : true;
    const rgb = (id, fallback) => (p.get(id) ? artColour(p.get(id).op, []) || fallback : fallback);
    const look = {
      fill: filled ? rgb(0x0181, 'FFFFFF') : null,
      line: lined ? rgb(0x01c0, '000000') : null,
      lineWidthPx: Math.max(0.5, (p.get(0x01cb)?.op ?? 9525) / 9525),
    };
    const blocks = boxBlocks(s);
    if (blocks) return { kind: 'textbox', ...at, name, blocks, ...look };
    if (LINES.has(s.type)) return { kind: 'shape', ...at, name, preset: 'line', fill: null, line: look.line || '000000', lineWidthPx: look.lineWidthPx, flipH: Boolean(s.flags & 0x40), flipV: Boolean(s.flags & 0x80) };
    const path = p.has(0x0145) ? freeformPath(p, { w: at.w, h: at.h }) : null;
    if (path) return { kind: 'shape', ...at, name, path, ...look, fill: path.filled ? look.fill : null };
    const preset = PRESETS[s.type];
    return preset ? { kind: 'shape', ...at, name, preset, ...look } : null;
  };

  // Notes: the references in the body, and each note's own paragraphs.
  for (const [kind, refIdx, txtIdx] of [['footnote', FC.PlcffndRef, FC.PlcffndTxt], ['endnote', FC.PlcfendRef, FC.PlcfendTxt]]) {
    const refs = readPlc(table, fib.pair(refIdx), 2);
    const txts = readPlc(table, fib.pair(txtIdx), 0);
    refs.cps.slice(0, refs.entries.length).forEach((cp, i) => {
      const start = txts.cps[i];
      const end = txts.cps[i + 1];
      if (start == null || end == null) return;
      notes[kind].push({ cp, id: i + 1, paragraphs: readParagraphs(base[kind] + start, base[kind] + end) });
    });
  }
  const noteAt = new Map();
  for (const kind of ['footnote', 'endnote']) for (const n of notes[kind]) noteAt.set(n.cp, { kind, id: n.id });

  // The sections' pages, and their headers and footers.
  const hdd = readPlc(table, fib.pair(FC.PlcfHdd), 0);
  const story = (i) => {
    const a = hdd.cps[i];
    const b = hdd.cps[i + 1];
    if (a == null || b == null || b <= a) return null;
    const paras = readParagraphs(base.header + a, base.header + b);
    // Each story ends with a paragraph mark of its own that the page does not show.
    if (paras.length > 1 && !paras[paras.length - 1].runs.length) paras.pop();
    return paras;
  };
  const sections = [];
  const last = { default: null, even: null, first: null };
  const lastFoot = { default: null, even: null, first: null };
  sed.entries.forEach((e, s) => {
    const fcSepx = i32(table, e + 2);
    const valid = fcSepx >= 0 && fcSepx < wd.length;
    const grpprl = readSepx ? (valid ? readSepx(fcSepx) : new Uint8Array(0)) : valid ? wd.subarray(fcSepx + 2, fcSepx + 2 + i16(wd, fcSepx)) : new Uint8Array(0);
    const sep = readSep(grpprl);
    // The section's stories, in the order even header, header, even footer,
    // footer, first-page header, first-page footer: six to a section from
    // Word 97 on, only those the section has before it.
    const k = headerSlots ? headerSlots(s, sep) : [0, 1, 2, 3, 4, 5].map((n) => 6 + s * 6 + n);
    // An empty story carries the previous section's on.
    const pick = (i, into, which) => { const st = i == null ? null : story(i); if (st) into[which] = st; return into[which]; };
    sections.push({
      cpEnd: sed.cps[s + 1],
      sep,
      headers: { even: pick(k[0], last, 'even'), default: pick(k[1], last, 'default'), first: pick(k[4], last, 'first') },
      footers: { even: pick(k[2], lastFoot, 'even'), default: pick(k[3], lastFoot, 'default'), first: pick(k[5], lastFoot, 'first') },
    });
  });
  if (!sections.length) sections.push({ cpEnd: fib.ccpText, sep: readSep(new Uint8Array(0)), headers: {}, footers: {} });

  // The body.
  const bodyParas = readParagraphs(0, fib.ccpText);
  // A note's reference: the auto-numbered mark at its CP in the body.
  for (const para of bodyParas) {
    for (const run of para.runs) {
      if (run.kind === 'noteMark' && noteAt.has(run.cp)) Object.assign(run, { kind: 'noteRef', note: noteAt.get(run.cp) });
    }
  }
  // Which section each paragraph closes: the one whose last character it ends at.
  let sIdx = 0;
  for (const para of bodyParas) {
    while (sIdx < sections.length - 1 && para.cpEnd >= sections[sIdx].cpEnd) { para.sectionEnd = sIdx; sIdx += 1; }
  }

  // The document's own settings: whether odd and even pages have headers of their own, and its layout rules.
  const dop = fib.pair(FC.Dop);
  return {
    format,
    nFib: fib.nFib,
    defaultFonts,
    evenAndOdd: dop.lcb ? Boolean(u8(table, dop.fc) & 1) : false,
    compat: dop.lcb ? compatOf(table, dop) : [],
    styles: styles.filter(Boolean),
    fonts,
    lists,
    sections,
    body: groupTables(bodyParas),
    notes: {
      footnote: notes.footnote.map((n) => ({ id: n.id, blocks: groupTables(n.paragraphs) })),
      endnote: notes.endnote.map((n) => ({ id: n.id, blocks: groupTables(n.paragraphs) })),
    },
    headers: sections.map((s) => ({
      headers: Object.fromEntries(Object.entries(s.headers).filter(([, v]) => v).map(([k, v]) => [k, groupTables(v)])),
      footers: Object.fromEntries(Object.entries(s.footers).filter(([, v]) => v).map(([k, v]) => [k, groupTables(v)])),
    })),
    images,
  };
}

/** A CHP as a key, so runs that look the same are one run. */
function chpKey(chp) {
  const { istd, ...rest } = chp;
  return JSON.stringify(rest) + '|' + (istd ?? '');
}

/**
 * Paragraphs into blocks: a run of in-table paragraphs becomes a table —
 * cells closed by their cell marks, rows by the row-end paragraph that
 * carries the row's properties. A nested table is laid out as its own
 * paragraphs inside the cell.
 */
export function groupTables(paras) {
  const blocks = [];
  let i = 0;
  while (i < paras.length) {
    const p = paras[i];
    if (!p.pap.inTable || (p.pap.itap ?? 1) < 1) {
      blocks.push({ type: 'paragraph', ...p });
      i += 1;
      continue;
    }
    const rows = [];
    let cells = [];
    let cell = [];
    while (i < paras.length && paras[i].pap.inTable) {
      const q = paras[i];
      // A nested table's own row end: its cells are already paragraphs of the outer cell.
      if (q.pap.innerTtp) { i += 1; continue; }
      if (q.pap.ttp && q.mark === 'cell') {
        rows.push({ cells, tap: q.pap.tap || {} });
        cells = [];
        cell = [];
      } else {
        cell.push({ type: 'paragraph', ...q });
        if (q.mark === 'cell' && (q.pap.itap ?? 1) <= 1) {
          cells.push(cell);
          cell = [];
        }
      }
      i += 1;
      if (q.sectionEnd != null) { rows.sectionEnd = q.sectionEnd; break; }
    }
    if (cell.length) cells.push(cell);
    if (cells.length) rows.push({ cells, tap: {} });
    const table = { type: 'table', rows };
    if (rows.sectionEnd != null) table.sectionEnd = rows.sectionEnd;
    blocks.push(table);
  }
  return blocks;
}

export { u8, u16, i16, u32, i32, FC, readPlc };
