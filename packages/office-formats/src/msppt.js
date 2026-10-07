// PowerPoint 97-2003 presentations (.ppt, .pps, .pot) — [MS-PPT].
//
// A .ppt is a compound file whose "PowerPoint Document" stream is a tree of
// records, written to in place by every fast save: the "Current User"
// stream says where the last edit's UserEditAtom is, and each edit's
// persist directory says where each object — the document, each slide,
// master and notes page — now lives, the newest edit's word winning. The
// document holds the slide size, the fonts, the pictures' index and the
// lists of slides, masters and notes; each slide holds its drawing (Office
// Art shapes, the same records Word and Excel draw with), its colour
// scheme and which master it follows. A placeholder's words live in the
// list of slides beside the slide, found by number; other shapes keep
// theirs in the shape.
//
// Read here into a plain model of slides — each its background and its
// shapes (text boxes with their paragraphs and runs, shapes with their
// fill, outline and words, pictures, lines) in slide pixels, and its
// speaker notes — for the deck builder to write as a .pptx. Text takes its
// look from the run first, then the master's style for its kind and level.
//
// The layouts follow [MS-PPT] and LibreOffice's PowerPoint import
// (filter/source/msfilter/svdfppt.cxx). Pure: bytes in, a model out.

import { CompoundFile } from './cfb.js';
import { findBlip } from './msdoc.js';

export class PptError extends Error {
  constructor(message) {
    super(message);
    this.name = 'PptError';
  }
}

const u8 = (b, at) => b[at] ?? 0;
const u16 = (b, at) => (b[at] ?? 0) | ((b[at + 1] ?? 0) << 8);
const i16 = (b, at) => { const v = u16(b, at); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (b, at) => (u16(b, at) | (u16(b, at + 2) << 16)) >>> 0;
const i32 = (b, at) => u32(b, at) | 0;

// Record types this reader follows.
const T = {
  Document: 0x03e8, DocumentAtom: 0x03e9, Slide: 0x03ee, SlideAtom: 0x03ef, Notes: 0x03f0, SlidePersistAtom: 0x03f3,
  Environment: 0x03f2, MainMaster: 0x03f8, Drawing: 0x040c, DrawingGroup: 0x040b, ColorScheme: 0x07f0, FontCollection: 0x07d5,
  FontEntity: 0x0fb7, SlideListWithText: 0x0ff0, UserEdit: 0x0ff5, PersistDirectory: 0x1772, OutlineTextRef: 0x0f9e,
  TextHeader: 0x0f9f, TextChars: 0x0fa0, StyleTextProp: 0x0fa1, TextBytes: 0x0fa8, TxMasterStyle: 0x0fa3, Placeholder: 0x0bc3,
  DggContainer: 0xf000, BStore: 0xf001, DgContainer: 0xf002, SpgrContainer: 0xf003, SpContainer: 0xf004, FBSE: 0xf007,
  FSPGR: 0xf009, FSP: 0xf00a, FOPT: 0xf00b, ClientTextbox: 0xf00d, ChildAnchor: 0xf00f, ClientAnchor: 0xf010, ClientData: 0xf011,
};

/** A record's header: version, instance, type, and where its body is. */
function header(b, at) {
  const vi = u16(b, at);
  const len = u32(b, at + 4);
  return { ver: vi & 0xf, inst: vi >> 4, type: u16(b, at + 2), len, body: at + 8, end: Math.min(b.length, at + 8 + len) };
}
/** A container's children. */
function children(b, h) {
  const out = [];
  if (!h) return out;
  let p = h.body;
  while (p + 8 <= h.end) {
    const c = header(b, p);
    out.push(c);
    if (c.len === 0 && c.ver !== 0xf) { p = c.body; continue; }
    p = c.body + c.len;
  }
  return out;
}
const child = (b, h, type, inst = null) => children(b, h).find((c) => c.type === type && (inst == null || c.inst === inst)) || null;

/* ── the document ─────────────────────────────────────────────────────── */

/**
 * Read a PowerPoint 97-2003 presentation: { size: { width, height },
 * slides: [{ background, shapes, notes }], images: [{ contentType, bytes }] }.
 */
export function readPpt(bytes) {
  const cfb = new CompoundFile(bytes);
  const app = cfb.application();
  if (app === 'encrypted') throw locked();
  const docEntry = cfb.find(['PowerPoint Document']);
  if (!docEntry) throw new PptError('not a PowerPoint 97-2003 presentation (no PowerPoint Document stream)');
  const doc = cfb.read(docEntry);
  const cuEntry = cfb.find(['Current User']);
  const cu = cuEntry ? cfb.read(cuEntry) : null;
  const picEntry = cfb.find(['Pictures']);
  const pictures = picEntry ? cfb.read(picEntry) : new Uint8Array(0);

  // The last edit, and the persist directory built back through every edit before it.
  let editAt = cu && u32(cu, 12) === 0xe391c05f ? u32(cu, 16) : null;
  if (cu && u32(cu, 12) === 0xf3d1c4df) throw locked();
  if (editAt == null || header(doc, editAt).type !== T.UserEdit) editAt = lastUserEdit(doc);
  if (editAt == null) throw new PptError('not a PowerPoint 97-2003 presentation (no edit record found) — a file from PowerPoint 95 or earlier');
  const persist = new Map();
  let docRef = null;
  for (let guard = 0; editAt != null && guard < 512; guard++) {
    const h = header(doc, editAt);
    if (h.type !== T.UserEdit) break;
    if (docRef == null) docRef = u32(doc, h.body + 16);
    if (h.len >= 32 && u32(doc, h.body + 28)) throw locked(); // an encryption session: the streams are encrypted
    const dirAt = u32(doc, h.body + 12);
    const dir = header(doc, dirAt);
    if (dir.type === T.PersistDirectory) {
      let p = dir.body;
      while (p + 4 <= dir.end) {
        const word = u32(doc, p);
        const first = word & 0xfffff;
        const count = word >>> 20;
        for (let i = 0; i < count; i++) if (!persist.has(first + i)) persist.set(first + i, u32(doc, p + 4 + i * 4));
        p += 4 + count * 4;
      }
    }
    const prev = u32(doc, h.body + 8);
    editAt = prev && prev !== editAt ? prev : null;
  }
  const at = (ref) => (persist.has(ref) ? header(doc, persist.get(ref)) : null);
  const docH = at(docRef);
  if (!docH || docH.type !== T.Document) throw new PptError('the presentation is damaged (its document record is missing)');

  // The slide size, in master units (576 to the inch) — slide pixels are a sixth of one.
  const docAtom = child(doc, docH, T.DocumentAtom);
  const size = { width: Math.round((docAtom ? i32(doc, docAtom.body) : 5760) / 6), height: Math.round((docAtom ? i32(doc, docAtom.body + 4) : 4320) / 6) };

  // Fonts, and the styles text takes when it says nothing.
  const env = child(doc, docH, T.Environment);
  const fonts = children(doc, child(doc, env, T.FontCollection)).filter((c) => c.type === T.FontEntity).map((c) => {
    let s = '';
    for (let i = 0; i < 32; i++) { const ch = u16(doc, c.body + i * 2); if (!ch) break; s += String.fromCharCode(ch); }
    return s;
  });
  const envStyles = readMasterStyles(doc, env);

  // The pictures' index: each FBSE says where its picture is in the Pictures stream.
  const store = children(doc, child(doc, child(doc, child(doc, docH, T.DrawingGroup), T.DggContainer), T.BStore));
  const blips = store.filter((c) => c.type === T.FBSE).map((c) => {
    const foDelay = u32(doc, c.body + 28);
    const cbName = u8(doc, c.body + 33);
    if (foDelay !== 0xffffffff && foDelay < pictures.length) return findBlip(pictures, foDelay, Math.min(pictures.length, foDelay + 8 + u32(pictures, foDelay + 4)));
    return findBlip(doc, c.body + 36 + cbName, c.end); // a picture kept in the record itself
  });
  const images = [];
  const imageIndex = new Map();
  const imageOf = (pib) => {
    if (!pib || !blips[pib - 1]) return null;
    if (!imageIndex.has(pib)) imageIndex.set(pib, images.push(blips[pib - 1]) - 1);
    return imageIndex.get(pib);
  };

  // The lists: slides (0), masters (1), notes (2) — each entry the object's persist number, its id, and its placeholders' words.
  const lists = {};
  for (const inst of [0, 1, 2]) {
    const entries = [];
    for (const c of children(doc, child(doc, docH, T.SlideListWithText, inst))) {
      if (c.type === T.SlidePersistAtom) entries.push({ ref: u32(doc, c.body), id: u32(doc, c.body + 12), texts: [] });
      else if (c.type === T.TextHeader && entries.length) entries[entries.length - 1].texts.push({ type: u32(doc, c.body), records: [] });
      const last = entries[entries.length - 1];
      const text = last?.texts[last.texts.length - 1];
      if (text && c.type !== T.TextHeader && c.type !== T.SlidePersistAtom) text.records.push(c);
    }
    lists[inst] = entries;
  }

  // The masters: each its colour scheme, text styles, background and the objects it puts on its slides.
  const masters = new Map();
  for (const m of lists[1]) {
    const h = at(m.ref);
    if (!h || (h.type !== T.MainMaster && h.type !== T.Slide)) continue;
    const scheme = readScheme(doc, h) || DEFAULT_SCHEME;
    const styles = mergeStyles(envStyles, readMasterStyles(doc, h));
    masters.set(m.id, { h, scheme, styles, entry: m });
  }
  const firstMaster = masters.values().next().value || { scheme: DEFAULT_SCHEME, styles: envStyles };

  const ctx = { doc, fonts, imageOf };
  const slides = [];
  for (const s of lists[0]) {
    const h = at(s.ref);
    if (!h || h.type !== T.Slide) continue;
    const atom = child(doc, h, T.SlideAtom);
    const flags = atom ? u16(doc, atom.body + 20) : 0x07;
    const master = (atom && masters.get(u32(doc, atom.body + 12))) || firstMaster;
    const scheme = flags & 0x02 || !readScheme(doc, h) ? master.scheme : readScheme(doc, h);
    const look = { ...ctx, scheme, styles: master.styles, slideNumber: slides.length + 1 };
    const own = readDrawing(doc, child(doc, h, T.Drawing), { ...look, texts: s.texts });
    const behind = master.h && flags & 0x01 ? readDrawing(doc, child(doc, master.h, T.Drawing), { ...look, texts: master.entry?.texts || [], masterObjects: true }) : { shapes: [], background: null };
    const background = flags & 0x04 || !own.background ? behind.background ?? readDrawing(doc, child(doc, master.h, T.Drawing), { ...look, texts: [] }).background : own.background;
    // The notes page's words, by its id.
    const notesId = atom ? u32(doc, atom.body + 16) : 0;
    const notesEntry = notesId ? lists[2].find((n) => n.id === notesId) : null;
    const notesH = notesEntry ? at(notesEntry.ref) : null;
    const notes = notesH && notesH.type === T.Notes ? readNotes(doc, notesH, { ...look, texts: notesEntry.texts }) : '';
    slides.push({ background, shapes: [...behind.shapes, ...own.shapes], notes });
  }
  return { size, slides, images, fonts };
}

function locked() {
  const err = new PptError('This presentation is password-protected.');
  err.encrypted = true;
  return err;
}

/** The last UserEditAtom in the stream, when the Current User stream cannot say. */
function lastUserEdit(doc) {
  let found = null;
  let p = 0;
  while (p + 8 <= doc.length) {
    const h = header(doc, p);
    if (h.type === T.UserEdit) found = p;
    if (h.len > doc.length) break;
    p = h.body + h.len;
  }
  return found;
}

/* ── colours and styles ───────────────────────────────────────────────── */

const DEFAULT_SCHEME = ['FFFFFF', '000000', '808080', '000000', '00CC99', '3333CC', 'CCCCFF', 'B2B2B2'];

/** A colour scheme: background, text and lines, shadows, title text, fills, accent, accent and link, accent and followed link. */
function readScheme(doc, h) {
  const c = child(doc, h, T.ColorScheme);
  if (!c || c.len < 32) return null;
  const out = [];
  for (let i = 0; i < 8; i++) out.push([0, 1, 2].map((k) => u8(doc, c.body + i * 4 + k).toString(16).padStart(2, '0')).join('').toUpperCase());
  return out;
}

/**
 * A paragraph's exception — its own bullet, alignment, spacing and
 * indents — in StyleTextPropAtom's order of fields, or (`master`) a
 * master style level's, which moves some of them.
 */
function readPf(doc, p, master = false, first = false) {
  const m = u32(doc, p);
  p += 4;
  const pf = {};
  const take2 = () => { const v = i16(doc, p); p += 2; return v; };
  if (m & 0x0f) { const f = u16(doc, p); p += 2; if (m & 1) pf.bulletOn = Boolean(f & 1); if (m & 2) pf.bulletHardFont = Boolean(f & 2); if (m & 4) pf.bulletHardColor = Boolean(f & 4); }
  if (m & 0x0080) pf.bulletChar = u16(doc, (p += 2) - 2);
  if (m & 0x0010) pf.bulletFont = u16(doc, (p += 2) - 2);
  if (m & 0x0040) pf.bulletSize = take2();
  if (m & 0x0020) { pf.bulletColor = u32(doc, p); p += 4; }
  if (!master) {
    if (m & 0x0800) pf.align = take2() & 3;
    if (m & 0x1000) pf.lineSpacing = take2();
    if (m & 0x2000) pf.spaceBefore = take2();
    if (m & 0x4000) pf.spaceAfter = take2();
    if (m & 0x0100) pf.leftMargin = take2();
    if (m & 0x0400) pf.indent = take2();
    if (m & 0x8000) take2();
    if (m & 0x100000) { const n = u16(doc, p); p += 2 + n * 4; }
    if (m & 0x10000) take2();
    if (m & 0xe0000) take2();
    if (m & 0x200000) take2();
  } else if (first) {
    if (m & 0xf00) pf.align = take2() & 3;
    if (m & 0x1000) pf.lineSpacing = take2();
    if (m & 0x2000) pf.spaceBefore = take2();
    if (m & 0x4000) pf.spaceAfter = take2();
    if (m & 0x8000) pf.leftMargin = take2();
    if (m & 0x10000) pf.indent = take2();
    if (m & 0x20000) take2();
    if (m & 0x200000) { const n = u16(doc, p); p += 2 + n * 4; }
    if (m & 0x40000) take2();
    if (m & 0x80000) take2();
    if (m & 0x100000) take2();
    for (let bits = m >>> 22; bits; bits >>>= 1) if (bits & 1) take2();
  } else {
    if (m & 0x0800) pf.align = take2() & 3;
    if (m & 0x1000) pf.lineSpacing = take2();
    if (m & 0x2000) pf.spaceBefore = take2();
    if (m & 0x4000) pf.spaceAfter = take2();
    if (m & 0x8000) take2();
    if (m & 0x0100) pf.leftMargin = take2();
    if (m & 0x0200) take2();
    if (m & 0x0400) pf.indent = take2();
    if (m & 0x10000) take2();
    if (m & 0xe0000) take2();
    if (m & 0x100000) { const n = u16(doc, p); p += 2 + n * 4; }
    if (m & 0x200000) take2();
    for (let bits = m >>> 22; bits; bits >>>= 1) if (bits & 1) take2();
  }
  return { pf, end: p };
}

/** A run's exception: bold, italic, underline, font, size, colour, raise. */
function readCf(doc, p) {
  const m = u32(doc, p);
  p += 4;
  const cf = {};
  if (m & 0xffff) {
    const f = u16(doc, p);
    p += 2;
    if (m & 1) cf.bold = Boolean(f & 1);
    if (m & 2) cf.italic = Boolean(f & 2);
    if (m & 4) cf.underline = Boolean(f & 4);
  }
  if (m & 0x10000) { cf.font = u16(doc, p); p += 2; }
  if (m & 0x200000) p += 2;
  if (m & 0x400000) p += 2;
  if (m & 0x800000) p += 2;
  if (m & 0x20000) { cf.size = u16(doc, p); p += 2; }
  if (m & 0x40000) { cf.color = u32(doc, p); p += 4; }
  if (m & 0x80000) { cf.position = i16(doc, p); p += 2; }
  if (m & 0x100000) p += 2;
  for (let bits = m >>> 24; bits; bits >>>= 1) if (bits & 1) p += 2;
  return { cf, end: p };
}

/** The master's text styles: for each kind of text (title, body, notes, other, subtitle…), its levels' paragraph and run looks. */
function readMasterStyles(doc, h) {
  const out = {};
  for (const c of children(doc, h).filter((x) => x.type === T.TxMasterStyle)) {
    const levels = [];
    const n = Math.min(5, u16(doc, c.body));
    let p = c.body + 2;
    for (let lv = 0; lv < n && p < c.end; lv++) {
      let first = lv === 0;
      if (c.inst >= 5) { p += 2; first = false; }
      const pf = readPf(doc, p, true, first);
      const cf = readCf(doc, pf.end);
      p = cf.end;
      levels.push({ pf: { ...(levels[lv - 1]?.pf || {}), ...pf.pf }, cf: { ...(levels[lv - 1]?.cf || {}), ...cf.cf } });
    }
    out[c.inst] = levels;
  }
  return out;
}

function mergeStyles(base, over) {
  const out = { ...base };
  for (const [k, v] of Object.entries(over)) out[k] = v;
  return out;
}

/** A text type's master style: subtitles and half/quarter bodies are the body's, centred titles the title's, where not their own. */
function styleFor(styles, textType) {
  const order = { 0: [0], 1: [1], 2: [2, 1], 4: [4, 1], 5: [5, 1], 6: [6, 0], 7: [7, 1], 8: [8, 1] }[textType] || [4, 1];
  for (const k of order) if (styles[k]?.length) return styles[k];
  return [];
}

/** A colour in a run, a bullet or a shape: a scheme entry by number, or its own RGB. */
function colourOf(value, scheme) {
  const index = value >>> 24;
  if (index === 0xfe) return [0, 8, 16].map((s) => ((value >>> s) & 0xff).toString(16).padStart(2, '0')).join('').toUpperCase();
  if (index < 8) return scheme[index] ?? null;
  return null;
}
/** An Office Art colour: its RGB, or a scheme entry when it says so. */
function artColour(value, scheme) {
  const flags = value >>> 24;
  if (flags & 0x08) return scheme[value & 0xff] ?? null;
  if (flags & 0x10) return null; // a system colour this has no table for
  return [0, 8, 16].map((s) => ((value >>> s) & 0xff).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/* ── shapes ───────────────────────────────────────────────────────────── */

/** An OfficeArtFOPT's properties, by id: { op, complex? }. */
function readFopt(doc, h) {
  const props = new Map();
  if (!h) return props;
  const n = h.inst;
  let complexAt = h.body + n * 6;
  for (let k = 0; k < n; k++) {
    const id = u16(doc, h.body + k * 6);
    const op = u32(doc, h.body + k * 6 + 2);
    const entry = { op };
    if (id & 0x8000) { entry.complex = doc.subarray(complexAt, Math.min(h.end, complexAt + op)); complexAt += op; }
    props.set(id & 0x3fff, entry);
  }
  return props;
}

// Office Art's shape types, by number, as DrawingML's preset geometries (LibreOffice's
// table of them, filter/source/msfilter/util.cxx). Lines and connectors, pictures,
// text boxes and WordArt are read as what they are, not as presets.
const PRESETS = {
  1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'triangle', 6: 'rtTriangle', 7: 'parallelogram', 8: 'trapezoid', 9: 'hexagon',
  10: 'octagon', 11: 'plus', 12: 'star5', 13: 'rightArrow', 14: 'rightArrow', 15: 'homePlate', 16: 'cube', 17: 'wedgeRoundRectCallout', 18: 'star16',
  19: 'arc', 21: 'plaque', 22: 'can', 23: 'donut', 41: 'callout1', 42: 'callout2', 43: 'callout3', 44: 'accentCallout1', 45: 'accentCallout2',
  46: 'accentCallout3', 47: 'borderCallout1', 48: 'borderCallout2', 49: 'borderCallout3', 50: 'accentBorderCallout1', 51: 'accentBorderCallout2',
  52: 'accentBorderCallout3', 53: 'ribbon', 54: 'ribbon2', 55: 'chevron', 56: 'pentagon', 57: 'noSmoking', 58: 'star8', 59: 'star16', 60: 'star32',
  61: 'wedgeRectCallout', 62: 'wedgeRoundRectCallout', 63: 'wedgeEllipseCallout', 64: 'wave', 65: 'foldedCorner', 66: 'leftArrow', 67: 'downArrow',
  68: 'upArrow', 69: 'leftRightArrow', 70: 'upDownArrow', 71: 'irregularSeal1', 72: 'irregularSeal2', 73: 'lightningBolt', 74: 'heart',
  76: 'quadArrow', 77: 'leftArrowCallout', 78: 'rightArrowCallout', 79: 'upArrowCallout', 80: 'downArrowCallout', 81: 'leftRightArrowCallout',
  82: 'upDownArrowCallout', 83: 'quadArrowCallout', 84: 'bevel', 85: 'leftBracket', 86: 'rightBracket', 87: 'leftBrace', 88: 'rightBrace',
  89: 'leftUpArrow', 90: 'bentUpArrow', 91: 'bentArrow', 92: 'star24', 93: 'stripedRightArrow', 94: 'notchedRightArrow', 95: 'blockArc',
  96: 'smileyFace', 97: 'verticalScroll', 98: 'horizontalScroll', 99: 'circularArrow', 100: 'notchedCircularArrow', 101: 'uturnArrow',
  102: 'curvedRightArrow', 103: 'curvedLeftArrow', 104: 'curvedUpArrow', 105: 'curvedDownArrow', 106: 'cloudCallout', 107: 'ellipseRibbon',
  108: 'ellipseRibbon2', 109: 'flowChartProcess', 110: 'flowChartDecision', 111: 'flowChartInputOutput', 112: 'flowChartPredefinedProcess',
  113: 'flowChartInternalStorage', 114: 'flowChartDocument', 115: 'flowChartMultidocument', 116: 'flowChartTerminator', 117: 'flowChartPreparation',
  118: 'flowChartManualInput', 119: 'flowChartManualOperation', 121: 'flowChartPunchedCard', 122: 'flowChartPunchedTape',
  123: 'flowChartSummingJunction', 124: 'flowChartOr', 125: 'flowChartCollate', 126: 'flowChartSort', 127: 'flowChartExtract', 128: 'flowChartMerge',
  129: 'flowChartOfflineStorage', 130: 'flowChartOnlineStorage', 131: 'flowChartMagneticTape', 132: 'flowChartMagneticDisk',
  133: 'flowChartMagneticDrum', 134: 'flowChartDisplay', 135: 'flowChartDelay', 176: 'flowChartAlternateProcess', 178: 'callout1',
  179: 'accentCallout1', 180: 'borderCallout1', 181: 'accentBorderCallout1', 182: 'leftRightUpArrow', 183: 'sun', 184: 'moon', 185: 'bracketPair',
  186: 'bracePair', 187: 'star4', 188: 'doubleWave', 189: 'actionButtonBlank', 190: 'actionButtonHome', 191: 'actionButtonHelp',
  192: 'actionButtonInformation', 193: 'actionButtonForwardNext', 194: 'actionButtonBackPrevious', 195: 'actionButtonEnd',
  196: 'actionButtonBeginning', 197: 'actionButtonReturn', 198: 'actionButtonDocument', 199: 'actionButtonSound', 200: 'actionButtonMovie',
  203: 'roundRect',
};
const LINES = new Set([20, 32, 33, 34, 35, 36, 37, 38, 39, 40]);

/**
 * A slide's (or master's) drawing: its shapes, in slide pixels, and its
 * background's colour. Groups are walked into, each child's box carried
 * from the group's own coordinates to the slide's.
 */
function readDrawing(doc, drawing, ctx) {
  const out = { shapes: [], background: null };
  const dg = child(doc, drawing, T.DgContainer);
  if (!dg) return out;
  const identity = (r) => r;
  for (const c of children(doc, dg)) {
    if (c.type === T.SpgrContainer) walkGroup(doc, c, identity, ctx, out.shapes, true);
    else if (c.type === T.SpContainer) {
      const props = readFopt(doc, child(doc, c, T.FOPT));
      const fsp = child(doc, c, T.FSP);
      if (fsp && u32(doc, fsp.body + 4) & 0x400) out.background = fillOf(props, ctx, true) || null;
    }
  }
  return out;
}

function walkGroup(doc, group, transform, ctx, shapes, top) {
  const kids = children(doc, group);
  if (!kids.length) return;
  // The group's own shape comes first: its box on the page and the coordinates its children use.
  const own = kids[0];
  let inner = transform;
  if (own.type === T.SpContainer && !top) {
    const spgr = child(doc, own, T.FSPGR);
    const box = anchorOf(doc, own, transform);
    if (spgr && box) {
      const cx = i32(doc, spgr.body);
      const cy = i32(doc, spgr.body + 4);
      const cw = i32(doc, spgr.body + 8) - cx || 1;
      const ch = i32(doc, spgr.body + 12) - cy || 1;
      inner = (r) => ({ x: box.x + ((r.x - cx) * box.w) / cw, y: box.y + ((r.y - cy) * box.h) / ch, w: (r.w * box.w) / cw, h: (r.h * box.h) / ch });
    }
  }
  for (const k of kids.slice(1)) {
    if (k.type === T.SpgrContainer) walkGroup(doc, k, inner, ctx, shapes, false);
    else if (k.type === T.SpContainer) {
      const shape = readShape(doc, k, inner, ctx);
      if (shape) shapes.push(shape);
    }
  }
}

/** A shape's box: a top-level shape's client anchor (master units), or a child's anchor in its group's coordinates. */
function anchorOf(doc, sp, transform) {
  const client = child(doc, sp, T.ClientAnchor);
  if (client) {
    let top;
    let left;
    let right;
    let bottom;
    if (client.len >= 16) { top = i32(doc, client.body); left = i32(doc, client.body + 4); right = i32(doc, client.body + 8); bottom = i32(doc, client.body + 12); }
    else { top = i16(doc, client.body); left = i16(doc, client.body + 2); right = i16(doc, client.body + 4); bottom = i16(doc, client.body + 6); }
    return { x: left / 6, y: top / 6, w: (right - left) / 6, h: (bottom - top) / 6 };
  }
  const ca = child(doc, sp, T.ChildAnchor);
  if (ca) {
    const left = i32(doc, ca.body);
    const top = i32(doc, ca.body + 4);
    return transform({ x: left, y: top, w: i32(doc, ca.body + 8) - left, h: i32(doc, ca.body + 12) - top });
  }
  return null;
}

/**
 * A shape's fill as a colour, or 'none'. What a shape does not say is
 * Office Art's own default — filled, white — not the drawing group's
 * defaults, which are only what PowerPoint gives a new shape.
 */
function fillOf(props, ctx, background = false) {
  const get = (id) => props.get(id);
  const bools = get(0x01bf)?.op;
  const filled = bools != null && bools & 0x100000 ? Boolean(bools & 0x10) : true;
  if (!filled) return 'none';
  const type = get(0x0180)?.op ?? 0;
  const colour = get(0x0181) ? artColour(get(0x0181).op, ctx.scheme) : background ? ctx.scheme[0] : 'FFFFFF';
  // A gradient or a texture is drawn as its first colour; a picture fill as nothing this can draw.
  if (type === 3) return background ? null : 'none';
  return colour ? '#' + colour : 'none';
}

/** A shape's outline, or 'none': black, three quarters of a point, where it does not say. */
function lineOf(props, ctx, defaultOn) {
  const get = (id) => props.get(id);
  const bools = get(0x01ff)?.op;
  const on = bools != null && bools & 0x80000 ? Boolean(bools & 0x08) : defaultOn;
  if (!on) return 'none';
  const colour = get(0x01c0) ? artColour(get(0x01c0).op, ctx.scheme) : '000000';
  return { color: '#' + (colour || '000000'), width: Math.max(0.25, (get(0x01cb)?.op ?? 9525) / 12700) };
}

/** One shape: a picture, a line, a shape with or without words, or a text box. */
function readShape(doc, sp, transform, ctx) {
  const fsp = child(doc, sp, T.FSP);
  if (!fsp) return null;
  const type = fsp.inst;
  const flags = u32(doc, fsp.body + 4);
  if (flags & 0x08 || flags & 0x400) return null; // deleted, or the background
  const box = anchorOf(doc, sp, transform);
  if (!box) return null;
  const props = readFopt(doc, child(doc, sp, T.FOPT));
  const data = child(doc, sp, T.ClientData);
  const ph = data ? child(doc, data, T.Placeholder) : null;
  const placement = ph ? u8(doc, ph.body + 4) : null;
  // On a master only its own objects come through: its placeholders are the slides' business.
  if (ctx.masterObjects && ph) return null;
  const name = props.get(0x0380)?.complex ? utf16(props.get(0x0380).complex) : null;
  const rotation = props.get(0x0004) ? i32(new Uint8Array(new Uint32Array([props.get(0x0004).op]).buffer), 0) / 65536 : 0;
  const base = { x: box.x, y: box.y, w: box.w, h: box.h, name, rotation, flipH: Boolean(flags & 0x40), flipV: Boolean(flags & 0x80) };

  const pib = props.get(0x0104)?.op;
  if (pib && (type === 75 || !child(doc, sp, T.ClientTextbox))) {
    const image = ctx.imageOf(pib);
    return image == null ? null : { ...base, type: 'picture', image };
  }
  if (LINES.has(type)) return { ...base, type: 'line', line: lineOf(props, ctx, true) };
  // WordArt: its words in the shape's own properties, at the size and in the font they name, in the shape's fill.
  const art = props.get(0x00c0)?.complex;
  if (art) {
    const words = utf16(art);
    if (!words.trim()) return null;
    const fill = fillOf(props, ctx);
    const run = { text: words.replace(/[\r\n]+/g, ' '), size: Math.max(8, Math.round((props.get(0x00c3)?.op ?? 36 * 65536) / 65536)) };
    if (props.get(0x00c5)?.complex) run.font = utf16(props.get(0x00c5).complex);
    if (fill !== 'none') run.color = fill;
    if ((props.get(0x00ff)?.op ?? 0) & 0x20) run.bold = true;
    return { ...base, type: 'text', paragraphs: [{ align: 'center', bullet: false, runs: [run] }], fill: 'none', line: 'none' };
  }

  const paragraphs = shapeText(doc, child(doc, sp, T.ClientTextbox), ctx, placement);
  const isText = type === 202 || ph;
  if (isText) {
    if (!paragraphs || !paragraphs.some((p) => p.runs.some((r) => r.text && r.text.trim()))) return null;
    const fill = props.has(0x0181) || props.get(0x01bf)?.op & 0x100000 ? fillOf(props, ctx) : 'none';
    const line = props.get(0x01ff)?.op & 0x80000 ? lineOf(props, ctx, false) : 'none';
    return { ...base, type: 'text', paragraphs, fill, line };
  }
  // A shape that carries its own points is drawn from them, whatever its
  // number says: a freeform, and every newer shape PowerPoint has no number
  // for, which it saves as points under a number of no meaning.
  const path = props.has(0x0145) ? freeformPath(props, box) : null;
  if (path) return { ...base, type: 'shape', preset: 'rect', path, fill: path.filled ? fillOf(props, ctx) : 'none', line: lineOf(props, ctx, true), paragraphs };
  const preset = PRESETS[type];
  if (!preset) return null;
  return { ...base, type: 'shape', preset, fill: fillOf(props, ctx), line: lineOf(props, ctx, true), paragraphs };
}

/** An Office Art array (IMsoArray): its elements' bytes, each `size` long. */
function msoArray(b) {
  if (!b || b.length < 6) return [];
  const n = u16(b, 0);
  let size = u16(b, 4);
  if (size === 0xfff0) size = 4;
  const out = [];
  for (let i = 0; i < n && 6 + (i + 1) * size <= b.length; i++) out.push(b.subarray(6 + i * size, 6 + (i + 1) * size));
  return out;
}

/**
 * A freeform's outline, as path commands in the shape's own pixels: its
 * points (pVertices) in the coordinates geoLeft..geoRight, geoTop..geoBottom
 * name, walked by its segments (pSegmentInfo) — move, line, curve, close —
 * or joined in order when it has none.
 */
function freeformPath(props, box) {
  const vertices = msoArray(props.get(0x0145)?.complex).map((e) => (e.length >= 8 ? [i32(e, 0), i32(e, 4)] : [i16(e, 0), i16(e, 2)]));
  if (vertices.length < 2) return null;
  const left = props.get(0x0140)?.op ?? 0;
  const top = props.get(0x0141)?.op ?? 0;
  const right = props.get(0x0142)?.op ?? 21600;
  const bottom = props.get(0x0143)?.op ?? 21600;
  const sx = box.w / ((right | 0) - (left | 0) || 1);
  const sy = box.h / ((bottom | 0) - (top | 0) || 1);
  const pt = (v) => [(v[0] - (left | 0)) * sx, (v[1] - (top | 0)) * sy];
  const commands = [];
  let k = 0;
  let closed = false;
  const segments = msoArray(props.get(0x0146)?.complex).map((e) => u16(e, 0));
  if (!segments.length) {
    commands.push({ op: 'M', pts: [pt(vertices[0])] });
    for (let i = 1; i < vertices.length; i++) commands.push({ op: 'L', pts: [pt(vertices[i])] });
  } else {
    for (const seg of segments) {
      const kind = seg >> 13;
      const count = seg & 0x1fff;
      if (kind === 2 && vertices[k]) commands.push({ op: 'M', pts: [pt(vertices[k++])] });
      else if (kind === 0) for (let i = 0; i < Math.max(1, count) && vertices[k]; i++) commands.push({ op: 'L', pts: [pt(vertices[k++])] });
      else if (kind === 1) for (let i = 0; i < Math.max(1, count) && vertices[k + 2]; i++) { commands.push({ op: 'C', pts: [pt(vertices[k]), pt(vertices[k + 1]), pt(vertices[k + 2])] }); k += 3; }
      else if (kind === 3) { commands.push({ op: 'Z' }); closed = true; }
      else if (kind === 4) break;
      else if (kind === 5) k += seg & 0xff; // an escape: its points skipped
    }
  }
  if (!commands.length || commands[0].op !== 'M') return null;
  return { commands, w: box.w, h: box.h, filled: closed };
}

const utf16 = (b) => { let s = ''; for (let i = 0; i + 1 < b.length; i += 2) { const c = u16(b, i); if (!c) break; s += String.fromCharCode(c); } return s; };

/** A placeholder's kind, as the kind of text it holds — what its master style is. */
const PLACEMENT_TEXT = { 1: 0, 2: 1, 3: 6, 4: 5, 13: 0, 14: 1, 15: 6, 16: 5, 17: 0, 18: 1 };

/** A shape's words: its own, or a placeholder's from the slide list by number. */
function shapeText(doc, box, ctx, placement) {
  if (!box) return null;
  const kids = children(doc, box);
  const ref = kids.find((k) => k.type === T.OutlineTextRef);
  if (ref) {
    const entry = ctx.texts?.[i32(doc, ref.body)];
    return entry ? paragraphsOf(doc, entry.records, entry.type, ctx) : null;
  }
  const head = kids.find((k) => k.type === T.TextHeader);
  const textType = head ? u32(doc, head.body) : placement != null ? PLACEMENT_TEXT[placement] ?? 4 : 4;
  return paragraphsOf(doc, kids, textType, ctx);
}

/**
 * Text and its style records as paragraphs of runs: the characters split
 * at each paragraph mark, each paragraph its exception over the master's
 * level style, each run its own over that.
 */
function paragraphsOf(doc, records, textType, ctx) {
  const chars = records.find((r) => r.type === T.TextChars);
  const bytes = records.find((r) => r.type === T.TextBytes);
  let text = '';
  if (chars) for (let p = chars.body; p + 1 < chars.end; p += 2) text += String.fromCharCode(u16(doc, p));
  else if (bytes) for (let p = bytes.body; p < bytes.end; p++) text += String.fromCharCode(doc[p]); // each byte a UTF-16 code's low half
  else return null;

  // The paragraph runs, then the character runs, each counted in characters.
  const pfRuns = [];
  const cfRuns = [];
  const style = records.find((r) => r.type === T.StyleTextProp);
  if (style) {
    let p = style.body;
    let covered = 0;
    while (covered <= text.length && p + 6 <= style.end) {
      const count = u32(doc, p);
      const level = u16(doc, p + 4);
      const { pf, end } = readPf(doc, p + 6);
      pfRuns.push({ count, level: Math.min(4, level), pf });
      covered += count;
      p = end;
      if (!count) break;
    }
    covered = 0;
    while (covered <= text.length && p + 8 <= style.end) {
      const count = u32(doc, p);
      const { cf, end } = readCf(doc, p + 4);
      cfRuns.push({ count, cf });
      covered += count;
      p = end;
      if (!count) break;
    }
  }
  const at = (runs, i) => {
    let sum = 0;
    for (const r of runs) { sum += r.count; if (i < sum) return r; }
    return runs[runs.length - 1] || null;
  };

  const master = styleFor(ctx.styles, textType);
  // A slide number field: a "*" in the text, shown as the slide's number.
  const numberAt = records.find((r) => r.type === 0x0fd8);
  const slidePos = numberAt && ctx.slideNumber != null ? i32(doc, numberAt.body) : -1;
  const paragraphs = [];
  let start = 0;
  const parts = text.split('\r');
  parts.forEach((part, pi) => {
    if (pi === parts.length - 1 && part === '' && parts.length > 1) return;
    const pr = at(pfRuns, start);
    const level = pr?.level ?? 0;
    const lv = master[level] || master[0] || { pf: {}, cf: {} };
    const pf = { ...lv.pf, ...(pr?.pf || {}) };
    const para = { level, runs: [] };
    if (pf.align != null) para.align = ['left', 'center', 'right', 'justify'][pf.align] || 'left';
    if (pf.bulletOn) {
      const ch = pf.bulletChar;
      para.bullet = { type: 'char', char: ch && ch >= 0x20 && !(ch >= 0xf000 && ch <= 0xf0ff) ? String.fromCharCode(ch) : '•' };
      if (pf.bulletHardColor && pf.bulletColor != null) { const c = colourOf(pf.bulletColor, ctx.scheme); if (c) para.bullet.color = '#' + c; }
    } else para.bullet = false;
    if (pf.leftMargin != null) para.indent = pf.leftMargin / 6;
    if (pf.indent != null && pf.leftMargin != null) para.hanging = (pf.indent - pf.leftMargin) / 6;
    if (pf.lineSpacing > 0 && pf.lineSpacing !== 100) para.lineHeight = pf.lineSpacing / 100;
    if (pf.spaceBefore < 0) para.spaceBefore = -pf.spaceBefore / 8;
    if (pf.spaceAfter < 0) para.spaceAfter = -pf.spaceAfter / 8;
    // The runs: where the character style changes, and at each line break.
    let run = null;
    let key = '';
    for (let k = 0; k < part.length; k++) {
      const ch = start + k === slidePos ? String(ctx.slideNumber) : part[k];
      const cr = at(cfRuns, start + k);
      const cf = { ...lv.cf, ...(cr?.cf || {}) };
      const look = {};
      if (cf.size) look.size = cf.size;
      if (cf.bold) look.bold = true;
      if (cf.italic) look.italic = true;
      if (cf.underline) look.underline = true;
      if (cf.color != null) { const c = colourOf(cf.color, ctx.scheme); if (c) look.color = '#' + c; }
      if (cf.font != null && ctx.fonts[cf.font]) look.font = ctx.fonts[cf.font];
      if (cf.position > 0) look.baseline = 'super';
      else if (cf.position < 0) look.baseline = 'sub';
      if (ch === '\u000b') { para.runs.push({ ...look, text: '\n' }); run = null; key = ''; continue; }
      const kk = JSON.stringify(look);
      if (!run || kk !== key) { run = { ...look, text: '' }; key = kk; para.runs.push(run); }
      run.text += ch.length === 1 && ch < ' ' && ch !== '\t' ? '' : ch;
    }
    if (!para.runs.length) {
      const cf = { ...lv.cf };
      para.runs.push({ text: '', ...(cf.size ? { size: cf.size } : {}) });
    }
    paragraphs.push(para);
    start += part.length + 1;
  });
  return paragraphs;
}

/** A notes page's words: its body placeholder's, as plain text. */
function readNotes(doc, h, ctx) {
  const dg = child(doc, child(doc, h, T.Drawing), T.DgContainer);
  if (!dg) return '';
  const found = [];
  const visit = (c) => {
    for (const k of children(doc, c)) {
      if (k.type === T.SpgrContainer) visit(k);
      else if (k.type === T.SpContainer) {
        const data = child(doc, k, T.ClientData);
        const ph = data ? child(doc, data, T.Placeholder) : null;
        const placement = ph ? u8(doc, ph.body + 4) : null;
        if (placement !== 0x0c && placement !== 0x06) continue;
        const paras = shapeText(doc, child(doc, k, T.ClientTextbox), ctx, 2);
        if (paras) found.push(paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n'));
      }
    }
  };
  visit(dg);
  return found.join('\n').trim();
}
