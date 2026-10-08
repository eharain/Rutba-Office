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
// fill, outline and words, pictures, lines, tables) in slide pixels, its
// speaker notes, and how it plays in the show (hidden or not, its
// transition, its shapes' effects) — for the deck builder to write as a
// .pptx. Text takes its look from the run first, then the master's style
// for its kind and level.
//
// The layouts follow [MS-PPT] and LibreOffice's PowerPoint import
// (filter/source/msfilter/svdfppt.cxx). Pure: bytes in, a model out.

import { CompoundFile } from './cfb.js';
import { findBlip } from './msdoc.js';
import { languageTag } from './lcid.js';
import { header, children, child, artColour, readFopt, PRESETS, LINES, freeformPath, artFill, artLine, artShadow } from './officeart.js';

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
  TertiaryFOPT: 0xf122, SlideShowSlideInfo: 0x03f9, AnimationInfo: 0x1014, AnimationInfoAtom: 0x0ff1,
  ProgTags: 0x1388, ProgBinaryTag: 0x138a, BinaryTagData: 0x138b, CString: 0x0fba, AnimGroup: 0xf144, AnimSubGroup: 0xf145,
  AnimPropertySet: 0xf13d, AnimAttributeValue: 0xf142, AnimReference: 0x2afb,
  RoundTripTheme: 0x040e, RoundTripColorMapping: 0x040f, RoundTripCompositeMasterId: 0x041d, RoundTripContentMasterInfo: 0x041e,
  RoundTripTextStyles: 0x0423, RoundTripTableStyles: 0x0428,
};


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

  // What PowerPoint 2007 and later keep beside all this: each master's theme, colour map and text
  // styles, each layout (a master of its own here) whole, and the table styles — the deck as they made it.
  const bytesOf = (h, type) => { const r = child(doc, h, type); return r && r.len ? doc.slice(r.body, r.end) : null; };
  const keptMasters = [...masters.entries()].map(([id, m]) => {
    const composite = child(doc, m.h, T.RoundTripCompositeMasterId);
    return {
      id, theme: bytesOf(m.h, T.RoundTripTheme), colourMap: bytesOf(m.h, T.RoundTripColorMapping), textStyles: bytesOf(m.h, T.RoundTripTextStyles),
      layout: bytesOf(m.h, T.RoundTripContentMasterInfo), composite: composite ? u32(doc, composite.body) : null,
    };
  });

  const ctx = { doc, fonts, imageOf };
  const slides = [];
  // A main master's own drawing, every object on it with what PowerPoint later kept of it.
  for (const k of keptMasters) {
    if (k.layout || !k.theme) continue;
    const m = masters.get(k.id);
    const drawn = readDrawing(doc, child(doc, m.h, T.Drawing), { ...ctx, scheme: m.scheme, styles: m.styles, texts: m.entry?.texts || [], allObjects: true });
    k.items = drawn.items;
    k.backgroundML = drawn.backgroundML;
  }
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
    slides.push({
      background, shapes: [...behind.shapes, ...own.shapes], notes, ...showOf(doc, child(doc, h, T.SlideShowSlideInfo)), timing: timingOf(doc, h),
      // For a deck rebuilt as PowerPoint 2007 kept it: its layout, its own drawings, its own background.
      kept: { layout: atom ? u32(doc, atom.body + 12) : null, items: own.items, backgroundML: flags & 0x04 ? null : own.backgroundML, masterObjects: Boolean(flags & 0x01), masterShapes: behind.shapes.length, ownBackground: !(flags & 0x04) },
    });
  }
  const keptTables = bytesOf(docH, T.RoundTripTableStyles) ?? bytesOf(child(doc, docH, T.Environment), T.RoundTripTableStyles);
  const kept = keptMasters.some((k) => k.theme && !k.layout) ? { masters: keptMasters, tableStyles: keptTables } : null;
  return { size, slides, images, fonts, kept };
}

/**
 * How a slide comes on in the show (its SlideShowSlideInfoAtom): whether it
 * is hidden, its transition — PowerPoint 97's effect and direction as the
 * classic effect a .pptx names, the eight that have none as the nearest —
 * at its speed, and whether a click or the time on it moves the show on.
 */
function showOf(doc, atom) {
  if (!atom) return { hidden: false, transition: null };
  const d = atom.body;
  const time = i32(doc, d);
  const dir = u8(doc, d + 8);
  const effect = u8(doc, d + 9);
  const flags = u16(doc, d + 10);
  const speed = u8(doc, d + 12);
  const SIDES = ['l', 'u', 'r', 'd'];
  const EIGHT = ['l', 'u', 'r', 'd', 'lu', 'ru', 'ld', 'rd'];
  const bars = (vertical) => ({ type: 'randomBar', direction: vertical ? 'vert' : 'horz' });
  const as = {
    // A cut, through black or not; any other direction is no effect at all (a hidden slide's, for one).
    0: dir === 1 ? { type: 'cut', direction: 'black' } : dir === 0 ? { type: 'cut', direction: 'smooth' } : null,
    1: { type: 'fade', direction: 'smooth' },
    2: bars(dir === 0), 3: bars(dir === 1), 4: { type: 'cover', direction: EIGHT[dir] ?? 'l' }, 5: { type: 'dissolve', direction: null },
    6: { type: 'fade', direction: 'black' }, 7: { type: 'pull', direction: EIGHT[dir] ?? 'l' }, 8: bars(dir === 1),
    9: { type: 'cover', direction: EIGHT[dir] ?? 'lu' }, 10: { type: 'wipe', direction: SIDES[dir] ?? 'l' },
    11: { type: 'zoom', direction: dir === 1 ? 'in' : 'out' }, 13: { type: 'split', direction: ['horz-out', 'horz-in', 'vert-out', 'vert-in'][dir] ?? 'horz-out' },
    17: { type: 'diamond', direction: null }, 18: { type: 'plus', direction: null }, 19: { type: 'circle', direction: null },
    20: { type: 'push', direction: SIDES[dir] ?? 'l' }, 21: bars(dir === 1), 22: { type: 'zoom', direction: 'in' },
    23: { type: 'fade', direction: 'smooth' }, 26: { type: 'circle', direction: null }, 27: { type: 'circle', direction: null },
    30: { type: 'fade', direction: 'smooth' },
  }[effect] ?? null;
  const advanceAfter = flags & 0x0400 ? Math.max(0, time) / 1000 : null;
  const advanceOnClick = Boolean(flags & 0x0001);
  const transition = as || advanceAfter != null || !advanceOnClick
    ? { type: as?.type ?? 'none', direction: as?.direction ?? null, duration: as ? [1, 0.75, 0.5][speed] ?? 1 : null, advanceOnClick, advanceAfter }
    : null;
  return { hidden: Boolean(flags & 0x0004), transition };
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

/**
 * A text type's master style: subtitles and half/quarter bodies take the
 * body's, centred titles the title's, each level with its own look over it
 * where it has one — a centred title's own style is often empty, and its
 * font the title's.
 */
function styleFor(styles, textType) {
  const order = { 0: [0], 1: [1], 2: [2, 1], 4: [4, 1], 5: [5, 1], 6: [6, 0], 7: [7, 1], 8: [8, 1] }[textType] || [4, 1];
  const out = [];
  for (const k of [...order].reverse()) {
    (styles[k] || []).forEach((lv, i) => { out[i] = { pf: { ...(out[i]?.pf || {}), ...lv.pf }, cf: { ...(out[i]?.cf || {}), ...lv.cf } }; });
  }
  return out;
}

/** A colour in a run, a bullet or a shape: a scheme entry by number, or its own RGB. */
function colourOf(value, scheme) {
  const index = value >>> 24;
  if (index === 0xfe) return [0, 8, 16].map((s) => ((value >>> s) & 0xff).toString(16).padStart(2, '0')).join('').toUpperCase();
  if (index < 8) return scheme[index] ?? null;
  return null;
}

/**
 * A slide's (or master's) drawing: its shapes, in slide pixels, and its
 * background's colour. Groups are walked into, each child's box carried
 * from the group's own coordinates to the slide's.
 */
function readDrawing(doc, drawing, ctx) {
  const out = { shapes: [], background: null, items: [], backgroundML: null };
  const dg = child(doc, drawing, T.DgContainer);
  if (!dg) return out;
  const identity = (r) => r;
  for (const c of children(doc, dg)) {
    if (c.type === T.SpgrContainer) walkGroup(doc, c, identity, ctx, out.shapes, true, out.items);
    else if (c.type === T.SpContainer) {
      const props = readFopt(doc, child(doc, c, T.FOPT));
      const fsp = child(doc, c, T.FSP);
      if (fsp && u32(doc, fsp.body + 4) & 0x400) { out.background = fillOf(props, ctx, true) || null; out.backgroundML = keptOf(doc, c); }
    }
  }
  return out;
}

/** The DrawingML a later PowerPoint kept for a shape (or a group, its own shape's), beside the older description: a package, or null. */
function keptOf(doc, sp) {
  const own = sp.type === T.SpgrContainer ? children(doc, sp)[0] : sp;
  if (!own || own.type !== T.SpContainer) return null;
  const blob = readFopt(doc, child(doc, own, T.TertiaryFOPT)).get(0x03a9)?.complex;
  return blob && blob[0] === 0x50 && blob[1] === 0x4b ? blob.slice() : null;
}

/** Groups inside groups deeper than any slide has: a file nesting past this is refused its inner shapes rather than the stack. */
const GROUP_DEPTH = 64;

function walkGroup(doc, group, transform, ctx, shapes, top, items = null, depth = 0) {
  if (depth > GROUP_DEPTH) return;
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
  // A table is a group its own shape marks as one: its cells, and the lines between them.
  if (own.type === T.SpContainer && !top && (readFopt(doc, child(doc, own, T.TertiaryFOPT)).get(0x03a0)?.op ?? 0) & 3) {
    const table = readTable(doc, kids.slice(1), inner, ctx);
    if (table) { shapes.push(table); return; }
  }
  for (const k of kids.slice(1)) {
    const before = shapes.length;
    if (k.type === T.SpgrContainer) walkGroup(doc, k, inner, ctx, shapes, false, null, depth + 1);
    else if (k.type === T.SpContainer) {
      const shape = readShape(doc, k, inner, ctx);
      if (shape) shapes.push(shape);
    }
    // At the top, each drawing in order: what a later PowerPoint kept of it, and what was read of it.
    if (items) {
      const own = k.type === T.SpgrContainer ? children(doc, k)[0] : k;
      const fsp = own?.type === T.SpContainer ? child(doc, own, T.FSP) : null;
      if (fsp && u32(doc, fsp.body + 4) & 0x408) continue; // deleted, or the background
      items.push({ spid: fsp ? u32(doc, fsp.body) : null, group: k.type === T.SpgrContainer, drawingML: keptOf(doc, k), shapes: shapes.slice(before) });
    }
  }
}

/**
 * A table, from the group PowerPoint keeps one as: each cell a box with its
 * words and fill, each border a line. Its rows are where the cells' tops
 * are and its columns where their lefts are (LibreOffice's reading); a cell
 * reaching over more of them spans them, and each line is the border of the
 * cells along it.
 */
function readTable(doc, kids, transform, ctx) {
  const boxes = [];
  const lines = [];
  for (const k of kids) {
    if (k.type !== T.SpContainer) continue;
    const fsp = child(doc, k, T.FSP);
    const box = anchorOf(doc, k, transform);
    if (!fsp || !box) continue;
    const props = readFopt(doc, child(doc, k, T.FOPT));
    if (LINES.has(fsp.inst)) { lines.push({ ...box, line: lineOf(props, ctx, true) }); continue; }
    const anchor = props.get(0x0087)?.op ?? 0;
    boxes.push({
      ...box,
      // A cell's shading drawn as its first colour: a table cell holds one.
      fill: ((f) => (f?.gradient ? f.gradient.stops[0].color : typeof f === 'object' ? f.color : f))(props.has(0x0181) || props.get(0x01bf)?.op & 0x100000 ? fillOf(props, ctx) : 'none'),
      paragraphs: shapeText(doc, child(doc, k, T.ClientTextbox), ctx, null) || [],
      anchor: [1, 4].includes(anchor) ? 'middle' : [2, 5, 7, 9].includes(anchor) ? 'bottom' : 'top',
    });
  }
  if (!boxes.length) return null;
  // Edges within a pixel of each other are one edge.
  const edges = (values) => values.sort((a, b) => a - b).filter((v, i, all) => i === 0 || v - all[i - 1] > 1);
  const tops = edges(boxes.map((b) => b.y));
  const lefts = edges(boxes.map((b) => b.x));
  const bottom = Math.max(...boxes.map((b) => b.y + b.h));
  const right = Math.max(...boxes.map((b) => b.x + b.w));
  const at = (list, v) => list.findIndex((e) => Math.abs(e - v) <= 1);
  const rows = tops.map((t, i) => (tops[i + 1] ?? bottom) - t);
  const columns = lefts.map((l, i) => (lefts[i + 1] ?? right) - l);
  const cells = tops.map(() => lefts.map(() => ({ paragraphs: [], fill: 'none', anchor: 'top', borders: {}, covered: true })));
  for (const b of boxes) {
    const r = at(tops, b.y);
    const c = at(lefts, b.x);
    if (r < 0 || c < 0) continue;
    const rowSpan = tops.filter((t) => t >= b.y - 1 && t < b.y + b.h - 1).length || 1;
    const colSpan = lefts.filter((l) => l >= b.x - 1 && l < b.x + b.w - 1).length || 1;
    cells[r][c] = { paragraphs: b.paragraphs, fill: b.fill, anchor: b.anchor, borders: {}, rowSpan, colSpan, covered: false };
    for (let i = 0; i < rowSpan; i++) {
      for (let j = 0; j < colSpan; j++) if ((i || j) && cells[r + i]?.[c + j]) cells[r + i][c + j] = { ...cells[r + i][c + j], covered: true, hMerge: j > 0, vMerge: i > 0 };
    }
  }
  // Each line, the border of the cells either side of it along its length.
  const edgeAt = (list, end, v) => (Math.abs(v - end) <= 1 ? list.length : at(list, v));
  for (const l of lines) {
    if (Math.abs(l.h) <= 1) {
      const r = edgeAt(tops, bottom, l.y);
      if (r < 0) continue;
      lefts.forEach((x, c) => {
        if (x + columns[c] / 2 < l.x || x + columns[c] / 2 > l.x + l.w) return;
        if (r < tops.length) cells[r][c].borders.top = l.line;
        if (r > 0) cells[r - 1][c].borders.bottom = l.line;
      });
    } else if (Math.abs(l.w) <= 1) {
      const c = edgeAt(lefts, right, l.x);
      if (c < 0) continue;
      tops.forEach((y, r) => {
        if (y + rows[r] / 2 < l.y || y + rows[r] / 2 > l.y + l.h) return;
        if (c < lefts.length) cells[r][c].borders.left = l.line;
        if (c > 0) cells[r][c - 1].borders.right = l.line;
      });
    }
  }
  // A merged cell's far edges are those of the last cells it covers; an edge no line runs along has no border.
  cells.forEach((row, r) => row.forEach((cell, c) => {
    if (cell.covered) return;
    if (cell.rowSpan > 1) cell.borders.bottom = cells[r + cell.rowSpan - 1]?.[c]?.borders.bottom;
    if (cell.colSpan > 1) cell.borders.right = cells[r]?.[c + cell.colSpan - 1]?.borders.right;
  }));
  for (const row of cells) for (const cell of row) for (const side of ['left', 'right', 'top', 'bottom']) cell.borders[side] ??= 'none';
  return { type: 'table', x: lefts[0], y: tops[0], w: right - lefts[0], h: bottom - tops[0], rows, columns, cells };
}

/**
 * A shape's build in the show, as PowerPoint 97 kept it (its
 * AnimationInfoAtom): its place in the slide's order, and its effect and
 * direction as the entrance effect a .pptx names — LibreOffice's table of
 * them, each of PowerPoint 97's as the nearest the deck writes — on a click,
 * or by itself after the one before. A later PowerPoint's effect comes as
 * the PowerPoint 97 one it saved beside it.
 */
function animationOf(doc, data) {
  const atom = data ? child(doc, child(doc, data, T.AnimationInfo), T.AnimationInfoAtom) : null;
  if (!atom || atom.len < 24) return null;
  const d = atom.body;
  if (!u8(doc, d + 20)) return null; // no build
  const flags = u32(doc, d + 4);
  const delay = i32(doc, d + 12);
  const method = u8(doc, d + 21);
  const dir = u8(doc, d + 22);
  const SIDES = ['left', 'top', 'right', 'bottom', 'top-left', 'top-right', 'bottom-left', 'bottom-right'];
  let effect = 'appear';
  let direction = null;
  let duration = null;
  if (method === 0x01 || method === 0x03 || method === 0x05) effect = 'fade';
  else if (method === 0x02 || method === 0x08 || method === 0x09) effect = 'wipe';
  else if (method === 0x0a) { effect = 'wipe'; direction = ['right', 'bottom', 'left', 'top'][dir] ?? null; }
  else if (method === 0x0b) effect = 'zoom';
  else if (method === 0x0d) { effect = 'split'; direction = ['horizontal-out', 'horizontal-in', 'vertical-out', 'vertical-in'][dir] ?? null; }
  else if (method === 0x0c) {
    if (dir <= 7) { effect = 'fly'; direction = SIDES[dir]; }
    else if (dir <= 0x0b) { effect = 'wipe'; direction = ['left', 'bottom', 'right', 'top'][dir - 8]; }
    else if (dir <= 0x0f) { effect = 'fly'; direction = SIDES[dir - 0x0c]; duration = 1; }
    else effect = 'zoom';
  }
  return {
    order: u16(doc, d + 16), effect, direction, duration,
    trigger: flags & 0x04 ? 'afterPrevious' : 'onClick',
    delay: delay > 0 && delay !== 0x7fffffff ? delay / 1000 : 0,
  };
}

/**
 * A slide's effects as PowerPoint 2002 and later keep them — beside the
 * PowerPoint 97 builds, in the slide's "___PPT10" binary tag: a tree of
 * time nodes the same shape as a .pptx's timing. Each effect node says its
 * class (entrance, emphasis, exit), its preset and subtype and how it
 * starts; the shape it acts on is named in a reference under it. In the
 * main sequence's order: [{ spid, kind, effect, direction, trigger }].
 */
function timingOf(doc, slideH) {
  const tags = child(doc, slideH, T.ProgTags);
  let root = null;
  for (const tag of children(doc, tags)) {
    if (tag.type !== T.ProgBinaryTag) continue;
    const name = child(doc, tag, T.CString);
    if (!name || utf16(doc.subarray(name.body, name.end)) !== '___PPT10') continue;
    root = child(doc, child(doc, tag, T.BinaryTagData), T.AnimGroup);
  }
  if (!root) return null;
  const props = (node) => {
    const out = new Map();
    for (const v of children(doc, child(doc, node, T.AnimPropertySet))) {
      if (v.type === T.AnimAttributeValue && v.len >= 5 && u8(doc, v.body) === 1) out.set(v.inst, i32(doc, v.body + 1));
    }
    return out;
  };
  const shapeOf = (node) => {
    for (const c of children(doc, node)) {
      if (c.type === T.AnimReference && c.len >= 12 && i32(doc, c.body + 4) === 1) return i32(doc, c.body + 8);
      if (c.ver === 0xf) { const found = shapeOf(c); if (found != null) return found; }
    }
    return null;
  };
  const SIDES = { 1: 'top', 2: 'right', 4: 'bottom', 8: 'left', 3: 'top-right', 6: 'bottom-right', 9: 'top-left', 12: 'bottom-left' };
  const SPLITS = { 37: 'vertical-out', 42: 'horizontal-out', 21: 'vertical-in', 26: 'horizontal-in' };
  const effects = [];
  const walk = (node) => {
    const p = props(node);
    const type = p.get(20);
    if (type === 5) return; // a sequence a shape's click starts: not the slide's own
    const cls = p.get(11);
    if (cls != null && type >= 1 && type <= 3) {
      const kind = { 1: 'entr', 2: 'exit', 3: 'emph' }[cls];
      const spid = shapeOf(node);
      if (kind && spid != null) {
        const id = p.get(9) ?? 0;
        const sub = p.get(10) ?? 0;
        let effect;
        let direction = null;
        if (kind === 'emph') effect = { 26: 'pulse', 8: 'spin', 6: 'grow' }[id] ?? 'pulse';
        else if (id === 1) effect = 'appear';
        else if (id === 2 || id === 7) { effect = 'fly'; direction = SIDES[sub] ?? 'bottom'; }
        else if (id === 42 || id === 47) { effect = 'float'; direction = (id === 42) === (kind === 'entr') ? 'up' : 'down'; }
        else if (id === 16) { effect = 'split'; direction = SPLITS[sub] ?? 'vertical-out'; }
        else if (id === 22 || id === 12) { effect = 'wipe'; direction = SIDES[sub] ?? 'bottom'; }
        else if ([3, 14, 18, 20, 21].includes(id)) effect = 'wipe';
        else if ([4, 6, 8, 13, 15, 17, 19, 23, 53].includes(id)) effect = 'zoom';
        else effect = 'fade';
        effects.push({ spid, kind, effect, direction, trigger: { 1: 'onClick', 2: 'withPrevious', 3: 'afterPrevious' }[type] });
        return;
      }
    }
    for (const c of children(doc, node)) if (c.type === T.AnimGroup || c.type === T.AnimSubGroup) walk(c);
  };
  walk(root);
  return effects;
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

/** A shape's fill: Office Art's own reading (officeart.js), in the slide's colour scheme. */
const fillOf = (props, ctx, background = false) => artFill(props, ctx.scheme, background);
/** A shape's outline, or 'none': black, three quarters of a point, where it does not say. */
const lineOf = (props, ctx, defaultOn) => artLine(props, ctx.scheme, defaultOn);

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
  // Its id, which a later PowerPoint's effects name it by, where its words sit in it when it says, and its PowerPoint 97 build.
  base.spid = u32(doc, fsp.body);
  if (ph) { base.placeholder = placement; base.placeholderIdx = i32(doc, ph.body); }
  // Its shadow.
  const shadow = artShadow(props, ctx.scheme);
  if (shadow) base.shadow = shadow;
  const anchorText = props.get(0x0087)?.op;
  if (anchorText != null) base.anchor = [1, 4].includes(anchorText) ? 'middle' : [2, 5, 7, 9].includes(anchorText) ? 'bottom' : 'top';
  const animation = animationOf(doc, data);
  if (animation) base.animation = animation;

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
  // A placeholder's words with only their own looks, for a deck whose layouts give it the rest.
  if (ph && paragraphs) base.ownParagraphs = shapeText(doc, child(doc, sp, T.ClientTextbox), { ...ctx, styles: {} }, placement);
  if (paragraphs?.words != null) base.words = paragraphs.words;
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
  // Each run's language (TextSpecialInfoAtom), and its alternate: a bidirectional text's order turns on them.
  const siRuns = [];
  const special = records.find((r) => r.type === 0x0faa);
  if (special) {
    let p = special.body;
    let covered = 0;
    while (covered <= text.length && p + 8 <= special.end) {
      const count = u32(doc, p);
      const m = u32(doc, p + 4);
      p += 8;
      if (m & 0x1) p += 2;
      const lang = m & 0x2 ? languageTag(u16(doc, (p += 2) - 2)) : null;
      const altLang = m & 0x4 ? languageTag(u16(doc, (p += 2) - 2)) : null;
      if (m & 0x40) p += 2;
      if (m & 0x20) p += 4;
      if (m & 0x200) p += 4 + u32(doc, p) * 4;
      siRuns.push({ count, lang, altLang });
      covered += count;
      if (!count) break;
    }
  }
  const at = (runs, i) => {
    let sum = 0;
    for (const r of runs) { sum += r.count; if (i < sum) return r; }
    return runs[runs.length - 1] || null;
  };

  const master = styleFor(ctx.styles, textType);
  // The words without their fields (a slide number, a date, a header or footer), whose values move.
  const fields = new Set(records.filter((r) => FIELDS.has(r.type)).map((r) => i32(doc, r.body)));
  const words = [...text].filter((_, i) => !fields.has(i)).join('');
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
    const para = { level, runs: [], own: pr?.pf || {} };
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
      const si = siRuns.length ? at(siRuns, start + k) : null;
      if (si?.lang) look.lang = si.lang;
      if (si?.altLang && si.altLang !== si.lang) look.altLang = si.altLang;
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
  paragraphs.words = words;
  return paragraphs;
}

/** The records that mark a field in a text: a slide number, dates, a header, a footer. */
const FIELDS = new Set([0x0fd8, 0x0ff7, 0x0ff8, 0x0ff9, 0x0ffa, 0x1015]);

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
