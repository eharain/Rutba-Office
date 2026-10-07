/**
 * Word documents in the formats no Office on hand still writes — Word 6.0/95,
 * Word 2.0, Windows Write and Word for DOS — laid out byte by byte from
 * their published layouts (LibreOffice's Word import, and libwps for Write
 * and Word for DOS), so the readers have files of each kind to open.
 *
 * Each builder takes a small description — paragraphs of runs, each with
 * its properties in the version's own form (one-byte sprms for Word 6 and
 * 2, Word 2's CHPX structure, Write's property bytes) — and returns the
 * file's bytes.
 */
import { writeCompoundFile } from '@rutba/office-formats/cfb-write';

const le16 = (v) => [v & 0xff, (v >> 8) & 0xff];
const le32 = (v) => [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
const bytesOf = (t) => (typeof t === 'string' ? [...t].map((c) => c.charCodeAt(0) & 0xff) : [...t]);
const even = (a) => (a.length % 2 ? [...a, 0] : a);

/** A file under construction: bytes appended, and written back at fixed places. */
class Out {
  constructor(size = 0) { this.b = new Array(size).fill(0); }
  get length() { return this.b.length; }
  at(pos, arr) { for (let i = 0; i < arr.length; i++) this.b[pos + i] = arr[i] & 0xff; }
  push(arr) { const at = this.b.length; this.b.push(...arr.map((x) => x & 0xff)); return at; }
  padTo(n) { while (this.b.length < n) this.b.push(0); }
  bytes() { return Uint8Array.from(this.b); }
}

/**
 * Pages of property runs, 512 bytes each: the runs' file positions, one
 * offset (and `bx - 1` bytes of layout cache) per run, the properties from
 * the page's end down, and the count in the last byte. Runs whose
 * properties are the same share them.
 */
function fkpPages(runs, bx) {
  const pages = [];
  let i = 0;
  while (i < runs.length) {
    const page = new Array(512).fill(0);
    const chosen = [];
    const shared = new Map();
    let top = 511;
    for (; i < runs.length; i++) {
      const r = runs[i];
      const key = r.blob ? r.blob.join(',') : '';
      let off = r.blob ? shared.get(key) : 0;
      let newTop = top;
      if (r.blob && off == null) { newTop = (top - r.blob.length) & ~1; off = newTop; }
      if ((chosen.length + 2) * 4 + (chosen.length + 1) * bx > newTop) break;
      if (r.blob && !shared.has(key)) { r.blob.forEach((x, k) => { page[newTop + k] = x; }); shared.set(key, newTop); top = newTop; }
      chosen.push({ r, off });
    }
    const n = chosen.length;
    chosen.forEach(({ r }, k) => le32(r.fcStart).forEach((x, j) => { page[k * 4 + j] = x; }));
    le32(chosen[n - 1].r.fcEnd).forEach((x, j) => { page[n * 4 + j] = x; });
    chosen.forEach(({ off }, k) => { page[(n + 1) * 4 + k * bx] = off / 2; });
    page[511] = n;
    pages.push({ page, fcStart: chosen[0].r.fcStart, fcEnd: chosen[n - 1].r.fcEnd });
  }
  return pages;
}

/**
 * A Word 6.0/95 document (`version: 6`, a compound file) or a Word 2.0 one
 * (`version: 2`, no container). The description:
 *   styles      Word 6: [{ sti, kind: 1|2, base, next, name, papx, chpx }] by istd
 *               Word 2: { cstcStd, entries: [{ name, chpx, papx: { stc, sprms }, next, base }] }
 *   fonts       [{ name, charset }]
 *   main, footnotes, header   paragraphs: [{ istd, papx, runs: [{ text, chpx }], mark }]
 *   sepx        the section's sprms
 *   pictures    [{ at: name, pic: bytes }] placed in the stream; a run's chpx may
 *               name one as { picture: name } to point at it
 */
export function buildOldWord({ version, styles, fonts, main, footnotes = [], header = null, sepx = [], pictures = [] }) {
  const w2 = version === 2;
  const fcMin = w2 ? 0x180 : 0x300;
  const text = [];
  const chp = [];
  const pap = [];
  let refCp = null;
  const pictureRefs = [];

  const story = (paras) => {
    for (const p of paras) {
      const start = text.length;
      for (const r of p.runs) {
        const b = bytesOf(r.text);
        if (r.noteRef) refCp = text.length;
        const run = { cpStart: text.length, cpEnd: text.length + b.length, chpx: r.chpx || [] };
        if (r.picture) pictureRefs.push({ run, name: r.picture });
        chp.push(run);
        text.push(...b);
      }
      chp.push({ cpStart: text.length, cpEnd: text.length + 1, chpx: p.markChpx || [] });
      text.push(p.mark ?? 13);
      pap.push({ cpStart: start, cpEnd: text.length, istd: p.istd ?? 0, papx: p.papx || [] });
    }
    return text.length;
  };
  const ccpText = story(main);
  const ccpFtn = footnotes.length ? story(footnotes) + 1 - ccpText : 0;
  if (footnotes.length) { text.push(13); chp.push({ cpStart: text.length - 1, cpEnd: text.length, chpx: [] }); pap.push({ cpStart: text.length - 1, cpEnd: text.length, istd: 0, papx: [] }); }
  const hddStart = text.length;
  const ccpHdd = header ? story(header) + 1 - hddStart : 0;
  if (header) { text.push(13); chp.push({ cpStart: text.length - 1, cpEnd: text.length, chpx: [] }); pap.push({ cpStart: text.length - 1, cpEnd: text.length, istd: 0, papx: [] }); }

  const f = new Out(fcMin);
  f.push(text);
  const fcMac = f.length;

  // Pictures go where the runs that show them can point.
  const pictureAt = new Map();
  f.padTo(Math.ceil(f.length / 2) * 2);
  for (const p of pictures) pictureAt.set(p.name, f.push(p.pic));
  for (const { run, name } of pictureRefs) {
    const fc = pictureAt.get(name);
    run.chpx = w2 ? [...run.chpx, ...new Array(20).fill(0)].slice(0, 20).concat(le32(fc)) : [...run.chpx, 68, 4, ...le32(fc)];
  }

  // The property pages.
  f.padTo(Math.ceil(f.length / 512) * 512);
  const chpBlob = (c) => (c.chpx.length ? [c.chpx.length, ...c.chpx] : null);
  const papBlob = (p) => {
    const data = even(w2 ? [p.istd, 0, 0, 0, 0, 0, 0, ...p.papx] : [...le16(p.istd), ...p.papx]);
    return [data.length / 2, ...data];
  };
  const fcOf = (cp) => fcMin + cp;
  const chpPages = fkpPages(chp.map((c) => ({ fcStart: fcOf(c.cpStart), fcEnd: fcOf(c.cpEnd), blob: chpBlob(c) })), 1);
  const papPages = fkpPages(pap.map((p) => ({ fcStart: fcOf(p.cpStart), fcEnd: fcOf(p.cpEnd), blob: papBlob(p) })), w2 ? 1 : 7);
  const pnChp = f.length / 512;
  for (const p of chpPages) f.push(p.page);
  const pnPap = f.length / 512;
  for (const p of papPages) f.push(p.page);

  // The tables, each recorded for the FIB.
  const pairs = {};
  const table = (index, arr) => { pairs[index] = { fc: f.push(arr), lcb: arr.length }; };
  const bte = (pages, pn) => [...pages.flatMap((p) => le32(p.fcStart)), ...le32(pages[pages.length - 1].fcEnd), ...pages.flatMap((_, k) => le16(pn + k))];

  if (w2) table(1, stsh2(styles));
  else table(1, stsh6(styles));
  if (refCp != null) {
    table(2, [...le32(refCp), ...le32(ccpText), ...le16(1)]);
    table(3, [...le32(0), ...le32(ccpFtn - 1), ...le32(ccpFtn)]);
  }
  const sepxAt = f.push(w2 ? [sepx.length, ...sepx] : [...le16(sepx.length), ...sepx]);
  table(6, [...le32(0), ...le32(ccpText), ...le16(0), ...le32(sepxAt), ...(w2 ? [] : [...le16(0), ...le32(0)])]);
  if (header) table(11, [...le32(0), ...le32(ccpHdd - 1), ...le32(ccpHdd)]);
  table(12, bte(chpPages, pnChp));
  table(13, bte(papPages, pnPap));
  table(15, w2 ? fonts2(fonts) : fonts6(fonts));
  table(31, [0, 0]);

  // The FIB.
  f.at(0, le16(w2 ? 0xa5db : 0xa5dc));
  f.at(2, le16(w2 ? 45 : 101));
  f.at(6, le16(0x0409));
  f.at(0x18, le32(fcMin));
  f.at(0x1c, le32(fcMac));
  f.at(0x20, le32(f.length));
  f.at(0x34, le32(ccpText));
  f.at(0x38, le32(ccpFtn));
  f.at(0x3c, le32(ccpHdd));
  for (const [i, p] of Object.entries(pairs)) {
    const at = 0x58 + Number(i) * (w2 ? 6 : 8);
    f.at(at, le32(p.fc));
    f.at(at + 4, w2 ? le16(p.lcb) : le32(p.lcb));
  }
  const tail = 0x58 + 38 * (w2 ? 6 : 8) + 2;
  f.at(tail, le16(pnChp));
  f.at(tail + 2, le16(pnPap));
  f.at(tail + 4, le16(chpPages.length));
  f.at(tail + 6, le16(papPages.length));

  const bytes = f.bytes();
  return w2 ? bytes : new Uint8Array(writeCompoundFile([{ path: ['WordDocument'], data: bytes }]));
}

/** Word 6's style sheet: the header, then each style's base, name and properties. */
function stsh6(styles) {
  const out = [...le16(14), ...le16(styles.length), ...le16(8), 0, 0, ...le16(0), ...le16(0), ...le16(0), ...le16(0)];
  styles.forEach((s, istd) => {
    if (!s) { out.push(...le16(0)); return; }
    const upx = s.kind === 1 ? [[...le16(istd), ...(s.papx || [])], s.chpx || []] : [s.chpx || []];
    let std = [...le16(s.sti), ...le16(s.kind | ((s.base ?? 0xfff) << 4)), ...le16(upx.length | ((s.next ?? istd) << 4)), ...le16(0)];
    std.push(s.name.length, ...bytesOf(s.name), 0);
    std = even(std);
    for (const u of upx) std = even([...std, ...le16(u.length), ...u]);
    out.push(...le16(std.length), ...std);
  });
  return out;
}

/** Word 2's style sheet: built-in styles first, then names, character and paragraph properties, next and base. */
function stsh2({ cstcStd, entries }) {
  const names = entries.flatMap((e) => (e.name == null ? [0xff] : [e.name.length, ...bytesOf(e.name)]));
  const chpxs = entries.flatMap((e) => (e.chpx == null ? [0xff] : [e.chpx.length, ...e.chpx]));
  const papxs = entries.flatMap((e) => (e.papx == null ? [0xff] : [7 + e.papx.sprms.length, e.papx.stc, 0, 0, 0, 0, 0, 0, ...e.papx.sprms]));
  return [
    ...le16(cstcStd),
    ...le16(names.length + 2), ...names,
    ...le16(chpxs.length + 2), ...chpxs,
    ...le16(papxs.length + 2), ...papxs,
    ...le16(entries.length), ...entries.flatMap((e) => [e.next ?? 222, e.base ?? 222]),
  ];
}

function fonts6(fonts) {
  const entries = fonts.flatMap((ft) => {
    const e = [0, 0x20 | 2, ...le16(400), ft.charset ?? 0, 0, ...bytesOf(ft.name), 0];
    e[0] = e.length - 1;
    return e;
  });
  return [...le16(entries.length + 2), ...entries];
}

function fonts2(fonts) {
  const entries = fonts.flatMap((ft) => {
    const e = [0, 0, ft.charset ?? 0, ...bytesOf(ft.name), 0];
    e[0] = e.length - 1;
    return e;
  });
  return [...le16(entries.length + 2), ...entries];
}

/** Word 6's ANLD: one level of numbering — format, text before and after, flags, font, start, indent. */
export function anld6({ nfc = 0, before = [], after = [], hang = true, font = 0, start = 1, indent = 360 }) {
  const a = [nfc, before.length, after.length, hang ? 0x08 : 0, 0, 0, ...le16(font), ...le16(0), ...le16(start), ...le16(indent), ...le16(0), 0, 0, 0, 0];
  const chars = [...before, ...after];
  for (let i = 0; i < 32; i++) a.push(chars[i] ?? 0);
  return [12, a.length, ...a];
}

/** A table row's cell edges and borders, as Word 6's sprmTDefTable. */
export function defTable6(edges, border = 0x0009) {
  const data = [edges.length - 1, ...edges.flatMap(le16)];
  for (let i = 0; i < edges.length - 1; i++) data.push(0, 0, ...le16(border), ...le16(border), ...le16(border), ...le16(border));
  return [190, ...le16(data.length + 1), ...data];
}

/** The smallest Windows metafile with a window: 1000 by 500 units. */
export function tinyWmf() {
  const records = [...le32(5), ...le16(0x020b), ...le16(0), ...le16(0), ...le32(5), ...le16(0x020c), ...le16(500), ...le16(1000), ...le32(3), ...le16(0)];
  const words = (18 + records.length) / 2;
  return [...le16(1), ...le16(9), ...le16(0x0300), ...le32(words), ...le16(0), ...le32(5), ...le16(0), ...records];
}

/** A Word 6 picture: its PIC header (68 bytes, the size it shows at), then the metafile. */
export function pic6(wmf, dxaGoal, dyaGoal) {
  const h = new Array(68).fill(0);
  const put = (at, arr) => arr.forEach((x, i) => { h[at + i] = x; });
  put(0, le32(68 + wmf.length));
  put(4, le16(68));
  put(6, le16(8));
  put(28, le16(dxaGoal));
  put(30, le16(dyaGoal));
  put(32, le16(1000));
  put(34, le16(1000));
  return [...h, ...wmf];
}

/**
 * A Windows Write (`dos: false`) or Word for DOS (`dos: true`) document:
 * the header, the text from byte 128, then pages of character runs,
 * paragraph runs, the section, the fonts (Write) and footnotes (DOS).
 *   paragraphs  [{ runs: [{ text, prop }], pap, object }] — `object` is a
 *               picture's bytes in place of text
 *   fonts       Write's font names
 *   notes       DOS: [{ refIndex: paragraph and offset of the reference, textIndex }]
 */
export function buildWriteDoc({ dos = false, codepage = 0, paragraphs, fonts = ['Arial'], notes = [], section = null }) {
  const text = [];
  const chars = [];
  const paras = [];
  const at = [];
  for (const p of paragraphs) {
    const start = 0x80 + text.length;
    at.push(start);
    if (p.object) {
      text.push(...p.object);
      chars.push({ fcLim: 0x80 + text.length, prop: [] });
    } else {
      for (const r of p.runs) {
        text.push(...bytesOf(r.text));
        chars.push({ fcLim: 0x80 + text.length, prop: r.prop || [] });
      }
      text.push(13, 10);
      chars.push({ fcLim: 0x80 + text.length, prop: p.runs[p.runs.length - 1]?.prop || [] });
    }
    paras.push({ fcLim: 0x80 + text.length, prop: p.pap || [] });
  }
  const f = new Out(128);
  f.push(text);
  const fcMac = f.length;
  f.padTo(Math.ceil(f.length / 128) * 128);

  // Property pages: each run's end and where its bytes are; the bytes from the page's end down.
  const fodPages = (runs) => {
    const first = f.length / 128;
    let i = 0;
    let fc = 0x80;
    while (i < runs.length) {
      const page = new Array(128).fill(0);
      le32(fc).forEach((x, k) => { page[k] = x; });
      let n = 0;
      let top = 127;
      for (; i < runs.length && n < 20; i++) {
        const r = runs[i];
        const need = 4 + (n + 1) * 6;
        let bf = 0xffff;
        if (r.prop.length) {
          const newTop = top - (r.prop.length + 1);
          if (newTop < need) break;
          page[newTop] = r.prop.length;
          r.prop.forEach((x, k) => { page[newTop + 1 + k] = x; });
          bf = newTop - 4;
          top = newTop;
        } else if (need > top) break;
        le32(r.fcLim).forEach((x, k) => { page[4 + n * 6 + k] = x; });
        le16(bf).forEach((x, k) => { page[4 + n * 6 + 4 + k] = x; });
        fc = r.fcLim;
        n += 1;
      }
      page[127] = n;
      f.push(page);
    }
    return first;
  };
  fodPages(chars);
  const pnPara = fodPages(paras);

  // The section: the page and the text's place on it.
  const pnSetb = f.length / 128;
  const setb = new Array(128).fill(0);
  setb[0] = 2;
  const sepAt = (pnSetb + 1) * 128;
  le32(0).forEach((x, k) => { setb[4 + k] = x; });
  le32(sepAt).forEach((x, k) => { setb[10 + k] = x; });
  f.push(setb);
  const sep = new Array(128).fill(0);
  const s = { yaMac: 15840, xaMac: 12240, yaTop: 1440, dyaText: 12960, xaLeft: 1800, dxaText: 8640, yaHeader: 1080, yaFooter: 14760, ...(section || {}) };
  sep[0] = 25;
  [[3, s.yaMac], [5, s.xaMac], [9, s.yaTop], [11, s.dyaText], [13, s.xaLeft], [15, s.dxaText], [19, s.yaHeader], [21, s.yaFooter]].forEach(([o, v]) => le16(v).forEach((x, k) => { sep[o + k] = x; }));
  f.push(sep);

  let pnFfntb = 0;
  if (!dos) {
    pnFfntb = f.length / 128;
    const page = [...le16(fonts.length)];
    for (const name of fonts) page.push(...le16(name.length + 2), 0x10, ...bytesOf(name), 0);
    page.push(0, 0);
    f.push(page);
    f.padTo(Math.ceil(f.length / 128) * 128);
  }
  let pnFntb = 0;
  if (dos && notes.length) {
    pnFntb = f.length / 128;
    const entries = [...notes.map((n) => ({ ref: at[n.ref[0]] + n.ref[1], text: at[n.text] })), { ref: fcMac, text: fcMac }];
    const page = [...le16(entries.length), ...le16(entries.length)];
    for (const e of entries) page.push(...le32(e.ref - 0x80), ...le32(e.text - 0x80));
    f.push(page);
    f.padTo(Math.ceil(f.length / 128) * 128);
  }
  const pnMac = f.length / 128;

  f.at(0, [0x31, 0xbe, 0, 0, 0, 0xab]);
  f.at(14, le32(fcMac));
  f.at(18, le16(pnPara));
  f.at(20, le16(dos ? pnFntb : 0));
  f.at(22, le16(dos ? pnFntb + 1 : 0));
  f.at(24, le16(pnSetb));
  f.at(26, le16(pnSetb + 2));
  if (!dos) f.at(28, le16(pnFfntb));
  if (!dos) f.at(96, le16(pnMac));
  if (dos) f.at(126, le16(codepage));
  return f.bytes();
}
