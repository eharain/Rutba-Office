// Excel's binary workbooks — BIFF8 (Excel 97-2003, .xls), BIFF5 (Excel
// 5.0 and 95) and BIFF2 to BIFF4 (Excel 2.x to 4.0) — read into a plain
// model that msxls-xlsx.js writes as the .xlsx this suite edits.
//
// A binary workbook is a stream of records, each an id, a length and its
// bytes: in a compound file's Workbook (BIFF8) or Book (BIFF5) stream, or
// the whole file before Excel 5. The workbook's own records come first —
// fonts, number formats, cell formats (XFs), the palette, the sheets and
// the defined names, and, in BIFF8, every string the cells use (the SST) —
// then each sheet's records between a BOF and an EOF: its cells, rows,
// columns, merged areas and window. Before Excel 5 a file is one sheet and
// all of it is in one run.
//
// A formula is stored compiled — tokens in reverse Polish order — and is
// read back here into the text Excel would show, references, functions,
// names, other sheets and constant arrays included; a cell that shares
// its formula with a block of others is given the formula as it reads in
// that cell. Each formula keeps the value Excel last calculated.
//
// A sheet's pictures are Office Art in BIFF8 (the workbook's drawing group
// holds their bytes); before it each is an object record with its bytes in
// the IMGDATA record after it. Notes and hyperlinks come with their cells.
// A chart is a substream of its own after the object that places it, or a
// chart sheet; msxls-chart.js reads it once every sheet is known.
//
// The layouts follow LibreOffice's Excel import (sc/source/filter/excel).
// Pure: bytes in, a model out.

import { CompoundFile } from './cfb.js';
import { decoderFor } from './codepage.js';
import { XLS_FUNCTIONS } from './xls-functions.js';
import { findBlip, dibFile } from './msdoc.js';
import { placeableWmf } from './msdoc-old.js';
import { readChart } from './msxls-chart.js';
import { header as artHeader, children as artChildren, child as artChild, readFopt, PRESETS, LINES as ART_LINES, freeformPath, artFill, artLine, artShadow } from './officeart.js';

export class XlsError extends Error {
  constructor(message) {
    super(message);
    this.name = 'XlsError';
  }
}

const u8 = (b, at) => b[at] ?? 0;
const u16 = (b, at) => (b[at] ?? 0) | ((b[at + 1] ?? 0) << 8);
const i16 = (b, at) => { const v = u16(b, at); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (b, at) => (u16(b, at) | (u16(b, at + 2) << 16)) >>> 0;
const f64 = (b, at) => (at + 8 <= b.length ? new DataView(b.buffer, b.byteOffset + at, 8).getFloat64(0, true) : 0);

const ERRORS = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A' };

// Excel's own colours by index, as BIFF8 has them before a workbook's PALETTE changes 8 to 63.
const PALETTE = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

// The record ids, which moved as the format grew.
const BOF = new Set([0x0009, 0x0209, 0x0409, 0x0809]);
const R = {
  EOF: 0x000a, CODEPAGE: 0x0042, DATEMODE: 0x0022, FILEPASS: 0x002f, PALETTE: 0x0092, BOUNDSHEET: 0x0085, WINDOW1: 0x003d,
  SST: 0x00fc, CONTINUE: 0x003c, EXTERNSHEET: 0x0017, SUPBOOK: 0x01ae, EXTERNNAME: 0x0023, EXTERNNAME34: 0x0223,
  FONT: 0x0031, FONT34: 0x0231, FONTCOLOR: 0x0045, FORMAT23: 0x001e, FORMAT: 0x041e,
  XF2: 0x0043, XF3: 0x0243, XF4: 0x0443, XF: 0x00e0, NAME: 0x0018, NAME34: 0x0218,
  BLANK2: 0x0001, INTEGER2: 0x0002, NUMBER2: 0x0003, LABEL2: 0x0004, BOOLERR2: 0x0005, FORMULA: 0x0006, STRING2: 0x0007,
  BLANK: 0x0201, NUMBER: 0x0203, LABEL: 0x0204, BOOLERR: 0x0205, FORMULA3: 0x0206, FORMULA4: 0x0406, STRING: 0x0207,
  RK: 0x027e, MULRK: 0x00bd, MULBLANK: 0x00be, LABELSST: 0x00fd, RSTRING: 0x00d6, IXFE: 0x0044,
  SHRFMLA: 0x04bc, ARRAY2: 0x0021, ARRAY: 0x0221,
  ROW2: 0x0008, ROW: 0x0208, COLINFO: 0x007d, COLWIDTH2: 0x0024, DEFCOLWIDTH: 0x0055, STANDARDWIDTH: 0x0099,
  DEFROWHEIGHT2: 0x0025, DEFROWHEIGHT: 0x0225, MERGEDCELLS: 0x00e5, WINDOW2_2: 0x003e, WINDOW2: 0x023e, PANE: 0x0041,
  BUNDLESHEET: 0x008f, MSODRAWINGGROUP: 0x00eb, MSODRAWING: 0x00ec, OBJ: 0x005d, TXO: 0x01b6, NOTE: 0x001c, HLINK: 0x01b8, HLINKTOOLTIP: 0x0800,
  IMGDATA: 0x007f, THEME: 0x0896, CONTINUEFRT: 0x0812, XFEXT: 0x087d, FEAT11: 0x0872, FEAT12: 0x0878, LIST12: 0x0877,
  CONDFMT: 0x01b0, CF: 0x01b1, CONDFMT12: 0x0879, CF12: 0x087a, CFEX: 0x087b,
};

/** Which BIFF a BOF record opens: 8, 5, 4, 3, 2 — or 0. */
function biffOf(id, data) {
  if (id === 0x0809) return u16(data, 0) === 0x0600 ? 8 : 5;
  if (id === 0x0409) return 4;
  if (id === 0x0209) return 3;
  if (id === 0x0009) return 2;
  return 0;
}

/** The record stream of a binary workbook, from a compound file or the file itself. */
function workbookStream(bytes) {
  if (!CompoundFile.is(bytes)) return bytes;
  const cfb = new CompoundFile(bytes);
  if (cfb.application() === 'encrypted') {
    const err = new XlsError('This workbook is password-protected.');
    err.encrypted = true;
    throw err;
  }
  const entry = cfb.find(['Workbook']) || cfb.find(['Book']) || cfb.find(['WORKBOOK']) || cfb.find(['BOOK']);
  if (!entry) throw new XlsError('not an Excel workbook (no Workbook stream)');
  return cfb.read(entry);
}

/** 'biff8', 'biff5', 'biff4', 'biff3', 'biff2' — or null for bytes that are none of them. */
export function xlsKind(bytes) {
  let stream;
  try { stream = workbookStream(bytes); } catch { return null; }
  if (!stream || stream.length < 8) return null;
  const v = BOF.has(u16(stream, 0)) ? biffOf(u16(stream, 0), stream.subarray(4)) : 0;
  return v ? 'biff' + v : null;
}

/** Each record: its id, its bytes, and where it starts in the stream. */
function records(stream) {
  const out = [];
  let at = 0;
  while (at + 4 <= stream.length) {
    const len = u16(stream, at + 2);
    out.push({ id: u16(stream, at), data: stream.subarray(at + 4, Math.min(stream.length, at + 4 + len)), pos: at });
    at += 4 + len;
  }
  return out;
}

/** A code page number, as a decoder's name. */
function codepageName(cp) {
  if (cp === 1200 || cp === 0x04b0) return 'utf-16le';
  if (cp === 10000 || cp === 0x8000) return 'macintosh';
  if (cp === 437 || cp === 850) return 'cp' + cp;
  if (cp >= 874 && cp <= 1258) return 'windows-' + cp;
  if (cp === 932) return 'shift_jis';
  if (cp === 936) return 'gbk';
  if (cp === 949) return 'euc-kr';
  if (cp === 950) return 'big5';
  return 'windows-1252';
}

/** Records with the CONTINUE records that carry their overflow: read as one, the segment edges known. */
class Segments {
  constructor(parts) {
    this.parts = parts;
    this.k = 0;
    this.at = 0;
  }
  get done() { return this.k >= this.parts.length; }
  _ensure() { while (this.k < this.parts.length && this.at >= this.parts[this.k].length) { this.k += 1; this.at = 0; } }
  u8() { this._ensure(); if (this.done) return 0; return this.parts[this.k][this.at++]; }
  u16() { return this.u8() | (this.u8() << 8); }
  u32() { return (this.u16() | (this.u16() << 16)) >>> 0; }
  skip(n) { for (let left = n; left > 0;) { this._ensure(); if (this.done) return; const take = Math.min(left, this.parts[this.k].length - this.at); this.at += take; left -= take; } }
  /** `cch` characters; where a CONTINUE splits them, the next part opens with its own flags byte. */
  chars(cch, high) {
    let s = '';
    let wide = high;
    let left = cch;
    while (left > 0) {
      if (this.k < this.parts.length && this.at >= this.parts[this.k].length) {
        this.k += 1;
        this.at = 0;
        if (this.done) break;
        wide = (this.parts[this.k][this.at++] & 1) === 1;
      }
      if (this.done) break;
      const part = this.parts[this.k];
      const avail = Math.floor((part.length - this.at) / (wide ? 2 : 1));
      const take = Math.min(left, Math.max(avail, 0));
      if (take === 0) { this.at = part.length; continue; }
      const codes = new Array(take);
      for (let i = 0; i < take; i++) codes[i] = wide ? part[this.at + i * 2] | (part[this.at + i * 2 + 1] << 8) : part[this.at + i];
      s += String.fromCharCode(...codes);
      this.at += take * (wide ? 2 : 1);
      left -= take;
    }
    return s;
  }
}

/* ── the workbook ─────────────────────────────────────────────────────── */

/**
 * Read a binary workbook: { biff, date1904, fonts, formats, xfs, palette,
 * sheets: [{ name, state, cells, rows, cols, merges, frozen, grid, ... }],
 * names, activeSheet }.
 */
export function readXls(bytes) {
  const stream = workbookStream(bytes);
  const recs = records(stream);
  if (!recs.length || !BOF.has(recs[0].id)) throw new XlsError('not an Excel workbook (it does not begin as one)');
  const biff = biffOf(recs[0].id, recs[0].data);
  const book = {
    biff, date1904: false, fonts: [], formats: new Map(), xfs: [], palette: PALETTE.slice(), sheets: [], names: [], activeSheet: 0,
    codepage: biff === 8 ? 'utf-16le' : 'windows-1252', externSheets: [], supbooks: [], biff5ExternNames: [],
  };
  const decode = (b) => decoderFor(book.codepage === 'utf-16le' ? 'windows-1252' : book.codepage).decode(b);

  // Strings, as each version writes them.
  /** BIFF8's Unicode string: a count, a flags byte, the characters (and any rich-text runs or extras skipped). */
  const uni = (b, at, countBytes = 2) => {
    const cch = countBytes === 2 ? u16(b, at) : u8(b, at);
    let p = at + countBytes;
    const flags = u8(b, p++);
    const high = flags & 1;
    let runs = 0;
    let ext = 0;
    if (flags & 8) { runs = u16(b, p); p += 2; }
    if (flags & 4) { ext = u32(b, p); p += 4; }
    let s = '';
    for (let i = 0; i < cch; i++) s += String.fromCharCode(high ? u16(b, p + i * 2) : u8(b, p + i));
    p += cch * (high ? 2 : 1) + runs * 4 + ext;
    return { text: s, end: p };
  };
  /** Before BIFF8: a count and the bytes in the workbook's code page. */
  const bytestr = (b, at, countBytes = 1) => {
    const cch = countBytes === 2 ? u16(b, at) : u8(b, at);
    return { text: decode(b.subarray(at + countBytes, at + countBytes + cch)), end: at + countBytes + cch };
  };
  const str = (b, at, countBytes) => (biff === 8 ? uni(b, at, countBytes) : bytestr(b, at, countBytes));

  let sst = [];
  let sstRuns = [];
  let depth = 0;
  let sheet = null;
  // The depth of a substream being passed over — a chart, a macro sheet, a
  // chart inside a sheet — whose records are not this workbook's: a chart
  // has fonts of its own, which would renumber the cells' fonts.
  let skipFrom = 0;
  const sheetAt = new Map();
  const bundle = [];
  let lastFont = null;
  let nextFormat = 0;
  let lastFormula = null;
  let ixfe = null;
  const newSheet = (name, state = 'visible') => {
    const s = { name, state, cells: new Map(), rows: new Map(), cols: [], merges: [], frozen: null, grid: true, headings: true, formulas: false, selected: false, defaultColWidth: null, defaultRowHeight: null, shared: new Map(), arrays: new Map(), pending: [], drawing: [], objectText: new Map(), notes: [], links: [], pictures: [], charts: [], shapes: [], tables: [], condFormats: [], lastObject: null };
    book.sheets.push(s);
    return s;
  };
  const put = (row, col, cell) => { if (sheet) sheet.cells.set(row * 0x4000 + col, { row, col, ...cell }); };
  const cellXf = (data, at) => (biff === 2 ? (ixfe != null ? ixfe : u8(data, at) & 0x3f) : u16(data, at));
  // BIFF2's three attribute bytes: the XF, then font and format bits the XF also has.
  const xfAt = (data) => { const xf = cellXf(data, 4); ixfe = null; return xf; };
  const valueStart = biff === 2 ? 7 : 6;

  // A chart's records, kept as they come — a chart inside a sheet, after the
  // object that places it, or a chart sheet — and read once every sheet is known.
  let chartFrom = 0;
  let chartRecs = null;
  for (let k = 0; k < recs.length; k++) {
    const { id, data, pos } = recs[k];
    if (chartRecs) {
      chartRecs.push(recs[k]);
      if (BOF.has(id)) depth += 1;
      if (id !== R.EOF) continue;
      if (depth === chartFrom) {
        sheet.charts.push({ records: chartRecs, object: chartFrom > 1 ? sheet.lastObject : null });
        if (chartFrom > 1) sheet.lastObject = null;
        chartRecs = null;
        chartFrom = 0;
      }
      depth = Math.max(0, depth - 1);
      if (depth === 0 && sheet) { finishSheet(sheet, book, biff, decode); sheet = null; }
      continue;
    }
    if (BOF.has(id)) {
      depth += 1;
      if (skipFrom) continue;
      const dt = u16(data, 2);
      if (depth > 1) {
        // A chart inside a sheet; anything else nested is passed over.
        if (sheet && dt === 0x0020 && biff >= 5) { chartFrom = depth; chartRecs = [recs[k]]; } else skipFrom = depth;
        continue;
      }
      if (k === 0 && (dt === 0x0005 || dt === 0x0100)) { sheet = null; continue; } // the workbook's own records
      if (dt === 0x0010 || (biff <= 4 && k === 0)) {
        sheet = sheetAt.get(pos) || bundle.shift() || newSheet('Sheet' + (book.sheets.length + 1));
      } else if (dt === 0x0020 && biff >= 5 && sheetAt.get(pos)) {
        // A chart sheet: a sheet holding its chart.
        sheet = sheetAt.get(pos);
        sheet.chartSheet = true;
        chartFrom = depth;
        chartRecs = [recs[k]];
      } else {
        // A macro sheet or a module: not a grid this suite draws.
        sheet = null;
        skipFrom = depth;
        const s = sheetAt.get(pos);
        if (s) s.skip = true;
      }
      continue;
    }
    if (id === R.EOF) {
      if (skipFrom && depth === skipFrom) skipFrom = 0;
      depth = Math.max(0, depth - 1);
      if (depth === 0 && !skipFrom) {
        if (sheet) finishSheet(sheet, book, biff, decode);
        sheet = null;
      }
      continue;
    }
    if (skipFrom) continue;

    switch (id) {
      case R.FILEPASS: {
        const err = new XlsError('This workbook is password-protected.');
        err.encrypted = true;
        throw err;
      }
      case R.CODEPAGE: book.codepage = codepageName(u16(data, 0)); break;
      case R.DATEMODE: book.date1904 = u16(data, 0) === 1; break;
      case R.PALETTE: {
        const n = u16(data, 0);
        for (let i = 0; i < n && 8 + i < 64; i++) book.palette[8 + i] = [0, 1, 2].map((c) => u8(data, 2 + i * 4 + c).toString(16).padStart(2, '0')).join('').toUpperCase();
        break;
      }
      case R.BOUNDSHEET: {
        const at = u32(data, 0);
        const state = ['visible', 'hidden', 'veryHidden'][u8(data, 4) & 3] || 'visible';
        const name = biff === 8 ? uni(data, 6, 1).text : bytestr(data, 6, 1).text;
        const s = { name, state };
        // A macro sheet or a module keeps its place but holds no grid; a chart sheet holds its chart.
        const kind = u8(data, 5) & 0x0f;
        if (kind !== 0 && kind !== 2) s.skip = true;
        sheetAt.set(at, Object.assign(newSheet(name, state), s));
        break;
      }
      case R.BUNDLESHEET: bundle.push(newSheet(bytestr(data, 6, 1).text)); break;
      case R.WINDOW1: if (biff >= 5) book.activeSheet = u16(data, 10); break;
      case R.SST: {
        const parts = [data.subarray(8)];
        while (recs[k + 1]?.id === R.CONTINUE) parts.push(recs[++k].data);
        const seg = new Segments(parts);
        const count = u32(data, 4);
        sst = [];
        sstRuns = [];
        for (let i = 0; i < count && !seg.done; i++) {
          const cch = seg.u16();
          const flags = seg.u8();
          const runs = flags & 8 ? seg.u16() : 0;
          const ext = flags & 4 ? seg.u32() : 0;
          sst.push(seg.chars(cch, (flags & 1) === 1));
          // Rich text: where each run starts, and its font.
          if (runs) {
            const list = [];
            for (let r = 0; r < runs; r++) list.push({ at: seg.u16(), font: seg.u16() });
            sstRuns[i] = list;
          }
          seg.skip(ext);
        }
        break;
      }
      case R.EXTERNSHEET:
        if (biff === 8) {
          const n = u16(data, 0);
          for (let i = 0; i < n; i++) book.externSheets.push({ supbook: u16(data, 2 + i * 6), first: u16(data, 4 + i * 6), last: u16(data, 6 + i * 6) });
        } else {
          // BIFF5: one sheet a record; a code 3 or 4 says it is in this workbook.
          const cch = u8(data, 0);
          const kind = u8(data, 1);
          book.externSheets.push({ internal: kind === 3 || kind === 4, name: decode(data.subarray(2, 1 + cch)), names: [] });
        }
        break;
      case R.SUPBOOK: {
        const cch = u16(data, 2);
        book.supbooks.push({ internal: cch === 0x0401, addin: cch === 0x3a01, names: [] });
        break;
      }
      case R.EXTERNNAME: case R.EXTERNNAME34: {
        const name = biff === 8 ? uni(data, 6, 1).text : bytestr(data, biff === 5 ? 6 : 2, 1).text;
        if (biff === 8) book.supbooks[book.supbooks.length - 1]?.names.push(name);
        else (book.externSheets[book.externSheets.length - 1]?.names || book.biff5ExternNames).push(name);
        break;
      }
      case R.FONT: case R.FONT34: {
        const f = { height: u16(data, 0), bold: false, italic: false, underline: 0, strike: false, colour: null, name: '', family: 0, charset: 0, script: 0 };
        const flags = u16(data, 2);
        f.italic = Boolean(flags & 0x02);
        f.strike = Boolean(flags & 0x08);
        if (biff >= 5) {
          f.colour = u16(data, 4);
          f.bold = u16(data, 6) >= 600;
          f.script = u16(data, 8);
          f.underline = u8(data, 10);
          f.family = u8(data, 11);
          f.charset = u8(data, 12);
          f.name = biff === 8 ? uni(data, 14, 1).text : bytestr(data, 14, 1).text;
        } else {
          f.bold = Boolean(flags & 0x01);
          f.underline = flags & 0x04 ? 1 : 0;
          if (id === R.FONT34) { f.colour = u16(data, 4); f.name = bytestr(data, 6, 1).text; } else f.name = bytestr(data, 4, 1).text;
        }
        book.fonts.push(f);
        lastFont = f;
        break;
      }
      case R.FONTCOLOR: if (lastFont) lastFont.colour = u16(data, 0); break;
      case R.FORMAT23: book.formats.set(nextFormat++, bytestr(data, 0, 1).text); break;
      case R.FORMAT:
        if (biff === 8) book.formats.set(u16(data, 0), uni(data, 2, 2).text);
        else if (biff === 5) book.formats.set(u16(data, 0), bytestr(data, 2, 1).text);
        else book.formats.set(nextFormat++, bytestr(data, 2, 1).text); // BIFF4: the index field is unused
        break;
      case R.XF2: case R.XF3: case R.XF4: case R.XF: book.xfs.push(readXf(data, biff)); break;
      case R.NAME: case R.NAME34: book.names.push(readName(data, biff, str, decode)); break;

      // The cells.
      case R.IXFE: ixfe = u16(data, 0); break;
      case R.NUMBER2: case R.NUMBER: put(u16(data, 0), u16(data, 2), { xf: xfAt(data), t: 'n', v: f64(data, valueStart) }); break;
      case R.INTEGER2: put(u16(data, 0), u16(data, 2), { xf: xfAt(data), t: 'n', v: u16(data, 7) }); break;
      case R.RK: put(u16(data, 0), u16(data, 2), { xf: u16(data, 4), t: 'n', v: rk(u32(data, 6)) }); break;
      case R.MULRK: {
        const row = u16(data, 0);
        const first = u16(data, 2);
        const n = Math.floor((data.length - 6) / 6);
        for (let i = 0; i < n; i++) put(row, first + i, { xf: u16(data, 4 + i * 6), t: 'n', v: rk(u32(data, 6 + i * 6)) });
        break;
      }
      case R.BLANK2: case R.BLANK: put(u16(data, 0), u16(data, 2), { xf: xfAt(data), t: null }); break;
      case R.MULBLANK: {
        const row = u16(data, 0);
        const first = u16(data, 2);
        const n = Math.floor((data.length - 6) / 2);
        for (let i = 0; i < n; i++) put(row, first + i, { xf: u16(data, 4 + i * 2), t: null });
        break;
      }
      case R.LABELSST: {
        const at = u32(data, 6);
        put(u16(data, 0), u16(data, 2), { xf: u16(data, 4), t: 's', v: sst[at] ?? '', ...(sstRuns[at] ? { runs: sstRuns[at] } : {}) });
        break;
      }
      case R.LABEL2: put(u16(data, 0), u16(data, 2), { xf: xfAt(data), t: 's', v: bytestr(data, 7, 1).text }); break;
      case R.LABEL: case R.RSTRING: {
        const read = str(data, 6, 2);
        const cell = { xf: u16(data, 4), t: 's', v: read.text };
        // A rich string's runs after its words: BIFF8's four bytes each, before it two.
        if (id === R.RSTRING && read.end < data.length) {
          const n = biff === 8 ? u16(data, read.end) : u8(data, read.end);
          const runs = [];
          for (let r = 0, p = read.end + (biff === 8 ? 2 : 1); r < n && p < data.length; r++) {
            runs.push(biff === 8 ? { at: u16(data, p), font: u16(data, p + 2) } : { at: u8(data, p), font: u8(data, p + 1) });
            p += biff === 8 ? 4 : 2;
          }
          if (runs.length) cell.runs = runs;
        }
        put(u16(data, 0), u16(data, 2), cell);
        break;
      }
      case R.BOOLERR2: case R.BOOLERR: {
        const xf = xfAt(data);
        const v = u8(data, valueStart);
        put(u16(data, 0), u16(data, 2), u8(data, valueStart + 1) ? { xf, t: 'e', v: ERRORS[v] || '#N/A' } : { xf, t: 'b', v: v ? 1 : 0 });
        break;
      }
      case R.FORMULA: case R.FORMULA3: case R.FORMULA4: {
        if (!sheet) break;
        const row = u16(data, 0);
        const col = u16(data, 2);
        const xf = xfAt(data);
        const at = biff === 2 ? 7 : 6;
        const cached = formulaResult(data, at);
        // After the value: options (a byte in BIFF2), BIFF5 and 8's four unused bytes, the token count.
        const ceAt = biff === 2 ? 17 : biff >= 5 ? 22 : 18;
        const cce = biff === 2 ? u8(data, 16) : u16(data, ceAt - 2);
        const rgce = data.subarray(ceAt, ceAt + cce);
        const cell = { xf, ...cached, formula: null };
        put(row, col, cell);
        sheet.pending.push({ cell: sheet.cells.get(row * 0x4000 + col), rgce, extra: data.subarray(ceAt + cce), row, col });
        lastFormula = cached.t === 'str' ? sheet.cells.get(row * 0x4000 + col) : null;
        break;
      }
      case R.STRING2: case R.STRING:
        if (lastFormula) { lastFormula.v = biff === 8 ? uni(data, 0, 2).text : bytestr(data, 0, biff === 2 ? 1 : 2).text; lastFormula = null; }
        break;
      case R.SHRFMLA: {
        if (!sheet) break;
        const cce = u16(data, 8);
        sheet.shared.set(u16(data, 0) * 0x4000 + u8(data, 4), { rgce: data.subarray(10, 10 + cce), extra: data.subarray(10 + cce) });
        break;
      }
      case R.ARRAY2: case R.ARRAY: {
        if (!sheet) break;
        const first = { row: u16(data, 0), col: u8(data, 4) };
        const range = { top: first.row, bottom: u16(data, 2), left: first.col, right: u8(data, 5) };
        const ceAt = biff === 2 ? 8 : biff <= 4 ? 10 : 14;
        const cce = biff === 2 ? u8(data, 7) : u16(data, ceAt - 2);
        sheet.arrays.set(first.row * 0x4000 + first.col, { range, rgce: data.subarray(ceAt, ceAt + cce), extra: data.subarray(ceAt + cce) });
        break;
      }

      // Rows and columns.
      case R.ROW2: case R.ROW: {
        if (!sheet) break;
        const raw = u16(data, 6);
        const flags = biff >= 3 ? u16(data, 12) : 0;
        const entry = {};
        // A height of its own: BIFF5 and 8 flag it; before them the top bit says "the default".
        const custom = biff >= 5 ? Boolean(flags & 0x40) : !(raw & 0x8000);
        if (custom && (raw & 0x7fff)) entry.height = raw & 0x7fff;
        // One sized from its fonts: the height Excel gave it, which its drawings were placed against.
        else if (biff >= 5 && (raw & 0x7fff) && (raw & 0x7fff) !== (sheet.defaultRowHeight ?? 255)) entry.fitted = raw & 0x7fff;
        if (flags & 0x20) entry.hidden = true;
        if (flags & 0x80 && biff >= 5) entry.xf = u16(data, 14) & 0x0fff;
        if (flags & 0x0007) entry.level = flags & 7;
        if (Object.keys(entry).length) sheet.rows.set(u16(data, 0), entry);
        break;
      }
      case R.COLINFO: if (sheet) sheet.cols.push({ first: u16(data, 0), last: u16(data, 2), width: u16(data, 4), xf: u16(data, 6), hidden: Boolean(u16(data, 8) & 1), level: (u16(data, 8) >> 8) & 7 }); break;
      case R.COLWIDTH2: if (sheet) sheet.cols.push({ first: u8(data, 0), last: u8(data, 1), width: u16(data, 2), xf: 0, hidden: false, level: 0 }); break;
      case R.DEFCOLWIDTH: if (sheet) sheet.defaultColWidth = u16(data, 0); break;
      case R.STANDARDWIDTH: if (sheet) sheet.standardWidth = u16(data, 0); break;
      case R.DEFROWHEIGHT: if (sheet) sheet.defaultRowHeight = u16(data, 2); break;
      case R.DEFROWHEIGHT2: if (sheet) sheet.defaultRowHeight = u16(data, 0) & 0x7fff; break;
      case R.MERGEDCELLS: {
        if (!sheet) break;
        const n = u16(data, 0);
        for (let i = 0; i < n; i++) sheet.merges.push({ top: u16(data, 2 + i * 8), bottom: u16(data, 4 + i * 8), left: u16(data, 6 + i * 8), right: u16(data, 8 + i * 8) });
        break;
      }
      case R.WINDOW2_2:
        if (sheet) { sheet.formulas = Boolean(u8(data, 0)); sheet.grid = Boolean(u8(data, 1)); sheet.headings = Boolean(u8(data, 2)); sheet.frozenFlag = Boolean(u8(data, 3)); }
        break;
      case R.WINDOW2:
        if (sheet) {
          const g = u16(data, 0);
          sheet.formulas = Boolean(g & 0x01);
          sheet.grid = Boolean(g & 0x02);
          sheet.headings = Boolean(g & 0x04);
          sheet.frozenFlag = Boolean(g & 0x08);
          sheet.selected = Boolean(g & 0x200);
        }
        break;
      case R.PANE:
        if (sheet && sheet.frozenFlag) {
          const cols = u16(data, 0);
          const rows = u16(data, 2);
          if (rows || cols) sheet.frozen = { rows, cols };
        }
        break;

      // Drawings: the workbook's pictures (its drawing group), each sheet's shapes, and the objects among them.
      case R.MSODRAWINGGROUP: {
        const parts = [data];
        while (recs[k + 1]?.id === R.CONTINUE) parts.push(recs[++k].data);
        book.drawingGroup = concat(parts);
        break;
      }
      case R.MSODRAWING:
        if (sheet) {
          sheet.drawing.push(data);
          while (recs[k + 1]?.id === R.CONTINUE) sheet.drawing.push(recs[++k].data);
        }
        break;
      case R.OBJ:
        if (!sheet) break;
        // BIFF8's object follows its shape in the drawing: where the drawing had got to says which shape it is.
        if (biff === 8) {
          sheet.lastObject = u16(data, 0) === 0x15 ? { type: u16(data, 4), id: u16(data, 6), at: sheet.drawing.reduce((n, p) => n + p.length, 0) } : null;
          if (sheet.lastObject) (sheet.objects ??= []).push(sheet.lastObject);
          // A large drawing goes on in the CONTINUEs after an object.
          while (recs[k + 1]?.id === R.CONTINUE) sheet.drawing.push(recs[++k].data);
          break;
        }
        // Before BIFF8 each object is one record: its type, flags and the cells
        // its corners are in (a picture's bytes in the IMGDATA after it), and from
        // BIFF5 a picture's name — "__BkgndObj", hidden, is the sheet's background.
        {
          let name = '';
          if (biff === 5 && u16(data, 4) === 8 && u16(data, 30) && data.length > 60) name = decode(data.subarray(61, 61 + u8(data, 60)));
          sheet.lastObject = {
            type: u16(data, 4), id: u16(data, 6), hidden: Boolean(u16(data, 8) & 0x0100), name,
            from: { col: u16(data, 10), dx: u16(data, 12) / 1024, row: u16(data, 14), dy: u16(data, 16) / 256 },
            to: { col: u16(data, 18), dx: u16(data, 20) / 1024, row: u16(data, 22), dy: u16(data, 24) / 256 },
          };
        }
        break;
      case R.IMGDATA: {
        const parts = [data];
        while (recs[k + 1]?.id === R.CONTINUE) parts.push(recs[++k].data);
        const obj = sheet?.lastObject;
        if (!sheet || biff === 8 || !obj || obj.type !== 8 || (obj.hidden && obj.name === '__BkgndObj')) break;
        sheet.lastObject = null;
        const blip = imageData(concat(parts), biff);
        if (blip) sheet.pictures.push({ blip, name: obj.name || null, from: obj.from, to: obj.to, order: obj.id });
        break;
      }
      case R.TXO: {
        // A text box's or a note's words, in the CONTINUEs after it — each a
        // flags byte, then characters — and then its formatting runs. Only as
        // many CONTINUEs as those take are its own: a large drawing goes on in
        // the CONTINUEs after them.
        if (!sheet) break;
        const cch = u16(data, 10);
        const cbRuns = biff === 8 ? u16(data, 12) : 0;
        const parts = [];
        let chars = 0;
        while (chars < cch && recs[k + 1]?.id === R.CONTINUE) {
          const part = recs[++k].data;
          parts.push(part);
          chars += biff === 8 ? Math.floor((part.length - 1) / (part[0] & 1 ? 2 : 1)) : part.length;
        }
        const runParts = [];
        let runBytes = 0;
        while (runBytes < cbRuns && recs[k + 1]?.id === R.CONTINUE) { const part = recs[++k].data; runParts.push(part); runBytes += part.length; }
        let text = '';
        if (cch && parts.length) {
          if (biff === 8) {
            const seg = new Segments(parts);
            const high = (seg.u8() & 1) === 1;
            text = seg.chars(cch, high);
          } else text = decode(concat(parts).subarray(0, cch));
        }
        const runData = concat(runParts);
        const runs = [];
        for (let p = 0; p + 4 <= runData.length; p += 8) if (u16(runData, p) < cch) runs.push({ at: u16(runData, p), font: u16(runData, p + 2) });
        const grbit = u16(data, 0);
        if (sheet.lastObject) sheet.objectText.set(sheet.lastObject.id, { text, runs, h: (grbit >> 1) & 7, v: (grbit >> 4) & 7, rotation: u16(data, 2) });
        if (biff === 8) while (recs[k + 1]?.id === R.CONTINUE) sheet.drawing.push(recs[++k].data);
        break;
      }
      case R.NOTE:
        if (!sheet) break;
        if (biff === 8) sheet.notes.push({ row: u16(data, 0), col: u16(data, 2), object: u16(data, 6), author: uni(data, 8, 2).text });
        else if (u16(data, 0) === 0xffff && sheet.notes.length) sheet.notes[sheet.notes.length - 1].text += decode(data.subarray(6, 6 + u16(data, 4)));
        else sheet.notes.push({ row: u16(data, 0), col: u16(data, 2), author: '', text: decode(data.subarray(6, 6 + u16(data, 4))) });
        break;
      // Conditional formatting: a block of ranges, then its rules — Excel 97's (CF) or Excel 2007's (CF12).
      case R.CONDFMT: case R.CONDFMT12: {
        if (!sheet || biff !== 8) break;
        const at = id === R.CONDFMT12 ? 12 : 0;
        const ranges = [];
        for (let i = 0, n = u16(data, at + 12); i < n; i++) {
          const p = at + 14 + i * 8;
          ranges.push({ top: u16(data, p), bottom: u16(data, p + 2), left: u16(data, p + 4), right: u16(data, p + 6) });
        }
        if (ranges.length) sheet.condFormats.push({ id: u16(data, at + 2) >> 1, ranges, rules: [] });
        break;
      }
      case R.CF: case R.CF12: {
        const block = sheet?.condFormats[sheet.condFormats.length - 1];
        if (!block) break;
        try { block.rules.push(readCondition(data, id === R.CF12)); } catch { /* a rule this cannot read is left out */ }
        break;
      }
      case R.CFEX: {
        // An Excel 97-style rule's priority, and the exact colours Excel 2007 kept for its formatting after a copy of it.
        if (!sheet || u32(data, 12) !== 0) break;
        const rule = sheet.condFormats.find((b) => b.id === u16(data, 16))?.rules[u16(data, 18)];
        if (!rule) break;
        rule.priority = u16(data, 22);
        if (!u8(data, 25) || !rule.dxf || data.length < 36) break;
        const end = 30 + u32(data, 26);
        let p = readDxf(data, 30).end;
        if (p + 8 > end) break;
        const count = u16(data, p + 6);
        p += 8;
        for (let i = 0; i < count && p + 4 <= end; i++) {
          const key = { 4: 'fillFg', 5: 'fillBg', 13: 'text' }[u16(data, p)];
          if (key) (rule.dxf.exact ??= {})[key] = fullColour(data, p + 4);
          p += Math.max(4, u16(data, p + 2));
        }
        break;
      }
      case R.FEAT11: case R.FEAT12: {
        // A table (ListObject): its range, and how many header and totals rows it has.
        if (!sheet || u16(data, 12) !== 5) break;
        const cref = u16(data, 19);
        const at = 27 + cref * 8;
        if (at + 16 > data.length) break;
        sheet.tables.push({
          id: u32(data, at + 4), header: u32(data, at + 8), totals: u32(data, at + 12),
          ref: { top: u16(data, 4), bottom: u16(data, 6), left: u16(data, 8), right: u16(data, 10) },
          name: null, style: null, stripes: true, firstColumn: false, lastColumn: false, columnStripes: false,
        });
        break;
      }
      case R.LIST12: {
        // A table's style (and which of its stripes and columns it shows), or its name.
        const table = sheet?.tables.find((t) => t.id === u32(data, 14));
        if (!table) break;
        const lsd = u16(data, 12);
        if (lsd === 1 && data.length >= 23) {
          const flags = u16(data, 18);
          Object.assign(table, { firstColumn: Boolean(flags & 1), lastColumn: Boolean(flags & 2), stripes: Boolean(flags & 4), columnStripes: Boolean(flags & 8) });
          const cch = u16(data, 20);
          const high = u8(data, 22) & 1;
          let name = '';
          for (let i = 0; i < cch; i++) name += String.fromCharCode(high ? u16(data, 23 + i * 2) : u8(data, 23 + i));
          if (name) table.style = name;
        } else if (lsd === 2 && data.length >= 21) table.name = uni(data, 18, 2).text || null;
        break;
      }
      case R.XFEXT: {
        // Excel 2007's exact colours for a cell format, beside the palette's nearest in its XF: fill, borders, text.
        const ixfe = u16(data, 14);
        const count = u16(data, 18);
        const ext = {};
        let p = 20;
        for (let i = 0; i < count && p + 4 <= data.length; i++) {
          const type = u16(data, p);
          const cb = u16(data, p + 2);
          const key = { 4: 'fillFg', 5: 'fillBg', 7: 'top', 8: 'bottom', 9: 'left', 10: 'right', 13: 'text' }[type];
          if (key && cb >= 12) ext[key] = fullColour(data, p + 4);
          p += Math.max(4, cb);
        }
        (book.xfExt ??= new Map()).set(ixfe, ext);
        break;
      }
      case R.THEME: {
        // Excel 2007's theme, zipped, after a future record's header and the theme's version.
        const parts = [data.subarray(16)];
        while (recs[k + 1]?.id === R.CONTINUEFRT) parts.push(recs[++k].data.subarray(12));
        book.theme = concat(parts);
        break;
      }
      case R.HLINK: if (sheet) { const link = readHyperlink(data); if (link) sheet.links.push(link); } break;
      case R.HLINKTOOLTIP: {
        if (!sheet || !sheet.links.length) break;
        let tip = '';
        for (let p = 10; p + 1 < data.length; p += 2) { const c = u16(data, p); if (!c) break; tip += String.fromCharCode(c); }
        sheet.links[sheet.links.length - 1].tooltip = tip;
        break;
      }
      default: break;
    }
  }
  if (sheet) finishSheet(sheet, book, biff, decode);

  // The names' formulas, now that every sheet is known by its number.
  if (!book.allNames) book.allNames = book.sheets.map((s) => s.name);
  for (const n of book.names) {
    try {
      n.formula = n.rgce?.length ? decompile(n.rgce, n.extra, { book, biff, row: 0, col: 0, decode, sheetNames: book.allNames }) : '';
    } catch {
      n.formula = '';
    }
    delete n.rgce;
    delete n.extra;
  }
  // The charts, now that every sheet is known by its number: each read from its records.
  const chartCtx = {
    biff, palette: book.palette,
    text: (b, at, countBytes) => str(b, at, countBytes).text,
    formula: (rgce) => decompile(rgce, new Uint8Array(0), { book, biff, row: 0, col: 0, decode, sheetNames: book.allNames }),
    font: (index) => {
      const f = book.fonts[index < 4 ? index : index - 1];
      return f ? { height: f.height, bold: f.bold, italic: f.italic, name: f.name, colour: f.colour != null && f.colour < 64 && f.colour !== 8 ? book.palette[f.colour] : null } : null;
    },
  };
  for (const s of book.sheets) {
    for (const block of s.condFormats) {
      const top = block.ranges[0];
      const read = (rgce) => { try { return rgce?.length ? decompile(rgce, new Uint8Array(0), { book, biff, row: top.top, col: top.left, decode, sheetNames: book.allNames }) : null; } catch { return null; } };
      for (const rule of block.rules) {
        rule.formulas = (rule.rgce || []).map(read).filter((x) => x != null);
        for (const t of rule.thresholds || []) if (t.rgce) t.formula = read(t.rgce);
        delete rule.rgce;
        for (const t of rule.thresholds || []) delete t.rgce;
      }
    }
    for (const c of s.charts) {
      try { c.chart = readChart(c.records, chartCtx); } catch { c.chart = null; }
      delete c.records;
      delete c.object;
    }
    s.charts = s.charts.filter((c) => c.chart);
  }
  // Modules and macro sheets leave; what pointed at sheets by number follows them.
  const all = book.sheets;
  const active = all[book.activeSheet] && !all[book.activeSheet].skip ? all[book.activeSheet] : null;
  book.sheets = all.filter((s) => !s.skip);
  for (const n of book.names) if (n.sheet != null) n.sheet = all[n.sheet] && !all[n.sheet].skip ? book.sheets.indexOf(all[n.sheet]) : null;
  // Excel's hidden "_xlfn." names only mark functions an older Excel lacked; a
  // built-in name written out in full (BIFF5 writes them so) is the built-in one.
  for (const n of book.names) if (BUILTIN_NAMES.includes('_xlnm.' + n.name)) n.name = '_xlnm.' + n.name;
  book.names = book.names.filter((n) => n.formula && !n.name.startsWith('_xlfn.') && !/^_xlnm\.(Auto_|Recorder|Data_Form)/.test(n.name));
  for (const s of book.sheets) { delete s.shared; delete s.pending; delete s.arrays; delete s.drawing; delete s.objectText; delete s.lastObject; delete s.objects; }
  if (!book.sheets.length) book.sheets.push(newSheet('Sheet1'));
  book.activeSheet = Math.max(0, active ? book.sheets.indexOf(active) : book.sheets.findIndex((s) => s.selected));
  delete book.allNames;
  return book;
}

/** A sheet's formulas, now that its shared and array formulas are known: each read into text. */
function finishSheet(sheet, book, biff, decode) {
  if (!book.allNames) book.allNames = book.sheets.map((s) => s.name);
  const sheetNames = book.allNames;
  for (const p of sheet.pending) {
    let { rgce, extra } = p;
    let base = { row: p.row, col: p.col };
    // A formula that is only a pointer: into a block sharing one formula, or an array formula.
    if (rgce[0] === 0x01 || rgce[0] === 0x02) {
      const row = u16(rgce, 1);
      const col = biff === 8 ? u16(rgce, 3) : u8(rgce, 3);
      const key = row * 0x4000 + col;
      const shared = sheet.shared.get(key);
      const array = sheet.arrays.get(key);
      if (array) {
        if (p.row === array.range.top && p.col === array.range.left) {
          p.cell.array = array.range;
          rgce = array.rgce;
          extra = array.extra;
        } else { delete p.cell.formula; continue; } // the rest of an array keeps its value
      } else if (shared) {
        rgce = shared.rgce;
        extra = shared.extra;
      } else { delete p.cell.formula; continue; }
    }
    try {
      p.cell.formula = decompile(rgce, extra, { book, biff, row: base.row, col: base.col, decode, sheetNames });
    } catch {
      delete p.cell.formula;
    }
  }
  sheet.pending = [];
  try { finishDrawings(sheet, book); } catch { /* a drawing that cannot be read leaves the cells */ }
  sheet.arrays = new Map();
}

/**
 * A picture in an IMGDATA record, as Excel 3 to 95 kept one: a format
 * (2 a Windows metafile after an 8-byte header, 9 a DIB), an environment and
 * a length, then the bytes. Excel 3 and 4 write a 32-bit DIB with three
 * stray bytes after its header; Excel 5 cannot read those either.
 */
function imageData(d, biff) {
  const format = u16(d, 0);
  const size = u32(d, 4);
  let body = d.subarray(8, 8 + size);
  if (format === 2 && body.length > 8) return { contentType: 'image/x-wmf', ext: 'wmf', bytes: placeableWmf(body.subarray(8)) };
  if (format !== 9 || body.length < 12) return null;
  if (biff <= 4 && u32(body, 0) === 12 && u16(body, 8) === 1 && u16(body, 10) === 32) body = concat([body.subarray(0, 12), body.subarray(15)]);
  return { contentType: 'image/bmp', ext: 'bmp', bytes: dibFile(body) };
}

/**
 * One conditional-formatting rule. Excel 97's (CF): a comparison of the
 * cell's value, or a formula, and the formatting it sets — font colour,
 * bold and italic, fill. Excel 2007's (CF12) add colour scales, data
 * bars and icon sets, each with its thresholds (Apache POI's reading of
 * the record) and its priority.
 */
function readCondition(d, v12) {
  const f64at = (at) => (at + 8 <= d.length ? new DataView(d.buffer, d.byteOffset + at, 8).getFloat64(0, true) : 0);
  let q = v12 ? 12 : 0;
  const ct = u8(d, q);
  const cp = u8(d, q + 1);
  const cce1 = u16(d, q + 2);
  const cce2 = u16(d, q + 4);
  q += 6;
  let dxf = null;
  if (v12) {
    const cb = u32(d, q);
    q += 4;
    if (cb) { dxf = readDxf(d, q); q += cb; } else q += 2;
  } else {
    dxf = readDxf(d, q);
    q = dxf.end;
  }
  const rgce = [d.subarray(q, q + cce1), d.subarray(q + cce1, q + cce1 + cce2)].filter((x) => x.length);
  q += cce1 + cce2;
  const rule = { type: ct === 2 ? 'expression' : 'cellIs', operator: OPERATORS[cp] || null, rgce, dxf, priority: null };
  if (!v12) return rule;
  q += 2 + u16(d, q); // the scale's formula
  q += 1; // its options
  rule.priority = u16(d, q);
  q += 4;
  q += 1 + u8(d, q); // the template's parameters
  const colour = () => {
    const type = u32(d, q);
    const value = u32(d, q + 4);
    const rgb = [0, 1, 2].map((c) => u8(d, q + 4 + c).toString(16).padStart(2, '0')).join('').toUpperCase();
    const tint = f64at(q + 8);
    q += 16;
    return type === 2 ? { rgb, tint } : type === 3 ? { theme: value, tint } : type === 1 ? { indexed: value, tint } : null;
  };
  const threshold = (extra) => {
    const type = u8(d, q);
    const cce = u16(d, q + 1);
    q += 3;
    const t = { type: CFVO[type] || 'num', value: null, rgce: cce ? d.subarray(q, q + cce) : null };
    q += cce;
    if (!cce && type !== 2 && type !== 3) { t.value = f64at(q); q += 8; }
    return Object.assign(t, extra ? extra() : {});
  };
  if (ct === 3) {
    q += 3;
    const n = u8(d, q);
    q += 3;
    const thresholds = [];
    for (let i = 0; i < n; i++) thresholds.push(threshold(() => { q += 8; return {}; }));
    const colours = [];
    for (let i = 0; i < n; i++) { q += 8; colours.push(colour()); }
    return { type: 'colorScale', priority: rule.priority, thresholds, colours };
  }
  if (ct === 4) {
    q += 3;
    const options = u8(d, q);
    const min = u8(d, q + 1);
    const max = u8(d, q + 2);
    q += 3;
    const c = colour();
    const thresholds = [threshold(), threshold()];
    return { type: 'dataBar', priority: rule.priority, thresholds, colours: [c], showValue: !(options & 1), minLength: min, maxLength: max };
  }
  if (ct === 6) {
    q += 3;
    const n = u8(d, q);
    const set = u8(d, q + 1);
    const options = u8(d, q + 2);
    q += 3;
    const thresholds = [];
    for (let i = 0; i < n; i++) thresholds.push(threshold(() => { const gte = u8(d, q) !== 0; q += 5; return { gte }; }));
    return { type: 'iconSet', priority: rule.priority, thresholds, iconSet: ICON_SETS[set] || '3TrafficLights1', showValue: !(options & 1), reverse: Boolean(options & 4) };
  }
  return rule;
}

/** A rule's own formatting (DXFN): a number format, font, alignment, border, fill and protection, those it has; the font's colour, bold, italic and struck, and the fill, read. */
function readDxf(d, at) {
  const flags = u32(d, at);
  const ext = u16(d, at + 4);
  let p = at + 6;
  const out = {};
  if (flags & 0x02000000) p += ext & 1 ? u16(d, p) : 2;
  if (flags & 0x04000000) {
    const options = u32(d, p + 68);
    const weight = u16(d, p + 72);
    const colour = u32(d, p + 80);
    const modified = u32(d, p + 88);
    if (!(modified & 0x02)) out.italic = Boolean(options & 0x02);
    if (!(modified & 0x80)) out.strike = Boolean(options & 0x80);
    if (u32(d, p + 100) === 0 && weight) out.bold = weight >= 600;
    if (colour !== 0xffffffff && colour < 64) out.colour = colour;
    p += 118;
  }
  if (flags & 0x08000000) p += 8;
  if (flags & 0x10000000) p += 8;
  if (flags & 0x20000000) {
    const style = u16(d, p) >> 10;
    const colours = u16(d, p + 2);
    if (!(flags & 0x00010000)) out.pattern = style;
    if (!(flags & 0x00020000)) out.fg = colours & 0x7f;
    if (!(flags & 0x00040000)) out.bg = (colours >> 7) & 0x7f;
    p += 4;
  }
  if (flags & 0x40000000) p += 2;
  out.end = p;
  return out;
}

const OPERATORS = { 1: 'between', 2: 'notBetween', 3: 'equal', 4: 'notEqual', 5: 'greaterThan', 6: 'lessThan', 7: 'greaterThanOrEqual', 8: 'lessThanOrEqual' };
const CFVO = { 1: 'num', 2: 'min', 3: 'max', 4: 'percent', 5: 'percentile', 7: 'formula' };
const ICON_SETS = ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters'];

/** A FullColorExt: automatic, a palette index, an RGB colour or a theme colour with its tint (-1 to 1). */
function fullColour(d, at) {
  const type = u16(d, at);
  const tint = ((u16(d, at + 2) << 16) >> 16) / 32767;
  if (type === 2) return { rgb: [0, 1, 2].map((c) => u8(d, at + 4 + c).toString(16).padStart(2, '0')).join('').toUpperCase(), tint };
  if (type === 3) return { theme: u32(d, at + 4), tint };
  if (type === 1) return { indexed: u32(d, at + 4), tint };
  return null;
}

/** Byte arrays as one. */
function concat(parts) {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) { out.set(p, at); at += p.length; }
  return out;
}

const URL_MONIKER = 'e0c9ea79f9bace118c8200aa004ba90b';
const FILE_MONIKER = '0303000000000000c000000000000046';

/**
 * A hyperlink (HLINK): the cells it is on, and where it goes — a web
 * address, a file, or a place in the workbook — with the words shown for
 * it. A link's parts are present as its flags say.
 */
function readHyperlink(d) {
  const range = { top: u16(d, 0), bottom: u16(d, 2), left: u16(d, 4), right: u16(d, 6) };
  let p = 8 + 16 + 4;
  const flags = u32(d, p);
  p += 4;
  // A length in characters, the terminating zero counted, then the characters.
  const str = () => {
    const n = u32(d, p);
    p += 4;
    let s = '';
    for (let i = 0; i < n && p + i * 2 + 1 < d.length; i++) { const c = u16(d, p + i * 2); if (c) s += String.fromCharCode(c); }
    p += n * 2;
    return s;
  };
  const hex = (at) => Array.from(d.subarray(at, at + 16), (b) => b.toString(16).padStart(2, '0')).join('');
  let display = null;
  let href = null;
  let location = null;
  if (flags & 0x10) display = str();
  if (flags & 0x80) str(); // a target frame
  if (flags & 0x01) {
    if (flags & 0x100) href = str();
    else {
      const clsid = hex(p);
      p += 16;
      if (clsid === URL_MONIKER) {
        const len = u32(d, p);
        let s = '';
        for (let i = 0; i + 1 < len && p + 4 + i + 1 < d.length; i += 2) { const c = u16(d, p + 4 + i); if (!c) break; s += String.fromCharCode(c); }
        href = s;
        p += 4 + len;
      } else if (clsid === FILE_MONIKER) {
        const up = u16(d, p);
        const ansiLen = u32(d, p + 2);
        let ansi = '';
        for (let i = 0; i < ansiLen - 1; i++) ansi += String.fromCharCode(d[p + 6 + i]);
        let q = p + 6 + ansiLen + 2 + 2 + 16 + 4;
        const cbUnicode = u32(d, q);
        let path = ansi;
        if (cbUnicode) {
          const bytes = u32(d, q + 4);
          path = '';
          for (let i = 0; i + 1 < bytes; i += 2) path += String.fromCharCode(u16(d, q + 10 + i));
          q += 10 + bytes;
        } else q += 4;
        href = '../'.repeat(up) + path;
        p = q;
      }
    }
  }
  if (flags & 0x08) location = str();
  return href || location ? { range, href, location, display } : null;
}

/** The workbook's pictures: its drawing group's picture store, each picture in its record. */
function groupBlips(book) {
  if (book.blips) return book.blips;
  book.blips = [];
  const dg = book.drawingGroup;
  if (!dg) return book.blips;
  const dgg = artHeader(dg, 0);
  for (const c of artChildren(dg, artChild(dg, dgg, 0xf001))) {
    book.blips.push(c.type === 0xf007 ? findBlip(dg, c.body + 36 + u8(dg, c.body + 33), c.end) : null);
  }
  return book.blips;
}

/**
 * A sheet's pictures, from its drawing (the MSODRAWING records, one Office
 * Art stream between them): each picture shape and the cells its corners
 * are in, with how far into them. And its notes, each its words from the
 * text object it names.
 */
function finishDrawings(sheet, book) {
  for (const n of sheet.notes) if (n.text == null) n.text = sheet.objectText.get(n.object)?.text ?? '';
  // Before BIFF8 a chart's object gives its corners itself.
  for (const c of sheet.charts) if (c.object?.from) { c.anchor = { from: c.object.from, to: c.object.to }; c.order = c.object.id; }
  if (!sheet.drawing.length) return;
  const dg = concat(sheet.drawing);
  const blips = groupBlips(book);
  const objectAt = new Map((sheet.objects || []).map((o) => [o.at, o]));
  const i32 = (b, at) => u32(b, at) | 0;
  // Each shape's place and name, by where its client data ends — where the object record that follows it was met.
  // Each in the drawing's own order, so what was drawn over what still is.
  const placed = new Map();
  let order = 0;
  /**
   * A container's shapes. A group's own shape comes first: its place (cells,
   * or a box in the group it is in) and the coordinates its members' boxes
   * are given in — its frame.
   */
  const walk = (container, frame) => {
    let inner = frame;
    artChildren(dg, container).forEach((c, index) => {
      if (c.type === 0xf003) { walk(c, inner); return; }
      if (c.type !== 0xf004) return;
      const fsp = artChild(dg, c, 0xf00a);
      const flags = fsp ? u32(dg, fsp.body + 4) : 0;
      if (flags & 0x4 || flags & 0x8) return; // the drawing's own top shape, or a deleted one
      // The group's own shape is placed in the group it is in; its members in it.
      const place = placeOf(c, index === 0 ? frame : inner);
      if (!place) return;
      if (index === 0 && flags & 0x1) {
        const spgr = artChild(dg, c, 0xf009);
        if (spgr) {
          const x = i32(dg, spgr.body);
          const y = i32(dg, spgr.body + 4);
          inner = { place, coords: { x, y, w: i32(dg, spgr.body + 8) - x, h: i32(dg, spgr.body + 12) - y } };
        }
        return;
      }
      const props = readFopt(dg, artChild(dg, c, 0xf00b));
      // Excel 2007 and later keep the shape as they drew it too, as DrawingML (metroBlob); the writer checks it is still this shape.
      const drawingML = readFopt(dg, artChild(dg, c, 0xf122)).get(0x03a9)?.complex || null;
      const name = props.get(0x0380)?.complex;
      let label = null;
      if (name) { label = ''; for (let i = 0; i + 1 < name.length; i += 2) { const ch = u16(name, i); if (!ch) break; label += String.fromCharCode(ch); } }
      const client = artChild(dg, c, 0xf011);
      const object = client ? objectAt.get(client.end) : null;
      order += 1;
      if (client && place.anchor) placed.set(client.end, { ...place.anchor, name: label, order });
      const pib = props.get(0x0104)?.op;
      if (pib && blips[pib - 1]) { sheet.pictures.push({ blip: blips[pib - 1], name: label, ...(place.anchor || { place }), order }); return; }
      // A drawn shape — not a chart's frame, a note's box or a form control.
      if (object && ![1, 2, 3, 4, 6, 9, 0x1e].includes(object.type)) return;
      if (!object && !fsp?.inst) return;
      const shape = drawnShape(fsp.inst, flags, props, object ? sheet.objectText.get(object.id) : null, book);
      if (shape) sheet.shapes.push({ ...shape, name: label, ...(place.anchor || { place }), order, drawingML });
    });
  };
  /** A shape's place: the cells its corners are in (a client anchor), or a box in its group's coordinates. */
  const placeOf = (c, frame) => {
    const anchor = artChild(dg, c, 0xf010);
    if (anchor && anchor.len >= 18) {
      const a = anchor.body;
      return {
        anchor: {
          from: { col: u16(dg, a + 2), dx: u16(dg, a + 4) / 1024, row: u16(dg, a + 6), dy: u16(dg, a + 8) / 256 },
          to: { col: u16(dg, a + 10), dx: u16(dg, a + 12) / 1024, row: u16(dg, a + 14), dy: u16(dg, a + 16) / 256 },
          // Whether it moves and sizes with its cells (0), only moves with them (2), or does neither (3).
          edit: { 2: 'oneCell', 3: 'absolute' }[u16(dg, a) & 3] || null,
        },
      };
    }
    const ca = artChild(dg, c, 0xf00f);
    if (ca && frame) {
      const x = i32(dg, ca.body);
      const y = i32(dg, ca.body + 4);
      return { box: { x, y, w: i32(dg, ca.body + 8) - x, h: i32(dg, ca.body + 12) - y }, frame };
    }
    return null;
  };
  walk(artHeader(dg, 0), null);
  for (const c of sheet.charts) {
    const shape = c.object?.at != null ? placed.get(c.object.at) : null;
    if (shape) { c.anchor = { from: shape.from, to: shape.to }; c.name = shape.name; c.order = shape.order; }
  }
}

/**
 * A shape drawn on a sheet, in the terms the writer needs: its outline (a
 * preset, a line or connector, or its own points in a 100000-unit box), its
 * fill, line and shadow (Office Art's own reading, colours direct or from the
 * palette), turned and flipped as it is, and its words — a text box's or a
 * shape's, from its TXO — in paragraphs of runs in the workbook's fonts.
 */
function drawnShape(type, flags, props, txo, book) {
  const rotation = props.get(0x0004) ? (props.get(0x0004).op | 0) / 65536 : 0;
  const out = {
    kind: 'shape', type, rotation, flipH: Boolean(flags & 0x40), flipV: Boolean(flags & 0x80),
    textBox: type === 202, line: null, fill: null, shadow: null, preset: null, path: null, connector: false, paragraphs: [],
  };
  if (ART_LINES.has(type)) {
    out.connector = true;
    out.preset = CONNECTORS[type] || 'line';
    out.line = artLine(props, book.palette, true);
  } else {
    out.path = props.has(0x0145) ? freeformPath(props, { w: 100000, h: 100000 }) : null;
    out.preset = out.path ? null : type === 202 ? 'rect' : PRESETS[type] || 'rect';
    out.fill = artFill(props, book.palette);
    if (out.path && !out.path.filled) out.fill = 'none';
    out.line = artLine(props, book.palette, true);
    out.shadow = artShadow(props, book.palette);
  }
  // Its words: a paragraph to a line, each run in its font from where it starts.
  if (txo && txo.text) {
    const font = (i) => book.fonts[i < 4 ? i : i - 1] || book.fonts[0];
    const cuts = [...(txo.runs || [])].sort((a, b) => a.at - b.at);
    const fontAt = (pos) => { let f = cuts.length && cuts[0].at === 0 ? cuts[0].font : 0; for (const c of cuts) if (c.at <= pos) f = c.font; return f; };
    let pos = 0;
    for (const line of txo.text.split(/\r\n|\n|\r/)) {
      const runs = [];
      for (let i = 0; i < line.length; i++) {
        const f = fontAt(pos + i);
        if (!runs.length || runs[runs.length - 1].font !== f) runs.push({ font: f, text: '' });
        runs[runs.length - 1].text += line[i];
      }
      out.paragraphs.push({ runs: runs.map((r) => ({ text: r.text, ...fontLook(font(r.font), book) })) });
      pos += line.length + 1;
    }
    out.align = { 1: 'l', 2: 'ctr', 3: 'r', 4: 'just', 7: 'dist' }[txo.h] || 'l';
    out.anchor = { 1: 't', 2: 'ctr', 3: 'b', 4: 'just' }[txo.v] || 't';
  }
  out.insets = [0x81, 0x82, 0x83, 0x84].map((k, i) => (props.has(k) ? props.get(k).op | 0 : [91440, 45720, 91440, 45720][i]));
  return out;
}

/** A workbook font as a run's look. */
function fontLook(f, book) {
  if (!f) return {};
  const colour = f.colour != null && f.colour < 64 && f.colour !== 8 ? book.palette[f.colour] : null;
  return { font: f.name, size: f.height / 20, bold: f.bold, italic: f.italic, underline: Boolean(f.underline), strike: f.strike, colour, script: f.script };
}

/** Office Art's connectors, as DrawingML's. */
const CONNECTORS = {
  20: 'line', 32: 'straightConnector1', 33: 'bentConnector2', 34: 'bentConnector3', 35: 'bentConnector4', 36: 'bentConnector5',
  37: 'curvedConnector2', 38: 'curvedConnector3', 39: 'curvedConnector4', 40: 'curvedConnector5',
};

/** An RK number: a 30-bit integer or the top of a double, perhaps a hundredth of it. */
function rk(v) {
  let n;
  if (v & 2) n = (v | 0) >> 2;
  else {
    const buf = new DataView(new ArrayBuffer(8));
    buf.setUint32(4, v & 0xfffffffc, true);
    n = buf.getFloat64(0, true);
  }
  return v & 1 ? n / 100 : n;
}

/** A formula's last value: a number, or a string (in the STRING record after), a truth value, an error, nothing. */
function formulaResult(data, at) {
  if (u16(data, at + 6) === 0xffff) {
    switch (u8(data, at)) {
      case 0: return { t: 'str', v: '' };
      case 1: return { t: 'b', v: u8(data, at + 2) ? 1 : 0 };
      case 2: return { t: 'e', v: ERRORS[u8(data, at + 2)] || '#N/A' };
      default: return { t: 'str', v: '' };
    }
  }
  return { t: 'n', v: f64(data, at) };
}

/**
 * A cell format: its font and number format, whether it is a style's or a
 * cell's, its alignment, borders and fill — each version packing them its
 * own way.
 */
function readXf(d, biff) {
  const xf = { font: 0, format: 0, style: false, locked: true, hidden: false, h: 0, v: 2, wrap: false, rotation: 0, indent: 0, shrink: false, border: {}, fill: { pattern: 0, fg: 64, bg: 65 } };
  const line = (style, colour) => (style ? { style, colour } : null);
  if (biff === 8) {
    xf.font = u16(d, 0); xf.format = u16(d, 2);
    const tp = u16(d, 4); xf.locked = Boolean(tp & 1); xf.hidden = Boolean(tp & 2); xf.style = Boolean(tp & 4);
    const al = u16(d, 6); xf.h = al & 7; xf.wrap = Boolean(al & 8); xf.v = (al >> 4) & 7; xf.rotation = (al >> 8) & 0xff;
    const misc = u16(d, 8); xf.indent = misc & 0x0f; xf.shrink = Boolean(misc & 0x10); xf.rtl = ((misc >> 6) & 3) === 2;
    const b1 = u32(d, 10); const b2 = u32(d, 14); const area = u16(d, 18);
    xf.border = { left: line(b1 & 0xf, (b1 >> 16) & 0x7f), right: line((b1 >> 4) & 0xf, (b1 >> 23) & 0x7f), top: line((b1 >> 8) & 0xf, b2 & 0x7f), bottom: line((b1 >> 12) & 0xf, (b2 >> 7) & 0x7f) };
    xf.fill = { pattern: (b2 >>> 26) & 0x3f, fg: area & 0x7f, bg: (area >> 7) & 0x7f };
  } else if (biff === 5) {
    xf.font = u16(d, 0); xf.format = u16(d, 2);
    const tp = u16(d, 4); xf.locked = Boolean(tp & 1); xf.hidden = Boolean(tp & 2); xf.style = Boolean(tp & 4);
    const al = u16(d, 6); xf.h = al & 7; xf.wrap = Boolean(al & 8); xf.v = (al >> 4) & 7;
    const area = u32(d, 8); const b = u32(d, 12);
    xf.fill = { pattern: (area >> 16) & 0x3f, fg: area & 0x7f, bg: (area >> 7) & 0x7f };
    xf.border = { top: line(b & 7, (b >> 9) & 0x7f), left: line((b >> 3) & 7, (b >> 16) & 0x7f), right: line((b >> 6) & 7, (b >>> 23) & 0x7f), bottom: line((area >> 22) & 7, (area >>> 25) & 0x7f) };
  } else if (biff === 3 || biff === 4) {
    xf.font = u8(d, 0); xf.format = u8(d, 1);
    const tp = u16(d, 2); xf.locked = Boolean(tp & 1); xf.hidden = Boolean(tp & 2); xf.style = Boolean(tp & 4);
    const al = u16(d, 4); xf.h = al & 7; xf.wrap = Boolean(al & 8); if (biff === 4) xf.v = (al >> 4) & 3;
    const area = u16(d, 6); const b = u32(d, 8);
    xf.fill = { pattern: area & 0x3f, fg: (area >> 6) & 0x1f, bg: (area >> 11) & 0x1f };
    xf.border = { top: line(b & 7, (b >> 3) & 0x1f), left: line((b >> 8) & 7, (b >> 11) & 0x1f), bottom: line((b >> 16) & 7, (b >> 19) & 0x1f), right: line((b >> 24) & 7, (b >>> 27) & 0x1f) };
  } else {
    xf.font = u8(d, 0); xf.format = u8(d, 2) & 0x3f; xf.locked = Boolean(u8(d, 2) & 0x40); xf.hidden = Boolean(u8(d, 2) & 0x80);
    const f = u8(d, 3); xf.h = f & 7;
    xf.border = { left: line(f & 0x08 ? 1 : 0, 8), right: line(f & 0x10 ? 1 : 0, 8), top: line(f & 0x20 ? 1 : 0, 8), bottom: line(f & 0x40 ? 1 : 0, 8) };
    xf.fill = { pattern: f & 0x80 ? 17 : 0, fg: 8, bg: 9 };
  }
  return xf;
}

/** A defined name: its own name (a built-in one by its code), where it applies, and its formula's tokens. */
function readName(d, biff, str, decode) {
  const flags = biff === 2 ? u8(d, 0) : u16(d, 0);
  const builtin = biff >= 5 ? Boolean(flags & 0x20) : false;
  const hidden = Boolean(flags & 1);
  let cch;
  let cce;
  let nameAt;
  let sheet = null;
  if (biff >= 5) {
    cch = u8(d, 3); cce = u16(d, 4); nameAt = 14;
    const itab = u16(d, 8);
    if (itab) sheet = itab - 1;
  } else if (biff >= 3) {
    cch = u8(d, 3); cce = u16(d, 4); nameAt = 6;
  } else {
    cch = u8(d, 3); cce = u8(d, 4); nameAt = 5;
  }
  let name;
  let end;
  if (biff === 8) {
    const high = u8(d, nameAt) & 1;
    name = '';
    for (let i = 0; i < cch; i++) name += String.fromCharCode(high ? u16(d, nameAt + 1 + i * 2) : u8(d, nameAt + 1 + i));
    end = nameAt + 1 + cch * (high ? 2 : 1);
  } else {
    name = decode(d.subarray(nameAt, nameAt + cch));
    end = nameAt + cch;
  }
  if (builtin) name = BUILTIN_NAMES[name.charCodeAt(0)] || name;
  return { name, builtin, hidden, sheet, rgce: d.subarray(end, end + cce), extra: d.subarray(end + cce) };
}

const BUILTIN_NAMES = ['_xlnm.Consolidate_Area', '_xlnm.Auto_Open', '_xlnm.Auto_Close', '_xlnm.Extract', '_xlnm.Database', '_xlnm.Criteria', '_xlnm.Print_Area', '_xlnm.Print_Titles', '_xlnm.Recorder', '_xlnm.Data_Form', '_xlnm.Auto_Activate', '_xlnm.Auto_Deactivate', '_xlnm.Sheet_Title', '_xlnm._FilterDatabase'];

/* ── formulas ─────────────────────────────────────────────────────────── */

const colName = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const BINARY = { 0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<', 0x0a: '<=', 0x0b: '=', 0x0c: '>=', 0x0d: '>', 0x0e: '<>', 0x0f: ' ', 0x10: ',', 0x11: ':' };
/** A sheet's name as a formula writes it: quoted when it is not a plain word. */
const sheetRef = (name) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
const num = (v) => (Number.isInteger(v) ? String(v) : String(v).toUpperCase());

/**
 * A formula's tokens as the text Excel shows (without its "="). `ctx`:
 * the version, the cell the formula is in (a shared formula's references
 * are offsets from it), the workbook (its names and other sheets).
 */
function decompile(rgce, extra, ctx) {
  const { biff, book, decode } = ctx;
  const b8 = biff === 8;
  const stack = [];
  let rgcb = 0;
  const pop = () => stack.pop() ?? '';
  const cellText = (row, col, rowRel, colRel) => `${colRel ? '' : '$'}${colName(col)}${rowRel ? '' : '$'}${row + 1}`;
  const maxRow = b8 ? 65535 : 16383;
  const areaText = (r1, r2, c1, c2, rr1, rr2, cr1, cr2) => {
    if (r1 === 0 && r2 === maxRow) return `${cr1 ? '' : '$'}${colName(c1)}:${cr2 ? '' : '$'}${colName(c2)}`;
    if (c1 === 0 && c2 === 255) return `${rr1 ? '' : '$'}${r1 + 1}:${rr2 ? '' : '$'}${r2 + 1}`;
    return `${cellText(r1, c1, rr1, cr1)}:${cellText(r2, c2, rr2, cr2)}`;
  };
  // A reference's row and column, with its relative flags; `relative` for one written as an offset (in a shared formula or a name).
  const ref = (b, at, relative) => {
    let row;
    let col;
    let rowRel;
    let colRel;
    if (b8) {
      row = u16(b, at);
      const c = u16(b, at + 2);
      col = c & 0x3fff;
      rowRel = Boolean(c & 0x8000);
      colRel = Boolean(c & 0x4000);
      if (relative) {
        if (rowRel) row = (ctx.row + ((row << 16) >> 16) + 65536) % 65536;
        if (colRel) col = (ctx.col + ((col << 24) >> 24) + 256) % 256;
      }
    } else {
      const r = u16(b, at);
      row = r & 0x3fff;
      rowRel = Boolean(r & 0x8000);
      colRel = Boolean(r & 0x4000);
      col = u8(b, at + 2);
      if (relative) {
        if (rowRel) row = (ctx.row + ((row << 18) >> 18) + 16384) % 16384;
        if (colRel) col = (ctx.col + ((col << 24) >> 24) + 256) % 256;
      }
    }
    return { row, col, rowRel, colRel };
  };
  const area = (b, at, relative) => {
    if (b8) {
      const top = ref(Uint8Array.of(b[at], b[at + 1], b[at + 4], b[at + 5]), 0, relative);
      const bottom = ref(Uint8Array.of(b[at + 2], b[at + 3], b[at + 6], b[at + 7]), 0, relative);
      return areaText(top.row, bottom.row, top.col, bottom.col, top.rowRel, bottom.rowRel, top.colRel, bottom.colRel);
    }
    const top = ref(Uint8Array.of(b[at], b[at + 1], b[at + 4]), 0, relative);
    const bottom = ref(Uint8Array.of(b[at + 2], b[at + 3], b[at + 5]), 0, relative);
    return areaText(top.row, bottom.row, top.col, bottom.col, top.rowRel, bottom.rowRel, top.colRel, bottom.colRel);
  };
  // Another sheet (or a range of sheets) before a reference, from BIFF8's EXTERNSHEET or BIFF5's sheet numbers.
  const sheetsPrefix8 = (ixti) => {
    const x = book.externSheets[ixti];
    const sup = x ? book.supbooks[x.supbook] : null;
    if (!x || !sup || !sup.internal || x.first >= 0xfffe) return null;
    const a = ctx.sheetNames[x.first];
    const z = ctx.sheetNames[x.last];
    if (a == null) return null;
    return sheetRef(x.first === x.last || z == null ? a : `${a}:${z}`) + '!';
  };
  const sheetsPrefix5 = (first, last) => {
    const a = ctx.sheetNames[first];
    const z = ctx.sheetNames[last];
    if (a == null) return null;
    return sheetRef(first === last || z == null ? a : `${a}:${z}`) + '!';
  };
  const nameOf = (index) => book.names[index - 1]?.name ?? '#NAME?';
  const funcCall = (name, argc) => {
    const args = stack.splice(Math.max(0, stack.length - argc), argc);
    stack.push(`${name}(${args.join(',')})`);
  };
  // A constant array, from the bytes after the formula: {1,2;3,4}.
  const arrayConstant = () => {
    const cols = u8(extra, rgcb) + 1;
    const rows = u16(extra, rgcb + 1) + 1;
    let p = rgcb + 3;
    const out = [];
    for (let r = 0; r < rows; r++) {
      const line = [];
      for (let c = 0; c < cols; c++) {
        const type = u8(extra, p++);
        if (type === 0x01) { line.push(num(f64(extra, p))); p += 8; }
        else if (type === 0x02) {
          if (b8) {
            const cch = u16(extra, p); const high = u8(extra, p + 2) & 1; let s = '';
            for (let i = 0; i < cch; i++) s += String.fromCharCode(high ? u16(extra, p + 3 + i * 2) : u8(extra, p + 3 + i));
            line.push(`"${s.replace(/"/g, '""')}"`); p += 3 + cch * (high ? 2 : 1);
          } else { const cch = u8(extra, p); line.push(`"${decode(extra.subarray(p + 1, p + 1 + cch)).replace(/"/g, '""')}"`); p += 1 + cch; }
        } else if (type === 0x04) { line.push(u8(extra, p) ? 'TRUE' : 'FALSE'); p += 8; }
        else if (type === 0x10) { line.push(ERRORS[u8(extra, p)] || '#N/A'); p += 8; }
        else { line.push(''); p += 8; }
      }
      out.push(line.join(','));
    }
    rgcb = p;
    return `{${out.join(';')}}`;
  };

  let i = 0;
  while (i < rgce.length) {
    const ptg = rgce[i++];
    if (BINARY[ptg]) { const b = pop(); const a = pop(); stack.push(`${a}${BINARY[ptg]}${b}`); continue; }
    switch (ptg) {
      case 0x01: case 0x02: i += biff === 2 ? 3 : 4; break; // a pointer handled before this
      case 0x12: stack.push('+' + pop()); break;
      case 0x13: stack.push('-' + pop()); break;
      case 0x14: stack.push(pop() + '%'); break;
      case 0x15: stack.push('(' + pop() + ')'); break;
      case 0x16: stack.push(''); break;
      case 0x17: {
        const cch = u8(rgce, i);
        let s;
        if (b8) {
          const high = u8(rgce, i + 1) & 1;
          s = '';
          for (let k = 0; k < cch; k++) s += String.fromCharCode(high ? u16(rgce, i + 2 + k * 2) : u8(rgce, i + 2 + k));
          i += 2 + cch * (high ? 2 : 1);
        } else { s = decode(rgce.subarray(i + 1, i + 1 + cch)); i += 1 + cch; }
        stack.push(`"${s.replace(/"/g, '""')}"`);
        break;
      }
      case 0x18: i += b8 ? 4 : 0; break; // extended tokens (BIFF8's ptgElf and the like): skipped
      case 0x19: { // attributes: SUM of one argument, or control tokens with nothing to show
        const opt = u8(rgce, i);
        const data = biff === 2 ? u8(rgce, i + 1) : u16(rgce, i + 1);
        i += biff === 2 ? 2 : 3;
        if (opt & 0x04) i += (data + 1) * (biff === 2 ? 1 : 2);
        else if (opt & 0x10) stack.push(`SUM(${pop()})`);
        break;
      }
      case 0x1c: stack.push(ERRORS[u8(rgce, i)] || '#N/A'); i += 1; break;
      case 0x1d: stack.push(u8(rgce, i) ? 'TRUE' : 'FALSE'); i += 1; break;
      case 0x1e: stack.push(String(u16(rgce, i))); i += 2; break;
      case 0x1f: stack.push(num(f64(rgce, i))); i += 8; break;
      default: {
        const base = (ptg & 0x1f) | 0x20;
        if (ptg < 0x20 || ptg > 0x7f) throw new XlsError('unknown formula token ' + ptg);
        switch (base) {
          case 0x20: i += biff === 2 ? 6 : 7; stack.push(arrayConstant()); break;
          case 0x21: {
            const id = biff <= 3 ? u8(rgce, i) : u16(rgce, i);
            i += biff <= 3 ? 1 : 2;
            const f = XLS_FUNCTIONS[id];
            funcCall(f ? f[0] : `FUNCTION.${id}`, f ? f[1] : 0);
            break;
          }
          case 0x22: {
            const argc = u8(rgce, i) & 0x7f;
            const id = (biff <= 3 ? u8(rgce, i + 1) : u16(rgce, i + 1)) & 0x7fff;
            i += biff <= 3 ? 2 : 3;
            if (id === 255) { // an add-in's or a newer Excel's function, named by its first argument
              const args = stack.splice(Math.max(0, stack.length - argc), argc);
              const name = args.shift() || '#NAME?';
              stack.push(`${name}(${args.join(',')})`);
            } else funcCall(XLS_FUNCTIONS[id]?.[0] ?? `FUNCTION.${id}`, argc);
            break;
          }
          case 0x23: stack.push(nameOf(u16(rgce, i))); i += { 2: 7, 3: 10, 4: 10, 5: 14, 8: 4 }[biff]; break;
          case 0x24: case 0x2a: {
            const r = ref(rgce, i, false);
            i += b8 ? 4 : 3;
            stack.push(base === 0x2a ? '#REF!' : cellText(r.row, r.col, r.rowRel, r.colRel));
            break;
          }
          case 0x25: case 0x2b: { const t = area(rgce, i, false); i += b8 ? 8 : 6; stack.push(base === 0x2b ? '#REF!' : t); break; }
          case 0x26: case 0x27: case 0x28: i += biff === 2 ? 4 : 6; break; // a subexpression's header: what follows is read as usual
          case 0x29: case 0x2e: case 0x2f: i += biff === 2 ? 1 : 2; break;
          case 0x2c: { const r = ref(rgce, i, true); i += b8 ? 4 : 3; stack.push(cellText(r.row, r.col, r.rowRel, r.colRel)); break; }
          case 0x2d: { stack.push(area(rgce, i, true)); i += b8 ? 8 : 6; break; }
          case 0x38: i += 2; break;
          case 0x39: { // a name in another book, or an add-in function's name
            if (b8) {
              const x = book.externSheets[u16(rgce, i)];
              const sup = x ? book.supbooks[x.supbook] : null;
              const name = sup?.names[u16(rgce, i + 2) - 1];
              stack.push(sup?.internal ? nameOf(u16(rgce, i + 2)) : name ?? '#NAME?');
              i += 6;
            } else {
              const ixals = i16(rgce, i);
              const idx = u16(rgce, i + 10);
              const list = ixals > 0 ? book.externSheets[ixals - 1]?.names : book.biff5ExternNames;
              stack.push(list?.[idx - 1] ?? nameOf(idx));
              i += 24;
            }
            break;
          }
          case 0x3a: case 0x3c: {
            let prefix;
            let r;
            if (b8) { prefix = sheetsPrefix8(u16(rgce, i)); r = ref(rgce, i + 2, false); i += 6; }
            else { prefix = sheetsPrefix5(u16(rgce, i + 10), u16(rgce, i + 12)); r = ref(rgce, i + 14, false); i += 17; }
            stack.push(base === 0x3c || prefix == null ? '#REF!' : prefix + cellText(r.row, r.col, r.rowRel, r.colRel));
            break;
          }
          case 0x3b: case 0x3d: {
            let prefix;
            let t;
            if (b8) { prefix = sheetsPrefix8(u16(rgce, i)); t = area(rgce, i + 2, false); i += 10; }
            else { prefix = sheetsPrefix5(u16(rgce, i + 10), u16(rgce, i + 12)); t = area(rgce, i + 14, false); i += 20; }
            stack.push(base === 0x3d || prefix == null ? '#REF!' : prefix + t);
            break;
          }
          default: throw new XlsError('unknown formula token ' + ptg);
        }
      }
    }
  }
  return stack.join('');
}

export { decompile as decompileFormula, rk as rkNumber };
