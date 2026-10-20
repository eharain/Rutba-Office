// SYLK (.slk), Multiplan's and Excel's symbolic link format: a line a
// record, fields after semicolons. C records hold cells (their value in K,
// their formula in E, in R1C1 references), F records formats and column
// widths, P records the number formats a cell names by number, NN records
// names. Read into the old spreadsheet readers' sheets (legacy-grid.js).

import { decode1252 } from './codepage.js';
import { Grid, cellRef } from './legacy-grid.js';

/** Whether bytes look like SYLK: it starts with an ID record. */
export const isSylk = (bytes) => bytes.length > 3 && bytes[0] === 0x49 && bytes[1] === 0x44 && bytes[2] === 0x3b;

/** A record's fields, `;;` being a semicolon inside one. */
function fields(line) {
  const out = [];
  let cur = '';
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === ';') {
      if (line[i + 1] === ';') { cur += ';'; i++; continue; }
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

/**
 * An R1C1 formula as A1 from the cell at (row, col): R[-1]C is the cell
 * above, R2C3 the absolute $C$2. Words in quotes are left as they are.
 */
export function r1c1ToA1(expr, row, col) {
  const part = (spec, here) => {
    if (spec == null || spec === '') return { at: here, abs: false };
    if (spec.startsWith('[')) return { at: here + Number(spec.slice(1, -1)), abs: false };
    return { at: Number(spec) - 1, abs: true };
  };
  return String(expr).split(/("(?:[^"]|"")*")/).map((piece, i) => {
    if (i % 2 === 1) return piece;
    return piece.replace(/(?<![A-Za-z0-9_.$])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(])/g, (_, r, c) => {
      const R = part(r, row);
      const C = part(c, col);
      if (R.at < 0 || C.at < 0) return '#REF!';
      return cellRef(R.at, C.at, { absRow: R.abs, absCol: C.abs });
    });
  }).join('');
}

/** A SYLK number format letter, as an .xlsx number format. */
function formatOf(code) {
  const m = /^([A-Z$%*])(\d*)/.exec(code || '');
  if (!m) return null;
  const places = Number(m[2] || 0);
  const dec = places ? `.${'0'.repeat(places)}` : '';
  switch (m[1]) {
    case 'F': return `0${dec}`;
    case '$': return `#,##0${dec}`;
    case '%': return `0${dec}%`;
    case 'E': return `0${dec}E+00`;
    case 'C': return `#,##0${dec}`;
    default: return null;
  }
}

/** A cell's value as SYLK writes it in K: a number, "text", TRUE, FALSE or an error. */
function valueOf(k) {
  if (k.startsWith('"')) return k.slice(1, k.endsWith('"') && k.length > 1 ? -1 : undefined).replace(/""/g, '"');
  if (k === 'TRUE') return true;
  if (k === 'FALSE') return false;
  if (/^#/.test(k)) return k;
  const n = Number(k);
  return Number.isFinite(n) ? n : k;
}

/** Read a SYLK file: the old spreadsheet readers' model, one sheet. */
export function readSylk(bytes, { name = 'Sheet1' } = {}) {
  const text = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? new TextDecoder('utf-8').decode(bytes.subarray(3)) : decode1252(bytes);
  const grid = new Grid(name);
  const pictures = [];
  const names = [];
  const notes = [];
  let row = 0;
  let col = 0;
  for (const raw of text.split(/\r\n|\n|\r/)) {
    if (!raw) continue;
    const [kind, ...rest] = fields(raw);
    if (kind === 'E') break;
    const f = {};
    for (const x of rest) if (x) f[x[0]] = x.slice(1);
    if (f.Y) row = Number(f.Y) - 1;
    if (f.X) col = Number(f.X) - 1;
    if (kind === 'P' && f.P != null) {
      pictures.push(f.P === 'General' ? null : f.P);
    } else if (kind === 'C') {
      const value = f.K != null ? valueOf(f.K) : null;
      if (f.E != null) grid.set(row, col, `=${r1c1ToA1(f.E, row, col)}`);
      else if (f.K != null) grid.set(row, col, value);
    } else if (kind === 'F') {
      if (f.W) {
        const [a, b, w] = f.W.split(/\s+/).map(Number);
        for (let c = a; c <= b; c++) grid.width(c - 1, w);
      }
      if ((f.Y || f.X) && (f.P != null || f.F || f.S)) {
        const style = {};
        const pic = f.P != null ? pictures[Number(f.P)] : null;
        const fmt = pic || (f.F ? formatOf(f.F) : null);
        if (fmt) style.numFmt = fmt;
        const align = /([LRC])$/.exec(f.F || '')?.[1];
        if (align) style.align = { L: 'left', R: 'right', C: 'center' }[align];
        if (f.S?.includes('D')) style.bold = true;
        if (f.S?.includes('I')) style.italic = true;
        if (Object.keys(style).length) grid.styles[cellRef(row, col)] = { ...(grid.styles[cellRef(row, col)] || {}), ...style };
      }
    } else if (kind === 'NN' && f.N && f.E) {
      names.push({ name: f.N, ref: r1c1ToA1(f.E, 0, 0).replace(/^=/, '') });
    }
  }
  if (!grid.cells.length) notes.push({ code: 'empty' });
  return { sheets: [grid.toSheet()], names, notes };
}
