// dBASE and FoxPro tables (.dbf), with their memo files (.dbt for dBASE,
// .fpt for FoxPro): a header naming the fields, then fixed-length records,
// each starting with a mark that says whether it was deleted. Read into the
// old spreadsheet readers' sheets (legacy-grid.js): a row of field names,
// then a row a record, deleted records left out and counted in the notes.

import { decodeIn } from './codepage.js';
import { Grid } from './legacy-grid.js';

/** The versions a .dbf opens with: what wrote it, and whether it has memos. */
const VERSIONS = {
  0x02: 'FoxBASE', 0x03: 'dBASE III', 0x04: 'dBASE IV', 0x05: 'dBASE V', 0x30: 'Visual FoxPro', 0x31: 'Visual FoxPro',
  0x32: 'Visual FoxPro', 0x43: 'dBASE IV', 0x63: 'dBASE IV', 0x83: 'dBASE III', 0x8b: 'dBASE IV', 0xcb: 'dBASE IV',
  0xf5: 'FoxPro', 0xfb: 'FoxBASE',
};

/** The language driver byte, as the code page it names. */
const CODE_PAGES = {
  0x01: 'cp437', 0x02: 'cp850', 0x03: 'windows-1252', 0x57: 'windows-1252', 0x58: 'windows-1252', 0x59: 'windows-1252',
  0x64: 'ibm852', 0x65: 'ibm866', 0x66: 'ibm865', 0x67: 'ibm861', 0x6a: 'ibm737', 0x6b: 'ibm857', 0x7d: 'windows-1255',
  0x7e: 'windows-1256', 0xc8: 'windows-1250', 0xc9: 'windows-1251', 0xca: 'windows-1254', 0xcb: 'windows-1253',
  0x4d: 'gb18030', 0x4e: 'euc-kr', 0x4f: 'big5', 0x13: 'shift_jis', 0x7b: 'shift_jis', 0x7a: 'gb18030', 0x78: 'big5', 0x79: 'euc-kr',
};

/** Whether bytes look like a .dbf: a known version, a sane header, and records that fit. */
export function isDbf(bytes) {
  if (bytes.length < 33 || !VERSIONS[bytes[0]]) return false;
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const headerSize = v.getUint16(8, true);
  const recordSize = v.getUint16(10, true);
  const month = bytes[2];
  const day = bytes[3];
  return headerSize >= 33 && headerSize <= bytes.length && recordSize > 0 && month >= 1 && month <= 12 && day >= 1 && day <= 31;
}

/** The fields: { name, type, length, decimals, offset } from the descriptors after the header. */
function fieldsOf(bytes, headerSize) {
  const out = [];
  let offset = 1;
  for (let at = 32; at + 32 <= headerSize && bytes[at] !== 0x0d; at += 32) {
    let end = at;
    while (end < at + 11 && bytes[end] !== 0) end++;
    const name = String.fromCharCode(...bytes.subarray(at, end)).trim();
    const type = String.fromCharCode(bytes[at + 11]);
    const length = bytes[at + 16];
    const decimals = bytes[at + 17];
    out.push({ name, type, length, decimals, offset });
    offset += length;
  }
  return out;
}

/** A memo file's reader: block number to text (or bytes, for a picture or other binary). */
function memoReader(memo, kind, decode) {
  if (!memo || memo.length < 512) return null;
  const v = new DataView(memo.buffer, memo.byteOffset, memo.byteLength);
  if (kind === 'fpt') {
    const size = v.getUint16(6, false) || 64;
    return (block) => {
      const at = block * size;
      if (block <= 0 || at + 8 > memo.length) return null;
      const type = v.getUint32(at, false);
      const length = v.getUint32(at + 4, false);
      const data = memo.subarray(at + 8, Math.min(memo.length, at + 8 + length));
      return type === 1 ? decode(data) : null;
    };
  }
  // dBASE IV writes its block size and starts a block with FF FF 08 00 and its length; dBASE III ends with 1A 1A.
  const size = v.getUint16(20, true) || 512;
  return (block) => {
    const at = block * (size >= 64 ? size : 512);
    if (block <= 0 || at >= memo.length) return null;
    if (memo[at] === 0xff && memo[at + 1] === 0xff && memo[at + 2] === 0x08 && memo[at + 3] === 0x00) {
      const length = v.getUint32(at + 4, true);
      return decode(memo.subarray(at + 8, Math.min(memo.length, at + length)));
    }
    let end = at;
    while (end < memo.length && !(memo[end] === 0x1a && memo[end + 1] === 0x1a)) end++;
    return decode(memo.subarray(at, end));
  };
}

const julianToDate = (jd) => {
  // Julian day number to a calendar date, as Visual FoxPro counts days.
  let l = jd + 68569;
  const n = Math.floor((4 * l) / 146097);
  l -= Math.floor((146097 * n + 3) / 4);
  const i = Math.floor((4000 * (l + 1)) / 1461001);
  l = l - Math.floor((1461 * i) / 4) + 31;
  const j = Math.floor((80 * l) / 2447);
  const day = l - Math.floor((2447 * j) / 80);
  l = Math.floor(j / 11);
  return { y: 100 * (n - 49) + i + l, m: j + 2 - 12 * l, d: day };
};

/** A date as the spreadsheet's serial number (days since 1899-12-30). */
const serial = (y, m, d, ms = 0) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000 + ms / 86400000;

/**
 * Read a .dbf, with its memo file's bytes when there is one
 * (`memo`, `memoKind` 'dbt' or 'fpt'): the old spreadsheet readers' model,
 * one sheet named after the table.
 */
export function readDbf(bytes, { name = 'Table', memo = null, memoKind = null, codePage = null } = {}) {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const version = bytes[0];
  const count = v.getUint32(4, true);
  const headerSize = v.getUint16(8, true);
  const recordSize = v.getUint16(10, true);
  const page = codePage || CODE_PAGES[bytes[29]] || 'cp437';
  const decode = (b) => decodeIn(page, b);
  const fields = fieldsOf(bytes, headerSize).filter((f) => f.type !== '0');
  const kind = memoKind || ((version & 0xf0) === 0x30 || version === 0xf5 ? 'fpt' : 'dbt');
  const memoText = memoReader(memo, kind, decode);
  const grid = new Grid(String(name).replace(/\.dbf$/i, '').slice(0, 31) || 'Table');
  const notes = [];
  fields.forEach((f, c) => {
    grid.set(0, c, f.name, { bold: true });
    grid.width(c, Math.max(f.name.length, Math.min(40, f.type === 'M' ? 40 : f.length)) + 2);
  });
  let row = 1;
  let deleted = 0;
  let missingMemo = false;
  for (let r = 0; r < count; r++) {
    const at = headerSize + r * recordSize;
    if (at + recordSize > bytes.length) break;
    if (bytes[at] === 0x2a) { deleted++; continue; }
    fields.forEach((f, c) => {
      const raw = bytes.subarray(at + f.offset, at + f.offset + f.length);
      const text = () => decode(raw).replace(/\0+$/, '').trimEnd();
      let value = null;
      let style = null;
      switch (f.type) {
        case 'C': value = text() || null; break;
        case 'N': case 'F': {
          const s = text().trim();
          value = s === '' || /^\*+$/.test(s) ? null : Number(s);
          if (Number.isNaN(value)) value = s;
          if (f.decimals) style = { numFmt: `0.${'0'.repeat(f.decimals)}` };
          break;
        }
        case 'D': {
          const s = text().trim();
          if (/^\d{8}$/.test(s)) { value = serial(Number(s.slice(0, 4)), Number(s.slice(4, 6)), Number(s.slice(6, 8))); style = { numFmt: 'yyyy-mm-dd' }; }
          break;
        }
        case 'L': {
          const ch = String.fromCharCode(raw[0]).toUpperCase();
          value = 'TY'.includes(ch) ? true : 'FN'.includes(ch) ? false : null;
          break;
        }
        case 'I': value = new DataView(raw.buffer, raw.byteOffset, 4).getInt32(0, true); break;
        case 'B': {
          if ((version & 0xf0) === 0x30) value = new DataView(raw.buffer, raw.byteOffset, 8).getFloat64(0, true);
          break;
        }
        case 'Y': {
          const dv = new DataView(raw.buffer, raw.byteOffset, 8);
          value = Number(dv.getBigInt64(0, true)) / 10000;
          style = { numFmt: '#,##0.0000' };
          break;
        }
        case 'T': {
          const dv = new DataView(raw.buffer, raw.byteOffset, 8);
          const jd = dv.getInt32(0, true);
          if (jd > 0) { const { y, m, d } = julianToDate(jd); value = serial(y, m, d, dv.getInt32(4, true)); style = { numFmt: 'yyyy-mm-dd hh:mm:ss' }; }
          break;
        }
        case 'M': case 'G': case 'P': {
          const block = f.length === 4 ? new DataView(raw.buffer, raw.byteOffset, 4).getInt32(0, true) : Number(text().trim() || 0);
          if (block > 0) {
            if (!memoText) missingMemo = true;
            else if (f.type === 'M') value = memoText(block);
          }
          break;
        }
        default: value = text() || null;
      }
      if (value != null) grid.set(row, c, value, style);
    });
    row++;
  }
  if (deleted) notes.push({ code: 'deleted', count: deleted });
  if (missingMemo) notes.push({ code: 'memo-missing' });
  return { sheets: [grid.toSheet()], names: [], notes, version: VERSIONS[version] || 'dBASE' };
}
