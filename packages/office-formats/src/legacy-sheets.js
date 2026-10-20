// The old spreadsheets Worksheets opens and saves on as .xlsx: Lotus 1-2-3
// (.wks, .wk1, .wk3, .wk4), SYLK (.slk), DIF (.dif) and dBASE and FoxPro
// tables (.dbf). Each format's reader makes sheets of cells (legacy-grid.js);
// this tells them apart and writes the workbook.

import { buildXlsx } from '@rutba/ooxml/build';
import { readSylk, isSylk } from './sylk.js';
import { readDif, isDif } from './dif.js';
import { readDbf, isDbf } from './dbf.js';
import { readLotus, lotusKind } from './lotus.js';

/** Which old spreadsheet bytes are: 'wk1', 'wk3', 'slk', 'dif', 'dbf' or null. `ext` settles a .dbf, whose header is loose. */
export function legacySheetKind(bytes, ext = '') {
  const lotus = lotusKind(bytes);
  if (lotus) return lotus;
  if (isSylk(bytes)) return 'slk';
  if (isDif(bytes)) return 'dif';
  if (isDbf(bytes) && (!ext || /^\.?dbf$/i.test(ext))) return 'dbf';
  return null;
}

/** Read an old spreadsheet as the readers' model, by kind. `memo` is a .dbf's memo file when there is one. */
export function readLegacySheets(bytes, kind, { name, memo = null, memoKind = null } = {}) {
  switch (kind) {
    case 'wk1': case 'wk3': return readLotus(bytes, { name });
    case 'slk': return readSylk(bytes, { name });
    case 'dif': return readDif(bytes, { name });
    case 'dbf': return readDbf(bytes, { name, memo, memoKind });
    default: throw new Error('not an old spreadsheet this suite reads');
  }
}

/** A sheet name as .xlsx allows: none of []:*?/\, at most 31 characters, each its own. */
function names(sheets) {
  const used = new Set();
  return sheets.map((s, i) => {
    const base = String(s.name || '').replace(/[[\]:*?/\\]/g, '_').trim().slice(0, 31) || `Sheet${i + 1}`;
    let name = base;
    for (let n = 2; used.has(name.toLowerCase()); n++) name = `${base.slice(0, 28)} (${n})`;
    used.add(name.toLowerCase());
    return name;
  });
}

/** The readers' model as .xlsx bytes, its names on the first sheet. */
export function legacySheetsToXlsx(model) {
  const sheetNames = names(model.sheets);
  const sheets = model.sheets.map((s, i) => ({ ...s, name: sheetNames[i], rows: s.rows.length ? s.rows : [[null]] }));
  const quoted = `'${sheetNames[0].replace(/'/g, "''")}'`;
  const definedNames = (model.names || [])
    .filter((n) => /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n.name) && !/^[A-Za-z]{1,3}\d+$/.test(n.name))
    .map((n) => ({ name: n.name, ref: `${quoted}!${n.ref}` }));
  return buildXlsx({ sheets, definedNames });
}
