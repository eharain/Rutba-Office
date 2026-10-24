// A workbook imported into Data Files: an .xlsx (or .xlsm) read by the
// Worksheets engine itself, an .ods by office-formats' OpenDocument reader,
// as Worksheets opens them. Each sheet is the model the old files' sheets
// are (database/src/import.js importSheets): its rows of values, each a
// number, words, true or false, and the number format of each cell, so a
// column formatted as a date imports as dates and one formatted as money as
// Currency. A formula's answer is imported, not the formula.

import { SheetView } from '@rutba/sheet-view';
import { readOdf } from '@rutba/office-formats/odf';
import { makeRef } from '@rutba/ooxml/workbook';

/** The workbooks Data Files imports, by their extension. */
export const WORKBOOK_EXTENSIONS = ['xlsx', 'xlsm', 'ods'];

const EPOCH = Date.UTC(1899, 11, 30);
/** An ODS date (2025-04-12, or with a time) as the serial a sheet keeps, so it imports as a sheet's date does. */
const serialOf = (iso) => {
  const t = Date.parse(/T/.test(iso) ? `${iso}Z` : `${iso}T00:00:00Z`);
  return Number.isFinite(t) ? (t - EPOCH) / 86400000 : null;
};

/** An .xlsx's sheets: each one's values as the grid works them out, and each cell's number format. */
function xlsxSheets(bytes) {
  // The package reader takes a Buffer: the bytes as read from the file are one, a copy of them may not be.
  const view = SheetView.open(Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength), { viewportWidth: 800, viewportHeight: 600 });
  return view.sheetNames().map((name) => {
    view.selectSheet(name);
    const { maxRow, maxCol } = view.calc.usedBounds(name);
    const rows = [];
    const styles = {};
    for (let r = 0; r <= maxRow; r++) {
      const row = [];
      for (let c = 0; c <= maxCol; c++) {
        const v = view.calc.getValue(name, r, c);
        // An error (#DIV/0!) or anything that is not a plain value imports as its words.
        row.push(v === null || v === undefined || v === '' ? null : typeof v === 'number' || typeof v === 'boolean' || typeof v === 'string' ? v : String(v?.text ?? v));
        const fmt = view.formatFor(r, c);
        if (fmt && fmt !== 'General' && r > 0) styles[makeRef(r, c)] = { numFmt: fmt };
      }
      rows.push(row);
    }
    return { name, rows, styles };
  });
}

/** An .ods's sheets: its cells' values by their own types, and the formats its data styles give them. */
function odsSheets(bytes) {
  const odf = readOdf(bytes);
  if (odf.flavour !== 'ods') throw new Error('That OpenDocument file is not a spreadsheet.');
  return (odf.sheets || []).map((s, i) => {
    const styles = {};
    for (const [ref, numFmt] of Object.entries(s.formats || {})) styles[ref] = { numFmt };
    const rows = (s.rows || []).map((row, r) => row.map((cell, c) => {
      if (!cell || cell.covered) return null;
      if (cell.type === 'float' || cell.type === 'percentage' || cell.type === 'currency') {
        if (cell.type === 'currency' && r > 0 && !styles[makeRef(r, c)]) styles[makeRef(r, c)] = { numFmt: '£#,##0.00' };
        const n = Number(cell.value ?? cell.text);
        return Number.isFinite(n) ? n : null;
      }
      if (cell.type === 'date' && cell.value) {
        if (r > 0) styles[makeRef(r, c)] = { numFmt: /T\d/.test(String(cell.value)) && !/T00:00(:00)?$/.test(String(cell.value)) ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd' };
        return serialOf(String(cell.value)) ?? cell.text ?? null;
      }
      if (cell.type === 'boolean') return String(cell.value ?? cell.text).toLowerCase() === 'true';
      return cell.text === undefined || cell.text === '' ? null : String(cell.text);
    }));
    return { name: s.name || `Sheet${i + 1}`, rows, styles };
  });
}

/**
 * A workbook's sheets for importSheets: [{ name, rows, styles }], the empty
 * ones left out; `only` names the sheet wanted, or null for every one.
 */
export function readWorkbook(bytes, ext, { only = null } = {}) {
  const sheets = (String(ext).replace(/^\./, '').toLowerCase() === 'ods' ? odsSheets(bytes) : xlsxSheets(bytes))
    .filter((s) => s.rows.some((r) => r.some((v) => v !== null && v !== '')));
  if (only == null) return sheets;
  const one = sheets.filter((s) => s.name === only);
  if (!one.length) throw new Error(`The workbook has no sheet called ${only} with anything in it.`);
  return one;
}
