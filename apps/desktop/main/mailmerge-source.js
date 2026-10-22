// Mailings: the recipient lists a mail merge reads.
//
// Select Recipients → Use an Existing List opens a workbook, a .csv or a
// .tsv; Type a New List saves one. A workbook is read by the Worksheets
// engine itself, sheet by sheet, and each cell is taken as the grid shows it
// — a date as a date, a price with its two places — which is what Word's
// own merge from Excel is always being asked to do. A delimited file is
// read with its byte-order mark gone and its quoted commas kept. The first
// row names the fields, as it does in Word. A Data Files database offers its
// tables and saved queries as Word offers an Access file's, read and never
// changed: its fields are the merge's, an amount with its two places.

import fs from 'node:fs';
import path from 'node:path';
import { SheetView } from '@rutba/sheet-view';
import { decodeText } from '@rutba/office-formats/text';
import { parseDelimited, sourceFromRows } from '@rutba/ooxml/mailmerge';
import { Base } from '@rutba/database';

const WORKBOOKS = new Set(['.xlsx', '.xlsm', '.xltx']);
const DELIMITED = { '.csv': null, '.tsv': '\t', '.tab': '\t', '.txt': null };

/** What a merge can read, for the Select Data Source dialog's filter. */
export const SOURCE_EXTENSIONS = ['xlsx', 'xlsm', 'csv', 'tsv', 'txt', 'rbase'];

/** A stored value as a merge field gives it: a yes/no as Yes or No, an amount with its two places. */
function mergeValue(type, v) {
  if (v === null || v === undefined) return '';
  if (type === 'yesno') return v ? 'Yes' : 'No';
  if (type === 'currency' && Number.isFinite(Number(v))) return Number(v).toFixed(2);
  return String(v);
}

/** A Data Files database's table or saved query as a recipient list, `sheet` naming which. */
function readDatabase(filePath, name, sheet) {
  let base;
  try {
    base = Base.read(filePath);
  } catch (err) {
    throw new Error(`${name} could not be read as a database: ${err.message}`);
  }
  try {
    // A protected database (D-2026-10-22-11) gives nothing to a reader that has not signed in.
    if (base.access?.locked) throw new Error(`${name} is protected. Open it in Data Files and sign in, or ask its Admin for a copy of what you need.`);
    const tables = base.hasDesignTable('rutba_tables') ? base.tables().map((x) => x.name) : [];
    const queries = base.hasDesignTable('rutba_queries') ? base.queries().map((x) => x.name).filter((n) => !tables.includes(n)) : [];
    const sheets = [...tables, ...queries];
    if (!sheets.length) throw new Error(`${name} has no tables to take recipients from.`);
    const chosen = sheet && sheets.includes(sheet) ? sheet : sheets[0];
    let rows;
    if (tables.includes(chosen)) {
      const fields = base.fields(chosen);
      const records = base.rows(chosen, { limit: 1_000_000 });
      rows = [fields.map((f) => f.name), ...records.map((r) => fields.map((f) => mergeValue(f.type, r[f.name])))];
    } else {
      const out = base.runQuery({ name: chosen, limit: 1_000_000 });
      rows = [out.columns, ...out.rows.map((r) => out.columns.map((c) => mergeValue(null, r[c])))];
    }
    return sourceFromRows(rows, { kind: 'rbase', name, path: filePath, sheet: chosen, sheets });
  } finally {
    base.close();
  }
}

/**
 * A recipient list from a file: `{ kind, name, path, sheet, sheets, fields,
 * records }`. `sheet` picks a workbook's worksheet by name; the first one
 * otherwise, as Word's Select Table dialog starts on the first.
 */
export function readMergeSource(filePath, { sheet = null } = {}) {
  const ext = path.extname(filePath).toLowerCase();
  const name = path.basename(filePath);
  let bytes;
  try {
    bytes = fs.readFileSync(filePath);
  } catch (err) {
    throw new Error(err.code === 'ENOENT' ? `${name} is not there any more — it was moved, renamed or deleted.` : `${name} could not be read: ${err.message}`);
  }
  if (WORKBOOKS.has(ext)) {
    const view = SheetView.open(bytes, { viewportWidth: 800, viewportHeight: 600 });
    const sheets = view.sheetNames();
    const chosen = sheet && sheets.includes(sheet) ? sheet : sheets[0];
    view.selectSheet(chosen);
    const { maxRow, maxCol } = view.calc.usedBounds(chosen);
    const rows = [];
    for (let r = 0; r <= maxRow; r++) {
      const row = [];
      for (let c = 0; c <= maxCol; c++) row.push(String(view.displayValue(r, c).text ?? ''));
      rows.push(row);
    }
    // A column with no heading and nothing under it is not a field.
    const used = rows[0] ? rows[0].map((_, c) => rows.some((row) => String(row[c] ?? '').trim() !== '')) : [];
    const trimmed = rows.map((row) => row.filter((_, c) => used[c]));
    return sourceFromRows(trimmed, { kind: 'xlsx', name, path: filePath, sheet: chosen, sheets });
  }
  if (ext === '.rbase') return readDatabase(filePath, name, sheet);
  if (ext in DELIMITED) {
    const rows = parseDelimited(decodeText(bytes), DELIMITED[ext]);
    return sourceFromRows(rows, { kind: ext === '.tsv' || ext === '.tab' ? 'tsv' : 'csv', name, path: filePath, sheet: null, sheets: null });
  }
  throw new Error(`${name} is not a list a merge reads. Choose a workbook (.xlsx), a Data Files database (.rbase) or a .csv or .tsv file.`);
}

/** A field as a CSV cell: quoted only when it has to be. */
const cell = (v) => {
  const s = v == null ? '' : String(v);
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
};

/**
 * Type a New List, saved: a UTF-8 .csv with the byte-order mark Excel looks
 * for (without it Excel reads "Müller" as "MÃ¼ller"), CRLF line ends.
 */
export function writeMergeList(filePath, fields, rows) {
  const lines = [fields, ...(rows || [])].map((r) => fields.map((_, i) => cell(r[i])).join(','));
  fs.writeFileSync(filePath, '\uFEFF' + lines.join('\r\n') + '\r\n', 'utf8');
  return filePath;
}
