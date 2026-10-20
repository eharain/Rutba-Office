// Access databases (.mdb, Jet 3 and 4; .accdb, ACE), read only: a sheet per
// table, its field names as the first row, a row a record. The tables are
// read by mdb-reader (MIT); forms, reports, macros and VBA are never read,
// and pictures, attachments and other binary fields are left out and said.
// Read into the old spreadsheet readers' sheets (legacy-grid.js).

import { Grid } from './legacy-grid.js';

/** Whether bytes are an Access database: "Standard Jet DB" or "Standard ACE DB" after the first page's header. */
export function isAccess(bytes) {
  if (bytes.length < 32) return false;
  const tag = String.fromCharCode(...bytes.subarray(4, 19));
  return tag === 'Standard Jet DB' || tag === 'Standard ACE DB';
}

const BINARY = new Set(['binary', 'ole', 'complex']);

/**
 * Read an Access database through `Reader` (mdb-reader's class): the old
 * spreadsheet readers' model, a sheet per table. `password` opens one that
 * asks for it.
 */
export function readAccess(bytes, { Reader, password = undefined, rowLimit = 1048575 } = {}) {
  if (!Reader) throw new Error('no reader for Access databases in this build');
  let reader;
  try {
    reader = new Reader(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength), password ? { password } : undefined);
  } catch (err) {
    if (/password/i.test(String(err?.message))) {
      const e = new Error('This database is protected by a password.');
      e.password = true;
      throw e;
    }
    throw err;
  }
  const sheets = [];
  const notes = [];
  let leftOut = 0;
  let cut = 0;
  for (const name of reader.getTableNames()) {
    const table = reader.getTable(name);
    const columns = table.getColumns();
    const grid = new Grid(String(name).slice(0, 31));
    const kept = columns.filter((c) => !BINARY.has(c.type));
    leftOut += columns.length - kept.length;
    kept.forEach((c, k) => {
      grid.set(0, k, c.name, { bold: true });
      grid.width(k, Math.min(40, Math.max(c.name.length + 2, c.type === 'memo' ? 40 : c.type === 'text' ? Math.min(40, (c.size || 20) / 2 + 2) : 12)));
    });
    const rows = table.getData({ columns: kept.map((c) => c.name), rowLimit: Math.min(rowLimit, table.rowCount) });
    if (table.rowCount > rowLimit) cut++;
    rows.forEach((record, r) => {
      kept.forEach((c, k) => {
        let value = record[c.name];
        let style = null;
        if (value == null) return;
        if (value instanceof Date) {
          // mdb-reader gives Access's day count as a Date on the UTC clock: back to the day count it was.
          style = { numFmt: value.getUTCHours() || value.getUTCMinutes() || value.getUTCSeconds() ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd' };
          value = (value.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
        } else if (typeof value === 'bigint') {
          value = Number(value);
        } else if (c.type === 'currency' || c.type === 'numeric') {
          const n = Number(value);
          if (Number.isFinite(n)) { value = n; if (c.type === 'currency') style = { numFmt: '#,##0.00' }; }
        }
        grid.set(r + 1, k, value, style);
      });
    });
    sheets.push(grid.toSheet());
  }
  if (!sheets.length) { sheets.push(new Grid('Sheet1').toSheet()); notes.push({ code: 'no-tables' }); }
  if (leftOut) notes.push({ code: 'binary-left-out', count: leftOut });
  if (cut) notes.push({ code: 'rows-cut', count: cut });
  notes.push({ code: 'access-read-only' });
  return { sheets, names: [], notes };
}

