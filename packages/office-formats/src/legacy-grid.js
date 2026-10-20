// What the old spreadsheet readers (SYLK, DIF, Lotus 1-2-3, dBASE) make:
// sheets of cells as @rutba/ooxml's buildXlsx takes them, so each reader is
// only the format it reads and the .xlsx is written in one place.
//
//   { sheets: [{ name, rows, styles, widths, merges }], names: [{ name, ref }],
//     notes: [string] }
//
// `rows` is rows of cells: a number, a string, a boolean, null, or a formula
// as a string starting with '='. `styles` maps "A1" (or "A1:C3") to a look
// buildXlsx knows ({ numFmt, bold, align, ... }); `widths` maps a column's
// letters to its width in characters. `notes` says what a reader met and
// could not keep, for the window to tell the person.

/** A column's letters, 0 being A. */
export const colName = (c) => {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

/** A cell's reference, `$` before each part held absolute. */
export const cellRef = (row, col, { absRow = false, absCol = false } = {}) => `${absCol ? '$' : ''}${colName(col)}${absRow ? '$' : ''}${row + 1}`;

/** A sheet being filled cell by cell, in any order. */
export class Grid {
  constructor(name) {
    this.name = name;
    this.cells = [];
    this.styles = {};
    this.widths = {};
    this.merges = [];
  }

  set(row, col, value, style = null) {
    if (!(row >= 0 && col >= 0) || row > 1048575 || col > 16383) return;
    while (this.cells.length <= row) this.cells.push(null);
    const line = this.cells[row] || (this.cells[row] = []);
    while (line.length <= col) line.push(null);
    line[col] = value;
    if (style && Object.keys(style).length) this.styles[cellRef(row, col)] = style;
  }

  get(row, col) {
    return this.cells[row]?.[col] ?? null;
  }

  /** A column's width, in characters. */
  width(col, chars) {
    if (chars > 0) this.widths[colName(col)] = Math.min(255, chars);
  }

  /** The sheet as buildXlsx takes it. */
  toSheet() {
    const rows = this.cells.map((line) => line || []);
    return { name: this.name, rows, styles: this.styles, widths: this.widths, merges: this.merges };
  }
}
