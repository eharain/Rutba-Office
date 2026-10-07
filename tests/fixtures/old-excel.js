/**
 * Excel 2.1, 3.0 and 4.0 worksheets, laid out record by record from their
 * published layouts (LibreOffice's Excel import), since no Excel on hand
 * writes them any more. A worksheet of the description's cells, in the
 * version's own records: its fonts, number formats and cell formats, the
 * cells (numbers, text, a truth value, formulas with their last values),
 * a column width, a row height, a frozen pane, a picture.
 */

const le16 = (v) => [v & 0xff, (v >> 8) & 0xff];
const le32 = (v) => [...le16(v & 0xffff), ...le16(v >>> 16)];
const f64 = (v) => { const b = new DataView(new ArrayBuffer(8)); b.setFloat64(0, v, true); return [...new Uint8Array(b.buffer)]; };
const bytes = (s) => [...s].map((c) => c.charCodeAt(0) & 0xff);
const rec = (id, data) => [...le16(id), ...le16(data.length), ...data];

/** A reference to a cell in a formula, relative both ways: BIFF2-5's three bytes. */
export const refTok = (row, col, cls = 0x44) => [cls, ...le16(row | 0xc000), col];
/** A range in a formula, relative: six bytes. */
export const areaTok = (r1, r2, c1, c2) => [0x25, ...le16(r1 | 0xc000), ...le16(r2 | 0xc000), c1, c2];
/** A function with its arguments counted: BIFF2 and 3 number it in a byte, BIFF4 in two. */
export const funcVar = (biff, argc, id) => (biff <= 3 ? [0x42, argc, id] : [0x42, argc, ...le16(id)]);

/**
 * @param {2|3|4} biff
 * @param {{ cells: Array<{ row, col, value?, text?, bool?, formula?: number[], result?: number, xf?: number }>,
 *           fonts: Array<{ name, height, bold?, italic? }>, formats: string[], colWidth?, rowHeight?, frozen?,
 *           picture?: { from: { col, row }, to: { col, row }, width, height, colour: [r, g, b] } }} sheet
 */
export function buildOldExcel(biff, { cells, fonts, formats, colWidth = null, rowHeight = null, frozen = null, picture = null }) {
  const out = [];
  const push = (id, data) => out.push(...rec(id, data));
  // BOF: the version, a worksheet.
  push({ 2: 0x0009, 3: 0x0209, 4: 0x0409 }[biff], biff === 2 ? [...le16(2), ...le16(0x10)] : [...le16(0), ...le16(0x10), ...le16(0)]);
  push(0x0042, le16(1252));
  for (const f of fonts) {
    const flags = (f.bold ? 1 : 0) | (f.italic ? 2 : 0);
    if (biff === 2) {
      push(0x0031, [...le16(f.height), ...le16(flags), f.name.length, ...bytes(f.name)]);
      push(0x0045, le16(0x7fff));
    } else push(0x0231, [...le16(f.height), ...le16(flags), ...le16(0x7fff), f.name.length, ...bytes(f.name)]);
  }
  for (const code of formats) {
    if (biff === 4) push(0x041e, [0, 0, code.length, ...bytes(code)]);
    else push(0x001e, [code.length, ...bytes(code)]);
  }
  // Cell formats: the default, then one in the second font, then one in the second number format.
  const xfs = [{ font: 0, fmt: 0 }, { font: 1, fmt: 0 }, { font: 0, fmt: 1 }];
  for (const x of xfs) {
    if (biff === 2) push(0x0043, [x.font, 0, x.fmt, 0]);
    else push(biff === 3 ? 0x0243 : 0x0443, [x.font, x.fmt, ...le16(0x0001), ...le16(0), ...le16(0), 0, 0, 0, 0]);
  }
  if (colWidth) {
    if (biff === 2) push(0x0024, [colWidth.col, colWidth.col, ...le16(colWidth.width)]);
    else push(0x007d, [...le16(colWidth.col), ...le16(colWidth.col), ...le16(colWidth.width), ...le16(0), ...le16(0), 0, 0]);
  }
  if (rowHeight) {
    const r = biff === 2 ? [...le16(rowHeight.row), 0, 0, ...le16(4), ...le16(rowHeight.height), 0, 0, 0, 0, 0, 0, 0, 0]
      : [...le16(rowHeight.row), ...le16(0), ...le16(4), ...le16(rowHeight.height), ...le16(0), ...le16(0), ...le16(0x40), ...le16(0)];
    push(biff === 2 ? 0x0008 : 0x0208, r);
  }
  // The cells: BIFF2's three attribute bytes (the XF first), or BIFF3 and 4's two-byte XF.
  const head = (c) => (biff === 2 ? [...le16(c.row), ...le16(c.col), c.xf ?? 0, 0, 0] : [...le16(c.row), ...le16(c.col), ...le16(c.xf ?? 0)]);
  for (const c of cells) {
    if (c.formula) {
      const rgce = c.formula;
      if (biff === 2) push(0x0006, [...head(c), ...f64(c.result ?? 0), 0, rgce.length, ...rgce]);
      else push(biff === 3 ? 0x0206 : 0x0406, [...head(c), ...f64(c.result ?? 0), ...le16(0), ...le16(rgce.length), ...rgce]);
    } else if (c.text != null) {
      push(biff === 2 ? 0x0004 : 0x0204, biff === 2 ? [...head(c), c.text.length, ...bytes(c.text)] : [...head(c), ...le16(c.text.length), ...bytes(c.text)]);
    } else if (c.bool != null) {
      push(biff === 2 ? 0x0005 : 0x0205, [...head(c), c.bool ? 1 : 0, 0]);
    } else if (biff === 2 && Number.isInteger(c.value) && c.value >= 0 && c.value < 65536) {
      push(0x0002, [...head(c), ...le16(c.value)]);
    } else {
      push(biff === 2 ? 0x0003 : 0x0203, [...head(c), ...f64(c.value)]);
    }
  }
  if (frozen) {
    push(biff === 2 ? 0x003e : 0x023e, biff === 2 ? [0, 1, 1, 1, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0] : [...le16(0x0002 | 0x0004 | 0x0008 | 0x0010), ...le16(0), ...le16(0), 0, 0, 0, 0]);
    push(0x0041, [...le16(frozen.cols), ...le16(frozen.rows), ...le16(frozen.rows), ...le16(frozen.cols), 0]);
  }
  if (picture && biff >= 3) {
    // A picture: an OBJ with the cells its corners are in, then its DIB in an
    // IMGDATA — 32 bits a pixel and three stray bytes after the 12-byte header,
    // as Excel 3 and 4 wrote one.
    const { from, to, width, height, colour } = picture;
    const corners = [from.col, 0, from.row, 0, to.col, 0, to.row, 0].flatMap(le16);
    push(0x005d, [...le32(1), ...le16(8), ...le16(1), ...le16(0), ...corners, ...le16(0), ...new Array(30).fill(0)]);
    const pixels = [];
    for (let i = 0; i < width * height; i++) pixels.push(colour[2], colour[1], colour[0], 0);
    const dib = [...le32(12), ...le16(width), ...le16(height), ...le16(1), ...le16(32), 0, 0, 0, ...pixels];
    push(0x007f, [...le16(9), ...le16(1), ...le32(dib.length), ...dib]);
  }
  push(0x000a, []);
  return Uint8Array.from(out);
}
