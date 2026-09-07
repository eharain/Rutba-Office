/**
 * Sheet geometry and virtualisation.
 *
 * A worksheet is 1,048,576 rows by 16,384 columns. Rendering it is not a
 * question of drawing quickly — it is a question of never drawing more than the
 * fifty or so rows a person can actually see. Everything here exists so the view
 * can ask "what is on screen, and where exactly" without walking a million rows.
 *
 * Column widths in OOXML are in CHARACTER units, not pixels: a width of 8.43
 * means "about 8.43 zeroes of the default font fit here". The conversion the
 * format documents is
 *
 *     pixels = trunc(((256 * width + trunc(128 / MDW)) / 256) * MDW)
 *
 * where MDW is the maximum digit width of the workbook's default font — 7px for
 * Calibri 11, which is what almost every real file uses. Getting this wrong by a
 * pixel per column visibly misaligns a wide sheet, so it is reproduced exactly
 * rather than approximated.
 *
 * Offsets are computed rather than accumulated. A prefix-sum array over a
 * million rows would cost 8MB per sheet to save a multiplication.
 */

export const DEFAULT_MAX_DIGIT_WIDTH = 7; // Calibri 11
/**
 * The stored default, which is NOT the 8.43 Excel shows in its UI.
 *
 * The `width` attribute already includes the 5px cell padding, so the number in
 * the file for a default column is 9.140625 — and 9.140625 * 7 is the familiar
 * 64 pixels. Using the UI's 8.43 here yields 59px and misaligns every default
 * column by five pixels.
 */
export const DEFAULT_COL_WIDTH_CHARS = 9.140625;
export const DEFAULT_ROW_HEIGHT_POINTS = 15;
export const POINTS_TO_PIXELS = 4 / 3;
export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

/** OOXML character width -> pixels, per the format's own formula. */
export function charWidthToPixels(width, mdw = DEFAULT_MAX_DIGIT_WIDTH) {
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
}
export const pointsToPixels = (pt) => Math.round(pt * POINTS_TO_PIXELS);

/**
 * Pixels -> character width, the exact inverse of the read above.
 *
 * The trick is to pick a width that is a multiple of 1/256: the format's
 * formula multiplies by 256 first, so a dyadic width survives the arithmetic
 * with no floating error, and `k = ceil(256·px / mdw) − trunc(128 / mdw)` is
 * the smallest such width whose truncated pixel value is exactly `px`. A
 * multiple of 1/256 also prints as a short exact decimal (9.140625), so the
 * number written into the file reopens as the same number.
 */
export function pixelsToCharWidth(px, mdw = DEFAULT_MAX_DIGIT_WIDTH) {
  const k = Math.ceil((256 * px) / mdw) - Math.trunc(128 / mdw);
  return Math.max(0, k) / 256;
}

/**
 * Pixels -> points, the inverse of `pointsToPixels`. `px · 3/4` is exact in
 * binary floating point and `round(px · 3/4 · 4/3)` is `px` again, so a row
 * height round-trips through the file without drifting.
 */
export const pixelsToPoints = (px) => (px * 3) / 4;

/** The smallest sizes a drag or an action may set. Below these, a column or
 *  row is effectively invisible and impossible to grab again. */
export const MIN_COL_WIDTH_PX = 16;
export const MIN_ROW_HEIGHT_PX = 12;

/**
 * Column and row sizing for one sheet.
 *
 * Overrides are sparse — a real sheet customises a handful of columns and leaves
 * the rest at the default — so they live in a Map and the common case costs one
 * multiplication.
 */
export class SheetGeometry {
  constructor({
    defaultColWidth = charWidthToPixels(DEFAULT_COL_WIDTH_CHARS),
    defaultRowHeight = pointsToPixels(DEFAULT_ROW_HEIGHT_POINTS),
    headerWidth = 46,
    headerHeight = 24,
  } = {}) {
    this.defaultColWidth = defaultColWidth;
    this.defaultRowHeight = defaultRowHeight;
    this.headerWidth = headerWidth;
    this.headerHeight = headerHeight;
    /** @type {Map<number, number>} column index -> pixels */
    this.colWidths = new Map();
    /** @type {Map<number, number>} row index -> pixels */
    this.rowHeights = new Map();
    this.hiddenCols = new Set();
    this.hiddenRows = new Set();
  }

  /** Read `<cols>` and row `ht` out of a sheet part's XML. */
  static fromSheetXml(xml, opts = {}) {
    const g = new SheetGeometry(opts);
    const colsBlock = /<cols\b[^>]*>([\s\S]*?)<\/cols>/.exec(xml);
    if (colsBlock) {
      for (const m of colsBlock[1].matchAll(/<col\b([^>]*)\/>/g)) {
        const a = m[1];
        const min = Number(/\bmin="(\d+)"/.exec(a)?.[1] ?? 0);
        const max = Number(/\bmax="(\d+)"/.exec(a)?.[1] ?? 0);
        const width = /\bwidth="([\d.]+)"/.exec(a)?.[1];
        const hidden = /\bhidden="(1|true)"/.test(a);
        if (!min || !max) continue;
        // A `<col>` spans a range; a file that styles every column writes
        // min="1" max="16384", so clamp before expanding or this loops a million
        // times to store one value.
        const upper = Math.min(max, min + 512);
        for (let c = min - 1; c < upper; c++) {
          if (width !== undefined) g.colWidths.set(c, charWidthToPixels(Number(width), opts.mdw));
          if (hidden) g.hiddenCols.add(c);
        }
      }
    }
    for (const m of xml.matchAll(/<row\b([^>]*?)(?:\/>|>)/g)) {
      const a = m[1];
      const r = Number(/\br="(\d+)"/.exec(a)?.[1] ?? 0);
      if (!r) continue;
      const ht = /\bht="([\d.]+)"/.exec(a)?.[1];
      if (ht !== undefined) g.rowHeights.set(r - 1, pointsToPixels(Number(ht)));
      if (/\bhidden="(1|true)"/.test(a)) g.hiddenRows.add(r - 1);
    }
    return g;
  }

  colWidth(index) {
    if (this.hiddenCols.has(index)) return 0;
    return this.colWidths.get(index) ?? this.defaultColWidth;
  }
  rowHeight(index) {
    if (this.hiddenRows.has(index)) return 0;
    return this.rowHeights.get(index) ?? this.defaultRowHeight;
  }

  /** Pixel offset of a column's left edge, relative to the grid origin. */
  colOffset(index) {
    let offset = index * this.defaultColWidth;
    for (const [c, w] of this.colWidths) if (c < index) offset += w - this.defaultColWidth;
    for (const c of this.hiddenCols) {
      if (c < index && !this.colWidths.has(c)) offset -= this.defaultColWidth;
    }
    return offset;
  }
  rowOffset(index) {
    let offset = index * this.defaultRowHeight;
    for (const [r, h] of this.rowHeights) if (r < index) offset += h - this.defaultRowHeight;
    for (const r of this.hiddenRows) {
      if (r < index && !this.rowHeights.has(r)) offset -= this.defaultRowHeight;
    }
    return offset;
  }

  /** Which column contains this x offset. Binary search over computed offsets. */
  colAt(x) {
    let lo = 0;
    let hi = MAX_COLS - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.colOffset(mid) + this.colWidth(mid) <= x) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }
  rowAt(y) {
    let lo = 0;
    let hi = MAX_ROWS - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.rowOffset(mid) + this.rowHeight(mid) <= y) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  /**
   * The window of cells to render for a given scroll position.
   *
   * `overscan` renders a little outside the viewport so a fast scroll does not
   * show blank rows before the next frame.
   *
   * @returns {{firstRow, lastRow, firstCol, lastCol, offsetX, offsetY}}
   */
  viewport({ scrollX = 0, scrollY = 0, width, height, overscan = 3 }) {
    const firstCol = Math.max(0, this.colAt(scrollX) - overscan);
    const firstRow = Math.max(0, this.rowAt(scrollY) - overscan);

    let lastCol = firstCol;
    let x = this.colOffset(firstCol);
    const rightEdge = scrollX + width;
    while (lastCol < MAX_COLS - 1 && x < rightEdge) {
      x += this.colWidth(lastCol);
      lastCol += 1;
    }
    lastCol = Math.min(MAX_COLS - 1, lastCol + overscan);

    let lastRow = firstRow;
    let y = this.rowOffset(firstRow);
    const bottomEdge = scrollY + height;
    while (lastRow < MAX_ROWS - 1 && y < bottomEdge) {
      y += this.rowHeight(lastRow);
      lastRow += 1;
    }
    lastRow = Math.min(MAX_ROWS - 1, lastRow + overscan);

    return {
      firstRow, lastRow, firstCol, lastCol,
      offsetX: this.colOffset(firstCol),
      offsetY: this.rowOffset(firstRow),
      rows: lastRow - firstRow + 1,
      cols: lastCol - firstCol + 1,
    };
  }

  /** Total scrollable size, bounded by the used range so scrollbars are sane. */
  totalSize({ maxRow, maxCol }) {
    const rows = Math.min(MAX_ROWS - 1, maxRow + 40);
    const cols = Math.min(MAX_COLS - 1, maxCol + 12);
    return {
      width: this.colOffset(cols) + this.colWidth(cols),
      height: this.rowOffset(rows) + this.rowHeight(rows),
    };
  }

  /** Scroll delta needed to bring a cell fully into view. */
  scrollToShow({ row, col, scrollX, scrollY, width, height }) {
    const cellX = this.colOffset(col);
    const cellY = this.rowOffset(row);
    const cellW = this.colWidth(col);
    const cellH = this.rowHeight(row);

    let nextX = scrollX;
    if (cellX < scrollX) nextX = cellX;
    else if (cellX + cellW > scrollX + width) nextX = cellX + cellW - width;

    let nextY = scrollY;
    if (cellY < scrollY) nextY = cellY;
    else if (cellY + cellH > scrollY + height) nextY = cellY + cellH - height;

    return { scrollX: Math.max(0, nextX), scrollY: Math.max(0, nextY) };
  }
}
