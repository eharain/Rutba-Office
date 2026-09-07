/**
 * Selection and keyboard navigation.
 *
 * Small, and worth getting exactly right, because it is the part users feel
 * every second. The behaviours below are the ones every spreadsheet has shared
 * for decades; they are conventions, not anyone's design:
 *
 *   - An anchor and an extent define a rectangle. Shift+arrow moves the extent
 *     and leaves the anchor; a plain arrow collapses to a single cell.
 *   - Ctrl+arrow jumps to the edge of the current block of data — the next
 *     boundary between filled and empty. From a filled cell next to a filled
 *     neighbour, it runs to the end of the run; otherwise it skips the blanks
 *     and lands on the next filled cell. That two-case rule is the whole
 *     behaviour and it is what makes large sheets navigable.
 *   - Tab and Enter move WITHIN a multi-cell selection and wrap inside it,
 *     rather than escaping it. That is what makes typing into a selected block
 *     work. With a single cell selected they just move.
 *   - Enter returns to the column Tab started from. Tabbing across a row then
 *     pressing Enter drops to the start of the next row, not wherever Tab left
 *     the cursor.
 */

/**
 * THREE cells, not two, and the distinction is the bug most implementations
 * ship with:
 *
 *   anchor  the fixed corner a shift-selection grows from
 *   extent  the moving corner — where shift+arrow goes
 *   active  the cell you are actually typing into
 *
 * Drag A1 to B2 and the range is A1:B2, the extent is B2, but the **active cell
 * stays at A1**. Tab then walks A1 → B1 → A2 → B2 → A1 inside the block. Treating
 * the extent as the active cell means typing after a drag lands in the wrong
 * corner, which is the kind of thing nobody reports as a bug — they just stop
 * trusting the grid.
 */
export class Selection {
  constructor({ row = 0, col = 0 } = {}) {
    this.anchor = { row, col };
    this.extent = { row, col };
    this.active = { row, col };
    /** Column that Enter returns to after a run of Tabs. */
    this.entryCol = col;
    this.entryRow = row;
  }

  static at(row, col) { return new Selection({ row, col }); }

  /** Normalised rectangle covering anchor and extent. */
  get range() {
    return {
      top: Math.min(this.anchor.row, this.extent.row),
      bottom: Math.max(this.anchor.row, this.extent.row),
      left: Math.min(this.anchor.col, this.extent.col),
      right: Math.max(this.anchor.col, this.extent.col),
    };
  }

  get isSingle() {
    return this.anchor.row === this.extent.row && this.anchor.col === this.extent.col;
  }

  contains(row, col) {
    const r = this.range;
    return row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;
  }

  collapseTo(row, col) {
    this.anchor = { row, col };
    this.extent = { row, col };
    this.active = { row, col };
    this.entryCol = col;
    this.entryRow = row;
    return this;
  }

  /** Grow the selection. The active cell does NOT follow — it stays at the anchor. */
  extendTo(row, col) {
    this.extent = { row, col };
    if (!this.contains(this.active.row, this.active.col)) this.active = { ...this.anchor };
    return this;
  }

  /**
   * @param {'up'|'down'|'left'|'right'} direction
   * @param {object} opts
   * @param {boolean} [opts.extend]  shift held
   * @param {boolean} [opts.jump]    ctrl held — run to the edge of the data
   * @param {(row:number,col:number)=>boolean} [opts.isFilled]
   * @param {{maxRow:number,maxCol:number}} [opts.bounds]
   */
  move(direction, { extend = false, jump = false, isFilled, bounds } = {}) {
    // shift+arrow grows from the extent; a plain arrow walks from the active cell
    const from = extend ? this.extent : this.active;
    const step = { up: [-1, 0], down: [1, 0], left: [0, -1], right: [0, 1] }[direction];
    if (!step) throw new Error('unknown direction: ' + direction);

    const limit = {
      maxRow: bounds?.maxRow ?? 1048575,
      maxCol: bounds?.maxCol ?? 16383,
    };
    const clampRow = (r) => Math.max(0, Math.min(limit.maxRow, r));
    const clampCol = (c) => Math.max(0, Math.min(limit.maxCol, c));

    let target;
    if (jump && isFilled) {
      target = jumpToEdge(from, step, isFilled, limit);
    } else {
      target = { row: clampRow(from.row + step[0]), col: clampCol(from.col + step[1]) };
    }

    if (extend) this.extendTo(target.row, target.col);
    else this.collapseTo(target.row, target.col);
    return this;
  }

  /**
   * Tab / Shift+Tab. Inside a multi-cell selection this wraps within the block;
   * otherwise it moves and drags the selection with it.
   */
  tab(back = false, { bounds } = {}) {
    if (this.isSingle) {
      const col = Math.max(0, Math.min(bounds?.maxCol ?? 16383, this.active.col + (back ? -1 : 1)));
      this.active = { row: this.active.row, col };
      this.anchor = { ...this.active };
      this.extent = { ...this.active };
      // entryCol deliberately NOT updated: Enter must return to where the run began
      this.entryRow = this.active.row;
      return this;
    }
    const r = this.range;
    const width = r.right - r.left + 1;
    const height = r.bottom - r.top + 1;
    let index = (this.active.row - r.top) * width + (this.active.col - r.left);
    index = (index + (back ? -1 : 1) + width * height) % (width * height);
    this.active = { row: r.top + Math.floor(index / width), col: r.left + (index % width) };
    return this;
  }

  /** Enter / Shift+Enter. Returns to the column the Tab run started from. */
  enter(back = false, { bounds } = {}) {
    if (this.isSingle) {
      const row = Math.max(0, Math.min(bounds?.maxRow ?? 1048575, this.active.row + (back ? -1 : 1)));
      this.active = { row, col: this.entryCol };
      this.anchor = { ...this.active };
      this.extent = { ...this.active };
      this.entryRow = row;
      return this;
    }
    const r = this.range;
    const width = r.right - r.left + 1;
    const height = r.bottom - r.top + 1;
    let index = (this.active.col - r.left) * height + (this.active.row - r.top);
    index = (index + (back ? -1 : 1) + width * height) % (width * height);
    this.active = { row: r.top + (index % height), col: r.left + Math.floor(index / height) };
    return this;
  }

  /** Every cell in the selection, row-major. Guarded against enormous ranges. */
  *cells({ limit = 100000 } = {}) {
    const r = this.range;
    let n = 0;
    for (let row = r.top; row <= r.bottom; row++) {
      for (let col = r.left; col <= r.right; col++) {
        if (++n > limit) return;
        yield { row, col };
      }
    }
  }

  clone() {
    const s = new Selection(this.anchor);
    s.extent = { ...this.extent };
    s.active = { ...this.active };
    s.entryCol = this.entryCol;
    s.entryRow = this.entryRow;
    return s;
  }

  toString() {
    const r = this.range;
    return r.top === r.bottom && r.left === r.right
      ? ref(r.top, r.left)
      : ref(r.top, r.left) + ':' + ref(r.bottom, r.right);
  }
}

/**
 * Ctrl+arrow. Two cases, and the distinction is the whole behaviour:
 *   from a filled cell whose neighbour is filled -> run to the end of that block
 *   otherwise                                     -> skip blanks to the next filled cell
 */
function jumpToEdge(from, step, isFilled, limit) {
  const inBounds = (r, c) => r >= 0 && c >= 0 && r <= limit.maxRow && c <= limit.maxCol;
  let { row, col } = from;
  const nextRow = row + step[0];
  const nextCol = col + step[1];
  if (!inBounds(nextRow, nextCol)) return { row, col };

  if (isFilled(row, col) && isFilled(nextRow, nextCol)) {
    // walk to the last filled cell of this run
    while (inBounds(row + step[0], col + step[1]) && isFilled(row + step[0], col + step[1])) {
      row += step[0];
      col += step[1];
    }
    return { row, col };
  }

  // skip blanks; land on the next filled cell, or the sheet edge
  let r = row + step[0];
  let c = col + step[1];
  while (inBounds(r, c)) {
    if (isFilled(r, c)) return { row: r, col: c };
    r += step[0];
    c += step[1];
  }
  return {
    row: Math.max(0, Math.min(limit.maxRow, r - step[0])),
    col: Math.max(0, Math.min(limit.maxCol, c - step[1])),
  };
}

export function colName(n) {
  let s = '';
  let v = n + 1;
  while (v > 0) {
    const rem = (v - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    v = Math.floor((v - 1) / 26);
  }
  return s;
}
export const ref = (row, col) => colName(col) + (row + 1);
