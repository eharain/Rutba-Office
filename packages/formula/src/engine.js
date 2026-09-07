/**
 * The calculation engine: a sheet model plus dependency-ordered recalculation.
 *
 * Three things a spreadsheet must get right, and they are the reason this is a
 * graph rather than a loop over cells:
 *
 * 1. **Order.** If C1 is `=B1*2` and B1 is `=A1+1`, changing A1 must recalculate
 *    B1 before C1. Evaluating in row order gives the right answer only by luck,
 *    and gives the WRONG answer silently when luck runs out.
 * 2. **Cycles.** `A1 = B1`, `B1 = A1` must be reported, not hang. Excel warns and
 *    yields 0; we yield `#CIRCULAR!` on every cell in the cycle, because a
 *    silent 0 in a finance sheet is worse than a visible error.
 * 3. **Minimal work.** Editing one cell should recalculate its dependents, not
 *    the whole workbook. A 50,000-cell sheet recalculated in full on every
 *    keystroke is unusable, and that is a correctness-shaped performance
 *    problem: people work around slow sheets by turning calculation off.
 *
 * Volatile functions (TODAY, NOW) are always dirty, since nothing in the graph
 * tells us the clock moved.
 */
import { parse, dependencies, indexToCol, colToIndex } from './parser.js';
import { evaluate } from './evaluator.js';
import { ERR, isError, isBlank } from './values.js';
import { isVolatile } from './functions.js';

const key = (sheet, row, col) => sheet + '!' + row + ':' + col;

export class Spreadsheet {
  /**
   * @param {object} [opts]
   * @param {() => Date} [opts.now] clock, injectable so tests are deterministic
   */
  constructor({ now = () => new Date() } = {}) {
    /** @type {Map<string, Map<string, object>>} sheet -> cellKey -> cell */
    this.sheets = new Map();
    /** @type {Map<string, {sheet: string|null, start: object, end: object}>} */
    this.names = new Map();
    /** @type {Map<string, object>} UPPER table name -> {sheet, top, left, bottom, right, headerRows, totalsRows, columns, name} */
    this.tables = new Map();
    this.now = now;
    this.dirty = new Set();
    this.volatileCells = new Set();
    /**
     * Spilled ranges, by their ANCHOR's key. A formula whose result is an
     * array writes its top-left value into its own cell and "spills" the
     * rest into empty neighbours — ghosts that display a value but hold no
     * input. A blocked spill (something real in the way) records its
     * would-be coverage with `blocked: true`, so clearing the blocker
     * re-seeds the anchor.
     * @type {Map<string, {sheet: string, top: number, left: number, h: number, w: number, values: any[][], blocked: boolean}>}
     */
    this.spills = new Map();
  }

  addSheet(name) {
    if (!this.sheets.has(name)) this.sheets.set(name, new Map());
    return this;
  }
  sheetNames() { return [...this.sheets.keys()]; }
  _sheet(name) {
    const s = this.sheets.get(name);
    if (!s) throw new Error('no such sheet: ' + name);
    return s;
  }

  /**
   * Set a cell. A leading `=` makes it a formula; anything else is a literal.
   * Parse errors are stored as the cell's value so the user sees the problem in
   * the cell rather than losing what they typed.
   */
  setCell(sheetName, row, col, input) {
    this.addSheet(sheetName);
    const cells = this._sheet(sheetName);
    const k = key(sheetName, row, col);

    if (isBlank(input)) {
      cells.delete(k);
      this.volatileCells.delete(k);
      this.spills.delete(k);
      this._invalidate(sheetName, row, col);
      return this;
    }

    const isFormula = typeof input === 'string' && input.startsWith('=');
    const cell = { sheet: sheetName, row, col, input, formula: null, ast: null, value: null, error: null };

    if (isFormula) {
      cell.formula = input;
      try {
        cell.ast = parse(input);
        cell.deps = dependencies(cell.ast);
        cell.volatile = hasVolatileCall(cell.ast);
      } catch (e) {
        cell.ast = null;
        cell.value = isError(e) ? e : ERR.VALUE(e.message ?? String(e));
        cell.error = cell.value;
      }
    } else {
      cell.value = input;
    }

    cells.set(k, cell);
    if (cell.volatile) this.volatileCells.add(k);
    else this.volatileCells.delete(k);
    this.spills.delete(k);
    this._invalidate(sheetName, row, col);
    return this;
  }

  setByRef(sheetName, a1, input) {
    const m = /^([A-Za-z]+)(\d+)$/.exec(a1);
    if (!m) throw new Error('bad cell reference: ' + a1);
    return this.setCell(sheetName, Number(m[2]) - 1, colToIndex(m[1]), input);
  }

  /** Stored cell record, or null. */
  cell(sheetName, row, col) {
    return this.sheets.get(sheetName)?.get(key(sheetName, row, col)) ?? null;
  }

  /** A spilled value covering a cell that holds nothing of its own. */
  spillValueAt(sheetName, row, col) {
    for (const s of this.spills.values()) {
      if (s.blocked) continue;
      if (s.sheet === sheetName && row >= s.top && row < s.top + s.h && col >= s.left && col < s.left + s.w) {
        return s.values[row - s.top][col - s.left] ?? '';
      }
    }
    return undefined;
  }

  /** Calculated value of a cell — the number a user sees. */
  getValue(sheetName, row, col) {
    if (this.dirty.size) this.recalculate();
    const c = this.cell(sheetName, row, col);
    if (!c) return this.spillValueAt(sheetName, row, col) ?? '';
    return c.value ?? '';
  }

  getByRef(sheetName, a1) {
    const m = /^([A-Za-z]+)(\d+)$/.exec(a1);
    if (!m) throw new Error('bad cell reference: ' + a1);
    return this.getValue(sheetName, Number(m[2]) - 1, colToIndex(m[1]));
  }

  /** What the user typed, for the formula bar. */
  getInput(sheetName, row, col) {
    return this.cell(sheetName, row, col)?.input ?? '';
  }

  defineName(name, target) { this.names.set(name.toUpperCase(), target); return this; }

  /**
   * Register a table (a ListObject) so structured references resolve.
   * @param {string} name
   * @param {{sheet: string, top: number, left: number, bottom: number, right: number,
   *          headerRows?: number, totalsRows?: number, columns?: string[]}} def
   */
  defineTable(name, def) {
    this.tables.set(String(name).toUpperCase(), { name, headerRows: 1, totalsRows: 0, columns: [], ...def });
    return this;
  }

  /** The table containing a cell, or null — what a bare `[@Col]` asks. */
  tableAt(sheetName, row, col) {
    for (const t of this.tables.values()) {
      if (t.sheet === sheetName && row >= t.top && row <= t.bottom && col >= t.left && col <= t.right) return t;
    }
    return null;
  }

  _invalidate(sheetName, row, col) {
    // Cheap and correct: any edit dirties the graph. The recalculation pass
    // then does the ordering. Fine-grained invalidation is an optimisation for
    // when a real sheet proves it necessary, not a correctness requirement.
    this.dirty.add(key(sheetName, row, col));
  }

  // ---- resolver ----------------------------------------------------------

  usedBounds(sheetName) {
    const cells = this.sheets.get(sheetName);
    let maxRow = 0;
    let maxCol = 0;
    if (cells) {
      for (const c of cells.values()) {
        if (c.row > maxRow) maxRow = c.row;
        if (c.col > maxCol) maxCol = c.col;
      }
    }
    // A spill is used space too — a whole-column SUM over it must see the
    // ghosts, and the grid must extend to them.
    for (const s of this.spills.values()) {
      if (s.blocked || s.sheet !== sheetName) continue;
      if (s.top + s.h - 1 > maxRow) maxRow = s.top + s.h - 1;
      if (s.left + s.w - 1 > maxCol) maxCol = s.left + s.w - 1;
    }
    return { maxRow, maxCol };
  }

  resolver() {
    const self = this;
    return {
      now: self.now,
      usedBounds: (sheet) => self.usedBounds(sheet),
      getCell(sheet, row, col) {
        const c = self.cell(sheet, row, col);
        if (!c) return self.spillValueAt(sheet, row, col) ?? '';
        return c.value ?? '';
      },
      getRange(sheet, start, end) {
        const grid = [];
        for (let r = start.row; r <= end.row; r++) {
          const line = [];
          for (let c = start.col; c <= end.col; c++) {
            const cell = self.cell(sheet, r, c);
            line.push(cell ? (cell.value ?? '') : (self.spillValueAt(sheet, r, c) ?? ''));
          }
          grid.push(line);
        }
        return grid;
      },
      getName(name) {
        return self.names.get(String(name).toUpperCase()) ?? null;
      },
      getTable(name) {
        return self.tables.get(String(name).toUpperCase()) ?? null;
      },
      tableAt(sheet, row, col) {
        return self.tableAt(sheet, row, col);
      },
      // The live spill at an anchor, for the `A1#` operator. A blocked spill
      // has no range to give.
      spillRange(sheet, row, col) {
        const s = self.spills.get(key(sheet, row, col));
        if (!s || s.blocked) return null;
        return s.values.map((line) => [...line]);
      },
    };
  }

  // ---- recalculation ------------------------------------------------------

  /** Every formula cell in the workbook, in insertion order. */
  _formulaCells() {
    const out = [];
    for (const cells of this.sheets.values()) {
      for (const c of cells.values()) if (c.ast) out.push(c);
    }
    return out;
  }

  /** Cell keys a formula cell reads. Ranges expand to their members. */
  _dependencyKeys(cell) {
    const out = new Set();
    for (const dep of cell.deps ?? []) {
      if (dep.type === 'cell') {
        out.add(key(dep.sheet ?? cell.sheet, dep.row, dep.col));
      } else if (dep.type === 'range') {
        const sheet = dep.sheet ?? dep.start.sheet ?? cell.sheet;
        const bounds = this.usedBounds(sheet);
        const r0 = Math.min(dep.start.row, dep.end.row);
        const r1 = Math.min(Math.max(dep.start.row, dep.end.row), bounds.maxRow);
        const c0 = Math.min(dep.start.col, dep.end.col);
        const c1 = Math.min(Math.max(dep.start.col, dep.end.col), bounds.maxCol);
        for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) out.add(key(sheet, r, c));
      } else if (dep.type === 'name') {
        const target = this.names.get(String(dep.name).toUpperCase());
        if (target && target.start && target.end) {
          const sheet = target.sheet ?? cell.sheet;
          for (let r = target.start.row; r <= target.end.row; r++) {
            for (let c = target.start.col; c <= target.end.col; c++) out.add(key(sheet, r, c));
          }
        }
      } else if (dep.type === 'spillref') {
        // The anchor is a real formula-to-formula edge, so the reader always
        // evaluates AFTER it; the coverage keys are the previous pass's — a
        // spill that changes shape re-runs its readers through the touched
        // ghosts of the follow-up pass, whose graph sees the new coverage.
        const sheet = dep.sheet ?? cell.sheet;
        const anchorKey = key(sheet, dep.row, dep.col);
        out.add(anchorKey);
        const s = this.spills.get(anchorKey);
        if (s && !s.blocked) {
          for (let r = s.top; r < s.top + s.h; r++) {
            for (let c = s.left; c < s.left + s.w; c++) out.add(key(s.sheet, r, c));
          }
        }
      } else if (dep.type === 'structref') {
        // The same shape arithmetic the evaluator does, minus the errors —
        // an unresolvable reference simply contributes no edges. `@` is
        // resolvable HERE because the asking cell is in hand: its own row.
        const table = dep.table
          ? this.tables.get(String(dep.table).toUpperCase())
          : this.tableAt(cell.sheet, cell.row, cell.col);
        if (!table) continue;
        const headerRows = table.headerRows ?? 1;
        const totalsRows = table.totalsRows ?? 0;
        const dataTop = table.top + headerRows;
        const dataBottom = table.bottom - totalsRows;
        let top = dep.area === 'all' ? table.top
          : dep.area === 'headers' ? table.top
            : dep.area === 'totals' ? dataBottom + 1 : dataTop;
        let bottom = dep.area === 'all' ? table.bottom
          : dep.area === 'headers' ? table.top + headerRows - 1
            : dep.area === 'totals' ? table.bottom : dataBottom;
        let left = table.left;
        let right = table.right;
        if (dep.startCol) {
          const names = (table.columns ?? []).map((c) => String(c).toLowerCase());
          const a = names.indexOf(String(dep.startCol).toLowerCase());
          const b = names.indexOf(String(dep.endCol ?? dep.startCol).toLowerCase());
          if (a < 0 || b < 0) continue;
          left = table.left + Math.min(a, b);
          right = table.left + Math.max(a, b);
        }
        if (dep.thisRow) {
          if (cell.row < dataTop || cell.row > dataBottom) continue;
          top = cell.row;
          bottom = cell.row;
        }
        for (let r = top; r <= bottom; r++) {
          for (let c = left; c <= right; c++) out.add(key(table.sheet ?? cell.sheet, r, c));
        }
      }
    }
    return out;
  }

  /**
   * Recalculate every formula in dependency order.
   *
   * Kahn's algorithm over the dependency graph. Anything left with a non-zero
   * in-degree at the end is in a cycle — that is the detection, and it costs
   * nothing extra.
   *
   * @returns {{calculated: number, cycles: string[]}}
   */
  recalculate({ all = false, _depth = 0 } = {}) {
    const formulas = this._formulaCells();
    const byKey = new Map(formulas.map((c) => [key(c.sheet, c.row, c.col), c]));

    // Edges: dependency -> dependents.
    //
    // Built for EVERY dependency, including literal cells, because that is what
    // makes an incremental pass possible: typing a number in B2 has to reach the
    // formulas that read B2, and B2 is not a formula. In-degree, by contrast,
    // counts only formula-to-formula edges — ordering matters only among things
    // that have to be evaluated.
    const dependents = new Map();
    const indegree = new Map();
    for (const k of byKey.keys()) indegree.set(k, 0);

    for (const cell of formulas) {
      const k = key(cell.sheet, cell.row, cell.col);
      for (const depKey of this._dependencyKeys(cell)) {
        if (depKey === k) { indegree.set(k, (indegree.get(k) ?? 0) + 1); continue; } // self-reference
        if (!dependents.has(depKey)) dependents.set(depKey, new Set());
        if (dependents.get(depKey).has(k)) continue;
        dependents.get(depKey).add(k);
        if (byKey.has(depKey)) indegree.set(k, (indegree.get(k) ?? 0) + 1);
      }
    }

    /*
     * Only what the edit can actually have changed.
     *
     * Recalculating every formula on every keystroke is correct and was fine
     * until a real sheet turned up: four thousand formulas cost 149ms a
     * keypress, and the sheet a customer sends is bigger than the one we tested
     * with. So the scope is the transitive closure of the dirty cells over the
     * dependency edges — everything downstream of what was touched, and nothing
     * else. A cell nothing changed keeps the value it already had.
     */
    const scope = new Set();
    if (all || this.dirty.size === 0) {
      for (const k of byKey.keys()) scope.add(k);
    } else {
      // Volatile cells are dirty by definition — nothing in the graph says the
      // clock moved or what INDIRECT will read this time — so every incremental
      // pass includes them, and their dependents through the closure below.
      const seeds = new Set(this.dirty);
      for (const k of this.volatileCells) if (byKey.has(k)) seeds.add(k);
      // An edit inside a spill's coverage re-seeds its ANCHOR: a blocker
      // typed in makes it #SPILL!, a blocker cleared lets it spill again.
      for (const k of this.dirty) {
        const anchor = this._spillAnchorTouching(k);
        if (anchor && byKey.has(anchor)) seeds.add(anchor);
      }
      const stack = [...seeds];
      for (const k of seeds) if (byKey.has(k)) scope.add(k);
      while (stack.length) {
        for (const d of dependents.get(stack.pop()) ?? []) {
          if (scope.has(d)) continue;
          scope.add(d);
          stack.push(d);
        }
      }
    }

    // In-degree within the scope: a dependency outside it already holds its
    // final value, so it imposes no ordering on this pass.
    const pending = new Map();
    for (const k of scope) {
      let n = 0;
      for (const depKey of this._dependencyKeys(byKey.get(k))) {
        if (depKey === k) { n += 1; continue; }
        if (scope.has(depKey)) n += 1;
      }
      pending.set(k, n);
    }

    const queue = [...pending.entries()].filter(([, n]) => n === 0).map(([k]) => k);
    const order = [];
    while (queue.length) {
      const k = queue.shift();
      order.push(k);
      for (const d of dependents.get(k) ?? []) {
        if (!pending.has(d)) continue;
        pending.set(d, pending.get(d) - 1);
        if (pending.get(d) === 0) queue.push(d);
      }
    }

    const resolver = this.resolver();
    // Ghost cells whose spilled value changed this pass — they are the dirty
    // seeds of the follow-up pass, so readers of a spill catch up.
    const spillTouched = new Set();
    for (const k of order) {
      const cell = byKey.get(k);
      const raw = evaluate(cell.ast, resolver, {
        sheet: cell.sheet, row: cell.row, col: cell.col, spill: true,
      });
      this._settle(cell, k, raw, spillTouched);
    }

    // Whatever never reached in-degree zero is circular. A Set, not
    // `order.includes` — that was O(n) inside a loop over n, sixteen million
    // comparisons on a four-thousand-formula sheet, and it looked innocent.
    const evaluated = new Set(order);
    const cycles = [...scope].filter((k) => !evaluated.has(k));
    for (const k of cycles) {
      const cell = byKey.get(k);
      cell.value = ERR.CIRCULAR('this cell is part of a circular reference');
      cell.error = cell.value;
    }

    this.dirty.clear();

    // A spill that changed shape or values is only half-delivered: readers of
    // its ghost cells may have computed BEFORE the anchor did. One follow-up
    // pass seeded with the touched ghosts settles them; the depth guard stops
    // a pathological pair of interfering spills from ping-ponging forever.
    if (spillTouched.size && _depth < 4) {
      for (const gk of spillTouched) this.dirty.add(gk);
      const again = this.recalculate({ _depth: _depth + 1 });
      return {
        calculated: order.length + again.calculated,
        cycles: [...cycles.map(prettyKey), ...again.cycles],
      };
    }
    return { calculated: order.length, cycles: cycles.map(prettyKey) };
  }

  /** The anchor whose spill (live or blocked) covers a cell key, or null. */
  _spillAnchorTouching(cellKey) {
    const [sheet, rc] = cellKey.split('!');
    const [row, col] = rc.split(':').map(Number);
    for (const [anchorKey, s] of this.spills) {
      if (s.sheet === sheet && row >= s.top && row < s.top + s.h && col >= s.left && col < s.left + s.w) {
        return anchorKey;
      }
    }
    return null;
  }

  /** Is (row, col) covered by an UNBLOCKED spill other than `exceptKey`? */
  _spillCovering(sheetName, row, col, exceptKey) {
    for (const [anchorKey, s] of this.spills) {
      if (anchorKey === exceptKey || s.blocked) continue;
      if (s.sheet === sheetName && row >= s.top && row < s.top + s.h && col >= s.left && col < s.left + s.w) {
        return anchorKey;
      }
    }
    return null;
  }

  /** Forget a spill, marking every ghost it was painting as touched. */
  _dropSpill(anchorKey, spillTouched) {
    const prev = this.spills.get(anchorKey);
    this.spills.delete(anchorKey);
    if (!prev || prev.blocked) return;
    for (let r = 0; r < prev.h; r++) {
      for (let c = 0; c < prev.w; c++) {
        if (r === 0 && c === 0) continue;
        spillTouched.add(key(prev.sheet, prev.top + r, prev.left + c));
      }
    }
  }

  /**
   * Land one formula's freshly evaluated result: a scalar into the cell, an
   * array as a SPILL — the anchor keeps the top-left value, empty neighbours
   * become ghosts, and anything real in the way makes the whole thing
   * #SPILL! rather than a silent partial overwrite.
   */
  _settle(cell, k, raw, spillTouched) {
    if (!Array.isArray(raw)) {
      if (this.spills.has(k)) this._dropSpill(k, spillTouched);
      cell.value = raw;
      cell.error = isError(raw) ? raw : null;
      return;
    }

    const grid = raw.length && Array.isArray(raw[0]) ? raw : [raw];
    const h = Math.max(1, grid.length);
    const w = Math.max(1, ...grid.map((line) => line.length));
    const values = grid.map((line) => Array.from({ length: w }, (_, j) => line[j] ?? ''));
    if (h === 1 && w === 1) {
      if (this.spills.has(k)) this._dropSpill(k, spillTouched);
      cell.value = values[0][0];
      cell.error = isError(cell.value) ? cell.value : null;
      return;
    }

    let blocked = false;
    for (let r = cell.row; !blocked && r < cell.row + h; r++) {
      for (let c = cell.col; c < cell.col + w; c++) {
        if (r === cell.row && c === cell.col) continue;
        if (this.cell(cell.sheet, r, c) || this._spillCovering(cell.sheet, r, c, k)) {
          blocked = true;
          break;
        }
      }
    }

    const prev = this.spills.get(k);
    if (blocked) {
      if (prev && !prev.blocked) this._dropSpill(k, spillTouched);
      this.spills.set(k, { sheet: cell.sheet, top: cell.row, left: cell.col, h, w, values: null, blocked: true });
      cell.value = ERR.SPILL('something is in the way of this formula’s spill');
      cell.error = cell.value;
      return;
    }

    if (prev && !prev.blocked) {
      const maxH = Math.max(prev.h, h);
      const maxW = Math.max(prev.w, w);
      for (let r = 0; r < maxH; r++) {
        for (let c = 0; c < maxW; c++) {
          if (r === 0 && c === 0) continue;
          const before = r < prev.h && c < prev.w ? prev.values[r][c] : undefined;
          const after = r < h && c < w ? values[r][c] : undefined;
          if (!sameCellValue(before, after)) spillTouched.add(key(cell.sheet, cell.row + r, cell.col + c));
        }
      }
    } else {
      for (let r = 0; r < h; r++) {
        for (let c = 0; c < w; c++) {
          if (r === 0 && c === 0) continue;
          spillTouched.add(key(cell.sheet, cell.row + r, cell.col + c));
        }
      }
    }
    this.spills.set(k, { sheet: cell.sheet, top: cell.row, left: cell.col, h, w, values, blocked: false });
    cell.value = values[0][0];
    cell.error = isError(cell.value) ? cell.value : null;
  }

  /** A snapshot for a UI or a test: displayed values, row-major. */
  grid(sheetName) {
    const { maxRow, maxCol } = this.usedBounds(sheetName);
    if (this.dirty.size) this.recalculate();
    const out = [];
    for (let r = 0; r <= maxRow; r++) {
      const line = [];
      for (let c = 0; c <= maxCol; c++) {
        const cell = this.cell(sheetName, r, c);
        line.push(cell ? (cell.value ?? '') : (this.spillValueAt(sheetName, r, c) ?? ''));
      }
      out.push(line);
    }
    return out;
  }
}

/** Value equality for spill diffing — two errors of one type are the same. */
function sameCellValue(a, b) {
  if (isError(a) && isError(b)) return a.type === b.type;
  return a === b;
}

/** Does a formula call a volatile function, at any depth? */
function hasVolatileCall(node) {
  if (!node || typeof node !== 'object') return false;
  if (node.type === 'call' && isVolatile(node.name)) return true;
  for (const k of ['left', 'right', 'operand']) if (hasVolatileCall(node[k])) return true;
  return (node.args ?? []).some(hasVolatileCall);
}

function prettyKey(k) {
  const [sheet, rc] = k.split('!');
  const [row, col] = rc.split(':').map(Number);
  return sheet + '!' + indexToCol(col) + (row + 1);
}

export { prettyKey };
