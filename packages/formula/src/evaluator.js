/**
 * Evaluate a parsed formula against a sheet.
 *
 * The `resolver` is the whole interface to the outside world:
 *   getCell(sheet, row, col) -> scalar
 *   getRange(sheet, start, end) -> 2-D array of scalars
 *   getName(name) -> { sheet, start, end } | scalar | null
 *   now() -> Date
 *
 * Keeping it that narrow is deliberate. The evaluator never sees a workbook, a
 * file or a UI, which is what lets it be tested exhaustively against plain
 * objects and reused unchanged by the recalculation engine, the binding refresh
 * path and eventually the editor.
 */
import {
  ERR, isError, isBlank, firstError,
  toNumber, toText, toBoolean, compareValues,
} from './values.js';
import { FUNCTIONS } from './functions.js';
import { parse, indexToCol } from './parser.js';

const MAX_ROWS = 1048576;
const MAX_COLS = 16384;

/** Clip whole-column/row ranges to what the sheet actually uses. */
function clip(resolver, sheet, start, end) {
  const used = resolver.usedBounds ? resolver.usedBounds(sheet) : null;
  const maxRow = used ? used.maxRow : 1000;
  const maxCol = used ? used.maxCol : 100;
  // Clipping trims a range that runs PAST the used cells; it must never pull
  // an end back BEFORE its start. A range wholly beyond them — D7:D9 on a
  // sheet whose last row is 5 — covers blanks, and clamping its end to that
  // row turned it into D5:D7 once the backwards-range rule normalised it:
  // =SUM(D7:D9) answered with the contents of D5.
  const endRow = Math.min(end.row, maxRow);
  const endCol = Math.min(end.col, maxCol);
  return {
    start: { row: Math.min(start.row, MAX_ROWS - 1), col: Math.min(start.col, MAX_COLS - 1) },
    end: {
      row: endRow < start.row ? Math.min(end.row, MAX_ROWS - 1) : endRow,
      col: endCol < start.col ? Math.min(end.col, MAX_COLS - 1) : endCol,
    },
  };
}

export function evaluate(ast, resolver, context = {}) {
  const ctx = {
    sheet: context.sheet ?? null,
    // Which cell is asking — what ROW() and COLUMN() mean with no argument.
    row: context.row ?? null,
    col: context.col ?? null,
    now: resolver.now ?? (() => new Date()),
    depth: context.depth ?? 0,
  };
  if (ctx.depth > 128) return ERR.NUM('formula nested too deeply');

  const evalNode = (node) => {
    switch (node.type) {
      case 'literal':
        return node.value;

      case 'errorLiteral':
        return ERR[
          { '#NULL!': 'NULL', '#DIV/0!': 'DIV0', '#VALUE!': 'VALUE', '#REF!': 'REF',
            '#NAME?': 'NAME', '#NUM!': 'NUM', '#N/A': 'NA', '#CIRCULAR!': 'CIRCULAR',
            '#SPILL!': 'SPILL', '#CALC!': 'CALC' }[node.value]
        ]();

      case 'cell': {
        const sheet = node.sheet ?? ctx.sheet;
        if (node.row < 0 || node.col < 0) return ERR.REF();
        return resolver.getCell(sheet, node.row, node.col);
      }

      case 'range': {
        const sheet = node.sheet ?? node.start.sheet ?? ctx.sheet;
        const bounds = clip(resolver, sheet, node.start, node.end);
        // A range written backwards (B2:A1) is normalised, as a spreadsheet does.
        const start = {
          row: Math.min(bounds.start.row, bounds.end.row),
          col: Math.min(bounds.start.col, bounds.end.col),
        };
        const end = {
          row: Math.max(bounds.start.row, bounds.end.row),
          col: Math.max(bounds.start.col, bounds.end.col),
        };
        return resolver.getRange(sheet, start, end);
      }

      case 'name': {
        const resolved = resolver.getName ? resolver.getName(node.name, ctx.sheet) : null;
        if (resolved === null || resolved === undefined) return ERR.NAME('unknown name "' + node.name + '"');
        if (typeof resolved === 'object' && resolved.start && resolved.end) {
          const bounds = clip(resolver, resolved.sheet ?? ctx.sheet, resolved.start, resolved.end);
          return resolver.getRange(resolved.sheet ?? ctx.sheet, bounds.start, bounds.end);
        }
        return resolved;
      }

      case 'structref':
        return structref(node);

      // `A1#` — the range a spill anchor covers. A cell that is not
      // currently a live spill anchor has no such range, which is #REF!
      // rather than its own scalar value: the operator PROMISES a range.
      case 'spillref': {
        const sheet = node.sheet ?? ctx.sheet;
        const grid = resolver.spillRange ? resolver.spillRange(sheet, node.row, node.col) : null;
        return grid ?? ERR.REF('nothing spills from ' + indexToCol(node.col) + (node.row + 1));
      }

      case 'unary': {
        const operand = evalNode(node.operand);
        if (Array.isArray(operand)) {
          return to2d(operand).map((line) => line.map((v) => applyUnary(node.op, v)));
        }
        return applyUnary(node.op, scalar(operand));
      }

      case 'binary':
        return binary(node);

      case 'call':
        return call(node);

      default:
        return ERR.VALUE('cannot evaluate node type ' + node.type);
    }
  };

  /** A range used where one value is expected collapses to its first cell. */
  const scalar = (v) => {
    if (Array.isArray(v)) {
      const flat = v.flat(Infinity);
      return flat.length ? flat[0] : '';
    }
    return v;
  };

  /**
   * A structured reference — `Table1[Qty]`, `[@Price]`, `Table1[#Totals]` —
   * resolved against the resolver's table registry. The table's shape does
   * the arithmetic: header rows on top, totals rows underneath, the data
   * body between them, and a column span narrowing the width. `@` means the
   * ASKING cell's row, which is why this lives in the evaluator.
   */
  function structref(node) {
    const table = node.table
      ? (resolver.getTable ? resolver.getTable(node.table) : null)
      : (resolver.tableAt && ctx.sheet !== null && ctx.row !== null && ctx.col !== null
        ? resolver.tableAt(ctx.sheet, ctx.row, ctx.col)
        : null);
    if (!table) {
      return node.table
        ? ERR.NAME('unknown table "' + node.table + '"')
        : ERR.VALUE('[@…] names "this table", and this cell is not in one');
    }

    const sheet = table.sheet ?? ctx.sheet;
    const headerRows = table.headerRows ?? 1;
    const totalsRows = table.totalsRows ?? 0;
    const dataTop = table.top + headerRows;
    const dataBottom = table.bottom - totalsRows;

    let top;
    let bottom;
    switch (node.area) {
      case 'all': top = table.top; bottom = table.bottom; break;
      case 'headers':
        if (!headerRows) return ERR.REF('this table has no header row');
        top = table.top; bottom = table.top + headerRows - 1; break;
      case 'totals':
        if (!totalsRows) return ERR.REF('this table has no totals row');
        top = dataBottom + 1; bottom = table.bottom; break;
      default: top = dataTop; bottom = dataBottom;
    }

    let left = table.left;
    let right = table.right;
    if (node.startCol) {
      const names = (table.columns ?? []).map((c) => String(c).toLowerCase());
      const a = names.indexOf(String(node.startCol).toLowerCase());
      const b = names.indexOf(String(node.endCol ?? node.startCol).toLowerCase());
      if (a < 0 || b < 0) {
        return ERR.REF('no column "' + (a < 0 ? node.startCol : node.endCol)
          + '" in table ' + (table.name ?? node.table ?? ''));
      }
      left = table.left + Math.min(a, b);
      right = table.left + Math.max(a, b);
    }

    if (node.thisRow) {
      if (ctx.row === null || ctx.row < dataTop || ctx.row > dataBottom) {
        return ERR.VALUE('[@…] means "this row", and this row is outside the table body');
      }
      top = ctx.row;
      bottom = ctx.row;
    }

    if (top > bottom) return ERR.REF();
    if (top === bottom && left === right) return resolver.getCell(sheet, top, left);
    return resolver.getRange(sheet, { row: top, col: left }, { row: bottom, col: right });
  }

  /**
   * An operator with an ARRAY on either side broadcasts elementwise, the way
   * `=A1:A5>2` or `=SEQUENCE(3)*10` behave in a modern spreadsheet: a
   * 1-row or 1-column side extends across the other's shape, and a genuine
   * size mismatch is #N/A in the cells that have no partner.
   */
  function binary(node) {
    const left = evalNode(node.left);
    const right = evalNode(node.right);
    if (Array.isArray(left) || Array.isArray(right)) {
      const A = to2d(left);
      const B = to2d(right);
      const rows = Math.max(A.length, B.length);
      const cols = Math.max(A[0]?.length ?? 0, B[0]?.length ?? 0);
      const out = [];
      for (let i = 0; i < rows; i++) {
        const line = [];
        for (let j = 0; j < cols; j++) line.push(applyBinary(node.op, pick2d(A, i, j), pick2d(B, i, j)));
        out.push(line);
      }
      return out;
    }
    return applyBinary(node.op, scalar(left), scalar(right));
  }

  function applyBinary(op, left, right) {
    const e = firstError([left, right]);
    if (e) return e;

    if (op === '&') {
      const a = toText(left);
      const b = toText(right);
      return firstError([a, b]) || a + b;
    }

    if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
      const cmp = compareValues(left, right);
      if (isError(cmp)) return cmp;
      switch (op) {
        case '=': return cmp === 0;
        case '<>': return cmp !== 0;
        case '<': return cmp < 0;
        case '<=': return cmp <= 0;
        case '>': return cmp > 0;
        case '>=': return cmp >= 0;
        default: return ERR.VALUE();
      }
    }

    const a = toNumber(left);
    const b = toNumber(right);
    const numErr = firstError([a, b]);
    if (numErr) return numErr;

    switch (op) {
      case '+': return a + b;
      case '-': return a - b;
      case '*': return a * b;
      case '/': return b === 0 ? ERR.DIV0() : a / b;
      case '^': {
        const r = a ** b;
        return Number.isFinite(r) ? r : ERR.NUM();
      }
      default: return ERR.VALUE('unknown operator ' + op);
    }
  }

  function applyUnary(op, value) {
    const n = toNumber(value);
    if (isError(n)) return n;
    if (op === '%') return n / 100;
    return op === '-' ? -n : n;
  }

  function call(node) {
    // ROW and COLUMN read a reference's POSITION, which only the AST still
    // knows — by the time arguments are evaluated a cell is just its value.
    // Handled here rather than in the library, and before evaluation, so
    // ROW(A5) is 5 whatever A5 holds.
    if ((node.name === 'ROW' || node.name === 'COLUMN') && node.args.length <= 1) {
      const wantRow = node.name === 'ROW';
      const arg = node.args[0];
      if (!arg) {
        const at = wantRow ? ctx.row : ctx.col;
        return at === null || at === undefined ? ERR.VALUE(node.name + '() needs a cell to be asked from') : at + 1;
      }
      if (arg.type === 'cell') return (wantRow ? arg.row : arg.col) + 1;
      if (arg.type === 'range') {
        return (wantRow ? Math.min(arg.start.row, arg.end.row) : Math.min(arg.start.col, arg.end.col)) + 1;
      }
      return ERR.VALUE(node.name + ' takes a reference');
    }

    // OFFSET and INDIRECT construct a NEW reference at evaluation time, which
    // needs the two things only the evaluator holds: the AST (for an anchor's
    // position) and the resolver (to read wherever the new reference lands).
    // Both are volatile in the engine — nothing static says what they read.
    if (node.name === 'OFFSET') {
      if (node.args.length < 3 || node.args.length > 5) {
        return ERR.VALUE('OFFSET takes 3 to 5 arguments');
      }
      const anchor = node.args[0];
      if (!anchor || (anchor.type !== 'cell' && anchor.type !== 'range')) {
        return ERR.VALUE('OFFSET needs a cell or range to start from');
      }
      const dr = toNumber(scalar(evalNode(node.args[1])));
      const dc = toNumber(scalar(evalNode(node.args[2])));
      const moveErr = firstError([dr, dc]);
      if (moveErr) return moveErr;
      const sheet = anchor.sheet ?? anchor.start?.sheet ?? ctx.sheet;
      const top = (anchor.type === 'cell' ? anchor.row : Math.min(anchor.start.row, anchor.end.row)) + Math.trunc(dr);
      const left = (anchor.type === 'cell' ? anchor.col : Math.min(anchor.start.col, anchor.end.col)) + Math.trunc(dc);
      let height = anchor.type === 'cell' ? 1 : Math.abs(anchor.end.row - anchor.start.row) + 1;
      let width = anchor.type === 'cell' ? 1 : Math.abs(anchor.end.col - anchor.start.col) + 1;
      if (node.args[3] !== undefined) {
        const h = toNumber(scalar(evalNode(node.args[3])));
        if (isError(h)) return h;
        height = Math.trunc(h);
      }
      if (node.args[4] !== undefined) {
        const w = toNumber(scalar(evalNode(node.args[4])));
        if (isError(w)) return w;
        width = Math.trunc(w);
      }
      if (height < 1 || width < 1) return ERR.REF('OFFSET height and width must be at least 1');
      if (top < 0 || left < 0 || top + height > MAX_ROWS || left + width > MAX_COLS) return ERR.REF();
      if (height === 1 && width === 1) return resolver.getCell(sheet, top, left);
      return resolver.getRange(sheet, { row: top, col: left }, { row: top + height - 1, col: left + width - 1 });
    }

    if (node.name === 'INDIRECT') {
      if (node.args.length < 1 || node.args.length > 2) {
        return ERR.VALUE('INDIRECT takes 1 or 2 arguments');
      }
      if (node.args[1] !== undefined) {
        const a1 = toBoolean(scalar(evalNode(node.args[1])));
        if (isError(a1)) return a1;
        if (a1 === false) return ERR.REF('R1C1-style references are not supported');
      }
      const text = toText(scalar(evalNode(node.args[0])));
      if (isError(text)) return text;
      let refAst;
      try {
        refAst = parse('=' + text);
      } catch {
        return ERR.REF('"' + text + '" is not a reference');
      }
      if (refAst.type !== 'cell' && refAst.type !== 'range' && refAst.type !== 'name') {
        return ERR.REF('"' + text + '" is not a reference');
      }
      // Evaluated as if the formula had contained the reference itself, so
      // sheet qualifiers, defined names and whole-column clipping all apply.
      return evalNode(refAst);
    }

    // IF broadcasts over an ARRAY condition — what a spilled comparison
    // implies. Laziness is a scalar's privilege: an array IF must look at
    // both branches, as Excel's does; a scalar one still takes only the
    // branch its condition picked.
    if (node.name === 'IF' && node.args.length >= 2 && node.args.length <= 3) {
      const cond = evalNode(node.args[0]);
      if (Array.isArray(cond)) {
        const C = to2d(cond);
        const T = to2d(evalNode(node.args[1]));
        const F = to2d(node.args[2] ? evalNode(node.args[2]) : [[false]]);
        const rows = Math.max(C.length, T.length, F.length);
        const cols = Math.max(C[0]?.length ?? 0, T[0]?.length ?? 0, F[0]?.length ?? 0);
        const out = [];
        for (let i = 0; i < rows; i++) {
          const line = [];
          for (let j = 0; j < cols; j++) {
            const c = pick2d(C, i, j);
            const b = isError(c) ? c : toBoolean(c);
            line.push(isError(b) ? b : (b ? pick2d(T, i, j) : pick2d(F, i, j)));
          }
          out.push(line);
        }
        return out;
      }
      const b = toBoolean(scalar(cond));
      if (isError(b)) return b;
      if (b) return evalNode(node.args[1]);
      return node.args[2] ? evalNode(node.args[2]) : false;
    }

    const entry = FUNCTIONS[node.name];
    if (!entry) return ERR.NAME('unknown function ' + node.name);

    if (entry.lazy) {
      // Thunks, so a branch that is not taken is never evaluated.
      const thunks = node.args.map((a) => () => scalarOrArray(evalNode(a)));
      try {
        return entry.fn(...thunks);
      } catch (e) {
        return isError(e) ? e : ERR.VALUE(e.message);
      }
    }

    const args = node.args.map((a) => scalarOrArray(evalNode(a)));
    if (!entry.wantsContext) {
      const e = firstError(args);
      // Aggregates still need to see errors inside ranges, and firstError walks
      // into arrays, so this catches both literal and in-range errors.
      if (e) return e;
    }
    try {
      return entry.wantsContext ? entry.fn(args[0], ctx) : entry.fn(...args);
    } catch (e) {
      return isError(e) ? e : ERR.VALUE(e.message);
    }
  }

  const scalarOrArray = (v) => v;

  try {
    const result = evalNode(ast);
    // In spill mode the caller WANTS the whole array — that is what spills.
    // Everywhere else a top-level array collapses to its first cell, the
    // legacy implicit intersection.
    if (context.spill) return result;
    return Array.isArray(result) ? scalar(result) : result;
  } catch (e) {
    return isError(e) ? e : ERR.VALUE(e.message ?? String(e));
  }
}

/** Any value as a rectangular 2-D array — the shape broadcasting works in. */
function to2d(v) {
  if (!Array.isArray(v)) return [[v]];
  return Array.isArray(v[0]) ? v : [v];
}

/** One element of a 2-D array under broadcast rules: 1 extends, else #N/A. */
function pick2d(grid, i, j) {
  const rows = grid.length;
  const cols = grid[0]?.length ?? 0;
  const ri = rows === 1 ? 0 : i;
  const cj = cols === 1 ? 0 : j;
  if (ri >= rows || cj >= (grid[ri]?.length ?? 0)) return ERR.NA();
  return grid[ri][cj];
}

/** Parse and evaluate in one step. Convenience for tests and one-off cells. */
export function calculate(formula, resolver, context) {
  let ast;
  try {
    ast = parse(formula);
  } catch (e) {
    return isError(e) ? e : ERR.VALUE(e.message ?? String(e));
  }
  return evaluate(ast, resolver, context);
}

export { isBlank };
