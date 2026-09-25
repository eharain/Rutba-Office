/**
 * Evaluate Formula, one part at a time.
 *
 * Excel's Formulas → Evaluate Formula shows a formula with the part it will
 * work out next underlined; Evaluate puts that part's value in its place, and
 * the formula shrinks a step at a time to its result. Step In opens the
 * formula of a cell the underlined reference points at, Step Out puts the
 * cell's value back in the formula it was stepped into from.
 *
 * The order is the evaluator's own: a reference is read before whatever
 * uses it, an operator's left side before its right, a function's arguments
 * left to right and the function last. A function that decides for itself
 * which arguments it needs — IF, IFERROR, CHOOSE, IFS, SWITCH — is asked:
 * its library entry is run with arguments that answer only once they have
 * been stepped, so the branch IF does not take is left as it was written,
 * exactly as Excel leaves it. Constants are never steps.
 *
 * Nothing here computes a value its own way. Each part's value is the
 * evaluator's answer for that part, so the steps always end on the number
 * the cell shows; the steps are only the order they are shown in.
 */
import { parse, indexToCol } from './parser.js';
import { evaluate } from './evaluator.js';
import { FUNCTIONS } from './functions.js';
import { isError, isBlank, formatNumber } from './values.js';

/** More than this many values of a range are shown as an ellipsis. */
const MAX_SHOWN = 1000;

/** How a value is written in the formula once it has been worked out. */
export function evaluationText(value, { blank = '0' } = {}) {
  if (Array.isArray(value)) {
    const rows = Array.isArray(value[0]) ? value : [value];
    let shown = 0;
    const lines = [];
    for (const row of rows) {
      const parts = [];
      for (const v of row) {
        if (shown >= MAX_SHOWN) { parts.push('…'); break; }
        parts.push(evaluationText(v, { blank }));
        shown += 1;
      }
      lines.push(parts.join(','));
      if (shown >= MAX_SHOWN) break;
    }
    return '{' + lines.join(';') + '}';
  }
  if (isError(value)) return value.type;
  if (value === true) return 'TRUE';
  if (value === false) return 'FALSE';
  // An empty cell reads as 0, the way Excel's dialog shows one.
  if (isBlank(value)) return blank;
  if (typeof value === 'number') return formatNumber(value);
  return '"' + String(value).replace(/"/g, '""') + '"';
}

const REFERENCE_TYPES = new Set(['cell', 'range', 'name', 'structref', 'spillref']);
/** Thrown by a stand-in argument that has not been stepped yet. */
const NEED = Symbol('need');

/**
 * One formula, stepped. `context` is the asking cell — `{ sheet, row, col }`
 * — which is what a reference without a sheet and ROW() mean.
 */
export class FormulaStepper {
  constructor(formula, resolver, context = {}) {
    this.formula = String(formula ?? '');
    this.source = this.formula.startsWith('=') ? this.formula.slice(1) : this.formula;
    this.resolver = resolver;
    this.context = { sheet: context.sheet ?? null, row: context.row ?? null, col: context.col ?? null };
    this.error = null;
    try {
      this.ast = parse('=' + this.source, { spans: true });
    } catch (e) {
      this.ast = null;
      this.error = e?.type ?? '#VALUE!';
    }
    this.restart();
  }

  restart() {
    /** Parts worked out so far, and the value each was replaced by. */
    this.reduced = new Map();
    this._values = new Map();
    this.recent = null;
    return this;
  }

  /** A part's value, from the evaluator — arrays kept whole. */
  valueOf(node) {
    if (this._values.has(node)) return this._values.get(node);
    const value = evaluate(node, this.resolver, { ...this.context, spill: true });
    this._values.set(node, value);
    return value;
  }

  /** Whether a node is written as it stands: a literal, or a sign on one. */
  _constant(node) {
    if (node.type === 'literal' || node.type === 'errorLiteral') return true;
    if (node.type === 'unary') return this._constant(node.operand);
    if (node.type === 'array') return node.rows.every((row) => row.every((n) => this._constant(n)));
    return false;
  }

  /** The next part to work out under `node`, or null when it has none left. */
  _next(node) {
    if (!node || this.reduced.has(node)) return null;
    if (this._constant(node)) return null;
    if (REFERENCE_TYPES.has(node.type)) return node;
    if (node.type === 'array') {
      for (const row of node.rows) for (const n of row) { const found = this._next(n); if (found) return found; }
      return null;
    }
    if (node.type === 'unary') return this._next(node.operand) ?? node;
    if (node.type === 'binary') return this._next(node.left) ?? this._next(node.right) ?? node;
    if (node.type !== 'call') return node;

    const args = node.args ?? [];
    // What only the whole call can answer: LET's names mean nothing alone,
    // and ROW, COLUMN and OFFSET's anchor are positions, never values.
    if (node.name === 'LET') return node;
    if ((node.name === 'ROW' || node.name === 'COLUMN') && args.length <= 1) return node;
    if (node.name === 'OFFSET') {
      for (const a of args.slice(1)) { const found = this._next(a); if (found) return found; }
      return node;
    }

    const entry = FUNCTIONS[node.name];
    if (entry?.lazy) {
      // An IF whose condition is a whole array looks at both branches.
      if (node.name === 'IF' && args[0] && this._next(args[0]) === null && Array.isArray(this.valueOf(args[0]))) {
        for (const a of args) { const found = this._next(a); if (found) return found; }
        return node;
      }
      // Ask the function which argument it wants next: each stand-in answers
      // with its argument's value once that has been stepped, and stops the
      // function where it has not.
      let wanted = null;
      const thunks = args.map((a, i) => () => {
        if (this._next(a) !== null) { wanted = i; throw NEED; }
        return this.valueOf(a);
      });
      try {
        entry.fn(...thunks);
      } catch (e) {
        if (e === NEED && wanted !== null) return this._next(args[wanted]);
      }
      return node;
    }
    for (const a of args) { const found = this._next(a); if (found) return found; }
    return node;
  }

  /** The part the next Evaluate works out, or null once the formula is done. */
  get pending() {
    if (!this.ast) return null;
    return this._next(this.ast);
  }

  get done() { return !this.ast || this.reduced.has(this.ast) || this.pending === null; }

  /** The formula's value — worked out by the evaluator, whatever has been stepped. */
  get value() {
    if (!this.ast) return null;
    return this.valueOf(this.ast);
  }

  /** Work out the underlined part. Returns false when there was nothing left. */
  evaluate() {
    const node = this.pending;
    if (!node) {
      if (this.ast && !this.reduced.has(this.ast)) this.reduced.set(this.ast, this.valueOf(this.ast));
      return false;
    }
    this.reduced.set(node, this.valueOf(node));
    this.recent = node;
    // A formula left with no part to step (its last call's arguments all
    // constants) is finished by the same press.
    return true;
  }

  /** Put a value in place of the underlined part — Step Out's half of the work. */
  settle(value) {
    const node = this.pending;
    if (!node) return false;
    this._values.set(node, value);
    this.reduced.set(node, value);
    this.recent = node;
    return true;
  }

  /**
   * The cell the underlined part is, when it is one cell — what Step In
   * opens. `{ sheet, row, col }`, or null.
   */
  get pendingCell() {
    const node = this.pending;
    if (!node || node.type !== 'cell' || node.row < 0 || node.col < 0) return null;
    return { sheet: node.sheet ?? this.context.sheet, row: node.row, col: node.col };
  }

  /**
   * The formula as it stands: `text` with its leading `=`, `underline` the
   * [start, end) of the part the next Evaluate works out, `recent` the part
   * last worked out (Excel sets it in italics).
   */
  view() {
    if (!this.ast) return { text: '=' + this.source, underline: null, recent: null, done: true };
    const target = this.pending;
    let underline = null;
    let recent = null;
    const src = this.source;
    const walk = (node, offset) => {
      if (this.reduced.has(node)) {
        const text = evaluationText(this.reduced.get(node));
        if (node === this.recent) recent = [offset, offset + text.length];
        return text;
      }
      const kids = children(node);
      let text;
      if (!kids.length) {
        text = src.slice(node.span[0], node.span[1]);
      } else {
        text = '';
        let cursor = node.span[0];
        for (const kid of kids) {
          text += src.slice(cursor, kid.span[0]);
          text += walk(kid, offset + text.length);
          cursor = kid.span[1];
        }
        text += src.slice(cursor, node.span[1]);
      }
      if (node === target) underline = [offset, offset + text.length];
      return text;
    };
    const head = '=' + src.slice(0, this.ast.span[0]);
    const body = walk(this.ast, head.length);
    return { text: head + body + src.slice(this.ast.span[1]), underline, recent, done: this.done };
  }
}

/** A node's parts in the order they appear in the text. */
function children(node) {
  switch (node.type) {
    case 'unary': return [node.operand];
    case 'binary': return [node.left, node.right];
    case 'call': return node.args ?? [];
    case 'array': return node.rows.flat();
    default: return [];
  }
}

/** How Excel names a cell in the Reference box: Sheet1!$C$5. */
export function referenceText(sheet, row, col) {
  const name = sheet == null ? '' : /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : "'" + String(sheet).replace(/'/g, "''") + "'";
  return (name ? name + '!' : '') + '$' + indexToCol(col) + '$' + (row + 1);
}

/**
 * The whole dialog: the cell being evaluated, and the cells stepped into
 * from it, innermost last. `getFormula(sheet, row, col)` answers a cell's
 * formula text (with its `=`), or null when it holds a constant or nothing.
 */
export class FormulaEvaluation {
  constructor(resolver, { sheet, row, col }, { getFormula, maxDepth = 32 } = {}) {
    this.resolver = resolver;
    this.getFormula = getFormula ?? (() => null);
    this.maxDepth = maxDepth;
    this.levels = [];
    this._open({ sheet, row, col });
  }

  _open(at) {
    const formula = this.getFormula(at.sheet, at.row, at.col);
    const level = { at, ref: referenceText(at.sheet, at.row, at.col) };
    if (typeof formula === 'string' && formula.startsWith('=')) {
      level.stepper = new FormulaStepper(formula, this.resolver, at);
    } else {
      // A constant: shown as it is, with nothing to step.
      level.constant = this.resolver.getCell(at.sheet, at.row, at.col);
    }
    this.levels.push(level);
  }

  get current() { return this.levels[this.levels.length - 1]; }

  _levelDone(level) { return !level.stepper || level.stepper.done; }

  get canEvaluate() {
    return this.levels.length > 1 || !this._levelDone(this.current);
  }

  get canStepIn() {
    const level = this.current;
    if (!level.stepper || this.levels.length >= this.maxDepth) return false;
    const cell = level.stepper.pendingCell;
    if (!cell) return false;
    // Excel steps into a cell that holds something: a formula, or a
    // constant it shows as it is. An empty cell has nothing to show.
    const formula = this.getFormula(cell.sheet, cell.row, cell.col);
    if (formula) return true;
    return !isBlank(this.resolver.getCell(cell.sheet, cell.row, cell.col));
  }

  get canStepOut() { return this.levels.length > 1; }

  /** Evaluate: the underlined part, or — on a finished inner level — Step Out. */
  evaluate() {
    const level = this.current;
    if (this._levelDone(level)) {
      if (this.levels.length > 1) return this.stepOut();
      return false;
    }
    return level.stepper.evaluate();
  }

  stepIn() {
    if (!this.canStepIn) return false;
    this._open(this.current.stepper.pendingCell);
    return true;
  }

  /** Back to the formula stepped in from, the cell's value in place of its reference. */
  stepOut() {
    if (this.levels.length < 2) return false;
    const inner = this.levels.pop();
    const value = this.resolver.getCell(inner.at.sheet, inner.at.row, inner.at.col);
    this.current.stepper.settle(value);
    return true;
  }

  /** What the dialog shows: every level, and which buttons are live. */
  state() {
    const levels = this.levels.map((level) => {
      if (!level.stepper) {
        const text = evaluationText(level.constant, { blank: '' });
        return { ref: level.ref, text, underline: null, recent: null, done: true, constant: true };
      }
      const v = level.stepper.view();
      return { ref: level.ref, ...v, constant: false, error: level.stepper.error };
    });
    const top = levels[levels.length - 1];
    let message;
    if (top.constant) message = 'The cell currently being evaluated contains a constant.';
    else if (top.error) message = 'This formula could not be read, so it cannot be stepped.';
    else if (this.levels.length === 1 && top.done) message = 'The formula is fully evaluated. Restart evaluates it again from the start.';
    else if (top.done) message = 'This cell is fully evaluated. Step Out, or Evaluate, returns its value to the formula above.';
    else message = 'To show the result of the underlined expression, click Evaluate. The most recent result appears in italics.';
    return {
      levels,
      canEvaluate: this.canEvaluate,
      canStepIn: this.canStepIn,
      canStepOut: this.canStepOut,
      done: this.levels.length === 1 && top.done,
      message,
    };
  }

  /**
   * Replay a list of presses — 'evaluate', 'stepIn', 'stepOut', 'restart' —
   * from a fresh start. What lets a window hold only the presses and ask for
   * the state after them, with nothing to keep alive in between.
   */
  static replay(resolver, at, actions = [], opts = {}) {
    let run = new FormulaEvaluation(resolver, at, opts);
    for (const action of actions) {
      if (action === 'restart') run = new FormulaEvaluation(resolver, at, opts);
      else if (action === 'evaluate') run.evaluate();
      else if (action === 'stepIn') run.stepIn();
      else if (action === 'stepOut') run.stepOut();
    }
    return run;
  }
}
