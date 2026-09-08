/**
 * Formula tokenizer and parser.
 *
 * Produces an AST from formula text. Operator precedence follows the spreadsheet
 * convention, which differs from most programming languages in two ways worth
 * knowing:
 *
 *   - `^` binds TIGHTER than unary minus, so `-2^2` is `-4`, not `4`.
 *   - `%` is a POSTFIX operator: `50%` is `0.5`, and `A1%` is `A1/100`.
 *
 * Precedence, loosest to tightest:
 *   comparison  =  <>  <  >  <=  >=
 *   concat      &
 *   additive    +  -
 *   multiplic.  *  /
 *   power       ^
 *   unary       -x  +x
 *   postfix     x%
 *   reference   A1  A1:B2  Sheet!A1  NAME
 *
 * Reference syntax handled: relative (A1), absolute ($A$1), mixed ($A1, A$1),
 * ranges (A1:B2), sheet-qualified (Sheet1!A1, 'My Sheet'!A1), whole columns
 * (A:A) and rows (1:1), and defined names.
 */
import { ERR } from './values.js';

// ---- tokenizer -----------------------------------------------------------

const T = {
  NUMBER: 'number', STRING: 'string', BOOL: 'bool', ERROR: 'error',
  REF: 'ref', NAME: 'name', FUNC: 'func', STRUCT: 'struct',
  OP: 'op', LPAREN: '(', RPAREN: ')', COMMA: ',', COLON: ':', PERCENT: '%',
};

const REF_RE = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w(])/;
const COL_RANGE_RE = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(])/;
const ROW_RANGE_RE = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\w(])/;
const NAME_RE = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/;
const ERROR_RE = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|CIRCULAR!|SPILL!|CALC!)/;

export function tokenize(input) {
  const src = String(input).startsWith('=') ? String(input).slice(1) : String(input);
  const tokens = [];
  let i = 0;

  const prev = () => tokens[tokens.length - 1];
  // A '-' is unary when nothing value-like precedes it.
  const expectsValue = () => {
    const p = prev();
    return !p || p.type === T.OP || p.type === T.LPAREN || p.type === T.COMMA || p.type === T.COLON;
  };

  // The text inside a structured reference's OUTER brackets, from `[` at
  // `from`. Brackets nest one level (`[[#Data],[Qty]]`) and `'` escapes the
  // next character inside a column name (`'[`, `']`, `''`).
  const scanBrackets = (from) => {
    let depth = 0;
    let j = from;
    while (j < src.length) {
      const c = src[j];
      if (c === "'") { j += 2; continue; }
      if (c === '[') depth += 1;
      else if (c === ']') {
        depth -= 1;
        if (depth === 0) return { inner: src.slice(from + 1, j), end: j + 1 };
      }
      j += 1;
    }
    throw ERR.VALUE('unterminated "[" in a table reference');
  };

  while (i < src.length) {
    const ch = src[i];

    if (/\s/.test(ch)) { i += 1; continue; }

    // A bare `[` opens a this-table reference — `[@Qty]` inside the table
    // the formula lives in.
    if (ch === '[') {
      const { inner, end } = scanBrackets(i);
      tokens.push({ type: T.STRUCT, table: null, inner });
      i = end;
      continue;
    }

    if (ch === '"') {
      let j = i + 1;
      let out = '';
      while (j < src.length) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') { out += '"'; j += 2; continue; } // "" is an escaped quote
          break;
        }
        out += src[j];
        j += 1;
      }
      if (j >= src.length) throw ERR.VALUE('unterminated string');
      tokens.push({ type: T.STRING, value: out });
      i = j + 1;
      continue;
    }

    const errMatch = ERROR_RE.exec(src.slice(i));
    if (errMatch) {
      tokens.push({ type: T.ERROR, value: errMatch[0] });
      i += errMatch[0].length;
      continue;
    }

    if (/\d/.test(ch) || (ch === '.' && /\d/.test(src[i + 1] ?? ''))) {
      const m = /^\d*\.?\d+([eE][+-]?\d+)?/.exec(src.slice(i));
      tokens.push({ type: T.NUMBER, value: Number(m[0]) });
      i += m[0].length;
      continue;
    }

    // References must be tried before names, and the widest form first.
    const slice = src.slice(i);
    const colRange = COL_RANGE_RE.exec(slice);
    if (colRange) {
      tokens.push({ type: T.REF, kind: 'colRange', text: colRange[0] });
      i += colRange[0].length;
      continue;
    }
    const rowRange = ROW_RANGE_RE.exec(slice);
    if (rowRange && /^[^A-Za-z]/.test(slice)) {
      tokens.push({ type: T.REF, kind: 'rowRange', text: rowRange[0] });
      i += rowRange[0].length;
      continue;
    }
    const ref = REF_RE.exec(slice);
    if (ref) {
      // `A1#` is the spill-range operator: the whole range A1 spills into.
      const spill = src[i + ref[0].length] === '#';
      tokens.push({ type: T.REF, kind: 'cell', text: ref[0], spill });
      i += ref[0].length + (spill ? 1 : 0);
      continue;
    }

    if (/[A-Za-z_\\]/.test(ch)) {
      const m = NAME_RE.exec(slice);
      let name = m[0];
      let j = i + name.length;
      // A name butted straight against `[` is a table reference: Table1[Qty].
      if (src[j] === '[') {
        const { inner, end } = scanBrackets(j);
        tokens.push({ type: T.STRUCT, table: name, inner });
        i = end;
        continue;
      }
      while (j < src.length && /\s/.test(src[j])) j += 1;
      if (src[j] === '(') {
        // Excel prefixes a function newer than 2007 with `_xlfn.` in the file
        // (and a worksheet-only one with `_xlws.`) so an old Excel shows a
        // #NAME? rather than a wrong number. The prefix is the file's, not
        // the function's: CEILING.MATH is CEILING.MATH.
        tokens.push({ type: T.FUNC, value: name.toUpperCase().replace(/^_XL(FN|WS|PM|LM|ETN|DLM|OP)\./, '') });

        tokens.push({ type: T.LPAREN });
        i = j + 1;
        continue;
      }
      // Sheet-qualified names, e.g. Budget!Total
      if (src[j] === '!') {
        const after = NAME_RE.exec(src.slice(j + 1));
        if (after) {
          name = name + '!' + after[0];
          j = j + 1 + after[0].length;
        }
      }
      const upper = name.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') tokens.push({ type: T.BOOL, value: upper === 'TRUE' });
      else tokens.push({ type: T.NAME, value: name });
      i = j;
      continue;
    }

    if (ch === '(') { tokens.push({ type: T.LPAREN }); i += 1; continue; }
    if (ch === ')') { tokens.push({ type: T.RPAREN }); i += 1; continue; }
    if (ch === ',' || ch === ';') { tokens.push({ type: T.COMMA }); i += 1; continue; }
    if (ch === ':') { tokens.push({ type: T.COLON }); i += 1; continue; }
    if (ch === '%') { tokens.push({ type: T.PERCENT }); i += 1; continue; }

    const two = src.slice(i, i + 2);
    if (['<=', '>=', '<>'].includes(two)) { tokens.push({ type: T.OP, value: two }); i += 2; continue; }
    if ('+-*/^&=<>'.includes(ch)) {
      const unary = (ch === '-' || ch === '+') && expectsValue();
      tokens.push({ type: T.OP, value: ch, unary });
      i += 1;
      continue;
    }

    throw ERR.VALUE('unexpected character "' + ch + '" at position ' + i);
  }
  return tokens;
}

// ---- reference parsing ---------------------------------------------------

export function colToIndex(letters) {
  let n = 0;
  for (const c of letters.toUpperCase()) n = n * 26 + (c.charCodeAt(0) - 64);
  return n - 1;
}
export function indexToCol(n) {
  let s = '';
  let v = n + 1;
  while (v > 0) {
    const r = (v - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    v = Math.floor((v - 1) / 26);
  }
  return s;
}

const unquoteSheet = (q, plain) => (q ? q.replace(/''/g, "'") : plain) ?? null;

/** Parse reference text into a structured node. */
export function parseReference(text) {
  const colRange = COL_RANGE_RE.exec(text);
  if (colRange && colRange[0] === text) {
    return {
      type: 'range',
      sheet: unquoteSheet(colRange[1], colRange[2]),
      start: { col: colToIndex(colRange[4]), row: 0, colAbs: !!colRange[3], rowAbs: true },
      end: { col: colToIndex(colRange[6]), row: 1048575, colAbs: !!colRange[5], rowAbs: true },
      wholeColumn: true,
    };
  }
  const rowRange = ROW_RANGE_RE.exec(text);
  if (rowRange && rowRange[0] === text) {
    return {
      type: 'range',
      sheet: unquoteSheet(rowRange[1], rowRange[2]),
      start: { col: 0, row: Number(rowRange[4]) - 1, colAbs: true, rowAbs: !!rowRange[3] },
      end: { col: 16383, row: Number(rowRange[6]) - 1, colAbs: true, rowAbs: !!rowRange[5] },
      wholeRow: true,
    };
  }
  const m = REF_RE.exec(text);
  if (!m || m[0] !== text) throw ERR.REF('cannot parse reference "' + text + '"');
  return {
    type: 'cell',
    sheet: unquoteSheet(m[1], m[2]),
    col: colToIndex(m[4]),
    row: Number(m[6]) - 1,
    colAbs: !!m[3],
    rowAbs: !!m[5],
  };
}

// ---- structured references -----------------------------------------------

/** Undo the `'` escaping a structured column name uses for [ ] ' # @. */
const unescapeStructName = (s) => s.replace(/'(.)/g, '$1');

const STRUCT_AREAS = {
  '#ALL': 'all', '#DATA': 'data', '#HEADERS': 'headers', '#TOTALS': 'totals',
  '#THIS ROW': 'thisRow',
};

/**
 * The inside of a structured reference's outer brackets, as an area and a
 * column span.
 *
 * Forms honoured — the ones Excel actually writes into files:
 *   ``            the data body            `#All` etc.   an area keyword
 *   `Qty`         one column's data        `@Qty`        this row's cell
 *   `@`           this row, all columns
 *   `[#Data],[Qty]`  area + column         `[Qty]:[Total]` a column span
 *   `[#All],[A]:[B]` both together
 */
function parseStructSpec(inner) {
  const spec = { area: 'data', thisRow: false, startCol: null, endCol: null };
  const text = inner.trim();
  if (text === '') return spec;

  // Split the two-part form on top-level commas; single-part text is one item.
  const items = [];
  if (text.startsWith('[')) {
    let depth = 0;
    let current = '';
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (c === "'") { current += text.slice(i, i + 2); i += 1; continue; }
      if (c === '[') depth += 1;
      if (c === ']') depth -= 1;
      if (c === ',' && depth === 0) { items.push(current); current = ''; continue; }
      current += c;
    }
    items.push(current);
  } else {
    items.push(text);
  }

  const stripBrackets = (s) => {
    const trimmed = s.trim();
    return trimmed.startsWith('[') && trimmed.endsWith(']') ? trimmed.slice(1, -1) : trimmed;
  };

  for (const rawItem of items) {
    const trimmed = rawItem.trim();
    // A column span: [A]:[B] at the top level of one item.
    const span = /^\[([\s\S]*?)\]\s*:\s*\[([\s\S]*?)\]$/.exec(trimmed);
    if (span) {
      spec.startCol = unescapeStructName(span[1]);
      spec.endCol = unescapeStructName(span[2]);
      continue;
    }
    const item = stripBrackets(trimmed);
    const area = STRUCT_AREAS[item.toUpperCase()];
    if (area === 'thisRow') { spec.thisRow = true; continue; }
    if (area) { spec.area = area; continue; }
    if (item.startsWith('@')) {
      spec.thisRow = true;
      const col = item.slice(1);
      if (col) spec.startCol = unescapeStructName(col);
      continue;
    }
    spec.startCol = unescapeStructName(item);
  }
  if (spec.endCol && !spec.startCol) spec.startCol = spec.endCol;
  if (spec.startCol && !spec.endCol) spec.endCol = spec.startCol;
  return spec;
}

// ---- parser --------------------------------------------------------------

const COMPARISON = ['=', '<>', '<', '>', '<=', '>='];

export function parse(input) {
  const tokens = tokenize(input);
  let pos = 0;

  const peek = () => tokens[pos];
  const next = () => tokens[pos++];
  const at = (type, value) => peek() && peek().type === type && (value === undefined || peek().value === value);
  const expect = (type) => {
    if (!at(type)) throw ERR.VALUE('expected ' + type + ' but found ' + (peek() ? (peek().value ?? peek().type) : 'end of formula'));
    return next();
  };

  function parseComparison() {
    let left = parseConcat();
    while (peek() && peek().type === T.OP && COMPARISON.includes(peek().value)) {
      const op = next().value;
      left = { type: 'binary', op, left, right: parseConcat() };
    }
    return left;
  }
  function parseConcat() {
    let left = parseAdditive();
    while (at(T.OP, '&')) {
      next();
      left = { type: 'binary', op: '&', left, right: parseAdditive() };
    }
    return left;
  }
  function parseAdditive() {
    let left = parseMultiplicative();
    while (peek() && peek().type === T.OP && !peek().unary && (peek().value === '+' || peek().value === '-')) {
      const op = next().value;
      left = { type: 'binary', op, left, right: parseMultiplicative() };
    }
    return left;
  }
  function parseMultiplicative() {
    let left = parseUnary();
    while (peek() && peek().type === T.OP && (peek().value === '*' || peek().value === '/')) {
      const op = next().value;
      left = { type: 'binary', op, left, right: parseUnary() };
    }
    return left;
  }
  function parseUnary() {
    if (peek() && peek().type === T.OP && peek().unary) {
      const op = next().value;
      // `^` binds tighter than unary minus: -2^2 === -4
      return { type: 'unary', op, operand: parseUnary() };
    }
    return parsePower();
  }
  function parsePower() {
    const base = parsePostfix();
    if (at(T.OP, '^')) {
      next();
      // right-associative, and its right side may itself be unary: 2^-1
      return { type: 'binary', op: '^', left: base, right: parseUnary() };
    }
    return base;
  }
  function parsePostfix() {
    let node = parsePrimary();
    while (at(T.PERCENT)) {
      next();
      node = { type: 'unary', op: '%', operand: node };
    }
    return node;
  }

  function parsePrimary() {
    const t = peek();
    if (!t) throw ERR.VALUE('formula ended unexpectedly');

    if (t.type === T.NUMBER) { next(); return { type: 'literal', value: t.value }; }
    if (t.type === T.STRING) { next(); return { type: 'literal', value: t.value }; }
    if (t.type === T.BOOL) { next(); return { type: 'literal', value: t.value }; }
    if (t.type === T.ERROR) { next(); return { type: 'errorLiteral', value: t.value }; }

    if (t.type === T.LPAREN) {
      next();
      const inner = parseComparison();
      expect(T.RPAREN);
      return inner;
    }

    if (t.type === T.FUNC) {
      const name = next().value;
      expect(T.LPAREN);
      const args = [];
      if (!at(T.RPAREN)) {
        args.push(parseComparison());
        while (at(T.COMMA)) { next(); args.push(parseComparison()); }
      }
      expect(T.RPAREN);
      return { type: 'call', name, args };
    }

    if (t.type === T.REF) {
      const first = next();
      const startNode = parseReference(first.text);
      // A spill reference stands alone — it IS a range, so it heads no other.
      if (first.spill) return { ...startNode, type: 'spillref' };
      if (at(T.COLON) && startNode.type === 'cell') {
        next();
        const endTok = expect(T.REF);
        const endNode = parseReference(endTok.text);
        if (endNode.type !== 'cell') throw ERR.REF('bad range end "' + endTok.text + '"');
        return {
          type: 'range',
          sheet: startNode.sheet ?? endNode.sheet,
          start: startNode,
          end: endNode,
        };
      }
      return startNode;
    }

    if (t.type === T.NAME) { next(); return { type: 'name', name: t.value }; }

    if (t.type === T.STRUCT) {
      next();
      return { type: 'structref', table: t.table, ...parseStructSpec(t.inner) };
    }

    throw ERR.VALUE('unexpected ' + (t.value ?? t.type));
  }

  const ast = parseComparison();
  if (pos < tokens.length) {
    throw ERR.VALUE('unexpected trailing input near "' + (peek().value ?? peek().type) + '"');
  }
  return ast;
}

/**
 * Shift a formula's RELATIVE references by (dr, dc) — what copying a formula
 * one cell down or right means. `$` pins an axis exactly as it does in a
 * spreadsheet, quoted strings and quoted sheet names pass through untouched,
 * and a reference pushed off the sheet's edge becomes `#REF!`, which is the
 * honest answer rather than a wrapped-around cell.
 *
 * Text-level rather than AST-level, deliberately: the result must preserve
 * every space, comma and function name exactly as typed, because this string
 * goes back into the cell as what the user "wrote".
 */
export function shiftFormula(formula, dr, dc) {
  const MAX_ROWS = 1048576;
  const MAX_COLS = 16384;
  const src = String(formula);
  let out = '';
  let i = 0;

  const copyQuoted = (quote) => {
    let j = i + 1;
    while (j < src.length) {
      if (src[j] === quote) {
        if (src[j + 1] === quote) { j += 2; continue; } // doubled quote is an escape
        break;
      }
      j += 1;
    }
    out += src.slice(i, Math.min(j + 1, src.length));
    i = j + 1;
  };

  while (i < src.length) {
    const ch = src[i];
    if (ch === '"' || ch === "'") { copyQuoted(ch); continue; }

    // A structured reference's brackets hold column LABELS, not references —
    // [B2] names a column called B2 — so everything to the matching `]`
    // passes through verbatim. `'` escapes one character inside, and the
    // two-part form nests brackets one level.
    if (ch === '[') {
      let depth = 0;
      let j = i;
      while (j < src.length) {
        const c = src[j];
        if (c === "'") { j += 2; continue; }
        if (c === '[') depth += 1;
        else if (c === ']') { depth -= 1; if (depth === 0) { j += 1; break; } }
        j += 1;
      }
      out += src.slice(i, j);
      i = j;
      continue;
    }

    // A reference must not begin mid-name: after `LOG` the `10(` is a call,
    // and after `TAX_` the `A1` is part of a defined name.
    const prev = out[out.length - 1] ?? '';
    if (!/[A-Za-z0-9_.$]/.test(prev)) {
      const slice = src.slice(i);

      const cell = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w(])/.exec(slice);
      if (cell) {
        let col = colToIndex(cell[2]);
        let row = Number(cell[4]) - 1;
        if (!cell[1]) col += dc;
        if (!cell[3]) row += dr;
        if (row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLS) {
          out += '#REF!';
        } else {
          out += cell[1] + indexToCol(col) + cell[3] + (row + 1);
        }
        i += cell[0].length;
        continue;
      }

      const colRange = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(])/.exec(slice);
      if (colRange) {
        const a = colToIndex(colRange[2]) + (colRange[1] ? 0 : dc);
        const b = colToIndex(colRange[4]) + (colRange[3] ? 0 : dc);
        if (a < 0 || b < 0 || a >= MAX_COLS || b >= MAX_COLS) out += '#REF!';
        else out += colRange[1] + indexToCol(a) + ':' + colRange[3] + indexToCol(b);
        i += colRange[0].length;
        continue;
      }

      const rowRange = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\w(])/.exec(slice);
      if (rowRange) {
        const a = Number(rowRange[2]) - 1 + (rowRange[1] ? 0 : dr);
        const b = Number(rowRange[4]) - 1 + (rowRange[3] ? 0 : dr);
        if (a < 0 || b < 0 || a >= MAX_ROWS || b >= MAX_ROWS) out += '#REF!';
        else out += rowRange[1] + (a + 1) + ':' + rowRange[3] + (b + 1);
        i += rowRange[0].length;
        continue;
      }
    }

    out += ch;
    i += 1;
  }
  return out;
}

/** Every cell and range a formula depends on. Drives the recalculation graph. */
export function dependencies(ast, out = []) {
  if (!ast || typeof ast !== 'object') return out;
  if (ast.type === 'cell' || ast.type === 'range' || ast.type === 'name'
    || ast.type === 'structref' || ast.type === 'spillref') out.push(ast);
  for (const key of ['left', 'right', 'operand']) if (ast[key]) dependencies(ast[key], out);
  // ROW(ref) and COLUMN(ref) read a reference's POSITION, never its value —
  // the evaluator answers them from the AST. Counting the reference as a
  // dependency here would invent edges (and with them, false cycles: a cell
  // asking =COLUMN(D1) does not depend on what D1 holds).
  if (ast.type === 'call' && (ast.name === 'ROW' || ast.name === 'COLUMN')) return out;
  // OFFSET's anchor is a POSITION too — the evaluator reads it from the AST,
  // and the function is volatile so recalculation covers wherever it lands.
  // Walking the anchor as a value dependency would make =OFFSET(A1,1,0) in A1
  // a false cycle. The remaining arguments are ordinary values.
  if (ast.type === 'call' && ast.name === 'OFFSET') {
    for (const a of (ast.args ?? []).slice(1)) dependencies(a, out);
    return out;
  }
  for (const a of ast.args ?? []) dependencies(a, out);
  return out;
}

export { T as TOKEN };
