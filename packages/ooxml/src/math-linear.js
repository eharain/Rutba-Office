/**
 * Word's linear format — UnicodeMath — into Office Math (OMML), the way Word
 * writes it.
 *
 * What a person types in Word's equation editor is a line of text:
 * `x=(-b±√(b^2-4ac))/2a`, `\sum_(i=1)^n i`, `\int_0^1 f(x)dx`, `a_i`, `\alpha`.
 * Word "builds it up" into a two-dimensional equation and stores it as OMML;
 * this does the same, and writes the OMML the way Word writes it — every
 * `m:r` in Cambria Math, every structure's `m:ctrlPr` italic Cambria Math, a
 * display equation in `m:oMathPara` — so a file saved here opens in Word as
 * an equation Word itself could have made.
 *
 * The grammar, as Word's:
 *   - an OPERAND is a run of letters and digits, a bracketed group, or a
 *     structure; a space ends one (`\sum_(i=1)^n i`: the `n` ends at the
 *     space, `i` is the sum's body);
 *   - `/` makes a fraction of the operands either side, and the brackets
 *     round an operand are dropped (`(a+b)/2`); `¦` a fraction with no bar
 *     (a binomial), `∕` a skewed one, `⊘` a linear one;
 *   - `^` and `_` take the operand after them as a superscript/subscript,
 *     brackets dropped; a script's base keeps its brackets (`(a+b)^2`);
 *   - `√(x)`, `√(n&x)` (an nth root), `∛`, `∜`;
 *   - a large operator (∑ ∏ ∫ ∬ ∮ …) takes its limits as scripts and its body
 *     after `▒`, or as the terms that follow it up to a `+`, `=` and the
 *     like;
 *   - `sin`, `cos`, `log`, `lim` … are functions: upright, their argument the
 *     operand after them; `lim_(n→∞)` puts the limit under the name;
 *   - `■(a&b@c&d)` is a matrix (`&` between cells, `@` between rows), and
 *     `(a&b@c&d)` a matrix in brackets; `█(x&=1@y&=2)` an equation array;
 *   - a combining accent after an operand (`x\u0302`, or `x\hat`) is an accent;
 *     `¯(x)` an overbar, `▁(x)` an underbar, `⏟(x)` a brace under;
 *     `a┬b` puts b below a, `a┴b` above;
 *   - `\alpha`, `\pm`, `\le`, `\infty`, `\rightarrow` … are the symbols of
 *     those names; `"text"` is normal text;
 *   - `〖…〗` groups without drawing brackets.
 *
 * Nothing here throws on bad input: an unknown keyword or an unclosed
 * bracket comes back as `{ ok: false, error, at }` for the editor to show.
 */
import { FUNCTION_NAMES } from './math.js';

/* ── keywords ──────────────────────────────────────────────────────────── */

const GREEK = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ϵ', varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ',
  iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', omicron: 'ο', pi: 'π', varpi: 'ϖ', rho: 'ρ', varrho: 'ϱ',
  sigma: 'σ', varsigma: 'ς', tau: 'τ', upsilon: 'υ', phi: 'ϕ', varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
};

/** Every `\name` the editor understands, and what it stands for. */
export const KEYWORDS = {
  ...GREEK,
  pm: '±', mp: '∓', times: '×', div: '÷', cdot: '⋅', ast: '∗', star: '⋆', circ: '∘', bullet: '∙',
  le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠', approx: '≈', equiv: '≡', sim: '∼', simeq: '≃', cong: '≅', propto: '∝',
  ll: '≪', gg: '≫', infty: '∞', partial: '∂', nabla: '∇', in: '∈', notin: '∉', ni: '∋', subset: '⊂', supset: '⊃',
  subseteq: '⊆', supseteq: '⊇', cup: '∪', cap: '∩', setminus: '∖', forall: '∀', exists: '∃', neg: '¬', wedge: '∧', vee: '∨',
  to: '→', rightarrow: '→', leftarrow: '←', leftrightarrow: '↔', Rightarrow: '⇒', Leftarrow: '⇐', Leftrightarrow: '⇔',
  uparrow: '↑', downarrow: '↓', mapsto: '↦', implies: '⟹', iff: '⟺',
  degree: '°', prime: '′', cdots: '⋯', ldots: '…', dots: '…', vdots: '⋮', ddots: '⋱', emptyset: '∅', hbar: 'ℏ', ell: 'ℓ',
  angle: '∠', perp: '⊥', parallel: '∥', therefore: '∴', because: '∵', aleph: 'ℵ', Re: 'ℜ', Im: 'ℑ',
  sqrt: '√', cbrt: '∛', qdrt: '∜', sum: '∑', prod: '∏', coprod: '∐', int: '∫', iint: '∬', iiint: '∭', oint: '∮', oiint: '∯',
  bigcup: '⋃', bigcap: '⋂', matrix: '■', eqarray: '█', box: '□', rect: '▭', below: '┬', above: '┴',
  overbar: '¯', underbar: '▁', overbrace: '⏞', underbrace: '⏟', phantom: '⟡', funcapply: '\u2061', naryand: '▒', of: '▒',
  hat: '\u0302', tilde: '\u0303', bar: '\u0305', vec: '\u20D7', dot: '\u0307', ddot: '\u0308', check: '\u030C', breve: '\u0306',
  acute: '\u0301', grave: '\u0300',
  langle: '⟨', rangle: '⟩', lfloor: '⌊', rfloor: '⌋', lceil: '⌈', rceil: '⌉', norm: '‖', lbrace: '{', rbrace: '}', mid: '│',
  atop: '¦', over: '/',
};

/* ── characters ────────────────────────────────────────────────────────── */

const OPENERS = { '(': ')', '[': ']', '{': '}', '⟨': '⟩', '⌊': '⌋', '⌈': '⌉', '|': '|', '‖': '‖', '〖': '〗', '├': null };
const CLOSERS = new Set([')', ']', '}', '⟩', '⌋', '⌉', '〗', '┤']);
const NARY = new Set(['∑', '∏', '∐', '∫', '∬', '∭', '∮', '∯', '∰', '⋃', '⋂', '⨁', '⨂']);
const INTEGRALS = new Set(['∫', '∬', '∭', '∮', '∯', '∰']);
const FRACTIONS = { '/': 'bar', '¦': 'noBar', '∕': 'skw', '⊘': 'lin' };
/** What ends a large operator's body when it has no `▒`: the next term. */
const BODY_ENDS = new Set(['+', '-', '−', '=', '<', '>', '≤', '≥', '≠', '±', '∓', ',', ';', '→', '≈', '≡', '∼', '⇒', '⇔']);
const PREFIX = { '□': 'box', '▭': 'borderBox', '⟡': 'phant', '¯': 'barTop', '▁': 'barBot', '⏟': 'groupBot', '⏞': 'groupTop', '⎵': 'groupBot', '⎴': 'groupTop' };
const isLetter = (c) => /\p{L}/u.test(c);
const isDigit = (c) => c >= '0' && c <= '9';
const isAlnum = (c) => isLetter(c) || isDigit(c);
const isCombining = (c) => /[\u0300-\u036F\u20D0-\u20FF]/.test(c);
/** Characters that stand for a quantity, not an operation: they join an operand. */
const OPERAND_SYMBOLS = new Set(['∞', '∂', '∇', 'ℏ', '∅', 'ℓ', '…', '⋯', '′', '″', "'", '!', '°', '⋮', '⋱', 'ℵ', 'ℜ', 'ℑ']);

class LinearError extends Error {
  constructor(message, at) { super(message); this.at = at; }
}

/** The input with every `\keyword` replaced by its symbol, each character keeping where it came from. */
function tokenize(src) {
  const chars = [];
  const s = String(src ?? '');
  for (let i = 0; i < s.length;) {
    const cp = s.codePointAt(i);
    const c = String.fromCodePoint(cp);
    if (c === '\\') {
      const m = /^[A-Za-z]+/.exec(s.slice(i + 1));
      if (m) {
        const sym = KEYWORDS[m[0]];
        if (sym === undefined) throw new LinearError(`\\${m[0]} is not a keyword this editor knows`, i);
        chars.push({ c: sym, at: i });
        i += 1 + m[0].length;
        // A space after a keyword only ends the keyword, as in Word.
        if (s[i] === ' ' && !isCombining(sym)) i += 1;
        continue;
      }
      if (i + 1 < s.length) { chars.push({ c: s[i + 1], at: i }); i += 2; continue; }
      i += 1;
      continue;
    }
    if (c === '\r' || c === '\n' || c === '\t') { chars.push({ c: ' ', at: i }); i += c.length; continue; }
    chars.push({ c, at: i });
    i += c.length;
  }
  return chars;
}

/* ── the parser ────────────────────────────────────────────────────────── */

/**
 * AST nodes: `{ t: 'char', c }` (one letter, digit run or operator),
 * `{ t: 'nor', s }`, `{ t: 'space' }`, and the structures — `f`, `rad`,
 * `script`, `pre`, `nary`, `d`, `grp`, `func`, `acc`, `bar`, `group`,
 * `lim`, `m`, `eqArr`, `box`, `borderBox`, `phant` — each holding rows
 * (arrays of nodes).
 */
class Parser {
  constructor(chars) { this.chars = chars; this.pos = 0; }
  peek(k = 0) { return this.chars[this.pos + k]?.c; }
  at() { return this.chars[this.pos]?.at ?? (this.chars.length ? this.chars[this.chars.length - 1].at + 1 : 0); }
  next() { return this.chars[this.pos++]?.c; }
  skipSpaces() { while (this.peek() === ' ') this.pos++; }

  /** A row up to (not including) one of `stops`, fractions built. */
  row(stops) {
    const items = [];
    while (this.pos < this.chars.length) {
      const c = this.peek();
      if (stops.has(c)) break;
      if (c === ' ') { this.pos++; items.push({ t: 'space' }); continue; }
      items.push(this.postfixed(stops));
    }
    return finishRow(items);
  }

  /** One operand and what follows it: scripts, an accent, a limit below or above. */
  postfixed(stops) {
    let base = this.primary(stops);
    for (;;) {
      const c = this.peek();
      if ((c === '^' || c === '_') && isOperand(base)) {
        base = this.scripts(base);
        continue;
      }
      if (c && isCombining(c) && isOperand(base)) {
        this.pos++;
        base = { t: 'acc', chr: c, e: unwrap(base) };
        continue;
      }
      if ((c === '┬' || c === '┴') && isOperand(base)) {
        this.pos++;
        base = { t: 'lim', pos: c === '┬' ? 'low' : 'upp', e: [base], lim: this.argument() };
        continue;
      }
      return base;
    }
  }

  /** `^x`, `_x`, or both in either order, on `base`. */
  scripts(base) {
    let sub = null;
    let sup = null;
    while ((this.peek() === '_' && sub === null) || (this.peek() === '^' && sup === null)) {
      const which = this.next();
      const arg = this.argument();
      if (which === '_') sub = arg; else sup = arg;
    }
    return { t: 'script', e: [base], sub, sup };
  }

  /**
   * A script's, a radical's, a limit's argument: one operand with its own
   * brackets dropped — `(i=1)` is `i=1` — or a run of letters and digits,
   * or a signed operand (`e^-x`).
   */
  argument() {
    const c = this.peek();
    if (c === undefined || c === ' ') return [];
    if (c === '(' || c === '〖') {
      const g = this.bracket();
      return g.t === 'd' && g.beg === '(' && g.end === ')' && g.parts.length === 1 ? g.parts[0] : g.t === 'grp' ? g.e : [g];
    }
    if (c === '-' || c === '+' || c === '−' || c === '±') {
      this.pos++;
      return [{ t: 'char', c }, ...this.argument()];
    }
    if (isAlnum(c)) {
      const out = [];
      while (this.peek() !== undefined && isAlnum(this.peek())) {
        if (isDigit(this.peek())) out.push(this.number());
        else out.push({ t: 'char', c: this.next() });
      }
      return out;
    }
    return [this.postfixed(new Set())];
  }

  number() {
    let n = '';
    while (this.peek() !== undefined && (isDigit(this.peek()) || (this.peek() === '.' && isDigit(this.peek(1) ?? '')))) n += this.next();
    return { t: 'char', c: n };
  }

  /** What an operand starts with. */
  primary(stops) {
    const c = this.peek();
    if (c in OPENERS) return this.bracket();
    if (CLOSERS.has(c)) { this.pos++; return { t: 'char', c }; } // a stray closing bracket is only a character
    if (c === '"') {
      const from = this.at();
      this.pos++;
      let s = '';
      while (this.peek() !== undefined && this.peek() !== '"') s += this.next();
      if (this.peek() !== '"') throw new LinearError('the text in quotes is not closed', from);
      this.pos++;
      return { t: 'nor', s };
    }
    if (c === '√' || c === '∛' || c === '∜') {
      this.pos++;
      if (c === '√' && this.peek() === '(') {
        const from = this.at();
        this.pos++;
        const first = this.row(new Set([')', '&']));
        if (this.peek() === '&') {
          this.pos++;
          const e = this.row(new Set([')']));
          if (this.next() !== ')') throw new LinearError('a bracket opened here is not closed', from);
          return { t: 'rad', deg: first, e };
        }
        if (this.next() !== ')') throw new LinearError('a bracket opened here is not closed', from);
        return { t: 'rad', deg: null, e: first };
      }
      return { t: 'rad', deg: c === '√' ? null : [{ t: 'char', c: c === '∛' ? '3' : '4' }], e: this.argument() };
    }
    if (NARY.has(c)) return this.nary(stops);
    if (c === '■' || c === '█') {
      this.pos++;
      const from = this.at();
      if (this.next() !== '(') throw new LinearError(c === '■' ? 'a matrix is written ■(a&b@c&d)' : 'an equation array is written █(a&=b@c&=d)', from);
      return c === '■' ? { t: 'm', rows: this.cells(from) } : { t: 'eqArr', rows: this.lines(from) };
    }
    if (c in PREFIX) {
      this.pos++;
      const kind = PREFIX[c];
      const e = this.argument();
      if (kind === 'barTop' || kind === 'barBot') return { t: 'bar', pos: kind === 'barTop' ? 'top' : 'bot', e };
      if (kind === 'groupTop' || kind === 'groupBot') return { t: 'group', chr: c, pos: kind === 'groupTop' ? 'top' : 'bot', e };
      return { t: kind, e };
    }
    if ((c === '_' || c === '^')) {
      // Scripts with nothing before them go before what follows: ₆¹⁴C.
      const s = this.scripts({ t: 'grp', e: [] });
      this.skipSpaces();
      return { t: 'pre', sub: s.sub, sup: s.sup, e: this.argument() };
    }
    if (isDigit(c)) return this.number();
    if (isLetter(c)) {
      const name = this.functionName();
      if (name) return this.func(name);
      this.pos++;
      return { t: 'char', c };
    }
    this.pos++;
    return { t: 'char', c };
  }

  /** The longest function name spelt from here, when it is a whole word. */
  functionName() {
    let word = '';
    for (let k = 0; isLetter(this.peek(k) ?? ''); k++) word += this.peek(k);
    for (let len = word.length; len >= 2; len--) {
      const w = word.slice(0, len);
      if (FUNCTION_NAMES.has(w) && len === word.length) return w;
    }
    return null;
  }

  func(name) {
    this.pos += name.length;
    let fName = [{ t: 'char', c: name, upright: true }];
    const under = name === 'lim' || name === 'max' || name === 'min' || name === 'sup' || name === 'inf';
    if (this.peek() === '_' || this.peek() === '┬') {
      const below = this.next() === '┬';
      const lim = this.argument();
      // `lim_(n→∞)` sets the limit under the name, as Word does; `log_2` is a subscript.
      fName = under || below ? [{ t: 'lim', pos: 'low', e: fName, lim }] : [{ t: 'script', e: fName, sub: lim, sup: null }];
    }
    if (this.peek() === '^') {
      this.pos++;
      fName = [{ t: 'script', e: fName, sub: null, sup: this.argument() }];
    }
    this.skipSpaces();
    if (this.peek() === '\u2061') { this.pos++; this.skipSpaces(); }
    const c = this.peek();
    const e = c === undefined || CLOSERS.has(c) || c === '&' || c === '@' ? [] : unwrapInvisible(this.postfixed(new Set()));
    return { t: 'func', name: fName, e };
  }

  nary(stops) {
    const chr = this.next();
    let sub = null;
    let sup = null;
    while ((this.peek() === '_' && sub === null) || (this.peek() === '^' && sup === null)) {
      const which = this.next();
      const arg = this.argument();
      if (which === '_') sub = arg; else sup = arg;
    }
    this.skipSpaces();
    let e;
    if (this.peek() === '▒') {
      this.pos++;
      this.skipSpaces();
      e = unwrapInvisible(this.postfixed(new Set()));
    } else {
      // No ▒: the body is the term that follows, up to a `+`, `=` or the like.
      e = this.row(new Set([...stops, ...BODY_ENDS]));
    }
    return { t: 'nary', chr, sub, sup, e };
  }

  /** A bracketed group, its delimiters kept: `(…)`, `[…]`, `|…|`, `〖…〗` (drawn without any). */
  bracket() {
    const from = this.at();
    const open = this.next();
    const match = OPENERS[open];
    const closers = open === '├' ? new Set([...CLOSERS, '|', '‖']) : new Set([match, '┤', ...(open === '|' || open === '‖' ? [] : [])]);
    const stops = new Set([...closers, '&', '@', '│']);
    const parts = [this.row(stops)];
    let matrix = null;
    if (this.peek() === '&' || this.peek() === '@') {
      // A matrix in brackets: (a&b@c&d).
      const rows = [[parts[0]]];
      while (this.peek() === '&' || this.peek() === '@') {
        const sep = this.next();
        if (sep === '@') rows.push([]);
        rows[rows.length - 1].push(this.row(stops));
      }
      matrix = rows;
    } else {
      while (this.peek() === '│') { this.pos++; parts.push(this.row(stops)); }
    }
    const close = this.peek();
    if (close === undefined || !closers.has(close)) throw new LinearError(`the ${open} opened here is not closed`, from);
    this.pos++;
    if (open === '〖') return { t: 'grp', e: parts[0] };
    const d = {
      t: 'd',
      beg: open === '├' ? '' : open,
      end: close === '┤' ? '' : close,
      parts: matrix ? [[{ t: 'm', rows: matrix }]] : parts,
    };
    if (parts.length > 1) d.sep = '│';
    return d;
  }

  /** A matrix's rows of cells, after `■(`. */
  cells(from) {
    const stops = new Set(['&', '@', ')']);
    const rows = [[this.row(stops)]];
    while (this.peek() === '&' || this.peek() === '@') {
      if (this.next() === '@') rows.push([]);
      rows[rows.length - 1].push(this.row(stops));
    }
    if (this.next() !== ')') throw new LinearError('the matrix opened here is not closed', from);
    return rows;
  }

  /** An equation array's lines, after `█(` — the `&` stays in the text to align on. */
  lines(from) {
    const stops = new Set(['@', ')']);
    const rows = [this.row(stops)];
    while (this.peek() === '@') { this.pos++; rows.push(this.row(stops)); }
    if (this.next() !== ')') throw new LinearError('the equation array opened here is not closed', from);
    return rows;
  }
}

/** Is a node something a script, an accent or a fraction can take hold of? */
function isOperand(node) {
  if (!node || node.t === 'space') return false;
  if (node.t !== 'char') return true;
  const c = node.c;
  return isAlnum(c[0]) || OPERAND_SYMBOLS.has(c) || isLetter(c);
}
/** A group's brackets dropped: `(a+b)` as the numerator of a fraction is `a+b`. */
function unwrap(node) {
  if (node.t === 'd' && node.beg === '(' && node.end === ')' && node.parts.length === 1) return node.parts[0];
  if (node.t === 'grp') return node.e;
  return [node];
}
/** Only the invisible 〖〗 dropped — a body or an argument keeps its parentheses. */
const unwrapInvisible = (node) => (node.t === 'grp' ? node.e : [node]);

/**
 * A row's fractions built — each `/` takes the operands either side, as far
 * as the nearest operator or space — then its spaces dropped, and a name
 * followed by the function-application mark made a function.
 */
function finishRow(items) {
  let list = items;
  for (;;) {
    const i = list.findIndex((n) => n.t === 'char' && FRACTIONS[n.c]);
    if (i < 0) break;
    let a = i;
    while (a > 0 && isOperand(list[a - 1])) a--;
    // The denominator is one bracketed group, or the operands up to the next
    // bracket: `1/2(α±β)` is a half of (α±β), as Word reads it.
    let b = i + 1;
    const bracketed = (n) => n && (n.t === 'd' || n.t === 'grp');
    if (bracketed(list[b])) b++;
    else while (b < list.length && isOperand(list[b]) && !bracketed(list[b])) b++;
    const left = list.slice(a, i);
    const right = list.slice(i + 1, b);
    const operand = (part) => (part.length === 1 ? unwrap(part[0]) : part);
    const f = { t: 'f', type: FRACTIONS[list[i].c], num: operand(left), den: operand(right) };
    list = [...list.slice(0, a), f, ...list.slice(b)];
  }
  list = list.filter((n) => n.t !== 'space');
  // `name\u2061arg` for a name Word does not know: the letters before the mark.
  for (let i = list.indexOf(list.find((n) => n.t === 'char' && n.c === '\u2061')); i > 0; i = list.findIndex((n) => n.t === 'char' && n.c === '\u2061')) {
    let a = i;
    while (a > 0 && list[a - 1].t === 'char' && isLetter(list[a - 1].c)) a--;
    const name = list.slice(a, i).map((n) => n.c).join('');
    const arg = list[i + 1] ? unwrapInvisible(list[i + 1]) : [];
    if (!name) { list = [...list.slice(0, i), ...list.slice(i + 1)]; continue; }
    list = [...list.slice(0, a), { t: 'func', name: [{ t: 'char', c: name, upright: true }], e: arg }, ...list.slice(i + 2)];
  }
  return list;
}

/* ── OMML, as Word writes it ───────────────────────────────────────────── */

const RPR = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
const CTRL = '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr></m:ctrlPr>';
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const attr = (name, value) => `<m:${name} m:val="${esc(value)}"/>`;

function run(text, mrPr = '') {
  const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
  return `<m:r>${mrPr ? `<m:rPr>${mrPr}</m:rPr>` : ''}${RPR}<m:t${space}>${esc(text)}</m:t></m:r>`;
}
const box = (name, row) => { const inner = rowXml(row); return inner ? `<m:${name}>${inner}</m:${name}>` : `<m:${name}/>`; };

/** A row as OMML: consecutive characters share one run, as Word writes `x=` and `-4ac`. */
function rowXml(row) {
  let out = '';
  let text = '';
  const flush = () => { if (text) out += run(text); text = ''; };
  for (const n of row || []) {
    if (n.t === 'char' && !n.upright) { text += n.c; continue; }
    flush();
    out += nodeXml(n);
  }
  flush();
  return out;
}

function nodeXml(n) {
  switch (n.t) {
    case 'char': return run(n.c, n.upright ? attr('sty', 'p') : '');
    case 'nor': return run(n.s, '<m:nor/>');
    case 'grp': return rowXml(n.e);
    case 'f':
      return `<m:f><m:fPr>${n.type !== 'bar' ? attr('type', n.type) : ''}${CTRL}</m:fPr>${box('num', n.num)}${box('den', n.den)}</m:f>`;
    case 'rad':
      return n.deg
        ? `<m:rad><m:radPr>${CTRL}</m:radPr>${box('deg', n.deg)}${box('e', n.e)}</m:rad>`
        : `<m:rad><m:radPr>${attr('degHide', '1')}${CTRL}</m:radPr><m:deg/>${box('e', n.e)}</m:rad>`;
    case 'script':
      if (n.sub && n.sup) return `<m:sSubSup><m:sSubSupPr>${CTRL}</m:sSubSupPr>${box('e', n.e)}${box('sub', n.sub)}${box('sup', n.sup)}</m:sSubSup>`;
      if (n.sub) return `<m:sSub><m:sSubPr>${CTRL}</m:sSubPr>${box('e', n.e)}${box('sub', n.sub)}</m:sSub>`;
      return `<m:sSup><m:sSupPr>${CTRL}</m:sSupPr>${box('e', n.e)}${box('sup', n.sup)}</m:sSup>`;
    case 'pre':
      return `<m:sPre><m:sPrePr>${CTRL}</m:sPrePr>${box('sub', n.sub || [])}${box('sup', n.sup || [])}${box('e', n.e)}</m:sPre>`;
    case 'nary': {
      const loc = INTEGRALS.has(n.chr) ? 'subSup' : 'undOvr';
      return `<m:nary><m:naryPr>${n.chr !== '∫' ? attr('chr', n.chr) : ''}${attr('limLoc', loc)}${n.sub ? '' : attr('subHide', '1')}${n.sup ? '' : attr('supHide', '1')}${CTRL}</m:naryPr>`
        + `${box('sub', n.sub || [])}${box('sup', n.sup || [])}${box('e', n.e)}</m:nary>`;
    }
    case 'd':
      return `<m:d><m:dPr>${n.beg !== '(' ? attr('begChr', n.beg) : ''}${n.sep ? attr('sepChr', n.sep) : ''}${n.end !== ')' ? attr('endChr', n.end) : ''}${CTRL}</m:dPr>`
        + `${n.parts.map((p) => box('e', p)).join('')}</m:d>`;
    case 'func':
      return `<m:func><m:funcPr>${CTRL}</m:funcPr><m:fName>${rowXml(n.name)}</m:fName>${box('e', n.e)}</m:func>`;
    case 'acc':
      return `<m:acc><m:accPr>${attr('chr', n.chr)}${CTRL}</m:accPr>${box('e', n.e)}</m:acc>`;
    case 'bar':
      return `<m:bar><m:barPr>${n.pos === 'top' ? attr('pos', 'top') : ''}${CTRL}</m:barPr>${box('e', n.e)}</m:bar>`;
    case 'group':
      return n.pos === 'top'
        ? `<m:groupChr><m:groupChrPr>${attr('chr', n.chr)}${attr('pos', 'top')}${attr('vertJc', 'bot')}${CTRL}</m:groupChrPr>${box('e', n.e)}</m:groupChr>`
        : `<m:groupChr><m:groupChrPr>${n.chr !== '⏟' ? attr('chr', n.chr) : ''}${CTRL}</m:groupChrPr>${box('e', n.e)}</m:groupChr>`;
    case 'lim':
      return n.pos === 'low'
        ? `<m:limLow><m:limLowPr>${CTRL}</m:limLowPr>${box('e', n.e)}${box('lim', n.lim)}</m:limLow>`
        : `<m:limUpp><m:limUppPr>${CTRL}</m:limUppPr>${box('e', n.e)}${box('lim', n.lim)}</m:limUpp>`;
    case 'm': {
      const cols = Math.max(1, ...n.rows.map((r) => r.length));
      return `<m:m><m:mPr><m:mcs><m:mc><m:mcPr>${attr('count', String(cols))}${attr('mcJc', 'center')}</m:mcPr></m:mc></m:mcs>${CTRL}</m:mPr>`
        + n.rows.map((r) => `<m:mr>${Array.from({ length: cols }, (_, i) => box('e', r[i] || [])).join('')}</m:mr>`).join('') + '</m:m>';
    }
    case 'eqArr':
      return `<m:eqArr><m:eqArrPr>${CTRL}</m:eqArrPr>${n.rows.map((r) => box('e', r)).join('')}</m:eqArr>`;
    case 'box': return `<m:box><m:boxPr>${CTRL}</m:boxPr>${box('e', n.e)}</m:box>`;
    case 'borderBox': return `<m:borderBox><m:borderBoxPr>${CTRL}</m:borderBoxPr>${box('e', n.e)}</m:borderBox>`;
    case 'phant': return `<m:phant><m:phantPr>${CTRL}</m:phantPr>${box('e', n.e)}</m:phant>`;
    default: return '';
  }
}

/**
 * Parse the linear form. `{ ok: true, ast }`, or `{ ok: false, error, at }`
 * with `at` the character of the input the problem starts at.
 */
export function parseLinear(src) {
  try {
    const text = String(src ?? '');
    if (!text.trim()) return { ok: false, error: 'Type an equation', at: 0 };
    const p = new Parser(tokenize(text));
    const ast = p.row(new Set());
    return { ok: true, ast };
  } catch (err) {
    return { ok: false, error: err.message, at: err.at ?? 0 };
  }
}

/**
 * The linear form as the OMML Word writes: `m:oMathPara` round a display
 * equation (one on a line of its own), a bare `m:oMath` for one in the
 * words. `{ ok: true, xml }` or `{ ok: false, error, at }`.
 */
export function linearToOmml(src, { display = true } = {}) {
  const parsed = parseLinear(src);
  if (!parsed.ok) return parsed;
  const body = rowXml(parsed.ast);
  const math = `<m:oMath>${body}</m:oMath>`;
  return { ok: true, xml: display ? `<m:oMathPara>${math}</m:oMathPara>` : math };
}

/**
 * Word's own built-in equations — Insert → Equation's gallery — in the
 * linear form, which builds each up exactly as the editor would.
 */
export const EQUATION_GALLERY = [
  { name: 'Area of Circle', linear: 'A=πr^2' },
  { name: 'Binomial Theorem', linear: '(x+a)^n=∑_(k=0)^n▒(n¦k) x^k a^(n-k)' },
  { name: 'Expansion of a Sum', linear: '(1+x)^n=1+nx/1!+(n(n-1)x^2)/2!+⋯' },
  { name: 'Fourier Series', linear: 'f(x)=a_0+∑_(n=1)^∞▒(a_n cos\u2061〖nπx/L〗+b_n sin\u2061〖nπx/L〗)' },
  { name: 'Pythagorean Theorem', linear: 'a^2+b^2=c^2' },
  { name: 'Quadratic Formula', linear: 'x=(-b±√(b^2-4ac))/2a' },
  { name: 'Taylor Expansion', linear: 'e^x=1+x/1!+x^2/2!+x^3/3!+⋯, -∞<x<∞' },
  { name: 'Trig Identity 1', linear: 'sin\u2061α±sin\u2061β=2 sin\u2061〖1/2(α±β)〗 cos\u2061〖1/2(α∓β)〗' },
  { name: 'Trig Identity 2', linear: 'cos\u2061α+cos\u2061β=2 cos\u2061〖1/2(α+β)〗 cos\u2061〖1/2(α-β)〗' },
];
