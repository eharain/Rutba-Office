// A formula as Excel writes it in a file, and as a person reads it.
//
// A function newer than Excel 2007 is written with a `_xlfn.` prefix — and
// FILTER, SORT and SORTBY with `_xlfn._xlws.` — so an Excel that does not
// know it shows #NAME? rather than a wrong answer; written without one, the
// Excel that DOES know it shows #NAME? too. A name LET or LAMBDA binds is
// written `_xlpm.name`, and the spill operator `A1#` is written
// `_xlfn.ANCHORARRAY(A1)`. `toFileFormula` puts all of that in; `fromFileFormula`
// takes it out again, so the formula bar shows what was typed.

/** Functions written `_xlfn._xlws.NAME`. */
const XLWS = new Set(['FILTER', 'SORT', 'SORTBY']);

/** Functions written `_xlfn.NAME` — every function Excel added after 2007 that this suite can meet. */
const XLFN = new Set(`ACOT ACOTH AGGREGATE ARABIC BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV
BITAND BITLSHIFT BITOR BITRSHIFT BITXOR CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV
CHISQ.INV.RT CHISQ.TEST COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC
CSCH DAYS DECIMAL ECMA.CEILING ERF.PRECISE ERFC.PRECISE EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST
FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY
FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS
HYPGEOM.DIST IFNA IFS IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISO.CEILING ISOWEEKNUM
LOGNORM.DIST LOGNORM.INV MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST
NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC
PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC QUERYSTRING RANK.AVG RANK.EQ RRI
SEC SECH SHEET SHEETS SKEW.P STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST
TEXTJOIN UNICHAR UNICODE VAR.P VAR.S WEBSERVICE WEIBULL.DIST WORKDAY.INTL XOR Z.TEST
ANCHORARRAY ARRAYTOTEXT BYCOL BYROW CHOOSECOLS CHOOSEROWS DROP EXPAND HSTACK ISOMITTED LAMBDA LET MAKEARRAY
MAP RANDARRAY REDUCE SCAN SEQUENCE SINGLE TAKE TEXTAFTER TEXTBEFORE TEXTSPLIT TOCOL TOROW UNIQUE
VALUETOTEXT VSTACK WRAPCOLS WRAPROWS XLOOKUP XMATCH REGEXTEST REGEXEXTRACT REGEXREPLACE`.split(/\s+/));

const IDENT = /[A-Za-z_\\][A-Za-z0-9_.\\]*/y;

/**
 * The formula cut into the pieces these rewrites care about: text in
 * double quotes and sheet names in single quotes passed through whole,
 * names (with whether a `(` follows), and everything else a character at
 * a time.
 */
function pieces(f) {
  const out = [];
  let i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < f.length) {
        if (f[j] === ch && f[j + 1] === ch) { j += 2; continue; }
        if (f[j] === ch) { j += 1; break; }
        j += 1;
      }
      out.push({ kind: ch === '"' ? 'string' : 'sheet', text: f.slice(i, j) });
      i = j;
      continue;
    }
    const prev = out[out.length - 1];
    const afterWord = prev && (prev.kind === 'name' || (prev.kind === 'char' && /[A-Za-z0-9_.$]/.test(prev.text)));
    if (!afterWord && /[A-Za-z_\\]/.test(ch)) {
      IDENT.lastIndex = i;
      const m = IDENT.exec(f);
      let j = i + m[0].length;
      while (f[j] === ' ') j += 1;
      out.push({ kind: 'name', text: m[0], call: f[j] === '(' });
      i += m[0].length;
      continue;
    }
    out.push({ kind: 'char', text: ch });
    i += 1;
  }
  return out;
}

/** A name with any file prefix taken off, upper case — what the function or parameter is called. */
const bare = (name) => name.toUpperCase().replace(/^(?:_XLFN\.|_XLWS\.|_XLPM\.)+/, '');

/** The names each LET and LAMBDA in the formula binds, upper case. */
function boundNames(ps) {
  const names = new Set();
  const frames = [];
  let depth = 0;
  for (let k = 0; k < ps.length; k++) {
    const p = ps[k];
    if (p.kind === 'name' && p.call) {
      const fn = bare(p.text);
      if (fn === 'LET' || fn === 'LAMBDA') {
        frames.push({ fn, depth: depth + 1, args: [[]] });
      }
      continue;
    }
    if (p.kind !== 'char') {
      const top = frames[frames.length - 1];
      if (top && depth === top.depth) top.args[top.args.length - 1].push(p);
      continue;
    }
    if (p.text === '(') { depth += 1; continue; }
    const top = frames[frames.length - 1];
    if (p.text === ',' && top && depth === top.depth) { top.args.push([]); continue; }
    if (p.text === ')') {
      if (top && depth === top.depth) {
        frames.pop();
        const args = top.args.map((a) => a.filter((x) => !(x.kind === 'char' && x.text === ' ')));
        const last = args.length - 1;
        args.forEach((a, n) => {
          const isName = a.length === 1 && a[0].kind === 'name' && !a[0].call;
          const binds = top.fn === 'LAMBDA' ? n < last : n % 2 === 0 && n < last;
          if (isName && binds) names.add(bare(a[0].text));
        });
      }
      depth -= 1;
      continue;
    }
    if (top && depth === top.depth) top.args[top.args.length - 1].push(p);
  }
  return names;
}

/** `A1#` → `_xlfn.ANCHORARRAY(A1)`, a sheet name before it kept; text in double quotes left alone. */
function anchorArrays(f) {
  const ref = /(?<![A-Za-z0-9_.$])((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+)#/g;
  return f.split(/("(?:[^"]|"")*")/).map((part, n) => (n % 2 ? part : part.replace(ref, (m, sheet, cell) => '_xlfn.ANCHORARRAY(' + (sheet ?? '') + cell + ')'))).join('');
}

/** A formula (without its `=`) as Excel writes it in a file. */
export function toFileFormula(formula) {
  const f = anchorArrays(String(formula ?? ''));
  const ps = pieces(f);
  const params = boundNames(ps);
  let out = '';
  for (const p of ps) {
    if (p.kind !== 'name') { out += p.text; continue; }
    if (/^_xl(?:fn|ws|pm)\./i.test(p.text)) { out += p.text; continue; }
    const upper = p.text.toUpperCase();
    if (params.has(upper)) out += '_xlpm.' + p.text;
    else if (p.call && XLWS.has(upper)) out += '_xlfn._xlws.' + upper;
    else if (p.call && XLFN.has(upper)) out += '_xlfn.' + upper;
    else out += p.text;
  }
  return out;
}

/** A formula (without its `=`) as a person reads it: the file's prefixes off, `A1#` back. */
export function fromFileFormula(formula) {
  const f = String(formula ?? '');
  if (!/_xl(?:fn|ws|pm)\./i.test(f)) return f;
  let out = '';
  for (const p of pieces(f)) out += p.kind === 'name' ? p.text.replace(/^(?:_xlfn\.|_xlws\.|_xlpm\.)+/i, '') : p.text;
  // ANCHORARRAY(ref) is the spill operator, written out.
  return out.replace(/\bANCHORARRAY\(\s*((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)?!?\$?[A-Za-z]{1,3}\$?\d+)\s*\)/gi, (m, ref) => ref + '#');
}
