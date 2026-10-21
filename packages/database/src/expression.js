// A calculated column, as Access writes one: [Quantity] * [Unit Price],
// [First Name] & " " & [Last Name], IIf([Paid], "Paid", "Due"),
// Format([Order Date], "yyyy-mm"). Read here into SQL from a fixed set of
// pieces: field names (quoted, and only those of the query's tables),
// numbers, text and dates (written as SQL literals, their quotes doubled),
// the operators and the functions listed below. Anything else is refused
// with what was not understood, so an expression can only ever compute a
// value: it cannot reach another table or change anything.

const text = (s) => `'${String(s).replace(/'/g, "''")}'`;

/** The words of an expression: [names], [Table].[Field], "text", #dates#, numbers, operators, words. */
export function tokens(src) {
  const out = [];
  const s = String(src);
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (/\s/.test(c)) { i++; continue; }
    if (c === '[') {
      const end = s.indexOf(']', i);
      if (end < 0) throw new Error(`A [ in "${s}" is not closed with ].`);
      const name = s.slice(i + 1, end);
      i = end + 1;
      // [Table].[Field]
      if (s[i] === '.' && s[i + 1] === '[') {
        const end2 = s.indexOf(']', i + 2);
        if (end2 < 0) throw new Error(`A [ in "${s}" is not closed with ].`);
        out.push({ t: 'name', table: name, field: s.slice(i + 2, end2) });
        i = end2 + 1;
      } else {
        out.push({ t: 'name', table: null, field: name });
      }
      continue;
    }
    if (c === '"' || c === "'") {
      let j = i + 1;
      let v = '';
      for (;;) {
        if (j >= s.length) throw new Error(`A quotation mark in "${s}" is not closed.`);
        if (s[j] === c && s[j + 1] === c) { v += c; j += 2; continue; }
        if (s[j] === c) break;
        v += s[j++];
      }
      out.push({ t: 'text', v });
      i = j + 1;
      continue;
    }
    if (c === '#') {
      const end = s.indexOf('#', i + 1);
      const d = end > 0 ? s.slice(i + 1, end) : '';
      if (!/^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?$/.test(d)) throw new Error(`Write a date as #2026-10-21#: "${s.slice(i, end > 0 ? end + 1 : s.length)}" is not one.`);
      out.push({ t: 'text', v: d.replace(' ', 'T') });
      i = end + 1;
      continue;
    }
    const rest = s.slice(i);
    const num = /^\d+(\.\d+)?/.exec(rest);
    if (num) { out.push({ t: 'num', v: num[0] }); i += num[0].length; continue; }
    const op = /^(<>|<=|>=|[-+*/&(),=<>])/.exec(rest);
    if (op) { out.push({ t: 'op', v: op[0] }); i += op[0].length; continue; }
    const word = /^[A-Za-z_][A-Za-z0-9_]*/.exec(rest);
    if (word) { out.push({ t: 'word', v: word[0] }); i += word[0].length; continue; }
    throw new Error(`"${c}" in "${s}" is not part of an expression here.`);
  }
  return out;
}

/** The functions an expression may use: [least, most] values, and the SQL each stands for. */
const FUNCTIONS = {
  round: [1, 2, (a) => `ROUND(${a[0]}, ${a[1] ?? '0'})`],
  nz: [1, 2, (a) => `COALESCE(${a[0]}, ${a[1] ?? '0'})`],
  iif: [3, 3, (a) => `(CASE WHEN ${a[0]} THEN ${a[1]} ELSE ${a[2]} END)`],
  abs: [1, 1, (a) => `ABS(${a[0]})`],
  year: [1, 1, (a) => `CAST(strftime('%Y', ${a[0]}) AS INTEGER)`],
  month: [1, 1, (a) => `CAST(strftime('%m', ${a[0]}) AS INTEGER)`],
  day: [1, 1, (a) => `CAST(strftime('%d', ${a[0]}) AS INTEGER)`],
  date: [0, 0, (a, o) => (o.today ? text(String(o.today).slice(0, 10)) : `date('now', 'localtime')`)],
  now: [0, 0, (a, o) => (o.today ? text(String(o.today).length > 10 ? String(o.today).replace(' ', 'T').slice(0, 19) : `${o.today}T12:00:00`) : `datetime('now', 'localtime')`)],
  len: [1, 1, (a) => `LENGTH(${a[0]})`],
  left: [2, 2, (a) => `SUBSTR(${a[0]}, 1, ${a[1]})`],
  right: [2, 2, (a) => `SUBSTR(${a[0]}, -(${a[1]}))`],
  mid: [2, 3, (a) => (a.length === 2 ? `SUBSTR(${a[0]}, ${a[1]})` : `SUBSTR(${a[0]}, ${a[1]}, ${a[2]})`)],
  ucase: [1, 1, (a) => `UPPER(${a[0]})`],
  lcase: [1, 1, (a) => `LOWER(${a[0]})`],
  trim: [1, 1, (a) => `TRIM(${a[0]})`],
  format: [2, 2, (a) => formatSql(a[0], literal(a[1]) ?? '')],
  weekday: [1, 1, (a) => `(CAST(strftime('%w', ${a[0]}) AS INTEGER) + 1)`],
  datediff: [3, 3, (a) => {
    const unit = literal(a[0])?.toLowerCase();
    const y = (d) => `CAST(strftime('%Y', ${d}) AS INTEGER)`;
    const m = (d) => `CAST(strftime('%m', ${d}) AS INTEGER)`;
    if (unit === 'd') return `CAST(julianday(${a[2]}) - julianday(${a[1]}) AS INTEGER)`;
    if (unit === 'm') return `((${y(a[2])} - ${y(a[1])}) * 12 + ${m(a[2])} - ${m(a[1])})`;
    if (unit === 'yyyy') return `(${y(a[2])} - ${y(a[1])})`;
    throw new Error('DateDiff counts "d" (days), "m" (months) or "yyyy" (years).');
  }],
};
export const FUNCTION_NAMES = ['Round', 'Nz', 'IIf', 'Abs', 'Year', 'Month', 'Day', 'Weekday', 'Date', 'Now', 'Format', 'DateDiff', 'Len', 'Left', 'Right', 'Mid', 'UCase', 'LCase', 'Trim'];

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const caseOf = (key, names, short) => `CASE ${key} ${names.map((n, i) => `WHEN ${i + (names === MONTHS ? 1 : 0)} THEN '${short ? n.slice(0, 3) : n}'`).join(' ')} END`;

/**
 * Format(value, pattern) as SQL, as Access formats:
 *   "00000"      a number padded with noughts (an invoice number: 00042)
 *   "0.00"       a number with its decimal places
 *   "#,##0.00"   with thousands separated (1,234.50); "#,##0" without decimals
 *   dates by their parts: yyyy yy mmmm mmm mm m dddd ddd dd d hh h nn ss,
 *   the rest kept as written ("d mmmm yyyy" is 21 October 2026,
 *   "dd/mm/yyyy" is 21/10/2026, "ddd" is Wed). Month and day names are English.
 */
export function formatSql(v, pattern) {
  if (/^0+$/.test(pattern)) return `printf('%0${pattern.length}d', ${v})`;
  let m = /^0\.(0+)$/.exec(pattern);
  if (m) return `printf('%.${m[1].length}f', ${v})`;
  m = /^#,##0(?:\.(0+))?$/.exec(pattern);
  if (m) {
    const places = m[1] ? m[1].length : 0;
    const r = `ROUND(${v}, ${places})`;
    const whole = `printf('%,d', CAST(ABS(${r}) AS INTEGER))`;
    const frac = places ? ` || SUBSTR(printf('%.${places}f', ABS(${r})), -${places + 1})` : '';
    return `(CASE WHEN ${v} IS NULL THEN NULL ELSE (CASE WHEN ${r} < 0 THEN '-' ELSE '' END) || ${whole}${frac} END)`;
  }
  const parts = [];
  const re = /yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|nn|ss|[^ymdhns]+|[ymdhns]/gi;
  let any = false;
  for (const tok of String(pattern).match(re) || []) {
    const k = tok.toLowerCase();
    const piece = {
      yyyy: `strftime('%Y', ${v})`, yy: `SUBSTR(strftime('%Y', ${v}), 3)`,
      mmmm: caseOf(`CAST(strftime('%m', ${v}) AS INTEGER)`, MONTHS, false), mmm: caseOf(`CAST(strftime('%m', ${v}) AS INTEGER)`, MONTHS, true),
      mm: `strftime('%m', ${v})`, m: `CAST(CAST(strftime('%m', ${v}) AS INTEGER) AS TEXT)`,
      dddd: caseOf(`CAST(strftime('%w', ${v}) AS INTEGER)`, DAYS, false), ddd: caseOf(`CAST(strftime('%w', ${v}) AS INTEGER)`, DAYS, true),
      dd: `strftime('%d', ${v})`, d: `CAST(CAST(strftime('%d', ${v}) AS INTEGER) AS TEXT)`,
      hh: `strftime('%H', ${v})`, h: `CAST(CAST(strftime('%H', ${v}) AS INTEGER) AS TEXT)`, nn: `strftime('%M', ${v})`, ss: `strftime('%S', ${v})`,
    }[k];
    if (piece) { any = true; parts.push(piece); } else parts.push(text(tok));
  }
  if (!any) throw new Error('Format takes noughts for a number ("00000", "0.00", "#,##0.00") or a date\'s parts ("d mmmm yyyy", "dd/mm/yyyy", "ddd").');
  return `(CASE WHEN ${v} IS NULL THEN NULL ELSE ${parts.join(' || ')} END)`;
}

/** The text of an SQL string literal this module wrote, or null. */
function literal(sql) {
  const m = /^'((?:[^']|'')*)'$/s.exec(sql);
  return m ? m[1].replace(/''/g, "'") : null;
}

/**
 * An expression as SQL. `resolve(table, field)` answers the quoted column
 * for [Field] or [Table].[Field], or throws when the query has none (or, for
 * [Field] alone, more than one table has it). `options.today` (YYYY-MM-DD)
 * is the day Date() and Now() answer, for a check of how a database looks
 * on another day.
 */
export function expressionSql(src, resolve, options = {}) {
  const ts = tokens(src);
  if (!ts.length) throw new Error('The expression is empty.');
  let at = 0;
  const peek = () => ts[at];
  const isOp = (v) => peek()?.t === 'op' && peek().v === v;
  const isWord = (v) => peek()?.t === 'word' && peek().v.toLowerCase() === v;
  const expect = (v) => { if (!isOp(v)) throw new Error(`"${src}": a ${v} is missing.`); at++; };

  const or = () => { let l = and(); while (isWord('or')) { at++; l = `(${l} OR ${and()})`; } return l; };
  const and = () => { let l = not(); while (isWord('and')) { at++; l = `(${l} AND ${not()})`; } return l; };
  const not = () => (isWord('not') ? (at++, `(NOT ${not()})`) : compare());
  const compare = () => {
    let l = concat();
    while (peek()?.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const o = ts[at++].v; l = `(${l} ${o} ${concat()})`; }
    return l;
  };
  // & joins text, an empty value joining as nothing, as Access's & does.
  const concat = () => { let l = sum(); while (isOp('&')) { at++; l = `(COALESCE(${l}, '') || COALESCE(${sum()}, ''))`; } return l; };
  const sum = () => { let l = product(); while (isOp('+') || isOp('-')) { const o = ts[at++].v; l = `(${l} ${o} ${product()})`; } return l; };
  const product = () => {
    let l = unary();
    for (;;) {
      if (isOp('*')) { at++; l = `(${l} * ${unary()})`; continue; }
      // Division in decimals, as Access divides: 7 / 2 is 3.5, not 3.
      if (isOp('/')) { at++; l = `(CAST(${l} AS REAL) / ${unary()})`; continue; }
      if (isWord('mod')) { at++; l = `(${l} % ${unary()})`; continue; }
      return l;
    }
  };
  const unary = () => {
    if (isOp('-')) { at++; return `(-${unary()})`; }
    if (isOp('+')) { at++; return unary(); }
    return primary();
  };
  const primary = () => {
    const t = peek();
    if (!t) throw new Error(`"${src}" ends too soon.`);
    at++;
    if (t.t === 'num') return t.v;
    if (t.t === 'text') return text(t.v);
    if (t.t === 'name') return resolve(t.table, t.field);
    if (t.t === 'op' && t.v === '(') { const e = or(); expect(')'); return e; }
    if (t.t === 'word') {
      const w = t.v.toLowerCase();
      if (w === 'true' || w === 'yes') return '1';
      if (w === 'false' || w === 'no') return '0';
      if (w === 'null') return 'NULL';
      const fn = FUNCTIONS[w];
      if (!fn) throw new Error(`${t.v} is not a function Data Files knows. It knows ${FUNCTION_NAMES.join(', ')}.`);
      expect('(');
      const args = [];
      if (!isOp(')')) {
        args.push(or());
        while (isOp(',')) { at++; args.push(or()); }
      }
      expect(')');
      const [lo, hi, sql] = fn;
      if (args.length < lo || args.length > hi) throw new Error(`${t.v} takes ${lo === hi ? lo : `${lo} or ${hi}`} value${hi === 1 ? '' : 's'}.`);
      return sql(args, options);
    }
    throw new Error(`"${src}": "${t.v}" is not where a value goes.`);
  };

  const sql = or();
  if (at < ts.length) {
    const t = ts[at];
    throw new Error(`"${src}": ${t.t === 'name' ? `[${t.field}]` : `"${t.v}"`} was not expected there.`);
  }
  return sql;
}

/** The fields an expression names, as [{ table, field }]. */
export function fieldsNamed(src) {
  return tokens(src).filter((t) => t.t === 'name').map((t) => ({ table: t.table, field: t.field }));
}
