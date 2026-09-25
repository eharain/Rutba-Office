/**
 * Flash Fill: from an example or two typed beside a block of data, the
 * transformation that makes them, and the rest of the column filled by it.
 *
 * Pure and deterministic. A program is a sequence of pieces, each either
 * a literal or an ATOM — a part of one source column's value: the whole
 * value, a token between delimiters (counted from the start or from the
 * end), a run of digits or letters, the first or last few characters, the
 * initials of its words, a part of a date — each optionally in upper,
 * lower or proper case, or cut to its first letter. The search walks every
 * example's output at once, trying the atoms that fit at the current place
 * in all of them (the longest first, then the plainest), then a literal
 * character they all share. The first program that spells out every
 * example exactly wins; none is an answer too, and nothing is filled.
 */

/** The delimiters a token is cut by, the structural ones before the space. */
const DELIMITERS = [', ', ',', ';', '|', '\t', ' - ', '/', '-', '@', '.', '_', ':', ' '];

const CASES = ['none', 'upper', 'lower', 'proper'];

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function proper(s) {
  return s.toLowerCase().replace(/(^|[^\p{L}\p{N}'])(\p{L})/gu, (m, a, b) => a + b.toUpperCase());
}

function cased(s, how) {
  if (how === 'upper') return s.toUpperCase();
  if (how === 'lower') return s.toLowerCase();
  if (how === 'proper') return proper(s);
  return s;
}

function tokens(text, delimiter) {
  const parts = delimiter === ' ' ? text.trim().split(/\s+/) : text.split(delimiter).map((t) => t.trim());
  return parts.filter((t) => t !== '');
}

/** An Excel date serial as its parts (1900 system, the phantom 29 February kept). */
function dateParts(serial) {
  if (typeof serial !== 'number' || !Number.isFinite(serial) || serial < 1) return null;
  const whole = Math.floor(serial);
  const days = whole > 59 ? whole - 1 : whole;
  const d = new Date(Date.UTC(1899, 11, 31) + days * 86400000);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), w: d.getUTCDay() };
}

/**
 * The base value of an atom for one row, before case and first-letter:
 * null when the row has no such part (a third token in two-token text).
 */
function base(atom, sources) {
  const src = sources[atom.col];
  if (!src) return null;
  const text = String(src.text ?? '');
  switch (atom.kind) {
    case 'whole':
      return text === '' ? null : text;
    case 'token': {
      const parts = tokens(text, atom.delimiter);
      if (parts.length < 2) return null;
      const i = atom.index < 0 ? parts.length + atom.index : atom.index;
      return parts[i] ?? null;
    }
    case 'digits':
    case 'letters': {
      const runs = text.match(atom.kind === 'digits' ? /\d+/g : /\p{L}+/gu) || [];
      const i = atom.index < 0 ? runs.length + atom.index : atom.index;
      return runs[i] ?? null;
    }
    case 'prefix':
      return text.length >= atom.n ? text.slice(0, atom.n) : null;
    case 'suffix':
      return text.length >= atom.n ? text.slice(text.length - atom.n) : null;
    case 'initials': {
      const words = text.split(/[\s,\-]+/).filter(Boolean);
      if (words.length < 2) return null;
      return words.map((w) => w[0]).join(atom.dot ? '.' : '') + (atom.dot ? '.' : '');
    }
    case 'date': {
      const p = dateParts(src.serial);
      if (!p) return null;
      switch (atom.part) {
        case 'yyyy': return String(p.y);
        case 'yy': return String(p.y % 100).padStart(2, '0');
        case 'm': return String(p.m);
        case 'mm': return String(p.m).padStart(2, '0');
        case 'mmm': return MONTHS[p.m - 1].slice(0, 3);
        case 'mmmm': return MONTHS[p.m - 1];
        case 'd': return String(p.d);
        case 'dd': return String(p.d).padStart(2, '0');
        case 'ddd': return DAYS[p.w].slice(0, 3);
        case 'dddd': return DAYS[p.w];
        default: return null;
      }
    }
    default:
      return null;
  }
}

/** An atom's value for one row, or null. */
export function evaluateAtom(atom, sources) {
  const b = base(atom, sources);
  if (b === null || b === '') return null;
  const one = atom.first ? b[0] : b;
  return cased(one, atom.casing);
}

/** How plain an atom is: the search tries plainer atoms first among equals. */
function rank(atom) {
  const kind = { whole: 0, token: 1, date: 1.5, digits: 2, letters: 2, initials: 3, prefix: 4, suffix: 4 }[atom.kind] ?? 9;
  const delimiter = atom.kind === 'token' ? DELIMITERS.indexOf(atom.delimiter) / 100 : 0;
  const index = atom.index !== undefined ? (atom.index < 0 ? 0.004 : 0) + Math.abs(atom.index) / 1000 : 0;
  return kind + (atom.first ? 0.5 : 0) + CASES.indexOf(atom.casing) * 0.1 + delimiter + index + atom.col / 100000;
}

const atomKey = (a) => JSON.stringify([a.kind, a.col, a.delimiter ?? null, a.index ?? null, a.n ?? null, a.part ?? null, a.dot ?? null, a.first ? 1 : 0, a.casing]);

/** Every atom one row's sources offer, each with its value in that row. */
function atomsOf(sources) {
  const out = [];
  // Only a word's own first letter is an initial: the first letter of the
  // last two characters of a word is a coincidence, and it made "Dr Ada"
  // read as a letter of "Ada" upper-cased.
  const add = (atom) => {
    const initials = ['whole', 'token', 'letters'].includes(atom.kind) ? [false, true] : [false];
    for (const casing of CASES) {
      for (const first of initials) {
        const a = { ...atom, casing, first };
        const value = evaluateAtom(a, sources);
        if (value) out.push({ atom: a, value });
      }
    }
  };
  sources.forEach((src, col) => {
    if (!src) return;
    const text = String(src.text ?? '');
    if (text === '') return;
    add({ kind: 'whole', col });
    for (const delimiter of DELIMITERS) {
      const n = tokens(text, delimiter).length;
      if (n < 2) continue;
      for (let i = 0; i < n; i++) {
        add({ kind: 'token', col, delimiter, index: i });
        add({ kind: 'token', col, delimiter, index: i - n });
      }
    }
    for (const kind of ['digits', 'letters']) {
      const n = (text.match(kind === 'digits' ? /\d+/g : /\p{L}+/gu) || []).length;
      for (let i = 0; i < n; i++) {
        add({ kind, col, index: i });
        add({ kind, col, index: i - n });
      }
    }
    for (const dot of [false, true]) add({ kind: 'initials', col, dot });
    for (let n = 2; n < text.length; n++) {
      add({ kind: 'prefix', col, n });
      add({ kind: 'suffix', col, n });
    }
    if (typeof src.serial === 'number') {
      for (const part of ['yyyy', 'yy', 'm', 'mm', 'mmm', 'mmmm', 'd', 'dd', 'ddd', 'dddd']) add({ kind: 'date', col, part });
    }
  });
  return out;
}

/**
 * The program that turns every example's sources into its output, or null.
 *
 * @param {Array<{ sources: Array<{ text: string, serial?: number } | null>, output: string }>} examples
 * @returns {{ pieces: Array<{ literal: string } | { atom: object }> } | null}
 */
export function inferProgram(examples, { maxNodes = 20000 } = {}) {
  const list = examples.filter((e) => e && String(e.output ?? '') !== '');
  if (!list.length) return null;
  const outputs = list.map((e) => String(e.output));
  const lead = atomsOf(list[0].sources);
  let nodes = 0;
  const failed = new Set();

  const search = (at, pieces) => {
    if (at.every((p, j) => p === outputs[j].length)) return pieces;
    if (at.some((p, j) => p >= outputs[j].length)) return null;
    const key = at.join(',');
    if (failed.has(key)) return null;
    if (++nodes > maxNodes) return null;
    const fits = [];
    const seen = new Set();
    for (const { atom, value } of lead) {
      if (!outputs[0].startsWith(value, at[0])) continue;
      const k = atomKey(atom);
      if (seen.has(k)) continue;
      seen.add(k);
      fits.push({ atom, value });
    }
    fits.sort((a, b) => b.value.length - a.value.length || rank(a.atom) - rank(b.atom));
    // A one-letter part stands at a word's edge — "AL", "A. Lovelace" — or
    // it is a letter of a word typed, and the "a" of "Navy" is not an initial.
    const letter = (ch) => ch !== undefined && /[\p{L}\p{N}]/u.test(ch);
    const after = pieces[pieces.length - 1];
    const edgeBefore = (j) => at[j] === 0 || !letter(outputs[j][at[j] - 1]) || (after && after.atom);
    for (const { atom, value } of fits) {
      if (value.length === 1 && !outputs.every((_, j) => edgeBefore(j))) continue;
      const next = [];
      let ok = true;
      for (let j = 0; j < list.length; j++) {
        const v = j === 0 ? evaluateAtom(atom, list[0].sources) : evaluateAtom(atom, list[j].sources);
        if (!v || !outputs[j].startsWith(v, at[j])) { ok = false; break; }
        next.push(at[j] + v.length);
      }
      if (!ok) continue;
      const found = search(next, [...pieces, { atom }]);
      if (found) return found;
    }
    // A character every output has here, as a literal — but not a letter
    // straight after a one-letter part, which would make that part a letter
    // of a typed word after all.
    const ch = outputs[0][at[0]];
    const lastAtom = pieces[pieces.length - 1]?.atom;
    const stuck = lastAtom && letter(ch) && outputs.every((_, j) => letter(outputs[j][at[j] - 1])) && evaluateAtom(lastAtom, list[0].sources)?.length === 1;
    if (!stuck && outputs.every((o, j) => o[at[j]] === ch)) {
      const last = pieces[pieces.length - 1];
      const merged = last && last.literal !== undefined
        ? [...pieces.slice(0, -1), { literal: last.literal + ch }]
        : [...pieces, { literal: ch }];
      const found = search(at.map((p) => p + 1), merged);
      if (found) return found;
    }
    failed.add(key);
    return null;
  };

  const pieces = search(outputs.map(() => 0), []);
  // All literals is a constant, not a pattern: Flash Fill does not copy words down.
  if (!pieces || !pieces.some((p) => p.atom)) return null;
  return { pieces };
}

/** Run a program on one row's sources: the text, or null where a part is missing. */
export function runProgram(program, sources) {
  let out = '';
  for (const p of program.pieces) {
    if (p.literal !== undefined) { out += p.literal; continue; }
    const v = evaluateAtom(p.atom, sources);
    if (v === null) return null;
    out += v;
  }
  return out;
}
