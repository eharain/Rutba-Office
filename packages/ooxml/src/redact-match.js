/**
 * What redaction looks for, and where in some words it is, the part the
 * window shares with the package pass (redact.js), so the marks drawn on
 * the page are the words Apply takes away.
 *
 * A mark is words (`{ text }`, found wherever they occur, matching case or
 * not, whole words or not), one of the patterns (`{ pattern: 'email' |
 * 'phone' | 'card' }`), or a pattern of one's own (`{ regex }`). A match is
 * replaced by black bars, one for each character, so a redacted line keeps
 * its length.
 */

/** The bar a redacted character becomes: U+2588, a full block. */
export const BAR = '█';

/** The patterns Redact offers, as regular expressions over plain words. */
export const PATTERNS = {
  // An address: a local part, an at sign, a domain with a dot.
  email: /[\p{L}\p{N}._%+'-]+@[\p{L}\p{N}-]+(?:\.[\p{L}\p{N}-]+)*\.\p{L}{2,}/gu,
  // A telephone number: an optional country code and area code, then groups of digits; seven to fifteen digits in all.
  phone: /(?<![\p{L}\p{N}])(?:\+\d{1,3}[\s.-]?)?(?:\(\d{1,5}\)[\s.-]?)?\d{2,5}(?:[\s.-]?\d{2,5}){1,4}(?![\p{L}\p{N}])/gu,
  // A payment card: thirteen to nineteen digits, spaced or dashed in groups, that pass the Luhn check.
  card: /(?<!\d)(?:\d[ -]?){12,18}\d(?!\d)/g,
};

const digits = (s) => s.replace(/\D/g, '');

/** The Luhn check every payment card number passes. */
export function luhn(number) {
  const d = digits(number);
  let sum = 0;
  for (let i = 0; i < d.length; i++) {
    let n = Number(d[d.length - 1 - i]);
    if (i % 2 === 1) {
      n *= 2;
      if (n > 9) n -= 9;
    }
    sum += n;
  }
  return d.length >= 13 && d.length <= 19 && sum % 10 === 0;
}

/** A pattern's match is kept only when it is really one: a phone number's digit count, a card's checksum. */
const KEEP = {
  phone: (m) => {
    const n = digits(m).length;
    if (n < 7 || n > 15) return false;
    // A date or a plain year range is not a telephone number.
    if (/^\d{4}-\d{2}-\d{2}$|^\d{2}[./-]\d{2}[./-]\d{2,4}$/.test(m)) return false;
    // A bare run of digits is a number unless it is long enough to dial.
    return /[\s.()+-]/.test(m) || n >= 10;
  },
  card: (m) => luhn(m),
  email: () => true,
};

const escapeRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Marks compiled to finders: each answers the `[from, to)` spans it finds
 * in some words. A pattern of one's own that is not a valid regular
 * expression is refused with a sentence.
 */
export function compileMarks(marks = []) {
  const out = [];
  for (const mark of marks) {
    if (!mark) continue;
    if (mark.pattern) {
      const re = PATTERNS[mark.pattern];
      if (!re) throw new Error('"' + mark.pattern + '" is not a pattern Redact knows.');
      out.push({ mark, re: new RegExp(re.source, re.flags), keep: KEEP[mark.pattern] });
    } else if (mark.regex != null) {
      let re;
      try {
        re = new RegExp(String(mark.regex), 'gu' + (mark.matchCase ? '' : 'i'));
      } catch {
        throw new Error('"' + mark.regex + '" is not a pattern that can be searched for.');
      }
      out.push({ mark, re, keep: (m) => m.length > 0 });
    } else if (mark.text != null && String(mark.text).length) {
      const body = escapeRe(mark.text).replace(/\s+/g, '\\s+');
      const re = new RegExp(mark.wholeWords ? '(?<![\\p{L}\\p{N}_])' + body + '(?![\\p{L}\\p{N}_])' : body, 'gu' + (mark.matchCase ? '' : 'i'));
      out.push({ mark, re, keep: () => true });
    }
  }
  return out;
}

/**
 * Where the compiled marks find something in `text`: spans `[from, to)`,
 * sorted, overlaps joined. `which` (optional) collects, for each span, the
 * index of the mark that found it.
 */
export function findSpans(text, finders, which = null) {
  const spans = [];
  const s = String(text ?? '');
  if (!s) return spans;
  finders.forEach((f, k) => {
    f.re.lastIndex = 0;
    let m;
    while ((m = f.re.exec(s))) {
      if (m[0].length === 0) { f.re.lastIndex += 1; continue; }
      if (f.keep(m[0])) spans.push([m.index, m.index + m[0].length, k]);
    }
  });
  spans.sort((a, b) => a[0] - b[0] || b[1] - a[1]);
  const merged = [];
  for (const sp of spans) {
    const last = merged[merged.length - 1];
    if (last && sp[0] <= last[1]) last[1] = Math.max(last[1], sp[1]);
    else merged.push([sp[0], sp[1]]);
    if (which) which.push(sp[2]);
  }
  return merged;
}

/** Some words with the spans blacked out, a bar for each character (a pair of surrogates is one character). */
export function barsOver(text, spans) {
  if (!spans.length) return text;
  let out = '';
  let at = 0;
  for (const [a, b] of spans) {
    out += text.slice(at, a) + BAR.repeat([...text.slice(a, b)].length);
    at = b;
  }
  return out + text.slice(at);
}
