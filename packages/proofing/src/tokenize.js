// Words to check, in a run of text.
//
// Office's spelling options skip, by default: words in UPPERCASE, words
// that contain numbers, and Internet and file addresses. Those are the
// options here too, with the same defaults. What is left is every word —
// letters, with an apostrophe or a hyphen inside — at its offset in the
// text, so a Change can replace exactly those characters.

export const DEFAULT_OPTIONS = { ignoreUppercase: true, ignoreNumbers: true, ignoreAddresses: true };

// An address in the text: a web address, an e-mail address, a path on a
// disk or a share. Matched first, so the words inside are never offered.
const ADDRESSES = new RegExp([
  String.raw`\b(?:https?|ftp|file|mailto):[^\s<>"]+`,
  String.raw`\bwww\.[^\s<>"]+`,
  String.raw`[\p{L}\p{N}._%+-]+@[\p{L}\p{N}.-]+\.[\p{L}]{2,}`,
  String.raw`(?:\b[A-Za-z]:|\\\\[^\s\\]+)\\[^\s<>"|]*`,
  String.raw`(?:^|(?<=\s))(?:\.{0,2}\/)?(?:[\w.-]+\/)+[\w.-]*`,
  String.raw`\b[\p{L}\p{N}-]+(?:\.[\p{L}\p{N}-]+)*\.(?:com|org|net|io|co|uk|gov|edu|info|biz|dev|app|pdf|docx|xlsx|pptx|txt|csv|html?|js|json|png|jpe?g|gif|zip)\b`,
].join('|'), 'giu');

// A word: letters and digits, with an apostrophe or a hyphen between two of them.
const WORD = /[\p{L}\p{M}\p{N}_]+(?:['’][\p{L}\p{M}]+)*/gu;

/**
 * The words of `text` worth checking: `{ word, offset, length }`, offsets in
 * UTF-16 units of `text`. `word` has a curly apostrophe made straight, which
 * is how the dictionaries spell "don't".
 */
export function tokenize(text, options = {}) {
  const o = { ...DEFAULT_OPTIONS, ...options };
  const s = String(text ?? '');
  const skip = [];
  if (o.ignoreAddresses) for (const m of s.matchAll(ADDRESSES)) skip.push([m.index, m.index + m[0].length]);
  const out = [];
  for (const m of s.matchAll(WORD)) {
    const start = m.index;
    const end = start + m[0].length;
    if (skip.some(([a, b]) => start < b && end > a)) continue;
    // A hyphenated word is checked a part at a time — each part is a word
    // of its own in the dictionaries.
    let raw = m[0];
    // "O'Brien's" is a word; a quotation's closing apostrophe is not part of it.
    if (/['’]$/.test(raw)) raw = raw.slice(0, -1);
    if (!/\p{L}/u.test(raw)) continue;
    if (o.ignoreNumbers && /[\p{N}_]/u.test(raw)) continue;
    const letters = raw.replace(/['’]/g, '');
    if (o.ignoreUppercase && letters.length > 1 && letters === letters.toUpperCase() && letters !== letters.toLowerCase()) continue;
    if (letters.length < 2 && !/^[aAI]$/.test(letters)) continue;
    out.push({ word: raw.replace(/’/g, "'"), offset: start, length: raw.length });
  }
  return out;
}

/** The same capitalisation on a suggestion as on the word it replaces. */
export function matchCase(original, replacement) {
  const o = String(original);
  const r = String(replacement);
  if (!o || !r) return r;
  if (o.length > 1 && o === o.toUpperCase()) return r.toUpperCase();
  if (o[0] === o[0].toUpperCase() && o[0] !== o[0].toLowerCase() && r[0] === r[0].toLowerCase()) return r[0].toUpperCase() + r.slice(1);
  return r;
}
