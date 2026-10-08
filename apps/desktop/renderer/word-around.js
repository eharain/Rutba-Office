// The word a caret is in, or a selection covers, in a paragraph or a cell's
// words: what Review → Thesaurus looks up and puts another in place of.

/** The word in `text` at or around `at` (or the one the range start..end covers): { word, start, end }, or null. */
export function wordAround(text, start, end = start) {
  const s = String(text || '');
  if (end > start) {
    const picked = s.slice(start, end).trim();
    if (picked && !/\s/.test(picked)) return { word: picked.replace(/^[^\p{L}]+|[^\p{L}'-]+$/gu, ''), start: start + s.slice(start, end).indexOf(picked), end: start + s.slice(start, end).indexOf(picked) + picked.length };
  }
  const isLetter = (ch) => /[\p{L}'-]/u.test(ch || '');
  let a = start;
  let b = start;
  while (a > 0 && isLetter(s[a - 1])) a -= 1;
  while (b < s.length && isLetter(s[b])) b += 1;
  if (a === b) return null;
  let word = s.slice(a, b);
  while (word && !/\p{L}/u.test(word[0])) { word = word.slice(1); a += 1; }
  while (word && !/\p{L}/u.test(word[word.length - 1])) { word = word.slice(0, -1); b -= 1; }
  return word ? { word, start: a, end: b } : null;
}
