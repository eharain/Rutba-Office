'use strict';
// Kashida: justified Arabic stretched the way a calligrapher stretches it,
// by lengthening the joins between letters with tatweel, rather than by
// widening the spaces between words. Word does it for a paragraph aligned
// Justify Low, Medium or High (w:jc lowKashida, mediumKashida, highKashida).
//
// A kashida goes only where two letters join: after a letter that joins
// the one after it, before a letter that joins the one before it. Never
// between lam and alef, whose ligature it would break, and never next to a
// space. Within a word the place nearest its end is used first, as the eye
// expects; the stretch is shared among the line's words, one at a time.

const { joiningType, isMark } = require('./shaping');

const TATWEEL = '\u0640';

/**
 * The places in a text where a kashida may go, as positions to insert at,
 * grouped by word, the word's last place first.
 */
function kashidaPlaces(text) {
  const chars = Array.from(String(text ?? ''));
  const words = [];
  let current = [];
  // The letter before, past any marks on it: [index of the letter, its joining type].
  let prev = null;
  // Where the marks after a letter end, the place a kashida after it goes.
  let after = 0;
  for (let i = 0; i < chars.length; i++) {
    const cp = chars[i].codePointAt(0);
    if (isMark(cp)) { after = i + 1; continue; }
    const type = joiningType(cp);
    if (type === null || type === 'C') {
      if (current.length) words.push(current.reverse());
      current = [];
      prev = type === 'C' ? [i, 'C'] : null;
      after = i + 1;
      continue;
    }
    const lamAlef = prev && chars[prev[0]].codePointAt(0) === 0x0644 && [0x0622, 0x0623, 0x0625, 0x0627].includes(cp);
    if (prev && (prev[1] === 'D') && (type === 'D' || type === 'R') && !lamAlef) current.push(after);
    prev = [i, type];
    after = i + 1;
  }
  if (current.length) words.push(current.reverse());
  return words;
}

/**
 * The text with `count` tatweels set into it, shared out over its words a
 * round at a time, each word's last place first. Places in UTF-16 units.
 * Returns the text as it came when it has nowhere to stretch.
 */
function stretchArabic(text, count) {
  const source = String(text ?? '');
  const n = Math.max(0, Math.floor(Number(count) || 0));
  const words = kashidaPlaces(source);
  if (!n || !words.length) return source;
  // How many tatweels go at each place: round robin, word by word, a word's
  // places in turn, until the count is spent.
  const at = new Map();
  let left = n;
  for (let round = 0; left > 0; round++) {
    let placed = false;
    for (const places of words) {
      if (!left) break;
      const place = places[round % places.length];
      at.set(place, (at.get(place) || 0) + 1);
      left -= 1;
      placed = true;
    }
    if (!placed) break;
  }
  // Code-point places to UTF-16 offsets, then the insertions, last first.
  const chars = Array.from(source);
  let out = '';
  for (let i = 0; i <= chars.length; i++) {
    if (at.has(i)) out += TATWEEL.repeat(at.get(i));
    if (i < chars.length) out += chars[i];
  }
  return out;
}

/** How much of a justified line's slack Word's three kashida alignments give to kashida, the rest to the spaces. */
const KASHIDA_SHARE = { lowKashida: 1 / 3, mediumKashida: 2 / 3, highKashida: 1 };

module.exports = { kashidaPlaces, stretchArabic, KASHIDA_SHARE, TATWEEL };
