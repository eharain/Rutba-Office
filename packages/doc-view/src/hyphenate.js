/**
 * Hyphenation for English: where a word may be broken at the end of a line.
 *
 * Frank Liang's method, the one TeX uses: a table of letter patterns, each
 * with numbers between its letters; every pattern that matches a stretch of
 * the word writes its numbers there, the greatest number at each gap wins,
 * and an odd number is a place the word may break. The same table serves the
 * page on screen (soft hyphens between the letters) and the printout (the
 * paginator's own line breaker), so a word breaks in the same place on both.
 *
 * The table is this suite's own, written for it: English suffixes and
 * prefixes, doubled consonants, and a vowel–consonant–consonant–vowel rule
 * that keeps the pairs English writes as one sound (ch, sh, th, ph, wh, gh,
 * ck, ng) and the blends a syllable starts with (br, pl, tr…) together. It
 * errs towards breaking less: a word it is not sure of is left whole. The
 * engine reads TeX's pattern notation, so a fuller table drops in as data.
 */

const VOWELS = 'aeiouy';
const CONSONANTS = 'bcdfghjklmnpqrstvwxz';
/** Two consonants English writes for one sound: never parted. */
const DIGRAPHS = ['ch', 'sh', 'th', 'ph', 'wh', 'gh', 'ck', 'ng', 'qu'];
/** The pairs a syllable starts with: the break goes before them. */
const BLENDS = ['bl', 'br', 'cl', 'cr', 'dr', 'fl', 'fr', 'gl', 'gr', 'pl', 'pr', 'tr', 'sc', 'sk', 'sp', 'wr'];

/** Patterns written out: endings, beginnings and the pairs that must stay whole. */
const WRITTEN = [
  // Endings that are syllables of their own.
  '1tion', '1sion', '1cian', '1tial', '1cial', '1tious', '1cious', '1ment', '1ness', '1less', '1ful.', '1fully',
  '1able.', '1ible.', '1ably.', '1ibly.', '1ture', '1ture.', '1sure.', '1ward', '1wards',
  // -ing after a consonant — and not where the consonant ends a word's first cluster.
  'k3ing.', 'd3ing.', 't3ing.', 's3ing.', 'n3ing.', 'w3ing.', 'y3ing.', 'h3ing.', 'l3ing.', 'm3ing.', 'r3ing.', 'p3ing.', 'v3ing.', 'z3ing.',
  'th4ing', '.spr4', '.str4', '.scr4', '.thr4', '.spl4', '.shr4',
  // Beginnings that are syllables of their own.
  '.con1', '.com1', '.dis1', '.pro1', '.inter1', '.over1', '.under1', '.trans1', '.sub1', '.super1', '.counter1', '.ex1', '.mis1', '.non1', '.out1',
  '.pre1', '.pre2s', '.un1', '.un2i', '.un2d', '.re2',
  // A doubled consonant before -ed ends the word whole: shipped, not ship-ped.
  ...CONSONANTS.split('').map((c) => `${c}2${c}ed.`),
  // A doubled consonant before -ing parts the pair, not the ending: run-ning.
  ...CONSONANTS.split('').map((c) => `${c}1${c}4ing.`),
  // -le after a consonant: ta-ble, sim-ple.
  ...CONSONANTS.split('').filter((c) => c !== 'l').map((c) => `1${c}le.`),
];

/** Patterns made from the letter classes: doubled consonants, and vowel–consonant–consonant–vowel. */
function fromClasses() {
  const out = [];
  for (const c of CONSONANTS) if (c !== 'h' && c !== 'w' && c !== 'x' && c !== 'j' && c !== 'q') out.push(`${c}1${c}`);
  out.push('ck1');
  for (const a of VOWELS) {
    for (const b of VOWELS) {
      for (const c1 of CONSONANTS) {
        for (const c2 of CONSONANTS) {
          const pair = c1 + c2;
          if (c1 === c2 || DIGRAPHS.includes(pair)) continue;
          if (c2 === 'h') continue; // an h after a consonant joins it (ch, sh, th, ph, wh, gh — and kh, rh)
          if (BLENDS.includes(pair)) out.push(`${a}1${c1}${c2}${b}`);
          else out.push(`${a}${c1}1${c2}${b}`);
        }
      }
    }
  }
  return out;
}

/** A pattern in TeX's notation as its letters and the numbers between them. */
function parse(pattern) {
  const letters = [];
  const values = [0];
  for (const ch of pattern) {
    if (/[0-9]/.test(ch)) values[values.length - 1] = Number(ch);
    else { letters.push(ch); values.push(0); }
  }
  return { key: letters.join(''), values };
}

/** A trie of patterns: letter by letter, each ending node holding its numbers. */
function buildTrie(patterns) {
  const root = new Map();
  for (const p of patterns) {
    const { key, values } = parse(p);
    let node = root;
    for (const ch of key) {
      if (!node.has(ch)) node.set(ch, new Map());
      node = node.get(ch);
    }
    const prev = node.get('$');
    node.set('$', prev ? prev.map((v, i) => Math.max(v, values[i] ?? 0)) : values);
  }
  return root;
}

let TRIE = null;
const trie = () => (TRIE ||= buildTrie([...WRITTEN, ...fromClasses()]));
const cache = new Map();

/**
 * The places a word may break: indices into the word, each the position of
 * the first letter after the break. `minLeft`/`minRight` keep that many
 * letters on each side (TeX's English: 2 and 3); a word shorter than six
 * letters, or with anything but letters in it, is left whole.
 */
export function hyphenPoints(word, { minLeft = 2, minRight = 3, minWord = 6 } = {}) {
  const w = String(word || '');
  if (w.length < minWord || !/^[A-Za-z]+$/.test(w)) return [];
  const key = w.toLowerCase();
  let points = cache.get(key);
  if (!points) {
    const s = '.' + key + '.';
    const values = new Array(s.length + 1).fill(0);
    const root = trie();
    for (let i = 0; i < s.length; i++) {
      let node = root;
      for (let j = i; j < s.length; j++) {
        node = node.get(s[j]);
        if (!node) break;
        const end = node.get('$');
        if (end) for (let k = 0; k < end.length; k++) values[i + k] = Math.max(values[i + k], end[k]);
      }
    }
    // values[i] is the gap before s[i]; the word's letter c is s[c + 1].
    // Two breaks a letter apart leave a letter on its own ("run-n-ing"):
    // the stronger of the two stays (the -ing ending is written 3).
    const found = [];
    for (let c = 1; c < key.length; c++) {
      const v = values[c + 1];
      if (v % 2 !== 1) continue;
      const last = found[found.length - 1];
      if (last && c - last.c < 2) { if (v > last.v) found[found.length - 1] = { c, v }; continue; }
      found.push({ c, v });
    }
    points = found.map((f) => f.c);
    cache.set(key, points);
    if (cache.size > 20000) cache.delete(cache.keys().next().value);
  }
  return points.filter((c) => c >= minLeft && key.length - c >= minRight);
}

/** A word with its breaks shown: "hy-phen-ation" — for the Manual dialog and for tests. */
export function hyphenated(word, opts) {
  const pts = hyphenPoints(word, opts);
  let out = '';
  let at = 0;
  for (const p of pts) { out += word.slice(at, p) + '-'; at = p; }
  return out + word.slice(at);
}

/**
 * The document's hyphenation, as the page and the printout need it: on or
 * off, the zone (px) a line's end must be short of the margin by before a
 * word is broken there, how many lines in a row may end in a hyphen (0 for
 * no limit), and whether a word in capitals may be broken.
 */
export function hyphenationRules(settings) {
  if (!settings) return null;
  return {
    auto: Boolean(settings.auto),
    zonePx: (Number(settings.zoneTwips ?? 360) || 0) / 15,
    limit: Math.max(0, Number(settings.limit) || 0),
    caps: settings.caps !== false,
  };
}

/** May this word be broken under these rules — a word in capitals only when capitals may be. */
export function breakableWord(word, rules) {
  if (!rules?.auto) return false;
  if (!rules.caps && /^[A-Z]+$/.test(word.replace(/[^A-Za-z]/g, '')) && /[A-Z]/.test(word)) return false;
  return true;
}

export const SOFT_HYPHEN = '­';
