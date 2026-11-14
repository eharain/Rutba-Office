// Meaning keys (D-2026-11-14-10 to -16): how a label's key is made.
//
// A key is `scope.name`. The scope says where the label is used (`locker`,
// `word.ribbon`, `common`) and is never shown. The name is short: the first
// words of the label's meaning, enough to fit its place, unique in its scope
// (`locker.names-apply-later`). With no language pack at all a window shows
// the name with spaces for its hyphens, cut to fit, the whole key as its
// tooltip (D-2026-11-14-14, -16).
//
// Each key's definition carries its space: the most display columns each
// version of its words may take, { short, medium, complete }, measured as the
// packages repository measures it (usage-from-office.mjs): an East Asian wide
// character counts 2 and a placeholder 8 (D-2026-11-14-12).
//
// tools/make-label-keys.mjs used this once to name every label Office had;
// tools/extract-messages.mjs uses it for a label new since.

/** Words that carry no meaning of their own in a name. Negations and quantities stay. */
const STOP = new Set(['a', 'an', 'the', 'to', 'of', 'for', 'and', 'or', 'in', 'on', 'at', 'by', 'with', 'from', 'into', 'onto',
  'is', 'are', 'was', 'were', 'be', 'been', 'this', 'that', 'these', 'those', 'it', 'its', 'as', 'your', 'you', 'yours',
  'will', 'can', 'so', 'than', 'then', 'there', 'here', 'has', 'have', 'had', 'which', 'who', 'whose', 'what', 'when',
  'where', 'their', 'them', 'they', 'our', 'we', 'us', 'i', 'my', 'me', 'his', 'her', 'he', 'she', 'if', 'but', 'up']);

/** A key's shape: lower-case words joined by hyphens, a scope of one or more such parts. */
export const KEY = /^[a-z0-9]+(?:-[a-z0-9]+)*(?:\.[a-z0-9]+(?:-[a-z0-9]+)*)+$/;
/** Whether a string is written as a key (`common.close`), not as words. */
export const isKey = (s) => typeof s === 'string' && KEY.test(s);
/** A key's scope and name: `word.ribbon.page-borders` is `word.ribbon` and `page-borders`. */
export function partsOf(key) {
  const at = key.lastIndexOf('.');
  return { scope: key.slice(0, at), name: key.slice(at + 1) };
}

/** The words of a label's English, lower case, ASCII, a placeholder as its name. */
export function wordsOf(english) {
  return String(english)
    .replace(/\{(\w+)\}/g, ' $1 ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .normalize('NFKD').replace(/[̀-ͯ]/g, '')
    .replace(/&/g, ' and ').replace(/['’]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim().split(/\s+/).filter(Boolean);
}

/**
 * The names a label may take, best first: its first content words, as many as
 * fit `room` columns (one at least, three at most), then more words for a
 * name already taken, then every word of it.
 */
export function namesFor(english, room = 24) {
  const words = wordsOf(english);
  // A label of three words or fewer keeps them all: "Save As" is not "Save".
  const content = words.length <= 3 ? words : words.filter((w) => !STOP.has(w));
  const base = content.length ? content : words;
  if (!base.length) return ['label'];
  const out = [];
  const add = (list) => {
    const n = list.join('-');
    if (n && !out.includes(n)) out.push(n);
  };
  let first = 1;
  while (first < Math.min(3, base.length) && base.slice(0, first + 1).join('-').length <= room) first += 1;
  for (let n = first; n <= Math.min(base.length, 6); n += 1) add(base.slice(0, n));
  add(words.slice(0, 8));
  return out;
}

/**
 * A name for a label in its scope: the first of namesFor() not taken, else
 * the last with a number. `taken` is the set of keys given so far; the key
 * chosen is added to it.
 */
export function nameIn(scope, english, taken, room) {
  const names = namesFor(english, room);
  for (const n of names) {
    if (!taken.has(`${scope}.${n}`)) {
      taken.add(`${scope}.${n}`);
      return n;
    }
  }
  const last = names[names.length - 1];
  for (let i = 2; ; i += 1) {
    if (!taken.has(`${scope}.${last}-${i}`)) {
      taken.add(`${scope}.${last}-${i}`);
      return `${last}-${i}`;
    }
  }
}

// ── space ────────────────────────────────────────────────────────────────

const WIDE = /[ᄀ-ᅟ⺀-〾ぁ-㏿㐀-䶿一-鿿ꀀ-꓏가-힣豈-﫿︰-﹏＀-｠￠-￦]|[\u{20000}-\u{3FFFD}]/u;
/** Display columns: East Asian wide and full-width characters count 2 (UAX #11); a placeholder counts as 8. */
export const widthOf = (s) => {
  let w = 0;
  for (const ch of String(s).replace(/\{\w+\}/g, '\u0000'.repeat(8))) w += WIDE.test(ch) ? 2 : 1;
  return w;
};
/** W3C and IBM growth from English to other languages, by the English's width. */
const growth = (w) => (w <= 10 ? 3 : w <= 20 ? 2 : w <= 30 ? 1.8 : w <= 50 ? 1.6 : w <= 70 ? 1.7 : 1.3);

/**
 * A new label's space, by the packages repository's rule for running text:
 * the English grown as other languages grow it; a short version drops words,
 * never a placeholder; the complete version half as wide again.
 */
export function spaceOf(english, plural = null) {
  const text = plural && plural.length > english.length ? plural : english;
  const w = widthOf(text);
  const medium = Math.max(w, Math.ceil(w * growth(w)));
  const held = 8 * (String(text).match(/\{\w+\}/g) || []).length;
  const short = Math.max(3, Math.ceil((medium - held) / 2)) + held;
  return { short, medium, complete: Math.ceil(medium * 1.5) };
}

// ── scope ────────────────────────────────────────────────────────────────

/** The suite's apps: a label used only inside one of them is scoped to it. */
const APPS = ['word', 'sheets', 'slides', 'video', 'mail', 'base', 'pdf', 'image', 'pictures', 'calendar', 'contacts',
  'talk', 'broadcast', 'colour', 'home', 'locker', 'recorder'];

/** The area a file's labels belong to: [app or module, part within it or null]. As the packages extractor gave scopes. */
function areaOf(file) {
  const f = file.replace(/\\/g, '/');
  let m;
  if ((m = /^apps\/desktop\/renderer\/apps\/([^/]+)\/([^/]+)\.js$/.exec(f))) return [m[1], m[2]];
  if ((m = /^apps\/desktop\/renderer\/apps\/([^/]+)\.js$/.exec(f))) {
    const base = m[1];
    if (APPS.includes(base)) return [base, null];
    const app = APPS.find((a) => base.startsWith(`${a}-`));
    if (app) return [app, base.slice(app.length + 1)];
    return [base, null];
  }
  if ((m = /^apps\/desktop\/renderer\/studio\/([^/]+)\.js$/.exec(f))) return ['studio', m[1]];
  if (f === 'apps/desktop/renderer/engine-words.js') return ['engine', null];
  if (f === 'apps/desktop/renderer/shared-words.js') return ['common', null];
  if (f === 'apps/desktop/renderer/registry-words.js') return ['registry', null];
  if ((m = /^apps\/desktop\/renderer\/([^/]+)\/([^/]+)\.js$/.exec(f))) return [m[1], m[2]];
  if ((m = /^apps\/desktop\/renderer\/([^/]+)\.js$/.exec(f))) return [m[1], null];
  if ((m = /^packages\/office-ui\/src\/([^/]+)\.js$/.exec(f))) return ['ui', m[1] === 'index' ? null : m[1]];
  if ((m = /^packages\/office-shell\/src\/electron\/([^/]+)\.js$/.exec(f))) return ['shell', m[1]];
  return ['common', null];
}

/** A new label's scope, from every file that uses it: one area's, else `common`. */
export function scopeOf(where) {
  const areas = where.map(areaOf);
  const tops = new Set(areas.map((a) => a[0]));
  if (tops.size !== 1) return 'common';
  const [top] = tops;
  const parts = new Set(areas.map((a) => (a[1] === 'index' ? null : a[1])));
  const one = parts.size === 1 ? [...parts][0] : null;
  const scope = one ? `${top}.${one}` : top;
  return scope.toLowerCase().replace(/[^a-z0-9.]+/g, '-').replace(/(^-|-$)/g, '');
}
