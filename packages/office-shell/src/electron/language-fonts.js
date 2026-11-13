// Fonts for a language's letters (L8).
//
// Office ships no fonts: a window draws its words in the computer's own. Most
// computers have a face for every script Office is translated into, so a
// font pack is added only when this one has none: when a language is added,
// and once at the first start of each version for the languages already
// here. A `font-<script>` pack, once in place, is served to the windows and
// put in their font list, ahead of the system's faces where the language asks
// for its own style (Urdu is written in Nastaliq, which an Arabic face that
// has the letters does not draw).
//
// The table is keyed by language where the script alone is not enough, then
// by script. Each row has letters a face must draw to count, and, where the
// style matters, the families that draw it.
import fs from 'node:fs';
import path from 'node:path';
import { listedLanguage } from './language-packs.js';

/**
 * @typedef {{ language?: string, script?: string, pack: string, test: string, families?: string[], first?: boolean }} FontNeed
 * `first`: the pack's face goes ahead of the system's in a window in this
 * language, not only behind them for the letters they lack.
 */

/** @type {FontNeed[]} */
export const FONT_NEEDS = [
  // Urdu in Nastaliq: an Arabic face has the letters, but not the style.
  { language: 'ur', pack: 'font-aran', test: 'ےڈ', families: ['Noto Nastaliq Urdu', 'Jameel Noori Nastaleeq', 'Urdu Typesetting', 'Nafees Nastaleeq', 'Alvi Nastaleeq', 'Mehr Nastaliq Web'], first: true },
  { script: 'Arab', pack: 'font-arab', test: 'بی' },
  { script: 'Hebr', pack: 'font-hebr', test: 'אש' },
  { script: 'Deva', pack: 'font-deva', test: 'हि' },
  { script: 'Cyrl', pack: 'font-cyrl', test: 'бЖ' },
  { script: 'Grek', pack: 'font-grek', test: 'αΩ' },
  // Simplified and Traditional Chinese: a letter only each writes.
  { script: 'Hans', pack: 'font-hans', test: '汉这' },
  { script: 'Hant', pack: 'font-hant', test: '漢這' },
];

const baseOf = (tag) => String(tag || '').split('-')[0];
const lower = (s) => String(s || '').toLowerCase();

/** Chinese by region, where the tag names a region rather than a script. */
const CHINESE = { 'zh-tw': 'Hant', 'zh-hk': 'Hant', 'zh-mo': 'Hant', 'zh-cn': 'Hans', 'zh-sg': 'Hans', 'zh-my': 'Hans', zh: 'Hans' };

/** The script a language is written in: the list's word, else its base's, else null. */
export function scriptOf(tag) {
  const t = lower(tag);
  const script = /^zh-(hans|hant)\b/.exec(t);
  if (script) return script[1] === 'hans' ? 'Hans' : 'Hant';
  return listedLanguage(tag)?.script || CHINESE[t] || listedLanguage(baseOf(tag))?.script || null;
}

/** The rows a language needs met, the language's own first, then its script's. */
export function needsOf(tag) {
  const t = lower(tag);
  const own = FONT_NEEDS.filter((r) => r.language && (lower(r.language) === t || lower(r.language) === baseOf(t)));
  const script = scriptOf(tag);
  return [...own, ...FONT_NEEDS.filter((r) => r.script && r.script === script)];
}

/** Whether a face meets a row: its letters, and its family where the row names them. */
export function faceMeets(face, row) {
  if (row.families && !row.families.some((f) => lower(f) === lower(face.family))) return false;
  return [...row.test].every((c) => face.has(c.codePointAt(0)));
}

/**
 * The font packs to add for some languages.
 * @param {object} o
 * @param {string[]} o.tags         the languages (window tags or lang-<tag> pack ids)
 * @param {(row: FontNeed) => boolean | Promise<boolean>} o.has   whether this computer has a face that meets the row
 * @param {Iterable<string>} [o.installed]  pack ids already here or staged
 * @param {Iterable<string>} [o.offered]    pack ids the feed lists
 * @returns {Promise<{ add: string[], lacking: FontNeed[] }>}  `lacking`: rows met by nothing, here or offered
 */
export async function fontsNeeded({ tags, has, installed = [], offered = [] }) {
  const here = new Set(installed);
  const listed = new Set(offered);
  const add = [];
  const lacking = [];
  for (const raw of tags || []) {
    const tag = String(raw).replace(/^lang-/, '');
    if (baseOf(tag) === 'en') continue;
    for (const row of needsOf(tag)) {
      if (here.has(row.pack) || add.includes(row.pack) || (await has(row))) break;
      if (listed.has(row.pack)) { add.push(row.pack); break; }
      // Not offered: the next row (Urdu's Arabic) may still be met or offered.
      if (!lacking.includes(row)) lacking.push(row);
    }
  }
  return { add, lacking };
}

/** Where each platform keeps the fonts of the computer and of the person. */
export function systemFontDirs({ platform = process.platform, env = process.env, home = env.USERPROFILE || env.HOME || '' } = {}) {
  if (platform === 'win32') {
    return [path.join(env.WINDIR || 'C:\\Windows', 'Fonts'), env.LOCALAPPDATA ? path.join(env.LOCALAPPDATA, 'Microsoft', 'Windows', 'Fonts') : null].filter(Boolean);
  }
  if (platform === 'darwin') return ['/System/Library/Fonts', '/Library/Fonts', home && path.join(home, 'Library', 'Fonts')].filter(Boolean);
  return ['/usr/share/fonts', '/usr/local/share/fonts', home && path.join(home, '.local', 'share', 'fonts'), home && path.join(home, '.fonts')].filter(Boolean);
}

const FONT_FILE = /\.(ttf|otf|ttc|otc)$/i;

/** The font files under some folders, a few levels deep. */
export function fontFiles(dirs, { depth = 4 } = {}) {
  const out = [];
  const walk = (dir, left) => {
    let entries;
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) { if (left > 0) walk(full, left - 1); } else if (FONT_FILE.test(e.name)) out.push(full);
    }
  };
  for (const d of dirs) walk(d, depth);
  return out;
}

/** A face reader that reads only the parts of a file `facesOf` asks for. */
export function readFacesOfFile(facesOf) {
  return (file) => {
    const fd = fs.openSync(file, 'r');
    try {
      const size = fs.fstatSync(fd).size;
      return facesOf((offset, length) => {
        if (offset < 0 || length < 0 || offset + length > size || length > 32 * 1024 * 1024) throw new Error('outside the font');
        const b = Buffer.alloc(length);
        fs.readSync(fd, b, 0, length, offset);
        return b;
      });
    } finally {
      fs.closeSync(fd);
    }
  };
}

/**
 * An async `has(row)` over this computer's fonts: the family of the first
 * face that meets the row, or null. Files are read only when a row is
 * asked about, and each once: a computer that has a face for every language
 * asked stops at the first that meets each.
 * @param {object} o
 * @param {() => string[]} o.files      the font files (fontFiles(systemFontDirs()))
 * @param {(file: string) => { family: string, has: (cp: number) => boolean }[]} o.faces
 */
export function createFontFinder({ files, faces }) {
  let list = null;
  const read = new Map();
  const answers = new Map();
  const facesIn = (file) => {
    if (!read.has(file)) {
      let got = [];
      try { got = faces(file) || []; } catch { /* not a font this reads */ }
      read.set(file, got);
    }
    return read.get(file);
  };
  return async (row) => {
    if (answers.has(row)) return answers.get(row);
    list ||= files() || [];
    // A row that names its families is looked for in the files whose names suggest them first.
    const hints = (row.families || []).map((f) => lower(f).replace(/\s+/g, '').slice(0, 6));
    const named = (file) => hints.some((h) => lower(path.basename(file)).replace(/[\s_-]+/g, '').includes(h));
    const ordered = hints.length ? [...list.filter(named), ...list.filter((f) => !named(f))] : list;
    // Asked once per row however many ask at once; the files read a few at a
    // time, so the main process answers windows between them.
    const asking = (async () => {
      for (let i = 0; i < ordered.length; i++) {
        if (i && i % 8 === 0) await new Promise((r) => setImmediate(r));
        const face = facesIn(ordered[i]).find((f) => faceMeets(f, row));
        if (face) return face.family || ordered[i];
      }
      return null;
    })();
    answers.set(row, asking);
    return asking;
  };
}

/**
 * The faces the installed font packs bring, for the windows: each file a
 * pack holds, by family, with the row it was made for.
 * @param {{ id: string, kind: string, dir: string, manifest: object }[]} addons
 * @param {(file: string) => { family: string }[]} faces
 * @returns {{ pack: string, family: string, file: string, rel: string, first: boolean, language: string|null, script: string|null }[]}
 */
export function fontPackFaces(addons, faces) {
  const out = [];
  for (const a of addons || []) {
    if (a.kind !== 'font' || !a.dir) continue;
    const row = FONT_NEEDS.find((r) => r.pack === a.id) || null;
    for (const file of fontFiles([a.dir], { depth: 2 })) {
      let family = '';
      try { family = faces(file)?.[0]?.family || ''; } catch { /* unreadable */ }
      if (!family) continue;
      const rel = path.relative(a.dir, file).split(path.sep).join('/');
      out.push({ pack: a.id, family, file, rel, first: Boolean(row?.first), language: row?.language || null, script: row?.script || a.manifest?.script || null });
    }
  }
  return out;
}

/** The faces a window in `tag` should be given: those for its language's rows. */
export function facesForLanguage(tag, packFaces) {
  const packs = new Set(needsOf(String(tag).replace(/^lang-/, '')).map((r) => r.pack));
  return (packFaces || []).filter((f) => packs.has(f.pack));
}

const quoted = (family) => `"${String(family).replace(/["\\]/g, '')}"`;

/**
 * What a window in `tag` is told about its fonts: the pack faces to load
 * (`url`), and the system's own face for a language that wants its style
 * first, as the last check found it (`url` null).
 * @param {string} tag
 * @param {{ packFaces: object[], urlOf: (face: object) => string, systemFirst?: Record<string, string> }} o
 * @returns {{ family: string, url: string|null, first: boolean }[]}
 */
export function windowFonts(tag, { packFaces = [], urlOf, systemFirst = {} }) {
  const out = facesForLanguage(tag, packFaces).map((f) => ({ family: f.family, url: urlOf(f), first: f.first }));
  if (!out.some((f) => f.first)) {
    for (const row of needsOf(String(tag).replace(/^lang-/, ''))) {
      const family = row.first && systemFirst[row.pack];
      if (family) { out.unshift({ family, url: null, first: true }); break; }
    }
  }
  return out;
}

/** The two font slots theme.css leaves for packs, each with its trailing comma, or '' when empty. */
export function fontSlots(fonts) {
  const list = (first) => [...new Set((fonts || []).filter((f) => Boolean(f.first) === first).map((f) => quoted(f.family)))];
  const slot = (names) => (names.length ? `${names.join(', ')},` : '');
  return { first: slot(list(true)), last: slot(list(false)) };
}
