// The language a window is shown in, settled before any app is loaded.
//
// Imported first by index.js, so every module after it, the tables of words
// they make as they load among them, already speaks the window's language:
// a window keeps one language for its whole life, as Office's do, and a
// change in the app menu takes effect in the windows opened after it.
//
// The language comes in the window's address (`lang`, from the "language"
// setting, put there by the shell), else the system's. A right-to-left
// language turns the window's own frame right to left — the ribbon, the
// panes, the dialogs — while a page, a grid or a slide keeps the direction
// its document gives it (they say `dir` for themselves).
//
// The words come from the language's pack when one is installed (L1): the
// main process makes the table from it (office-shell electron/language-packs.js)
// and the window asks for it once, here, before anything else. An installed
// pack replaces the built-in Urdu, Arabic or Hindi table for its tag, never
// merged with it; with no pack, the built-in table stands. English comes from
// lang-en the same way, and a label the window's language lacks is shown in it.
// A pack of format 2 gives its labels' short and complete versions as well,
// which a ribbon too narrow for its buttons' names shows (L3).

import { registerCatalogue, setLanguage, language, setDigits, numberTag, tFilled, setFallback, setVersions } from '@rutba/office-ui';
import { CATALOGUES } from '@rutba/office-ui/catalogues';
import { setErrorWords } from '@rutba/office-shell/client';
import { registryInLanguage } from './registry-words.js';

/** The languages written right to left, by their base tag. */
const RIGHT_TO_LEFT = new Set(['ar', 'ckb', 'dv', 'fa', 'he', 'ps', 'sd', 'ug', 'ur', 'yi']);

/** The directions the list of languages gives (L4), filled in once the main process has answered below. */
let listed = {};

/** Whether a language tag is written right to left: as the list of languages says, else by the set above. */
export function rightToLeft(tag) {
  const t = String(tag || '').toLowerCase().replace(/_/g, '-');
  const said = listed[t] || listed[t.split('-')[0]];
  if (said === 'rtl' || said === 'ltr') return said === 'rtl';
  return RIGHT_TO_LEFT.has(t.split('-')[0]);
}

for (const [tag, messages] of Object.entries(CATALOGUES)) registerCatalogue(tag, messages);

const query = typeof location === 'undefined' ? new URLSearchParams() : new URLSearchParams(location.search);
const asked = query.get('lang');
const wanted = asked || (typeof navigator === 'undefined' ? 'en' : navigator.language) || 'en';

/** The words the main process made for this language from its pack, or null. */
function packWords(tag) {
  try {
    const w = globalThis.rutbaOffice?.__words?.(tag);
    return w && typeof w === 'object' ? w : null;
  } catch {
    return null;
  }
}
const words = packWords(wanted);
if (words?.directions && typeof words.directions === 'object') listed = words.directions;
if (words?.tag && words.table && typeof words.table === 'object') registerCatalogue(words.tag, words.table);
setFallback(words?.english || null);
// The labels' short and complete versions (L3): the language's own, and lang-en's for a label shown in English.
setVersions(words?.versions || null, words?.englishVersions || null);
/** The installed languages, as the main process found them: for the language menu (L4). */
export const installedLanguages = () => (Array.isArray(words?.languages) ? words.languages : []);

const chosen = setLanguage(wanted);
// The digits: as the setting says (Western 0 to 9, Arabic's ٠ to ٩, Urdu's
// ۰ to ۹), else 0 to 9 in an Arabic window, as Office in Arabic writes them
// by default, else the language's own.
const digitsAsked = query.get('digits');
setDigits(digitsAsked || (chosen === 'ar' ? 'latn' : null));
if (typeof document !== 'undefined') {
  document.documentElement.lang = chosen;
  // The pack's own word on its direction, else the list above.
  const packDirection = words?.tag === chosen ? words.direction : null;
  document.documentElement.dir = packDirection === 'rtl' || packDirection === 'ltr' ? packDirection : rightToLeft(chosen) ? 'rtl' : 'ltr';
  useFonts(words?.fonts);
}

/**
 * The fonts the main process found for this language (L8): a font pack's
 * faces, loaded from the add-on's own folder (rutba://app/addons/), and the two slots theme.css leaves for them,
 * ahead of the frame's faces for a language that wants its own style, or
 * before the generic ones for letters this computer has no face for.
 */
function useFonts(fonts) {
  if (!fonts || typeof fonts !== 'object') return;
  const cssString = (s) => `"${String(s).replace(/["\\\n\r]/g, '')}"`;
  const faces = (Array.isArray(fonts.faces) ? fonts.faces : []).filter((f) => f && typeof f.family === 'string' && typeof f.url === 'string' && f.url.startsWith('rutba://app/addons/'));
  if (faces.length) {
    const style = document.createElement('style');
    style.dataset.rutbaFonts = '';
    style.textContent = faces.map((f) => `@font-face { font-family: ${cssString(f.family)}; src: url(${cssString(f.url)}); font-display: swap; }`).join('\n');
    document.head.appendChild(style);
  }
  // As fontSlots writes them: quoted families, each followed by a comma.
  const slot = (v) => (typeof v === 'string' && /^"[^"\\\n\r]*"(, "[^"\\\n\r]*")*,$/.test(v) ? v : '');
  const root = document.documentElement.style;
  const first = slot(fonts.first);
  const last = slot(fonts.last);
  if (first) root.setProperty('--font-pack-first', first);
  if (last) root.setProperty('--font-pack', last);
}
// The app names and kinds of file every window shows, in the same language.
registryInLanguage();
// And what the engines and services say, as each message reaches the window.
setErrorWords(tFilled);

/**
 * Dates and numbers in the window's language too. Every place the windows
 * format one asks for the default locale (toLocaleDateString(undefined, …)),
 * which is the system's; a window set to another language makes that default
 * its own, so a date in an Urdu window reads 9 اکتوبر 2026, as Urdu writes it.
 * A call that names its locale (a document's own, a sheet's) is left as it is,
 * and a window in the system's language, its digits not chosen, is not
 * touched at all. Chosen digits come along: 10 أكتوبر 2026 becomes
 * ١٠ أكتوبر ٢٠٢٦ when Arabic's own digits are set.
 */
function localeDefaults(tag) {
  const own = (locales) => (locales === undefined || (Array.isArray(locales) && !locales.length) ? tag : locales);
  const wrap = (proto, name) => {
    const base = proto[name];
    Object.defineProperty(proto, name, {
      configurable: true,
      writable: true,
      value: function localised(locales, options) { return base.call(this, own(locales), options); },
    });
  };
  for (const name of ['toLocaleString', 'toLocaleDateString', 'toLocaleTimeString']) wrap(Date.prototype, name);
  wrap(Number.prototype, 'toLocaleString');
}
const systemTag = typeof navigator === 'undefined' ? 'en' : String(navigator.language || 'en').toLowerCase();
if ((chosen !== 'en' && !systemTag.startsWith(chosen)) || digitsAsked) localeDefaults(numberTag());

export { language };
