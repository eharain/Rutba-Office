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

import { registerCatalogue, setLanguage, language, tFilled } from '@rutba/office-ui';
import { CATALOGUES } from '@rutba/office-ui/catalogues';
import { setErrorWords } from '@rutba/office-shell/client';
import { registryInLanguage } from './registry-words.js';

/** The languages written right to left, by their base tag. */
const RIGHT_TO_LEFT = new Set(['ar', 'ckb', 'dv', 'fa', 'he', 'ps', 'sd', 'ug', 'ur', 'yi']);

/** Whether a language tag is written right to left. */
export const rightToLeft = (tag) => RIGHT_TO_LEFT.has(String(tag || '').toLowerCase().split(/[-_]/)[0]);

for (const [tag, messages] of Object.entries(CATALOGUES)) registerCatalogue(tag, messages);

const asked = typeof location === 'undefined' ? null : new URLSearchParams(location.search).get('lang');
const chosen = setLanguage(asked || (typeof navigator === 'undefined' ? 'en' : navigator.language) || 'en');
if (typeof document !== 'undefined') {
  document.documentElement.lang = chosen;
  document.documentElement.dir = rightToLeft(chosen) ? 'rtl' : 'ltr';
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
 * and a window in the system's language is not touched at all.
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
if (chosen !== 'en' && !systemTag.startsWith(chosen)) localeDefaults(chosen);

export { language };
