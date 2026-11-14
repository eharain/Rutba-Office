// The words the main process shows itself (the question before a window
// with unsaved work closes, the menu) in the language the windows are in.
//
// The same catalogue as the windows' (@rutba/office-ui/messages), with the
// language taken from the "language" setting, else the system's, each time
// the main process is about to speak: a language chosen in a window's app
// menu reaches the next question at once, and the menu at the next start.
//
// Its words come from the language's pack when one is installed (L1,
// language-packs.js), laid over any built-in words as the windows lay theirs
// (registerPack, D-2026-11-14-16). Office carries none of its own since then:
// Urdu, Arabic and Hindi are packs like every other language. The windows ask the same
// source for theirs as they start (WORDS_CHANNEL, below), so a window and
// the main process never disagree on a word.

import { app, ipcMain } from 'electron';
import { t, tn, setLanguage, registerCatalogue, registerPack, setFallback } from '@rutba/office-ui/messages';
import { CATALOGUES } from '@rutba/office-ui/catalogues';
import { WORDS_CHANNEL } from '../contract.js';
import { listedDirections } from './language-packs.js';

for (const [tag, messages] of Object.entries(CATALOGUES)) registerCatalogue(tag, messages);

let source = null;
const registered = new Map();
/** The language packs each open window's words came from, and the main process's: what may not be removed now (L6). */
const windowsIn = new Map();
let mainIn = [];

/**
 * A language's words from the packs (language-packs.js createLanguageWords),
 * or null when there is no source yet or it fails: the built-in tables stand.
 */
export function wordsFor(tag) {
  try {
    return source?.wordsFor(tag) || null;
  } catch {
    return null;
  }
}

/**
 * Where the words come from this run, from the moment the add-ons are found.
 * Each window asks for its own language once, as its preload starts; the
 * answer is the table made for that language, made once.
 */
export function useLanguageWords(words) {
  source = words;
  registered.clear();
  ipcMain.removeAllListeners(WORDS_CHANNEL);
  ipcMain.on(WORDS_CHANNEL, (event, tag) => {
    const w = wordsFor(String(tag || ''));
    const sender = event.sender;
    if (sender && typeof sender.id === 'number') {
      if (!windowsIn.has(sender.id)) sender.once?.('destroyed', () => windowsIn.delete(sender.id));
      windowsIn.set(sender.id, w?.uses || []);
    }
    const fonts = safeFonts(w?.tag || String(tag || ''));
    event.returnValue = w ? { ...w, languages: safeLanguages(), directions: listedDirections(), fonts } : { tag: null, table: null, english: null, languages: safeLanguages(), directions: listedDirections(), fonts };
  });
}

/** The language packs in use now, by id: every open window's and the main process's. */
export const languagePacksInUse = () => [...new Set([...[...windowsIn.values()].flat(), ...mainIn])].sort();

/** The fonts a window in this language is given (L8): { faces: [{ family, url }], first, last }, or null. */
const safeFonts = (tag) => {
  try {
    return source?.fontsFor?.(tag) || null;
  } catch {
    return null;
  }
};

const safeLanguages = () => {
  try {
    return source?.languages() || [];
  } catch {
    return [];
  }
};

/** The table a language's words give, registered once for the main process's own catalogue. */
function adopt(tag) {
  const w = wordsFor(tag);
  if (!w) return [];
  if (w.table && w.tag && registered.get(w.tag) !== w.table) {
    registerPack(w.tag, w.table);
    registered.set(w.tag, w.table);
  }
  setFallback(w.english || null);
  return w.uses || [];
}

/** Settle the main process's language from the settings, before it says something. */
export function speakAs(stores) {
  let chosen = null;
  try {
    chosen = stores?.settings?.get('language', null) || null;
  } catch {
    chosen = null;
  }
  const tag = chosen || (app.isReady() ? app.getLocale() : null) || 'en';
  mainIn = adopt(tag);
  return setLanguage(tag);
}

export { t, tn };
