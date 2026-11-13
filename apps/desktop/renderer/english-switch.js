// The English switch (L14, D-2026-11-13-30): one button, in the same place
// in every window's title bar, while a language other than English is
// shown. One click shows the window in English, at once and with no
// restart: the window opens again in English, its document with it. In
// English, the same button goes back to the language it was in. Its name
// and tooltip are in the window's language and in English, it is a real
// button (Tab, Enter, Space), and a window with changes not saved is not
// opened again: it says so, in both languages, and the windows opened
// after it are in English.

import React, { useCallback, useEffect, useState } from 'react';
import { Icon, useToast, t, msg, language } from '@rutba/office-ui';
import { BEFORE_ENGLISH_KEY, switchState, bothLanguages, afterClick } from './english-switch-rule.js';

export const ENGLISH_SWITCH_WORDS = [
  msg('Show this window in English'),
  msg('Save your changes first: this window opens in English once it opens again. The windows you open from now on are in English.'),
];

const TO_ENGLISH = 'Show this window in English';
const SAVE_FIRST = 'Save your changes first: this window opens in English once it opens again. The windows you open from now on are in English.';

/** A window's address with its language set, or taken off for the system's. */
const withLanguage = (href, tag) => {
  const url = new URL(href);
  if (tag) url.searchParams.set('lang', tag);
  else url.searchParams.delete('lang');
  return url.toString();
};

/**
 * @param {object} p
 * @param {object} p.shell
 * @param {boolean} [p.dirty]   the window has changes not saved
 * @param {(tag: string) => string} [p.nameOf]  a language's name, in its own words
 */
export function EnglishSwitch({ shell, dirty = false, nameOf = (tag) => tag }) {
  const toast = useToast();
  const [before, setBefore] = useState(undefined);
  useEffect(() => {
    let alive = true;
    shell.store.get({ key: BEFORE_ENGLISH_KEY, fallback: null }).then((v) => { if (alive) setBefore(v || null); }).catch(() => { if (alive) setBefore(null); });
    return () => { alive = false; };
  }, [shell]);
  const state = before === undefined ? (language() === 'en' ? null : switchState({ showing: language() })) : switchState({ showing: language(), before });

  const click = useCallback(async () => {
    if (!state) return;
    const setting = await shell.store.get({ key: 'language', fallback: null }).catch(() => null);
    const next = afterClick(state, { setting });
    await shell.store.set({ key: 'language', value: next.setting || null });
    await shell.store.set({ key: BEFORE_ENGLISH_KEY, value: next.before });
    if (dirty) {
      toast(state.to === 'en' ? bothLanguages(t(SAVE_FIRST), SAVE_FIRST) : `Save your changes first: this window opens in ${nameOf(state.to)} once it opens again. The windows you open from now on are in ${nameOf(state.to)}.`, { ms: 8000 });
      return;
    }
    window.location.replace(withLanguage(window.location.href, next.lang));
  }, [shell, state, dirty, toast, nameOf]);

  if (!state) return null;
  // Back from English: said in English, which the window is in, with the language named in its own words.
  const words = state.to === 'en' ? bothLanguages(t(TO_ENGLISH), TO_ENGLISH) : `Show this window in ${nameOf(state.to)}`;
  return (
    <button type="button" className="rw-btn ghost english-switch" data-tip={words} aria-label={words} lang={state.to === 'en' ? undefined : 'en'} onClick={click}>
      <Icon name="globe" size={13} />
      <span className="english-switch-to" aria-hidden="true" lang={state.to === 'en' ? 'en' : state.to}>{state.to === 'en' ? 'EN' : nameOf(state.to)}</span>
    </button>
  );
}

export const ENGLISH_SWITCH_CSS = `
.english-switch { display: inline-flex; align-items: center; gap: 3px; }
.english-switch-to { font-size: 11px; font-weight: 600; letter-spacing: 0.02em; }
`;
