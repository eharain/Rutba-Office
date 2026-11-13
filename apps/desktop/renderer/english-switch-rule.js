// What the title bar's English switch does (L14, D-2026-11-13-30): one
// click shows the window in English; in English, the same click goes back
// to the language it was in. Kept apart from the button so it can be tried
// without a window.

/** Where the language the switch went from is kept: { setting, shown }. */
export const BEFORE_ENGLISH_KEY = 'language.beforeEnglish';

const lower = (s) => String(s || '').toLowerCase().replace(/_/g, '-');
const isEnglish = (tag) => lower(tag).split('-')[0] === 'en';

/**
 * The switch now, or null when it is not shown (an English window not
 * switched from another language).
 * @param {{ showing: string, before?: { setting: string|null, shown: string }|null }} o
 * @returns {null | { to: 'en', from: string } | { to: string, setting: string|null, from: 'en' }}
 */
export function switchState({ showing, before = null }) {
  const now = lower(showing) || 'en';
  if (!isEnglish(now)) return { to: 'en', from: now };
  const back = before && typeof before === 'object' ? lower(before.shown) : '';
  if (back && !isEnglish(back)) return { to: back, setting: typeof before.setting === 'string' && before.setting ? before.setting : null, from: now };
  return null;
}

/**
 * The words of a switch, for its tooltip and its name to a screen reader:
 * in the language shown and in English, so that someone who cannot read
 * the one finds it by the other. Each once when they are the same.
 * @param {string} shown     the words in the window's language
 * @param {string} english   the same words in English
 */
export const bothLanguages = (shown, english) => (shown && shown !== english ? `${shown} · ${english}` : english);

/**
 * The language setting and the window's address after a click.
 * @returns {{ setting: string|null, before: object|null, lang: string|null }}
 *   `setting` to store as 'language', `before` to store under
 *   BEFORE_ENGLISH_KEY (null to clear it), `lang` for the window's address
 *   (null: the system's).
 */
export function afterClick(state, { setting = null } = {}) {
  if (!state) return null;
  if (state.to === 'en') return { setting: 'en', before: { setting: setting || null, shown: state.from }, lang: 'en' };
  return { setting: state.setting, before: null, lang: state.setting || state.to };
}
