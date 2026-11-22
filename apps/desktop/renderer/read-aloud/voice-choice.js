// Which voice reads (2.26.1): the choice the reader makes for a language, and
// the order the spoken words offer the added voices in. Kept apart from the
// reader and the windows, with nothing imported, so the engine tests hold it.
//
// A voice chosen: { kind: 'system', name } for one of the system's voices by
// name, { kind: 'system' } for the system's own, or { kind: 'added', id }.

/** The default voice in the profile's settings: an added voice's id, set in Voices and languages; '' for none. */
export const DEFAULT_VOICE_KEY = 'speech.defaultVoice';

/** Whether an added voice ({ language, also }) reads a language; any language when none is said. */
export function reads(voice, language) {
  const base = String(language || '').toLowerCase().split('-')[0];
  if (!base) return true;
  return [voice?.language, ...(voice?.also || [])].some((l) => String(l || '').toLowerCase().startsWith(base));
}

/**
 * The voice for a language: the one chosen for it in the Read Aloud bar (the
 * system's own is a choice too, and stands over the default), else the
 * default voice when it is added and reads that language, else null, the
 * system's own.
 * @param {object} o
 * @param {object} o.chosen        { [language or '*']: voice } as the reader keeps them
 * @param {string} o.defaultId     the default voice's id, or ''
 * @param {object[]} o.added       the added voices that can speak: [{ id, language, also }]
 * @param {string|null} o.language
 * @param {boolean} [o.engineFailed]  the speech engine stopped: added voices are passed over
 */
export function voiceFor({ chosen = {}, defaultId = '', added = [], language = null, engineFailed = false }) {
  const key = String(language || '').toLowerCase() || '*';
  const base = key.split('-')[0];
  const v = chosen[key] || chosen[base] || chosen['*'] || null;
  if (v) return v.kind === 'added' && engineFailed ? null : v;
  const d = defaultId && !engineFailed ? added.find((a) => a.id === defaultId) : null;
  return d && (key === '*' || reads(d, key)) ? { kind: 'added', id: d.id } : null;
}

/**
 * The language nothing on this computer reads: the sentence's own, when no
 * system voice speaks it and the voice that reads is not an added one. The
 * system's voice then reads words it does not know, badly or not at all, and
 * the bar says why. Null when the language is read, when none is said, and
 * while the system has not listed its voices (an empty list says nothing).
 * @param {object} o
 * @param {string|null} o.language
 * @param {object|null} o.voice     the voice that reads it, as voiceFor answers
 * @param {{ lang: string }[]} o.system  the system's voices
 */
export function unread({ language = null, voice = null, system = [] } = {}) {
  const base = (tag) => String(tag || '').toLowerCase().split(/[-_]/)[0];
  if (!base(language) || voice?.kind === 'added' || !system.length) return null;
  return system.some((v) => base(v?.lang) === base(language)) ? null : language;
}

/** The added voices with the default first, so it is the one offered until another is picked. */
export function defaultFirst(voices, defaultId) {
  return defaultId ? [...voices.filter((v) => v.id === defaultId), ...voices.filter((v) => v.id !== defaultId)] : voices;
}
