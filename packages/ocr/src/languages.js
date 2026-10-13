// The languages text recognition reads, each a model shipped inside the
// suite (vendor/tessdata). The names are English: the window translates
// them. Urdu is marked beta: its model reads Naskh type well and the
// Nastaliq most Urdu is printed in poorly (docs/OCR.md has the numbers).

export const LANGUAGES = Object.freeze([
  { code: 'eng', name: 'English', script: 'Latin', direction: 'ltr' },
  { code: 'urd', name: 'Urdu', script: 'Arabic', direction: 'rtl', beta: true },
  { code: 'ara', name: 'Arabic', script: 'Arabic', direction: 'rtl' },
  { code: 'hin', name: 'Hindi', script: 'Devanagari', direction: 'ltr' },
  { code: 'chi_sim', name: 'Chinese (Simplified)', script: 'Han', direction: 'ltr' },
  { code: 'jpn', name: 'Japanese', script: 'Japanese', direction: 'ltr' },
].map(Object.freeze));

export const CODES = Object.freeze(LANGUAGES.map((l) => l.code));

/** Languages written without spaces between words. */
export const UNSPACED = Object.freeze(['chi_sim', 'jpn']);

export class OcrLanguageError extends Error {
  constructor(code) {
    super(`There is no text recognition for "${code}".`);
    this.name = 'OcrLanguageError';
    this.code = 'OCR_LANGUAGE';
  }
}

/** The languages asked for, checked, without repeats; English when none are named. */
export function checkLanguages(languages) {
  const list = (Array.isArray(languages) ? languages : String(languages || 'eng').split('+')).map((l) => String(l).trim()).filter(Boolean);
  const out = [];
  for (const code of list.length ? list : ['eng']) {
    if (!CODES.includes(code)) throw new OcrLanguageError(code);
    if (!out.includes(code)) out.push(code);
  }
  return out;
}

/** What to read first for a window speaking this language: its own script, then English. */
export function languagesFor(uiLanguage) {
  const own = { ur: 'urd', ar: 'ara', hi: 'hin', zh: 'chi_sim', ja: 'jpn' }[String(uiLanguage || '').slice(0, 2)];
  return own ? [own, 'eng'] : ['eng'];
}
