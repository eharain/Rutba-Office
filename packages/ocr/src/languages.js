// The languages text recognition reads: one row a language, the one table
// everything else follows. Text recognition follows the languages the suite
// ships, so a language the suite gains is added here as one row, and its
// model file goes into vendor/tessdata (tools/vendor-ocr.mjs takes the
// list of models from this table and nothing else).
//
//   code       the language as recognition names it (Tesseract's code)
//   ui         the suite's own code for it (its catalogue, its window language)
//   model      the model's file name in vendor/tessdata, without .traineddata
//   script, direction, and beta while its rates are not yet good enough
//              to promise (docs/OCR.md has them)
//
// The names are English: the window translates them. Urdu is beta: its
// model reads Naskh type well and the Nastaliq most Urdu is printed in
// poorly.

export const LANGUAGES = Object.freeze([
  { code: 'eng', ui: 'en', model: 'eng', name: 'English', script: 'Latin', direction: 'ltr' },
  { code: 'urd', ui: 'ur', model: 'urd', name: 'Urdu', script: 'Arabic', direction: 'rtl', beta: true },
  { code: 'ara', ui: 'ar', model: 'ara', name: 'Arabic', script: 'Arabic', direction: 'rtl' },
  { code: 'hin', ui: 'hi', model: 'hin', name: 'Hindi', script: 'Devanagari', direction: 'ltr' },
  { code: 'chi_sim', ui: 'zh', model: 'chi_sim', name: 'Chinese (Simplified)', script: 'Han', direction: 'ltr', unspaced: true },
  { code: 'jpn', ui: 'ja', model: 'jpn', name: 'Japanese', script: 'Japanese', direction: 'ltr', unspaced: true },
].map(Object.freeze));

/** A language's row, by its code. */
export const languageOf = (code) => LANGUAGES.find((l) => l.code === code) || null;

/** The model file name for a language code. */
export const modelOf = (code) => languageOf(code)?.model || code;

export const CODES = Object.freeze(LANGUAGES.map((l) => l.code));

/** Languages written without spaces between words. */
export const UNSPACED = Object.freeze(LANGUAGES.filter((l) => l.unspaced).map((l) => l.code));

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
  const ui = String(uiLanguage || '').toLowerCase();
  const own = LANGUAGES.find((l) => l.ui === ui) || LANGUAGES.find((l) => l.ui === ui.slice(0, 2));
  return own && own.code !== 'eng' ? [own.code, 'eng'] : ['eng'];
}
