// What Video says of a film from the Locker (owner, 13 Nov, D-2026-11-13-22:
// a locked film plays, but Video does not edit or export it). Video imports
// this only in a build that has the Locker, and its words are kept here, apart
// from the shared catalogues, as the Locker's own are (locker-words.js): a
// build made without the Locker must not carry a word of it.
import { addToCatalogue } from '@rutba/office-ui';
import { t } from '../locker-words.js';

export const LOCKED_FILM = '{name} is in the Locker. Video plays it, but does not edit or export it. To edit it, unlock it in the Locker first.';

export const LOCKED_FILM_WORDS = {
  ur: { [LOCKED_FILM]: '{name} لاکر میں ہے۔ ویڈیو اسے چلاتی ہے، لیکن اس میں ترمیم یا ایکسپورٹ نہیں کرتی۔ ترمیم کرنی ہو تو پہلے اسے لاکر میں ان لاک کریں۔' },
  ar: { [LOCKED_FILM]: '{name} في الخزانة. يشغّله تطبيق الفيديو، لكنه لا يحرّره ولا يصدّره. لتحريره، افتح قفله في الخزانة أولًا.' },
  hi: { [LOCKED_FILM]: '{name} लॉकर में है। वीडियो इसे चलाता है, पर इसमें संपादन या निर्यात नहीं करता। संपादन करना हो तो पहले इसे लॉकर में अनलॉक करें।' },
};

for (const [tag, words] of Object.entries(LOCKED_FILM_WORDS)) addToCatalogue(tag, words);

/** The sentence, in the window's language, for a film of this name. */
export const lockedFilm = (name) => t(LOCKED_FILM, { name });
