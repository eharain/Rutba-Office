// What Word, Sheets and Slides say when a file from the Locker opens as it was
// saved, because the unsaved work sealed for it would not open here: that work
// stays in the Locker, and is offered again when its copy opens. Imported only in
// a build that has the Locker (set-aside.js), and its words are kept here,
// apart from the shared catalogues, as the Locker's own are (locker-words.js).
import { addToCatalogue } from '@rutba/office-ui';
import { t } from './locker-words.js';

export const SET_ASIDE = '{name} opened as it was last saved. Its unsaved work could not be opened here, so the Locker keeps it and offers it again the next time it unlocks.';
export const SET_ASIDE_UNREAD = '{name} opened as it was last saved. Its unsaved work could not be opened.';

export const SET_ASIDE_WORDS = {
  ur: {
    [SET_ASIDE]: '{name} آخری بار محفوظ کی گئی حالت میں کھلی ہے۔ اس کا غیر محفوظ کام یہاں نہیں کھل سکا، اس لیے لاکر اسے رکھتا ہے اور اگلی بار ان لاک ہونے پر دوبارہ پیش کرتا ہے۔',
    [SET_ASIDE_UNREAD]: '{name} آخری بار محفوظ کی گئی حالت میں کھلی ہے۔ اس کا غیر محفوظ کام نہیں کھل سکا۔',
  },
  ar: {
    [SET_ASIDE]: 'فُتح {name} كما حُفظ آخر مرة. تعذّر فتح عمله غير المحفوظ هنا، لذا تحتفظ به الخزانة وتعرضه مجددًا عند فتح قفلها في المرة القادمة.',
    [SET_ASIDE_UNREAD]: 'فُتح {name} كما حُفظ آخر مرة. تعذّر فتح عمله غير المحفوظ.',
  },
  hi: {
    [SET_ASIDE]: '{name} वैसी खुली है जैसी आख़िरी बार सहेजी गई थी। इसका बिना सहेजा काम यहाँ नहीं खुल सका, इसलिए लॉकर उसे रखता है और अगली बार अनलॉक होने पर फिर से दिखाता है।',
    [SET_ASIDE_UNREAD]: '{name} वैसी खुली है जैसी आख़िरी बार सहेजी गई थी। इसका बिना सहेजा काम नहीं खुल सका।',
  },
};

for (const [tag, words] of Object.entries(SET_ASIDE_WORDS)) addToCatalogue(tag, words);

/** The sentence, in the window's language, for a file of this name: offered later ('offered'), or not ('unread'). */
export const setAsideLine = (name, how) => t(how === 'unread' ? SET_ASIDE_UNREAD : SET_ASIDE, { name });
