// Windows language ids (LCIDs), as the language tags a .docx, .pptx or
// .xlsx names a run's language by: the ones Office writes most, each region
// its own. An id not here has no tag. Pure.

const TAGS = {
  0x0401: 'ar-SA', 0x0801: 'ar-IQ', 0x0c01: 'ar-EG', 0x1001: 'ar-LY', 0x1401: 'ar-DZ', 0x1801: 'ar-MA', 0x1c01: 'ar-TN', 0x2001: 'ar-OM',
  0x2401: 'ar-YE', 0x2801: 'ar-SY', 0x2c01: 'ar-JO', 0x3001: 'ar-LB', 0x3401: 'ar-KW', 0x3801: 'ar-AE', 0x3c01: 'ar-BH', 0x4001: 'ar-QA',
  0x0402: 'bg-BG', 0x0403: 'ca-ES', 0x0404: 'zh-TW', 0x0804: 'zh-CN', 0x0c04: 'zh-HK', 0x1004: 'zh-SG', 0x0405: 'cs-CZ', 0x0406: 'da-DK',
  0x0407: 'de-DE', 0x0807: 'de-CH', 0x0c07: 'de-AT', 0x0408: 'el-GR', 0x0409: 'en-US', 0x0809: 'en-GB', 0x0c09: 'en-AU', 0x1009: 'en-CA',
  0x1409: 'en-NZ', 0x1809: 'en-IE', 0x1c09: 'en-ZA', 0x4009: 'en-IN', 0x4809: 'en-SG', 0x040a: 'es-ES_tradnl', 0x080a: 'es-MX', 0x0c0a: 'es-ES',
  0x040b: 'fi-FI', 0x040c: 'fr-FR', 0x080c: 'fr-BE', 0x0c0c: 'fr-CA', 0x100c: 'fr-CH', 0x040d: 'he-IL', 0x040e: 'hu-HU', 0x040f: 'is-IS',
  0x0410: 'it-IT', 0x0411: 'ja-JP', 0x0412: 'ko-KR', 0x0413: 'nl-NL', 0x0813: 'nl-BE', 0x0414: 'nb-NO', 0x0814: 'nn-NO', 0x0415: 'pl-PL',
  0x0416: 'pt-BR', 0x0816: 'pt-PT', 0x0418: 'ro-RO', 0x0419: 'ru-RU', 0x041a: 'hr-HR', 0x041b: 'sk-SK', 0x041c: 'sq-AL', 0x041d: 'sv-SE',
  0x041e: 'th-TH', 0x041f: 'tr-TR', 0x0420: 'ur-PK', 0x0820: 'ur-IN', 0x0421: 'id-ID', 0x0422: 'uk-UA', 0x0423: 'be-BY', 0x0424: 'sl-SI',
  0x0425: 'et-EE', 0x0426: 'lv-LV', 0x0427: 'lt-LT', 0x0429: 'fa-IR', 0x042a: 'vi-VN', 0x042b: 'hy-AM', 0x042c: 'az-Latn-AZ', 0x042d: 'eu-ES',
  0x042f: 'mk-MK', 0x0436: 'af-ZA', 0x0437: 'ka-GE', 0x0439: 'hi-IN', 0x043e: 'ms-MY', 0x043f: 'kk-KZ', 0x0441: 'sw-KE', 0x0445: 'bn-IN',
  0x0845: 'bn-BD', 0x0446: 'pa-IN', 0x0447: 'gu-IN', 0x0449: 'ta-IN', 0x044a: 'te-IN', 0x044b: 'kn-IN', 0x044c: 'ml-IN', 0x044e: 'mr-IN',
  0x0461: 'ne-NP', 0x0463: 'ps-AF', 0x048c: 'prs-AF', 0x045a: 'syr-SY', 0x0465: 'dv-MV', 0x0464: 'fil-PH', 0x0452: 'cy-GB', 0x083c: 'ga-IE',
  0x081a: 'sr-Latn-CS', 0x0c1a: 'sr-Cyrl-CS', 0x141a: 'bs-Latn-BA', 0x0443: 'uz-Latn-UZ', 0x0428: 'tg-Cyrl-TJ', 0x0440: 'ky-KG', 0x0442: 'tk-TM',
  0x0480: 'ug-CN', 0x0492: 'ku-Arab-IQ', 0x0459: 'sd-Arab-PK', 0x0859: 'sd-Arab-PK', 0x045b: 'si-LK', 0x0453: 'km-KH', 0x0454: 'lo-LA', 0x0455: 'my-MM',
};

/** A Windows language id's tag ("ar-SA"), or null. */
export function languageTag(lcid) {
  return TAGS[lcid] ?? null;
}

/** Whether a language reads right to left. */
export const rightToLeft = (tag) => /^(ar|he|fa|ur|ps|prs|syr|dv|ug|sd|ku-Arab|yi)(-|$)/.test(tag || '');
