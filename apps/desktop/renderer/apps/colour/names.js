// The palette's words that the rest of the suite shares, from the shared
// catalogues: the theme's slots, Office's standard colours, Word's highlight
// colours and the palette's buttons. The palette's own new words come from
// colour-words.js instead (see colour-translations.js).
import { t } from '@rutba/office-ui';

export const slotName = (slot) => ({
  lt1: t('Background 1'), dk1: t('Text 1'), lt2: t('Background 2'), dk2: t('Text 2'),
}[slot] || t('Accent {n}', { n: slot.replace('accent', '') }));

export const standardName = (key) => ({
  darkRed: t('Dark red'), red: t('Red'), orange: t('Orange'), yellow: t('Yellow'), lightGreen: t('Light green'),
  green: t('Green'), lightBlue: t('Light blue'), blue: t('Blue'), darkBlue: t('Dark blue'), purple: t('Purple'),
}[key] || key);

export const highlightName = (key) => ({
  yellow: t('Yellow'), green: t('Bright green'), cyan: t('Turquoise'), magenta: t('Pink'), blue: t('Blue'),
  red: t('Red'), darkBlue: t('Dark blue'), darkCyan: t('Teal'), darkGreen: t('Green'), darkMagenta: t('Violet'),
  darkRed: t('Dark red'), darkYellow: t('Dark yellow'), darkGray: t('Grey 50%'), lightGray: t('Grey 25%'), black: t('Black'),
}[key] || key);

/** The palette's other shared words. */
export const words = () => ({
  automatic: t('Automatic'), themeColours: t('Theme colours'), colours: t('Colours'), hue: t('Hue'), cancel: t('Cancel'), ok: t('OK'),
  red: t('Red'), green: t('Green'), blue: t('Blue'), newer: t('New'),
});
