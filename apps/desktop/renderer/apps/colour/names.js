// The palette's words that the rest of the suite shares, through the shared
// t(): the theme's slots, Office's standard colours, Word's highlight colours
// and the palette's buttons. The palette's own new words come from
// colour-words.js instead.
import { t } from '@rutba/office-ui';

export const slotName = (slot) => ({
  lt1: t('common.background-1'), dk1: t('common.text-1'), lt2: t('common.background-2'), dk2: t('common.text-2'),
}[slot] || t('common.accent-n', { n: slot.replace('accent', '') }));

export const standardName = (key) => ({
  darkRed: t('common.dark-red'), red: t('common.red'), orange: t('common.orange'), yellow: t('common.yellow'), lightGreen: t('common.light-green'),
  green: t('common.green'), lightBlue: t('common.light-blue'), blue: t('common.blue'), darkBlue: t('common.dark-blue'), purple: t('common.purple'),
}[key] || key);

export const highlightName = (key) => ({
  yellow: t('common.yellow'), green: t('colour.names.bright-green'), cyan: t('colour.names.turquoise'), magenta: t('common.pink'), blue: t('common.blue'),
  red: t('common.red'), darkBlue: t('common.dark-blue'), darkCyan: t('common.teal'), darkGreen: t('common.green'), darkMagenta: t('colour.names.violet'),
  darkRed: t('common.dark-red'), darkYellow: t('colour.names.dark-yellow'), darkGray: t('colour.names.grey-50'), lightGray: t('colour.names.grey-25'), black: t('common.black'),
}[key] || key);

/** The palette's other shared words. */
export const words = () => ({
  automatic: t('common.automatic'), themeColours: t('common.theme-colours'), colours: t('common.colours'), hue: t('common.hue'), cancel: t('common.cancel'), ok: t('common.ok'),
  red: t('common.red'), green: t('common.green'), blue: t('common.blue'), newer: t('common.new'),
});
