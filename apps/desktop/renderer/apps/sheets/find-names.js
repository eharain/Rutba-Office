// The words Find in Sheets shares with the rest of the suite, through the
// shared t(). Its own new words come from sheets-words.js.
import { t } from '@rutba/office-ui';

export function findNames() {
  return {
    find: t('common.find'),
    replace: t('common.replace'),
    replaceWith: t('common.replace-with'),
    findNext: t('sheets.find-names.find-next'),
    replaceAll: t('common.replace-all'),
    close: t('common.close'),
    title: t('common.find-and-replace'),
    matchCase: t('common.match-case'),
    sheet: t('common.sheet'),
    workbook: t('common.workbook'),
    cell: t('sheets.cell'),
    value: t('common.value'),
    formula: t('common.formula'),
    nothing: t('common.nothing-matched'),
    replacedOne: t('sheets.find-names.replaced-one'),
    nothingLeft: t('sheets.find-names.nothing-left-replace'),
    replacedAll: t('common.replaced-every'),
  };
}
