// The words Find in Sheets shares with the rest of the suite, taken from
// the shared catalogues so the list of messages keeps them. Its own new
// words are in sheets-translations.js.
import { t } from '@rutba/office-ui';

export function findNames() {
  return {
    find: t('Find'),
    replace: t('Replace'),
    replaceWith: t('Replace with'),
    findNext: t('Find next'),
    replaceAll: t('Replace all'),
    close: t('Close'),
    title: t('Find and replace'),
    matchCase: t('Match case'),
    sheet: t('Sheet'),
    workbook: t('Workbook'),
    cell: t('Cell'),
    value: t('Value'),
    formula: t('Formula'),
    nothing: t('Nothing matched.'),
    replacedOne: t('Replaced one.'),
    nothingLeft: t('Nothing left to replace.'),
    replacedAll: t('Replaced every match.'),
  };
}
