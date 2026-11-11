// Sheets' own new words (sheets/find.js) go through the same t() as every
// other window. They are translated here (sheets-translations.js), so the
// shared catalogues and the translators' list are not changed by them; the
// words Sheets shares with the rest of the suite stay in the shared
// catalogues (sheets/find-names.js takes those).
import { addToCatalogue } from '@rutba/office-ui';
import { SHEETS_WORDS } from './sheets-translations.js';

for (const [tag, words] of Object.entries(SHEETS_WORDS)) addToCatalogue(tag, words);

export { t } from '@rutba/office-ui';
