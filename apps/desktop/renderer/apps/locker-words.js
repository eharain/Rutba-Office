// The Locker's strings go through the same t() as every other window. They are
// imported from here, and their translations added from here, so that the
// translators' list and the shared catalogues do not carry them: a build made
// without the Locker's switch must not carry a word of it.
import { addToCatalogue } from '@rutba/office-ui';
import { LOCKER_WORDS } from './locker-translations.js';

for (const [tag, words] of Object.entries(LOCKER_WORDS)) addToCatalogue(tag, words);

export { t, tn } from '@rutba/office-ui';
