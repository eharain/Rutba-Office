// The Locker's strings go through the same t() as every other window. They are
// imported from here, and their translations added from here, so that the
// shared catalogues do not carry them: a build made without the Locker's
// switch carries none of the Locker's translations. Its labels are still named
// in the labels' definitions (catalogues/labels.json, part "locker"), which
// every build carries, so a language pack words them; a
// build without the Locker does not count them (main.js codeKeys).
import { addToCatalogue, addKeys } from '@rutba/office-ui';
import { LOCKER_WORDS } from './locker-translations.js';
import { LOCKER_KEYS } from './locker-keys.js';

addKeys(LOCKER_KEYS);

for (const [tag, words] of Object.entries(LOCKER_WORDS)) addToCatalogue(tag, words);

export { t, tn } from '@rutba/office-ui';
