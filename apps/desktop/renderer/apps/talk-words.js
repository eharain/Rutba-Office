// Talk's own words for Office 2.25 and on go through the same t() as every
// other window. They are imported from here, and their translations added
// from here (talk-translations.js), so the shared catalogues and the
// translators' list are not changed by Talk's work alone. Talk's words from
// before stay where they were, in the shared catalogues.
import { addToCatalogue } from '@rutba/office-ui';
import { TALK_WORDS } from './talk-translations.js';

for (const [tag, words] of Object.entries(TALK_WORDS)) addToCatalogue(tag, words);

export { t, tn } from '@rutba/office-ui';
