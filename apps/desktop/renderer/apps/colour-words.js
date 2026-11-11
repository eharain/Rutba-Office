// The colour palette's words (colour/palette.js) go through the same t() as
// every other window. Its new words are translated here
// (colour-translations.js), so the shared catalogues and the translators'
// list are not changed by the palette alone; the words it shares with the
// rest of the suite stay in the shared catalogues.
import { addToCatalogue } from '@rutba/office-ui';
import { COLOUR_WORDS } from './colour-translations.js';

for (const [tag, words] of Object.entries(COLOUR_WORDS)) addToCatalogue(tag, words);

export { t } from '@rutba/office-ui';
