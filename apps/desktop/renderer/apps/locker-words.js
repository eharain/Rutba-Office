// The Locker's labels go through the same t() as every other window, taken
// from here (locker.js, locker-tile.js), so they carry the part's name in the
// labels' definitions (catalogues/labels.json, part "locker"), which every
// build carries, and a language pack words them. Their keys are added from
// here (locker-keys.js), so a build made without the Locker's switch carries
// none of them and does not count its labels (main.js codeKeys).
import { addKeys } from '@rutba/office-ui';
import { LOCKER_KEYS } from './locker-keys.js';

addKeys(LOCKER_KEYS);

export { t, tn } from '@rutba/office-ui';
