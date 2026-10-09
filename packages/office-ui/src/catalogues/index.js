// The catalogues of the languages the windows can be shown in besides
// English.
//
// messages.json beside this file lists every message the windows show, in
// English, and which files use it — made by tools/extract-messages.mjs, and
// checked by the tests to be current. A language is a module here named by
// its tag — fr.js, ar.js, pt-br.js — whose default export maps each English
// message to its translation, `{name}` placeholders kept as they are:
//
//   export default {
//     'Close': 'Fermer',
//     'Save changes to {name}?': 'Enregistrer les modifications de {name} ?',
//     '{count} word': { one: '{count} mot', other: '{count} mots' },
//   };
//
// A counted message is keyed by its English singular, with the forms the
// language's plural rules name (zero, one, two, few, many, other). A
// message left out shows in English. Import the module below and list it
// under its tag; the window picks it from the system's language, or from
// the "language" setting when there is one.

import ur from './ur.js';
import ar from './ar.js';

/**
 * Urdu, since 1.38.0, and Arabic, since 1.40.0: each the whole list, its
 * frame right to left (renderer/language.js).
 */
export const CATALOGUES = { ur, ar };
