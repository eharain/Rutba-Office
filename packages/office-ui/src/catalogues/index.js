// The languages the windows carry built in besides English: none since
// D-2026-11-14-16. Urdu (from 1.38.0), Arabic (from 1.40.0) and Hindi (from
// 1.42.0) were catalogues here; their words are their language packs' now
// (lang-ur, lang-ar, lang-hi), as every other language's are, and English is
// the lang-en copy Office carries (resources/languages).
//
// The add-ons contract's languages.builtIn names these, and its leftCore the
// languages that left (tests/language-built-in.test.js holds the two to each
// other); the menu (renderer/shell.js), the first start and the update gate
// read them.

/** A language's built-in table, by its tag: none. */
export const CATALOGUES = {};
