# @rutba/proofing

Review's tools for Rutba Word, Worksheets and Presentation.

- **Check Accessibility** — `accessibility.js` holds the rules (Office's
  Errors / Warnings / Tips and its rule names); `word.js`, `sheet.js` and
  `deck.js` describe each kind of document for them and write the fixes:
  alternative text on `wp:docPr`, `xdr:cNvPr` and `p:cNvPr` (`descr`, and
  "Mark as decorative" as Office's `adec:decorative` extension), a header
  row, a slide title, a readable text colour, a document title.
- **Spelling** — `tokenize.js` finds the words (skipping UPPERCASE, words
  with numbers, Internet and file addresses, as Office does by default),
  `speller.js` and `speller-host.js` hold the dictionaries on a worker
  thread, `spelling.js` walks a document's text from the caret, and the
  three app modules list the text and change it where each word is.

The main process uses it through `apps/desktop/main/proofing.js`; the panes
and dialogs are `@rutba/office-ui/proofing`.

## Third-party components and their licences

The dictionaries are read by **nspell** 2.1.5 (MIT, © 2016 Titus Wormer),
a JavaScript implementation of Hunspell.

Two dictionaries ship, each from npm with its licence file beside it in the
installed package:

- **dictionary-en** 4.0.0 — English (United States)
- **dictionary-en-gb** 3.0.0 — English (United Kingdom, "-ise" spelling)

Their package code is MIT (© Titus Wormer). The dictionary and affix files
are the Hunspell dictionaries made from **SCOWL** (Spell Checker Oriented
Word Lists, http://wordlist.sourceforge.net, version 2020.12.07), which is:

> Copyright 2000-2018 by Kevin Atkinson
>
> Permission to use, copy, modify, distribute and sell these word lists,
> the associated scripts, the output created from the scripts, and its
> documentation for any purpose is hereby granted without fee, provided
> that the above copyright notice appears in all copies and that both that
> copyright notice and this permission notice appear in supporting
> documentation. Kevin Atkinson makes no representations about the
> suitability of this array for any purpose. It is provided "as is"
> without express or implied warranty.

SCOWL itself draws on further word lists under similarly permissive terms —
the Ispell word lists (© 1993 Geoff Kuenning, BSD-style), the WordNet
inflections (© 1997 Princeton University), the UK English lists (©
1993-1999 J Ross Beresford, © 2000-2016 Kevin Atkinson, © 2016 Benjamin
Titze) and public-domain sources. Each notice is reproduced in full in the
`license` file of `node_modules/dictionary-en` and
`node_modules/dictionary-en-gb`, which travel with the application.
