# The roadmap, reviewed — and the order from here

A review of what landed between 1.26.0 (24 September 2026) and 1.29.10
(8 October 2026) against the build order in [GAPS.md](GAPS.md), and the
order the work should take next. GAPS.md stays the register of what is and
is not built; this file says what to build, and why in that order.

## What landed, against the order

The ten-item order at the foot of GAPS.md is complete. Every item is struck
through as of 1.29.4: floating layout in the deck, shape formatting and text
properties in the deck writer, fields through to a table of contents that
refreshes, recording tracked changes, and `LAMBDA`. The "absent, not
incomplete" table went the same way — equations, password-protected files,
a spelling pass, an accessibility check and right-to-left layout are all
built; localisation is started.

Thirteen releases went out in that fortnight, 316 commits by two people,
each a pre-release with installers on GitHub. The engine suite grew from
893 to 1,737 tests passing on this Linux head (the README says 1,807; see
below), the application checks from 434 to 1,174.

Three kinds of work made up the fortnight, and all three fit the suite's
own rule — a button either works or is not there:

- **The order itself**, finished in the first days (1.27.0, 1.28.0).
- **Greyed-out buttons made to work**: sixteen in 1.29.1, seventeen in
  1.29.2, sixty in 1.29.3 (the Draw tab, media and narration, Compare,
  SmartArt), tables that do what Word's do in 1.29.6, Excel-like selection
  and the clipboard in 1.29.7.
- **Reach**: every age of Word, Excel and PowerPoint file in 1.29.5; Power
  Query, a thesaurus and CalDAV/CardDAV in 1.29.8; motion paths, 3D models
  and Cameo in 1.29.9; scripts in 1.29.10.

One line was crossed on purpose and crossed well. GAPS.md keeps "macros
that run" below the line, and still does: scripts are the user's own, kept
on this computer, never carried in a file. The worker they run in is served
under a policy of `default-src 'none'` with only `script-src 'self'
'unsafe-eval'` (packages/office-shell/src/electron/protocol.js), so it can
reach no network, no file and no frame; it is handed a snapshot and hands
back edits applied as one undo step, and is stopped at fifteen seconds. That
is the right shape, and it is the shape anything else that runs user code
must take.

## Where it drifted

Nothing landed off the roadmap. Five things slipped around it:

1. **Features under patch numbers.** 1.29.1 to 1.29.10 are feature releases
   — Power Query, scripts, old file formats — and a user, or the updater,
   cannot tell a fix from a feature by the number. Features take the minor
   number; a patch release carries fixes only.
2. **The gate is not on the record.** No release commit or release note since
   1.26.0 says the Windows gate ran, or what it counted. The notes'
   "Proven, not promised" sections prove each feature; nothing proves the
   release.
3. **The download page is thirteen releases behind.** The site in
   Rutba-Management still says 1.26.0 and 893 tests (last changed 24
   September). Someone arriving at the site is told a version that is
   two weeks old.
4. **Two engine counts disagree.** The README says 1,807; this head runs
   1,737 here with none skipped. Either seventy tests exist only on Windows
   or the README is written ahead of the suite. A count written by hand
   drifts; the gate should write it.
5. **Localisation has started with no policy.** A message catalogue is a
   change to every string in the suite. Before it spreads, it needs the
   catalogue's format, the languages that ship, who translates them, a
   right-to-left interface (not only right-to-left paragraphs), and a
   check that no English string is left outside the lookup.

## The order from here

Depth before more breadth. The surface is wide now; what a person meets
next is a file that opens slowly, a count that is wrong, or a page that is
behind.

1. **1.30.0 as a consolidation release.** [REVIEW-1.29.md](REVIEW-1.29.md)
   holds what a reading of the code found, fixed and left open — the
   500-page pagination cap, the document service on the main process and
   the shared Dialog's focus are its three high items, and belong here. The gate run on Windows and its
   counts written into the release note and the README by the gate script
   itself; `tools/fuzz-open.js` run over the binary readers (Word 1.0 to
   2003, Write, `.xls`, `.ppt`, encrypted CFB) with its iteration count
   recorded, since those readers take hostile input and must refuse in a
   sentence; the site brought to the current release (`npm run
   import:releases`, `check:facts`) as a step of every release, not a
   separate job; the version policy above, written into docs/RELEASING.md
   or wherever the release steps live.
2. **A performance budget, as checks.** No gap entry mentions size. Open and
   scroll a 300-page document, a 100,000-row workbook and a 200-slide deck
   within a stated time and memory, as application checks that fail when
   the budget is missed. Fix what they find before anything new.
3. **Localisation to one shipped language, end to end.** Urdu first, given
   the right-to-left work already done: every string through the catalogue,
   the interface mirrored, dates and numbers formatted, and a check that
   finds any string left in English. A second language only after the
   first is complete.
4. **The remainders.** Page setup kept in the file for a document and a deck
   (order item 2 was built for workbooks only); a real watermark and
   multilevel list definitions; a default template; edits from the Pictures
   viewer that change the file; thumbnails on Linux; **WordArt's Transform
   presets** — words drawn along an arc, a circle, a button and the waves,
   on screen and in print, offered from Text Effects → Transform and written
   as `a:prstTxWarp` in all three apps (asked for 2026-10-08; the arc and the
   circle first, since those are what people reach for).
5. **Deepen what 1.29 opened** before widening it: Power Query's remaining
   transforms and a refresh that re-reads its source; scripts' API to cover
   what Record Actions can record; the old-format readers against a corpus
   (`npm run verify:corpus` over a folder of real files of each age).
6. **The field's gaps and the two editors.** [COMPETITORS.md](COMPETITORS.md)
   reviews the suite against WPS Office, Microsoft 365, Google Workspace,
   LibreOffice, ONLYOFFICE, SoftMaker and Zoho, and sets out what to add:
   a PDF app, OCR and a recorder on this computer, templates, forms, and
   Image and Video taken in three tiers each from crop-and-trim to layers,
   tracks, captions and MP4. Its release-by-release order (1.31 to 1.33)
   follows items 1 to 5 above and runs beside localisation. The two
   editors and the recorder now have their own plan and a dedicated worker:
   [STUDIO.md](STUDIO.md), built on `packages/studio`, the renderer brought
   over from the consumer suite's Studio on 2026-10-09.
7. **Then new breadth**, each item added to GAPS.md first, built second.
8. **More languages after Urdu: Arabic, Hindi, Chinese and Japanese.**
   1.38 and 1.39 shipped the suite in Urdu. The next languages, their
   order and what each needs beyond a catalogue are set out under
   [Languages after Urdu](#languages-after-urdu) below. They run beside the
   other items, one language at a time, each finished before the next.

Everything below the line in GAPS.md stays there: macros that run, cloud
services, anything that would make a request this suite has not declared.

## Languages after Urdu

1.38 and 1.39 put every word the windows show through the message catalogue
(`packages/office-ui/src/messages.js`) and shipped Urdu
(`packages/office-ui/src/catalogues/ur.js`, with `GLOSSARY-ur.md`), the frame
mirrored right to left. The next languages are **Arabic, Hindi, Simplified
Chinese, Japanese and Traditional Chinese**, in that order. A catalogue is
the smaller part of each: the documents themselves have to be typed, laid
out, printed and saved correctly in the script, and that is where most of
the work below is.

### What every language needs

- **A catalogue and a glossary**: `catalogues/<tag>.js` and
  `GLOSSARY-<tag>.md`, following the words Microsoft Office and Windows
  already use in that language. Drafted by machine against the glossary,
  marked as awaiting review in the release note, and reviewed by a native
  speaker before the note may drop that line.
- **Its entry in the Language menu** (`LANGUAGES` in
  `apps/desktop/renderer/shell.js`), named in its own script.
- **Plural forms by its own rules.** `tn()` already chooses the form with
  `Intl.PluralRules` and reads `{ zero, one, two, few, many, other }` from the
  catalogue. Arabic uses all six; Hindi `one` and `other`; Chinese and
  Japanese `other` only.
- **The checks generalised.** `tests/catalogue-ur.test.js` becomes one test
  per catalogue: every message has an entry, every plural form the
  language's rules can select is given, and no Latin word is left in. The
  window checks that open Documents, Worksheets and Presentations in Urdu
  run in each language.
- **Engine and service messages** (`renderer/engine-words.js`) translated
  as 1.39 did for Urdu.
- **Dates, times and numbers** through `Intl` in the window's language,
  including the calendar's first day of the week.
- **The installer** in each language (it is English only today, Urdu
  included).
- **Fonts.** Windows ships a font for each of these scripts (Segoe UI and
  Arabic Typesetting for Arabic, Nirmala UI for Devanagari, Microsoft YaHei
  and Microsoft JhengHei for Chinese, Yu Gothic and Meiryo for Japanese).
  The interface names them in its font stack; documents keep the fonts
  their files name and fall back to these.
- **PDF export** keeps the script: `packages/pdf` embeds a subset of a
  TrueType font (`truetype.js`) and joins Arabic itself (`shaping.js`).
  Each new script states what it adds there (below).

### Arabic (first)

Most of it is done by Urdu's work: the mirrored frame, right-to-left
paragraphs, `w:bidi` on runs, bidirectional reordering and Arabic joining in
PDF.

- Catalogue in Modern Standard Arabic, with a glossary.
- Six plural forms in every counted message.
- Digits: Western digits by default, Arabic-Indic digits (٠١٢٣) as a choice
  in settings, applied to the interface and offered as a number format in
  Worksheets.
- The Hijri calendar as a date format in Worksheets and as a second calendar
  in Calendar (`Intl` with `-u-ca-islamic-umalqura`).
- Find and Replace that can ignore diacritics (harakat) and treat the forms
  of alef and of yeh as one, as Word does for Arabic.
- Kashida justification in Documents, as Word draws justified Arabic.
- Spelling: a Hunspell dictionary for Arabic in `@rutba/proofing`, where
  its licence allows bundling.

### Hindi

Devanagari is a shaped script: consonants join into conjuncts and vowel
signs sit before, after, above or below the letter they follow.

- Catalogue and glossary in standard Hindi.
- **Caret and selection by grapheme cluster** in Documents, Worksheets'
  cell editor and Presentations: moving, deleting and selecting a whole
  syllable (`Intl.Segmenter` with `granularity: 'grapheme'`), never half of
  a conjunct.
- Line breaking at word boundaries that keeps a cluster whole, in the page
  layout and in PDF.
- **PDF shaping for Devanagari**: the writer embeds glyphs by code point
  today, which is right for Latin, Greek and Cyrillic and handled by
  `shaping.js` for Arabic. Devanagari needs the font's OpenType substitution
  and positioning (GSUB and GPOS) applied, for conjuncts, the reordered
  short i and the marks. Write it in `packages/pdf` for Devanagari first;
  it is the same work Bengali, Tamil and the other Indic scripts will need.
- Indian digit grouping (1,00,000 and 1,00,00,000) as the default number
  format in Worksheets for `hi-IN`, from `Intl.NumberFormat`.
- Devanagari digits (०१२३) as a choice, as Arabic's are.
- Spelling: a Hunspell dictionary for Hindi where the licence allows.

### Chinese (Simplified, then Traditional)

- Two catalogues: `zh-Hans` (Simplified, for China and Singapore) and
  `zh-Hant` (Traditional, for Taiwan and Hong Kong), each with its own
  glossary, since the Office terms differ between them as well as the
  characters.
- **Typing through an input method.** Composition events
  (`compositionstart`, `compositionupdate`, `compositionend`) handled in
  every editing surface: the text being composed shown at the caret,
  underlined, and only committed text recorded as an edit, undone as one
  step. Documents, Worksheets' cells and formula bar, Presentations' text
  boxes, Mail's editor and every dialog field.
- **Line breaking without spaces**: a line may break between most
  characters, except that some punctuation may not start a line and some
  may not end one (the rules called kinsoku in Japanese). The page layout,
  text boxes, cells that wrap and PDF use `Intl.Segmenter` and a table of
  those characters.
- **East Asian fonts in the files**: a run's `w:eastAsia` font and language
  (`packages/ooxml/src/runs.js` already writes the language) read and
  written, and the document's East Asian font used for those characters.
- Full-width and half-width characters measured correctly in the grid and on
  the page.
- **PDF**: CJK fonts on Windows are TrueType collections (`.ttc`) or
  OpenType with CFF outlines. `truetype.js` reads single TrueType files, so
  it needs collections and CFF before a Chinese document exports with its
  text. A subset is essential; the fonts are tens of megabytes.
- Word count as Word counts Chinese: each character is a word.
- Sorting by pinyin (Simplified) and by stroke (Traditional) through
  `Intl.Collator`, in Worksheets' Sort and in Contacts.
- No spelling check; the proofing pane says so rather than marking every
  word.

### Japanese

Everything listed for Chinese applies: the input method, line breaking with
kinsoku, East Asian fonts in files and in PDF, full-width measurement, the
character word count and no spelling check. Added for Japanese:

- Catalogue and glossary in Japanese.
- The Japanese era calendar as a date format in Worksheets and in Calendar
  (`ja-JP-u-ca-japanese`).
- Sorting by reading through `Intl.Collator`; Contacts keeps a phonetic name
  (furigana) beside each name, as vCard's `X-PHONETIC` fields and Outlook do,
  and sorts by it.
- Find and Replace that can treat hiragana and katakana, and full-width and
  half-width forms, as the same.
- Later: phonetic guides (ruby, `w:ruby` in Word files) and vertical text in
  Documents and Presentations.

### Order and proof

| Release | Language | Done when |
|---|---|---|
| Next | Arabic | Catalogue whole, six plural forms, window checks in Arabic pass, an Arabic document typed, saved, reopened and exported to PDF with its words joined and in order |
| Then | Hindi | Caret and deletion by grapheme cluster, and a Hindi PDF whose conjuncts match the screen, each a check |
| Then | Simplified Chinese | Input method composition in every editing surface, line breaking by kinsoku rules, a Chinese PDF with its text, each a check |
| Then | Japanese | As Chinese, with the era calendar and reading sort |
| Then | Traditional Chinese | The second Chinese catalogue and stroke sort |

Each language ships as a beta catalogue first, with the release note saying
it awaits a native speaker's review, as Urdu's did. Studio's editors
([STUDIO.md](STUDIO.md)) follow the same rules: titles and captions in each
of these scripts are part of their checks.
