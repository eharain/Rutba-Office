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

1. **1.30.0 as a consolidation release.** The gate run on Windows and its
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
   follows items 1 to 5 above and runs beside localisation.
7. **Then new breadth**, each item added to GAPS.md first, built second.

Everything below the line in GAPS.md stays there: macros that run, cloud
services, anything that would make a request this suite has not declared.
