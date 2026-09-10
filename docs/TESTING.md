# How Rutba Office is tested, and what each test can and cannot tell you

A green build proves that the code compiles. Nothing below is satisfied by
that. Each layer here exists because the one above it missed something a
person found first; the layers are listed with what they prove and, as
importantly, what they do not.

## The layers

| Layer | Command | What it drives | What it proves | What it cannot see |
| :-- | :-- | :-- | :-- | :-- |
| Engine tests | `npm test` | The packages, in node | A format reads and writes as specified; an edit changes only what it should; a formula computes what Excel computes | Anything about a window: layout, paint, timing, a click |
| Editing checks | `npm run verify:edit` | Two real windows, real key events | Typing and formatting reach the engine through the DOM, and undo puts them back | Everything outside those two windows |
| Application checks | `npm run verify:apps` | Every app's window, the ribbon's buttons, the menus, the dialogs | A control does what its label says, on a fixture, and the engine agrees | Real files: the fixtures are small and well formed |
| Smoke | `npm run smoke` | Every window, photographed | Every app paints without an exception on first render | Whether what it painted is right |
| **Corpus** | `npm run verify:corpus` | Every file in a folder, opened in the real application one window at a time | The file is shown, refused with a reason, hangs, or crashes — and how long it takes | Whether what is shown matches Word, Excel or PowerPoint |
| Fidelity comparison | `tools/compare-word.ps1`, `compare-excel.ps1`, `pdf-pages.ps1` + a capture | A file in Microsoft's application and in ours, side by side | What differs on the page | It is a person looking, not a check |

`npm run gate` runs the first four before a release. The corpus run is
separate because it takes an hour over a thousand files; it is run before a
release and its report is read, and the failures it names are fixed or
listed in the release notes as known.

## What the corpus run reports

For each file: `ok` with the time to first paint, `refused` with the
application's own reason, `timeout`, `hung` (the renderer stopped answering),
`crashed` (the renderer went away), or `skipped` (no app claims the
extension). Console errors the window logged are kept with the file even when
it opened. The report lists the failures, then files that opened but logged
an error, then the slowest twenty, then the numbers per format with median,
90th-percentile and maximum open times.

Point it at any folder:

```bash
RUTBA_CORPUS_DIRS="D:\documents;E:\more" npm run verify:corpus
```

It opens each file in a profile of its own, so a run over a thousand files
does not write a thousand entries into your recent list. It keeps its windows
off the desktop — to the right of every display, out of the taskbar, never
activated — so they are not in your way and you cannot close them by
mistake, and they never take the keyboard from you. Set
`RUTBA_WINDOW_DISPLAY=secondary` to watch a run on a second display
instead. `RUTBA_SMOKE_THEME=dark npm run smoke` photographs every window
in the dark theme, which is how a change to the look is checked in both.
`RUTBA_VERIFY_ONLY=pages,grips,panes npm run verify:apps` runs only the
named blocks of application checks — the ones for a feature being worked
on, in seconds rather than minutes — and `RUTBA_VERIFY_CAPTURE=<dir>`
makes those blocks photograph the window at the moment that matters (a
page boundary in Rutba Word, each of Presentation's panes).

## When the main process is late

The application checks print a line whenever the main process kept a timer
waiting and a clock line every minute, so a slow block shows itself. Two
switches say why: `RUTBA_VERIFY_PROFILE=1` samples the main process while
the checks run and prints the top functions by own time, the same by file
and the idle share — a low idle share means our own code, a high one means
the machine; `RUTBA_IPC_TRACE=1` (any run of the application) prints the
busiest bridge channels and the windows behind them every five seconds.
Together they found an idle Worksheets window asking for a wider frame
eight times a second, which a day of reading the code had not.

## What "opens" means, per app

The run waits for the element each app renders only once it is showing the
file, not for the page to load — and so does every window the application
checks open (a fixed pause stood there before, and a loaded machine beat it):

| App | The element | Refusal |
| :-- | :-- | :-- |
| Rutba Word | a paragraph block on the page | "This file could not be opened", with the engine's reason |
| Worksheets | the cell layer | same |
| Presentation | a slide thumbnail or the slide | same |
| Pictures | the picture decoded (`naturalWidth`), the PDF frame, or the media element with its metadata | same |
| Image | the canvas | same |
| Video | a clip on the timeline | same |

An element is not proof: an `<img>` whose file never decoded is still an
`<img>`, and counting it as shown is how a zero-byte PNG passed the run. The
picture apps are waited on for a decoded image or a media element that has
read its metadata, and each of them refuses in words when the decoder does.

## The cases the checks are being broadened to cover

Written 2026-09-08 after the owner used the suite and found problems opening
documents and with usability that none of the checks had caught. In order:

1. **Every real file, in the real application** — the corpus run above.
   Built. The first run over the research corpus (475 `.docx`, 445 `.pdf`,
   247 `.txt`, 98 `.doc`, 96 `.xlsx`, 44 `.pptx`, and the rest) is the
   baseline; every failure it names becomes a fixture and a check.
2. **Time budgets.** A file that opens in twelve seconds has not "opened" in
   any sense a person would accept. The corpus report's per-format
   percentiles set the budgets: the check fails when the median for a format
   regresses past its budget, not only when a file fails.
3. **Round trips for every format.** Open, edit one thing, save, reopen, and
   compare: the edit is there and nothing else changed. Today this exists for
   `.docx`, `.xlsx` and `.pptx` on fixtures; it is extended to every format
   the suite claims to write, on corpus files.
4. **The first five minutes.** New document, type a paragraph, format it,
   save with a name, close, reopen from the launcher's recent list. For every
   app, driven through the real windows, keyboard only and mouse only.
5. **Error paths.** A corrupt file, a file that is locked by another program,
   a file that vanishes while open, a save to a read-only folder, a picture
   that is not a picture. Each must produce a sentence a person can act on,
   never a spinner and never a blank page.
6. **Large and hostile files.** A 200-page contract, a workbook of two
   million cells, a deck of fifteen megabytes, a picture of fifty megapixels,
   an hour of video: opened, scrolled to the end, edited at the end, saved,
   within a budget.
7. **Usability journeys.** Every menu item and every ribbon control reached by
   keyboard; every dialog closed by Escape with nothing changed; window
   resize down to the minimum with nothing clipped; the ribbon collapsed and
   expanded; zoom in every app; the recent list, drag-and-drop onto a window,
   and double-click from the operating system.
8. **Fidelity against Microsoft**, measured, for a fixed set of corpus files
   per release, so a regression on a page is caught by a picture rather than
   by a person.

Items 1 and 2 exist as of this writing, and the first parts of 5 and 7: a
broken file in each document app must produce the sentence, and the
journeys block covers Escape on dialogs, the minimum window size, Ctrl+F1
and the launcher's recent list. Items 3, 4, 6, 8 and the rest of 5 and 7
are the work.

Item 3 and most of item 5 arrived with `tests/document-service.test.js`,
which drives the service every window uses — open, convert, save, export,
refuse — in node, where a check costs milliseconds rather than a window. It
covers the round trip for every format the suite writes, a file that is not
the kind the window asked for, and what the disk failing says. What it
cannot see is the window: that a refusal is drawn, and drawn instead of a
spinner, is still the corpus run's job.

## What the first corpus runs taught, 2026-09-08

- A window never closed its document session; a hundred and forty files in,
  the main process stalled for two minutes. Sessions belong to their window
  now. The harness prints the open-session count and the main process's
  memory per file so the next leak shows as a number that only grows.
- The Presentation window's error return sat above a hook, so a refused
  file was a blank window. Every early return in a component goes below its
  last hook; the broken-file check guards it.
- A launcher in a check run made the announcement request, and an aborted
  one threw a dialog onto the owner's screen. Check runs switch the request
  off; the main process logs uncaught errors to `errors.log` in the profile
  instead of showing a box.
- Closing a check run's launcher by hand ends the run: the run keeps one
  window open on purpose, and that window must not be closed. Two runs were
  ended by hand this way — once on the laptop's screen, once on the second
  display — so the windows now sit off the desktop, where nobody can.
- A 46 MB workbook of eighteen million plain values took 135 seconds and
  2.7 GB and gave up: the calculation model copied every cell of a sheet
  with no formula, and the view kept a map of every cell's style, which a
  JavaScript map cannot hold past 16.7 million entries. Past a size, a
  formula-free sheet stays in its part and styles are read on demand. Eight
  seconds now. A million-row sheet still spends most of that parsing rows;
  a byte-indexed part would be the next step.
- The owner's "focus was not moving with my keys" was the formula bar
  keeping focus after Enter. The harness clicked elements from inside the
  page and never saw it; the real-input block sends the mouse and the
  keyboard through the window and reads the engine. A click on an empty
  cell must be aimed by the headers, because an empty cell has no element.



## What a corpus of hostile files taught, 2026-09-08

The runs above were over real documents. This one was over 76 files written
on purpose to be awkward: every format the suite claims, in ordinary shapes
and then zero-byte, truncated at the header, truncated in the middle, with a
byte flipped, under the wrong extension, in four text encodings, with the
name of an Office owner file, and with names full of spaces, brackets and
scripts other than Latin. Four things came out of it that a thousand real
documents had not.

- **A document under a workbook's name hung the window.** The extension
  decides which app opens a file, so `report.xlsx` that is really a document
  opened Worksheets, which drew a document model and threw on geometry it
  does not have: thirty seconds of blank window, then the timeout. The window
  now says which kind it edits when it asks for a file, and a file of another
  kind is refused by name — "docx-named.xlsx is a document, not a workbook.
  Open it in Word." — in a third of a second.
- **A UTF-16 text file was refused as an MP3.** `FF FE` is the byte-order
  mark Notepad writes for "Unicode" and PowerShell writes whenever it
  redirects; it is also a valid MPEG frame sync, and the audio test came
  first. Every UTF-16 text file on Windows was unopenable. The mark is read
  before the magic numbers now, and `tests/sniff.test.js` exists because the
  code that decides what a file is had no test of its own.
- **A file the decoder refuses left the media window empty for ever.** The
  Video window builds its timeline on `loadedmetadata` and listened for
  nothing else, so a file with a RIFF header and rubbish behind it produced
  no timeline, no message and no end. Pictures and the Image editor had the
  same shape in a milder form: a toast that fades, and then a window that
  looks like nothing was opened. All three now say what happened, and keep
  saying it.
- **A picture that never decoded counted as shown.** An `<img>` is an `<img>`
  whether or not there is an image behind it, so the corpus run called a
  zero-byte PNG "ok" in 672 ms. It waits for `naturalWidth` now, and for a
  media element's metadata, which turned three false passes into three
  refusals with sentences.

And one that was not about files: **the disk's own failures reached people as
error codes.** Saving over a file Word still had open said "EPERM: operation
not permitted, open 'D:\work\report.docx'". A save into a folder that had
been moved said ENOENT. Each of them now says what happened and what to do
about it, and `tests/document-service.test.js` holds them to it.

## Before a release

The gate proves the code. Three more runs prove the build, in this order,
and each has caught something the gate cannot:

1. **The packaged smoke, from outside the repository.** `npm run dist`, then
   copy `apps/desktop/release/win-unpacked` somewhere that is not under the
   repository and run it there with `RUTBA_OFFICE_SMOKE=1 RUTBA_SMOKE_SEED=1
   RUTBA_WINDOW_DISPLAY=offscreen` and a `--user-data-dir` of its own. Run
   in place, the unpacked build finds workspace packages the installer never
   packed — Node climbs from `release/win-unpacked` to the repository's own
   `node_modules` — and passes for a build that fails once installed.
   `tests/desktop-deps.test.js` reads the imports for the same reason.
2. **The install test.** `tools/install-test.ps1 -Installer <the setup exe>
   -Label <version>` installs silently, checks the registered copy, the
   icon behind every file type, the tiles unpacked beside the archive, and
   that the installed copy opens a Rutba Word window from `--app=word` in a
   profile of its own, off the desktop. When it does not, the script prints
   what the copy showed instead — the text of an "Error" box included — and
   exits 3. It refuses to install over a running copy (exit 2).
3. **Publish**, with `node tools/publish-release.js`, a pre-release by
   default; the tag goes on the commit the notes and the version bump are
   in, after the two runs above, never before.

1.8.0 is why the first two are written down: the gate was green, the
unpacked smoke was green, and the installed copy opened with "Cannot find
package '@rutba/contacts'" because the two new packages were not declared
as the desktop application's dependencies. The install test said so in one
word — a window titled "Error" — and the word was nearly missed.

## The rich fixtures, written by Office itself

`tests/fixtures/rich/` holds a workbook, a document and a deck that use what
people actually put in files, written by Excel, Word and PowerPoint through
COM automation — `tools/make-rich-fixtures.ps1` — as OOXML and as
OpenDocument from the same content, so the OpenDocument readers are judged
against the same thing the OOXML readers are.

- **showcase.xlsx / .ods** — five sheets: twelve months by four regions with
  totals, averages, IFERROR growth, number formats, borders, fills, a colour
  scale, data bars, an icon set, sparklines, an autofilter, frozen headings,
  a note, a hyperlink and a validation list; a Summary sheet whose every
  cell reaches into another sheet — `SUM`, a named range, `INDEX/MATCH`,
  `VLOOKUP`, `COUNTIF`, `SUMIF`, `IF`, `TEXT`, a sheet whose name has a
  space, a table — and a pie chart; a Charts sheet with four chart kinds,
  sixteen preset shapes with fills, gradients, outlines and shadows, a
  picture, a text box, a group, a rotation and a connector; a Data types
  sheet with one of everything a cell can hold; and a real table with a
  style, a total row and a structured reference.
- **showcase.docx** — a title, a table of contents Word built, headings,
  every run format, a justified shaded paragraph, three kinds of list, a
  styled table with a merged total, a picture with a caption, five floating
  shapes and a text box, a footnote, an endnote, a comment, a bookmark, a
  hyperlink, DATE/PAGE/NUMPAGES fields, a tracked insertion and deletion, an
  equation, a landscape section in two columns, a header, page numbers and
  a DRAFT watermark. Taken from Word as Flat OPC (`Document.WordOpenXML`)
  and packed by `tools/flat-opc-to-docx.mjs`, because Word's SaveAs never
  returns from an automated session on the build machine — so there is no
  `.odt` from Word; the corpus carries OpenDocument text from other hands.
- **showcase.pptx / .odp** — eight slides: a gradient title slide with
  notes, three levels of bullets and a picture, a gallery of sixteen preset
  shapes with fills, gradients, transparency, shadows, a group, a rotation
  and a hyperlink, a styled table, two charts, WordArt and a formatted text
  box, animations on a hidden slide, sections, footers and slide numbers.

`tests/rich-fixtures.test.js` opens every one through the document service
the windows use and holds the engine to Excel's own results — the totals,
the names, the formats, the chart count, the merged cells, the frozen panes
— and to what the document and the deck carry. Regenerate with

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File tools/make-rich-fixtures.ps1
```

on a machine with Office; the script drives the applications invisibly
(PowerPoint shows itself minimised for its charts, which a windowless
presentation refuses), reaps the processes it started, and reports each
decoration Office declined rather than stopping.

What the first generation found, on 2026-09-10: `SUBTOTAL` was missing, so
every Excel table's total row read `#NAME?`; an `.ods` lost its named
ranges (`$$Name` in the file), merged cells and number formats; built-in
date format 14 was rendered in the American order rather than the
machine's; and a chart on a slide was an empty frame. All four are fixed
and pinned. Still open, in docs/GAPS.md: OpenDocument drawings, equations,
and tracked changes shown as changes.

### The downloaded corpus

`.corpus/samples/` (ignored by git) holds 81 real files from two public
sample sites — getsamplefiles.com and mzeeshan.me — in every format the
suite opens, including corrupted, password-protected, non-Latin, equation,
chart, 73-page and 36 MB media cases. Run them through the real windows
with

```bash
RUTBA_CORPUS_DIRS="D:\Rutba2.0\office\.corpus\samples" npm run verify:corpus
```

and turn each failure into a fixture and a check, as the section above
describes.
