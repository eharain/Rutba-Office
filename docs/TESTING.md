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
instead.

## What "opens" means, per app

The run waits for the element each app renders only once it is showing the
file, not for the page to load:

| App | The element | Refusal |
| :-- | :-- | :-- |
| Word | a paragraph block on the page | "This file could not be opened", with the engine's reason |
| Worksheets | the cell layer | same |
| Presentation | a slide thumbnail or the slide | same |
| Pictures | the picture, the PDF frame, the video or audio element | — |
| Image | the canvas | — |
| Video | a clip on the timeline | — |

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


