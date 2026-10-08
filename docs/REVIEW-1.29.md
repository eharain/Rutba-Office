# Review of 1.26.0 → 1.29.10: code, tests and gates

A reading of the 316 commits between the two tags, done 2026-10-08 on the
dev head, in three tracks — hostile input and sandboxes, feature
correctness, and renderer resources and performance — with tests written
for what was found and small fixes landed where the fix was certain. The
gates were run on Linux under Xvfb; the Windows gate remains the release's.
What is fixed is fixed on dev; what is open is listed here with its place,
for the consolidation release ROADMAP.md asks for.

## Gates on this head

- Engine suite: 1,737 tests under `node --test`, and 1,827 under `npm test`
  (which adds the PDF package's 70; six of those skip where pdfjs is not
  installed). All pass. The README's 1,807 is `npm test` before this
  review's additions, so the two counts were never at odds.
- Fuzz: `node tools/fuzz-open.js 150 4242` damaged 450 files across the
  formats; 158 opened, 292 were refused in a sentence, no faults.
- Editing checks 12 of 12; smoke 10 of 10 windows rendered clean; the full
  application pass 1,146 of 1,170 under Xvfb. Of the 24 that did not pass,
  eighteen want what the container has not got — an OS keystore (the
  CalDAV account chain of four and Mail's import), a camera and a
  microphone (the cameo's three, Record's four, Insert → Audio), a desktop
  to capture (Screenshot in two apps, Screen Recording and its saved deck),
  an Arabic font for the PDF text checks in Worksheets and Presentations,
  and a real mouse for the slicer's corner handle. Six are not explained by
  the platform and want a run on Windows before 1.30: the table of
  contents' entries drawn on the page (none found); Formula → Your own
  formula drawing =B2*2 (worked out, not drawn); Preview Results → Next
  record's address block; the Track Changes button reading pressed; the
  slicer's resize; and Summary Zoom's second target reading null. If any of
  them fails there too, it is a 1.30 fix; if all pass, each is a check that
  needs the time or the fonts this container lacks, and the check should
  say so as the `float` and `pages` blocks learned to.

## Fixed on dev by this review

Correctness, each with a test in `tests/review-*.test.js`:

- Reject All on a field Word had tracked as deleted left `w:delInstrText`
  behind, so the field lost its code (document.js, `_resolveParagraphChanges`).
- A new revision's id counted only some kinds, and only in the body; it
  could collide with a table, row, cell or section change, or one in a
  footnote, header or footer (`nextTrackChangeId`).
- Update all formulas worked a nested table's `=SUM(ABOVE)` against the
  outer table (`updateTableFormulas`, with `blankNestedTables`).
- Table Design edits ran over a `w:tblPrChange` record and stripped what
  rejecting it would restore (`_editTableProps`).
- Power Query's Refresh after the loaded sheet was renamed made a second
  sheet and a second table (`query-load.js`, `syncLoadSheet`).
- Append lost a column when two headings matched or were blank
  (`queries.js`, `appendQuery`).
- A deck hit after a dotted capital I cut the wrong slice: the search ran on
  a lower-cased copy whose offsets differed (`deck.js`, `find`).
- A `.doc` carrying a macro project converts to a `.docx` with none of it,
  no macro part and no macro-enabled content type (test only; the reader
  already went through a model and never copied the storage).

Hostile input, each held by `tests/binary-hostile.test.js`:

- A compound file whose directory entry named itself as its sibling
  recursed until the stack went; a looping sector chain with a 2 GB stream
  size allocated the lot; a DIFAT pointing at itself used 1.4 GB; a name
  over 64 characters and a DIFAT sector past a cut file's end each threw
  from a DataView. The tree keeps a visited set, chains and the DIFAT are
  bounded by the table and the file, a stream never allocates past the
  source, the name is clamped and a short sector ends the walk (`cfb.js`).
- Old Word's piece and PLC counts came from a 32-bit field and ran billions
  of times before "Invalid array length"; `entriesIn` clamps them to what
  the stream holds (`msdoc.js`, `msdoc-old.js`).
- A message that was nothing but unclosed tags took the preview's and the
  junk filter's tag patterns to the end of the text at every angle bracket,
  40 s for 200 KB with no click from the reader; a single pass replaces
  them (`mime.js`, `mail-junk.js`, `mail/parts.js`).
- An agile-encrypted file's spin count ran as many hashing rounds as it
  named, hours on the main process; refused above Office's own ten million
  (`crypt.js`).
- On Windows an attachment named `update.exe ` with a trailing space, or
  with an alternate data stream, missed the "this is a program" warning
  and ran; the extension is now taken as Windows takes it (`mail.js`,
  `attachmentExtension` in `mime.js`).
- A script could write cells past Excel's last row and column, which Excel
  then calls damaged, and a whole-sheet clear walked every address there
  is; both keep to the cells that exist (`script-apply.js`, `scripts.js`).

Performance, held by `tests/perf-budget.test.js`:

- A formula cell's row was found by searching the part's rows: 100,000
  rows recalculated in 20.4 s, now 7.0 (`recalc.js`, `writeCachedValue`).
- Every modified part was deflated at level nine, 7.8 s for a 34 MB sheet
  part against 0.76 s at level six for a file eight percent larger, on the
  main process where it froze every window (`zip.js`).
- Mark All Read, flag, Empty Folder and delete in Mail rewrote and
  re-parsed the folder index once per message, about 20 s for 2,000
  messages; `setFlagsMany` and `removeMany` write once (`store.js`, `mail.js`).
- A catalogue tagged `pt_BR` made `tn()` throw; tags are normalised
  (`messages.js`).

The budget as measured: a 300-page document opens, paginates, edits and
saves in well under a second; a 100,000-cell workbook recalculates and
saves in 5.2 s; a 200-slide deck renders every slide in about 0.1 s; a
3,000-message store marks all read in 0.23 s.

## Open findings, by severity

### High

- **Pagination stops at 500 pages.** `paginate.js` guards runaway layout
  with `maxPages = 500`; a 2,400-paragraph document lays out as 501 pages
  with the rest neither drawn nor editable, and nothing says so. The saved
  file is whole. Derive the guard from the content (say twice the flow's
  length, never under 500) and keep the runaway test on a tiny page.
- **The document service runs on the main process.** Open, edit, save,
  autosave, print and thumbnails for every document run synchronously in
  Electron's one main process (`documents.js`); a 100 MB workbook freezes
  every window, mail sync included, for tens of seconds, and the autosave
  back-off only spaces the freezes out. Move the service into a
  `utilityProcess`; until then, open through `OoxmlPackage.readAsync` and
  asynchronous reads, as `Deck.openAsync` already does.
- **The shared Dialog manages no focus.** About a hundred dialogs inherit
  `role=dialog` and `aria-modal` from `office-ui`'s Dialog but focus moves in
  only where an author wrote `autoFocus`, Tab walks into the page behind,
  focus is not restored on close, and one Escape closes every stacked
  dialog since the handler is on the window. Fifteen lines fix all four.

### Medium

- **A mail-merge source path stored in the document is opened on open**
  (`documents.js`, `reattachMergeSource`), with no prompt; a document
  naming a UNC share leaks Windows credentials to it and merges its rows.
  Refuse UNC paths and ask before attaching any path a document supplies.
- **`shell.openPath` and `fs.write` from the renderer check no extension
  or place**, and `rutba://file` serves any path; mitigated by the sandbox,
  context isolation and `script-src 'self'`, but `openPath` in main should
  refuse the extensions that run when opened.
- Tracked table rows (`w:ins`/`w:del` in `trPr`): Accept All leaves the
  marker, Reject All empties the row but keeps it; the resolver works
  paragraph by paragraph and needs a row level.
- Mail `move` and the junk sweep still call `store.remove` per message
  inside a loop; collect the ids and call `removeMany` once.
- `export-video.js` has no `try/finally`: a failed slide image or recorder
  start leaves the canvas capture track, the audio context and the recorder
  running; and every slide image is decoded up front (about 1.6 GB for 200
  slides at 1080p) with each picture inlined again per slide.
- `record.js`: a second press of Record before the microphone opens starts
  a second recorder and loses the first; closing while opening leaves the
  mic on; raw Float32 audio grows unbounded until close. `recordShow` in
  slides.js has the same re-entry.
- `screen-record.js` holds the whole recording as blobs and crosses IPC as
  one array, about 19 MB a minute.
- A deck picture's old blob is never released when its bytes change
  (`blobs.js`, `documents.js` `resolveImage`), and `hold` copies a Buffer it
  could keep as a view, so every picture is held twice.
- The mail store's folder-index cache is never evicted; the recent list
  stats up to sixty paths on every read, which a lost network drive turns
  into a hang; calendar and contacts rewrite one JSON file synchronously on
  every change, with photos as base64 sent on every list.
- Folder listing and thumbnails stat per entry synchronously; a 40 MP JPEG
  is decoded synchronously per tile on Linux; the thumbnail queue keeps
  abandoned jobs; the JPEG cache has no eviction.
- `whatsnew.js` and the About panel are hand-built scrims with no role,
  label, Escape or focus; the password reveal toggle is unreachable by
  keyboard.

### Low

- A Power Query source path stored in the workbook is read on Refresh with
  UNC paths accepted; limited to text files and to a click, but main never
  checks the path came from a dialog.
- The junk verdict trusts Safe Senders, contacts and Safe Recipients before
  the server's spam flag, on unauthenticated From and To.
- `msppt.js` `walkGroup` recurses on nested groups without a depth bound;
  ten thousand give a stack-overflow refusal, not a hang. Bound it at 64.
- Control characters in a cell string reach the sheet XML unfiltered; the
  writer has no XML 1.0 character filter.
- `tools/fuzz-open.js` holds only `.docx`, `.xlsx` and `.pptx`, so no binary
  reader is in its corpus, and its judge tests the message for a pattern
  that never matches, so a DataView range error or a stack overflow counts
  as a clean refusal. Judge on the error's name, add a two-second limit,
  and add `tests/fixtures/binary` to the corpus.
- One multi-piece tracked delete shares a `w:id` across its `w:del`
  elements; Word tolerates it, the schema does not.
- Accept All and Reject All resolve body paragraphs only, not headers,
  footers or footnotes.
- `updateTableFormulas` reads `w:fldSimple` only; Word's own `=` fields are
  complex fields and never recalculate.
- Merging cells while recording writes no `w:cellMerge`; the fill handle
  copies values over a merge but not the merge; `fillSeries` writes into a
  merge's hidden cells and, with step 0 and a stop, writes 10,000 cells.
- `splitColumn` and `parseDelimited` turn a quoted "007" into 7; query
  steps drop the `types` side table through append, merge and group.
- The legacy `findText`/`replaceText` pair disagrees with itself on matches
  across runs; the pane uses `find`/`replaceAll`, which agree.
- Drag handlers in sheets, word, slides and drawings remove their window
  listeners and auto-scroll timer only on mouseup; a lost mouseup leaves
  them running. Frame object URLs are never revoked; `posterFrame`'s timer
  is never cleared; `split.js` clones the page DOM on every change.
- `compare` and `photoAlbum` sessions with no window are never freed;
  `fs.write` writes in place rather than temp-and-rename; `capture.grab`
  renders every screen's thumbnail to pick one.
- Icon-only buttons in comments and sheet objects have no accessible name;
  the export progress is not a live region.

## Localisation, measured

The catalogue works: `t` by English key with `{name}` fills, `tn` through
`Intl.PluralRules`, a missing key falling back to English, 292 messages
extracted and a test that the catalogue is current. No language exists
yet. Of roughly 3,000 user-facing strings in the renderer, about 160 are
routed; `word.js`, `sheets.js`, `slides.js` and `mail.js` have none.
Localisation is five percent started, which is what ROADMAP.md item 3
should budget for.

## Sound

Checked and found right: the scripts worker's policy of `default-src 'none'`
blocks fetch, XHR, WebSocket, EventSource and beacons, `importScripts` and
nested workers reach only the app's own origin, a fresh worker runs each
script and is ended at fifteen seconds, prototype pollution stays inside
it and `applyScriptEdits` reads fixed fields; the IPC contract has lost
`fs.rename`, `fs.copy` and the secrets group, `openExternal` allows only
http, https, mailto and tel, and no `eval` or `new Function` exists outside
the worker; the mail frame is sandboxed under `default-src 'none'` with
remote pictures rewritten and off by default; the binary readers' byte
accessors return zero past the end and their record loops always advance;
the one new dependency group (`@rutba/proofing` with its dictionaries) is
pure JavaScript with no network. Also: Word's own revisions keep ids and authors through
an edit and a save; a Word table with `gridSpan`, `vMerge` and a nested
table saves byte for byte untouched; `transposeFormula` across ranges,
mixed anchors and sheets; cut-paste re-pointing and undoing whole; Power
Query refreshing twice without change and merging left, inner, anti and
full with case-blind keys; a match across two runs skipped consistently by
`find` and `replaceAll`; the old-format fixtures' styles, lists, tables
and floating drawings; the cameo stream's reference counting; every
ResizeObserver disconnected; the scripts worker terminated on every path;
blob release on window close; autosave back-off; the paginate line cache
bounded; the deck undo stack capped.
