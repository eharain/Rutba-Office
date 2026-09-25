# What is missing, between this and a suite somebody could use for everything

Written 2026-09-09, from the code rather than from memory. Three sources: the
suite's own register of unbuilt controls (257 of them, each a `<Soon>` in a
ribbon with a sentence saying what it needs), a reading of what the engines
write as against what they read, and the things a person expects that are not
in either list because nothing in the interface offers them at all.

The order below is not the order of difficulty. It is the order of what a
person notices: the first section is the work that stops somebody using this
instead of what they use now, whatever else is finished.

---

## 0. Absent, not incomplete

Nothing in the interface offers these. They are what separates an editor from
a suite.

| | State | Notes |
| :-- | :-- | :-- |
| **Print** | **Built, 2026-09-09** | Ctrl+P, a dialog that counts the pages, printers, copies, paper, orientation, margins; sheets add fit-to-width, print area, repeated rows, gridlines and headings; decks add slides, notes pages and handouts. PDF export goes through the same door for all three. |
| **Autosave and recovery** | **Built, 2026-09-09** | A dirty document is copied into the profile every half minute and the copy is deleted when it is saved or closed, so what is left at start-up is what a crash took. The launcher offers each one back by name, with where it belonged and when it was taken. |
| **Calendar** | **Built, 2026-09-09** | Month, week, day and agenda; calendars with a colour each; events with repeats, exceptions and zones; `.ics` opened, offered, imported and exported. An invitation in a message is shown in Mail's reading pane, kept as an event, and answered with the reply attached to a message to the organizer. |
| **Contacts** | **Built, 2026-09-09** | A book in the profile: list, card, editor, search. `.vcf` and CSV opened, offered and merged by address; vCards exported. Compose completes from the book and then from the people mail has seen; a sender is kept from the reading pane in one click. |
| **OpenDocument drawings** | **Missing** | An `.ods` written by Excel carries its charts as embedded objects and its shapes as `draw:custom-shape`; neither is read, so the Charts sheet of the showcase workbook opens empty from `.ods` and full from `.xlsx`. Its frozen panes (settings.xml) and hidden columns are not read either. An `.odp` reads only its text frames: the showcase deck's pictures, shapes, table and gradients are all dropped from `.odp` and all present from `.pptx`. Found 2026-09-10 by tools/make-rich-fixtures.ps1. |
| **Equations** | **Built for documents, 2026-09-25** | Office Math is read, drawn through the browser's own MathML, edited in Word's linear format, written as Word writes it and printed. A deck's equations (`a14:m`) are not read yet, though the same converter could draw them. |
| **Tracked changes, shown** | **Missing** | `w:ins` text arrives as plain text and `w:del` text is dropped — the accepted view, which is right for reading, but nothing says a change was made or by whom; there is no view of the markup and no accept or reject. Recording them is the older gap below. |
| **Page setup as a document property** | **Built for workbooks, 2026-09-09** | Paper, orientation, margins, scaling, gridlines, headings, centring, the print area and the repeated rows are read from and written to the file where Excel keeps them. A document and a deck still choose theirs in the dialog each time. |
| **Password-protected files** | **Missing** | Refused with a sentence that says so, since 2026-09-10 — the corpus found an encrypted `.docx` opening as an empty page and an encrypted `.pptx` refused as "a document". An encrypted OOXML package (`EncryptedPackage` in a compound file, agile encryption) is neither read nor written. |
| **A spell-check pass** | **Partly** | Chromium underlines as you type in the places that are contenteditable. There is no pass over a document, no dictionary the person owns, no add-to-dictionary, and nothing at all in the grid or on a slide. |
| **Accessibility check** | **Missing** | No alt-text prompt, no heading-order check, no contrast check. The windows are keyboard-reachable; nothing verifies it but a person. |
| **Right-to-left layout** | **Missing** | Arabic and Hebrew text draws, and paragraph direction, mirrored margins and RTL ribbons are not built. |
| **Localisation** | **Missing** | Every string in the interface is English, written inline. There is no message catalogue and no way to add one. |

---

## 1. Rutba Word — the document

Reads more than it writes. What it draws faithfully and cannot yet produce:

- **Floating layout.** On screen, since 2026-09-14, a picture or a text box
  anchored at the left or right floats there and the words wrap round it;
  Wrap Text and Position on the Layout tab move a picked picture between the
  line, the sides and the middle, and the file keeps it. Since 2026-09-20
  print and PDF do the same for a picture: the paginator lays the lines
  beside it shorter, for as far down as it reaches and into the paragraphs
  after it, and a centred or right-aligned block keeps its side, and since
  the same day a floating text box prints beside the words the same way.
  A picture behind or in front of the words is drawn in the flow. Align, Group, the selection
  pane and z-order for a document's drawings are still to come.
- **Columns and section control.** Columns arrived 2026-09-24: Layout →
  Columns writes `w:cols`, and print and PDF flow the words into them,
  each column filled to the foot of the page before the next starts, its
  own floats and footnotes at its foot and a separator drawn down the gap
  when asked for. On screen the editable flow is laid at the first
  column's width in print layout, as Word's own draft view does, and at
  the browser's own columns in draft layout. Line numbers, hyphenation
  and page borders are the same story still. (Pages themselves arrived on
  screen on 2026-09-10: the flow is laid onto sheets by measuring the
  drawn page, paragraphs split at
  a line and tables at a row, and a manual page break heads a fresh page.
  Since 2026-09-20 a footnote sits at the foot of the page its reference
  lands on, on screen as in print. A paragraph
  splits at a line, a table at a row and, since 2026-09-20, a paragraph of
  pictures between its picture lines — a paragraph of only pictures lays
  them side by side as Word does, on screen and in print; a single line,
  row, picture or text box taller than a page still runs into the gap
  rather than being cut. Since
  2026-09-14 the ruler moves a paragraph's four indents, its tab stops,
  the page margins and a table's columns by hand, and grips on a table's
  borders move its columns and rows, and the print paginator lays a
  table on the same grid.)
- **Fields.** Bookmarks arrived 2026-09-24, Insert → Bookmark naming a span
  of paragraphs (`w:bookmarkStart`/`w:bookmarkEnd`) that Go To finds again
  and a bookmarked paragraph stays editable through it. Cross-references
  arrived the same day: Insert → Cross-reference puts a REF field on a
  bookmark, and Update Fields (F9) refreshes them from the bookmark's
  words. Captions arrived the same day too: Insert → Caption writes a
  labelled, numbered paragraph — a SEQ field, `w:fldChar` begin/separate/end
  around ` SEQ Figure \* ARABIC `, the complex-field shape Word itself
  writes for one — and Update Fields renumbers every label's captions from
  where they now sit, the same pass that refreshes a REF. The table of
  contents arrived 2026-09-25 as the field Word writes — a TOC field in its
  building block, entries linked to `_Toc` bookmarks with PAGEREF page
  numbers from the window's own page layout — and Update Table rebuilds
  it. An index, a table of authorities, citations and a bibliography,
  `PAGE` outside the header — none are written yet.
- **Mail merge**, envelopes and labels: a data source, merge fields, a
  preview and a run. Nothing of it exists, and Mail beside it is exactly
  what would make it worth having.
- **Recording tracked changes.** Recorded since 2026-09-25: Review → Track
  Changes writes `w:trackRevisions`, typing and deleting are kept as
  `w:ins`/`w:del` with author and date, four markup views, Accept and Reject
  one or all, Previous and Next. Not recorded yet: a paragraph split or
  joined, a paste, find and replace, a sort, and formatting changes
  (`w:rPrChange`). Document comparison and Restrict Editing sit behind it.
- **Text effects and equations.** Outline, shadow and glow arrived
  2026-09-24 from Home → Text Effects, drawn on the page and, outline and
  shadow, in print; a glow is not printed. Drop caps arrived the same day:
  Insert → Drop Cap, dropped, two lines, or in margin, the letter framed
  with `w:framePr` and laid as a float on screen and in print. WordArt,
  SmartArt, OMML equations: read past, not written.
- **A real watermark** (the header holds text only), multilevel list
  definitions. (Paragraph shading and borders, and the page colour, are
  written from the ribbon since 2026-09-21, and printed.)
- **Building blocks** (Quick Parts) and a default template.

## 2. Worksheets — the workbook

The calculation engine is the strongest part of the suite; the sheet around
it is where the gaps are.

- **Things in a sheet that are not cells**: pictures and shapes are drawn
  when a file has them; since 2026-09-21 Insert → Pictures puts a picture
  from a file at the cell, at its own proportions, written as Excel keeps
  one (a media part, the drawing part's one-cell anchor and relationship).
  Since 2026-09-20 a hyperlink is read with its
  target, drawn underlined, followed with Ctrl+click (an address opens
  outside, a place in the workbook is gone to), and put on a cell, changed
  or taken off with Ctrl+K — an external relationship, as Excel writes
  one. A cell comment is drawn with its red corner and read on hover, and
  the status bar counts a sheet's notes; since the same day a note is put
  on a cell, changed or taken off with Shift+F2, the Review tab or the
  cell's menu, written as Excel keeps one — the comments part, the VML box
  it draws the words in, and the sheet pointing at both. Threaded comments
  (Excel 365's replies) are read through their legacy shadow only. Since
  2026-09-21 Format as Table (Home, or Insert → Table) makes the selection
  or the block of data round the cell a table — a header row, banded rows,
  a style Excel knows by name — written as Excel keeps one: the table part
  with its columns and autofilter, the sheet's rels and `tableParts`; an
  empty header cell is given its column's name. Painted at once, in this
  window's own palette, as tables from files already were.
- **Frozen panes** are read, written, chosen from the View tab and, since
  2026-09-20, pinned on screen: the frozen rows sit under the column
  headings, the frozen columns beside the row headings, the corner at both,
  their headings with them. The fill handle keeps out of a frozen pane; a
  merged range that crosses the freeze line is drawn by the pane its
  top-left cell is in.
- **A cell's alignment** is written in full since 2026-09-21: the Home
  tab's indent arrows move the words in and out by Excel's units, and Text
  orientation angles, turns or stacks them, on screen and on the printed
  page, kept in the file as Excel keeps them.
- **Slicers and sparklines.** Sparklines arrived 2026-09-24: Insert →
  Sparklines draws a line or a column chart in a cell, written as Excel
  writes them — the x14 sparkline groups in the worksheet's extension
  list — and drawn in the cell layer under the cell's own words. Slicers
  are still read, not written.
- **Data tools**: consolidate and forecast sheets. (Text to columns,
  remove duplicates and a sort by up to three keys are built, 2026-09-21;
  group and outline, subtotals, advanced filter with a criteria range and
  flash fill, 2026-09-25 — Auto Outline, SUBTOTAL 101–111 and an advanced
  filter copied to another sheet are not.)
- **Auditing**: evaluate step by step. Error checking and the watch window
  both arrived 2026-09-24: Formulas → Error Checking opens a pane listing
  every error cell and circular reference, with Next and Previous to walk
  them; Formulas → Watch Window opens a pane listing cells chosen with Add
  Watch — sheet, cell, value and formula — that keeps reading their values
  live as the workbook recalculates, for as long as the document stays
  open (nothing is written to the file, the same as Excel's own watch
  window). (Trace precedents and dependents draw the dependency graph on
  the grid since 2026-09-21.)
- **Page-break preview** and a sheet background. (The print settings live
  in the file, and since 2026-09-21 the Page Layout tab writes them:
  margins, orientation, paper, print titles, scale to fit and manual page
  breaks — before that its buttons changed a note in the window that
  nothing read.)
- **Workbook protection** (structure), editable ranges on a protected sheet,
  custom views, manual calculation.
- **The formula language**, where it stops: `LAMBDA`. Array constants and
  `LET` came on 2026-09-09; `TEXT()` speaks every format code the grid does
  (the formatter moved in beside the engine), and serial 60 is 29 February
  1900 with Excel's own weekday arithmetic, as of the same day.
- **Power Query, Office Scripts, macros**: not built, and macros are
  deliberately never run — a workbook that runs code it arrived with is how
  ransomware starts. VBA in a file is preserved untouched.

## 3. Presentation — the deck

The largest register of the three: the deck writer is younger than the
document and workbook writers.

- **Shape formatting**: since 2026-09-14 the Format pane sets a shape's
  fill and outline (colour, weight, dashes) and Quick Styles gives it the
  theme's looks, and every shape moves and resizes by hand on the stage.
  Since 2026-09-25 the fill is also a gradient, a picture or a see-through
  colour, and Shape Effects writes glow, soft edges and reflection beside
  the shadow; Arrange aligns, distributes, rotates, flips, groups and
  ungroups a selection of shapes gathered with Shift+click. Bevel, 3-D
  rotation, pattern fills and editing points are still read-only.
- **Text inside a shape**: since 2026-09-21 bullets, numbering, list
  levels, line spacing, strikethrough, character spacing, change case and
  a highlight are written from the Home tab, and a slide's own bullets,
  spacing and run looks survive an edit (until then a format press or an
  edited line dropped them), and the box's own anchor, text direction
  and columns are written from the same tab. WordArt is still read past.
- **What can be put on a slide**: a table and a chart are read and drawn —
  a chart since 2026-09-10, through the same chart kit a worksheet uses,
  found by the corpus deck whose six chart slides showed an empty frame
  each. Both arrived to insert 2026-09-24: Insert → Table puts one on the
  slide with its cells edited in place, and rows and columns added and
  removed from a right-click; Insert → Chart puts a cached chart part on
  the slide, no embedded workbook, its data edited from a dialog. Video,
  audio, SmartArt and a hyperlink are still named and boxed, not
  insertable.
- **The master and layouts.** Masters and layouts are read and inherited
  from, and since 2026-09-10 the Designs pane puts an existing slide on
  another of the deck's layouts and starts a new slide from one (since
  2026-09-21 the Home tab's Layout does the same, and Reset puts moved
  placeholders back where the layout has them); editing
  them, swapping a theme, headers and footers, slide numbers and date
  fields all need the writer to reach the master. Background styles
  arrived 2026-09-24: Design → Background Styles gives a slide its own
  `p:bg` — a solid colour, a theme colour or a gradient — with Apply to
  all, though the master's own background is still not editable there.
  The Layers pane of the same day gives the drawing order, hiding and naming;
  grouping and alignment across shapes arrived 2026-09-25.
- **Structure**: since 2026-09-24 Home → Section adds, renames and removes
  sections, drawn as headings above their first slide in the thumbnail strip
  and in Slide Sorter view, written as PowerPoint's own `p14:sectionLst`
  extension list and kept in step as slides are added, duplicated, moved or
  deleted; hidden slides arrived the same day, Slide Show → Hide Slide
  marking a slide `show="0"` so the show steps over it while the strip,
  Slide Sorter and printing still draw it; custom shows and saving as a
  show are still to come.
- **The show**: recorded narration and timings, transitions and animations
  as anything but ribbon tabs, export to video.
- **Review**: comments on a slide, tracked changes. Find and replace across
  a deck arrived 2026-09-24: Home → Find (Ctrl+F) and Replace (Ctrl+H) open
  a small pane at the stage's own top right, the way Word's find pane does
  rather than a dialog over the slide; Next and Previous walk every hit in
  slide order, reading "n of m" and landing on the hit's slide and shape,
  Match case narrows the search, and Replace and Replace All rewrite a hit
  or every hit while keeping each run's own look. A match that would
  straddle two runs is not offered, since rewriting across the boundary
  where a run's formatting changes would either lose the second run's look
  or have to invent a blend of the two.

## 4. Mail — the one with no register

Mail carries no `<Soon>` markers at all: IMAP and SMTP, sign-in to Google and
Microsoft, rules, search, conversations, attachments in one place, tracker
blocking and unsubscribe are built. What it is missing is not inside mail:

- the **calendar** and **contacts** above, which is where a mail client
  stops being half an application;
- **local search across archives** is there, but there is no index, so a
  large mailbox searches by walking;
- ~~**out-of-office and send-later**~~ — both built 2026-09-25: Send later
  holds a message in an Outbox on this computer until its time and sends it
  while the suite runs (at start-up if it fell due while closed), a failed
  send staying there with its error; automatic replies are per account,
  follow RFC 3834's skips and answer each sender once while on. Both go out
  only while Rutba Office is running — holding them on the server (Gmail's
  vacation responder, Microsoft 365's automatic replies) needs a sign-in
  permission the suite does not ask for yet.
- ~~adding an account asked for the servers and guessed them from the domain~~ — built 2026-09-09: an address and a password; the server is found from the provider table, MX, SRV, autoconfig, Microsoft autodiscover and a knock on the conventional names, with Advanced always a click away.
- ~~signatures~~ — built 2026-09-24: plain text, per account, edited from that account's own settings; a new message carries it after the "-- " line, a reply or forward puts it above the quote, and switching the From account swaps the block for the one it replaces, but only while it is still exactly what was inserted.

## 5. The formats, read against written

| Format | Read | Write |
| :-- | :-- | :-- |
| `.docx` `.xlsx` `.pptx` | yes, preserving | yes, preserving |
| `.doc` `.xls` `.ppt` | text only — an approximation, and only `.xls` says so | no |
| `.odt` `.ods` `.odp` | yes — an `.ods` reads its named ranges, merged cells and number formats since 2026-09-10; its charts and shapes are not read | **yes**, 2026-09-09 — what the suite models: text, structure, tables and run formatting; values, formulas and value types; text boxes, pictures and notes |
| `.rtf` | yes | **yes**, 2026-09-09 |
| `.csv` `.tsv` `.txt` `.md` `.html` | yes | yes |
| `.pdf` | viewed | written, for all three kinds since printing landed |
| `.eml` `.msg` `.mbox` `.pst` `.ost` `.olm` | yes | mbox only |
| `.ics` `.vcf` | yes | **yes**, 2026-09-09 — events and cards, read and written; a reply to an invitation written as `.ics` |
| Encrypted OOXML | no | no |

RTF and OpenDocument were the trap — the installer told Windows this suite
was the **editor** of `.odt`, `.ods` and `.odp` while Ctrl+S on one refused —
and both are written now. What is left in the table is the binary Office
formats, which are read as text and never claimed as editable, and the
encrypted package.

## 6. Pictures — the viewer

Rebuilt 2026-09-20 for the large, mixed folder: a tile is the platform's
thumbnail kept on disk (never the file), a clip's tile is a frame rather than
a live player, the grid and the filmstrip draw only what is in view, and the
stage keeps a picture until the next has decoded. What is still missing:

- **Thumbnails on Linux.** The platform there makes none; a picture is read
  and shrunk by Electron's own decoder, which knows PNG and JPEG, and a clip's
  frame is drawn by the window. HEIC, AVIF, TIFF and SVG tiles show the kind's
  icon there. On Windows and macOS the system knows what it has a codec for.
- **A clip's length on its tile** arrived 2026-09-24, read once from the
  clip's own metadata and kept for the window's life, as the frames are.
- **A slideshow with transitions** arrived 2026-09-24: View → Slideshow
  (F5) crossfades through the folder's pictures at an interval you choose.
  A folder's counts before it is entered, ratings, tags and a search across
  folders are still to come.
- **Edits from the viewer** — rotate the file, not only the view — belong to
  the Image tool, which the viewer opens in one click.

---

## The order this should be built in

1. ~~**Autosave and recovery.**~~ Built 2026-09-09.
2. ~~**Page setup written into the file**~~, with the print area and the print
   titles. Built 2026-09-09 for workbooks; a document and a deck still choose
   theirs per print.
3. ~~**Write ODF**, or stop claiming it.~~ RTF and `.odt`, `.ods`, `.odp` are
   all written as of 2026-09-09; a file that came in as OpenDocument goes out
   as OpenDocument.
4. ~~**Calendar and contacts**, as a seventh and eighth app on the same shell.~~
   Both built 2026-09-09, and Mail answers an invitation from the message.
5. ~~**Floating layout in the document paginator.**~~ One piece of work
   unlocks eleven controls in Word and the same again in the deck. Pictures
   and text boxes beside the words print since 2026-09-20; the deck's
   align, distribute, rotate, flip and group arrived 2026-09-25.
6. ~~**Shape formatting and text properties in the deck writer**~~ —
   gradients, picture fills, transparency, glow, soft edges and reflection
   written since 2026-09-25; bevel and 3-D remain.
7. ~~**Things in a sheet that are not cells**: comments, hyperlinks~~ (both
   read and written since 2026-09-20), ~~format as table, pictures~~ (both
   written since 2026-09-21).
8. ~~**Fields in the document**: bookmarks first, then cross-references,
   captions~~ (all three built 2026-09-24) ~~**and a table of contents that
   refreshes**~~ (built 2026-09-25).
9. ~~**Recording tracked changes**~~ (built 2026-09-25), then comparison and
   restricted editing.
10. ~~**The formula language's remaining corners**: array constants, `LET`,
    `TEXT()` through the formatter the suite already owns.~~ Built 2026-09-09,
    with the phantom day; `LAMBDA` remains.

Everything above this line is work somebody can start on Monday. Below it
sit the things that are deliberately not built — macros that run, cloud
services (translation, dictation, design ideas, co-authoring), and anything
that would make a request this suite has not declared.
