# What is missing, between this and a suite somebody could use for everything

Written 2026-09-09, from the code rather than from memory. Three sources: the
suite's own register of unbuilt controls (257 of them then, 36 in
1.29.5, 28 in 1.29.8, 24 in 1.29.9, 20 in 1.29.10, 19 in 1.33.0 — 4 in Documents, 5 in Worksheets, 10 in Presentations — each a
`<Soon>` in a ribbon with a sentence saying what it needs), a reading of what the engines
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
| **OpenDocument drawings** | **Built, 2026-10-08** | An `.odp` opens as it was made: slides at its own page size holding each drawing where it stood — text boxes, pictures, shapes in their fill and outline with their words, lines and tables — and its notes; since 2026-10-08 its gradients, backgrounds and charts too, and its shapes as their own outlines, all written back on save. An `.ods` keeps its column widths, hidden rows and columns and frozen panes, and since 2026-10-08 its charts (embedded objects, plotting their cells), shapes, pictures and text boxes, read and written back. An `.odt` opens whole since 2026-10-08 — headings, run looks, links, lists at their levels, tables with their spans and shading, pictures, the page — where it came in as plain lines. Found 2026-09-10 by tools/make-rich-fixtures.ps1. |
| **Equations** | **Built, 2026-09-25** | Office Math is read, drawn through the browser's own MathML, edited in Word's linear format, written as Word writes it and printed. A deck's equations are read, drawn and inserted the same way since the same day. |
| **Tracked changes, shown** | **Built for documents, 2026-09-25** | Insertions and deletions are kept with their author and date, shown as All Markup, Simple Markup, No Markup or Original, walked with Previous and Next, and accepted or rejected one at a time or all at once (since 1.30.0 all at once reaches inserted and deleted table rows, and the headers, footers and notes); recording them arrived the same day (below). A deck's and a workbook's tracked changes are not read. |
| **Page setup as a document property** | **Built for workbooks, 2026-09-09; documents and decks, 1.32.0** | Paper, orientation, margins, scaling, gridlines, headings, centring, the print area and the repeated rows are read from and written to the file where Excel keeps them. Since 1.32.0 a document's print shows and changes the document's own paper, orientation and margins, as Word's does, and a deck's print choices (slides, notes or a handout, slides to a page, a frame) are kept in `p:prnPr`, as PowerPoint keeps them. |
| **Password-protected files** | **Built, 2026-09-25** | An encrypted `.docx`, `.xlsx` or `.pptx` opens through a Password dialog, reading Agile and Standard Encryption and checking the file's HMAC. Info → Encrypt with Password saves Agile AES-256/SHA-512 inside a compound file as Office writes one, and the autosave copy is encrypted too. Since 2026-10-08 a password-protected Office 97–2003 file opens too — RC4 CryptoAPI and Office 97/2000 RC4 in `.doc`, `.xls` and `.ppt`, and Excel 95's XOR; Word's own older XOR obfuscation is refused with a message saying so. |
| **A spell-check pass** | **Built, 2026-09-25** | Review → Spelling (F7) runs an Editor pane over a document, a workbook or a deck with Change, Change All, Ignore and Add to Dictionary. It uses English (UK) and English (US) Hunspell dictionaries on the machine and a personal dictionary shared by the three apps. Since 1.29.8 Review → Thesaurus (Shift+F7) opens a pane of words of like meaning from a thesaurus of the suite's own, some six thousand words in British spelling, in all three apps. Suggestions are not ranked by how common a word is ("brwn" offers bran before brown). The language is chosen once per document, not per paragraph; words marked as another language or not to be checked (Review → Language, 1.29.1) are read past rather than checked in their own. The grid and the slides have no spelling right-click. No other languages ship yet. |
| **Accessibility check** | **Built, 2026-09-25** | Review → Check Accessibility lists Errors, Warnings and Tips under Office's rule names in all three apps, with one-click fixes, Alt Text, Mark as decorative and a status-bar indicator that follows edits. Since 2026-10-08 a table cell's shading — its own and its table style's — is read for contrast, and merged cells in a document unmerge in one click; merged cells elsewhere and unclear links have no one-click fix. |
| **Right-to-left layout** | **Built, 2026-10-08** | A paragraph Word marks right to left, or whose style does, runs from the right margin with its indents, hang, borders and drop cap mirrored; Home → Paragraph turns one either way, and its alignment is written mirrored as Word writes it. Since 2026-10-08 the ruler and tab stops count from the right margin in such a paragraph, a table runs from the right (`w:bidiVisual`, Table Layout → Right to Left), a section lays its columns from the right, Page Layout → Sheet Right-to-Left mirrors a worksheet, and a slide's paragraphs run right to left from Home → Paragraph. Since 1.29.8 a document's PDF prints Arabic, Hebrew, Greek and Cyrillic in a font the computer has that may be embedded, subset to the glyphs used with a ToUnicode map, Arabic joined (Urdu's and Persian's letters among them), right-to-left lines in drawing order by the bidirectional algorithm. A workbook's and a deck's PDF, laid out and printed by Chromium, carry those scripts as text too, which a window check now holds them to. Since 1.38.0 the windows' own frame turns right to left in Urdu, the ribbon, panes and dialogs with it, while a page, a grid and a stage keep their documents' direction. |
| **Localisation** | **Urdu, 1.38.0** | A message catalogue: each string looked up by its English words (`t`, and `tn` for counts under each language's plural rules), the language settled before a window loads anything (renderer/language.js) from the app menu's choice or the system, `messages.json` made from the sources by tools/extract-messages.mjs and checked current. Since 1.38.0 every word the windows show goes through it, nearly 4,900, held there by a check (tools/find-english.mjs, tests/english-left.test.js), and Urdu is translated whole, machine-drafted against a glossary (catalogues/GLOSSARY-ur.md), its frame right to left and its dates in Urdu. Since 1.39.0 what the engines and services say themselves reaches the window in its language too: some 310 messages they throw, listed from their sources (renderer/engine-words.js) and found by the message as it arrives, its names put back (messages.js `tFilled`), and the names they hand the windows (citation styles, source types, envelope sizes, a query's applied steps) through the catalogue as well. Not yet: a native speaker's review of the Urdu; the installer; a second language. |

---

## 1. Documents — the document

Reads more than it writes. What it draws faithfully and cannot yet produce:

- **Floating layout.** On screen, since 2026-09-14, a picture or a text box
  anchored at the left or right floats there and the words wrap round it;
  Wrap Text and Position on the Layout tab move a picked picture between the
  line, the sides and the middle, and the file keeps it. Since 2026-09-20
  print and PDF do the same for a picture: the paginator lays the lines
  beside it shorter, for as far down as it reaches and into the paragraphs
  after it, and a centred or right-aligned block keeps its side, and since
  the same day a floating text box prints beside the words the same way.
  Since 2026-09-25 Insert → Text Box writes a box of words, and Layout →
  Arrange puts a drawing behind or in front of the words, on screen and on
  paper, with z-order, Align, Distribute, Rotate, Flip, Group and a
  Selection Pane.
- **Columns and section control.** Columns arrived 2026-09-24: Layout →
  Columns writes `w:cols`, and print and PDF flow the words into them,
  each column filled to the foot of the page before the next starts, its
  own floats and footnotes at its foot and a separator drawn down the gap
  when asked for. On screen the editable flow is laid at the first
  column's width in print layout, as Word's own draft view does, and at
  the browser's own columns in draft layout. Hyphenation arrived 2026-09-25 (Layout → Hyphenation, with the suite's own
  English hyphenator); line numbers and page borders are the same story still. (Pages themselves arrived on
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
- **Tables.** Since 2026-10-08 a table is drawn and printed in its own
  lines and shading and in its table style — the header row, total row,
  first and last columns and bands its `w:tblLook` turns on — a merged cell
  as one cell over the rows it covers, and its header rows again at the head
  of each page it runs onto. Table Design picks a style (Table Grid, Plain
  Table 1, Grid Table 1 Light, Grid Table 4 and List Table 4 in each accent,
  written in the document's theme colours when it has none), its options,
  shading and borders with a pen; Table Layout puts rows and columns in and
  takes them out, merges cells and splits one into columns and rows (or a
  merged one back), draws a new table or a line through a cell with the
  Draw Table pen and rubs one out with the Eraser, splits the table, sets a column's
  width, AutoFit and Distribute Columns, aligns the table and a cell's
  words, turns a cell's words to read up or down, sets the cells' margins,
  repeats header rows, sorts the rows by a column, works out a formula
  (=SUM(ABOVE), =B2*C2, =AVERAGE(B2:B4) and the like) and updates them
  all, converts the table to text and back, turns it right to left and
  shows or hides its gridlines; Properties sets its width, alignment and
  alt text and the caret's row and column. A formula is worked out again
  on Update all formulas, as Word does on F9, not as its numbers are typed.
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
  it. Citations and a bibliography arrived 2026-09-25 in APA, MLA, Chicago,
  Harvard and IEEE, written as Word writes them. An index (XE and INDEX
  fields, with Mark Entry), a table of figures, and PAGE, NUMPAGES, DATE,
  FILENAME, AUTHOR and TITLE fields in the body arrived 2026-09-26; the
  index is laid in one column here. A table of authorities (TA and TOA
  fields, with Mark Citation) arrived 2026-10-07.
- ~~**Mail merge**, envelopes and labels~~ — built 2026-09-25: sources
  from a workbook, a .csv, Contacts or a typed list; Word's merge fields,
  Address Block, Greeting Line and rules; preview; a merge to a new
  document, to the printer or to e-mail through Mail; envelopes and Avery
  label sheets. The Ask, Fill-in and Set Bookmark rules arrived in 1.29.9,
  their questions answered before the merge runs; since 1.29.10 the
  recipient list's ticks and sort are kept in the file, in a part of the
  suite's own (Word keys its own by its data source's record ids).
- **Recording tracked changes.** Recorded since 2026-09-25: Review → Track
  Changes writes `w:trackRevisions`, typing and deleting are kept as
  `w:ins`/`w:del` with author and date, four markup views, Accept and Reject
  one or all, Previous and Next. Since 1.29.4 a paragraph split or joined,
  each line of a paste, Replace All, a sort, and changes of formatting to
  words and to paragraphs (`w:rPrChange`, `w:pPrChange`) are recorded too.
  Restrict Editing arrived 2026-09-25 and is enforced in the window; Compare
  arrived in 1.29.3. Since 1.30.0 a change made over several runs or
  paragraphs is saved with an id for each piece, as the schema asks.
- **Text effects and equations.** Outline, shadow and glow arrived
  2026-09-24 from Home → Text Effects, drawn on the page and, outline and
  shadow, in print; a glow is not printed. Drop caps arrived the same day:
  Insert → Drop Cap, dropped, two lines, or in margin, the letter framed
  with `w:framePr` and laid as a float on screen and in print. WordArt,
  SmartArt, OMML equations: read past, not written.
- **WordArt's shapes.** A run's WordArt look (outline, glow, shadow, no
  fill) is written. Since 1.31.0 Shape Format → Text Effects → Transform
  lays the words along Arch Up, Arch Down, Circle or Button in Documents,
  Presentations and Worksheets: the `a:prstTxWarp` preset on the text body
  is read (a deck's adjust handles with it) and written back first in the
  body's properties, as Office writes it, and the words are drawn along
  the preset's path (`@rutba/drawing/warp`) on the page, the slide and the
  sheet, in a slide's thumbnail and in the show. In Documents the words
  show straight while the caret is in them. Since 1.32.0 Transform → More
  gives twenty of the warps (the waves, inflate and deflate, slants,
  triangles, chevrons, fades and curves), each letter stretched between
  the warp's two curves; since 1.34.0 thirty, with Square, the poured
  arches, cans, rings, cascades and Stop; since 1.36.0 all thirty-six,
  with Fade: Up and Down, the deflate-inflate pairs and the poured Circle
  and Button, and a Documents PDF writes a transformed box's words a
  letter at a time as the page draws them. Still to build: the handles,
  kept from the file but neither drawn nor dragged here. Asked for
  2026-10-08.
- **Multilevel list definitions.** Since 1.34.0 Home → Multilevel List is
  Word's gallery of list libraries (1. 1.1. 1.1.1., 1) a) i), I. A. 1.,
  Article I. with Section 1.01, and the bullets), legal levels and zero
  padding labelled as Word labels them. Since 1.36.0 Define New Multilevel
  List makes a list of one's own level by level, and a level can be linked
  to a heading style, every heading of that style numbered by it. Still to
  build: Word's More options (a level's font, the tab after the number,
  restarting after a level), and changing a list already defined.
  (A picture watermark arrived in
  1.32.0: Design → Watermark → Picture watermark, washed out behind every
  page, written in the header as Word writes one.) (Paragraph shading and borders, and the page colour, are
  written from the ribbon since 2026-09-21, and printed.)
- **A default template.** Design → Set as Default starts new blank
  documents in a document's theme and styles, and since 1.33.0 templates
  of one's own are kept where Office keeps them (Save as Template, from the
  app menu of all three apps, into Documents\Custom Office Templates) and
  listed on Home, each making a new untitled document. Not yet: a Normal
  template file whose own text, headers and page setup new documents take.
  (Building blocks — Insert → Quick Parts, the Quick Part and AutoText
  galleries and the Organizer — arrived in 1.29.4.)

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
- **Themes**: since 2026-09-26 Page Layout → Themes, Colours, Fonts and
  Effects write the workbook's theme as Excel does, from the same eleven
  themes Presentations offers, and cells, tables, charts and shapes follow
  it. A chart keeps the suite's colours while the workbook is on an Office
  theme.
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
  and PivotCharts arrived 2026-09-25: Insert → Slicer filters a table or a
  pivot from a panel of buttons, and Insert → PivotChart draws a chart
  bound to a pivot. Both are written as Excel writes them.
- **Data tools**: consolidate and forecast sheets arrived 2026-09-25. (Text to columns,
  remove duplicates and a sort by up to three keys are built, 2026-09-21;
  group and outline, subtotals, advanced filter with a criteria range and
  flash fill, 2026-09-25; SUBTOTAL leaves out the rows a filter hides,
  and 101–111 every hidden row, Auto Outline builds the outline from the
  formulas, and an advanced filter copies a list from another sheet, since
  2026-10-08.)
- **Auditing**: Evaluate Formula steps a formula part by part since
  2026-09-25, into the cells it reads and back. Error checking and the watch window
  both arrived 2026-09-24: Formulas → Error Checking opens a pane listing
  every error cell and circular reference, with Next and Previous to walk
  them; Formulas → Watch Window opens a pane listing cells chosen with Add
  Watch — sheet, cell, value and formula — that keeps reading their values
  live as the workbook recalculates, for as long as the document stays
  open (nothing is written to the file, the same as Excel's own watch
  window). (Trace precedents and dependents draw the dependency graph on
  the grid since 2026-09-21.)
- (A sheet background, Page Layout view, Page Break Preview and Split all
  arrived 2026-09-25. The print settings live
  in the file, and since 2026-09-21 the Page Layout tab writes them:
  margins, orientation, paper, print titles, scale to fit and manual page
  breaks — before that its buttons changed a note in the window that
  nothing read.)
- (Workbook protection, editable ranges on a protected sheet, custom views,
  manual calculation and threaded comments all arrived 2026-09-25.)
- **The formula language**: `LAMBDA`, named or called where it is written,
  with MAP, REDUCE, SCAN, BYROW, BYCOL and MAKEARRAY, XMATCH, TEXTSPLIT and
  the array-shaping functions arrived in 1.29.4, saved with the `_xlfn.`
  prefixes Excel needs. Array constants and `LET` came on 2026-09-09; `TEXT()` speaks every format code the grid does
  (the formatter moved in beside the engine), and serial 60 is 29 February
  1900 with Excel's own weekday arithmetic, as of the same day.
- **Power Query** arrived in 1.29.8: Data → From Table/Range and From
  Text/CSV shape a table, a range or a delimited file in the Power Query
  Editor — columns removed, renamed and typed, rows filtered, sorted,
  deduplicated and grouped, text split and tidied, each an applied step —
  and Close & Load writes the result as a table on a sheet of its own;
  Queries & Connections and Refresh All run it again on its source. The
  queries live in a part of the suite's own, which Excel keeps but does not
  read as its own Power Query (that is a binary package this suite neither
  reads nor writes), and there is no M language, no web or database source.
  Merge Queries (left outer, inner, left anti, full outer) and Append
  Queries arrived in 1.29.10, a query reading another query's result.
  Since 1.31.0 a column's type follows it through every step, and a column
  made dates loads in the short date format. Since 1.33.0 Fill Down and Up,
  Unpivot Other Columns, Merge Columns and Extract are steps too, and since
  1.35.0 Pivot Column and Conditional Column; custom columns, which want
  M, are not.
- **A cut pasted** carries every reference with it since 1.29.8, as Excel's
  does: formulas on any sheet and defined names that name only the moved
  cells follow them, onto another sheet too.
- **Scripts** arrived in 1.29.10: Automate → New Script, All Scripts and
  Record Actions, scripts in the shape of Office Scripts (`function
  main(workbook)`) kept on this computer, run in a worker that reaches
  nothing but a snapshot of the workbook, their edits one undo step. Only
  the commonest of Office Scripts' calls are there (ranges, values,
  formulas, number formats, fonts, fills, alignment, sheets added and
  renamed, and since 1.35.0 column widths, row heights, AutoFit, merges,
  strikethrough and vertical alignment, which Record Actions writes too);
  a script that awaits is refused, and rows and columns are not inserted
  or deleted by one.
- **Macros** are deliberately never run — a workbook that runs code it arrived with is how
  ransomware starts. VBA in a file is preserved untouched.

## 3. Presentations — the deck

The largest register of the three: the deck writer is younger than the
document and workbook writers.

- **Shape formatting**: since 2026-09-14 the Format pane sets a shape's
  fill and outline (colour, weight, dashes) and Quick Styles gives it the
  theme's looks, and every shape moves and resizes by hand on the stage.
  Since 2026-09-25 the fill is also a gradient, a picture or a see-through
  colour, and Shape Effects writes glow, soft edges and reflection beside
  the shadow; Arrange aligns, distributes, rotates, flips, groups and
  ungroups a selection of shapes gathered with Shift+click. Since 1.29.4
  pattern fills, a bevel, depth and 3-D rotation are written, and Edit
  Points changes a shape's outline point by point. Since 1.29.5 a heart, a
  smiley, callouts, block arrows, a can, a cube, a cloud and the flowchart
  shapes are drawn as their outlines, not as a box; rarer presets still are.
- **Text inside a shape**: since 2026-09-21 bullets, numbering, list
  levels, line spacing, strikethrough, character spacing, change case and
  a highlight are written from the Home tab, and a slide's own bullets,
  spacing and run looks survive an edit (until then a format press or an
  edited line dropped them), and the box's own anchor, text direction
  and columns are written from the same tab. WordArt's run look is written,
  and since 1.31.0 its Transform's arcs, circle and button are drawn and
  written from Shape Format, since 1.32.0 twenty of the warps, and since
  1.36.0 the whole gallery (see Documents, "WordArt's shapes"; the same
  drawing serves all three apps, since the preset lives on the shared text
  body).
- **What can be put on a slide**: a table and a chart are read and drawn —
  a table in its table style since 1.29.5 (its header, bands, borders and
  first and last rows and columns, which Table Design turns on and off since
  2026-10-08, with its cells' shading), and a chart since 2026-09-10, through the same chart kit a worksheet uses,
  found by the corpus deck whose six chart slides showed an empty frame
  each. Both arrived to insert 2026-09-24: Insert → Table puts one on the
  slide with its cells edited in place, and rows and columns added and
  removed from a right-click; Insert → Chart puts a cached chart part on
  the slide, no embedded workbook, its data edited from a dialog. Video,
  audio, SmartArt and hyperlinks are insertable too, video and audio from a
  file or recorded. 3D models arrived in 1.29.9, in Documents too: a glTF
  file drawn as a picture by the suite's own renderer, the model kept beside
  it and turned from the 3D Model tab; PowerPoint's own 3D model element is
  not written, so PowerPoint shows the picture and cannot turn it. Cameo
  arrived the same release: the camera, live, in a shape on the slide, and
  since 1.33.0 Record with the camera puts each slide's recording in its
  cameo's place, as PowerPoint does, with Reset to Cameo to take it off.
- **Design Ideas** arrived in 1.29.8, worked out on this computer rather
  than asked of a service: a pane of layouts the slide's own title, words
  and pictures suit, each drawn by the deck's renderer, one applied in a
  single undo step with a picture cropped (`a:srcRect`) to keep its
  proportions.
- **The master and layouts.** Masters and layouts are read and inherited
  from, and since 2026-09-10 the Designs pane puts an existing slide on
  another of the deck's layouts and starts a new slide from one (since
  2026-09-21 the Home tab's Layout does the same, and Reset puts moved
  placeholders back where the layout has them). Since 2026-09-25 View →
  Slide Master edits the master and its layouts on the stage — text
  styles, placeholders, shapes, backgrounds, new and renamed layouts — and
  Design swaps the theme (eleven of our own, with variants), its colours,
  fonts and effects. The handout and notes masters, and a second master,
  are editable since 2026-10-07. Background styles
  arrived 2026-09-24: Design → Background Styles gives a slide its own
  `p:bg` — a solid colour, a theme colour or a gradient — with Apply to
  all, and in Slide Master view the Slide Master tab's Background Styles
  sets the master's own or a layout's.
  The Layers pane of the same day gives the drawing order, hiding and naming;
  grouping and alignment across shapes arrived 2026-09-25.
- **Structure**: since 2026-09-24 Home → Section adds, renames and removes
  sections, drawn as headings above their first slide in the thumbnail strip
  and in Slide Sorter view, written as PowerPoint's own `p14:sectionLst`
  extension list and kept in step as slides are added, duplicated, moved or
  deleted; hidden slides arrived the same day, Slide Show → Hide Slide
  marking a slide `show="0"` so the show steps over it while the strip,
  Slide Sorter and printing still draw it; custom shows and saving as a
  show (`.ppsx`) arrived 2026-10-07.
- **The show**: transitions and animations are written and played since
  2026-09-25 (eleven transitions; entrance, emphasis and exit effects with
  the Animation Pane). Recorded narration and timings, triggers and
  export to video arrived 2026-10-06 and 2026-10-07; motion paths (written
  as PowerPoint writes a path drawn by hand) and effects by paragraph in
  1.29.9.
- **Review**: comments on a slide arrived 2026-09-25, written as
  PowerPoint 365's modern comments; tracked changes on a deck are not, but
  since 1.29.10 Review → Show Changes lists what is different about a deck
  since it was last open on this computer, slide by slide. Find and replace across
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

- ~~the **calendar** and **contacts** above~~ — both built 2026-09-09, on
  this computer; since 1.29.8 they are kept in step with a CalDAV and
  CardDAV server (iCloud, Fastmail, Nextcloud and the like, with an app
  password) — Google's and Microsoft's own calendars, which want their own
  sign-in, are not;
- **search** is indexed per account — an inverted index, with `from:`,
  `subject:`, `has:attachment`, `is:unread` and quoted phrases, an
  imported archive being an account of its own — but there is no one search
  across every account and archive at once;
- ~~**out-of-office and send-later**~~ — both built 2026-09-25: Send later
  holds a message in an Outbox on this computer until its time and sends it
  while the suite runs (at start-up if it fell due while closed), a failed
  send staying there with its error; automatic replies are per account,
  follow RFC 3834's skips and answer each sender once while on. Both go out
  only while Rutba Office is running — holding them on the server (Gmail's
  vacation responder, Microsoft 365's automatic replies) needs a sign-in
  permission the suite does not ask for yet.
- ~~adding an account asked for the servers and guessed them from the domain~~ — built 2026-09-09: an address and a password; the server is found from the provider table, MX, SRV, autoconfig, Microsoft autodiscover and a knock on the conventional names, with Advanced always a click away.
- ~~signatures~~ — built 2026-09-24: plain text, per account, edited from that account's own settings; a new message carries it after the "-- " line, a reply or forward puts it above the quote, and switching the From account swaps the block for the one it replaces, but only while it is still exactly what was inserted. Rich text since 1.29.4: bold, italic, underline, colour, size, a link and a picture, which goes as an attachment of the message.
- ~~a junk filter~~ — built in 1.29.4: Home → Junk marks a message junk or not junk and blocks or trusts a sender or a domain; Junk Email Options sets off, low, high or safe lists only, the Safe and Blocked Senders lists, and trust in contacts. Arriving mail is filed by the lists, the server's own spam headers and a filter that learns from what is marked, on this computer. Since 2026-10-08 its options are set per account (what it learns is shared), with Safe Recipients and "Never block this group or mailing list", and blocked top-level domains and encodings.

## 5. The formats, read against written

| Format | Read | Write |
| :-- | :-- | :-- |
| `.docx` `.xlsx` `.pptx` | yes, preserving | yes, preserving |
| `.doc` `.xls` `.ppt` | **yes**, 1.29.5 — Word 97–2003, 6.0/95, 2.0, 1.0, Write and DOS Word; Excel 2.1 to 97–2003 with charts, shapes, tables and conditional formats; PowerPoint 97–2003; one Office 2007 or later saved is opened from the newer description it keeps inside | no — saved as `.docx`, `.xlsx` or `.pptx` |
| `.odt` `.ods` `.odp` | yes — an `.ods` reads its named ranges, merged cells and number formats since 2026-09-10, and its column widths, hidden rows and columns and frozen panes since 1.29.4, its charts and shapes since 2026-10-08; an `.odp` its pictures, shapes, lines and tables where they stood since 1.29.4, its gradients, backgrounds and charts since 2026-10-08; an `.odt` whole since 2026-10-08 | **yes**, 2026-09-09 — what the suite models: values, formulas and value types; text boxes, pictures and notes; since 2026-10-08 an `.odt` written whole from the document (it had written each heading, list, table and picture as a plain paragraph), an `.ods`'s charts and shapes and an `.odp`'s shapes, tables, charts and backgrounds |
| `.rtf` | yes — whole since 2026-10-08: headings, lists at their levels, tables with their widths and merges, links, pictures, the page | **yes**, 2026-09-09 — since 2026-10-08 with headings' outline levels, real lists, tables' widths and merges, links and pictures |
| `.html` | yes — a page opens as a document since 2026-10-08: headings, looks, links, lists, tables with their spans, embedded pictures | yes — since 2026-10-08 with lists, tables and pictures |
| `.csv` `.tsv` `.txt` `.md` | yes | yes — a document's own numbered list kept as one in `.md` and `.txt` since 2026-10-08 |
| `.pdf` | viewed | written, for all three kinds since printing landed |
| `.eml` `.msg` `.mbox` `.pst` `.ost` `.olm` | yes | mbox only |
| `.ics` `.vcf` | yes | **yes**, 2026-09-09 — events and cards, read and written; a reply to an invitation written as `.ics` |
| Encrypted OOXML | **yes**, 2026-09-25 — Agile and Standard Encryption, the HMAC checked | **yes**, 2026-09-25 — Agile AES-256/SHA-512, as Office writes it |

RTF and OpenDocument were the trap — the installer told Windows this suite
was the **editor** of `.odt`, `.ods` and `.odp` while Ctrl+S on one refused —
and both are written now. The binary Office formats are read in full since
1.29.5 and saved as their newer counterparts, never written back; since
2026-10-08 their older RC4 and XOR encryption is read too.

## 6. Pictures — the viewer

Rebuilt 2026-09-20 for the large, mixed folder: a tile is the platform's
thumbnail kept on disk (never the file), a clip's tile is a frame rather than
a live player, the grid and the filmstrip draw only what is in view, and the
stage keeps a picture until the next has decoded. What is still missing:

- **Thumbnails on Linux.** The platform there makes none; a small PNG or
  JPEG is shrunk by Electron's own decoder, and since 1.32.0 anything else
  the window can decode (a large photo, an SVG, an AVIF, a WebP) is drawn by
  the window at the tile's size and kept, as a clip's frame is, on any
  platform whose system has no tile for it. HEIC and TIFF tiles still show
  the kind's icon on Linux. Not yet run on Linux: checked on Windows, where
  an SVG has no system tile.
- **A clip's length on its tile** arrived 2026-09-24, read once from the
  clip's own metadata and kept for the window's life, as the frames are.
- **A slideshow with transitions** arrived 2026-09-24: View → Slideshow
  (F5) crossfades through the folder's pictures at an interval you choose.
- **Ratings, tags, folder counts and Subfolders** arrived in 1.30.0: a
  picture is rated one to five stars (the keys 0 to 5) and tagged from the
  details, kept on this computer by its path and never written into the
  file; the grid filters by name or tag, by stars, and sorts Highest rated;
  a folder's tile says how much it holds; and Subfolders searches the
  folders inside for a name or a tag. XMP ratings written into the file, as
  Windows Photos and Lightroom write them, are not read or written.
- **Edits from the viewer.** Since 1.33.0 Rotate the picture (Ctrl+R) turns
  the file and saves it, as Windows Photos does: a JPEG by its EXIF
  orientation, losslessly, and a PNG redrawn. Other formats, and crops and
  adjustments, belong to the Image tool, which the viewer opens in one
  click.

---

## 7. Image and Video — the editors

Both are real but small, and until now unregistered here. What each does,
and what it does not, is set out with the plan in
[COMPETITORS.md](COMPETITORS.md).

- **Image** has rotate, flip horizontal, crop to a ratio, seven sliders
  (brightness, contrast, saturation, hue, blur, sepia, greyscale), undo and
  export to PNG, JPEG or WebP, over a non-destructive pipeline. The pipeline
  also knows flip vertical, resize, invert and annotations (rectangle,
  ellipse, arrow, text, pen) that no button offers yet. No selections,
  layers, masks, retouching, background removal, RAW, batch or project file.
- **Video** opens one file, splits, trims the ends, removes a clip, previews
  at four speeds (the model's per-clip speed has no button), and exports
  WebM by recording a canvas at the speed the film plays. One track; no
  joining of files, transitions, titles, captions, audio mixing, keyframes,
  colour, chroma key, stabilisation or MP4.
- **Pictures** opens either in one click; "edits from the viewer" belong to
  Image, and a batch over a folder belongs there too.

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
   written since 2026-09-25; bevel, 3-D and pattern fills since 1.29.4.
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
    with the phantom day; `LAMBDA` in 1.29.4.

This order is complete; [ROADMAP.md](ROADMAP.md) reviews how it was met
and sets the order from here. Everything above this line is work somebody can start on Monday. Below it
sit the things that are deliberately not built — macros that run, cloud
services (translation, dictation, co-authoring), and anything
that would make a request this suite has not declared.
