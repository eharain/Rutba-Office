# PDF

The suite's tenth app, Rutba PDF: read a PDF, and put its pages in order.
Part 1 is built (2026-10-13). This page says what it is, how pages are drawn,
and what part 2 adds.

## What is built

**The window** (`apps/desktop/renderer/apps/pdf/`). A strip of thumbnails
beside the pages. A press picks a page, Ctrl adds or takes away one, Shift
takes the run from the last one picked, Ctrl+A in the strip takes them all,
and dragging a picked page moves every picked page to where the line shows.
The ribbon:

- **Home**: Open, Save, Save As, Print; Zoom In, Zoom Out, Fit Width, Fit
  Page; Find; Rotate Right and Delete.
- **Pages**: Rotate Left and Rotate Right, Delete, Extract, Insert from File,
  Split, Merge Files.
- **View**: Single Page, Two Pages, Continuous; Thumbnails; Fit Width, Fit
  Page, Actual Size.

Ctrl and the wheel zoom about the pointer (the suite's `useWheelZoom`), and the
zoom slider is in the status bar beside "Page 3 of 12". A window opens at the
width of the window, never past 100%. The words on each page are laid over
it as pdf.js's text layer, so text is selected and copied as it is in any
reader. Find (Ctrl+F, F3 and Shift+F3) searches every page with the suite's own
matching (`@rutba/editing/find`, so Arabic's marks and letter forms are passed
over as in the editors) and marks the matches with the browser's own
highlights. Undo and redo cover every page operation. The thumbnails' context
menu has the page operations too.

Saving is careful. Save writes a new file through Save As unless the pages
came from a file, and then asks once a window before it replaces it ("Replace
the Original", or "Save a Copy…"). Extract writes the selected pages to a new
file and leaves the document as it is; Split writes files of N pages, or a new
file at each selected page, into a folder chosen (`name_Part1.pdf`, as Acrobat
names them); Merge Files adds other PDFs at the end, or, in an empty window,
combines them into a new document. While there is unsaved work, a recovery
copy (the pages as a PDF) is kept through `doc.keepCopy`, as Video and Image
keep theirs, and the launcher offers it back after a crash.

A protected PDF that opens without a password is read, with a bar saying its
pages cannot be changed or saved yet. One that needs a password to be read
says so in a sentence.

**The engine** (`packages/pdf/src/reader.js`, `assemble.js`). The package
that wrote PDFs now reads them: cross-reference tables and streams (with PNG
and TIFF predictors), hybrid files, object streams, every incremental update
chained by `/Prev` (the newest definition wins), bytes before the header, the
trailer, and the page tree with the resources, boxes and rotation a page
inherits from the nodes above it. A file whose cross-references are wrong is
read again by scanning for its objects. Encryption is recognised and refused
with a sentence. New files are made from pages without drawing anything: a
page is copied with every object it reaches, renumbered, its content streams,
resources, fonts and pictures carried as the bytes they were; what a page
inherited is written onto it; references to the page tree, the catalog and
pages left behind are cut, and a link to a page that is kept follows it to
its new place. The language, layers (`/OCProperties`) and viewer preferences
go with the first file's pages. `pageText` reads the words of a page set in a
simple font, which the checks use to tell pages apart.

**The main process** (`apps/desktop/main/pdf.js`, contract namespace `pdf`):
`open` reads a file once and holds it for its window (so a file can be saved
over itself), `assemble` writes the pages named (whole or not at all),
`addText` writes a recognised text layer, `text`, `inspect` and `close`.

**Recognise Text** (Home and Pages, the thumbnails' menu): the words in
scanned pages found on this computer and written into the file as an
invisible text layer, so Find, selecting and copying work on scans. A dialog
chooses all pages, the selected ones or this one, and the languages (the six
the suite reads, Urdu marked beta); pages that already have text are left
alone; a progress bar counts the pages, with Cancel. The layer is an
incremental update (`packages/pdf/src/textlayer.js`): the original bytes
stay the start of the file and no page is redrawn. Undo takes it away, and a
Save of the recognised file as it is writes the original with the layer
after it. docs/OCR.md has the details and the error rates.

**Comments and markup** (the Comment tab, `annots.js`; engine
`packages/pdf/src/annotations.js`). Highlight, Underline and Strikethrough
over the words selected (or select words with the tool in hand), Sticky Note,
Text Box, Draw (freehand), Rectangle, Ellipse, Line and Arrow, in six colours.
Each is written as Acrobat writes it (Text Markup annotations with
QuadPoints, Text, FreeText, Ink, Square, Circle, Line with its line endings)
with an appearance stream drawn by the engine, so every reader shows it as
this window does: pdf.js draws it from that appearance. Each change is an
incremental update (`update.js`): the original bytes stay the start of the
file, the file with the update is a new source, and Undo takes it away. The
Comments pane lists every comment by page with its author and time; a
comment is answered (a reply in reply to it, /IRT), marked resolved (a hidden
review state, Completed, as Acrobat's Set Status writes), its words changed,
or deleted with its replies (their objects freed in the update). Comments
other writers made are read and listed; one saved without an appearance is
drawn by the window from its kind and geometry; all of them are kept by
every save. Changes are written one after another, each to the newest
version of its file, so quick work in a row is never lost.

**Forms and Fill & Sign** (the Fill & Sign tab, `fields.js`; engine
`forms.js`). An AcroForm's text fields (multi-line, aligned, combed, sized to
fit), check boxes, radio groups, drop-downs and lists get a control laid
over the page; a value is written when committed, as an incremental update
with appearance streams in the form's own font, so any reader shows it.
Words Windows-1252 cannot spell are written as the value and the form marked
/NeedAppearances, so the reader draws them in a font it has. Read-only
fields cannot be changed. Fill & Sign places a signature or initials,
typed, drawn on a pad or taken from a picture of one signed on paper (its
paper made see-through), as a stamp; Add Text, Date (today's) and Checkmark
put words and ticks where the form has no field. The tab and the signature
dialog say plainly that these are pictures, not digital signatures.

**Redaction** (the Redact tab, `redact.js`; engine `packages/pdf/src/redact.js`).
Mark Text, Mark Area and Mark Matches (every match of the words in Find, on
every page) write Redact annotations, reviewed in the Comments pane. Apply
Redactions says what it does and writes a new file whole, never as an
update (an update keeps the old bytes): each marked page's content is read
with its text state, and every glyph whose box lies under a mark is taken
out of its string, the space it took kept by a TJ number so the words beside
it stay put; invisible text (a recognised scan's layer) goes the same way;
pictures under a mark are painted out (decoded here, or a JPEG decoded by
the window and handed back), and a picture that cannot be decoded is left
off the page; a form XObject is redacted as a new object, so other pages
drawing it keep it; /ActualText over removed glyphs goes; comments, links and
widgets over a mark go; each mark is filled black with its overlay text.
The words the marks covered (learnt from the glyphs removed, through each
font's /ToUnicode, and from the words a mark records: the selection, or the
words searched for) are then taken out of every other place in the file:
the document's title, subject, keywords, author and creator, XMP metadata
(the document's, a page's, a picture's), bookmarks, comments and their
replies, links, attached files, form values (blanked, with the appearance
that showed them), field tooltips and names, and page labels. A comment, a
bookmark or a link with a word in it goes whole. Before anything is saved
the dialog lists the words, what else it found by kind ("1 bookmark, 1
comment and the title"), and, by page, any picture under a mark it cannot
paint out and will leave off the page. Lines and shapes under a mark are
covered, not cut away, and a marked word left unmarked elsewhere on a page
stays there (Mark Matches marks every one). The engine's tests prove the
removal three ways, on a file that carries the word in every one of those
places: this reader, pdf.js, and a search of every byte and inflated
stream.

**Bookmarks, page labels and links** (`nav.js`; engine `navigation.js`,
`assemble.js`). The side pane shows Pages or Bookmarks; a bookmark goes to
its page where the page now stands (one whose page was deleted is greyed);
thumbnails and the status bar show each page's label; links on a page go to
their page or open their web address. Save, Extract, Split and Merge carry
bookmarks (a bookmark to a page left out gives way to its own), page labels
(as the fewest ranges), links and named destinations (made explicit), and a
form's fields whose widgets are on kept pages (a merge's second field of a
name renamed).

**Export as Pictures and Reduce File Size** (Pages tab, `export.js`; engine
`compress.js`). Pages are written as PNG or JPEG at 72 to 600 dots an inch.
Reduce File Size draws pictures shown at more detail than the page needs at
225, 150 or 96 dots an inch (JPEG from the window's canvas, or resampled and
deflated by the engine) and writes the file whole as a new one.

**Wiring.** The launcher's card (its New opens the window asking for files to
combine), the app menu and the Apps menu, `.pdf` owned by the PDF app (Pictures
still shows a PDF met among pictures, with Open in PDF on its navigator and
In PDF on its ribbon), the taskbar tile `resources/apps/pdf.ico`, and every
word in Urdu, Arabic and Hindi, the window right to left in Urdu and Arabic
with the pages left as they are.

## How pages are drawn

**pdf.js** (Mozilla, Apache-2.0, which the AGPL accepts), version 6.2.108,
vendored in `apps/desktop/vendor/pdfjs` by `tools/vendor-pdfjs.mjs` and copied
beside the bundle by `build/bundle.js`. The window loads it from the app's
own files the first time a PDF is shown; it parses and draws in a worker of
its own (served under a policy that lets it compile its WebAssembly image
decoders and fetch the app's own files, and nothing else), so a large file
never holds the window. Nothing reaches the network.

Why not another route: Chromium's own PDF viewer draws a whole document in a
plugin frame, with no way to draw a page into a thumbnail, lay text over it
or know which page is in view; drawing PDF content ourselves is a renderer
the size of pdf.js; a native library is a native module, which the suite has
none of. pdf.js is the renderer Firefox ships.

What is vendored: the library and its worker, the standard fonts a PDF names
without embedding (Foxit and Liberation, under their own licences), the
WebAssembly decoders for JPEG 2000 and JBIG2 scans and colour profiles
(OpenJPEG, PDFium's JBIG2, qcms, each under its own licence), and the CMYK
profile. Not vendored: the character maps for CJK text in fonts a PDF does not
embed (169 files), the scripting sandbox and the viewer. A PDF of Chinese,
Japanese or Korean text that does not embed its fonts draws without that
text until they are added.

Pages are drawn when they come near the screen, two at a time, those on
screen first; a page drawn is drawn again only when its size or turn
changes. The page operations never draw anything: the window keeps a list
of (file, page, turn), and the engine writes it.

## Checks

- `tests/pdf-pages.test.js`: the reader and the assembler on fixtures made by
  the suite's writer, by the incremental updater and by hand (object streams,
  a cross-reference stream with a predictor, a hybrid file, an incremental
  update that redefines a page's content, a damaged table, an encrypted
  file, links to kept and deleted pages, strings and names that must survive).
- `apps/desktop/main/verify-pdf-app.js`, key `pdf-app`: the window, pressed
  and read back from the files it writes.
- `packages/pdf/test/annotations.test.js`, `forms.test.js`, `redact.test.js`,
  `navigation.test.js`, `compress.test.js`, with pdf.js loaded in Node
  (`test/pdfjs.js`) as a second reader; window checks `pdf-annot`,
  `pdf-forms`, `pdf-redact`, `pdf-outline` and `pdf-export`
  (`main/verify-pdf-*.js`); `tools/fuzz-pdf.js` runs every new reader and a
  redaction over damaged files.

## Part 2

Built (2026-10-14): comments and markup, forms and Fill & Sign, redaction,
bookmarks, page labels and links through every page operation, Export as
Pictures and Reduce File Size (above). Still to come, in the order they
matter to a person:

1. **Digital signatures**: see and verify a signed PDF; sign with a
   certificate, as an incremental update that leaves the signed bytes alone
   (`update.js` is the start). Fill & Sign's signatures are pictures.
2. **Text in place**: change a line of words where it stands.
3. **Password-protected PDFs**: open with the password (RC4 and AES), and
   change and save their pages; protect a PDF on save.
4. **Export to Word**: a page's words and pictures into a `.docx` through the
   Documents engine.
5. **OCR**: built (Recognise Text, docs/OCR.md). Still to come: pages
   that mix text with a scanned picture, sideways scans, and a choice to
   recognise only the pictures on a page.
6. **Comments**: a text box's words in Urdu, Arabic or Hindi are kept in the
   file but drawn by its appearance only in Windows-1252's letters (an
   embedded font for the box is next); polygons and pop-ups are read and
   kept, not drawn or made; a whole-document "Sanitize" that strips
   metadata, attachments and scripts without marks.

Also: the CJK character maps; insert a blank page or a page from a picture;
duplicate pages; print through the system's own PDF printing rather than as
pictures; the engine's assembly on the document thread for very large files
(it runs on the main process today, which is quick for ordinary files);
object streams written when a file is made smaller.
