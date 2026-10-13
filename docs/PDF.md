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
`text`, `inspect` and `close`.

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

## Part 2

In the order they matter to a person:

1. **Forms**: fill and save AcroForm fields, keep fields when pages are
   reordered (the form's field tree is not copied yet: a page keeps its
   widgets' appearance, not their being fields).
2. **Bookmarks and named destinations** carried through a reorder, a merge
   and a split (outlines are dropped today), and page labels.
3. **Annotation**: highlight, underline, strike-out, notes, ink and shapes,
   written as annotations an incremental update appends.
4. **Signatures**: see and verify a signed PDF; sign, as an incremental
   update that leaves the signed bytes alone (`incremental.js` is the start).
5. **Redaction** that removes the words, not a black box over them.
6. **Text in place**: change a line of words where it stands.
7. **Password-protected PDFs**: open with the password (RC4 and AES), and
   change and save their pages; protect a PDF on save.
8. **Export to Word**: a page's words and pictures into a `.docx` through the
   Documents engine.
9. **Compress**: images resampled and object streams written, for a smaller
   file.
10. **OCR** for scans, later, as a declared download at most once.

Also: the CJK character maps; insert a blank page or a page from a picture;
duplicate pages; print through the system's own PDF printing rather than as
pictures; the engine's assembly on the document thread for very large files
(it runs on the main process today, which is quick for ordinary files).
