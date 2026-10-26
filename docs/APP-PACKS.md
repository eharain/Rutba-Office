# App packs

Only the apps a person chose are on their computer. An app not chosen is
absent, not turned off: no file of it, no IPC channel, no file type, no
tile and no AI tool. Rutba Office is then a **core** (the Electron runtime,
the shell, the launcher, settings, the updater and the engines several apps
share) and one **pack** per app, each a signed archive the core finds at
start and loads only when it checks out.

This is the design, from the code as it is. It has five parts: the map of
what each app takes from another (section 1), the pack format (2), the tools
that make and sign packs (3), the core's loader (4) and the move of the app
code into packs (5). Sections 1 to 4 are built and tested; nothing here is
wired into `main.js` yet, and no app code has moved.

## 1. The map

`tools/app-map.mjs` walks the import graph from each app's entry points and
says, per app, what it imports from another app, which engines and npm
packages it reaches, its heavy parts, its IPC namespaces and events, the
file types it claims, its tiles, its AI tools, its window checks and its
size.

```
node tools/app-map.mjs                the headline
node tools/app-map.mjs --json <file>  the whole map
node tools/app-map.mjs --doc          the tables below, written again
node tools/app-map.mjs --write-edges  the fixture the test reads
```

How it reads the code:

- **The graph** is esbuild's, the same resolver the renderer's build uses,
  one build for the windows (`platform: browser`) and one for the main
  process (`platform: node`). esbuild does not follow
  `new URL('./x.js', import.meta.url)`, which is how the main process and
  the engines start their worker threads, so the map reads those itself; a
  `new URL` that names a folder or a data file (text recognition's
  `../vendor/`) is kept as a run-time read of that folder.
- **The entry points** are read from the wiring rather than listed:
  `APP_LOADERS` in `renderer/index.js` gives each app's window module, and
  the `namespaces` function in `main/main.js` gives each IPC namespace and
  the module whose factory builds it. A few are named by hand
  (`EXTRA_ENTRIES`): the two window workers `build/bundle.js` builds on
  their own, the AI Tools window (`renderer/ai-tools.js`), the text
  recognition window code (`renderer/ocr-ui.js`) and the AI tools'
  program run as Node (`main/mcp/node.js`).
- **Every file has a home.** In the renderer, `apps/<app>.js` and
  `apps/<app>/` are that app's, the rest the core's. In the main process a
  file's name decides (`MAIN_HOMES`: `mail*.js` and `discover.js` are
  Mail's, `base*.js` and `engines/` Data Files', `mcp/` and `ai-tools.js` AI
  Tools', `live.js` Recorder's, `present.js` and `legacy-deck*.js`
  Presentations', `mailmerge-source.js` Documents'). The window checks
  (`verify-*.js`, the smoke and screens runs) are their own home and never
  walked.
- **A walk does not enter another app's files.** The import that would is a
  cross-app import, the thing to cut. The core's walk (from
  `renderer/index.js`, `main/main.js` and the preload) enters no app at
  all, so what the core reaches it reaches on its own.
- **Owners.** An engine (a package under `packages/`) or an npm package one
  pack alone reaches goes into that pack; one the core or two packs reach
  stays in the core. A shared file in the renderer or the main process that
  one app alone reaches moves into it. Two engines are split by file
  (`SPLIT_PACKAGES`), because the map shows they divide cleanly:
  `@rutba/ocr` (the engine and its 15 MB of data are text recognition's;
  the result helpers four apps' windows read are small and shared) and
  `@rutba/templates` (Data Files' databases are Data Files'; the gallery is
  the core's).
- **IPC.** A namespace's owner is named in `NAMESPACE_PACK`; a namespace not
  named is the core's. Its channels are the contract's methods
  (`packages/office-shell/src/contract.js`), `rutba-office/<ns>:<method>`.
  The map also reads each app's window code for calls into another pack's
  namespace (`shell.mail.send(...)`) and for windows of another app it
  opens (`win.create({ app: 'mail' })`). Those are not imports and need no
  cut, but an absent pack must not be offered.
- **Sizes** are bytes as shipped, before the installer's compression: the
  window code minified and split as the build makes it (attributed per
  source file from esbuild's metafile), the main-process files, and the
  folders of the engines and npm packages the main process reaches (as
  electron-builder copies them, without tests, notes or source maps).

`tests/app-map.test.js` runs the walk (not the sizes) and compares it with
`tests/fixtures/app-map-edges.json`: a cross-app import that is not in the
fixture fails, and so does an engine or npm package changing owner (Mail's
`imapflow` reached by Calendar, say). The failure names the file and line.
A cut makes the fixture shorter; `--write-edges` writes it again.

### What the map shows

**The core** is the shell, the launcher, the document service the three
document apps share (`main/documents.js` and its thread), every engine that
reads and writes Office files (`@rutba/ooxml`, `@rutba/office-formats`,
`@rutba/doc-view`, `@rutba/sheet-view`, `@rutba/presentation`,
`@rutba/drawing`, `@rutba/formula`), spelling, printing, the updater, the
gallery and the words in other languages. It is most of the code: the
document apps' engines are reached by the core's own document service, by
AI Tools and by Data Files, so none of them can be one pack's. Documents,
Worksheets and Presentations are therefore small packs: their windows,
their checks and their file types. Splitting `main/documents.js` by app is
possible later; it is one 4,000-line service today, and the map counts it
as the core's.

**The packs with weight** are text recognition (its engine and language
data), Mail (its protocols, `imapflow` and `nodemailer`), Data Files (its
sample databases and the PostgreSQL and MySQL clients) and PDF (pdf.js).
The others are window code of a few hundred kilobytes.

**The cuts** fall into a few kinds, all in the windows except four:

1. *Shared window code living in one app's folder.* `word/dialogs.js` and
   `word/equations.js` (used by Worksheets and Presentations),
   `slides/design.js`, `slides/ink.js` and `slides/ink-geometry.js` (used by
   Documents, Worksheets and Image), `image/stage.js` (used by Video),
   `pictures/library.js` (used by Worksheets). These move to a shared
   folder of the core; nothing about them is one app's.
2. *Recording.* `video/recorder/` is Recorder's engine and Presentations'
   Cameo and screen recording too. With the `capture` namespace already the
   core's, the recording engine goes to the core with it, and Recorder and
   Video stop needing each other.
3. *Text recognition in the windows.* `renderer/ocr-ui.js` is imported by
   PDF, Pictures, Image, Worksheets and Mail. The window code is small and
   moves to the core; what it offers (Recognise Text, Copy Text from
   Picture, Data from Picture, a PDF attachment's note) shows only when the
   text recognition pack is installed, which the core knows from the
   loader.
4. *pdf.js outside PDF.* Mail's attachment note and Worksheets' receipt row
   read PDFs with `pdf/pdfjs.js` and `pdf/ocr.js`. Either pdf.js goes to the
   core (2.9 MB, and PDF becomes a small pack) or those two features read
   PDFs only when PDF is installed. The second keeps the core smaller and
   is the proposal.
5. *Pictures and Image.* Pictures' batch conversion imports Image's
   (`image/batch.js`). The batch code moves to the core or Pictures offers
   batch conversion only with Image installed.
6. *The core reaching an app.* The document service imports
   `mailmerge-source.js` (Documents' mail merge, which reads workbooks,
   CSV and Data Files databases) and `legacy-deck.js` (Presentations' old
   `.ppt` reader); both belong to the document service and move into the
   core, which then keeps `@rutba/database` reachable only through mail
   merge. `main.js` uses Data Files' trust list when a recent file is
   renamed (`base-trust.js`): a rename hook the Data Files pack registers.
   The shell's title bar imports the AI Tools window (`renderer/ai-tools.js`):
   the badge and the menu item come from the AI Tools pack when it is there.
   AI Tools' tools call text recognition's service directly
   (`mcp/make-tools.js`): the tool is listed only when that pack is
   installed.

**What the map found besides the cuts:**

- *Every IPC channel exists whatever is installed.* `installIpc`
  (`packages/office-shell/src/electron/ipc.js`) registers a handler for
  every method in the contract's `METHODS`, and the preload exposes every
  one; a namespace with no service answers "no backend". For "no IPC
  channel" the contract must become the core's methods plus each loaded
  pack's own, carried in its manifest (section 2), and the preload must be
  told the list at run time rather than built with it.
- *The registry is one table.* `APPS`, `APP_EXTENSIONS`, `WRITABLE`,
  `OLD_EXTENSIONS` and the New templates in
  `packages/office-formats/src/registry.js` name every app. The core keeps
  the table's shape; each pack's manifest supplies its rows.
- *File types are registered by the installer for every app*
  (`fileAssociations` in `electron-builder.config.cjs`), then trimmed by
  `main/defaults.js` to the apps chosen. With packs, the installer
  registers only the types of the packs it copies.
- *Calendar and Contacts share one sync service* (`dav`, `dav-sync.js`, and
  `renderer/dav-accounts.js`); `@rutba/contacts` itself depends on
  `@rutba/calendar`. The map puts the sync and both small engines in the
  core (about 100 KB) so neither app needs the other.
- *`mdb-reader`* (3 MB) is reached by the core's document service
  (Worksheets opening `.mdb` and `.accdb`) and by Data Files, so it is the
  core's. If Worksheets' Access reading moves into the Worksheets pack it
  becomes a dependency shared by two packs, which the core then holds; it
  stays in the core either way.
- *What the archive carries and nothing in the main process reaches*: the
  folders of engines the windows alone use (their code is in the bundle)
  and `mailparser`, which nothing imports. Listed at the end of the tables.

<!-- app-map:begin -->
### The packs

Sizes as shipped before compression: the window code minified as the build makes it, the main-process files, the engines and npm packages only that pack reaches (as folders, without tests and notes) and its other files. "Imports" are the imports to cut (below); "calls" and "opens" are the other packs an app reaches at run time, through IPC or by opening a window.

| Pack | Size | Windows | Main | Engines | npm | Other | Imports to cut, to | Calls | Opens | File types it claims | IPC namespaces (channels) | AI tools | Checks |
|---|---:|---:|---:|---:|---:|---:|---|---|---|---|---|---|---:|
| Documents (`word`) | 482 KB | 482 KB | 0 KB | 0 KB | 0 KB | 0 KB | Presentations (2) | Contacts, Mail | none | docx docm dotx doc dot wri odt rtf txt md markdown html htm | none (0) | create_document, export_pdf | 50 |
| Worksheets (`sheets`) | 351 KB | 351 KB | 0 KB | 0 KB | 0 KB | 0 KB | Documents (2), Presentations (3), Text recognition (3), Pictures (2), PDF (2) | Text recognition | none | xlsx xlsm xltx xls xlt ods csv tsv | none (0) | create_workbook | 30 |
| Presentations (`slides`) | 362 KB | 359 KB | 3 KB | 0 KB | 0 KB | 0 KB | Documents (2), Video (3) | none | none | pptx pptm potx ppsx ppt pps pot odp | present (2) | create_presentation | 49 |
| PDF (`pdf`) | 3.0 MB | 137 KB | 18 KB | 0 KB | 0 KB | 2.9 MB | Text recognition (3) | Text recognition | none | pdf | pdf (18) | combine_pdf_pages | 10 |
| Pictures (`pictures`) | 54 KB | 54 KB | 0 KB | 0 KB | 0 KB | 0 KB | Image (1), Text recognition (1) | none | Image, PDF, Video | jpg jpeg png gif webp avif heic bmp tif tiff ico svg | none (0) | none | 4 |
| Image (`image`) | 181 KB | 181 KB | 0 KB | 0 KB | 0 KB | 0 KB | Text recognition (1), Presentations (1) | none | none | rimg | none (0) | none | 7 |
| Video (`video`) | 225 KB | 225 KB | 0 KB | 0 KB | 0 KB | 0 KB | Image (1) | none | none | rvid mp4 m4v mov webm mkv ogv mp3 m4a wav flac ogg | none (0) | none | 9 |
| Recorder (`recorder`) | 85 KB | 49 KB | 10 KB | 26 KB | 0 KB | 0 KB | Video (4) | none | Image, Pictures, Video | none | live (8) | none | 4 |
| Mail (`mail`) | 5.6 MB | 228 KB | 214 KB | 351 KB | 4.8 MB | 0 KB | Text recognition (1), PDF (2) | Calendar, Contacts, Text recognition | none | eml emlx msg mbox pst ost olm | mail, discover (75) | none | 16 |
| Calendar (`calendar`) | 62 KB | 45 KB | 17 KB | 0 KB | 0 KB | 0 KB | none | none | Mail | ics | calendar (12) | none | 2 |
| Contacts (`contacts`) | 32 KB | 19 KB | 13 KB | 0 KB | 0 KB | 0 KB | none | none | Mail | vcf | contacts (11) | none | 0 |
| Data Files (`base`) | 4.1 MB | 201 KB | 99 KB | 2.5 MB | 1.3 MB | 0 KB | none | none | none | rbase | base (84) | none | 19 |
| Talk (`talk`) | 78 KB | 43 KB | 35 KB | 0 KB | 0 KB | 0 KB | none | Contacts, Mail | Mail | none | talk (21) | none | 4 |
| AI Tools (`ai`) | 161 KB | 12 KB | 54 KB | 95 KB | 0 KB | 0 KB | Text recognition (1) | none | none | none | ai (6) | list_folders, list_files, read_file, list_templates, create_from_template | 1 |
| Text recognition (`ocr`) | 15.3 MB | 13 KB | 3 KB | 15.3 MB | 0 KB | 0 KB | none | none | none | none | ocr (4) | read_text_in_picture | 3 |
| **Core** | 22.4 MB | 5.5 MB | 427 KB | 10.9 MB | 5.6 MB | 0 KB | | | | | app, win, fs, thumbs, dialog, shell, store, doc, dav, update, defaults, oauth, announce, print, clipboard, capture (143) |  | 58 |

All together: 52.4 MB. The window checks (`main/verify-*.js` and the smoke and screens runs, 3.5 MB) are not in these sizes.

### The engines and who reaches them

An engine one pack alone reaches moves into that pack; one the core or two packs reach stays in the core. "Windows" means it reaches the windows only, where it is bundled; its folder in the archive is then not needed at run time.

| Engine | Owner | Size | Reached from | Reached by | First import, by owner |
|---|---|---:|---|---|---|
| @rutba/barcode | core | 27 KB | windows | Worksheets, Presentations, Talk, Documents | core: `apps/desktop/renderer/qrcode.js:10` |
| @rutba/calendar | core | 40 KB | main process | Calendar, Contacts, core | core: `apps/desktop/main/dav-sync.js:18`; Calendar: `apps/desktop/main/calendar.js:17` |
| @rutba/contacts | core | 22 KB | main process | Contacts, core | core: `apps/desktop/main/dav-sync.js:19`; Contacts: `apps/desktop/main/contacts.js:13` |
| @rutba/database | Data Files | 223 KB | main process, windows | Data Files | Data Files: `apps/desktop/main/base-sources.js:18` |
| @rutba/doc-view | core | 444 KB | main process, windows | AI Tools, core, Documents | Documents: `apps/desktop/renderer/apps/word.js:66`; core: `apps/desktop/main/documents.js:20`; AI Tools: `apps/desktop/main/mcp/make-tools.js:7` |
| @rutba/drawing | core | 145 KB | main process, windows | AI Tools, Data Files, core, Worksheets, Presentations, Documents | core: `apps/desktop/main/documents.js:50`; Documents: `apps/desktop/renderer/apps/word/drawings.js:24`; Worksheets: `apps/desktop/renderer/apps/sheets/ribbon.js:12`; Presentations: `apps/desktop/renderer/apps/slides/ribbon.js:13` |
| @rutba/editing | core | 11 KB | main process, windows | AI Tools, Data Files, core, PDF, Video | Video: `apps/desktop/renderer/apps/video/index.js:18`; PDF: `apps/desktop/renderer/apps/pdf/index.js:51`; core: `apps/desktop/main/documents.js:16` |
| @rutba/formula | core | 239 KB | main process, windows | AI Tools, Data Files, core, Worksheets | Worksheets: `apps/desktop/renderer/apps/sheets/ribbon.js:13`; core: `packages/ooxml/src/pivot.js:33` |
| @rutba/imaging | core | 124 KB | main process, windows | core, Image, Pictures, Presentations, Documents | core: `apps/desktop/main/documents.js:46`; Documents: `apps/desktop/renderer/apps/word/ribbon.js:35`; Presentations: `apps/desktop/renderer/apps/slides/ribbon.js:16`; Pictures: `apps/desktop/renderer/apps/pictures.js:244`; Image: `apps/desktop/renderer/apps/image.js:18` |
| @rutba/mailbox | Mail | 351 KB | main process, windows | Mail | Mail: `apps/desktop/main/mail-archive.js:27` |
| @rutba/media | core | 80 KB | windows | Pictures, Video | Pictures: `apps/desktop/renderer/apps/pictures.js:245`; Video: `apps/desktop/renderer/apps/video/export.js:18` |
| @rutba/ocr | split: Text recognition 15.3 MB, core 82 KB | 15.4 MB | main process, windows | AI Tools, Mail, Text recognition, PDF, Worksheets | Worksheets: `apps/desktop/renderer/apps/sheets/receipt-row.js:16`; PDF: `apps/desktop/renderer/apps/pdf/found.js:16`; Mail: `apps/desktop/renderer/apps/mail/docnote.js:16`; Text recognition: `apps/desktop/main/ocr.js:12`; AI Tools: `apps/desktop/main/mcp/make-tools.js:9` |
| @rutba/office-formats | core | 837 KB | main process, windows | AI Tools, Data Files, Calendar, Contacts, core, Image, Mail, PDF, Pictures, Recorder, Worksheets, Presentations, Talk, Video, Documents | core: `apps/desktop/main/announce.js:33`; Documents: `apps/desktop/renderer/apps/word.js:21`; Worksheets: `apps/desktop/renderer/apps/sheets.js:29`; Presentations: `apps/desktop/renderer/apps/slides.js:19`; Mail: `apps/desktop/renderer/apps/mail.js:27`; AI Tools: `apps/desktop/main/mcp/tools.js:19`; Data Files: `apps/desktop/main/base-sources.js:13` |
| @rutba/office-shell | core | 174 KB | main process, windows | Data Files, Calendar, Contacts, core, Image, Mail, PDF, Pictures, Recorder, Worksheets, Presentations, Talk, Video, Documents | core: `apps/desktop/main/main.js:13`; Mail: `apps/desktop/renderer/apps/mail/compose.js:19`; Recorder: `apps/desktop/main/live.js:23` |
| @rutba/office-ui | core | 4.0 MB | main process, windows | AI Tools, Data Files, Calendar, Contacts, core, Image, Mail, Text recognition, PDF, Pictures, Recorder, Worksheets, Presentations, Talk, Video, Documents | core: `apps/desktop/renderer/apps-choice.js:8`; Documents: `apps/desktop/renderer/apps/word.js:17`; Worksheets: `apps/desktop/renderer/apps/sheets.js:15`; Presentations: `apps/desktop/renderer/apps/slides.js:13`; Pictures: `apps/desktop/renderer/apps/pictures.js:25`; Image: `apps/desktop/renderer/apps/image.js:17`; Video: `apps/desktop/renderer/apps/video/export.js:17`; Recorder: `apps/desktop/renderer/apps/recorder.js:20`; Talk: `apps/desktop/renderer/apps/talk.js:14`; Data Files: `apps/desktop/renderer/apps/base.js:9`; PDF: `apps/desktop/renderer/apps/pdf/annots.js:18`; Mail: `apps/desktop/renderer/apps/mail.js:25`; Calendar: `apps/desktop/renderer/apps/calendar.js:15`; Contacts: `apps/desktop/renderer/apps/contacts.js:15`; AI Tools: `apps/desktop/renderer/ai-tools.js:12`; Text recognition: `apps/desktop/renderer/ocr-ui.js:12` |
| @rutba/ooxml | core | 1.3 MB | main process, windows | AI Tools, Data Files, core, Mail, Worksheets, Presentations, Documents | Documents: `apps/desktop/renderer/apps/word.js:45`; core: `apps/desktop/main/documents.js:21`; Worksheets: `apps/desktop/renderer/apps/sheets.js:24`; Presentations: `apps/desktop/renderer/apps/slides.js:31`; Mail: `packages/mailbox/src/olm.js:12`; AI Tools: `apps/desktop/main/mcp/tools.js:14`; Data Files: `apps/desktop/main/base-sources.js:17` |
| @rutba/pdf | core | 342 KB | main process | AI Tools, core, Mail, PDF | core: `apps/desktop/main/system-fonts.js:10`; PDF: `apps/desktop/main/pdf.js:43`; AI Tools: `apps/desktop/main/mcp/tools.js:21` |
| @rutba/presentation | core | 605 KB | main process, windows | AI Tools, core, Presentations | Presentations: `apps/desktop/renderer/apps/slides.js:28`; core: `apps/desktop/main/documents.js:40`; AI Tools: `apps/desktop/main/mcp/tools.js:16` |
| @rutba/proofing | core | 204 KB | main process | core | core: `apps/desktop/main/proofing.js:29` |
| @rutba/sheet-view | core | 689 KB | main process, windows | AI Tools, Data Files, core, Worksheets | Worksheets: `apps/desktop/renderer/apps/sheets.js:53`; core: `apps/desktop/main/documents.js:19`; AI Tools: `packages/templates/src/sheets.js:12`; Data Files: `apps/desktop/main/base-sources.js:12` |
| @rutba/stream | Recorder | 26 KB | main process, windows | Recorder | Recorder: `apps/desktop/main/live.js:21` |
| @rutba/studio | core | 901 KB | windows | Image, Recorder, Presentations, Video | Presentations: `apps/desktop/renderer/apps/slides.js:46`; Image: `apps/desktop/renderer/apps/image.js:21`; Video: `apps/desktop/renderer/apps/video/export.js:21`; Recorder: `apps/desktop/renderer/apps/recorder.js:21` |
| @rutba/templates | split: Data Files 2.3 MB, core 1.7 MB, AI Tools 95 KB | 4.1 MB | main process | AI Tools, Data Files, core | core: `apps/desktop/main/documents.js:97`; AI Tools: `apps/desktop/main/mcp/tools.js:17`; Data Files: `apps/desktop/main/base.js:20` |

| npm package (main process) | Owner | Size with what it needs | Reached by |
|---|---|---:|---|
| dictionary-en | core | 559 KB | core |
| dictionary-en-gb | core | 559 KB | core |
| electron-updater | core | 1.1 MB | core |
| imapflow | Mail | 3.7 MB | Mail |
| mdb-reader | core | 2.9 MB | Data Files, core |
| mysql2 | Data Files | 1.4 MB | Data Files |
| nodemailer | Mail | 1.5 MB | Mail |
| nspell | core | 32 KB | core |
| pg | Data Files | 329 KB | Data Files |

### The heavy parts

| What | Where | Size | Owner |
|---|---|---:|---|
| Text recognition's engine and language data | `packages/ocr/vendor` | 15.2 MB | Text recognition |
| pdf.js, which draws PDF pages (copied beside the bundle) | `apps/desktop/vendor/pdfjs` | 2.9 MB | PDF |
| Data Files' ready-made and sample databases | `packages/templates/data-files` | 2.3 MB | Data Files |
| The template gallery: files and thumbnails | `packages/templates/files, packages/templates/thumbs` | 1.6 MB | core |
| The words in Urdu, Arabic and Hindi | `packages/office-ui/src/catalogues` | 3.8 MB | core |
| The studio renderer (layers for Video, Image and Recorder) | `packages/studio/src` | 616 KB | core |
| Video's containers and export plan (the codecs are Chromium's own WebCodecs; none ship) | `packages/media/src` | 79 KB | core |
| AI Tools' server (the Model Context Protocol, local and over the network) | `apps/desktop/main/mcp` | 46 KB | AI Tools |
| Talk's calling and messages | `apps/desktop/renderer/apps/talk, apps/desktop/renderer/apps/talk.js, apps/desktop/main/talk.js, apps/desktop/main/talk-signin.js` | 96 KB | Talk |
| Spelling dictionaries | `dictionary-en, dictionary-en-gb` | 1.1 MB | core |
| Database engine clients | `pg, mysql2` | 1.7 MB | Data Files |
| Mail protocols | `imapflow, nodemailer` | 5.2 MB | Mail |

### In the archive, never reached by the main process

electron-builder copies the folder of every dependency `apps/desktop/package.json` names. These the main process never imports: either the windows' bundle already holds what they use, or nothing uses them. Neither the core nor a pack needs their folders.

| Dependency | Size | Why |
|---|---:|---|
| @rutba/barcode | 27 KB | the windows only: bundled |
| @rutba/media | 80 KB | the windows only: bundled |
| @rutba/studio | 901 KB | the windows only: bundled |
| mailparser | 4.8 MB | reached by nothing |

### The imports to cut (39)

Each is a file of one owner importing a file of another app. Before the move each is cut (the app asks through IPC or an event, or the shared part moves to the core). Grouped by who imports whom.

| From | To | Where | Imports | Kind |
|---|---|---|---|---|
| AI Tools | Text recognition | `apps/desktop/main/mcp/make-tools.js:12` | `apps/desktop/main/ocr.js` | static |
| core | AI Tools | `apps/desktop/renderer/shell.js:16` | `apps/desktop/renderer/ai-tools.js` | static |
| core | Data Files | `apps/desktop/main/main.js:48` | `apps/desktop/main/base-trust.js` | static |
| core | Presentations | `apps/desktop/main/documents.js:87` | `apps/desktop/main/legacy-deck.js` | static |
| core | Documents | `apps/desktop/main/documents.js:76` | `apps/desktop/main/mailmerge-source.js` | static |
| Image | Text recognition | `apps/desktop/renderer/apps/image.js:50` | `apps/desktop/renderer/ocr-ui.js` | static |
| Image | Presentations | `apps/desktop/renderer/apps/image/type-editor.js:19` | `apps/desktop/renderer/apps/slides/design.js` | static |
| Mail | Text recognition | `apps/desktop/renderer/apps/mail/docnote.js:17` | `apps/desktop/renderer/ocr-ui.js` | static |
| Mail | PDF | `apps/desktop/renderer/apps/mail/docnote.js:18` | `apps/desktop/renderer/apps/pdf/pdfjs.js` | static |
| Mail | PDF | `apps/desktop/renderer/apps/mail/docnote.js:19` | `apps/desktop/renderer/apps/pdf/ocr.js` | static |
| PDF | Text recognition | `apps/desktop/renderer/apps/pdf/found.js:17` | `apps/desktop/renderer/ocr-ui.js` | static |
| PDF | Text recognition | `apps/desktop/renderer/apps/pdf/index.js:61` | `apps/desktop/renderer/ocr-ui.js` | static |
| PDF | Text recognition | `apps/desktop/renderer/apps/pdf/ocr.js:18` | `apps/desktop/renderer/ocr-ui.js` | static |
| Pictures | Image | `apps/desktop/renderer/apps/pictures.js:37` | `apps/desktop/renderer/apps/image/batch.js` | static |
| Pictures | Text recognition | `apps/desktop/renderer/apps/pictures.js:38` | `apps/desktop/renderer/ocr-ui.js` | static |
| Recorder | Video | `apps/desktop/renderer/apps/recorder.js:24` | `apps/desktop/renderer/apps/video/recorder/session.js` | static |
| Recorder | Video | `apps/desktop/renderer/apps/recorder.js:25` | `apps/desktop/renderer/apps/video/recorder/panel.js` | static |
| Recorder | Video | `apps/desktop/renderer/apps/recorder.js:26` | `apps/desktop/renderer/apps/video/recorder/css.js` | static |
| Recorder | Video | `apps/desktop/renderer/apps/recorder/live-dialog.js:12` | `apps/desktop/renderer/apps/video/recorder/panel.js` | static |
| Worksheets | Text recognition | `apps/desktop/renderer/apps/sheets.js:45` | `apps/desktop/renderer/ocr-ui.js` | static |
| Worksheets | Text recognition | `apps/desktop/renderer/apps/sheets/picture-data.js:13` | `apps/desktop/renderer/ocr-ui.js` | static |
| Worksheets | Text recognition | `apps/desktop/renderer/apps/sheets/receipt-row.js:17` | `apps/desktop/renderer/ocr-ui.js` | static |
| Worksheets | PDF | `apps/desktop/renderer/apps/sheets/receipt-row.js:19` | `apps/desktop/renderer/apps/pdf/pdfjs.js` | static |
| Worksheets | PDF | `apps/desktop/renderer/apps/sheets/receipt-row.js:20` | `apps/desktop/renderer/apps/pdf/ocr.js` | static |
| Worksheets | Pictures | `apps/desktop/renderer/apps/sheets/picture-data.js:14` | `apps/desktop/renderer/apps/pictures/library.js` | static |
| Worksheets | Pictures | `apps/desktop/renderer/apps/sheets/receipt-row.js:18` | `apps/desktop/renderer/apps/pictures/library.js` | static |
| Worksheets | Presentations | `apps/desktop/renderer/apps/sheets.js:38` | `apps/desktop/renderer/apps/slides/design.js` | static |
| Worksheets | Presentations | `apps/desktop/renderer/apps/sheets.js:46` | `apps/desktop/renderer/apps/slides/ink.js` | static |
| Worksheets | Presentations | `apps/desktop/renderer/apps/sheets/ribbon.js:18` | `apps/desktop/renderer/apps/slides/ink-geometry.js` | static |
| Worksheets | Documents | `apps/desktop/renderer/apps/sheets.js:30` | `apps/desktop/renderer/apps/word/dialogs.js` | static |
| Worksheets | Documents | `apps/desktop/renderer/apps/sheets.js:47` | `apps/desktop/renderer/apps/word/equations.js` | static |
| Presentations | Video | `apps/desktop/renderer/apps/slides.js:45` | `apps/desktop/renderer/apps/video/recorder/panel.js` | static |
| Presentations | Video | `apps/desktop/renderer/apps/slides/cameo.js:13` | `apps/desktop/renderer/apps/video/recorder/camera.js` | static |
| Presentations | Video | `apps/desktop/renderer/apps/slides/screen-record.js:7` | `apps/desktop/renderer/apps/video/recorder/session.js` | static |
| Presentations | Documents | `apps/desktop/renderer/apps/slides.js:36` | `apps/desktop/renderer/apps/word/equations.js` | static |
| Presentations | Documents | `apps/desktop/renderer/apps/slides.js:37` | `apps/desktop/renderer/apps/word/dialogs.js` | static |
| Video | Image | `apps/desktop/renderer/apps/video/stage.js:24` | `apps/desktop/renderer/apps/image/stage.js` | static |
| Documents | Presentations | `apps/desktop/renderer/apps/word.js:59` | `apps/desktop/renderer/apps/slides/ink.js` | static |
| Documents | Presentations | `apps/desktop/renderer/apps/word/ribbon.js:32` | `apps/desktop/renderer/apps/slides/ink-geometry.js` | static |

### The core's wiring (24)

What `renderer/index.js` and `main/main.js` import from each app today. These are not cut: the loader takes them over, one pack at a time.

| Pack | Where | Imports |
|---|---|---|
| Mail | `apps/desktop/main/main.js:31` | `apps/desktop/main/mail.js` |
| Presentations | `apps/desktop/main/main.js:35` | `apps/desktop/main/present.js` |
| Mail | `apps/desktop/main/main.js:38` | `apps/desktop/main/discover.js` |
| Contacts | `apps/desktop/main/main.js:39` | `apps/desktop/main/contacts.js` |
| Calendar | `apps/desktop/main/main.js:40` | `apps/desktop/main/calendar.js` |
| PDF | `apps/desktop/main/main.js:42` | `apps/desktop/main/pdf.js` |
| Text recognition | `apps/desktop/main/main.js:43` | `apps/desktop/main/ocr.js` |
| AI Tools | `apps/desktop/main/main.js:44` | `apps/desktop/main/ai-tools.js` |
| Recorder | `apps/desktop/main/main.js:45` | `apps/desktop/main/live.js` |
| Talk | `apps/desktop/main/main.js:46` | `apps/desktop/main/talk.js` |
| Data Files | `apps/desktop/main/main.js:47` | `apps/desktop/main/base.js` |
| Documents | `apps/desktop/renderer/index.js:22` | `apps/desktop/renderer/apps/word.js` |
| Worksheets | `apps/desktop/renderer/index.js:23` | `apps/desktop/renderer/apps/sheets.js` |
| Presentations | `apps/desktop/renderer/index.js:24` | `apps/desktop/renderer/apps/slides.js` |
| Pictures | `apps/desktop/renderer/index.js:25` | `apps/desktop/renderer/apps/pictures.js` |
| Image | `apps/desktop/renderer/index.js:26` | `apps/desktop/renderer/apps/image.js` |
| Video | `apps/desktop/renderer/index.js:27` | `apps/desktop/renderer/apps/video/index.js` |
| Recorder | `apps/desktop/renderer/index.js:28` | `apps/desktop/renderer/apps/recorder.js` |
| Talk | `apps/desktop/renderer/index.js:29` | `apps/desktop/renderer/apps/talk.js` |
| Data Files | `apps/desktop/renderer/index.js:30` | `apps/desktop/renderer/apps/base.js` |
| PDF | `apps/desktop/renderer/index.js:31` | `apps/desktop/renderer/apps/pdf/index.js` |
| Mail | `apps/desktop/renderer/index.js:32` | `apps/desktop/renderer/apps/mail.js` |
| Calendar | `apps/desktop/renderer/index.js:33` | `apps/desktop/renderer/apps/calendar.js` |
| Contacts | `apps/desktop/renderer/index.js:34` | `apps/desktop/renderer/apps/contacts.js` |

### Shared files one app alone uses (1)

Files in the shared folders that only one app reaches: they move into its pack.

| File | Moves into |
|---|---|
| `apps/desktop/renderer/defaults.js` | Mail |

### Calls and windows across packs

Not imports, so nothing to cut, but a pack that is absent must not be offered: the button that would call it or open its window is shown only when it is installed.

| Pack | Calls (namespace: first place) | Opens a window of |
|---|---|---|
| Documents | Contacts (3): `apps/desktop/renderer/apps/word/envelopes.js:48 contacts.list`; Mail (2): `apps/desktop/renderer/apps/word/mailings.js:850 mail.accounts` | none |
| Worksheets | Text recognition (4): `apps/desktop/renderer/apps/sheets/receipt-row.js:83 ocr.cancel` | none |
| PDF | Text recognition (1): `apps/desktop/renderer/apps/pdf/index.js:566 ocr.cancel` | none |
| Pictures | none | Image: `apps/desktop/renderer/apps/pictures.js:628`; Video: `apps/desktop/renderer/apps/pictures.js:630`; PDF: `apps/desktop/renderer/apps/pictures.js:632` |
| Recorder | none | Video: `apps/desktop/renderer/apps/recorder.js:98`; Pictures: `apps/desktop/renderer/apps/recorder.js:615`; Image: `apps/desktop/renderer/apps/recorder.js:616` |
| Mail | Calendar (3): `apps/desktop/renderer/apps/mail.js:1049 calendar.openFile`; Contacts (2): `apps/desktop/renderer/apps/mail.js:1102 contacts.fromMail`; Text recognition (1): `apps/desktop/renderer/apps/mail/docnote.js:117 ocr.cancel` | none |
| Calendar | none | Mail: `apps/desktop/renderer/apps/calendar.js:353` |
| Contacts | none | Mail: `apps/desktop/renderer/apps/contacts.js:217` |
| Talk | Mail (2): `apps/desktop/renderer/apps/talk/invite.js:88 mail.accounts`; Contacts (1): `apps/desktop/renderer/apps/talk/invite.js:98 contacts.suggest` | Mail: `apps/desktop/renderer/apps/talk/invite.js:191` |
<!-- app-map:end -->

## 2. The pack format

A pack is three things side by side in the core's packs folder
(`resources/packs/` beside `app.asar` in an installed copy):

```
mail.asar             the pack's files: an Electron archive
mail.asar.unpacked/   the files the operating system reads from disk (tiles), when there are any
mail.pack.json        the signed manifest
```

**Why an Electron archive (asar).** `@electron/asar` is already in the
tree: electron-builder uses it to make `app.asar`, and the lockfile has it
(3.4.1) as electron-builder's dependency. It is used only when packs are
made, at build time. At run time Electron reads into an `.asar` by path, so
a pack's main module is imported straight from its archive, its neighbours
import each other, and a file beside it is read by `import.meta.url`
(checked under Electron 44 with two packs made by `tools/make-packs.mjs`:
both imported and answered from inside their archives, and a spoiled one
was refused). The format is a JSON table followed by the files' bytes, so
the core reads the table and hashes the files itself (`readArchive` in the
loader, no dependency at run time). A zip would need unpacking into a
writable place before any code ran; a plain folder would make a partial or
mixed copy hard to see and slow to check (on Windows, opening a file is
what costs, and a pack of hundreds of files is then hundreds of opens;
the archive is one). A pack's own `package.json`, at the root of its
archive, says `"type": "module"`.

**The manifest file** (`<id>.pack.json`):

```json
{
  "format": "rutba-office-pack",
  "keyId": "3f0c2a9b81d4e6f7",
  "manifest": { "format": 1, "id": "mail", "name": "Mail", "version": "...", "...": "..." },
  "signature": "<Ed25519 signature of the manifest, base64>"
}
```

| Field of `manifest` | What it is |
|---|---|
| `format` | 1. A core refuses a format it does not read. |
| `id` | The app's key, as the registry and the windows use it (`mail`, `word`, `base`); the pack's files are named by it. |
| `name` | The name shown in a refusal and in Add or remove apps. |
| `version` | The version it was built with. It must equal the core's. |
| `needs` | Packs that must be installed and good for this one to load (refused otherwise, with their names). |
| `uses` | Packs whose features this one offers when they are installed, and hides when not (a PDF attachment's note in Mail uses PDF and text recognition). Never refused for. |
| `archive` | `<id>.asar`, in the same folder. |
| `main` | The IPC entry module in the archive. It exports `register(core)` and answers the namespaces it serves (section 5). |
| `windows` | The window module in the archive, which `renderer/index.js` imports for a window of this app. |
| `ipc` | `{ namespace: [method, ...] }`: the channels this pack adds. The core registers them, and the preload's bridge offers them, only when the pack is loaded. |
| `events` | The events it sends (`mail:new`). |
| `tiles` | `[{ app, icon }]`: the taskbar tile and the file types' icon, unpacked, because Windows reads them from disk. |
| `fileTypes` | The registry rows it adds: `ext`, `name`, `mimeType`, `role`, `icon`. |
| `tools` | The AI tools it adds (`create_document`). |
| `files` | Every file in the archive and beside it: `{ path: { size, sha512, unpacked } }`, SHA-512 in base64, paths relative with forward slashes. |

**The signature** is Ed25519 (`crypto.sign(null, ...)` in `node:crypto`)
over the manifest as canonical JSON: keys in order at every level, no
spaces, a key whose value is undefined left out, UTF-8. `canonical()` in
the loader is the one definition; the signer imports it. The whole of the
manifest is signed, the files' hashes with it, so a changed file shows as
a hash that differs and a changed manifest as a signature that fails.
`keyId` (the first 16 hex digits of the SHA-256 of the public key's SPKI
form) is a label for people; the core does not trust it and tries each key
it holds, so it can carry the next key before the last is retired.

**The rules the loader keeps** (section 4): the manifest is believed only
once its signature holds; the id is the file's name; the version is the
core's; every path is relative, with no `..`; the archive is in the same
folder; the archive holds exactly the files listed, no more, no fewer, and
no links; the unpacked folder holds exactly the unpacked files listed;
every file has its listed size and SHA-512; `main` and `windows` are files
of the pack; and everything in `needs` loaded too.

## 3. Making and signing packs

```
node tools/make-pack-key.mjs <folder outside the repository>
node tools/make-packs.mjs --tree <built tree> --out <folder> [--key <file>] [--version <v>] [--only mail,calendar]
```

`tools/make-pack-key.mjs` makes an Ed25519 key pair: `pack-key.pem`, the
private half (PKCS #8, written readable only by its owner where the system
allows) and `pack-key.pub.pem`, the public half (SPKI). It refuses a folder
inside the repository and never overwrites a key. It prints the public half
and its id, never the private one. The real key is made once and kept with
the other release secrets; its public half is the one the core carries
(a file in the core, public, committed with the move).

`tools/make-packs.mjs` reads a built tree with one folder per pack, named by
its id. In each, `pack.json` says what the pack is, and every other file is
the pack:

```json
{
  "name": "Mail",
  "needs": [],
  "uses": ["calendar", "contacts", "pdf", "ocr"],
  "main": "main/index.js",
  "windows": "renderer/mail.js",
  "unpack": ["tiles"],
  "ipc": { "mail": ["accounts", "list", "send"], "discover": ["scan"] },
  "events": ["mail:new", "mail:progress"],
  "tiles": [{ "app": "mail", "icon": "tiles/mail.ico" }],
  "fileTypes": [{ "ext": "eml", "name": "Email Message", "mimeType": "message/rfc822", "role": "Editor", "icon": "tiles/filetypes/mail.ico" }],
  "tools": []
}
```

For each it writes the archive with `@electron/asar` (the `unpack` folders
beside it), reads the archive back and checks it holds what was listed,
hashes every file, and writes the signed `<id>.pack.json`. The version is
the repository's `package.json` version unless `--version` names another.
The private key comes from `--key <file>`, or `RUTBA_PACK_KEY` (the PEM
text), or `RUTBA_PACK_KEY_FILE`; a key file inside the repository is refused
before it is read. At the move, `@electron/asar` is named in the root
`devDependencies` at the version the lockfile already has, so the build
does not lean on how npm hoists electron-builder's dependencies.

`tests/packs.test.js` makes a key pair in memory, a small built tree in a
temporary folder (Contacts; Mail, which needs Contacts, with a tile kept on
disk and a file past a megabyte; text recognition) and packs from it, and
throws all of it away. Nothing secret is in the repository.

## 4. The core's loader

`packages/office-shell/src/electron/packs.js`, not yet wired into
`main.js`:

```js
import { findPacks } from '@rutba/office-shell/electron/packs'; // export added with the move
const { packs, refused } = findPacks({ dir, publicKey, version });
```

- `dir`: the packs folder; none there means no packs and nothing refused.
- `publicKey`: the key packs must be signed with (a KeyObject, PEM or DER),
  or a list of keys.
- `version`: the core's own version.

It answers `packs`, in an order that puts what a pack needs before it, each
with what the core loads: `main` and `windows` (paths into the archive,
which `import()` takes under Electron), `ipc`, `events`, `tiles` (paths on
disk), `fileTypes`, `tools`, `needs` and `uses`. And `refused`, each with
the pack's id, its manifest file, a code (`unreadable`, `format`,
`signature`, `id`, `version`, `missing`, `unknown`, `changed`, `needs`) and a
sentence for a person, for example:

- The Mail pack is version 2.0.0 and this Rutba Office is 2.0.1. A pack
  must match the Rutba Office it runs in: update it with Add or remove apps.
- The mail pack is not signed by Rutba, or was changed after it was
  signed. It was not loaded.
- The Contacts pack's file main/index.js has been changed since it was
  signed. It may have been tampered with; install the pack again with Add
  or remove apps.
- The Mail pack holds a file its manifest does not list: main/payload.js.
  It may have been tampered with.
- The Mail pack was not loaded. It needs contacts, which is not installed:
  add it with Add or remove apps. (Or: which could not be loaded either.)

A refused pack's name comes from its file until its signature holds, so a
forged manifest cannot choose the words shown. A pack is refused in turn
when one it needs is refused. Under Electron the archive is read with
`original-fs`, which sees an `.asar` as the file it is; under Node (the
tests) with `node:fs`.

**Cost.** Every file of every installed pack is hashed at each start. The
archive is opened once per pack and read in megabyte pieces. Measured on
the development computer while other work ran: text recognition (28 files,
15.4 MB) and Mail (its engine, `imapflow` and `nodemailer`, 431 files,
3.6 MB) checked together in 210 to 280 ms. With every pack installed
(about 30 MB of packs) that is under half a second, run beside the first
window's start rather than before it, since no pack's IPC is needed until
its window asks. If it ever matters, a pack checked once can be trusted by
the archive's size and time until either changes; the full check stays for
after an update and for Add or remove apps.

**The tests** (`tests/packs.test.js`): good packs load in order with paths
into their archives; no folder, no packs; a pack of another version refused,
naming both; a file changed inside the archive refused, and the pack that
needs it refused in turn; a file changed beside the archive refused; a pack
signed with another key refused, and either of two keys accepted; a
manifest edited after signing refused; a pack whose needs are missing
refused, naming them; a file not in the manifest refused, in the archive (an
archive made again with one more file) or beside it; a missing archive or
file refused; a signed manifest naming a path out of its folder, or another
id than its file's, refused; the private key never read from or written
into the repository; the loader's reading of an archive agrees with
Electron's own tool; the signed bytes do not depend on the order of keys.
