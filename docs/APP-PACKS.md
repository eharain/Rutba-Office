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
code into packs (5). Sections 1 to 4 are built and tested; section 5 is the
plan, and its first part, the core that loads packs, is being built: each
item says what is built. No app code has moved yet.

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

**The cuts** fall into a few kinds, all in the windows except four. Those
made are marked; the tables below list what is left.

1. *Shared window code living in one app's folder.* Made. The Symbol
   dialog (from `word/dialogs.js`, now `renderer/symbol-dialog.js`),
   `equations.js`, `design.js` with `markup.js`, `ink.js` and
   `ink-geometry.js` (from Documents' and Presentations' folders), the
   turned corner Video's stage keeps as Image's does
   (`renderer/turned-corner.js`, from `image/stage.js`) and the file
   addresses Worksheets shows a picture by (`renderer/file-urls.js`, from
   `pictures/library.js`) are the core's, at the top of `renderer/`.
2. *Recording.* `video/recorder/` is Recorder's engine and Presentations'
   Cameo and screen recording too. With the `capture` namespace already the
   core's, the recording engine goes to the core with it, and Recorder and
   Video stop needing each other.
3. *Text recognition in the windows.* Made. `renderer/ocr-ui.js` (imported
   by PDF, Pictures, Image, Worksheets and Mail) is the core's; its calls go
   to the `ocr` namespace, which says Text recognition is not installed
   when it is not.
4. *pdf.js outside PDF.* Made. Mail's attachment note and Worksheets'
   receipt row read PDFs, so the wrapper round pdf.js and the page pixels
   text recognition reads (`renderer/pdfjs.js`) are the core's, and pdf.js
   with them (2.9 MB). PDF is a small pack.
5. *Pictures and Image.* Made. Pictures' batch conversion is in Pictures'
   own folder (`pictures/batch.js`); the image worker and the window's side
   of it (`renderer/imaging/`), which Image and the batch conversion both
   use, are the core's.
6. *The core reaching an app.* Made, but for AI Tools. The document service
   imports `mailmerge-source.js` (Documents' mail merge, which reads
   workbooks, CSV and Data Files databases) and `legacy-deck.js`
   (Presentations' old `.ppt` reader): both are the core's, and mail merge
   reads a database with `@rutba/database/base`, the file reader alone,
   which is all of Data Files' engine the core keeps. `main.js` tells Data
   Files of a renamed recent file through a hook the pack registers. The
   shell's title bar imports the AI Tools window (`renderer/ai-tools.js`):
   the badge and the menu item come from the AI Tools pack when it is there.
   AI Tools' text recognition tool asks for the `ocr` service the core
   hands it (`ai-pack.js`), or, in the program AI apps run, loads the
   installed Text recognition pack itself (`mcp/node.js`); without it the
   tool says Text recognition is not installed.

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
  folders of engines the windows alone use (their code is in the bundle).
  Listed at the end of the tables. `mailparser`, which nothing imported,
  is no longer a dependency.

<!-- app-map:begin -->
### The packs

Sizes as shipped before compression: the window code minified as the build makes it, the main-process files, the engines and npm packages only that pack reaches (as folders, without tests and notes) and its other files. "Imports" are the imports to cut (below); "calls" and "opens" are the other packs an app reaches at run time, through IPC or by opening a window.

| Pack | Size | Windows | Main | Engines | npm | Other | Imports to cut, to | Calls | Opens | File types it claims | IPC namespaces (channels) | AI tools | Checks |
|---|---:|---:|---:|---:|---:|---:|---|---|---|---|---|---|---:|
| Documents (`word`) | 466 KB | 466 KB | 0 KB | 0 KB | 0 KB | 0 KB | none | Contacts, Mail | none | docx docm dotx doc dot wri odt rtf txt md markdown html htm | none (0) | create_document, export_pdf | 51 |
| Worksheets (`sheets`) | 353 KB | 353 KB | 0 KB | 0 KB | 0 KB | 0 KB | none | Text recognition | none | xlsx xlsm xltx xls xlt ods csv tsv | none (0) | create_workbook | 31 |
| Presentations (`slides`) | 343 KB | 340 KB | 3 KB | 0 KB | 0 KB | 0 KB | Video (3) | none | none | pptx pptm potx ppsx ppt pps pot odp | present (2) | create_presentation | 52 |
| PDF (`pdf`) | 155 KB | 137 KB | 19 KB | 0 KB | 0 KB | 0 KB | none | Text recognition | none | pdf | pdf (18) | combine_pdf_pages | 11 |
| Pictures (`pictures`) | 69 KB | 69 KB | 0 KB | 0 KB | 0 KB | 0 KB | none | none | Image, PDF, Video | jpg jpeg png gif webp avif heic bmp tif tiff ico svg | none (0) | none | 5 |
| Image (`image`) | 166 KB | 166 KB | 0 KB | 0 KB | 0 KB | 0 KB | none | none | none | rimg | none (0) | none | 8 |
| Video (`video`) | 225 KB | 225 KB | 0 KB | 0 KB | 0 KB | 0 KB | none | none | none | rvid mp4 m4v mov webm mkv ogv mp3 m4a wav flac ogg | none (0) | none | 9 |
| Recorder (`recorder`) | 85 KB | 49 KB | 11 KB | 26 KB | 0 KB | 0 KB | Video (4) | none | Image, Pictures, Video | none | live (8) | none | 4 |
| Mail (`mail`) | 5.7 MB | 238 KB | 275 KB | 377 KB | 4.8 MB | 0 KB | AI Tools (1) | Calendar, Contacts, Text recognition | none | eml emlx msg mbox pst ost olm | mail, discover (84) | none | 17 |
| Calendar (`calendar`) | 71 KB | 50 KB | 20 KB | 0 KB | 0 KB | 0 KB | none | none | Mail | ics | calendar (12) | none | 3 |
| Contacts (`contacts`) | 43 KB | 25 KB | 18 KB | 0 KB | 0 KB | 0 KB | none | none | Mail | vcf | contacts (12) | none | 0 |
| Data Files (`base`) | 4.0 MB | 202 KB | 104 KB | 2.3 MB | 1.3 MB | 0 KB | AI Tools (1) | AI Tools | none | rbase | base (85) | none | 20 |
| Talk (`talk`) | 80 KB | 43 KB | 37 KB | 0 KB | 0 KB | 0 KB | none | Contacts, Mail | Mail | none | talk (21) | none | 4 |
| AI Tools (`ai`) | 258 KB | 21 KB | 143 KB | 95 KB | 0 KB | 0 KB | none | none | none | none | ai (11) | list_folders, list_files, read_file, list_templates, create_from_template | 2 |
| Text recognition (`ocr`) | 15.3 MB | 0 KB | 3 KB | 15.3 MB | 0 KB | 0 KB | none | none | none | none | ocr (4) | read_text_in_picture | 3 |
| **Core** | 26.2 MB | 5.8 MB | 507 KB | 11.5 MB | 5.6 MB | 2.9 MB | | | | | app, win, fs, thumbs, dialog, shell, store, doc, dav, update, defaults, oauth, announce, print, profiles, packs, clipboard, capture (163) |  | 65 |

All together: 53.4 MB. The window checks (`main/verify-*.js` and the smoke and screens runs, 3.7 MB) are not in these sizes.

### The engines and who reaches them

An engine one pack alone reaches moves into that pack; one the core or two packs reach stays in the core. "Windows" means it reaches the windows only, where it is bundled; its folder in the archive is then not needed at run time.

| Engine | Owner | Size | Reached from | Reached by | First import, by owner |
|---|---|---:|---|---|---|
| @rutba/barcode | core | 27 KB | windows | Worksheets, Presentations, Talk, Documents | core: `apps/desktop/renderer/qrcode.js:10` |
| @rutba/calendar | core | 40 KB | main process, windows | AI Tools, Calendar, Contacts, core | Calendar: `apps/desktop/main/calendar.js:18`; core: `apps/desktop/main/dav-sync.js:18`; AI Tools: `apps/desktop/main/mcp/reach-tools.js:27` |
| @rutba/contacts | core | 29 KB | main process | Contacts, core | core: `apps/desktop/main/dav-sync.js:19`; Contacts: `apps/desktop/main/contacts.js:13` |
| @rutba/database | split: core 209 KB, Data Files 35 KB | 244 KB | main process, windows | Data Files, core | Data Files: `apps/desktop/main/base-sources.js:18`; core: `apps/desktop/main/mailmerge-source.js:18` |
| @rutba/doc-view | core | 445 KB | main process, windows | AI Tools, core, Documents | Documents: `apps/desktop/renderer/apps/word.js:66`; core: `apps/desktop/main/documents.js:20`; AI Tools: `apps/desktop/main/mcp/make-tools.js:7` |
| @rutba/drawing | core | 145 KB | main process, windows | AI Tools, Data Files, core, Worksheets, Presentations, Documents | core: `apps/desktop/main/documents.js:50`; Documents: `apps/desktop/renderer/apps/word/drawings.js:24`; Worksheets: `apps/desktop/renderer/apps/sheets/ribbon.js:12`; Presentations: `apps/desktop/renderer/apps/slides/ribbon.js:13` |
| @rutba/editing | core | 11 KB | main process, windows | AI Tools, Data Files, core, PDF, Video | Video: `apps/desktop/renderer/apps/video/index.js:18`; PDF: `apps/desktop/renderer/apps/pdf/index.js:51`; core: `apps/desktop/main/documents.js:16` |
| @rutba/formula | core | 239 KB | main process, windows | AI Tools, Data Files, core, Worksheets | Worksheets: `apps/desktop/renderer/apps/sheets/dialogs.js:11`; core: `packages/ooxml/src/pivot.js:33` |
| @rutba/imaging | core | 124 KB | main process, windows | core, Image, Pictures, Presentations, Documents | core: `apps/desktop/main/documents.js:46`; Documents: `apps/desktop/renderer/apps/word/ribbon.js:35`; Presentations: `apps/desktop/renderer/apps/slides/ribbon.js:16`; Pictures: `apps/desktop/renderer/apps/pictures.js:26`; Image: `apps/desktop/renderer/apps/image.js:18` |
| @rutba/mailbox | Mail | 377 KB | main process, windows | Mail | Mail: `apps/desktop/main/mail-archive.js:27` |
| @rutba/media | core | 80 KB | windows | Pictures, Video | Pictures: `apps/desktop/renderer/apps/pictures.js:27`; Video: `apps/desktop/renderer/apps/video/export.js:18` |
| @rutba/ocr | split: Text recognition 15.3 MB, core 89 KB | 15.4 MB | main process, windows | AI Tools, Image, Mail, Text recognition, PDF, Pictures, Worksheets | core: `apps/desktop/renderer/ocr-ui.js:13`; Worksheets: `apps/desktop/renderer/apps/sheets/receipt-row.js:16`; PDF: `apps/desktop/renderer/apps/pdf/found.js:16`; Mail: `apps/desktop/renderer/apps/mail/docnote.js:16`; AI Tools: `apps/desktop/main/mcp/make-tools.js:9`; Text recognition: `apps/desktop/main/ocr.js:12` |
| @rutba/office-formats | core | 837 KB | main process, windows | AI Tools, Data Files, Calendar, Contacts, core, Image, Mail, PDF, Pictures, Recorder, Worksheets, Presentations, Talk, Video, Documents | core: `apps/desktop/main/announce.js:33`; Documents: `apps/desktop/renderer/apps/word.js:21`; Worksheets: `apps/desktop/renderer/apps/sheets.js:29`; Presentations: `apps/desktop/renderer/apps/slides.js:19`; Mail: `apps/desktop/renderer/apps/mail.js:27`; Data Files: `apps/desktop/main/base-sources.js:13`; AI Tools: `apps/desktop/main/mcp/tools.js:24` |
| @rutba/office-shell | core | 295 KB | main process, windows | AI Tools, Data Files, Calendar, Contacts, core, Image, Mail, PDF, Pictures, Recorder, Worksheets, Presentations, Talk, Video, Documents | core: `apps/desktop/main/app-packs.js:24`; Mail: `apps/desktop/renderer/apps/mail/compose.js:19`; Recorder: `apps/desktop/main/live.js:23`; AI Tools: `apps/desktop/main/ai-tools.js:27` |
| @rutba/office-ui | core | 4.2 MB | main process, windows | AI Tools, Data Files, Calendar, Contacts, core, Image, Mail, PDF, Pictures, Recorder, Worksheets, Presentations, Talk, Video, Documents | core: `apps/desktop/renderer/apps-choice.js:8`; Documents: `apps/desktop/renderer/apps/word.js:17`; Worksheets: `apps/desktop/renderer/apps/sheets.js:15`; Presentations: `apps/desktop/renderer/apps/slides.js:13`; Pictures: `apps/desktop/renderer/apps/pictures.js:25`; Image: `apps/desktop/renderer/apps/image.js:17`; Video: `apps/desktop/renderer/apps/video/export.js:17`; Recorder: `apps/desktop/renderer/apps/recorder.js:20`; Talk: `apps/desktop/renderer/apps/talk.js:14`; Data Files: `apps/desktop/renderer/apps/base.js:9`; PDF: `apps/desktop/renderer/apps/pdf/annots.js:18`; Mail: `apps/desktop/renderer/apps/mail.js:25`; Calendar: `apps/desktop/renderer/apps/calendar.js:15`; Contacts: `apps/desktop/renderer/apps/contacts.js:15`; AI Tools: `apps/desktop/renderer/ai-tools.js:19` |
| @rutba/ooxml | core | 1.3 MB | main process, windows | AI Tools, Data Files, core, Mail, Worksheets, Presentations, Documents | Documents: `apps/desktop/renderer/apps/word.js:45`; core: `apps/desktop/main/documents.js:21`; Worksheets: `apps/desktop/renderer/apps/sheets.js:24`; Presentations: `apps/desktop/renderer/apps/slides.js:31`; Mail: `packages/mailbox/src/olm.js:12`; Data Files: `apps/desktop/main/base-sources.js:17`; AI Tools: `apps/desktop/main/mcp/tools.js:19` |
| @rutba/pdf | core | 344 KB | main process | AI Tools, core, PDF | core: `apps/desktop/main/system-fonts.js:10`; PDF: `apps/desktop/main/pdf.js:43`; AI Tools: `apps/desktop/main/mcp/tools.js:26` |
| @rutba/presentation | core | 609 KB | main process, windows | AI Tools, core, Presentations | Presentations: `apps/desktop/renderer/apps/slides.js:28`; core: `apps/desktop/main/documents.js:40`; AI Tools: `apps/desktop/main/mcp/tools.js:21` |
| @rutba/proofing | core | 204 KB | main process | core | core: `apps/desktop/main/proofing.js:29` |
| @rutba/sheet-view | core | 689 KB | main process, windows | AI Tools, Data Files, core, Worksheets | Worksheets: `apps/desktop/renderer/apps/sheets.js:53`; core: `apps/desktop/main/documents.js:19`; Data Files: `apps/desktop/main/base-sources.js:12`; AI Tools: `packages/templates/src/sheets.js:12` |
| @rutba/stream | Recorder | 26 KB | main process, windows | Recorder | Recorder: `apps/desktop/main/live.js:21` |
| @rutba/studio | core | 901 KB | windows | Image, Recorder, Presentations, Video | Presentations: `apps/desktop/renderer/apps/slides.js:46`; Image: `apps/desktop/renderer/apps/image.js:21`; core: `apps/desktop/renderer/turned-corner.js:4`; Video: `apps/desktop/renderer/apps/video/export.js:21`; Recorder: `apps/desktop/renderer/apps/recorder.js:21` |
| @rutba/templates | split: Data Files 2.3 MB, core 1.7 MB, AI Tools 95 KB | 4.1 MB | main process | AI Tools, Data Files, core | core: `apps/desktop/main/documents.js:97`; Data Files: `apps/desktop/main/base.js:20`; AI Tools: `apps/desktop/main/mcp/tools.js:22` |

| npm package (main process) | Owner | Size with what it needs | Reached by |
|---|---|---:|---|
| @tech-style/secure-keystore | core | 84 KB | AI Tools, core, Recorder |
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
| pdf.js, which draws PDF pages (copied beside the bundle) | `apps/desktop/vendor/pdfjs` | 2.9 MB | core |
| Data Files' ready-made and sample databases | `packages/templates/data-files` | 2.3 MB | Data Files |
| The template gallery: files and thumbnails | `packages/templates/files, packages/templates/thumbs` | 1.6 MB | core |
| The words in Urdu, Arabic and Hindi | `packages/office-ui/src/catalogues` | 4.0 MB | core |
| The studio renderer (layers for Video, Image and Recorder) | `packages/studio/src` | 616 KB | core |
| Video's containers and export plan (the codecs are Chromium's own WebCodecs; none ship) | `packages/media/src` | 79 KB | core |
| AI Tools' server (the Model Context Protocol, local and over the network) | `apps/desktop/main/mcp` | 123 KB | AI Tools |
| Talk's calling and messages | `apps/desktop/renderer/apps/talk, apps/desktop/renderer/apps/talk.js, apps/desktop/main/talk.js, apps/desktop/main/talk-signin.js` | 97 KB | Talk |
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

### The imports to cut (11)

Each is a file of one owner importing a file of another app. Before the move each is cut (the app asks through IPC or an event, or the shared part moves to the core). Grouped by who imports whom.

| From | To | Where | Imports | Kind |
|---|---|---|---|---|
| Data Files | AI Tools | `apps/desktop/renderer/apps/base/ai-tools.js:11` | `apps/desktop/renderer/ai-tools.js` | static |
| core | AI Tools | `apps/desktop/renderer/dav-accounts.js:16` | `apps/desktop/renderer/ai-tools.js` | static |
| core | AI Tools | `apps/desktop/renderer/shell.js:16` | `apps/desktop/renderer/ai-tools.js` | static |
| Mail | AI Tools | `apps/desktop/renderer/apps/mail/dialogs.js:17` | `apps/desktop/renderer/ai-tools.js` | static |
| Recorder | Video | `apps/desktop/renderer/apps/recorder.js:24` | `apps/desktop/renderer/apps/video/recorder/session.js` | static |
| Recorder | Video | `apps/desktop/renderer/apps/recorder.js:25` | `apps/desktop/renderer/apps/video/recorder/panel.js` | static |
| Recorder | Video | `apps/desktop/renderer/apps/recorder.js:26` | `apps/desktop/renderer/apps/video/recorder/css.js` | static |
| Recorder | Video | `apps/desktop/renderer/apps/recorder/live-dialog.js:12` | `apps/desktop/renderer/apps/video/recorder/panel.js` | static |
| Presentations | Video | `apps/desktop/renderer/apps/slides.js:45` | `apps/desktop/renderer/apps/video/recorder/panel.js` | static |
| Presentations | Video | `apps/desktop/renderer/apps/slides/cameo.js:13` | `apps/desktop/renderer/apps/video/recorder/camera.js` | static |
| Presentations | Video | `apps/desktop/renderer/apps/slides/screen-record.js:7` | `apps/desktop/renderer/apps/video/recorder/session.js` | static |

### The core's wiring (28)

What the core imports from each app today: `renderer/index.js` its window module, and `main/app-packs.js` its register module (`<id>-pack.js`), which builds its services. These are not cut: they are the registration, which a pack carrying its own code takes over from the core.

| Pack | Where | Imports |
|---|---|---|
| Documents | `apps/desktop/main/app-packs.js:32` | `apps/desktop/main/word-pack.js` |
| Worksheets | `apps/desktop/main/app-packs.js:33` | `apps/desktop/main/sheets-pack.js` |
| Presentations | `apps/desktop/main/app-packs.js:34` | `apps/desktop/main/slides-pack.js` |
| PDF | `apps/desktop/main/app-packs.js:35` | `apps/desktop/main/pdf-pack.js` |
| Pictures | `apps/desktop/main/app-packs.js:36` | `apps/desktop/main/pictures-pack.js` |
| Image | `apps/desktop/main/app-packs.js:37` | `apps/desktop/main/image-pack.js` |
| Video | `apps/desktop/main/app-packs.js:38` | `apps/desktop/main/video-pack.js` |
| Recorder | `apps/desktop/main/app-packs.js:39` | `apps/desktop/main/recorder-pack.js` |
| Mail | `apps/desktop/main/app-packs.js:40` | `apps/desktop/main/mail-pack.js` |
| Calendar | `apps/desktop/main/app-packs.js:41` | `apps/desktop/main/calendar-pack.js` |
| Contacts | `apps/desktop/main/app-packs.js:42` | `apps/desktop/main/contacts-pack.js` |
| Data Files | `apps/desktop/main/app-packs.js:43` | `apps/desktop/main/base-pack.js` |
| Talk | `apps/desktop/main/app-packs.js:44` | `apps/desktop/main/talk-pack.js` |
| AI Tools | `apps/desktop/main/app-packs.js:45` | `apps/desktop/main/ai-pack.js` |
| Text recognition | `apps/desktop/main/app-packs.js:46` | `apps/desktop/main/ocr-pack.js` |
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

### Shared files one app alone uses (2)

Files in the shared folders that only one app reaches: they move into its pack.

| File | Moves into |
|---|---|
| `apps/desktop/main/check-smtp.js` | Mail |
| `apps/desktop/renderer/defaults.js` | Mail |

### Calls and windows across packs

Not imports, so nothing to cut, but a pack that is absent must not be offered: the button that would call it or open its window is shown only when it is installed.

| Pack | Calls (namespace: first place) | Opens a window of |
|---|---|---|
| Documents | Contacts (3): `apps/desktop/renderer/apps/word/envelopes.js:48 contacts.list`; Mail (2): `apps/desktop/renderer/apps/word/mailings.js:850 mail.accounts` | none |
| Worksheets | Text recognition (4): `apps/desktop/renderer/apps/sheets/receipt-row.js:82 ocr.cancel` | none |
| PDF | Text recognition (1): `apps/desktop/renderer/apps/pdf/index.js:594 ocr.cancel` | none |
| Pictures | none | Image: `apps/desktop/renderer/apps/pictures.js:667`; Video: `apps/desktop/renderer/apps/pictures.js:669`; PDF: `apps/desktop/renderer/apps/pictures.js:671` |
| Recorder | none | Video: `apps/desktop/renderer/apps/recorder.js:98`; Pictures: `apps/desktop/renderer/apps/recorder.js:615`; Image: `apps/desktop/renderer/apps/recorder.js:616` |
| Mail | Calendar (3): `apps/desktop/renderer/apps/mail.js:1077 calendar.openFile`; Contacts (2): `apps/desktop/renderer/apps/mail.js:1130 contacts.fromMail`; Text recognition (1): `apps/desktop/renderer/apps/mail/docnote.js:116 ocr.cancel` | none |
| Calendar | none | Mail: `apps/desktop/renderer/apps/calendar.js:433` |
| Contacts | none | Mail: `apps/desktop/renderer/apps/contacts.js:226` |
| Data Files | AI Tools (3): `apps/desktop/renderer/apps/base/ai-tools.js:19 ai.stepOf` | none |
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

## 5. The move

What changes when the app code moves into packs, in the order the map
suggests. Technical only: each step lands with the whole gate passing, and
the map's fixture shrinks as the cuts are made.

### 5.1 The core learns to load packs, before any app moves

1. **IPC from the packs, not from one list.** `installIpc` takes the core's
   methods (the contract's `METHODS` less the namespaces apps own) plus each
   loaded pack's `ipc`, and registers handlers for those alone. A pack may
   not declare a namespace the core or another loaded pack has; such a
   pack is refused with the others. The sandboxed preload cannot import
   anything at run time, so it asks once, as it starts
   (`ipcRenderer.sendSync('rutba-office/contract')`), for the methods and
   events of this run, and builds the bridge from the answer. A channel of
   an absent pack then exists neither in the main process nor in the
   window's `shell`.

   *Built.* The contract (`packages/office-shell/src/contract.js`) keeps
   `METHODS` whole and says who owns each namespace (`NAMESPACE_PACK`) and
   each event (`EVENT_PACK`); `coreMethods()`, `methodsFor(ids)` and
   `eventsFor(ids)` cut it by owner. `installIpc(impls, methods)` registers
   only the methods it is given, and `installContract(answer)` answers the
   preload's one synchronous call on `CONTRACT_CHANNEL` with the methods,
   events, loaded packs and refusals of this run. A preload that gets no
   answer builds the core's channels alone. In the window, the client
   (`client.js`) keeps a stub for every method of the contract: one of an
   absent pack rejects at once, nothing crossing to the main process, with
   an error whose `code` is `absent` and whose words say which app to add
   ("Mail is not installed on this computer. Add it with Add or remove
   apps."). `installedPacks()`, `isInstalled(id)`, `refusedPacks()` and
   `packWindows(id)` say what this run has; `on()` for an absent pack's
   event listens to nothing. The menu offers only the installed apps, and a
   window asked for an absent app is made for another installed app that
   opens the file, or for the launcher, told which app was missing.
2. **main.js registers each pack.** At start, before the first window:

   ```js
   const { packs, refused } = findPacks({ dir: path.join(process.resourcesPath, 'packs'), publicKey: PACK_KEYS, version: electron.getVersion() });
   // ... in namespaces(ctx), after the core's own services:
   for (const pack of packs) {
     const { register } = await import(pathToFileURL(pack.main).href);
     const served = await register({ ...ctx, doc, oauth, broadcast, packs: view });
     // only the namespaces the signed manifest declared
     for (const ns of Object.keys(served)) if (!pack.ipc[ns]) throw new Error(`${pack.name} serves ${ns}, which it did not declare.`);
     Object.assign(services, served);
   }
   ```

   `view` is how one pack reaches another's service without importing it:
   `view.has('mail')`, `view.service('mail')`, answering `undefined` when
   Mail is absent. Today those reaches go through the `services` closure in
   `main.js` (Contacts' people and Calendar's accounts from Mail, Talk's
   invitations sent through Mail, Calendar's and Contacts' servers kept by
   `dav`); each becomes a `view.service(...)` call that copes with
   `undefined`. These are not imports, so the map does not list them; they
   are listed here. A refused pack is said on the launcher ("Mail could
   not be loaded: ... Install it again with Add or remove apps.") and in
   the log.

   *Built.* `main.js` imports no app's service. Each app has a register
   module beside its code (`main/<id>-pack.js`, `register(core)` answering
   the namespaces it serves), and `main/app-packs.js` decides which run:
   `planPacks` (every app built in when there is no packs folder: a run
   from the source, the portable build; otherwise only the signed packs
   found in `resources/packs`, or in `RUTBA_PACKS_DIR` with
   `RUTBA_PACK_PUBLIC_KEY` for a run from the source), then
   `registerPacks`, which imports each pack's main module (or, for a pack
   whose code is still in the core, hands it `builtIn`, the core's own
   register module), refuses one that throws or serves a namespace it did
   not declare, and leaves its channels out. `describePack(id)` is what a
   pack's manifest says of it (channels, events, tiles, file types, AI
   tools and its rows of the registry), read by the build and by a core
   with every app built in alike. The guards: Contacts' suggestions read
   Mail's people through `packs` (`peopleFrom`), Calendar's and Talk's
   accounts through `mailFor`, Mail's check for a known contact through
   `contactsFor`; each answers nothing while the other app is absent, and
   Talk refuses to send an invitation from Mail that is not there. `dav`
   is built after the packs, with the calendar and address book that are
   there. Data Files' rename hook is one it registers
   (`onRecentRenamed`).
3. **The windows' code from the packs.** The build stays one esbuild run
   with splitting, so chunk names agree across the core and every pack of
   one build (and one build is one version, which the loader enforces).
   The outputs are divided by the map's owners: a chunk only one app's
   window module reaches goes into that pack's archive; a chunk the
   launcher or two apps reach goes into the core. The protocol handler
   (`protocol.js`, `rutba://app/`) serves `packs/<id>/...` from that pack's
   archive, and `APP_LOADERS` in `renderer/index.js` becomes the list the
   preload hands the window: the installed apps and their window modules.
   A window for an absent app is never made: `win.create` refuses an app
   that is not loaded, and so does a file whose app is absent (it opens in
   whichever loaded app also opens it, or the launcher says which app to
   add).

   *Built, as far as it goes before a move.* While an app's code is in the
   core, its chunks are in the core's build too; dividing them is part of
   each app's move. A pack's manifest may name its window module
   (`windows`). The protocol handler serves `rutba://app/packs/<id>/<file>`
   from that pack's archive, from the roots `main.js` hands it (the loaded
   packs' archives), and never a path outside that archive.
   `renderer/index.js` imports a pack's own window module when the start
   says it has one (`packWindows`), the built-in one otherwise, and opens
   the launcher for an app that is not installed.
4. **The registry from the packs.** The tables in
   `packages/office-formats/src/registry.js` stay the build's source of
   truth; the build writes each app's rows (its `APPS` entry, extensions,
   writable formats, New templates, tiles) into its pack's manifest, and
   the core's registry at run time is the core's rows plus the loaded
   packs'. Choose apps, the Open dialog's filters, Save As, the launcher and
   `defaults.js` read that.

   *Built.* Each manifest carries its app's rows (`describePack(id).registry`)
   beside its tiles and file types. In the window, the launcher's tiles,
   Choose apps and the Open dialog's filters are the installed apps' only
   (`isInstalled`, `installedPacks`, `openFilters(app, { apps })`); Choose
   apps keeps the stored choice of an app not installed as it was, so adding
   it back brings it back as the person left it. The tables stay in the
   core's registry module until the apps move: they are a few kilobytes of
   data, and the build reads them.
5. **AI tools from the packs.** `TOOL_LIST` is split by owner (the map's
   `TOOL_PACK`): AI Tools keeps the folder tools, and each pack's
   `register` adds its own tools to a list the AI Tools pack serves. The
   AI tools' program run as Node (`main/mcp/node.js`, started with
   `ELECTRON_RUN_AS_NODE`) runs the loader too, so an AI app sees only the
   tools of packs that are installed and good.

   *Built.* `main/pack-tools.js` says whose each tool is (`TOOL_PACK`) and
   whether it is offered (`toolOffered`). The core writes the packs it
   loaded into the settings at each start (`packs.loaded`); the tools'
   list and every call read it, in the tools Rutba Office serves and in
   the program an AI app starts alike. Without AI Tools no tool is
   offered; with it, its own and those of the packs loaded. A call to a
   tool not offered is refused in plain words. The tools themselves stay in
   one list in AI Tools' code until the apps move.
6. **Development.** `npm start` and the window checks build every pack into
   `build/packs/` and sign them with a development key made once per
   checkout outside the repository (`make-pack-key.mjs`, in the profile
   folder). A packaged core holds only the release key, so a development
   pack never loads in an installed copy.

   *Built, otherwise.* `npm start` runs with every app built in, as before.
   The packs' window checks (`npm run verify:packs`,
   `build/verify-packs-run.js` and `main/verify-packs.js`) make every pack
   with a key made for the run and thrown away after it, hand the core its
   public half (`RUTBA_PACK_PUBLIC_KEY`), and start Rutba Office twice over
   one profile and a packs folder of Documents, Mail, AI Tools and Talk
   (`RUTBA_PACKS_DIR`). The first start checks that an app not there has no
   channel in the main process or the window, no tile and no window (its
   file opens the launcher, which says it is not installed), and that the AI
   tools follow; then Add or remove apps adds Calendar from a folder of the
   release's pack files (`RUTBA_PACKS_SOURCE`) and removes Mail. The second
   start has Calendar and not Mail, and Mail's files are gone. Every app
   but AI Tools carries its code (item 7) and is told to the window as
   `code: 'own'`, AI Tools as `'core'`; the windows open, Talk's,
   Documents' and Mail's windows' chunks come from their packs, and text recognition reads a picture with the engine
   and models in its pack. A release build reads `RUTBA_PACK_KEY_FILE`.
7. **A pack that carries its code.** Three things let an app's code leave
   the core's archive for its pack, and an app moves when all three hold
   for it:

   - **The renderer's chunks by owner.** `build/bundle.js` asks esbuild
     what went into each file it writes and keeps the answer in
     `build/meta/renderer.json` (each output's source files and packages,
     the outputs it imports, and the entry it starts from). An app's window module is already its own chunk
     (`chunks/<id>-<hash>.js`, imported when a window asks for it); the
     list says whether anything of another owner is in it, and whether a
     shared chunk holds some of the app.
   - **Packages from the core.** A module in a pack imports what it carries
     by relative paths and its own `node_modules`. A package it does not
     carry (`@rutba/office-shell`, an engine the core has) is not under
     any folder Node walks up to from the packs folder, so
     `@rutba/office-shell/electron/pack-resolve` (`installPackResolution`,
     a `module.registerHooks` resolve hook, set before the packs are
     registered) answers a bare import from inside the packs folder that
     fails as if the core's main module had made it. The pack gets the
     core's own copy, the one the rest of the run shares. A `require()`
     (Mail's PDF writer requires `@rutba/pdf`) is answered by the core's
     own require, as Node resolves a require from the module that made it
     whatever the hook is told. A relative import a pack cannot answer
     stays an error.
   - **Copying the app in, and leaving it out of the core.**
     `build/pack-contents.js` reads the map and the renderer's list and
     works out what the app takes with it:

     - *its main-process files*, those the map gives the app, its register
       module among them;
     - *its window chunks*, by the windows that reach them. The build also
       says what each output imports and which entry it starts from, so
       the walk starts at each app's window module and stops at another
       app's. A chunk only this app's window reaches goes with it, even
       when it holds a file the map calls the core's (a time-zone helper
       only Calendar's window runs). A chunk that holds the app's code and
       is reached by another window too is refused, with the windows that
       reach it;
     - *a worker its window starts by name* (Worksheets' `scripts-worker.js`,
       a bundle of its own); the image worker Image and Pictures share is
       the core's;
     - *its npm packages*, each with what it needs: a package needed by one
       owner's packages alone is that owner's, one two owners need
       (`iconv-lite`, which Mail's parsers and the MySQL client share)
       stays in the core;
     - *its workspace packages*, whole, or by file for an engine the map
       splits (`SPLIT_PACKAGES`: text recognition's, the templates, Data
       Files'). The pack carries such a package's code whole and the data
       that is the app's; the core leaves out only the files the app alone
       reaches. Text recognition takes its engine and models and the core
       keeps the result helpers the windows read; Data Files takes its
       engine and its sample databases and the core keeps the file reader
       (`@rutba/database/base`) Documents' mail merge reads a database
       with. A folder the engine reads from disk stays unpacked beside the
       pack's archive as beside the core's (`@rutba/ocr/vendor`).

     It refuses a move with the reason: a chunk that mixes the app's code
     with another's (with the import to cut), a file of the app that
     imports one of the core's by a relative path, or a file left in the
     core that imports one the pack would take. `build/make-app-packs.js`
     copies these into the pack (`main/`, `renderer/chunks/`, `renderer/`,
     `node_modules/`) for the apps in `CARRY_CODE`, makes the register
     module the pack's main module, and writes `build/packs/moved.json`:
     the core's paths each such app took with it, its tile among them.
     `electron-builder.config.cjs` leaves those out of `app.asar`.

   In the window, a chunk is fetched at the address the core would serve
   it from (`rutba://app/chunks/<name>`), so the shared chunks it imports
   stay the one copy every window module uses (React only once). The app
   host looks in each loaded pack's `renderer/` first and serves the
   chunk from there, with an `x-rutba-pack` header naming the pack. Any
   other file the core does not have (a worker started by name) is looked
   for in the loaded packs' `renderer/` the same way. The contract's
   `packs` says for each pack whether its code is its own (`code: 'own'`)
   or the core's (`'core'`, a main module that hands back to
   `core.builtIn`).

   The sources stay where they are in the repository: a run from the
   source and the window checks use them as before, and only a release
   build moves them.

   *Built.* `CARRY_CODE` holds Talk, Text recognition, Data Files,
   Calendar, Contacts, PDF, Pictures, Image, Documents, Worksheets,
   Presentations and Mail. Mail is handed the system font a PDF of mail is
   written in (`unicodeFont`, from `system-fonts.js`) in the context the
   core registers it with. Two apps cannot move yet:

   - AI Tools: its main side imports `pack-tools.js`, `system-fonts.js`
     and `app-packs.js`, which stay in the core, and its window code
     (`renderer/ai-tools.js`) is in a chunk every window reaches (the
     title bar's badge).
   - Video: `video/recorder/` is in a chunk Recorder and Presentations
     reach too (cut 2).

   Recorder could move as it is, its own `@rutba/stream` with it, but it
   waits for Video, which shares its recording engine.

### 5.2 The apps, in this order

From the map: least entangled first, the shared code moved to the core as
each cut is made.

1. **Talk, Calendar, Contacts.** Moved. No imports to cut; `dav` and the
   two small engines stay in the core.
2. **Data Files.** Moved. The rename hook is one the pack registers. Its
   engine, its sample databases (`@rutba/templates/data-files`) and the
   PostgreSQL and MySQL clients go with it; the core keeps the database
   file reader mail merge uses.
3. **Text recognition.** Moved. `@rutba/ocr` is split: the engine and its
   data into the pack, the result helpers (`record`, `languages`, `png`)
   and `renderer/ocr-ui.js` into the core. PDF, Pictures, Image,
   Worksheets, Mail and AI Tools name it in `uses`; its calls say it is
   not installed when it is not.
4. **AI Tools.** The title bar's badge and menu item come from the pack
   (the cut in `renderer/shell.js`); its main side reaches what it imports
   from the core through the context the core registers it with. Text
   recognition is already reached through its service.
5. **Recorder and Video, then Image and Pictures.** Image and Pictures
   have moved (cut 5). `video/recorder/` goes to the core with the
   `capture` namespace; `@rutba/stream` goes with Recorder.
6. **PDF.** Moved, pdf.js and its wrapper in the core (cut 4).
7. **Mail.** Moved. The largest service and the most reached (`view`
   callers in Calendar, Contacts and Talk); `imapflow`, `nodemailer` and
   `@rutba/mailbox` go with it. `oauth` and the system fonts stay in the
   core.
8. **Presentations, Worksheets, Documents.** Moved: the shared window code
   is the core's (cut 1). `present.js` goes with Presentations;
   `legacy-deck*.js` and `mailmerge-source.js` stay with the document
   service in the core.

### 5.3 The installer

- The packs are built before electron-builder runs, into `build/packs/`,
  and are not part of `app.asar` or the installer's main archive. The
  NSIS script (`resources/installer.nsh`) embeds each pack's files and, in
  `customInstall`, writes a pack's files only when its box is ticked
  (`File` inside an `${If}`), so a pack not chosen is never on the disk,
  not even for a moment.
- The page that already asks which apps (`rutbaAppsPageShow`) gains AI Tools
  as a pack (today an add-on box) and Text recognition, ticks what a ticked
  app needs, and says each pack's size.
- File types: `fileAssociations` leaves `electron-builder.config.cjs`
  (which registers every type for every app); the installer registers the
  types of the packs it writes, from their manifests, and the uninstaller
  removes every type Rutba Office registered (a `customUnInstall` that
  reads the list the core keeps in the registry).
- The installed list: the installer writes `HKCU\Software\Rutba Office`
  `Packs` (the ids it wrote), and the core writes it again at each start
  from what it loaded, so the next installer opens on what is there.
- The portable build, the macOS disk image and the Linux packages carry
  every pack: there is no installer step to choose in, and a read-only
  bundle (an AppImage, a signed `.app`) cannot lose a pack. Choosing there
  is a later step (packs in the user's own folder).

*Built.* `build/make-app-packs.js` runs after the renderer's build and
before electron-builder (`npm run dist`). For each pack it makes a folder
from `describePack` (its main module, its tile, the icons of its kinds of
file, kept on disk for the taskbar and Explorer), signs it with
`tools/make-packs.mjs` into `build/packs/out`, and writes the release's
file for it (`Rutba-Office-<version>-<id>.pack`) beside the installers.
The private key is read from `--key`, `RUTBA_PACK_KEY` or
`RUTBA_PACK_KEY_FILE`, never from the repository; the public half goes
into the core as `build/out/pack-keys.json`, with each pack's size. It
also writes `build/packs/installer-packs.nsh`, which
`resources/installer.nsh` includes:

- **Apps to install**, a page before the apps page: a box per pack, all
  ticked on a first install and on an update from a version with one
  archive, the installed ones on any later run (the `Packs` value), or
  those `/PACKS=a,b` names. One app at least must be ticked.
- **Choose apps** greys the apps not ticked there, and treats an app not
  installed as on, so that adding it later shows it.
- `customInstall` writes each ticked pack (archive, then its folder, then
  its manifest) and registers its kinds of file and nothing else; it
  writes `Packs`, and `PacksKnown`: every pack this installer has. The
  core writes both again at each start of an installed copy, `Packs` from
  what it loaded.
- `customUnInstall` lets go of every kind of file any pack registers. The
  pack files go with the install folder.

The Windows build has no `fileAssociations` of its own any more; the
macOS and Linux builds keep them.

`npm run dist` runs `build/dist.js` after the renderer's build: the packs
first, then electron-builder. On Windows it runs twice, the installer and
then the portable build; the portable build, the macOS disk image and the
Linux packages carry every pack in `resources/packs`
(`RUTBA_ALL_PACKS=1` adds them as extra resources), so an app whose code
is in its pack is there too. Such a copy loads its packs as an installed
one does, has nothing to add or remove, and writes nothing for an
installer to read (`main.js`, `everyPack`). The window checks and the screens
build (`main/verify-*.js`, `main/screen-*.js`) are left out of the
archive.

### 5.4 Add or remove apps

The launcher's Choose apps becomes Add or remove apps.

- **Remove:** the app's windows close; the pack is written down as to go,
  and at the next start, before the loader runs (Windows will not delete a
  file Electron holds open), its archive, unpacked folder and manifest are
  deleted and its file types given back. The person's documents, mail,
  settings and profile data stay unless they ask.
- **Add:** each release carries, beside the installer, one file per pack
  (`Rutba-Office-<version>-<id>.pack`, a zip of its archive, unpacked
  folder and manifest). The core downloads it from the release of its own
  version, unzips it into a staging folder (with the zip reader the engines
  already have), runs the loader on the staging folder with its own key and
  version, and only then moves it into the packs folder and loads it at the
  next start. Offline, Add or remove apps runs the full installer the person
  has (`/S /PACKS=mail,calendar`), which writes just those packs.
- **Per-machine installs** (in Program Files) are not writable by the
  person: there Add or remove apps runs the installer elevated, with the
  usual prompt, rather than writing the folder itself.

*Built.* `main/packs-service.js` is the `packs` namespace (`list`, `add`,
`remove`, `keep`, `restart`), and the launcher's **Add or remove apps…**
opens it (`renderer/packs-dialog.js`); the launcher also says which app a
file or a window asked for is not installed, and which packs on this
computer would not load and why.

- **Add** fetches the release's file for the core's own version
  (`${SITE.releases}/download/v<version>/Rutba-Office-<version>-<id>.pack`),
  or takes the pack from a folder named by `RUTBA_PACKS_SOURCE` (loose
  packs or the release's files: the checks use this). A release file is
  refused if it holds anything but that pack's three things. The pack is
  checked in a staging folder by the loader, with the core's keys and
  version, and put in place with its manifest last; it loads at the next
  start. Offline, or a file missing from the release, is said in plain
  words, and nothing is put in place.
- **Remove** writes the pack down to go (`pending.json` in the packs
  folder); the next start deletes it before the loader runs, and leaves it
  out even when another running copy still holds its files. **Keep** takes
  that back. A pack added in this run and removed again goes at once. The
  last app cannot be removed: uninstalling is Windows Settings' job.
- **Restart** starts Rutba Office again. In a copy installed for everyone
  the folder is not the person's to write: the checked packs wait in the
  staging folder, and Restart runs one PowerShell script with the
  administrator's permission (`Start-Process -Verb RunAs`) that waits for
  Rutba Office to close, puts the packs in place or deletes them, registers
  or lets go of their kinds of file for the machine, and starts it again.
- In a per-user copy the kinds of file of a pack added are registered for
  the person at once, and a pack removed lets go of its own at once, except
  those an app that stays opens too. Restart writes `Packs` for the next
  installer. A copy with every app built in says there is nothing to add
  or remove.

### 5.5 Updates

- electron-updater still fetches the one installer of the new version and
  runs it silently on quit. The silent installer writes the packs in
  `Packs` and no others, so the core and every installed pack move to the
  new version in one step, and the loader's version check holds the line:
  a pack left behind (an interrupted update) is refused with its reason
  and offered again, never run against a core it was not built with.
- **Updating from a version with one archive.** There is no `Packs` value
  yet. The silent installer then writes every pack: everything was
  installed before, so nothing the person used goes missing. The apps they
  had turned off in Choose apps stay turned off (the stored choice is
  read as before), and Add or remove apps offers to remove them, with the
  space each takes.
- A pack downloaded by Add or remove apps is always of the core's own
  version, so an add never needs an update and an update never leaves an
  add behind.
- **A pack new in a version.** `Packs` alone would leave out an app the
  version before did not have. So the installer and the core also write
  `PacksKnown`, every pack that version had, and an update writes the
  packs in `Packs` and each pack of its own that is not in `PacksKnown`:
  an app new in this version is installed, and one the person left out
  before stays out. With `Packs` and no `PacksKnown`, every pack is taken
  as known and none is added.

### 5.6 The release check: a core with one pack holds nothing of another

`tools/check-packs.mjs`, run by the gate on the built tree:

1. **Files.** Every file in the core's archive and every window chunk is
   owned by the core in the map (`tools/app-map.mjs` over the build's
   metafile and the main-process files); no file of any app's home is in
   it.
2. **One pack installed.** For each pack: a temporary install of the core
   with that pack alone (and what it needs); the loader loads exactly it.
3. **What the running core offers.** The core started as a check run with
   that one pack: the IPC handlers registered are the core's and the
   pack's namespaces only; the preload's bridge offers the same; the
   registry's apps, file types and tiles are the pack's; the AI tools
   listed are AI Tools' and the pack's, or none.
4. **The installer.** An install with a short list of apps writes those
   packs and registers only their file types; an uninstall leaves no file
   type behind.

*Built.* `node tools/check-packs.mjs` after `npm run dist` reads the
packaged core (`release/win-unpacked/resources/app.asar`, or `--core`),
the packs the installer carries (`build/packs/out`) and its include
(`build/packs/installer-packs.nsh`), and exits 1 on a failure:

- **The core's files by owner**, by the map's owners: the main process's
  files, the window chunks by the windows that reach them as
  `build/pack-contents.js` reckons (the renderer's list,
  `build/meta/renderer.json` or `--meta`; by what went into each when it
  does not say what imports what, and by the app each is named after
  without it), the tiles, and
  each engine and npm package by who reaches it (an npm package's own
  dependencies go with it). A pack whose main module still registers the
  core's own code is reported with how much of it the core holds; a pack
  that carries its code fails while the core holds any file of it, or a
  chunk that mixes some of it with another owner's code. The
  window checks (`main/verify-*.js`, `main/screen-*.js`) in the core fail
  too.
- **Each pack alone**, copied into a folder of its own and planned as a
  start plans it with the core's keys and version: it loads exactly
  itself; the channels and events are the core's and its own; the windows
  are told of it alone; its tiles, kinds of file and AI tools are its own;
  the Open dialog offers its kinds of file only; no AI tool is offered
  without AI Tools, and with AI Tools its own are.
- **The installer's part for each pack** writes that pack's three things
  only, its manifest among them, and registers its kinds of file only.

Starting the running core with one pack, and installing with a short list,
are the window checks' part (`main/verify-packs.js`) and the installer run
at release.

### 5.7 The riskiest parts

- **The window chunks.** A shared chunk that holds one app's file puts
  that app's code in the core; an import the map did not see breaks a
  window when the other pack is absent. The map's test keeps new imports
  out, and the release check's first part reads the real build.
- **The services reaching each other through main.js.** Not imports, so not
  in the map: every `services.mail` reached from another service is a
  place where an absent pack must be met with `undefined` rather than a
  throw.
- **The contract and the preload.** Turning one static list into a list per
  run touches every window and the sandboxed preload.
- **File types and uninstall.** Moving registration out of
  electron-builder means the uninstaller no longer cleans up for free.
- **Write access and timing.** A per-user install's packs folder is
  writable by the person (as `app.asar` is today), so the check at start is
  against corruption and swapped files, not against a local
  administrator. Hashing every pack costs a few hundred milliseconds at
  start.
- **The update from one archive.** The old version's uninstaller runs first
  and removes `app.asar`; the new installer must then write every pack. The
  update test from the last single-archive version proves it.
