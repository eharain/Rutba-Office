# Rutba Office

A free, open-source desktop office suite. Twelve apps in one download, working
with the network switched off. It starts with the Office apps; the others are
added from Choose apps whenever you want them.

**[office.rutba.io](https://office.rutba.io)** ·
[Download](https://github.com/eharain/Rutba-Office/releases) ·
[Contact us](https://office.rutba.io/contact)

| | | |
|---|---|---|
| **Mail** | Every account, every archive, one inbox | An address and a password — the servers are found from the domain's own records — with sign-in to Gmail, Outlook.com and Microsoft 365, or a file of accounts set up at once; plus the archives other clients leave behind — Outlook `.pst` and `.ost`, `.olm`, mbox, `.eml`, `.msg`. Names the trackers in a message, and offers the way off the list |
| **Calendar** | Meetings, kept on your own machine | Month, week, day and agenda; `.ics` in and out, repeats and time zones included; an invitation in a message answered from Mail, with a reply the organizer's calendar understands |
| **Contacts** | The address book Mail completes from | `.vcf` and CSV in, vCards out; a sender kept from the reading pane in one click |
| **Documents** | Documents that open the same everywhere, laid out on pages | `.docx` on an engine we own, opening and saving `.odt`, `.rtf` and web pages whole, reading `.doc` and text — and a GitHub-flavoured Markdown editor that gives a README back unchanged |
| **Worksheets** | Real formulas, real recalculation | `.xlsx` with a full calculation engine, reading `.xls` of every age, `.ods` and `.csv`, and old files to save as `.xlsx`: Lotus 1-2-3, dBASE and FoxPro tables, Access databases (read only), SYLK and DIF |
| **Presentations** | Slides that survive the round trip, shapes you move and recolour by hand | `.pptx` — read, edit, render, present — pictures and shapes on a slide, speaker notes, and a presenter view for the other screen; reading `.ppt` and `.odp` |
| **PDF** | Pages put in order, without redrawing them | Read, find and copy; reorder, rotate, delete, extract, insert, split and merge, each page copied into the new file exactly as it was; Recognise Text makes scanned pages searchable with an invisible text layer; pages drawn by pdf.js, offline |

| **Pictures** | A viewer that opens before you blink, and stays quick on a folder of thousands | Every common format, EXIF, orientation, PDFs, clips and sound in the same folder; thumbnails made once by the platform and kept; only what is in view is drawn; star ratings and tags kept on this computer, a search through the folders inside, and a whole folder resized, converted and renamed in one go; Copy Text from Picture reads the words in a photo or scan |
| **Image** | Crop, correct, select, paint, layer, export | Levels, curves, white balance, straighten; selections, brushes, clone and heal, masks and perspective; layers with blend modes and adjustments, saved as a project; Edit in Image from the other apps; non-destructive, so your original is never touched |
| **Video** | A player first, then an editor on a timeline, a screen recorder and introduction videos, without a render farm | A video or sound file opens in a player (speed, repeat, captions, picture in picture, full screen) until Edit Video; clips, stills and sound on lanes, split and trim, transitions, titles and captions, a library of titles, word art, shapes, icons and sound effects, starters to begin from, speed with the pitch kept, Freeze Frame and Reverse, a preview on a second screen, the screen recorded with the camera beside it, introductions made scene by scene with a teleprompter, and export with sound to MP4 or WebM for where it is going; no ffmpeg, no native binaries, nothing to install |
| **Recorder** | The screen, the camera and the voice made into one video as you record | Scenes to switch between while it records (the screen with the camera in a corner, the screen alone, the camera alone, the two side by side, a Be right back card), a title across the foot and a logo in a corner, Region to record a part of a screen or window, Take Picture for a still of what is shown, the screen or the camera, the microphone and the computer's sound with their levels, pause and a key that stops it from anywhere; one finished WebM to play, or to edit in Video; and Go Live, streaming the same picture to YouTube, Facebook, Twitch or any RTMP server at once, the stream keys kept encrypted on this computer |
| **Talk** | Messages and calls with your contacts | Someone becomes a contact once both of you have said so; messages and calls go straight between the two computers, a small Rutba service only introducing them, with a safety code to read out on a call; a relay of your choice, or your own, where the two cannot reach each other; signed in with your Rutba account in the browser |
| **Data Files** | Tables of records, related to each other, in one file | A table and its design (text, numbers, currency, dates, yes/no; required, unique, a default), records a page at a time, relationships the file keeps, queries designed with criteria and totals or written in SQL, forms made by a wizard, reports grouped with totals to print or save as PDF, ready-made databases to start from, a recipient list for Documents' mail merge, and imports from CSV and the old files Worksheets reads; one `.rbase` SQLite file, and any SQLite database opens too |

Dual-licensed: **GNU AGPL v3.0**, or a commercial licence — see
[COMMERCIAL-LICENSE.md](COMMERCIAL-LICENSE.md).

## Video

[![Rutba Office: ten free office apps for Windows, no account needed](https://raw.githubusercontent.com/eharain/Rutba-Office/main/docs/media/video-thumbnail.jpg)](https://www.youtube.com/watch?v=vPo8xXONjjk)

A 1:40 introduction on YouTube, with English captions.

---

## Why this exists

Most free office software asks you to give something up: your file format, your
privacy, your old mail, or your patience. This asks for none of those.

- **It opens what you already have.** Microsoft formats are the point, not an
  afterthought. So are the mail archives you have been carrying between
  computers for fifteen years.
- **It works offline.** No account, no telemetry, nothing to sign up for. It
  makes exactly two requests of its own — an update check, and one announcement
  check a day — and both are switches. Unplug the machine and everything in the
  list above still works.
- **It does not damage your files.** The engine rewrites only the parts of a
  document it deliberately edited; everything else — charts, pivot caches,
  macros, signatures, embedded media — is returned byte-for-byte as it arrived.
- **The source is published.** Under the AGPL, which means you can read it,
  change it and share it.

## Install

**Rutba Office is in beta.** Every release is published as a pre-release,
the window title says so, and an installed copy follows the beta releases.
It opens and edits real files today; it also still meets documents it draws
wrongly and places where the way to do something is not yet obvious. The
issue tracker is where those go, and the release notes say what each
release fixed.

Downloads are built for Windows. The build configuration has macOS and Linux targets, but
no release has carried them yet; build them from source with `npm run dist:mac` or
`npm run dist:linux`, and they are named as below:


| Platform | File |
|---|---|
| Windows | `Rutba-Office-<version>-win-x64.exe` (installer) or `-portable.exe` |
| macOS | `Rutba-Office-<version>-mac-<arch>.dmg` |
| Linux | `Rutba-Office-<version>-linux-x86_64.AppImage`, or the `.deb` |

### Code signing

The Windows builds are not code-signed, so SmartScreen warns on first run.
That warning is telling you the truth: Windows cannot name a publisher for an
unsigned file. This note will change when a signed release ships.

Privacy policy: Rutba Office collects no user data at all — see [Privacy](#privacy)
below.

The Windows installer keeps **one copy** of Rutba Office on a machine: it
removes whatever version is already installed — for this user or for all
users, in whichever directory — before it installs, and it does the same when
an installed copy updates itself. Your documents and settings stay where they
are. The portable build installs nothing and is not counted.

Once Rutba Office is the default, each kind of file wears the icon of the
app that opens it — a document, a workbook, a deck, a picture, a video, a
message and a PDF each look like what they are in Explorer — and on
Windows each app's windows carry that app's tile and sit under their own
taskbar icon, so Documents and Worksheets can be pinned and opened on their own.


### What the website links to

[office.rutba.io](https://office.rutba.io) is the download page, and these are
the addresses it can rely on. Names are produced by `electron-builder` from
`artifactName`, so they are stable across releases:

```
https://github.com/eharain/Rutba-Office/releases/latest
https://github.com/eharain/Rutba-Office/releases/download/v<version>/Rutba-Office-<version>-win-x64.exe
https://github.com/eharain/Rutba-Office/releases/download/v<version>/Rutba-Office-<version>-portable.exe
https://github.com/eharain/Rutba-Office/releases/download/v<version>/Rutba-Office-<version>-mac-<arch>.dmg
https://github.com/eharain/Rutba-Office/releases/download/v<version>/Rutba-Office-<version>-linux-x86_64.AppImage
```

Each release also carries `latest.yml`, `latest-mac.yml` and `latest-linux.yml`
— the feeds the installed application reads to find a newer version. They are
generated by the build; a release published without them leaves existing
installations on the version they have.

## Build it yourself

```bash
git clone https://github.com/eharain/Rutba-Office.git office
cd office
npm install
npm test          # 2,997 tests, no network needed
npm run build     # bundle the renderer
npm start         # run the app
npm run dev       # run the app, rebuilding the window's code on every save
npm run dist      # installers for this platform, into apps/desktop/release
```

Node 20 or later. Nothing else — no Python, no C++ toolchain, no native modules.

## What is in the repository

```
packages/
  ooxml            the preserving .xlsx/.docx package layer — the engine core
  formula          spreadsheet calculation: parser, evaluator, dependency order
  sheet-view       the renderable model of a sheet: formats, geometry, selection
  doc-view         the editable model of a document: runs, positions, formatting
  presentation     the .pptx engine: package, scene, layout inheritance, SVG
  drawing          format-neutral scene graph, SVG renderer, chart layout
  editing          the undo history, and the rules that make it feel right
  pdf              a PDF writer, and a reader for page work, with no dependencies
  office-formats   what a file is: sniffing, ODF, RTF, compound file, text
  mailbox          MIME, mbox, .msg, .olm, .pst/.ost, and the local mail store
  imaging          image probing, EXIF, and the non-destructive edit pipeline
  media            container probing
  studio           the layer renderer the Video and Image editors are built on
  templates        the templates gallery, made by the engines and kept as files
  barcode          QR codes encoded and read, with no dependencies
  ocr              text recognition on worker threads: Tesseract 5 in WebAssembly, our own page analysis, layout and tables
  proofing         spelling, the thesaurus and the accessibility check
  calendar         iCalendar read and written, recurrence, time zones
  contacts         vCard and CSV address books, read and written
  office-shell     the platform contract, and its Electron backend
  office-ui        the design system every app renders with

apps/desktop       the application: main process, renderer, twelve apps
tools/             probes for looking inside real files during development
tests/             the engine suite
```

### The engines have a second consumer

The Rutba consumer line (the `rutba-suite` repo, `consumer/` beside this one
in the estate) runs on these packages too: Workspace's editors, Sign's drafting
surface and PDF writer, the assistant's report writers, comms' whiteboard and
Studio's undo all resolve `@rutba/ooxml`, `formula`, `sheet-view`, `doc-view`,
`drawing`, `editing` and `pdf` from *this* repository, linked in by `file:`
path from `consumer/package.json` since 2026-09-09. There is one engine, not
two. Two things follow:

- **The export maps are a contract.** `tests/engine-seam.test.js` pins every
  subpath the consumer line imports; renaming one is a build broken in another
  repository, so the test says so here.
- **The seam holds on both sides.** `drawing`, `editing` and `doc-view` import
  no format layer outside a backend (same test); the consumer line's
  `workspace/tests/shared-seam.test.js` checks the same through its links, and
  its own suite runs against these files on every change there.

### The two decisions that shape everything else

**Documents live in the backend.** The OOXML engine inflates with zlib and works
in Buffers — it is a Node engine, not a browser one. So a document opens, edits
and saves in Node, on a thread of the main process's own (since 1.37.0, so the
other windows answer while one document is busy), and the window draws the view
model it is sent.
A spreadsheet asks for the viewport it is about to paint and nothing else, which
is why the size of a workbook has no bearing on the size of a window's heap.

**We only rewrite what we edit.** An `.xlsx` is a zip of XML parts, and almost
everything we will not implement for years lives in a part that can be carried
through untouched. Fidelity is therefore high on day one and degrades only in
the specific places we choose to touch — which are also the places we test.

## Formats

**Read and write:** `.docx` `.xlsx` `.pptx` `.csv` `.tsv` `.txt` `.md` `.html`
`.pdf` (export) `.png` `.jpg` `.webp` `.webm` `.eml` `.mbox`

**Read:** `.doc` `.xls` `.ppt` `.odt` `.ods` `.odp` `.rtf` `.pst` `.ost` `.olm`
`.msg` `.emlx` and every image and video format the browser engine decodes.

A format we can read but not write is opened by converting it into one we can —
an `.odt` becomes a document, an `.ods` a workbook — and the title bar says so,
so Save As is never a surprise.

### Outlook data files

`.pst` and `.ost` are read in full: the node and block B-trees, the heap,
property and table contexts, folders, messages, recipients and attachments.
Verified against a real 906 MB `.ost` — 467 folders, 12,172 messages indexed in
under a second — which is how three things were found that no amount of reading
the specification would have given:

- Outlook 2013 and later compress the larger data blocks of an `.ost` with
  zlib. It is not in the published structure, and a reader that does not notice
  sees every message in the file as empty.
- In an `.ost` the hierarchy and contents tables are mostly not materialised —
  449 folders with five hierarchy tables between them — so the folder tree has
  to come from the node B-tree's parent pointers, which are always complete.
- A view table that does exist may still carry no subject column.

The substitution tables for Outlook's block encoding are transcribed from
Microsoft's published [MS-PST] open specification, and checked three ways when
they load: each is a permutation of 0–255, and the encode and decode tables are
exact inverses in both directions. A damaged table decodes nothing rather than
decoding wrongly.

## Updates

Installed copies check [github.com/eharain/Rutba-Office/releases](https://github.com/eharain/Rutba-Office/releases)
for a newer version — twenty-five seconds after launch, then every six hours.
A new release downloads in the background and installs the next time you quit;
nothing is ever installed while you are working. You are told at every step:
a card in every window when a release is found, with the download's progress
and then **Restart and update**; a chip in every status bar meanwhile; and,
the first time the new version opens, a card that says which version it is
now and which it was, with **What's new** opening the release's own notes
(they ship with the build, so they read with the network off).

That is one HTTPS request to GitHub's release feed. It carries no identifier and
nothing about you, your machine or your files. **About → Check for updates
automatically** turns it off, and when it is off no request is made at all.

Running from source never checks: updates apply to an installed copy.

## Privacy

Three requests, and here they all are. If a network monitor ever shows a fourth,
that is a bug and we want to hear about it. (Mail, calendars and contacts talk
to the servers you add, which is what they are for.)

| What | When | What it carries | Off |
|---|---|---|---|
| Update check | 25 seconds after launch, then every 6 hours | Nothing about you | About → Check for updates automatically |
| [Announcement](docs/ANNOUNCEMENT.md) | Once a day at most, from the launcher | The version and the operating system. **No identifier of any kind** | About → Show announcements |
| Mail server lookup | Only when you press Add (or Add for several accounts), for an address whose provider is not in the built-in list | The address's domain, asked of that domain's own DNS records (SRV, MX) and its own autoconfig and autodiscover addresses (autoconfig.domain, domain/.well-known, autodiscover.domain), then a knock on the servers they name, over https only (a redirect to plain http is refused). No lookup service of anyone else's. A server they name outside the address's own domain, and not in the built-in list, is shown by name and gets the password only if you agree; no server is ever sent it unencrypted | Fill the servers in yourself under Add account → Advanced |

The announcement is also how we know anyone is using this. The server counts
requests, and a request carries nothing that identifies the copy making it — so
what is counted is *"a copy of Rutba Office opened somewhere today"*, never you.
Turning it off turns off the notice board too: there is deliberately no setting
that reports without showing you something back.

Everything else:

- No account, no telemetry, no crash reports. Your documents, your mail, and
  what you do with them never leave this computer.
- Mail passwords go to the operating system's keystore — DPAPI, Keychain,
  libsecret — never to a file the application can read back in clear text.
- Message bodies render in a sandboxed frame with no scripts, no same-origin
  access, and a policy that blocks every remote fetch. Remote images load only
  when you ask, per message, so a tracking pixel cannot report that you opened
  the mail.
- Deleting a file moves it to the operating system's trash. An office suite
  should never be the reason something is unrecoverable.
- Mail goes further than blocking trackers: it **names** them. See
  [what a message is doing to you](apps/desktop/main/mail-insight.js).

## What is missing

[GAPS.md](docs/GAPS.md) lists, area by area, what the suite does not do yet,
written from the code. A gap nobody has written down is a gap nobody fixes.
[ACCESSIBILITY.md](docs/ACCESSIBILITY.md) says how the suite is checked against
WCAG 2.2 A and AA for screen readers, the keyboard and High Contrast, and what
it is known not to meet.

Other things worth reading:

- [AI-TOOLS.md](docs/AI-TOOLS.md): letting an AI app on this computer make and
  read documents with the suite, off unless you turn it on, with nothing
  listening on the network.
- [OAUTH.md](docs/OAUTH.md) — signing in to Gmail and Outlook.com, and why a
  client id is not a secret
- [ANNOUNCEMENT.md](docs/ANNOUNCEMENT.md) — the announcement contract, in full

## Testing

```bash
npm run gate            # all four passes, in the order that finds problems soonest
npm test                # the engine suite: 2,997 checks, no windows
npm run verify:edit     # do keystrokes reach the document?
npm run verify:apps     # does each app open, change and save a real file?
npm run smoke           # boot the real app, photograph every window, report
npm run smoke -- word   # just one
RUTBA_CORPUS_DIRS="D:\docs" npm run verify:corpus   # open every file in a folder in the real app, one window at a time
node tools/fuzz-open.js 500                         # damage good files at random; every one must draw or refuse in a sentence
```

`npm run gate` is what runs before a release: 3,005 engine tests, 12 editing checks,
2,150 application checks (most of them pressing the real ribbon buttons and reading what the page paints, a block of them driving the actual mouse and keyboard) and 14 window captures. The windows sit off the desktop, so a run never gets in your way and you cannot close it by mistake; it takes about twenty minutes.

The engine suite includes `tests/rich-fixtures.test.js`, which opens a workbook, a document and a deck that Excel, Word and PowerPoint themselves wrote — charts, shapes, pictures, cross-sheet formulas, names, number formats, notes, footnotes, tracked changes, fields, a watermark — and holds the engine to Excel's own results. `tools/make-rich-fixtures.ps1` regenerates them on a machine with Office; docs/TESTING.md says what they hold and what they found.

The window runs are the ones that find what a build cannot. They launch the
actual application — same main process, same preload, same bundle — in a profile
of their own against generated fixtures, so they pass or fail on the code rather
than on whatever happens to be on the machine. Several defects in this
repository's history were caught by them and by nothing else, including a video
timeline that was never built when the file loaded faster than the window did.

## Licence

Copyright © 2026 Tech Style Ltd (Company No. 11101491), registered in England &
Wales — [tech-style.co](https://tech-style.co). The product lives at
[office.rutba.io](https://office.rutba.io), and questions go to
[office.rutba.io/contact](https://office.rutba.io/contact).

Dual-licensed under the GNU Affero General Public License v3.0 — see
[LICENSE](LICENSE) — and a separate commercial licence for use without the
AGPL's obligations: [COMMERCIAL-LICENSE.md](COMMERCIAL-LICENSE.md).

Notice: the spelling dictionaries are third-party work under their own
permissive terms — nspell (MIT, © Titus Wormer) and the English Hunspell
dictionaries made from SCOWL (© 2000-2018 Kevin Atkinson and the other
authors named in their licence files; MIT and BSD-style terms), shipped as
the npm packages dictionary-en and dictionary-en-gb with their licence files.
[packages/proofing/README.md](packages/proofing/README.md) reproduces the
notice.

Notice: the PDF app draws pages with pdf.js (Apache-2.0, © Mozilla
Foundation), vendored with its licence in apps/desktop/vendor/pdfjs, together
with the standard fonts it ships (PDFium's, BSD-style; Liberation, SIL Open
Font Licence) and its WebAssembly decoders (OpenJPEG, BSD; PDFium's JBIG2,
BSD-style; qcms, MIT), each with its own licence file there.

Notice: the accessibility checks (apps/desktop/main/verify-a11y-audit.js) run
axe-core (MPL-2.0, © Deque Systems, Inc.) over the windows, as a development
dependency the checks alone load; nothing of it is in the installers. Its
licence is in node_modules/axe-core/LICENSE once the dependencies are installed.

Notice: text recognition reads with Tesseract 5 compiled to WebAssembly, as
tesseract.js 7 ships it (tesseract.js-core, Apache-2.0, © the tesseract.js
authors; Tesseract, Apache-2.0, © Google and the Tesseract contributors), and
the fast models of tessdata_fast (Apache-2.0, © Google and the Tesseract
contributors), vendored with their licences in packages/ocr/vendor.
