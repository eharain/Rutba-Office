# Rutba Office

A free, open-source desktop office suite. Seven apps in one download, working
with the network switched off.

**[office.rutba.io](https://office.rutba.io)** ·
[Download](https://github.com/eharain/Rutba-Office/releases) ·
[Contact us](https://office.rutba.io/contact)

| | | |
|---|---|---|
| **Mail** | Every account, every archive, one inbox | An address and a password — the servers are found from the domain's own records — with sign-in to Gmail, Outlook.com and Microsoft 365, or a file of accounts set up at once; plus the archives other clients leave behind — Outlook `.pst` and `.ost`, `.olm`, mbox, `.eml`, `.msg`. Names the trackers in a message, and offers the way off the list |
| **Calendar** | Meetings, kept on your own machine | Month, week, day and agenda; `.ics` in and out, repeats and time zones included; an invitation in a message answered from Mail, with a reply the organizer's calendar understands |
| **Contacts** | The address book Mail completes from | `.vcf` and CSV in, vCards out; a sender kept from the reading pane in one click |
| **Word** | Documents that open the same everywhere | `.docx` on an engine we own, reading `.odt`, `.rtf`, `.doc` and text — and a GitHub-flavoured Markdown editor that gives a README back unchanged |
| **Worksheets** | Real formulas, real recalculation | `.xlsx` with a full calculation engine, reading `.ods` and `.csv` |
| **Presentation** | Slides that survive the round trip | `.pptx` — read, edit, render, present — pictures and shapes on a slide, speaker notes, and a presenter view for the other screen |

| **Pictures** | A viewer that opens before you blink | Every common format, EXIF, orientation, and PDFs |
| **Image** | Crop, correct, annotate, export | Non-destructive: your original is never touched |
| **Video** | Trim and export without a render farm | No ffmpeg, no native binaries, nothing to install |

Dual-licensed: **GNU AGPL v3.0**, or a commercial licence — see
[COMMERCIAL-LICENSE.md](COMMERCIAL-LICENSE.md).

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

Downloads are built for Windows, macOS and Linux:


| Platform | File |
|---|---|
| Windows | `Rutba-Office-<version>-win-x64.exe` (installer) or `-portable.exe` |
| macOS | `Rutba-Office-<version>-mac-<arch>.dmg` |
| Linux | `Rutba-Office-<version>-linux-x86_64.AppImage`, or the `.deb` |

### Code signing

Code signing for the Windows builds is provided by [SignPath.io](https://signpath.io),
with a free code-signing certificate granted by the
[SignPath Foundation](https://signpath.org) to open-source projects.
**Thank you to the SignPath Foundation for signing Rutba Office.**

Certificates issued under that programme are held in SignPath Foundation's
name, so that — rather than Tech Style Ltd — is the publisher Windows shows.
It is the trade the programme asks for and it is worth stating plainly rather
than leaving somebody to notice it in a dialog.

> **Status:** the application is in progress and the binaries published so far
> are NOT yet signed, so SmartScreen still warns on first run. That warning is
> telling you the truth until this note says otherwise. This paragraph goes
> when the first signed release ships.

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
taskbar icon, so Word and Worksheets can be pinned and opened on their own.


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
npm test          # 688 tests, no network needed
npm run build     # bundle the renderer
npm start         # run the app
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
  pdf              a PDF writer with no dependencies
  office-formats   what a file is: sniffing, ODF, RTF, compound file, text
  mailbox          MIME, mbox, .msg, .olm, .pst/.ost, and the local mail store
  imaging          image probing, EXIF, and the non-destructive edit pipeline
  media            container probing and the video edit timeline
  office-shell     the platform contract, and its Electron backend
  office-ui        the design system every app renders with

apps/desktop       the application: main process, renderer, seven apps
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
and saves in the main process, and the window draws the view model it is sent.
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
nothing is ever installed while you are working.

That is one HTTPS request to GitHub's release feed. It carries no identifier and
nothing about you, your machine or your files. **About → Check for updates
automatically** turns it off, and when it is off no request is made at all.

Running from source never checks: updates apply to an installed copy.

## Privacy

Two requests, and here they both are. If a network monitor ever shows a third,
that is a bug and we want to hear about it.

| What | When | What it carries | Off |
|---|---|---|---|
| Update check | 25 seconds after launch, then every 6 hours | Nothing about you | About → Check for updates automatically |
| [Announcement](docs/ANNOUNCEMENT.md) | Once a day at most, from the launcher | The version and the operating system. **No identifier of any kind** | About → Show announcements |

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

## Where this stands against the alternatives

[COMPETITORS.md](docs/COMPETITORS.md) is an audit, app by app, against what you
would otherwise install — Microsoft 365, LibreOffice, OnlyOffice, Thunderbird,
Outlook, eM Client, Google. It lists the gaps as carefully as the leads, because
a gap nobody has written down is a gap nobody fixes. Track changes, macros, a
calendar and encryption are all on it.

Other things worth reading:

- [OAUTH.md](docs/OAUTH.md) — signing in to Gmail and Outlook.com, and why a
  client id is not a secret
- [ANNOUNCEMENT.md](docs/ANNOUNCEMENT.md) — the announcement contract, in full

## Testing

```bash
npm run gate            # all four passes, in the order that finds problems soonest
npm test                # the engine suite — 688 checks, no windows
npm run verify:edit     # do keystrokes reach the document?
npm run verify:apps     # does each app open, change and save a real file?
npm run smoke           # boot the real app, photograph every window, report
npm run smoke -- word   # just one
RUTBA_CORPUS_DIRS="D:\docs" npm run verify:corpus   # open every file in a folder in the real app, one window at a time
node tools/fuzz-open.js 500                         # damage good files at random; every one must draw or refuse in a sentence
```

`npm run gate` is what runs before a release: 688 engine tests, 9 editing checks,
165 application checks — most of them pressing the real ribbon buttons and reading what the page paints, a block of them driving the actual mouse and keyboard — and 10 window captures. The windows sit off the desktop, so a run never gets in your way and you cannot close it by mistake; it takes about two and a half minutes.

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
