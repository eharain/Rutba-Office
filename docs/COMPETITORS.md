# Where Rutba Office stands

An honest audit, app by app, against what people would otherwise install. It is
written to be useful rather than flattering: the gaps are listed with the same
care as the leads, because a gap nobody has written down is a gap nobody fixes.

Everything marked ✅ is built and covered by a check in `npm run verify:apps` or
`npm test`. Everything marked ⬜ is not built.

---

## The suite

| | Rutba Office | Microsoft 365 | LibreOffice | OnlyOffice | WPS | Google Docs |
| :-- | :--: | :--: | :--: | :--: | :--: | :--: |
| Price | Free | Subscription | Free | Free | Free tier + ads | Free with account |
| Source published | ✅ AGPL | ⬜ | ✅ MPL | ✅ AGPL | ⬜ | ⬜ |
| Works with no network | ✅ | Partly | ✅ | ✅ | ✅ | ⬜ |
| Works with no account | ✅ | ⬜ | ✅ | ✅ | Nags | ⬜ |
| Telemetry | **None** | Yes | Opt-in | Yes | Yes | Yes |
| Mail client included | ✅ | Outlook | ⬜ | ⬜ | ⬜ | ⬜ |
| Reads Outlook .pst/.ost | ✅ | ✅ | ⬜ | ⬜ | ⬜ | ⬜ |
| Picture, image and video tools | ✅ | ⬜ | ⬜ | ⬜ | ⬜ | ⬜ |
| One download for all of it | ✅ | ✅ | ✅ | ✅ | ✅ | n/a |

**The one request outward.** An update check against GitHub, and one announcement
check a day against office.rutba.io that carries the version and the platform and
no identifier. Both are listed in About and both switch off. See
[ANNOUNCEMENT.md](ANNOUNCEMENT.md).

---

## Mail

Against **Outlook**, **Thunderbird**, **eM Client**, **Mailbird**, **Apple Mail**,
**Spark**, **Superhuman**.

### Where we are ahead

| | Rutba | Outlook | Thunderbird | eM Client |
| :-- | :--: | :--: | :--: | :--: |
| Opens a .pst | ✅ free | ✅ | ⬜ **cannot** | Paid tier |
| Opens a .ost | ✅ free | Own profile only | ⬜ | ⬜ |
| Opens .olm (Outlook for Mac) | ✅ | ⬜ | ⬜ | ⬜ |
| Finds your other clients' mail for you | ✅ | ⬜ | ⬜ | ⬜ |
| Names the tracking networks in a message | ✅ | ⬜ | ⬜ | ⬜ |
| Counts blocked pixels, tells sender's own logo apart | ✅ | ⬜ | ⬜ | ⬜ |
| Shows SPF / DKIM / DMARC per message | ✅ | Add-in | Add-on | ⬜ |
| One-click unsubscribe (RFC 8058) | ✅ | ✅ | ✅ | ✅ |
| Undo send | ✅ | Paid/365 | ⬜ | ⬜ |
| Scheduled send | ✅ | ✅ | ⬜ | ✅ |
| Unified inbox | ✅ | ✅ | ✅ | ✅ |
| Attachment view across the mailbox | ✅ | ⬜ | ⬜ | ✅ |
| People view derived from your mail | ✅ | ⬜ | ⬜ | ⬜ |
| Account limit | None | None | None | **2 on free** |

The `.pst`/`.ost` reader is the headline. Thunderbird cannot open one at all,
eM Client puts it behind the paid tier, and Microsoft's new Outlook dropped PST
support. Verified against 906 MB of real mail: 467 folders, 12,172 messages.

The privacy report is the thing no client does. Everyone blocks remote images;
nobody says *"Mailchimp and Google Analytics tried to record that you opened
this, and the 180-pixel image is the sender's own logo"*. See
`apps/desktop/main/mail-insight.js`.

### Where we are behind — the real list

| Missing | Who has it | Why it matters |
| :--- | :--- | :--- |
| ✅ ~~OAuth for Gmail and Outlook.com~~ | All of them | **Built.** RFC 8252 flow in the system browser with PKCE and a loopback redirect. Needs a registered client id per build — see [OAUTH.md](OAUTH.md). |
| ✅ ~~Message rules and filters~~ | Outlook, Thunderbird, eM | **Built**, and unlike theirs it shows what a rule *would* do to the folder in front of you, live, before it does anything. |
| ✅ ~~A search index~~ | Outlook, Thunderbird | **Built.** Inverted index with `from:`, `subject:`, `has:attachment`, `is:unread` and quoted phrases. 20,000 messages: 494 ms to build, under 10 ms per search. |
| ⬜ Calendar and contacts (CalDAV / CardDAV) | Outlook, Thunderbird, eM | A mail client without a calendar is half a client to many people |
| ⬜ OpenPGP / S/MIME | Thunderbird (built in) | Encryption is why some people choose Thunderbird |
| ⬜ Junk filtering | All of them | We show that a message is bulk; we do not sort it |
| ⬜ RSS feeds | Thunderbird | Small, but Thunderbird users expect it |

---

## Word

Against **Microsoft Word**, **LibreOffice Writer**, **OnlyOffice**, **Google Docs**.

### Where we are ahead

- ✅ **A Markdown editor that will not rewrite your README.** 264 of 268 real
  `.md` files in this estate round-trip byte for byte — front matter, task
  lists, table alignment, reference links and footnotes all preserved. Word
  cannot open Markdown at all; LibreOffice converts it and loses the structure;
  Google Docs has no Markdown file support. See `packages/office-formats/src/markdown.js`.
- ✅ **Preserving OOXML.** An edit rewrites only the parts it touches, so
  macros, content controls, custom XML, revisions and signatures survive a round
  trip untouched. LibreOffice rewrites the whole document through its own model.
- ✅ Opens `.docx`, `.odt`, `.rtf`, `.txt`, `.md`, `.html`, and extracts what it
  can from `.doc`.

### Where we are behind

| Missing | Who has it |
| :--- | :--- |
| ⬜ Track changes | Word, Writer, OnlyOffice, Docs |
| ⬜ Footnotes and endnotes | all |
| ⬜ Table of contents | all |
| ⬜ Mail merge | Word, Writer |
| ⬜ Citations and bibliography | Word, Writer |
| ⬜ Equations | Word, Writer |
| ⬜ Columns and section breaks | all |
| ⬜ Editing a style's definition | all |

---

## Worksheets

Against **Excel**, **LibreOffice Calc**, **OnlyOffice**, **Google Sheets**.

### Where we are ahead

- ✅ **Opens a 100,000-row workbook without loading it into a window.** The grid
  asks for the viewport it is about to paint; the model stays in the backend.
- ✅ Real recalculation with a dependency graph, not a formula cache.
- ✅ Goal seek and what-if tables in the free product — Google Sheets has neither.

### Where we are behind

| Missing | Who has it |
| :--- | :--- |
| ⬜ Macros | Excel (VBA), Calc (Basic), OnlyOffice (JS) |
| ⬜ Sparklines, slicers, timelines | Excel |
| ⬜ Power Query / external data | Excel |
| ⬜ Array formulas and dynamic spill | Excel, Sheets |
| ⬜ Chart editing after insertion | all |
| ⬜ Solver | Excel, Calc |

---

## Presentation

Against **PowerPoint**, **Impress**, **Keynote**, **Google Slides**.

### Where we are ahead

- ✅ **One renderer for the editor, the sorter and the print.** A slide cannot
  look one way while you edit it and another way when it is shown.
- ✅ Preserving package: animations and transitions we do not model survive a
  round trip rather than being dropped.

### Where we are behind

| Missing | Who has it |
| :--- | :--- |
| ⬜ Transitions and animations (authoring) | all |
| ✅ ~~Presenter view with notes, timer and next slide~~ | all |
| ⬜ Master and layout editing | all |
| ⬜ Themes to choose from | all |
| ⬜ Tables and charts on a slide | all |

---

## Pictures, Image and Video

Against **Windows Photos**, **IrfanView**, **XnView**, **Paint.NET**, **GIMP**,
**Photos (macOS)**, **VLC**, **Shotcut**.

### Where we are ahead

- ✅ **One viewer for stills, animations, video and audio**, with an auto-advance
  navigator, a filmstrip, and a slideshow that waits for a video to finish
  before moving on. Windows Photos does not play video in the same surface;
  IrfanView needs a plug-in.
- ✅ Opens PDFs in the same viewer.
- ✅ Ships in the same download as the office apps, which no other free suite does
  at all.

### Where we are behind

| Missing | Who has it |
| :--- | :--- |
| ⬜ Layers | Paint.NET, GIMP, Photoshop |
| ⬜ RAW decoding | IrfanView, XnView, Photos |
| ⬜ Batch conversion | IrfanView, XnView |
| ⬜ Video effects, transitions, multi-track | Shotcut, DaVinci |

---

## What to build next, in order

Ranked by how many people it stops from using this at all.

1. ~~OAuth for Gmail and Outlook.com.~~ **Built** — see [OAUTH.md](OAUTH.md).
   A build still needs its own registered client id before it works.
2. ~~Message rules.~~ **Built.**
3. ~~A search index.~~ **Built.**
4. ~~Presenter view.~~ **Built** — a second window for the other screen.
5. **Footnotes and a table of contents** in Word — the two most common reasons a
   document has to go back to Word.
6. **Calendar.** The largest single piece of work on this list, and the one that
   turns a mail client into a replacement for Outlook rather than for Thunderbird.
