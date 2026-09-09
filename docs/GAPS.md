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
| **Page setup as a document property** | **Built for workbooks, 2026-09-09** | Paper, orientation, margins, scaling, gridlines, headings, centring, the print area and the repeated rows are read from and written to the file where Excel keeps them. A document and a deck still choose theirs in the dialog each time. |
| **Password-protected files** | **Missing** | Refused with a sentence. An encrypted OOXML package (`EncryptedPackage` in a compound file, agile encryption) is neither read nor written. |
| **A spell-check pass** | **Partly** | Chromium underlines as you type in the places that are contenteditable. There is no pass over a document, no dictionary the person owns, no add-to-dictionary, and nothing at all in the grid or on a slide. |
| **Accessibility check** | **Missing** | No alt-text prompt, no heading-order check, no contrast check. The windows are keyboard-reachable; nothing verifies it but a person. |
| **Right-to-left layout** | **Missing** | Arabic and Hebrew text draws, and paragraph direction, mirrored margins and RTL ribbons are not built. |
| **Localisation** | **Missing** | Every string in the interface is English, written inline. There is no message catalogue and no way to add one. |

---

## 1. Word — the document

Reads more than it writes. What it draws faithfully and cannot yet produce:

- **Floating layout.** A picture or a text box sits in its own paragraph;
  text does not wrap round it. Position, Wrap Text, Align, Group, the
  selection pane and z-order all wait on this one thing.
- **Columns and section control.** `w:cols` is read and not honoured: the
  paginator lays one column. Manual page breaks, line numbers, hyphenation
  and page borders are the same story.
- **Fields.** Bookmarks, cross-references, captions, an index, a table of
  authorities, citations and a bibliography, `SEQ`, `PAGE` outside the
  header — none are written, so the table of contents this suite inserts is
  text rather than a field that refreshes.
- **Mail merge**, envelopes and labels: a data source, merge fields, a
  preview and a run. Nothing of it exists, and Mail beside it is exactly
  what would make it worth having.
- **Recording tracked changes.** Tracked changes are read and shown;
  `w:ins`/`w:del` are not written as edits happen, so Accept and Reject have
  nothing to act on. Document comparison and Restrict Editing sit behind it.
- **Text effects and equations.** WordArt, outline/shadow/glow, drop caps,
  SmartArt, OMML equations: read past, not written.
- **Paragraph shading and borders**, page colour, a real watermark (the
  header holds text only), multilevel list definitions.
- **Building blocks** (Quick Parts) and a default template.

## 2. Worksheets — the workbook

The calculation engine is the strongest part of the suite; the sheet around
it is where the gaps are.

- **Things in a sheet that are not cells**: pictures, shapes, hyperlinks and
  cell comments are drawn when a file has them and cannot be added — each
  needs a part the writer does not yet write (a drawing anchor, the
  hyperlinks list, the comments part plus its VML).
- **Format as Table**, slicers, sparklines: read, not written.
- **Data tools**: text to columns, remove duplicates, multi-key sort,
  advanced filter with a criteria range, consolidate, flash fill, subtotals
  and outlining, forecast sheets.
- **Auditing**: trace precedents and dependents, evaluate step by step, the
  watch window, error checking. The dependency graph exists; nothing draws
  it.
- **Print settings in the file** (see above), page-break preview, print
  titles, a sheet background.
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

- **Shape formatting**: fill, outline, effects, quick styles. A shape can be
  drawn and moved; it cannot be recoloured.
- **Text inside a shape**: bullets and numbering, indent levels, line
  spacing, character spacing, strikethrough, highlight, text direction,
  vertical alignment, columns in a box, WordArt.
- **What can be put on a slide**: a table, a chart, video, audio, SmartArt,
  a hyperlink. All four are read and drawn; none can be inserted.
- **The master and layouts.** Masters and layouts are read and inherited
  from; editing them, swapping a theme, changing an existing slide's layout,
  background styles, headers and footers, slide numbers and date fields all
  need the writer to reach the master.
- **Structure**: sections, hidden slides, custom shows, saving as a show.
- **The show**: recorded narration and timings, transitions and animations
  as anything but ribbon tabs, export to video.
- **Review**: comments on a slide, tracked changes, find and replace across
  a deck.

## 4. Mail — the one with no register

Mail carries no `<Soon>` markers at all: IMAP and SMTP, sign-in to Google and
Microsoft, rules, search, conversations, attachments in one place, tracker
blocking and unsubscribe are built. What it is missing is not inside mail:

- the **calendar** and **contacts** above, which is where a mail client
  stops being half an application;
- **local search across archives** is there, but there is no index, so a
  large mailbox searches by walking;
- **signatures, out-of-office, and send-later** are not built.
- ~~adding an account asked for the servers and guessed them from the domain~~ — built 2026-09-09: an address and a password; the server is found from the provider table, MX, SRV, autoconfig, Microsoft autodiscover and a knock on the conventional names, with Advanced always a click away.

## 5. The formats, read against written

| Format | Read | Write |
| :-- | :-- | :-- |
| `.docx` `.xlsx` `.pptx` | yes, preserving | yes, preserving |
| `.doc` `.xls` `.ppt` | text only — an approximation, and only `.xls` says so | no |
| `.odt` `.ods` `.odp` | yes | **yes**, 2026-09-09 — what the suite models: text, structure, tables and run formatting; values, formulas and value types; text boxes, pictures and notes |
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
5. **Floating layout in the document paginator.** One piece of work unlocks
   eleven controls in Word and the same again in the deck.
6. **Shape formatting and text properties in the deck writer** — the largest
   single block of unbuilt controls in the suite.
7. **Things in a sheet that are not cells**: comments, pictures, hyperlinks,
   format as table.
8. **Fields in the document**: bookmarks first, then cross-references,
   captions and a table of contents that refreshes.
9. **Recording tracked changes**, then comparison and restricted editing.
10. ~~**The formula language's remaining corners**: array constants, `LET`,
    `TEXT()` through the formatter the suite already owns.~~ Built 2026-09-09,
    with the phantom day; `LAMBDA` remains.

Everything above this line is work somebody can start on Monday. Below it
sit the things that are deliberately not built — macros that run, cloud
services (translation, dictation, design ideas, co-authoring), and anything
that would make a request this suite has not declared.
