# Office Open XML: notes from real files

What the files Word, Excel and PowerPoint write show that ISO/IEC 29500 does
not describe, or that the applications do differently from it, and what
Rutba Office does in each case. The notes come from building Strict support
(2.19.0) and from the files in `tests/fixtures/`: `rich/` holds a document, a
workbook and a deck in the usual form, and `strict/` the same three saved by
Word, Excel and PowerPoint (version 16) as Strict Open XML.

Each note says what the standard says, what the files do, and what Rutba
Office does. "The standard" is ISO/IEC 29500:2016, which matches ECMA-376
5th edition: Part 1 is the markup, Part 2 the packaging (OPC), Part 3 Markup
Compatibility (MCE) and Part 4 the Transitional features.

The checks behind the notes:

| Check | What it proves |
|---|---|
| `tests/strict-ooxml.test.js` | A Strict file reads as its usual twin does; an unedited part is saved as the bytes it came as; a Strict file stays Strict when edited; both directions of Save As; every Strict save passes a structural check that needs no Office. |
| `tests/rich-fixtures.test.js` | The Strict files are held to the same results as the usual ones. |
| `tools/check-strict-office.ps1` | Word, Excel and PowerPoint open Office's Strict files, our Strict saves of the usual files and our edited Strict files, and find the same content in each. Run under the shared lock, on a machine with Office. |
| `RUTBA_VERIFY_ONLY=strictsave npm run verify:apps` | The menu entries in Documents, Worksheets and Presentations write the form they name. |

## 1. Transitional and Strict in practice

### 1.1 Telling the two apart

**The standard.** Strict and Transitional are two conformance classes with
different namespaces. Strict's namespaces start `http://purl.oclc.org/ooxml/`,
Transitional's `http://schemas.openxmlformats.org/`. The relationship types
move with them.

**The files.** The namespaces of Part 2 and Part 3 do not move: Office's
Strict files keep `http://schemas.openxmlformats.org/package/2006/...` for the
content types, the relationships and the core properties, and
`.../markup-compatibility/2006` for MCE. The reliable test is the type of the
root relationship to the main part in `_rels/.rels`:
`http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument`.
The file name and the content types are the same for both forms.

**Rutba Office.** A package is Strict when that root relationship says so.
The test is made once, when the file is read, in
`packages/ooxml/src/package.js`, and the answer is kept with the open file.

### 1.2 Not every relationship type moves

**The standard.** Every Transitional relationship type has a Strict one.

**The files.** Two stay in the Transitional spelling inside Office's Strict
files. Excel writes the `vmlDrawing` relationship (the comment shapes) as
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing`,
and Word writes `schemaLibrary` the same way. Two others change their
spelling as well as their prefix: `extended-properties` becomes
`extendedProperties`, and `custom-properties` becomes `customProperties`.

**Rutba Office.** The mapping table in `packages/ooxml/src/strict.js`
carries the two renames and leaves `vmlDrawing` and `schemaLibrary` as
Office does, so a Strict file we write uses the same types as one Office
writes.

### 1.3 The conformance attribute

**The standard.** The main part's root may say `conformance="strict"` or
`conformance="transitional"`, the default being Transitional.

**The files.** Excel puts `conformance="strict"` on `<workbook>`, and
PowerPoint on `<p:presentation>`. Word puts it in its own namespace, as
`w:conformance="strict"` on `<w:document>`. The usual files leave it out.

**Rutba Office.** On reading, the attribute is set aside; on writing Strict,
it is added to the main part as each application spells it.

### 1.4 Fallbacks in a Strict file carry Transitional namespaces

**The standard.** MCE lets a part offer alternatives in
`mc:AlternateContent`. Nothing ties the namespaces of a fallback to the
conformance class of the file.

**The files.** In PowerPoint's Strict deck, the slide transitions are a
`p14` choice whose `mc:Fallback` declares the Transitional PresentationML,
DrawingML and relationship namespaces again, inside a part that is otherwise
Strict.

**Rutba Office.** The structural check in `tests/strict-ooxml.test.js` reads
a part with its `mc:AlternateContent` and `extLst` blocks taken out before
it looks for Transitional namespaces, so Office's own Strict files pass it.
Those blocks are kept, not rewritten.

## 2. Values the two forms write differently

### 2.1 Percentages in drawings

**The standard.** Transitional DrawingML writes a percentage as a whole
number of thousandths of a percent (`60000` is 60 per cent). Strict writes it
as a number and a percent sign (`60%`).

**The files.** PowerPoint and Excel write `<a:alpha val="60%"/>`,
`<a:gs pos="0%"/>` and `<a:spcPct val="90%"/>` in Strict. Text sizes (`sz`
on `<a:defRPr>`) stay plain numbers in both forms, as they are hundredths of
a point, not percentages.

**Rutba Office.** On reading, any unprefixed attribute of a DrawingML or
PresentationML element whose value is a number and a percent sign is turned
into thousandths. On writing Strict, only the attributes the schema types as
percentages are turned back (the list is `PERCENT_ATTRS` in `strict.js`), so
a text attribute that happens to hold a number is never touched.

### 2.2 Measures in documents

**The standard.** Strict WordprocessingML lets a length be written with a
unit (`36pt`, `2.54cm`, `1in`). Transitional writes twentieths of a point
for most lengths and half-points for font sizes.

**The files.** Word writes indents with a unit (`<w:ind w:start="36pt"/>`)
and table widths with one too, while keeping the type that names the old
unit: `<w:tblW w:w="451.30pt" w:type="dxa"/>`. Font sizes stay half-point
numbers (`<w:sz w:val="22"/>`).

**Rutba Office.** On reading, a length with a unit is turned into
twentieths of a point, or into half-points for the font size attributes. On
writing, plain numbers are kept: Strict allows them, and they are what the
engines and Word itself read.

### 2.3 Percentages in documents

**The standard.** A Transitional table width of type `pct` is in fiftieths
of a percent (`5000` is the full width). Strict writes `100%`. The same
change applies to the zoom setting.

**The files.** Word's Strict settings say `<w:zoom w:percent="60%"/>`.

**Rutba Office.** Read as fiftieths for a `pct` width and as a whole number
for the zoom; written back with the percent sign in a Strict save.

### 2.4 Start and end instead of left and right

**The standard.** Strict names the sides of a paragraph, a table and a
border by reading direction: `start` and `end`. Transitional uses `left` and
`right`. This covers indents, alignment, tabs, table and cell borders, and
cell margins in documents, and the borders in a workbook's styles.

**The files.** Word writes `<w:ind w:start="36pt" w:end="36pt"/>`; Excel
writes `<border><start style="thin">...</start><end .../>`.

**Rutba Office.** Each is renamed on reading and renamed back on writing
Strict, in the containers where the schema has them.

### 2.5 A font's character set

**The standard.** Transitional says a font's character set by its Windows
number, in hex (`<w:charset w:val="00"/>`). Strict says it by name
(`w:characterSet`).

**The files.** Word writes `<w:charset w:characterSet="iso-8859-1"/>`.

**Rutba Office.** The common sets are mapped both ways by a table in
`strict.js`. A name the table does not hold is kept as it came; the font is
still found by its name.

### 2.6 Dates in cells

**The standard.** Strict SpreadsheetML lets a cell hold an ISO 8601 date
(`t="d"`). A Transitional cell holds a serial number in the workbook's date
system. Strict also adds `dateCompatibility` on `<workbookPr>`: when it is
on, serial numbers follow the 1900 system with its day for 29 February 1900,
which never existed.

**The files.** Excel writes dates as `<c t="d"><v>2026-09-10</v></c>` and a
time on its own as `<v>18:00:00.000</v>`, and sets
`dateCompatibility="0"`. The usual workbook holds the same cells as serial
numbers.

**Rutba Office.** An ISO date is turned into the serial number the workbook's
date system gives it, with the 1904 system and `dateCompatibility` taken
into account, so formulas and number formats see what they see in the usual
file. A date cell we edit is written back as a serial number, which Strict
also allows. A sheet we do not edit keeps Excel's ISO dates, byte for byte.

### 2.7 Shapes and text boxes in documents

**The standard.** Strict has no `wps` or `wpg` namespaces: a shape or group
in a document is in the WordprocessingML Drawing namespace itself, as
`wp:wsp` and `wp:wgp`.

**The files.** Word writes `<wp:wsp>` and puts the text of a text box in
`<wne:txbxContent>`, in the Word 2006 namespace that it also adds to
`mc:Ignorable`. The usual file has `wps:wsp` and `w:txbxContent`, inside an
`mc:Choice Requires="wps"` with a VML fallback.

**Rutba Office.** Renamed both ways. A Strict save keeps the DrawingML
choice and drops the VML fallback, and leaves out `w:hdrShapeDefaults` and
`w:shapeDefaults`, which are VML.

## 3. VML in Strict files

**The standard.** VML belongs to Part 4, Transitional. A Strict file may not
use it except through MCE, where a reader that does not understand it can
pass over it.

**The files.** Office still writes VML in its Strict files, behind a choice:

- Word: the header watermark is `mc:Choice Requires="v"` holding the VML
  shape, with an `mc:Fallback` holding the same picture as a DrawingML
  anchor.
- Excel: a sheet's `<legacyDrawing>` (the shape of each comment) is inside
  `mc:Choice Requires="v"` with no fallback. The comment's place is given
  again for readers without VML: the drawing part has an empty VML choice
  and a fallback with a hidden DrawingML shape, and the comments part an
  empty choice and a fallback with `<commentPr>` and its anchor.

How Excel reads a Strict workbook matters to any other writer. Tried through
COM, Excel passes over a choice that requires VML, though it writes one, and
takes the comment from the DrawingML fallback: a Strict workbook whose
`<legacyDrawing>` is behind a choice, with no DrawingML fallback, opens with
its comments gone. The same `<legacyDrawing>`
written without the choice is read, and the comments are kept. Word, by
contrast, reads its watermark whether the VML is behind a choice or not.

**Rutba Office.** On reading, a choice that requires VML is taken, so the
engines see what the usual file shows. On writing Strict:

- documents: `<w:pict>` and `<w:object>` go behind `mc:Choice Requires="v"`,
  with no DrawingML fallback made for them;
- workbooks: `<legacyDrawing>` is written without a choice, since that is
  the form Excel reads; no DrawingML fallback or `<commentPr>` is made.

A Strict save from Rutba Office therefore carries VML where Office does,
and the structural check allows a bare `<legacyDrawing>` and nothing else.

## 4. Microsoft's extension namespaces

**The standard.** Part 3 lets a part use namespaces a reader may not know,
listed in `mc:Ignorable`, and Part 1 gives many elements an `extLst` of
`ext` blocks, each named by a `uri`. A reader that does not know a
namespace or an extension passes over it.

**The files.** Office uses both in every file, in the usual form and in
Strict alike. The namespaces are Microsoft's own (`schemas.microsoft.com`),
and they do not move to `purl.oclc.org` in a Strict file.

| File | Ignorable prefixes, across the parts | Extensions (`extLst` uris) | Choices |
|---|---|---|---|
| Document, usual | `cr` `w14` `w15` `w16` `w16cex` `w16cid` `w16du` `w16sdtdh` `w16sdtfl` `w16se` `wp14` | 2 | `wps` (shapes) |
| Document, Strict | the same, and `rd` `w16cei` `wne` | 4 | `v` (watermark) |
| Workbook, usual | `x14ac` `x15` `x16r2` `xr` `xr2` `xr3` `xr6` `xr10` | 14 | `x15` (workbook), `c14` (charts) |
| Workbook, Strict | the same, and `a14` | 17 | `x15`, `c14`, `v` (comments) |
| Deck, usual and Strict | none on the root | 13 and 14 | `p14` (transitions), `c14` (charts) |

Some of what these carry:

- `w14`, `w15` and the `w16` family: Word 2010 and later features, such as
  paragraph ids, comment threads and their dates, content control
  appearance and colours, and text effects. `wp14` gives drawings their
  relative sizes and ids.
- `x14ac`, `x15` and the `xr` family: Excel's extra row and column details,
  revision ids for co-authoring, and the workbook's `x15ac:absPath`, the
  folder the file was last saved in on the machine that saved it.
- `a14`, `a16`: picture effects and the creation ids of shapes.
- `c14` and later: chart styles, where `c:style` sits in a choice with a
  plain `c:style` fallback.
- `p14`, `p15`: transitions, sections and guides in decks.
- Dynamic arrays: Excel 365 marks a spilled formula with a `cm` attribute
  pointing at the workbook's metadata part. A reader that does not know it
  reads an ordinary array formula.

**Rutba Office.** It rewrites only the parts it edits. A part that is not
edited is saved as the bytes it came as, extensions and all, in either
form. A part that is edited keeps the extension elements it holds. The Strict
mapping changes no Microsoft namespace. The structural check accepts a prefix
in a Strict part only when it is one of ISO's or is listed in that part's
`mc:Ignorable`.

## 5. Other places the applications depart from the standard

### 5.1 Word's Restrict Editing password

**The standard.** `w:documentProtection` holds a salted, repeated hash of
the password.

**The files.** Word hashes not the password but the 32-bit key Word 97 made
from it, written as eight hex digits, low byte first. Microsoft documents
this in [MS-OI29500].

**Rutba Office.** `packages/ooxml/src/protection.js` makes and checks the
hash as Word does, so a password set in either application opens in the
other.

### 5.2 Element order

**The standard.** The children of most elements are a sequence in a fixed
order.

**The files.** Excel enforces the order strictly: a worksheet element out
of place makes Excel repair the file by throwing the whole part away.

**Rutba Office.** Everything written after a sheet's cells goes in through
one place in `packages/ooxml/src/workbook.js` that takes its position from
the schema's order. The document engine does the same for paragraph, run,
table and section properties.

### 5.3 Notes parts

**The standard.** The footnotes and endnotes parts hold the notes; the
separator and continuation separator are optional.

**The files.** Word always writes the two separators first, at ids -1 and 0,
and repairs a notes part that lacks them when it opens the file.

**Rutba Office.** A notes part it creates starts with both separators.

## 6. What Strict support does not cover yet

- Dates we edit are written as serial numbers, not as ISO 8601 text.
- VML in a document is put behind a choice with no DrawingML fallback made
  for it, and a workbook's comment shapes are written as Excel reads them,
  without the DrawingML fallback or `<commentPr>` Excel writes.
- Form controls, ActiveX controls and embedded OLE objects are kept as they
  came; they are not converted.
- Drawing canvases (`wpc`) and ink are not mapped.
- The character set table covers the common sets only.
- Office keeps chart values as plain numbers in Strict, so charts need no
  mapping; a chart from another writer that uses percent signs in chart
  markup is read as written.

Files from other applications are measured by the corpus run described in
[COMPATIBILITY.md](COMPATIBILITY.md), and what it finds is added here.
