# Text recognition

Reading the words in a picture or a scanned page, on this computer, with
the network off. This page records how well the engine reads each of the
suite's six languages, measured before any of it was built into the apps,
and how to measure it again.

## The engine

Tesseract 5 compiled to WebAssembly, as tesseract.js 7 ships it
(tesseract.js-core 7.0.0, Apache-2.0), with the "fast" LSTM models of
tessdata_fast (Apache-2.0) for English (`eng`), Urdu (`urd`), Arabic
(`ara`), Hindi (`hin`), Chinese, Simplified (`chi_sim`) and Japanese
(`jpn`). Both are vendored in the repository and ship inside the suite, so
nothing is fetched at first use or ever. The SIMD, LSTM-only build of the
engine is used: 2.9 MB of WebAssembly. The six models are 13.0 MB:

| Model | Size |
| :-- | --: |
| eng | 4.1 MB |
| chi_sim | 2.5 MB |
| jpn | 2.5 MB |
| ara | 1.4 MB |
| urd | 1.4 MB |
| hin | 1.1 MB |

## Languages

Text recognition follows the languages the suite ships: the first six are
the suite's English, Urdu, Arabic and Hindi and the Chinese and Japanese it
is gaining, and each language the suite adds gains recognition too. The
list is one table, `packages/ocr/src/languages.js`: a row a language with
the suite's own code for it, its model's file name, its script, its
direction and whether it is beta. Adding a language is a row and its model
file; `tools/vendor-ocr.mjs` takes the models it copies from that table,
the language picker in the windows lists its rows, a window speaking the
language suggests it first, and the package's tests check every row has its
model and every model shipped has its row.

## Where it runs

`@rutba/ocr` (packages/ocr) answers `recognize(image, { languages })` with
the text, its lines and its words, each word with its box in the picture's
pixels and a confidence from 0 to 100. Pages are read on a pool of worker
threads (`src/host.js`): one page a thread, at most one thread a core and
never more than four, the rest waiting in a queue in the order asked, so a
long scan holds nothing else up; a job can be cancelled while it waits or
while it is read. Nothing is read on the main process or on a window's
thread. The windows reach it through the `ocr` namespace
(apps/desktop/main/ocr.js): `recognize`, `cancel`, `languages` and
`stats`. A window that closes cancels what it was waiting for.

The engine and the models are vendored in `packages/ocr/vendor` by
`tools/vendor-ocr.mjs` and unpacked beside the installed app's archive,
where the threads read them. The engine is handed its WebAssembly as bytes
and its models are written into its own memory, so it has nowhere to fetch
anything from; the package's tests run a recognition with every way to the
network refused, in the process and its threads, and check that none was
tried.

## In the apps

**PDF: Recognise Text** (Home and Pages, and the thumbnails' menu). A
dialog chooses the pages (all, the selected ones, or this one) and the
languages, Urdu marked beta with a sentence on why. Pages that already have
text are left as they are. Each scanned page is drawn by pdf.js at up to
300 dots per inch (its long side at most 4,200 pixels) and read on the pool;
each word's box is turned into the page's own coordinates by the viewport
that drew it, so a rotated or cropped page comes out right. The main process
then writes an invisible text layer (render mode 3) as an incremental update
(`packages/pdf/src/textlayer.js`): each page redefined with one content
stream more and a font for the layer, the original bytes left as the start
of the file, and the update's cross-references written as a table or as a
stream to match the file's own. The font is a Type0 font over an embedded
TrueType program of empty glyphs half an em wide, whose codes are the text's
UTF-16 units mapped back by its ToUnicode map; each word is stretched to its
width on the page, so selecting it highlights the word, and Urdu and Arabic
words are written in the order a page shows them, as PDFs write those
scripts. Find, selecting and copying then work on the scans; Undo takes the
layer away; a Save of the recognised file as it is writes its original
bytes with the layer after them. Refused by name: files protected with a
password, and files whose cross-references had to be rebuilt (save a copy
first).

**Pictures and Image: Copy Text from Picture** (Pictures: the Home tab
and the viewer's bar under the picture, where the ribbon folds away, and
the picture's menu; Image: the Home tab). The picture is read once on the
pool, every word goes on the clipboard in reading order, and the lines are
laid over the picture as invisible text, outlined faintly, where a drag
selects and copies words or a line alone; a bar over the picture says how
many lines were found, with Copy All and Hide. The lines follow the picture
as it is zoomed, scrolled or turned in Pictures; in Image the picture is read
as it is edited, and any edit takes the words away. Languages chooses which
of the six are read (Urdu marked beta), remembered for every window.

**Worksheets: Data from Picture** (Data, From Picture: Picture From File or
Picture From Clipboard). The picture is read with the table pass (below)
and shown beside a preview grid of its cells, each one editable; cells the
reading was unsure of (a word under 75 of 100) are marked to check. Insert
Data puts the grid in at the selected cell as one step Undo takes back,
numbers as numbers. The languages can be changed and the picture read again
in the dialog.

**The table pass** (`recognize(image, { table: true })`). The engine's page
segmentation is made for pages, not grids: on a table set with space alone
it dropped a narrow column of figures whole, and its sparse mode, which
keeps the column, dropped a figure standing alone. So a table is read both
ways and the words merged (one word where two overlap, the surer kept),
the grid is found from the table's ruling lines when they make one and
otherwise from its lines of words and the white every row leaves, and a
cell that came out empty though there is ink in it is read again on its own
as one line. Each cell carries its least sure word's confidence. Both test
tables (one ruled, one not) come back cell for cell.

## Page analysis and layout of our own

Around the engine, in `packages/ocr/src/analyse.js` and `layout.js`,
classical image processing and geometry that the engine tests prove on
pages drawn with known answers:

- **Ink and paper**: Otsu's threshold for evenly lit pages, Sauvola's
  local threshold where the paper's brightness varies across the page (a
  shadow down a fold, a photo lit from one side), chosen by measuring the
  paper in a six by six grid. On a page lit from one side Otsu gets more
  than a tenth of the pixels wrong; Sauvola under 2 percent.
- **Skew**: the rows of ink projected at angles from minus to plus eight
  degrees, coarse then fine; the angle where text lines fall into the
  sharpest peaks is the page's turn, measured to a tenth of a degree.
- **Text height**: the median height of the letter-like blobs, each counted
  by its ink, so Arabic's dots and the parts of a Han character do not pull
  it down.
- **Ruling lines**: long runs of ink along rows and columns, short breaks
  bridged, thick blocks and letters refused.
- **Lines, columns and table cells**: words gathered into lines by their
  vertical overlap; columns from white strips no line crosses (a heading
  wider than the body is allowed); a table's rows and columns from its
  ruling lines when they make a grid, otherwise from its lines of words and
  the white every row leaves, each word in the cell its centre falls in,
  right to left for Urdu and Arabic.

Before reading, a page turned by two degrees or more is turned back level,
text smaller than 10 pixels is scaled up (a screenshot) and text larger
than 64 down, and an unevenly lit page is handed over as ink and paper.
After reading, the engine's lines are put in reading order by geometry
(columns, then lines top to bottom, right to left for Urdu and Arabic), a
line the engine split into two blocks is joined again, a word the engine
boxed as a sliver takes its line's height (judged where the engine read it,
and only when it is thin against the page's usual line too, so small type
the engine put on one line with a large heading keeps its own height), and
in Urdu and Arabic a Latin comma, semicolon or question mark after a letter
becomes the Arabic one.

Two of these were measured and changed: turning a page back by resampling
it cost more than the 1.2 degree tilt of the simulated scans did (Japanese
3.1 to 13.3 percent CER, Nastaliq 44.5 to 59.7), because the engine reads
a small tilt well itself, so only two degrees or more is turned back; and
scaling 300 dpi pages towards a smaller text height cost Hindi and
Nastaliq, so only very large text is scaled down.

## How it was measured

**Corpus.** `tools/ocr-corpus.js` sets four sentences of each language
(`tools/ocr-texts.json`: a notice, an invoice line with numbers, a request
and a description) as one page of 12 point type in a 6.5 inch column, drawn
by the suite's own renderer (Chromium in Electron), so the text of every
page is known exactly. Each page is made three ways: clean at 150 dots per
inch, clean at 300, and at 300 as a simulated scan (turned 1.2 degrees,
softened by half a pixel, on a paper grain, saved as a JPEG at quality 55).
Fonts: Times New Roman and Arial for English, Noto Naskh Arabic for Arabic
and for Urdu in Naskh, Noto Nastaliq Urdu for Urdu in Nastaliq, Noto Sans
Devanagari for Hindi, SimSun for Chinese and Yu Gothic for Japanese. 24
pages in all.

**Rates.** The character error rate (CER) is the edit distance between the
text read and the text drawn, over the length of the text drawn, counted in
Unicode code points after NFC and with runs of white space made one space.
The word error rate (WER) is the same over words. Chinese and Japanese are
written without spaces, so their CER ignores spaces and they have no WER.

**Time and memory.** `tools/ocr-measure.mjs` reads each page in a Node
process of its own, twice. "First" is the time from nothing to the first
page's text: starting the engine, loading the model and reading the page.
"Page" is the second read alone, the cost of each further page. Peak
memory is the process's peak resident set over both. The machine is the
reference laptop: an Intel Core i5-8350U (four cores, eight threads, 1.7
GHz), 16 GB of memory, Windows 11, Node 22. Other work ran on it at the same
time, so single times vary by a third or more; the medians are steadier.

## Results

Fast models, SIMD build. CER and WER in percent; time per further page.

| Page | Clean, 150 dpi | Clean, 300 dpi | Simulated scan, 300 dpi |
| :-- | :-- | :-- | :-- |
| English, serif | 0.2 / 2.2, 1.1 s | 0.0 / 0.0, 1.2 s | 0.0 / 0.0, 1.3 s |
| English, sans | 0.2 / 2.2, 0.8 s | 0.2 / 2.2, 1.2 s | 0.0 / 0.0, 1.5 s |
| Arabic, Naskh | 1.0 / 5.7, 0.8 s | 3.2 / 8.6, 1.0 s | 3.4 / 10.0, 1.1 s |
| Urdu, Naskh | 4.0 / 7.5, 1.1 s | 5.5 / 8.6, 1.1 s | 7.3 / 10.8, 1.1 s |
| Urdu, Nastaliq | 20.9 / 64.5, 0.7 s | 19.4 / 62.4, 0.9 s | 50.0 / 86.0, 1.5 s |
| Hindi | 0.0 / 0.0, 1.2 s | 0.0 / 0.0, 1.8 s | 0.0 / 0.0, 1.9 s |
| Chinese | 1.5, 1.3 s | 2.9, 1.4 s | 2.2, 1.5 s |
| Japanese | 0.0, 1.1 s | 0.0, 1.8 s | 3.1, 1.5 s |

- **Time.** A page of this size takes about a second once the engine is
  running (median 1.2 s) and about two seconds from a cold start (median
  1.8 s, of which starting the engine and loading a model is 0.2 to 0.6 s).
  The engine reads on one thread; the pool reads pages on up to four
  threads at once, so a long scan goes up to four times as fast.
- **Memory.** 100 to 135 MB of peak resident memory per process for one
  page, Node's own share included; the engine and one model hold roughly
  40 to 80 MB of it. Four engines reading at once would stay under half a
  gigabyte.
- **English, Hindi and Japanese** are read almost without error on clean
  pages and simulated scans alike.
- **Arabic and Urdu in Naskh** are read well. Most of the errors are
  punctuation: the Arabic comma read as a Latin comma, or lost. Without
  punctuation the rates are 0.0 to 2.3 percent (CER) for Arabic and 2.9 to
  6.1 percent for Urdu.
- **Urdu in Nastaliq**, the style most Urdu is printed in, is read poorly:
  one character in five is wrong on a clean page, and three words in five;
  half the characters on a simulated scan. Urdu recognition is therefore
  labelled beta in the suite, good on Naskh and on Urdu typed in a Naskh
  font, not yet for Nastaliq print. A recogniser of our own for Nastaliq is
  being prepared (below, "A model of our own").
- **Chinese** loses a character or two a page (1.5 to 2.9 percent).
- **150 or 300 dots per inch.** The models read type of 20 to 30 pixels best;
  at 300 dots per inch 12 point type is 50 pixels high and some scripts read
  a little worse than at 150, so scaling a page's text towards that size
  before reading is worth doing.

### Through @rutba/ocr

The same pages read as the suite reads them, through `@rutba/ocr` with the
analysis above (`tools/ocr-measure.mjs --ocr`), against the engine alone:

| Page | Engine alone | Through @rutba/ocr |
| :-- | :-- | :-- |
| Urdu, Naskh, 150 / 300 / scan | 4.0 / 5.5 / 7.3 (WER 7.5 / 8.6 / 10.8) | 3.1 / 4.7 / 6.4 (WER 3.2 / 5.4 / 7.5) |
| Urdu, Nastaliq, 150 / 300 / scan | 20.9 / 19.4 / 50.0 (WER 64.5 / 62.4 / 86.0) | 19.0 / 18.7 / 44.5 (WER 65.6 / 60.2 / 75.3) |
| Arabic, 150 / 300 / scan | 1.0 / 3.2 / 3.4 (WER 5.7 / 8.6 / 10.0) | 2.2 / 2.2 / 2.4 (WER 2.9 / 2.9 / 4.3) |
| English, Hindi | as above | the same, but for one English page at 150 dpi (0.6, one letter) |
| Chinese, 150 / 300 / scan | 1.5 / 2.9 / 2.2 | 2.2 / 3.7 / 2.2 |
| Japanese, 150 / 300 / scan | 0.0 / 0.0 / 3.1 | 0.0 / 2.6 / 3.1 |

Urdu and Arabic gain most, from the punctuation and the reading order.
Chinese and Japanese lose a character or two on clean pages: the engine
reads the picture as grey pixels from the suite rather than decoding the
colour PNG itself, and a few glyphs come out differently. A page read
through the package takes the same time (median 1.07 s a further page,
1.50 s the first) and up to 176 MB at peak, the analysis's copies of the
page included.

### The engine's builds

Median time per further page over the 24 pages, fast models:

| Build | WebAssembly | Page | First page | CER |
| :-- | --: | --: | --: | :-- |
| SIMD, LSTM only (shipped) | 2.9 MB | 1.18 s | 1.82 s | as above |
| relaxed SIMD, LSTM only | 2.9 MB | 1.17 s | 1.67 s | identical |
| no SIMD, LSTM only | 2.9 MB | 2.99 s | 3.30 s | identical |

SIMD makes reading two and a half times faster. Relaxed SIMD adds nothing
measurable on this processor, and the plain SIMD build runs on more
machines, so it is the one shipped.

### Fast and "best_int" models

The integer-quantised "best" models (`4.0.0_best_int` from tesseract.js's
model packages) were measured on the same pages with the same build:

| | eng | ara | hin | jpn | chi_sim, urd | Six together |
| :-- | :-- | :-- | :-- | :-- | :-- | --: |
| Size, fast | 4.1 MB | 1.4 MB | 1.1 MB | 2.5 MB | the same file | 13.0 MB |
| Size, best_int | 5.2 MB | 2.5 MB | 1.7 MB | 3.0 MB | the same file | 16.3 MB |
| CER, fast | 0.0 to 0.2 | 1.0 to 3.4 | 0.0 | 0.0 to 3.1 | | |
| CER, best_int | 0.0 | 1.0 to 1.2 | 0.0 to 0.5 | 0.5 to 4.6 | | |

The best_int models read English and Arabic slightly better, Japanese
slightly worse, at the same speed and 3.3 MB more. The fast models are
shipped; English and Arabic could move to best_int later if real scans
show the same gain.

### Real scans

Three printed pages scanned by others, public domain, each with a
transcription proofread by Wikisource's volunteers (validated or
proofread pages), read as a 1,800 to 1,920 pixel wide JPEG of the page:

| Page | Engine alone | Through @rutba/ocr |
| :-- | :-- | :-- |
| English: Alice's Adventures in Wonderland (1866 printing), page 3 | 2.3 / 3.4 | 2.3 / 3.4 |
| Hindi: आग और धुआं (Chatursen), page 148 | 2.7 / 12.5 | 2.8 / 12.5 |
| Urdu, Nastaliq: بیوی کی تربیت (Khwaja Hasan Nizami), page 22, lithographed | 50.4 / 90.4 | 50.2 / 89.3 |

CER / WER in percent. Each page's running head is not in its transcription
but is read, so a point or two of each rate is the head. The real pages
agree with the corpus: English and Hindi read well, and Nastaliq print is
not read reliably, which is why Urdu is beta. The pages are not kept in the
repository; `Page:` titles on en, hi and ur.wikisource.org find them.

### Not measured yet

- Photographs of pages (receipts, letters, a newspaper column) and more
  real scans in every language.
- A second engine for Chinese, Japanese and difficult layouts.

## A model of our own

Nastaliq is where the shipped model is weakest, so it is the first script
for a recogniser trained by us. The training data is made by the suite
itself: `tools/ocr-lines.js` sets lines of text in an OFL font (Noto
Nastaliq Urdu first) with the suite's renderer, so every line is shaped as
the suite shows it, and spoils them as print and scans do (a slight turn,
softening, paper grain, a shadow, drawn small and enlarged, JPEG loss), each
choice from a seed so a set can be made again exactly. The text can be a
file, or the suite's own catalogue for a language (`--catalogue ur`: our own
translations, so nothing of anyone else's is trained on). Each line is a PNG
beside its text in logical order, the form line recognisers and Tesseract's
own training read, with a manifest of every line's settings.

`tools/ocr-measure.mjs --lines <folder>` reads such a folder line by line
through @rutba/ocr and gives the rates over the whole set. The shipped
`urd` model on lines the generator made:

| Lines | CER | WER | Lines read without an error |
| :-- | --: | --: | --: |
| Nastaliq, spoiled as scans (390) | 37.3 | 78.1 | 2.8 |
| Nastaliq, clean (300) | 29.3 | 73.2 | 3.7 |
| Naskh, spoiled as scans (300) | 15.9 | 23.6 | 42.7 |

A line read alone is harder than a page (no neighbouring lines to steady
the layout, and many catalogue lines are two or three words), so these
rates are higher than the page corpus's; they are what a line model is
compared with. A model of our own replaces it, behind the same `recognize()`, only when it
reads these sets and a set of real printed Nastaliq pages clearly better,
with its rates published here.

## Bills, invoices and receipts

`recognize(image, { document: true })` reads a page as a document and
answers `document` beside the text: what kind of page it is (invoice,
receipt, letter, form or other), its fields with their boxes and
confidence (who it is from, the number, the date, the due date, the
subtotal, the tax and its rate, the total and the currency), its lines of
items, whether its sums agree, and its tick boxes, ticked or not. The rules
are in `packages/ocr/src/fields.js` and `checkboxes.js`; the engine is the
same, and nothing is learnt or fetched. No window uses it yet.

How a page is read:

- **Twice, then merged.** The page is read as a page and again as sparse
  text, and the words of both are merged as the table pass merges them: the
  page reading keeps the figures of a table that sparse text drops, sparse
  text keeps a label or a figure standing alone that the page reading
  drops. Words the sparse reading is unsure of (under 40 of 100) are left
  out, since it also reads a dashed rule or the grain of the paper as
  letters.
- **Rows and segments.** The words are levelled (the median slope of the
  engine's lines taken out, for a page photographed a little turned) and
  gathered into rows by the middle band of each row's words, so a large
  title beside two lines of small type takes neither into its row. Each row
  is cut into segments where a gap wider than a text height opens.
- **Labels and values.** A label is the start of a segment matching a
  phrase of the field's dictionary: six languages, several wordings each, a
  misread letter allowed in five. Its value is the rest of its segment,
  else the next segment along the row in the page's direction (leftwards
  in Urdu and Arabic), else the segment below it, never past another label.
- **Reading figures again.** Every figure beside a label and every item's
  amount is cropped and read again as one line, by English's model (which
  reads Western digits best, also on an Arabic or Hindi page) and by the
  page's own. Where a label has no value beside it at all, the stretch of
  its row where the value would lie is read.
- **Locales.** Amounts are read by their own separators (a comma or a point
  for the decimals, a space or an apostrophe between thousands, Indian
  grouping, the Arabic decimal and thousands marks), with Arabic-Indic,
  Urdu and Devanagari digits and full-width ones, and a figure the engine
  spaced about its separators joined again. Dates are read year first, day
  first (month first when the first number cannot be a month, or when asked
  for), with the month's name in English, Arabic (both sets of names),
  Urdu or Hindi, and as 年月日.
- **The sums decide.** The candidates for the subtotal, the tax and the
  total (each reading of each, a lost decimal point put back, a rupee sign
  read as a leading 7 or 2 taken off) and each item's readings are tried
  together; the choice where the items add up to the subtotal, the subtotal
  and the tax to the total and the tax is the rate of the subtotal wins. A
  tax the page lost is the total less the subtotal when that is the rate.
  The sums are said to agree only when the total itself is confirmed, by its
  parts or, on a page with neither subtotal nor tax, by its items.
- **Who it is from.** A supplier label's value where there is one, else the
  largest writing near the top in the page's own script that is not the
  title, a label, a value, the customer under "Bill to", or writing the
  reading was unsure of.
- **The kind of page.** Scores from its words (invoice, receipt and form
  titles in six languages, a letter's greeting and closing), its labels (a
  due date, "Bill to", cash and change), a till roll's shape, a time of day,
  long lines of prose, tick boxes and whether a total was found.
- **Tick boxes.** A blob of ink with four straight, thin sides of a like
  length, square corners, paper to its left and right and about the size of
  the text; ticked when ink fills more than a twentieth of its inside. A
  box much smaller than the form's others, a bullet, a round or open letter
  and a Chinese or Japanese character drawn as a square are not boxes; an
  empty box the engine read as 口 is.

### How it was measured

`tools/docread-corpus.js` draws 90 pages with the suite's renderer at 200
dots per inch, fifteen in each of the six languages: four invoices laid out
as the gallery's Word invoice (the heading, the bill to, number and date
block, the items with the sums under them, the heading row white on the
accent colour), three as its Worksheets invoice (labels in one column,
values beside them, a gridded items table), four till rolls, two letters
and two forms with six tick boxes each. `tools/docread-docs.mjs` chooses
each page's wording, font, figures, dates, currency and digits from a seed
and writes what the page holds beside its picture. Fonts: Arial, Times New
Roman, Calibri, Georgia, Consolas and Courier New for English; Noto Naskh
Arabic, Tahoma and Arial for Arabic; Noto Naskh Arabic, Tahoma and Noto
Nastaliq Urdu for Urdu; Noto Sans Devanagari and Nirmala UI for Hindi;
SimSun and Microsoft YaHei for Chinese; Yu Gothic and MS Gothic for
Japanese. Six of the 66 pages with figures print them in their own digits
(two each in Arabic-Indic, Urdu and Devanagari). Each page is clean, a scan (turned about a
degree, softened, on a grain of paper, a JPEG at quality 55) or a photo
(turned two and a half degrees, softened more, darkening towards one
corner, a JPEG at quality 60).

`tools/docread-measure.mjs` reads every page through `@rutba/ocr` with the
page's language and English, as a window speaking it would, and compares
each field with the truth: amounts to the cent (the yen for yen), dates to
the day, the number by its characters, the supplier by its text with
spaces and punctuation aside, and the items by their count and each line's
amount. A field is right, wrong or missed; the rates below are the share
right. 66 pages carry figures (42 invoices, 24 receipts).

### Results

Read as a document (`document: true`), percent right:

| Language | Supplier | Number | Date | Due date | Subtotal | Tax | Total | Items | Kind |
| :-- | --: | --: | --: | --: | --: | --: | --: | --: | --: |
| All | 76 | 88 | 71 | 64 | 73 | 76 | 77 | 47 | 99 |
| English | 82 | 100 | 91 | 100 | 100 | 100 | 100 | 82 | 100 |
| Arabic | 55 | 82 | 73 | 57 | 45 | 36 | 45 | 27 | 100 |
| Urdu | 55 | 64 | 45 | 29 | 27 | 36 | 36 | 9 | 93 |
| Hindi | 91 | 100 | 73 | 43 | 91 | 82 | 100 | 36 | 100 |
| Chinese | 91 | 82 | 73 | 57 | 91 | 100 | 91 | 55 | 100 |
| Japanese | 82 | 100 | 73 | 100 | 82 | 100 | 91 | 73 | 100 |

By how the page was made, all languages:

| Page | Supplier | Number | Date | Due date | Subtotal | Tax | Total | Items |
| :-- | --: | --: | --: | --: | --: | --: | --: | --: |
| Clean | 83 | 100 | 72 | 67 | 67 | 72 | 72 | 56 |
| Scan | 83 | 80 | 67 | 56 | 77 | 77 | 73 | 57 |
| Photo | 56 | 89 | 78 | 75 | 72 | 78 | 89 | 22 |
| Western digits | 75 | 90 | 75 | 67 | 77 | 80 | 82 | 52 |
| Arabic-Indic or Urdu digits (4 pages) | | | 0 | 0 | 0 | 0 | 0 | 0 |
| Western digits, Nastaliq aside (57 pages) | 79 | 93 | 79 | 72 | 81 | 84 | 86 | 54 |

Read once as a page (`recognize` alone, the same rules over its words, no
second readings), the amounts are much worse outside English and the
Chinese and Japanese pages:

| Language | Subtotal | Tax | Total | Items |
| :-- | --: | --: | --: | --: |
| All | 53 | 56 | 61 | 38 |
| Arabic | 9 | 0 | 18 | 9 |
| Urdu | 18 | 27 | 18 | 0 |
| Hindi | 18 | 18 | 45 | 0 |

- **The sums as a signal.** 34 of the 66 pages had sums that agree, and on
  every one of them the subtotal, the tax and the total were all right.
  Of the 32 whose sums did not agree, 17 still had the right total. A
  window can therefore show an agreeing page's figures as read, and ask for
  a look at the rest.
- **What the engine cannot read.** The shipped models do not read
  Arabic-Indic or Urdu digits: ٧ comes back as V, ٤ as £, ٢ as Y, on clean
  pages as on scans, by the Arabic model and the English one alike. No rule
  recovers them, so a page printed in those digits gives no figures and no
  dates. A small recogniser of our own for the twenty digit shapes, trained
  on pages the suite draws, is the way to them; it is a new model and is
  not part of this work. Urdu set in Nastaliq gives nothing reliable either,
  as the page rates above lead one to expect.
- **Where it is good.** English pages are read whole; Chinese, Japanese and
  Hindi pages give their totals nine times in ten or more, and Hindi's
  model reads Devanagari digits (both such pages right).
- **Items** are the weakest field: one line misread in five or six spoils a
  page's items, and on photos the heading row (white on colour, darkened)
  is often lost.
- **Kind of page**: 89 of 90 right; the miss is an Urdu Nastaliq invoice read
  as other.
- **Tick boxes**: all 72 boxes on the 12 forms found and read as ticked or
  not (ticks, crosses and filled dots), clean and scanned, no other square
  taken for a box.
- **Time.** On the reference laptop with four pages read at once, the median
  A4 page took 7.2 s as a document against 4.0 s read once; a till roll 7.2 s
  against 1.9 s. Most of the difference on pages not in English is the
  engine loading English's model for the second readings and the page's
  model again after them.

No real receipts were measured: the public-domain scans that could be used
were not fetched for this round. That is the next measurement, with
photographs of real till rolls in each language.

### What the results support in the windows

- **Worksheets: an expense row from a receipt.** A receipt or an invoice
  read into one row (date, supplier, number, subtotal, tax, total,
  currency), the figures of a page whose sums agree filled as read and the
  others marked to check, as Data from Picture marks unsure cells. Supported
  for pages in English, Chinese, Japanese and Hindi, and for Arabic and Urdu
  pages in Western digits marked to check.
- **Mail: the fields of an attached invoice.** A note over a PDF or picture
  attachment that reads as an invoice (kind 99 percent right): who it is
  from, the total and the due date, each with its box to show where it was
  read. The due date is the weakest of the three outside English.
- **PDF: fields on a scanned form or invoice.** After Recognise Text, the
  document's fields and tick boxes offered as form values; tick boxes are
  reliable enough to fill, the other fields to propose.
- Not yet: items as a table into Worksheets (one page in two right), and
  anything on pages printed in Arabic-Indic or Urdu digits.

## Measuring again

```bash
npm install --prefix <scratch> tesseract.js@7
npx electron tools/ocr-corpus.js --fonts <folder with the Noto fonts> --out <corpus>
node tools/ocr-measure.mjs --corpus <corpus> --core <scratch>/node_modules/tesseract.js/node_modules/tesseract.js-core --models <tessdata_fast folder> --out results.json
```

Lines for training and testing:

```bash
npx electron tools/ocr-lines.js --catalogue ur --fonts <fonts> --language urd --font NotoNastaliqUrdu.ttf --out <lines> --count 4000 --test 0.1 --seed 1
node tools/ocr-measure.mjs --lines <lines>/test --language urd
```

Bills, invoices and receipts (the readings are kept beside the pages, so a
change to the rules is measured again without reading them; `--again` reads
them anew, `--mode auto` reads each page once, `--misses` lists every field
not read right):

```bash
npx electron tools/docread-corpus.js --fonts <folder with the Noto fonts> --out <corpus>
node tools/docread-measure.mjs --corpus <corpus> --mode document --threads 4 --out results.json
```

The Noto fonts (SIL Open Font Licence) are not kept in the repository:
`NotoNaskhArabic.ttf`, `NotoNastaliqUrdu.ttf` and `NotoSansDevanagari.ttf`,
from the Noto project or Google Fonts. `--build` picks the engine's build
(`simd-lstm`, `relaxedsimd-lstm`, `lstm`) and `--only` a page name's start.
