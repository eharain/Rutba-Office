# @rutba/pdf

Writing PDFs, with nothing installed to do it.

The suite could *read* a PDF and could not write one. So finance could email
an invoice and never attach it, and "send me that as a PDF" — the most
ordinary request a customer makes of an accounts department — had no answer.

```js
const { renderPaper } = require('@rutba/pdf');

const bytes = renderPaper({
  title: 'Invoice',
  reference: 'INV-0001',
  letterhead: { name: 'Tech-Style (UK) Ltd', lines: ['London WC2H 9JQ'] },
  party: { label: 'Billed to', name: 'Acme (Northern) Ltd' },
  meta: [['Invoice date', '2026-08-01'], ['Due date', '2026-08-31']],
  table: {
    columns: [
      { key: 'description', label: 'Description', flex: 3 },
      { key: 'amount', label: 'Amount (GBP)', width: 90, align: 'right' },
    ],
    rows: [{ description: 'Consultancy', amount: '100.00' }],
  },
  totals: [['Subtotal', '100.00'], ['Balance due (GBP)', '120.00']],
});
```

## Why there is no dependency

The two usual answers are a headless browser — a Chromium download and a
process per document — or a drawing library that still leaves you to write
the layout. Neither is much use for the thing actually wanted: a page of A4
that looks like an invoice.

What makes the short way work is that **the base-14 fonts are the reader's**.
Their advance widths are fixed by the format, so measuring with the AFM table
in `metrics.js` gives the same answer the viewer will. No font files, no
embedding, and — unlike the HTML path, which has to guess at what a browser
will do — measurement and rendering agree by construction.

`test/readback.test.js` puts that to `pdfjs-dist`, the reader the suite
already ships: our widths and its widths agree to within a hundredth of a
point on real strings.

## The four layers

| file | what it knows |
| --- | --- |
| `metrics.js` | how wide a glyph is |
| `encoding.js` | which byte a character becomes (WinAnsi) |
| `document.js` | objects, pages, the cross-reference table |
| `sheet.js` | a cursor, wrapping, tables, page breaks, footers |
| `paper.js` | what a business document looks like |

Only `document.js` knows that PDF measures y upwards. Everything above it
counts downwards from the top of the page, which is how people think about
paper.

## What it will not do

- **No Times, no Symbol.** Adding a font means adding its width table. A
  font whose widths are wrong fails silently, in the reader's viewer and
  nowhere else.
- **No Greek, Cyrillic or CJK.** WinAnsi has no glyphs for them, so they
  become `?` — visibly a substitution, rather than a name quietly shortened.
  Real Unicode means an embedded font with a CMap: a much bigger machine,
  worth building when somebody needs it and not before.
- **No compression** — except image pixel streams. An invoice is a few
  kilobytes and stays greppable; raw RGB would triple a file for nothing, so
  images alone are deflated at assembly.
- **Images: PNG, the honest subset.** Sign asked (a drawn signature on the
  certificate of completion), and it was indeed an XObject and a `Do`:
  `images.js` decodes 8-bit non-interlaced PNGs (grey, RGB, palette, and both
  alpha flavours — alpha becomes an `/SMask`), `PdfDocument.addImage` +
  `Page.image`/`Sheet.image` draw them, and `renderPaper` renders a
  `signatures` block of adopted marks. 16-bit, low-bit-depth and Adam7
  interlaced PNGs are refused by name rather than decoded wrongly.

## Tests

```bash
npm test --workspace @rutba/pdf
```

48 checks. The interesting ones: every cross-reference offset points at the
object it claims, a bracket in a company name does not corrupt the page, a
table takes its headings onto the next sheet, and a cell too tall for the
page continues rather than being drawn off the bottom.
