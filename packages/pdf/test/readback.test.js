'use strict';

/**
 * Read our own output back with somebody else's parser.
 *
 * Every other test in this package checks that we wrote what we meant to.
 * This one checks the only thing that actually matters to a customer: that a
 * real PDF reader opens the file and finds the document in it. The suite
 * already carries one - `pdfjs-dist`, for the viewer - so the writer can be
 * judged by the reader without adding anything.
 *
 * The width comparison at the end is the strongest claim this package makes,
 * tested the only honest way. `metrics.js` asserts that our advance widths
 * ARE the viewer's. pdfjs computes its own from its own base-14 tables, so
 * if the two agree to a hundredth of a point on real strings, that assertion
 * is true; if the table had a transposed digit, this is where it surfaces.
 *
 * Skipped rather than failed where pdfjs cannot be resolved, because the
 * package itself has no dependencies and must stay testable without one.
 */

const test = require('node:test');
const assert = require('node:assert');
const { PdfDocument } = require('../src/document');
const { renderPaper } = require('../src/paper');

let pdfjs = null;
const load = async () => {
  if (pdfjs) return pdfjs;
  try {
    pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  } catch {
    pdfjs = false;
  }
  return pdfjs;
};

const open = async (buffer) => {
  const lib = await load();
  if (!lib) return null;
  return lib.getDocument({ data: new Uint8Array(buffer), useSystemFonts: false }).promise;
};

const invoice = (rowCount) => renderPaper({
  title: 'Invoice',
  reference: 'INV-2026-0041',
  letterhead: {
    name: 'Tech-Style (UK) Ltd',
    lines: ['71-75 Shelton Street', 'London WC2H 9JQ'],
    contact: ['accounts@example.test'],
  },
  party: { label: 'Billed to', name: 'Müller & Sons (Northern) Ltd', lines: ['Manchester M1 2AB'] },
  meta: [['Invoice date', '2026-08-01'], ['Due date', '2026-08-31']],
  intro: 'Dear Müller & Sons, please find the invoice below.',
  table: {
    columns: [
      { key: 'd', label: 'Description', flex: 3 },
      { key: 'a', label: 'Amount (GBP)', width: 90, align: 'right' },
    ],
    rows: Array.from({ length: rowCount }, (_, i) => ({
      d: `Consultancy — week ${i + 1} at the “agreed” rate`,
      a: (100 * (i + 1)).toFixed(2),
    })),
  },
  totals: [['Subtotal', '1,000.00'], ['Balance due (GBP)', '1,200.00']],
  footer: 'Registered in England and Wales.',
  created: '2026-09-01T09:00:00Z',
});

test('an independent reader opens what we wrote', async (t) => {
  const doc = await open(invoice(3));
  if (!doc) return t.skip('pdfjs-dist is not installed');
  assert.strictEqual(doc.numPages, 1);
  const info = (await doc.getMetadata()).info;
  assert.strictEqual(info.Title, 'Invoice INV-2026-0041');
  assert.strictEqual(info.Producer, 'Rutba');
});

test('the words come back out, accents and typography included', async (t) => {
  const doc = await open(invoice(3));
  if (!doc) return t.skip('pdfjs-dist is not installed');
  const text = (await (await doc.getPage(1)).getTextContent()).items.map((i) => i.str).join(' ');

  assert.match(text, /INVOICE/);
  assert.match(text, /INV-2026-0041/);
  // The round trip that WinAnsi exists for: through our encoder, into
  // octal escapes, out of somebody else's decoder, still correct.
  assert.match(text, /Müller & Sons \(Northern\) Ltd/);
  assert.match(text, /Consultancy — week 1 at the “agreed” rate/);
  assert.match(text, /Amount \(GBP\)/);
});

test('a long document paginates, and every page is numbered and headed', async (t) => {
  const doc = await open(invoice(120));
  if (!doc) return t.skip('pdfjs-dist is not installed');
  assert.ok(doc.numPages > 2, `${doc.numPages} pages`);
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const text = (await page.getTextContent()).items.map((i) => i.str).join(' ');
    assert.ok(text.includes(`Page ${p} of ${doc.numPages}`), `page ${p} is not numbered`);
    assert.ok(text.includes('Description'), `page ${p} lost its column headings`);
  }
});

test('nothing is drawn outside the margins', async (t) => {
  const doc = await open(invoice(60));
  if (!doc) return t.skip('pdfjs-dist is not installed');
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const { width, height } = page.getViewport({ scale: 1 });
    for (const item of (await page.getTextContent()).items) {
      if (!item.str.trim()) continue;
      const x = item.transform[4];
      const y = item.transform[5];
      assert.ok(x >= 40, `"${item.str}" starts at ${x.toFixed(1)}, inside the left margin`);
      assert.ok(x + item.width <= width - 40 + 0.5,
        `"${item.str}" ends at ${(x + item.width).toFixed(1)} of ${width}`);
      assert.ok(y > 20 && y < height - 30, `"${item.str}" is at y ${y.toFixed(1)}`);
    }
  }
});

/**
 * The claim in metrics.js, put to an independent implementation.
 */
test('our advance widths are the reader\'s advance widths', async (t) => {
  const strings = [
    'Invoice INV-2026-0041',
    'Müller & Sons (Northern) Ltd',
    '1,234,567.89',
    'The quick brown fox jumps over the lazy dog',
    'illiWWW—“…”',
  ];
  const size = 11;
  const pdf = new PdfDocument({ created: '2026-09-01T09:00:00Z' });
  const page = pdf.addPage();
  strings.forEach((s, i) => page.text(s, 40, 60 + i * 20, { size }));
  const mine = strings.map((s) => pdf.widthOf(s, { size }));

  const doc = await open(pdf.toBuffer());
  if (!doc) return t.skip('pdfjs-dist is not installed');
  const items = (await (await doc.getPage(1)).getTextContent()).items.filter((i) => i.str.trim());
  assert.strictEqual(items.length, strings.length);

  items.forEach((item, i) => {
    assert.strictEqual(item.str, strings[i]);
    assert.ok(Math.abs(item.width - mine[i]) < 0.01,
      `"${strings[i]}": we measured ${mine[i].toFixed(4)}, the reader makes it ${item.width.toFixed(4)}`);
  });
});

test('bold is measured as bold, not as roman in a heavier colour', async (t) => {
  const size = 14;
  const pdf = new PdfDocument({ created: '2026-09-01T09:00:00Z' });
  const page = pdf.addPage();
  page.text('Balance due', 40, 60, { size, font: 'Helvetica-Bold' });
  const mine = pdf.widthOf('Balance due', { size, font: 'Helvetica-Bold' });
  assert.ok(mine > pdf.widthOf('Balance due', { size }), 'bold is the wider of the two');

  const doc = await open(pdf.toBuffer());
  if (!doc) return t.skip('pdfjs-dist is not installed');
  const item = (await (await doc.getPage(1)).getTextContent()).items[0];
  assert.ok(Math.abs(item.width - mine) < 0.01, `${mine} against ${item.width}`);
});
