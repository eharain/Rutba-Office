'use strict';

const test = require('node:test');
const assert = require('node:assert');
const { PdfDocument, SIZES, colour } = require('../src/document');

const simple = (opts = {}) => {
  const doc = new PdfDocument({ created: '2026-09-01T09:00:00Z', ...opts });
  const page = doc.addPage();
  page.text('Invoice INV-0001', 50, 60, { font: 'Helvetica-Bold', size: 14 });
  return doc;
};

test('a document with no pages is refused rather than written', () => {
  const doc = new PdfDocument();
  assert.throws(() => doc.toBuffer(), /no pages is not a document/);
});

test('the file is bracketed the way every reader expects', () => {
  const text = simple().toBuffer().toString('latin1');
  assert.ok(text.startsWith('%PDF-1.4\n'), 'header');
  assert.ok(text.includes('\n%%EOF\n'), 'trailer marker');
  assert.match(text, /\/Type \/Catalog/);
  assert.match(text, /\/Type \/Pages/);
  assert.match(text, /\/Type \/Page /);
  // Four high bytes on line two: the convention that tells a transport this
  // is not text to be helpfully re-encoded.
  assert.ok([10, 11, 12, 13].every((i) => text.charCodeAt(i) > 127), 'binary marker in the header');
});

/**
 * The structural check. Everything else about a PDF can be slightly wrong
 * and still open; a cross-reference table that points at the wrong byte
 * cannot. This walks the table the way a reader does.
 */
test('every xref offset points at the object it claims', () => {
  const buffer = simple().toBuffer();
  const text = buffer.toString('latin1');

  const startxref = Number(text.match(/startxref\n(\d+)\n%%EOF/)[1]);
  assert.strictEqual(text.slice(startxref, startxref + 4), 'xref');

  const table = text.slice(startxref);
  const size = Number(table.match(/\/Size (\d+)/)[1]);
  const entries = table.match(/^\d{10} \d{5} [nf] $/gm);
  assert.strictEqual(entries.length, size, 'one entry per object, including the free head');
  assert.ok(entries[0].endsWith('f '), 'object 0 is the free head');

  entries.slice(1).forEach((entry, i) => {
    const offset = Number(entry.slice(0, 10));
    assert.strictEqual(text.slice(offset, offset + `${i + 1} 0 obj`.length), `${i + 1} 0 obj`,
      `object ${i + 1} is not at the offset the table gives`);
  });

  // Fixed-width entries are not decoration: readers seek by multiplying.
  for (const entry of entries) assert.strictEqual(entry.length + 1, 20);
});

test('the same content twice is the same bytes, so a change is visible in a diff', () => {
  assert.ok(simple().toBuffer().equals(simple().toBuffer()));
  const other = new PdfDocument({ created: '2026-09-01T09:00:00Z' });
  other.addPage().text('Invoice INV-0002', 50, 60, { font: 'Helvetica-Bold', size: 14 });
  assert.ok(!simple().toBuffer().equals(other.toBuffer()), 'different content, different file id');
});

test('a name with a bracket in it does not corrupt the page', () => {
  const doc = new PdfDocument({ created: '2026-09-01T09:00:00Z', title: 'Acme (UK) Ltd' });
  doc.addPage().text('Acme (UK) Ltd \\ Partners', 40, 40);
  const text = doc.toBuffer().toString('latin1');
  assert.ok(text.includes('(Acme \\(UK\\) Ltd \\\\ Partners)'));
  // And the object structure is still intact after it.
  assert.match(text, /startxref\n\d+\n%%EOF/);
});

test('only the fonts actually used are declared', () => {
  const doc = new PdfDocument({ created: '2026-09-01T09:00:00Z' });
  const page = doc.addPage();
  page.text('roman', 40, 40);
  page.text('bold', 40, 60, { font: 'Helvetica-Bold' });
  const text = doc.toBuffer().toString('latin1');
  assert.strictEqual((text.match(/\/Type \/Font/g) || []).length, 2);
  assert.ok(text.includes('/BaseFont /Helvetica '), 'the roman');
  assert.ok(text.includes('/BaseFont /Helvetica-Bold'), 'the bold');
  assert.ok(text.includes('/Encoding /WinAnsiEncoding'), 'without which the accents are wrong');
});

test('y is measured down from the top, once, here', () => {
  const doc = new PdfDocument({ created: '2026-09-01T09:00:00Z' });
  const page = doc.addPage();
  page.text('top of the page', 40, 50);
  const text = doc.toBuffer().toString('latin1');
  // A4 is 841.89 tall, so 50 from the top is 791.89 in PDF's own frame.
  assert.ok(text.includes('1 0 0 1 40 791.89 Tm'), text.slice(text.indexOf('BT'), text.indexOf('ET')));
});

test('right and centre alignment use the real width of the text', () => {
  const doc = new PdfDocument({ created: '2026-09-01T09:00:00Z' });
  const page = doc.addPage();
  const width = doc.widthOf('TOTAL', { font: 'Helvetica-Bold', size: 12 });
  page.textRight('TOTAL', 500, 100, { font: 'Helvetica-Bold', size: 12 });
  const text = doc.toBuffer().toString('latin1');
  const x = Number(text.match(/1 0 0 1 ([\d.]+) [\d.]+ Tm/)[1]);
  assert.ok(Math.abs(x - (500 - width)) < 0.01, `${x} should be 500 minus ${width}`);
});

test('paper sizes are points, and landscape is the same paper turned', () => {
  assert.deepStrictEqual(SIZES.A4, [595.28, 841.89]);
  const portrait = new PdfDocument({ size: 'A4' });
  const landscape = new PdfDocument({ size: 'A4', landscape: true });
  assert.deepStrictEqual(landscape.size, [portrait.size[1], portrait.size[0]]);
  assert.throws(() => new PdfDocument({ size: 'Foolscap' }), /not a paper size/);
});

test('colours come in as hex or components and leave as PDF numbers', () => {
  assert.deepStrictEqual(colour('#ffffff'), ['1', '1', '1']);
  assert.deepStrictEqual(colour('#000'), ['0', '0', '0']);
  assert.deepStrictEqual(colour([0.5, 0, 1]), ['0.5', '0', '1']);
  assert.throws(() => colour('rebeccapurple'), /not a colour/);
});

test('a coordinate that is not a number is refused, not drawn at NaN', () => {
  const page = new PdfDocument().addPage();
  assert.throws(() => page.text('x', undefined, 10), /not a usable coordinate/);
  assert.throws(() => page.line(0, 0, NaN, 10), /not a usable coordinate/);
});
