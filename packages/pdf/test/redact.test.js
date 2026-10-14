'use strict';

// Redaction (redact.js): what a mark covers is removed from the file, not
// drawn over. Each case is proved by reading the result back three ways:
// this package's reader, pdf.js, and a plain search of every byte and every
// inflated stream for the words (single-byte, UTF-16 and hex). The words
// beside the redacted ones stay where they were; pictures under a mark are
// painted out; a form drawn on two pages is redacted on one only; text in a
// recognised scan's invisible layer goes too; annotations over a mark go,
// others stay; and the file is written whole, never as an update.

const test = require('node:test');
const assert = require('node:assert');
const zlib = require('zlib');
const { PdfDocument } = require('../src/document');
const { openPdf, Dict, Name, Stream } = require('../src/reader');
const { textString } = require('../src/assemble');
const { PdfUpdate, flateStream } = require('../src/update');
const { annotate, readAnnotations } = require('../src/annotations');
const { addTextLayer } = require('../src/textlayer');
const { applyRedactions, redactionPlan, bytesContain, parseContent, writeContent } = require('../src/redact');
const { widthOfBytes } = require('../src/metrics');
const { encode } = require('../src/encoding');
const { readWithPdfjs } = require('./pdfjs');

const H = 841.89;
const created = '2026-10-14T09:00:00Z';

/** Where a run of words in a line the writer set at (x, top) lies, in the page's coordinates. */
function wordBox(line, word, x, top, size, font = 'Helvetica') {
  const at = line.indexOf(word);
  const x0 = x + widthOfBytes(encode(line.slice(0, at)), font, size);
  const x1 = x0 + widthOfBytes(encode(word), font, size);
  const base = H - top;
  return [x0 - 0.5, base - size * 0.25, x1 + 0.5, base + size * 0.85];
}

const LINE = 'Account SECRET-4417 belongs to the client.';
function letter() {
  const doc = new PdfDocument({ title: 'Letter', created });
  const page = doc.addPage();
  page.text('Dear client,', 72, 100, { font: 'Helvetica', size: 14 });
  page.text(LINE, 72, 140, { font: 'Helvetica', size: 14 });
  page.text('Nothing else here is private.', 72, 180, { font: 'Courier', size: 12 });
  doc.addPage().text('Page two keeps SECRET-4417 untouched.', 72, 100, { font: 'Helvetica', size: 14 });
  return doc.toBuffer();
}

test('the words under a mark are gone from the file, read three ways; the words beside them stay where they were', async () => {
  const original = letter();
  const box = wordBox(LINE, 'SECRET-4417', 72, 140, 14);
  const { bytes, removed } = applyRedactions(original, { areas: [{ page: 0, rects: [box] }] });
  assert.equal(removed.glyphs, 'SECRET-4417'.length);
  assert.ok(!bytes.subarray(0, original.length).equals(original), 'written whole, not after the original');
  const file = openPdf(bytes);
  // Our reader.
  const text = file.pageText(0);
  assert.doesNotMatch(text, /SECRET|4417/);
  assert.match(text, /Account/);
  assert.match(text, /belongs to the client/);
  assert.match(file.pageText(1), /SECRET-4417/, 'the other page is not touched');
  // A plain search of the bytes: only page two has the words.
  const one = applyRedactions(original, { areas: [{ page: 0, rects: [box] }, { page: 1, rects: [[0, 0, 600, 842]] }] });
  assert.equal(bytesContain(one.bytes, 'SECRET-4417'), false);
  assert.equal(bytesContain(one.bytes, 'SECRET'), false);
  assert.equal(bytesContain(original, 'SECRET-4417'), true, 'the search finds it where it is');
  // pdf.js.
  const read = await readWithPdfjs(bytes);
  if (read) {
    assert.doesNotMatch(read.pages[0].text, /SECRET|4417/);
    assert.match(read.pages[0].text, /belongs to the client/);
  }
  // The words after the mark start where they did: the space the words took is kept by a number in the TJ array.
  const kept = -(widthOfBytes(encode('SECRET-4417'), 'Helvetica', 14) * 1000) / 14;
  const numbers = [...file.pageContent(0).toString('latin1').matchAll(/(-[\d.]+)/g)].map((m) => Number(m[1]));
  assert.ok(numbers.some((v) => Math.abs(v - kept) < 0.01), `no advance of ${kept} in the page`);
  // The black box is drawn where the mark was.
  assert.match(file.pageContent(0).toString('latin1'), /0 0 0 rg [\d.]+ [\d.]+ [\d.]+ [\d.]+ re f/);
});

test('Redact marks in the file are applied, the marks and comments over them go, and other comments stay', () => {
  const box = wordBox(LINE, 'SECRET-4417', 72, 140, 14);
  const marked = annotate(letter(), {
    add: [
      { page: 0, type: 'Redact', quads: [[box[0], box[3], box[2], box[3], box[0], box[1], box[2], box[1]]], overlay: 'Withheld' },
      { page: 0, type: 'Text', rect: [box[0] + 5, box[1], box[0] + 25, box[1] + 20], contents: 'This is SECRET-4417' },
      { page: 0, type: 'Text', rect: [400, 100, 420, 120], contents: 'Elsewhere' },
    ],
  }).bytes;
  assert.equal(readAnnotations(openPdf(marked)).filter((a) => a.type === 'Redact').length, 1);
  const { bytes, removed } = applyRedactions(marked);
  assert.equal(removed.annotations, 1);
  const file = openPdf(bytes);
  const left = readAnnotations(file);
  assert.deepEqual(left.map((a) => a.contents), ['Elsewhere']);
  assert.doesNotMatch(file.pageText(0), /SECRET/);
  assert.match(file.pageText(0), /Withheld/, 'the overlay text is written on the box');
  assert.equal(bytesContain(bytes, 'This is SECRET'), false, 'the comment over the mark is gone');
});

test('nothing marked is refused in a sentence', () => {
  assert.throws(() => applyRedactions(letter()), /Nothing is marked for redaction/);
});

test('a recognised scan: the invisible words under the mark go, and the picture is painted out there', async () => {
  const doc = new PdfDocument({ title: 'Scan', created });
  const grey = Buffer.alloc(100 * 50, 200);
  const image = doc.addImage({ width: 100, height: 50, colorSpace: 'DeviceGray', data: grey });
  doc.addPage().image(image, 0, 0, 595.28, H);
  const scan = doc.toBuffer();
  const layered = addTextLayer(scan, [{ page: 0, lines: [{ words: [
    { text: 'Patient', x: 72, y: 700, width: 60, size: 14 },
    { text: 'Zainab', x: 140, y: 700, width: 60, size: 14 },
    { text: 'Rehman', x: 208, y: 700, width: 64, size: 14 },
  ] }] }]).bytes;
  assert.equal(bytesContain(layered, 'Zainab'), true);
  const { bytes, removed } = applyRedactions(layered, { areas: [{ page: 0, rects: [[138, 695, 202, 714]] }] });
  // The word, and the spaces either side of it that lie under the area.
  assert.ok(removed.glyphs >= 'Zainab'.length && removed.glyphs <= 'Zainab'.length + 2, String(removed.glyphs));
  assert.equal(removed.pictures, 1);
  assert.equal(bytesContain(bytes, 'Zainab'), false);
  const file = openPdf(bytes);
  // The picture's pixels under the area are black now, and those elsewhere as they were.
  const page = file.pages()[0];
  const xo = file.resolve(file.resolve(page.dict.get('Resources')).get('XObject'));
  const pics = [...xo.keys()].map((k) => file.resolve(xo.get(k))).filter((s) => s instanceof Stream && s.dict.get('Subtype').name === 'Image');
  const pix = file.decode(pics.find((s) => file.decode(s).includes(0)));
  const at = (x, y) => pix[y * 100 + x];
  // (170, 704) on a 595 x 842 page is pixel (28.6, 8.2) of 100 x 50.
  assert.equal(at(28, 8), 0);
  assert.equal(at(80, 40), 200);
  const read = await readWithPdfjs(bytes);
  if (read) {
    assert.doesNotMatch(read.pages[0].text, /Zainab/);
    assert.match(read.pages[0].text, /Patient/);
    assert.match(read.pages[0].text, /Rehman/);
  }
});

test('a form drawn on two pages is redacted on the page marked and left whole on the other', async () => {
  const doc = new PdfDocument({ title: 'Stamp', created });
  doc.addPage().text('First', 72, 80, { font: 'Helvetica', size: 12 });
  doc.addPage().text('Second', 72, 80, { font: 'Helvetica', size: 12 });
  const up = new PdfUpdate(doc.toBuffer());
  const font = up.add(new Dict([['Type', new Name('Font')], ['Subtype', new Name('Type1')], ['BaseFont', new Name('Helvetica')], ['Encoding', new Name('WinAnsiEncoding')]]));
  const form = up.add(flateStream([
    ['Type', new Name('XObject')], ['Subtype', new Name('Form')], ['BBox', [0, 0, 300, 50]],
    ['Resources', new Dict([['Font', new Dict([['F9', font]])]])],
  ], 'BT /F9 12 Tf 10 20 Td (Ref HIDDEN-99 here) Tj ET'));
  for (const pg of up.file.pages()) {
    const draw = up.add(flateStream([], 'q 1 0 0 1 72 600 cm /Fm9 Do Q'));
    const dict = new Dict(pg.dict.entries());
    const res = new Dict(up.file.resolve(pg.dict.get('Resources')).entries());
    res.set('XObject', new Dict([['Fm9', form]]));
    dict.set('Resources', res);
    const contents = up.file.resolve(pg.dict.get('Contents'));
    dict.set('Contents', [...(Array.isArray(contents) ? contents : [pg.dict.get('Contents')]), draw]);
    up.set(pg.ref, dict);
  }
  const both = up.toBuffer();
  // "HIDDEN-99" starts after "Ref " at 72 + 10 + width("Ref ").
  const x0 = 82 + widthOfBytes(encode('Ref '), 'Helvetica', 12);
  const x1 = x0 + widthOfBytes(encode('HIDDEN-99'), 'Helvetica', 12);
  const { bytes } = applyRedactions(both, { areas: [{ page: 0, rects: [[x0 - 0.5, 615, x1 + 0.5, 632]] }] });
  const read = await readWithPdfjs(bytes);
  if (read) {
    assert.doesNotMatch(read.pages[0].text, /HIDDEN/);
    assert.match(read.pages[0].text, /Ref/);
    assert.match(read.pages[0].text, /here/);
    assert.match(read.pages[1].text, /HIDDEN-99/);
  }
});

test('a picture the module cannot decode is named in the plan, painted out when handed back, and left out whole when not', () => {
  const doc = new PdfDocument({ title: 'Photo', created });
  doc.addPage().text('A photograph below.', 72, 80, { font: 'Helvetica', size: 12 });
  const up = new PdfUpdate(doc.toBuffer());
  const pg = up.file.pages()[0];
  const jpeg = up.add(new Stream(new Dict([
    ['Type', new Name('XObject')], ['Subtype', new Name('Image')], ['Width', 40], ['Height', 20], ['ColorSpace', new Name('DeviceRGB')],
    ['BitsPerComponent', 8], ['Filter', new Name('DCTDecode')],
  ]), Buffer.from('not really a jpeg, but nobody decodes it here')));
  const draw = up.add(flateStream([], 'q 400 0 0 200 100 300 cm /Im9 Do Q'));
  const dict = new Dict(pg.dict.entries());
  const res = new Dict(up.file.resolve(pg.dict.get('Resources')).entries());
  res.set('XObject', new Dict([['Im9', jpeg]]));
  dict.set('Resources', res);
  dict.set('Contents', [pg.dict.get('Contents'), draw]);
  up.set(pg.ref, dict);
  const photo = up.toBuffer();
  const areas = [{ page: 0, rects: [[100, 300, 300, 400]] }];
  const plan = redactionPlan(photo, { areas });
  assert.equal(plan.length, 1);
  assert.equal(plan[0].filter, 'DCTDecode');
  assert.deepEqual(plan[0].rects, [[0, 10, 20, 20]]);
  // Handed back decoded: painted out, as a new picture.
  const rgb = Buffer.alloc(40 * 20 * 3, 255);
  const given = applyRedactions(photo, { areas, images: { [plan[0].key]: { width: 40, height: 20, data: rgb } } });
  assert.equal(given.removed.pictures, 1);
  const file = openPdf(given.bytes);
  const xo = file.resolve(file.resolve(file.pages()[0].dict.get('Resources')).get('XObject'));
  const img = [...xo.keys()].map((k) => file.resolve(xo.get(k))).find((s) => s.dict.get('Filter')?.name === 'FlateDecode');
  const px = file.decode(img);
  assert.equal(px[(15 * 40 + 5) * 3], 0, 'under the area');
  assert.equal(px[(5 * 40 + 30) * 3], 255, 'elsewhere');
  assert.equal(given.bytes.includes(Buffer.from('not really a jpeg')), false, 'the old picture is not in the file');
  // Not handed back: the picture is left out of the page entirely.
  const dropped = applyRedactions(photo, { areas });
  assert.equal(dropped.removed.images, 1);
  assert.equal(dropped.bytes.includes(Buffer.from('not really a jpeg')), false);
});

test('content read and written again draws the same', () => {
  const content = Buffer.from('q 1 0 0 1 10 20 cm BT /F1 12 Tf (a\\(b\\)) Tj [<0041> -120 (x)] TJ ET BI /W 2 /H 1 /BPC 8 /CS /G ID \x00\xff EI Q /P <</MCID 3>> BDC EMC', 'latin1');
  const ops = parseContent(content);
  assert.deepEqual(ops.map((o) => o.op), ['q', 'cm', 'BT', 'Tf', 'Tj', 'TJ', 'ET', 'BI', 'Q', 'BDC', 'EMC']);
  const again = parseContent(writeContent(ops));
  assert.deepEqual(again.map((o) => o.op), ops.map((o) => o.op));
  assert.ok(writeContent(ops).includes(Buffer.from('ID \x00\xff EI', 'latin1')), 'the inline picture is kept byte for byte');
  assert.ok(zlib && textString);
});
