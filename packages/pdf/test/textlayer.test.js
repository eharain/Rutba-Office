'use strict';

// The invisible text layer on scanned pages (textlayer.js): written as an
// incremental update after the original bytes, on files with classic
// cross-reference tables and with cross-reference and object streams, on
// pages that inherit their resources and rotation, with right-to-left
// words drawn in the order a page shows them.

const test = require('node:test');
const assert = require('node:assert');
const zlib = require('zlib');
const { PdfDocument } = require('../src/document');
const { openPdf, Dict, Ref } = require('../src/reader');
const { readTrueType } = require('../src/truetype');
const { visualText } = require('../src/bidi');
const { addTextLayer, pageHasText, glyphlessFont, TextLayerError } = require('../src/textlayer');

/** A page that is one grey picture and no text, as a scanner makes. */
function scannedPdf(pages = 1) {
  const doc = new PdfDocument();
  const data = Buffer.alloc(40 * 20, 200);
  const image = doc.addImage({ width: 40, height: 20, colorSpace: 'DeviceGray', data });
  for (let i = 0; i < pages; i++) doc.addPage().image(image, 0, 0, 595, 842);
  return doc.toBuffer();
}

/** The layer's words on a page: each `<hex> Tj` of its content, as UTF-16 code units. */
function layerText(file, index) {
  const content = file.pageContent(index).toString('latin1');
  const out = [];
  for (const m of content.matchAll(/<([0-9a-f]+)>\s*Tj/g)) {
    let s = '';
    for (let i = 0; i < m[1].length; i += 4) s += String.fromCharCode(parseInt(m[1].slice(i, i + 4), 16));
    out.push(s);
  }
  return out;
}

const word = (text, x, y = 700, width = 50, size = 12) => ({ text, x, y, width, size });

test('a layer is appended after the original bytes, the page redefined with its content kept and one stream more', () => {
  const original = scannedPdf(2);
  const before = openPdf(original);
  const contentBefore = before.pageContent(0);
  assert.equal(pageHasText(before, 0), false);
  const { bytes, pages } = addTextLayer(original, [{ page: 0, lines: [{ words: [word('Invoice', 72), word('total', 130)] }, { words: [word('59.20', 72, 680)] }] }]);
  assert.deepEqual(pages, [0]);
  assert.ok(bytes.subarray(0, original.length).equals(original), 'the original bytes are the start of the file');
  assert.ok(bytes.length > original.length);
  const after = openPdf(bytes);
  assert.equal(after.repaired, false, 'read through its cross-references, not by scanning');
  assert.equal(after.pageCount, 2);
  assert.deepEqual(layerText(after, 0), ['Invoice ', 'total', '59.20']);
  assert.ok(after.pageContent(0).includes(contentBefore), 'the page draws exactly what it drew');
  assert.match(after.pageContent(0).toString('latin1'), /3 Tr/);
  assert.equal(pageHasText(after, 0), true);
  assert.equal(pageHasText(after, 1), false, 'a page not named is left alone');
  const page = after.pages()[0].dict;
  const fonts = after.resolve(after.resolve(page.get('Resources')).get('Font'));
  const font = after.resolve(fonts.get('ROCR'));
  assert.equal(font.get('Subtype').name, 'Type0');
  assert.equal(font.get('Encoding').name, 'Identity-H');
  assert.ok(after.resolve(font.get('ToUnicode')));
  assert.ok(after.resolve(page.get('Resources')).get('XObject'), 'the page keeps its picture');
  // The update chains to the original's cross-references.
  assert.equal(after.trailer.get('Prev'), before.startXref());
});

test('each word is stretched to its width and placed at its baseline; a word on a turned baseline turns with it', () => {
  const { bytes } = addTextLayer(scannedPdf(), [{ page: 0, lines: [{ words: [{ text: 'abcd', x: 100, y: 500, width: 40, size: 10 }, { text: 'up', x: 300, y: 200, width: 20, size: 10, angle: 90 }] }] }]);
  const content = openPdf(bytes).pageContent(0).toString('latin1');
  // Four units at half an em: scale 40 / 2 = 20.
  assert.match(content, /\/ROCR 1 Tf\n20 0 0 10 100 500 Tm\n/);
  // The next word a size step on (2), its matrix halved: the same stretch and place.
  assert.match(content, /\/ROCR 2 Tf\n0 10 -5 0 300 200 Tm\n/);
});

test('every word is a size step from the one before it, so a reader that runs a line together still keeps each word over its place', () => {
  const words = ['Rutba', 'Office', 'reads', 'the', 'words'].map((text, i) => ({ text, x: 72 + i * 60, y: 700, width: 10 * text.length, size: 12 }));
  const { bytes } = addTextLayer(scannedPdf(), [{ page: 0, lines: [{ words }] }]);
  const content = openPdf(bytes).pageContent(0).toString('latin1');
  const sizes = [...content.matchAll(/\/ROCR (\d+) Tf/g)].map((m) => Number(m[1]));
  assert.deepEqual(sizes, [1, 2, 1, 2, 1]);
  // Size times matrix is each word's own stretch (its width over its units at half an em) and height.
  const steps = [...content.matchAll(/\/ROCR (\d+) Tf\n([-\d.]+) [-\d.]+ [-\d.]+ ([-\d.]+) ([-\d.]+) ([-\d.]+) Tm/g)];
  assert.equal(steps.length, words.length);
  steps.forEach((m, i) => {
    assert.equal(Number(m[1]) * Number(m[2]), words[i].width / (words[i].text.length * 0.5));
    assert.equal(Number(m[1]) * Number(m[3]), 12);
    assert.equal(Number(m[4]), words[i].x);
  });
});

test('right-to-left words are written in the order a page shows them; Chinese has no space between characters', () => {
  const urdu = 'رتبہ';
  const { bytes } = addTextLayer(scannedPdf(), [{ page: 0, lines: [
    { words: [word(urdu, 300), word('آفس', 240)] },
    { words: [word('文件', 72, 600), word('请求', 120, 600), word('Office', 200, 600)] },
  ] }]);
  const got = layerText(openPdf(bytes), 0);
  assert.equal(got[0], `${visualText(urdu)} `);
  assert.equal([...got[0].trim()].reverse().join(''), urdu);
  assert.deepEqual(got.slice(2), ['文件', '请求 ', 'Office']);
});

test('the font is a real TrueType program of empty glyphs half an em wide', () => {
  const font = readTrueType(glyphlessFont());
  assert.equal(font.unitsPerEm, 1000);
  assert.equal(font.advance(1), 500);
  assert.equal(font.advance(0), 500);
});

/** A PDF 1.5 file with a cross-reference stream and its page inside an object stream; resources and rotation inherited. */
function streamedPdf() {
  const parts = ['%PDF-1.5\n%\xe2\xe3\xcf\xd3\n'];
  const offsets = {};
  let length = Buffer.byteLength(parts[0], 'latin1');
  const put = (n, body) => {
    offsets[n] = length;
    const text = `${n} 0 obj\n${body}\nendobj\n`;
    parts.push(text);
    length += Buffer.byteLength(text, 'latin1');
  };
  put(1, '<< /Type /Catalog /Pages 2 0 R >>');
  put(2, '<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 400 300] /Rotate 90 /Resources << /Font << /F1 6 0 R >> >> >>');
  const content = 'BT /F1 12 Tf 10 10 Td ET';
  put(4, `<< /Length ${content.length} >>\nstream\n${content}\nendstream`);
  put(6, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
  // Object 3 (the page) lives in object stream 5.
  const pageBody = '<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>';
  const head = '3 0 ';
  const objStm = zlib.deflateSync(Buffer.from(head + pageBody, 'latin1'));
  offsets[5] = length;
  const stmText = Buffer.concat([Buffer.from(`5 0 obj\n<< /Type /ObjStm /N 1 /First ${head.length} /Filter /FlateDecode /Length ${objStm.length} >>\nstream\n`, 'latin1'), objStm, Buffer.from('\nendstream\nendobj\n', 'latin1')]);
  parts.push(stmText);
  length += stmText.length;
  const xrefAt = length;
  const row = (type, a, b) => { const r = Buffer.alloc(7); r[0] = type; r.writeUInt32BE(a, 1); r.writeUInt16BE(b, 5); return r; };
  const rows = [row(0, 0, 65535), row(1, offsets[1], 0), row(1, offsets[2], 0), row(2, 5, 0), row(1, offsets[4], 0), row(1, offsets[5], 0), row(1, offsets[6], 0), row(1, xrefAt, 0)];
  const data = Buffer.concat(rows);
  const xref = Buffer.concat([Buffer.from(`7 0 obj\n<< /Type /XRef /Size 8 /W [1 4 2] /Root 1 0 R /Length ${data.length} >>\nstream\n`, 'latin1'), data, Buffer.from(`\nendstream\nendobj\nstartxref\n${xrefAt}\n%%EOF\n`, 'latin1')]);
  return Buffer.concat([...parts.slice(0, 1).map((p) => Buffer.from(p, 'latin1')), ...parts.slice(1, -1).map((p) => Buffer.from(p, 'latin1')), stmText, xref]);
}

test('a file with a cross-reference stream gets one in its update; a page from an object stream, inheriting resources and a rotation, is redefined with them', () => {
  const original = streamedPdf();
  const before = openPdf(original);
  assert.equal(before.repaired, false);
  assert.equal(before.pageInfo(0).rotate, 90);
  const { bytes } = addTextLayer(original, [{ page: 0, lines: [{ words: [word('Scan', 20, 30)] }] }]);
  assert.ok(bytes.subarray(0, original.length).equals(original));
  const update = bytes.subarray(original.length).toString('latin1');
  assert.match(update, /\/Type \/XRef/);
  assert.doesNotMatch(update, /\ntrailer\n/);
  const after = openPdf(bytes);
  assert.equal(after.repaired, false);
  assert.deepEqual(layerText(after, 0), ['Scan']);
  assert.equal(after.pageInfo(0).rotate, 90, 'the rotation it inherited is kept');
  assert.deepEqual(after.pageInfo(0).mediaBox, [0, 0, 400, 300]);
  const resources = after.resolve(after.pages()[0].dict.get('Resources'));
  const fonts = after.resolve(resources.get('Font'));
  assert.ok(fonts.get('F1') instanceof Ref && fonts.get('ROCR') instanceof Ref, 'its own font kept beside the layer\'s');
  assert.match(after.pageContent(0).toString('latin1'), /BT \/F1 12 Tf 10 10 Td ET/);
});

test('a second layer chains to the first update; an encrypted file and a page that does not exist are refused by name; no words, no update', () => {
  const once = addTextLayer(scannedPdf(2), [{ page: 0, lines: [{ words: [word('one', 72)] }] }]).bytes;
  const twice = addTextLayer(once, [{ page: 1, lines: [{ words: [word('two', 72)] }] }]).bytes;
  const file = openPdf(twice);
  assert.deepEqual([layerText(file, 0), layerText(file, 1)], [['one'], ['two']]);
  assert.ok(twice.subarray(0, once.length).equals(once));

  const original = scannedPdf();
  const encrypted = Buffer.from(original.toString('latin1').replace(/trailer\n<</, 'trailer\n<</Encrypt << /Filter /Standard >> '), 'latin1');
  assert.throws(() => addTextLayer(encrypted, [{ page: 0, lines: [{ words: [word('x', 1)] }] }]), (e) => e instanceof TextLayerError && e.code === 'encrypted');
  assert.throws(() => addTextLayer(original, [{ page: 4, lines: [{ words: [word('x', 1)] }] }]), (e) => e.code === 'page');
  const none = addTextLayer(original, [{ page: 0, lines: [] }]);
  assert.equal(none.bytes, original);
  assert.deepEqual(none.pages, []);
  assert.ok(new Dict() instanceof Dict);
});
