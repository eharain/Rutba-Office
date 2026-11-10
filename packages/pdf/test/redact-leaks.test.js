'use strict';

// Where words can hide on a page besides the page's own text, each found by a
// peer review on 10 Nov 2026 to survive a redaction: a picture's masks, a
// pattern, a soft-mask group, a reading-aloud text kept in the page's
// property lists, text whose font is set by a graphics state, and a composite
// font whose codes are not two bytes each.

const test = require('node:test');
const assert = require('node:assert');
const zlib = require('zlib');
const { applyRedactions, bytesContain } = require('../src/redact');
const { openPdf } = require('../src/reader');

const SECRET = 'SECRET-4417';
const FONT = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>';

/** A PDF of one page from the objects given, numbered from 4. */
function pdf({ content, resources = '', objects = [] }) {
  const streamOf = (data, dict = '') => `<< ${dict} /Length ${Buffer.byteLength(data, 'latin1')} >>\nstream\n${data}\nendstream`;
  const all = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> ${resources} >> >>`,
    streamOf(content),
    FONT,
    ...objects.map((o) => (typeof o === 'string' ? o : streamOf(o.data, o.dict))),
  ];
  let out = '%PDF-1.7\n';
  const offsets = [];
  all.forEach((body, i) => {
    offsets.push(Buffer.byteLength(out, 'latin1'));
    out += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const at = Buffer.byteLength(out, 'latin1');
  out += `xref\n0 ${all.length + 1}\n0000000000 65535 f \n${offsets.map((o) => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}`;
  out += `trailer\n<< /Size ${all.length + 1} /Root 1 0 R >>\nstartxref\n${at}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

const AREA = [{ page: 0, rects: [[100, 100, 300, 300]] }];
const plainText = (bytes) => openPdf(bytes).pages().length;

function picture(extra = '') {
  const data = Buffer.alloc(40 * 40, 200).toString('latin1');
  return { data, dict: `/Type /XObject /Subtype /Image /Width 40 /Height 40 /BitsPerComponent 8 /ColorSpace /DeviceGray ${extra}` };
}
const maskHolding = (word) => {
  const bytes = Buffer.alloc(40 * 40, 255);
  bytes.write(word, 300, 'latin1');
  return bytes.toString('latin1');
};

test('a picture under a mark: words in its soft mask do not survive, and the mask is still there', () => {
  const file = pdf({
    content: 'q 200 0 0 200 100 100 cm /Im0 Do Q',
    resources: '/XObject << /Im0 6 0 R >>',
    objects: [picture('/SMask 7 0 R'), { data: maskHolding(SECRET), dict: '/Type /XObject /Subtype /Image /Width 40 /Height 40 /BitsPerComponent 8 /ColorSpace /DeviceGray' }],
  });
  assert.ok(bytesContain(file, SECRET), 'the test file holds it to begin with');
  const out = applyRedactions(file, { areas: AREA });
  assert.equal(bytesContain(out.bytes, SECRET), false);
  assert.equal(out.removed.pictures, 1);
  assert.match(out.bytes.toString('latin1'), /\/SMask/, 'the painted mask is written');
});

test('a picture under a mark: words in its stencil mask do not survive', () => {
  const file = pdf({
    content: 'q 200 0 0 200 100 100 cm /Im0 Do Q',
    resources: '/XObject << /Im0 6 0 R >>',
    objects: [picture('/Mask 7 0 R'), { data: `${SECRET}${'x'.repeat(200)}`, dict: '/Type /XObject /Subtype /Image /Width 40 /Height 40 /BitsPerComponent 1 /ImageMask true' }],
  });
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('a pattern that draws words is emptied, and the person is told something was left off', () => {
  const file = pdf({
    content: '/Pattern cs /P0 scn 100 100 200 200 re f',
    resources: '/Pattern << /P0 6 0 R >>',
    objects: [{ data: `BT /F1 12 Tf 10 10 Td (${SECRET}) Tj ET`, dict: '/Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 100 100] /XStep 100 /YStep 100 /Resources << /Font << /F1 5 0 R >> >>' }],
  });
  assert.ok(bytesContain(file, SECRET));
  const out = applyRedactions(file, { areas: AREA });
  assert.equal(bytesContain(out.bytes, SECRET), false);
  assert.ok(out.unpainted.length, 'said');
  assert.equal(plainText(out.bytes), 1);
});

test('a soft-mask group that draws words does not survive', () => {
  const file = pdf({
    content: '/G0 gs 0 0 1 rg 100 100 200 200 re f',
    resources: '/ExtGState << /G0 << /SMask << /Type /Mask /S /Luminosity /G 6 0 R >> >> >>',
    objects: [{ data: `BT /F1 12 Tf 10 10 Td (${SECRET}) Tj ET`, dict: '/Type /XObject /Subtype /Form /BBox [0 0 612 792] /Group << /S /Transparency /CS /DeviceGray >> /Resources << /Font << /F1 5 0 R >> >>' }],
  });
  assert.ok(bytesContain(file, SECRET));
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('a reading-aloud text kept in the page\'s property lists goes with the words it spoke for', () => {
  const file = pdf({
    content: '/Span /MC0 BDC BT /F1 14 Tf 110 200 Td (hidden words) Tj ET EMC BT /F1 14 Tf 400 600 Td (kept words) Tj ET',
    resources: `/Properties << /MC0 << /ActualText (${SECRET}) >> >>`,
  });
  assert.ok(bytesContain(file, SECRET));
  const out = applyRedactions(file, { areas: AREA });
  assert.equal(bytesContain(out.bytes, SECRET), false);
  assert.ok(out.removed.glyphs > 0);
});

test('a property list on a span the mark did not reach is left alone', () => {
  const file = pdf({
    content: '/Span /MC0 BDC BT /F1 14 Tf 400 600 Td (other words) Tj ET EMC BT /F1 14 Tf 110 200 Td (hidden) Tj ET',
    resources: '/Properties << /MC0 << /ActualText (stays here) >> >>',
  });
  assert.ok(bytesContain(applyRedactions(file, { areas: AREA }).bytes, 'stays here'));
});

test('text whose font comes from a graphics state is measured by that font', () => {
  // Narrow letters, so a default width would run far past the mark and take the word after it too.
  const narrow = 'i'.repeat(30);
  const line = `(${narrow}) Tj (${SECRET}) Tj`;
  const file = pdf({ content: `/G0 gs BT 10 200 Td ${line} (after) Tj ET`, resources: '/ExtGState << /G0 << /Font [5 0 R 12] >> >>' });
  // The narrow letters are 30 x 2.66 wide, so the secret starts near 90 and ends near 169.
  const out = applyRedactions(file, { areas: [{ page: 0, rects: [[88, 190, 168, 215]] }] });
  assert.equal(bytesContain(out.bytes, SECRET), false, 'the word under the mark goes');
  assert.ok(bytesContain(out.bytes, 'after'), 'the word past the mark stays');
});

test('a composite font with a non-Identity encoding: a string near the mark goes whole', () => {
  const type0 = '<< /Type /Font /Subtype /Type0 /BaseFont /MSGothic /Encoding /90ms-RKSJ-H /DescendantFonts [7 0 R] >>';
  const cid = '<< /Type /Font /Subtype /CIDFontType2 /BaseFont /MSGothic /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 2 >> /DW 500 >>';
  const file = pdf({ content: `BT /F2 12 Tf 110 200 Td (${SECRET}) Tj ET BT /F2 12 Tf 400 600 Td (farther off) Tj ET`, objects: [type0, cid] });
  // The page's own Font resource is replaced with one that names the composite font.
  const patched = Buffer.from(file.toString('latin1').replace('/Font << /F1 5 0 R >>', '/Font << /F2 6 0 R >>'), 'latin1');
  const fixed = rebuild(patched);
  const out = applyRedactions(fixed, { areas: AREA });
  assert.equal(bytesContain(out.bytes, SECRET), false);
  assert.ok(bytesContain(out.bytes, 'farther off'));
});

/** The cross-reference table again after a file was edited as text. */
function rebuild(bytes) {
  let text = bytes.toString('latin1');
  text = text.slice(0, text.indexOf('xref\n'));
  const offsets = [];
  for (const m of text.matchAll(/^(\d+) 0 obj$/gm)) offsets[Number(m[1]) - 1] = m.index;
  const at = text.length;
  text += `xref\n0 ${offsets.length + 1}\n0000000000 65535 f \n${offsets.map((o) => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}`;
  text += `trailer\n<< /Size ${offsets.length + 1} /Root 1 0 R >>\nstartxref\n${at}\n%%EOF\n`;
  return Buffer.from(text, 'latin1');
}

test('the file written can still be read', () => {
  const file = pdf({ content: 'BT /F1 14 Tf 110 200 Td (hidden) Tj ET' });
  assert.equal(plainText(applyRedactions(file, { areas: AREA }).bytes), 1);
  assert.ok(zlib);
});
