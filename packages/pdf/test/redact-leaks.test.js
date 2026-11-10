'use strict';

// Where words can hide on a page besides the page's own text, each found by a
// peer review on 10 Nov 2026 to survive a redaction: a picture's masks, a
// pattern, a soft-mask group, a reading-aloud text kept in the page's
// property lists, text whose font is set by a graphics state, and a composite
// font whose codes are not two bytes each.

const test = require('node:test');
const assert = require('node:assert');
const { applyRedactions, bytesContain, pageGlyphs } = require('../src/redact');
const { openPdf, Ref } = require('../src/reader');

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

// Words about a picture, kept beside it, can name what it showed: each goes with the picture painted out.
for (const [what, extra, objects] of [
  ['its description', '/Metadata 7 0 R', [{ data: `<x:xmpmeta><dc:description>${SECRET}</dc:description></x:xmpmeta>`, dict: '/Type /Metadata /Subtype /XML' }]],
  ['its prepress link', `/OPI << /2.0 << /Type /OPI /Version 2.0 /F (${SECRET}.tif) >> >>`, []],
  ['its alternative text', `/Alt (${SECRET})`, []],
]) {
  test(`a picture under a mark does not keep ${what}`, () => {
    const file = pdf({ content: 'q 200 0 0 200 100 100 cm /Im0 Do Q', resources: '/XObject << /Im0 6 0 R >>', objects: [picture(extra), ...objects] });
    assert.ok(bytesContain(file, SECRET), 'the test file holds it to begin with');
    const out = applyRedactions(file, { areas: AREA });
    assert.equal(out.removed.pictures, 1);
    assert.equal(bytesContain(out.bytes, SECRET), false);
  });
}

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
  // It starts left of the mark at 40 and reaches it only as eleven glyphs of 6 points, one a byte, not as six of two bytes.
  const file = pdf({ content: `BT /F2 12 Tf 40 200 Td (${SECRET}) Tj ET BT /F2 12 Tf 400 600 Td (farther off) Tj ET`, objects: [type0, cid] });
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

test('the file written can still be read, and the words outside the mark are still in it', () => {
  const file = pdf({ content: 'BT /F1 14 Tf 110 200 Td (hidden) Tj ET BT /F1 14 Tf 400 600 Td (stays) Tj ET' });
  const out = applyRedactions(file, { areas: AREA }).bytes;
  assert.equal(plainText(out), 1);
  assert.equal(bytesContain(out, 'hidden'), false);
  assert.ok(bytesContain(out, 'stays'));
  // Read back as a reader would: the word that stays is drawn, where it was.
  const read = pageGlyphs(out, 0);
  assert.match(read.text, /stays/);
  assert.doesNotMatch(read.text, /hidden/);
  const first = read.boxes[read.text.indexOf('stays')];
  assert.ok(first && Math.abs(first[0] - 400) < 1 && Math.abs(first[1] - 600) < 20, `drawn at ${JSON.stringify(first)}`);
});

test('a soft mask larger than its picture is painted out over the part the mark covers, scaled to the mask', () => {
  // The picture is 40 wide, its mask 80. The mark covers the left half of the picture: in the mask that is 0 to 40, and the word is at 25.
  const mask = Buffer.alloc(80 * 80, 255);
  mask.write(SECRET, 10 * 80 + 25, 'latin1');
  const file = pdf({
    content: 'q 200 0 0 200 100 100 cm /Im0 Do Q',
    resources: '/XObject << /Im0 6 0 R >>',
    objects: [picture('/SMask 7 0 R'), { data: mask.toString('latin1'), dict: '/Type /XObject /Subtype /Image /Width 80 /Height 80 /BitsPerComponent 8 /ColorSpace /DeviceGray' }],
  });
  const out = applyRedactions(file, { areas: [{ page: 0, rects: [[100, 100, 200, 300]] }] });
  assert.equal(out.removed.pictures, 1);
  assert.equal(bytesContain(out.bytes, SECRET), false);
});

test('reading-aloud text around a picture under a mark goes with the picture', () => {
  const file = pdf({
    content: `/Figure << /ActualText (${SECRET}) /Alt (${SECRET} too) >> BDC q 200 0 0 200 100 100 cm /Im0 Do Q EMC BT /F1 12 Tf 400 600 Td (kept) Tj ET`,
    resources: '/XObject << /Im0 6 0 R >>',
    objects: [picture()],
  });
  assert.ok(bytesContain(file, SECRET));
  const out = applyRedactions(file, { areas: AREA });
  assert.equal(bytesContain(out.bytes, SECRET), false);
  assert.ok(bytesContain(out.bytes, 'kept'));
});

// ── found by the second review, 10 Nov 2026 ──────────────────────────────

/** One page whose Font resource names the composite font in object 6 (its CIDFont in 7). */
function composite({ content, encoding, descendant = '/DW 1000', extra = [] }) {
  const type0 = `<< /Type /Font /Subtype /Type0 /BaseFont /X /Encoding ${encoding} /DescendantFonts [7 0 R] >>`;
  const cid = `<< /Type /Font /Subtype /CIDFontType2 /BaseFont /X /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> ${descendant} >>`;
  const file = pdf({ content, objects: [type0, cid, ...extra] });
  return rebuild(Buffer.from(file.toString('latin1').replace('/Font << /F1 5 0 R >>', '/Font << /F2 6 0 R >>'), 'latin1'));
}
const hex = (word) => Buffer.from(word, 'utf16le').swap16().toString('hex');

test('words written down a column are found where they are drawn: the column under the mark goes, the one beside it stays', () => {
  const content = `BT /F2 20 Tf 200 280 Td <${hex(SECRET)}> Tj ET BT /F2 20 Tf 500 280 Td <${hex('KEPT')}> Tj ET`;
  const out = applyRedactions(composite({ content, encoding: '/Identity-V' }), { areas: [{ page: 0, rects: [[190, 40, 230, 290]] }] });
  assert.equal(out.removed.glyphs, SECRET.length, 'every glyph of the column');
  assert.equal(bytesContain(out.bytes, SECRET), false);
  assert.equal(bytesContain(out.bytes, SECRET.slice(1)), false, 'not one letter left behind');
  assert.ok(bytesContain(out.bytes, 'KEPT'));
});

test('a column only the lower part of which is marked keeps the part above it', () => {
  // Eleven glyphs from y=280 down at 20 a glyph; the mark starts below the sixth.
  const content = `BT /F2 20 Tf 200 280 Td <${hex('ABCDEFGHIJK')}> Tj ET`;
  const out = applyRedactions(composite({ content, encoding: '/Identity-V' }), { areas: [{ page: 0, rects: [[190, 40, 230, 165]] }] });
  assert.ok(bytesContain(out.bytes, 'ABCDE'), 'the top of the column');
  assert.equal(bytesContain(out.bytes, 'GHIJK'), false, 'the bottom of it');
  assert.ok(out.removed.glyphs >= 5 && out.removed.glyphs <= 6, `cut ${out.removed.glyphs}`);
});

test('a vertical encoding other than Identity is judged whole, down the column', () => {
  const content = `BT /F2 20 Tf 200 280 Td <${hex(SECRET)}> Tj ET`;
  const out = applyRedactions(composite({ content, encoding: '/UniJIS-UTF16-V' }), { areas: [{ page: 0, rects: [[190, 40, 230, 120]] }] });
  assert.equal(bytesContain(out.bytes, SECRET), false);
});

test('a one-byte composite font: a mark at the end of a long string takes the string', () => {
  const content = `BT /F2 10 Tf 100 200 Td (Name: John Q ${SECRET}) Tj ET`;
  // 24 bytes of width 1 em at size 10: the secret is drawn from x=230 on, past half the string's bytes.
  const cmap = { data: '/CIDInit /ProcSet findresource begin 12 dict begin begincmap endcmap end end', dict: '/Type /CMap /CMapName /X' };
  const out = applyRedactions(composite({ content, encoding: '8 0 R', extra: [cmap] }), { areas: [{ page: 0, rects: [[230, 195, 340, 215]] }] });
  assert.equal(bytesContain(out.bytes, SECRET), false);
});

test('character spacing counts towards the length of a string judged whole', () => {
  const content = `BT /F2 10 Tf 5 Tc 100 200 Td (${SECRET}) Tj ET`;
  // Eleven bytes of 10 points plus 5 of spacing reach x=265; without the spacing, 210.
  const out = applyRedactions(composite({ content, encoding: '/90ms-RKSJ-H' }), { areas: [{ page: 0, rects: [[225, 195, 270, 215]] }] });
  assert.equal(bytesContain(out.bytes, SECRET), false);
});

test('strings after one judged whole are measured from the farthest it could have reached', () => {
  const content = `BT /F2 10 Tf 100 200 Td [(AB) (${SECRET})] TJ ET`;
  // At one byte a code the second string starts at 120 and ends at 230; the shortest guess would end it at 170.
  const out = applyRedactions(composite({ content, encoding: '/90ms-RKSJ-H' }), { areas: [{ page: 0, rects: [[200, 195, 240, 215]] }] });
  assert.equal(bytesContain(out.bytes, SECRET), false);
});

test('a form that draws itself does not hang the program, and is treated as unreadable', () => {
  const body = `${'/X Do '.repeat(20)} BT /F1 12 Tf 10 10 Td (${SECRET}) Tj ET`;
  const file = pdf({
    content: 'q 1 0 0 1 100 100 cm /X Do Q',
    resources: '/XObject << /X 6 0 R >>',
    objects: [{ data: body, dict: '/Type /XObject /Subtype /Form /BBox [0 0 200 200] /Resources << /XObject << /X 6 0 R >> /Font << /F1 5 0 R >> >>' }],
  });
  const started = Date.now();
  const out = applyRedactions(file, { areas: AREA });
  assert.ok(Date.now() - started < 8000, 'came back');
  assert.equal(bytesContain(out.bytes, SECRET), false);
  assert.ok(out.unpainted.length, 'said');
});

test('reading a page whose forms call each other a dozen levels deep comes back', () => {
  const objects = [];
  const levels = 12;
  for (let i = 0; i < levels; i++) {
    const next = i + 1 < levels ? `/Xobject << /N ${7 + i} 0 R >>`.replace('Xobject', 'XObject') : '';
    objects.push({ data: i + 1 < levels ? '/N Do '.repeat(20) : 'BT /F1 12 Tf 10 10 Td (deep) Tj ET', dict: `/Type /XObject /Subtype /Form /BBox [0 0 612 792] /Resources << ${next} /Font << /F1 5 0 R >> >>` });
  }
  const file = pdf({ content: '/N Do', resources: '/XObject << /N 6 0 R >>', objects });
  const started = Date.now();
  const read = pageGlyphs(file, 0);
  assert.ok(Date.now() - started < 8000, `took ${Date.now() - started} ms`);
  assert.equal(read.unsure, true, 'and says it could not read it all');
});

test('a form too deep to follow, under a mark, leaves the page', () => {
  const objects = [];
  for (let i = 0; i < 14; i++) {
    const last = i === 13;
    objects.push({ data: last ? `BT /F1 12 Tf 10 10 Td (${SECRET}) Tj ET` : '/N Do', dict: `/Type /XObject /Subtype /Form /BBox [0 0 200 200] /Resources << ${last ? '' : `/XObject << /N ${7 + i} 0 R >>`} /Font << /F1 5 0 R >> >>` });
  }
  const file = pdf({ content: 'q 1 0 0 1 100 100 cm /N Do Q', resources: '/XObject << /N 6 0 R >>', objects });
  assert.ok(bytesContain(file, SECRET));
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('a form that cannot be unpacked, under a mark, leaves the page', () => {
  const file = pdf({
    content: 'q 1 0 0 1 100 100 cm /X Do Q',
    resources: '/XObject << /X 6 0 R >>',
    objects: [{ data: `not deflate ${SECRET}`, dict: '/Type /XObject /Subtype /Form /BBox [0 0 200 200] /Filter /FlateDecode' }],
  });
  const out = applyRedactions(file, { areas: AREA });
  assert.equal(out.bytes.includes(Buffer.from(SECRET)), false);
  assert.ok(out.unpainted.length, 'said');
});

test('a pattern two forms share goes from both', () => {
  const tile = '/Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 100 100] /XStep 100 /YStep 100 /Resources << /Font << /F1 5 0 R >> >>';
  const form = '/Type /XObject /Subtype /Form /BBox [0 0 200 200] /Resources << /Pattern << /P 8 0 R >> >>';
  const paint = '/Pattern cs /P scn 0 0 100 100 re f';
  const file = pdf({
    content: 'q 1 0 0 1 100 100 cm /A Do Q q 1 0 0 1 120 120 cm /B Do Q',
    resources: '/XObject << /A 6 0 R /B 7 0 R >>',
    objects: [{ data: paint, dict: form }, { data: paint, dict: form }, { data: `BT /F1 12 Tf 10 10 Td (${SECRET}) Tj ET`, dict: tile }],
  });
  assert.ok(bytesContain(file, SECRET));
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('a pattern and a soft mask nowhere near the mark are left as they are', () => {
  const file = pdf({
    content: '/Pattern cs /P0 scn 400 600 100 100 re f /G0 gs 0 0 1 rg 400 400 50 50 re f',
    resources: '/Pattern << /P0 6 0 R >> /ExtGState << /G0 << /SMask << /Type /Mask /S /Luminosity /G 7 0 R >> >> >>',
    objects: [
      { data: 'BT /F1 12 Tf 10 10 Td (tile words) Tj ET', dict: '/Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 100 100] /XStep 100 /YStep 100 /Resources << /Font << /F1 5 0 R >> >>' },
      { data: 'BT /F1 12 Tf 410 410 Td (mask words) Tj ET', dict: '/Type /XObject /Subtype /Form /BBox [400 400 450 450] /Group << /S /Transparency /CS /DeviceGray >> /Resources << /Font << /F1 5 0 R >> >>' },
    ],
  });
  const out = applyRedactions(file, { areas: AREA });
  assert.ok(bytesContain(out.bytes, 'tile words'));
  assert.ok(bytesContain(out.bytes, 'mask words'));
  assert.deepEqual(out.unpainted, []);
});

test('a pattern that paints words under the mark goes, though no shape was drawn with it', () => {
  const file = pdf({
    content: `/Pattern cs /P0 scn BT /F1 12 Tf 110 200 Td (covered) Tj ET`,
    resources: '/Pattern << /P0 6 0 R >>',
    objects: [{ data: `BT /F1 12 Tf 10 10 Td (${SECRET}) Tj ET`, dict: '/Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 100 100] /XStep 100 /YStep 100 /Resources << /Font << /F1 5 0 R >> >>' }],
  });
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('a span the stream never closes still gives up what it carried for reading aloud', () => {
  const file = pdf({ content: `/Span << /ActualText (${SECRET}) >> BDC BT /F1 12 Tf 110 200 Td (covered) Tj ET` });
  assert.ok(bytesContain(file, SECRET));
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('a reply to a comment under the mark goes with it, and a reply to that', () => {
  const note = (rect, words, extra = '') => `<< /Type /Annot /Subtype /Text /Rect [${rect}] /Contents (${words}) ${extra} >>`;
  const base = pdf({ content: 'BT /F1 12 Tf 110 200 Td (covered) Tj ET', objects: [note('120 120 140 140', SECRET), note('500 700 520 720', 'REPLY-ONE', '/IRT 6 0 R'), note('540 700 560 720', 'REPLY-TWO', '/IRT 7 0 R'), note('580 700 600 720', 'OTHER-NOTE')] });
  const file = rebuild(Buffer.from(base.toString('latin1').replace('/Contents 4 0 R', '/Contents 4 0 R /Annots [6 0 R 7 0 R 8 0 R 9 0 R]'), 'latin1'));
  const out = applyRedactions(file, { areas: AREA });
  assert.equal(bytesContain(out.bytes, SECRET), false, 'the comment');
  assert.equal(bytesContain(out.bytes, 'REPLY-ONE'), false);
  assert.equal(bytesContain(out.bytes, 'REPLY-TWO'), false);
  assert.ok(bytesContain(out.bytes, 'OTHER-NOTE'));
  assert.equal(out.removed.annotations, 3);
});

for (const [filter, encode] of [
  ['/RunLengthDecode', () => Buffer.concat(Array.from({ length: 600000 }, () => Buffer.from([129, 65]))).toString('latin1')],
]) {
  test(`a stream under ${filter} that would unpack past the limit is refused, and quickly`, () => {
    const file = pdf({ content: 'BT ET', objects: [{ data: encode(), dict: `/Filter ${filter}` }] });
    const reader = openPdf(file);
    const stream = reader.getObject(new Ref(6, 0));
    const started = Date.now();
    assert.throws(() => reader.decode(stream), /more than this program can hold/);
    assert.ok(Date.now() - started < 8000);
  });
}

// ── found by the third review, 10 Nov 2026 ───────────────────────────────

const pdfReader = require('../src/reader');
const redactModule = require('../src/redact');

const withFont = (file, dict) => rebuild(Buffer.from(file.toString('latin1').replace(FONT, dict), 'latin1'));
const withAnnots = (file, list) => rebuild(Buffer.from(file.toString('latin1').replace('/Contents 4 0 R', `/Contents 4 0 R /Annots [${list}]`), 'latin1'));
const withCatalog = (file, words) => rebuild(Buffer.from(file.toString('latin1').replace('<< /Type /Catalog /Pages 2 0 R >>', `<< /Type /Catalog /Pages 2 0 R ${words} >>`), 'latin1'));
const TIMES = '<< /Type /Font /Subtype /Type1 /BaseFont /Times-Roman >>';
const COVERED = `BT /F1 12 Tf 110 200 Td (${SECRET}) Tj ET`;

test('a font without widths this program does not know: a mark at the end of a long string takes the string', () => {
  // A W of Times is 0.944 em, so the secret is drawn from about x=445. At the
  // old guess of 0.55 em it would have ended at 430, short of the mark.
  const file = withFont(pdf({ content: `BT /F1 20 Tf 100 200 Td (${'W'.repeat(18)} ${SECRET}) Tj ET` }), TIMES);
  const out = applyRedactions(file, { areas: [{ page: 0, rects: [[445, 190, 600, 230]] }] });
  assert.equal(bytesContain(out.bytes, SECRET), false);
  assert.equal(out.removed.glyphs, 18 + 1 + SECRET.length, 'one code a byte');
});

test('a font without widths this program does not know: a string away from the mark stays', () => {
  const file = withFont(pdf({ content: `BT /F1 20 Tf 100 600 Td (${SECRET}) Tj ET` }), TIMES);
  assert.ok(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET));
});

test('a font without widths this program does not know: reading the page finds the words and says the places are a guess', () => {
  const read = pageGlyphs(withFont(pdf({ content: `BT /F1 20 Tf 100 600 Td (${SECRET}) Tj ET` }), TIMES), 0);
  assert.match(read.text, new RegExp(SECRET));
  assert.equal(read.unsure, true);
});

test('words the same font spells out for itself are told one code a byte', () => {
  const cmap = { data: 'beginbfchar\n<41> <0058>\nendbfchar', dict: '' };
  const base = pdf({ content: 'BT /F1 20 Tf 110 200 Td (AAAA) Tj ET', objects: [cmap] });
  const file = withFont(base, TIMES.replace('>>', '/ToUnicode 6 0 R >>'));
  const out = applyRedactions(file, { areas: AREA });
  assert.deepEqual(out.words, ['XXXX']);
});

const FORM = '/Type /XObject /Subtype /Form /BBox [0 0 612 792]';

test('a form that draws with the font of the page around it: its words go', () => {
  const file = pdf({ content: 'BT /F1 14 Tf ET /Fm0 Do', resources: '/XObject << /Fm0 6 0 R >>', objects: [{ data: `BT 110 200 Td (${SECRET}) Tj ET`, dict: `${FORM} /Resources << >>` }] });
  assert.ok(bytesContain(file, SECRET));
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('a form measures its words with the spacing the page set', () => {
  // With 20 points between letters the end of the string lies at x=338 to 426; without them it would end by 175.
  const file = pdf({ content: 'BT /F1 10 Tf 20 Tc ET /Fm0 Do', resources: '/XObject << /Fm0 6 0 R >>', objects: [{ data: 'BT 100 200 Td (ABCDEFGHIXYZQ) Tj ET', dict: `${FORM} /Resources << >>` }] });
  const out = applyRedactions(file, { areas: [{ page: 0, rects: [[330, 190, 430, 220]] }] });
  assert.equal(bytesContain(out.bytes, 'XYZQ'), false);
  assert.ok(bytesContain(out.bytes, 'ABCDEFGHI'));
});

test('a form drawing with the font of the page uses that font, though it has another under the same name', () => {
  // Under the form's own F1 each letter would be 200 points wide and most of the word would lie beyond the mark.
  const wide = `<< /Type /Font /Subtype /Type1 /BaseFont /Wide /FirstChar 32 /LastChar 126 /Widths [${Array(95).fill(20000).join(' ')}] >>`;
  const file = pdf({ content: 'BT /F1 10 Tf ET /Fm0 Do', resources: '/XObject << /Fm0 6 0 R >>', objects: [{ data: `BT 110 200 Td (${SECRET}) Tj ET`, dict: `${FORM} /Resources << /Font << /F1 7 0 R >> >>` }, wide] });
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('a form drawn under the page\'s own pattern: the pattern goes when its words do', () => {
  const file = pdf({
    content: '/Pattern cs /P0 scn /Fm0 Do',
    resources: '/XObject << /Fm0 6 0 R >> /Pattern << /P0 7 0 R >>',
    objects: [
      { data: 'BT /F1 12 Tf 110 200 Td (covered) Tj ET', dict: `${FORM} /Resources << /Font << /F1 5 0 R >> >>` },
      { data: `BT /F1 12 Tf 10 10 Td (${SECRET}) Tj ET`, dict: '/Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 100 100] /XStep 100 /YStep 100 /Resources << /Font << /F1 5 0 R >> >>' },
    ],
  });
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('words outlined with a pattern under the mark: the pattern goes', () => {
  const file = pdf({
    content: '/Pattern CS /P0 SCN 1 Tr BT /F1 12 Tf 110 200 Td (covered) Tj ET',
    resources: '/Pattern << /P0 6 0 R >>',
    objects: [{ data: `BT /F1 12 Tf 10 10 Td (${SECRET}) Tj ET`, dict: '/Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 100 100] /XStep 100 /YStep 100 /Resources << /Font << /F1 5 0 R >> >>' }],
  });
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('a text map of a great many ranges comes back quickly and is held to what a font can name', () => {
  const started = Date.now();
  const map = redactModule.parseToUnicode(Buffer.from(`beginbfrange\n${'<0000> <FFFF> <0041>\n'.repeat(1000)}endbfrange`, 'latin1'));
  assert.ok(Date.now() - started < 3000, `${Date.now() - started} ms`);
  assert.ok(map.size <= 65536);
});

test('a text map a font could really have is read whole', () => {
  const map = redactModule.parseToUnicode(Buffer.from('beginbfrange\n<0000> <FFFF> <0041>\nendbfrange', 'latin1'));
  assert.equal(map.size, 65536);
  assert.equal(map.get(1), 'B');
});

const stubFile = { resolve: (x) => x, decode: () => Buffer.alloc(0) };
const compositeDict = (encoding, descendant) => new pdfReader.Dict([['Subtype', new pdfReader.Name('Type0')], ['Encoding', new pdfReader.Name(encoding)], ['DescendantFonts', [new pdfReader.Dict(descendant)]]]);

test('a width list of a great many ranges comes back quickly, and its strings are then judged whole', () => {
  const started = Date.now();
  const info = redactModule.fontInfo(stubFile, compositeDict('Identity-H', [['W', Array.from({ length: 20000 }, () => [0, 65535, 500]).flat()]]));
  assert.ok(Date.now() - started < 3000, `${Date.now() - started} ms`);
  assert.equal(info.unsure, true);
  assert.ok(info.widest >= 2);
  assert.equal(info.narrowest, 0);
});

test('a column-width list of a great many ranges comes back quickly, and its strings are then judged whole', () => {
  const started = Date.now();
  const info = redactModule.fontInfo(stubFile, compositeDict('Identity-V', [['W2', Array.from({ length: 20000 }, () => [0, 65535, -1000, 500, 880]).flat()]]));
  assert.ok(Date.now() - started < 3000, `${Date.now() - started} ms`);
  assert.equal(info.unsure, true);
  assert.ok(info.widest >= 2);
});

test('a width list a font could really have stays exact', () => {
  const info = redactModule.fontInfo(stubFile, compositeDict('Identity-H', [['W', [0, 65535, 500]]]));
  assert.equal(info.unsure, false);
  assert.equal(info.width(7), 0.5);
});

test('a page that saves its state without end is refused before it can fill the memory', () => {
  const file = pdf({ content: 'q\n'.repeat(5000) });
  assert.throws(() => applyRedactions(file, { areas: AREA }), (e) => e.name === 'PdfRedactError' && e.code === 'nested');
});

test('saves that are given back do not count against a page', () => {
  const file = pdf({ content: `${'q Q\n'.repeat(5000)}${COVERED}` });
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});

test('a layer named with the word: the name goes, the layer and its place in the list stay', () => {
  const base = pdf({ content: COVERED, objects: [`<< /Type /OCG /Name (${SECRET} layer) >>`, '<< /Type /OCG /Name (Plain layer) >>'] });
  const file = withCatalog(base, `/OCProperties << /OCGs [6 0 R 7 0 R] /D << /Name (${SECRET} view) /Order [(${SECRET} group) 6 0 R [(Other) 7 0 R]] >> >>`);
  const out = applyRedactions(file, { areas: AREA });
  assert.equal(bytesContain(out.bytes, SECRET), false);
  const read = openPdf(out.bytes);
  const props = read.resolve(read.catalog().get('OCProperties'));
  assert.equal(read.resolve(props.get('OCGs')).length, 2);
  const order = read.resolve(read.resolve(props.get('D')).get('Order'));
  assert.equal(order.length, 3);
  assert.equal(read.resolve(order[1]).get('Type').name, 'OCG', 'the list still names the layer, under its new number');
  assert.equal(read.resolve(order[2])[0].text(), 'Other');
  assert.ok(bytesContain(out.bytes, 'Plain layer'));
});

test('a layer whose usage names the word loses the usage and keeps its name', () => {
  const base = pdf({ content: COVERED, objects: [`<< /Type /OCG /Name (Plain layer) /Usage << /User << /Type /Ind /Name (${SECRET}) >> >> >>`] });
  const file = withCatalog(base, '/OCProperties << /OCGs [6 0 R] /D << /Order [6 0 R] >> >>');
  const out = applyRedactions(file, { areas: AREA });
  assert.equal(bytesContain(out.bytes, SECRET), false);
  assert.ok(bytesContain(out.bytes, 'Plain layer'));
});

test('a form value kept as a stream with the word in it is blanked', () => {
  const base = pdf({ content: COVERED, objects: ['<< /Type /Annot /Subtype /Widget /FT /Tx /T (plain) /V 7 0 R /Rect [400 600 450 620] >>', { data: `Owner: ${SECRET}`, dict: '' }] });
  const out = applyRedactions(withAnnots(base, '6 0 R'), { areas: AREA });
  assert.equal(bytesContain(out.bytes, SECRET), false);
});

test('a form value kept as a stream without the word stays', () => {
  const base = pdf({ content: COVERED, objects: ['<< /Type /Annot /Subtype /Widget /FT /Tx /T (plain) /V 7 0 R /Rect [400 600 450 620] >>', { data: 'Nothing to hide', dict: '' }] });
  assert.ok(bytesContain(applyRedactions(withAnnots(base, '6 0 R'), { areas: AREA }).bytes, 'Nothing to hide'));
});

test('a comment whose rich text is a stream with the word in it goes', () => {
  const base = pdf({ content: COVERED, objects: ['<< /Type /Annot /Subtype /Text /Rect [500 700 520 720] /Contents (plain) /RC 7 0 R >>', { data: `<body><p>${SECRET}</p></body>`, dict: '' }] });
  const out = applyRedactions(withAnnots(base, '6 0 R'), { areas: AREA });
  assert.equal(bytesContain(out.bytes, SECRET), false);
  assert.equal(out.elsewhere.comment, 1);
});

test('a comment whose rich text splits the word across its markup goes too', () => {
  const base = pdf({ content: COVERED, objects: ['<< /Type /Annot /Subtype /Text /Rect [500 700 520 720] /Contents (plain) /RC 7 0 R >>', { data: '<body><p>SECR<b>ET-4417</b></p></body>', dict: '' }] });
  assert.equal(applyRedactions(withAnnots(base, '6 0 R'), { areas: AREA }).elsewhere.comment, 1);
});

test('a comment whose rich text is a stream that cannot be read goes: what it says cannot be checked', () => {
  const base = pdf({ content: COVERED, objects: ['<< /Type /Annot /Subtype /Text /Rect [500 700 520 720] /Contents (plain) /RC 7 0 R >>', { data: 'not flate at all', dict: '/Filter /FlateDecode' }] });
  assert.equal(applyRedactions(withAnnots(base, '6 0 R'), { areas: AREA }).elsewhere.comment, 1);
});

test('an array that names itself does not bring the scrub down', () => {
  const base = pdf({ content: COVERED, objects: ['[6 0 R]'] });
  const file = rebuild(Buffer.from(base.toString('latin1').replace('<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Catalog /Pages 2 0 R /OCProperties << /D << /Order 6 0 R >> >> >>'), 'latin1'));
  assert.equal(bytesContain(applyRedactions(file, { areas: AREA }).bytes, SECRET), false);
});
