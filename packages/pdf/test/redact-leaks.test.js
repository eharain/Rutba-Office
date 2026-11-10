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

test('the file written can still be read, and the words outside the mark are still in it', () => {
  const file = pdf({ content: 'BT /F1 14 Tf 110 200 Td (hidden) Tj ET BT /F1 14 Tf 400 600 Td (stays) Tj ET' });
  const out = applyRedactions(file, { areas: AREA }).bytes;
  assert.equal(plainText(out), 1);
  assert.equal(bytesContain(out, 'hidden'), false);
  assert.ok(bytesContain(out, 'stays'));
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
