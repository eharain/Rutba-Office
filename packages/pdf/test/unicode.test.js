'use strict';

/**
 * Text the base-14 fonts cannot draw: Arabic joined, right-to-left lines put
 * in drawing order, and an embedded TrueType font subset to the glyphs used.
 * The joining and the ordering need no font; the embedding uses one this
 * computer has, and is skipped where it has none.
 */

const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert');
const zlib = require('node:zlib');
const { PdfDocument } = require('../src/document');
const { shapeArabic } = require('../src/shaping');
const { visualText, visualPieces } = require('../src/bidi');
const { readTrueType, subsetTrueType } = require('../src/truetype');

const hex = (s) => [...s].map((c) => c.codePointAt(0).toString(16).toUpperCase());

test('Arabic letters take the shape their neighbours give them, lam and alef one ligature, Urdu letters too', () => {
  // س initial, لا final ligature (the lam joins the seen), م alone after the alef.
  assert.deepEqual(hex(shapeArabic('سلام').text), ['FEB3', 'FEFC', 'FEE1']);
  // پاکستان: pe initial, alef final, kaf initial, seen and teh medial, alef final, noon alone.
  assert.deepEqual(hex(shapeArabic('پاکستان').text), ['FB58', 'FE8E', 'FB90', 'FEB4', 'FE98', 'FE8E', 'FEE5']);
  // A vowel mark does not break the join: beh-fatha-teh joins as beh-teh.
  assert.deepEqual(hex(shapeArabic('بَت').text), ['FE91', '64E', 'FE96']);
  assert.equal(shapeArabic('plain words').text, 'plain words');
});

test('A right-to-left line is put in drawing order: words reversed, numbers and Latin kept as they read, brackets mirrored', () => {
  assert.equal(visualText('abc אבג def'), 'abc גבא def');
  assert.equal(visualText('אב 123 ג', { rtl: true }), 'ג 123 בא');
  assert.equal(visualText('(אב)', { rtl: true }), '(בא)', 'the brackets mirrored as the run is reversed');
  assert.equal(visualText('אב Word 2026.', { rtl: true }), '.Word 2026 בא');
  const pieces = visualPieces([{ text: 'שלום ', bold: true }, { text: 'עולם' }], { rtl: true });
  assert.deepEqual(pieces.map((p) => [p.text, Boolean(p.bold)]), [['םלוע', false], [' םולש', true]], 'each run kept with its own look, the runs in drawing order');
});

const FONT = (() => {
  const dirs = [path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts'), '/usr/share/fonts/truetype/dejavu', '/System/Library/Fonts/Supplemental'];
  for (const dir of dirs) {
    for (const name of ['arial.ttf', 'DejaVuSans.ttf', 'Arial.ttf']) {
      try {
        const bytes = fs.readFileSync(path.join(dir, name));
        const f = readTrueType(bytes);
        if (f.embeddable && f.glyphOf(0x0628) && f.glyphOf(0x05D0)) return bytes;
      } catch { /* not here */ }
    }
  }
  return null;
})();

test('A TrueType font is read, and subset to the glyphs a page uses, the rest emptied', { skip: !FONT && 'no font with Arabic and Hebrew on this computer' }, () => {
  const font = readTrueType(FONT);
  assert.ok(font.unitsPerEm > 0 && font.numGlyphs > 100);
  const beh = font.glyphOf(0x0628);
  const subset = subsetTrueType(font, new Set([beh]));
  assert.ok(subset.length < FONT.length / 4, `${subset.length} of ${FONT.length} bytes`);
  const back = readTrueType(subset);
  assert.equal(back.numGlyphs, font.numGlyphs, 'glyph numbers kept');
  assert.equal(back.advance(beh), font.advance(beh));
  const [a, b] = back.glyphRange(beh);
  assert.ok(b > a, 'the glyph used is there');
  const other = font.glyphOf(0x05D0);
  const [c, d] = back.glyphRange(other);
  assert.equal(d, c, 'a glyph not used is emptied');
});

test('A document with an embedded font draws what WinAnsi cannot in it, glyph by glyph, with a ToUnicode map; Latin stays in Helvetica', { skip: !FONT && 'no font with Arabic and Hebrew on this computer' }, () => {
  const doc = new PdfDocument({ created: '2026-01-01T00:00:00Z', unicodeFont: FONT });
  const page = doc.addPage();
  page.text('Invoice', 40, 60);
  const w = page.text('مرحبا', 40, 90, { size: 12 });
  assert.ok(w > 0 && Math.abs(w - doc.widthOf('مرحبا', { size: 12 })) < 1e-9, 'the width drawn is the width measured');
  page.text('שלום', 40, 120);
  const text = doc.toBuffer().toString('latin1');
  assert.match(text, /\/Subtype \/Type0 \/BaseFont \/[A-Z]{6}\+\w+ \/Encoding \/Identity-H/);
  assert.match(text, /\/CIDToGIDMap \/Identity/);
  assert.match(text, /\/FontFile2 \d+ 0 R/);
  assert.match(text, /\/Type \/Font \/Subtype \/Type1 \/BaseFont \/Helvetica /, 'the Latin word in the viewer\'s own font');
  assert.match(text, /\(Invoice\) Tj/);
  assert.match(text, /<[0-9a-f]{8,}> Tj/, 'the Arabic and Hebrew written as glyph numbers');
  // The ToUnicode map gives the letters back.
  const streams = [...text.matchAll(/stream\n([\s\S]*?)\nendstream/g)].map((m) => m[1]);
  const cmap = streams.find((s) => s.includes('beginbfchar'));
  assert.ok(cmap && /<05E9>/.test(cmap) && /<0645>/.test(cmap), 'shin and meem mapped back');
  // The font program is a subset, deflated.
  const program = streams.map((s) => { try { return zlib.inflateSync(Buffer.from(s, 'latin1')); } catch { return null; } }).find((b) => b && b.readUInt32BE(0) === 0x00010000);
  assert.ok(program && program.length < FONT.length / 4);
});

test('Without an embedded font, text WinAnsi cannot say prints as ? as it always did', () => {
  const doc = new PdfDocument({ created: '2026-01-01T00:00:00Z' });
  doc.addPage().text('שלום', 40, 60);
  assert.match(doc.toBuffer().toString('latin1'), /\(\?\?\?\?\) Tj/);
  assert.equal(new PdfDocument().useUnicodeFont(Buffer.from('not a font')), false);
});
