'use strict';

const test = require('node:test');
const assert = require('node:assert');
const { byteWidth, widthOfBytes, HELVETICA_ASCII, HELVETICA_BOLD_ASCII, FONTS } = require('../src/metrics');
const { encode } = require('../src/encoding');

test('every width table covers exactly the range it claims', () => {
  assert.strictEqual(HELVETICA_ASCII.length, 95, 'ASCII 32..126 is 95 glyphs');
  assert.strictEqual(HELVETICA_BOLD_ASCII.length, 95);
  for (const [name, font] of Object.entries(FONTS)) {
    if (font.fixed) continue;
    assert.strictEqual(font.ascii.length, 95, `${name} ascii`);
    assert.strictEqual(font.high.length, 128, `${name} high`);
  }
});

test('the obliques are the romans sheared, so their metrics are identical', () => {
  for (let b = 32; b <= 126; b++) {
    assert.strictEqual(byteWidth('Helvetica-Oblique', b), byteWidth('Helvetica', b));
    assert.strictEqual(byteWidth('Helvetica-BoldOblique', b), byteWidth('Helvetica-Bold', b));
  }
});

test('Courier is monospaced, which is the whole of its metrics', () => {
  const widths = new Set();
  for (let b = 32; b <= 255; b++) widths.add(byteWidth('Courier', b));
  assert.deepStrictEqual([...widths], [600]);
});

test('a font with no width table is refused by name, not silently substituted', () => {
  assert.throws(() => byteWidth('Times-Roman', 65), /not one of the fonts/);
  assert.throws(() => byteWidth('Comic Sans', 65), /width table, not just its name/);
});

test('bold is wider than roman where it should be, and equal where it should be', () => {
  // Digits are the same width in both faces - which is why a bold total
  // lines up under a roman column - and letters are not.
  for (const digit of '0123456789') {
    const b = digit.charCodeAt(0);
    assert.strictEqual(byteWidth('Helvetica-Bold', b), byteWidth('Helvetica', b), digit);
  }
  assert.ok(byteWidth('Helvetica-Bold', 98) > byteWidth('Helvetica', 98), 'bold b is wider');
});

/**
 * The check this file exists for.
 *
 * `@rutba/drawing/measure` was written independently, from a different
 * source, and describes itself as approximate proportional-sans averages.
 * If it lands on the same font as the AFM table here, both are very
 * probably right; if it did not, one of them is wrong and this is where
 * that shows up.
 */
test('the estate\'s own width table is this font', async () => {
  const { measureText } = await import('@rutba/drawing');

  // Glyphs whose widths could not agree by coincidence: nothing about
  // "average sans" predicts that W is 944 and l is 222.
  const distinctive = 'WmilfIjt^`\'|{}';
  for (const ch of distinctive) {
    const theirs = measureText(ch, { size: 1000 });
    const mine = byteWidth('Helvetica', ch.charCodeAt(0));
    assert.ok(Math.abs(theirs - mine) <= 6,
      `${JSON.stringify(ch)}: measure says ${theirs}, the AFM says ${mine}`);
  }

  // Across the whole of ASCII, all but a handful agree to a rounding step,
  // and none is further out than the "few percent" that file claims for
  // itself. A transposed digit in either table would be far larger.
  let close = 0;
  let worst = 0;
  for (let b = 32; b <= 126; b++) {
    const theirs = measureText(String.fromCharCode(b), { size: 1000 });
    const mine = byteWidth('Helvetica', b);
    if (Math.abs(theirs - mine) <= 6) close += 1;
    worst = Math.max(worst, Math.abs(theirs - mine) / mine);
  }
  assert.ok(close >= 90, `only ${close} of 95 glyphs agree to a rounding step`);
  assert.ok(worst < 0.07, `worst disagreement ${(worst * 100).toFixed(1)}% is too large to be rounding`);
});

test('measuring works on bytes, so it measures the glyph that will be drawn', () => {
  // A character with no WinAnsi glyph becomes '?', and must be measured as
  // '?' - measuring the original would reserve space for ink nobody sees.
  const bytes = encode('你');
  assert.deepStrictEqual(bytes, [0x3F]);
  assert.strictEqual(
    widthOfBytes(bytes, 'Helvetica', 10),
    widthOfBytes(encode('?'), 'Helvetica', 10),
  );
});

test('width scales with size exactly', () => {
  const bytes = encode('Invoice INV-0001');
  assert.ok(Math.abs(widthOfBytes(bytes, 'Helvetica', 20) - widthOfBytes(bytes, 'Helvetica', 10) * 2) < 1e-9);
});
