'use strict';

/**
 * Base-14 font widths.
 *
 * == Why this file is the reason the package exists =====================
 *
 * Every PDF reader in the world ships the same fourteen fonts, and their
 * advance widths are not a rendering choice - they are fixed by the format.
 * A viewer laying out Helvetica MUST use the widths below, because they are
 * the widths Adobe published and every implementation since has copied.
 *
 * That turns a hard problem into an easy one. `@rutba/drawing/measure` says
 * of itself that it is "an APPROXIMATION", and `doc-view/paginate` names the
 * cost honestly: line breaks are close to, not identical to, what the reader
 * will draw, "and for print-exact fidelity it is not good enough". For HTML
 * that is unavoidable - the browser has its own fonts and its own opinion.
 * For a PDF drawn in a base-14 font it simply is not true: we measure with
 * the same numbers the viewer will, so measurement and rendering agree by
 * construction rather than by luck.
 *
 * The whole of a printable invoice follows from that. No font embedding, no
 * headless browser, no dependency - just arithmetic that happens to be the
 * same arithmetic on both sides.
 *
 * == How much to trust the table ========================================
 *
 * ASCII is cross-checked in `test/metrics.test.js` against the estate's own
 * `@rutba/drawing/measure` table, which was written as "proportional-sans
 * averages" and turns out to be this font: the distinctive widths - W 944,
 * m 833, i 222, asciicircum 469, quotesingle 191, bar 260 - match to a
 * rounding step, which a table of averages would not do by luck.
 *
 * It is not a byte-for-byte match, and the test does not pretend otherwise.
 * 91 of the 95 ASCII glyphs agree to within a rounding step; the four that
 * do not are space and exclam (0.26 there against 0.278 here, 6.5%) and
 * percent and at (about 2%) - which is exactly the "may differ by a few
 * percent" that file claims for itself. Two tables written from different
 * sources landing on the same font is the check; the residue is why the
 * approximate one cannot be used to place glyphs.
 *
 * The Latin-1 and typographic ranges are transcribed from the AFM files and
 * are far less exercised: a stray digit there would widen one accented
 * character by a fraction of a millimetre. Worth knowing, not worth
 * pretending otherwise.
 *
 * The obliques are not a transcription error. A base-14 oblique is the
 * roman sheared, so Helvetica-Oblique has Helvetica's metrics exactly.
 */

/** ASCII 32..126, in 1000ths of the em. */
const HELVETICA_ASCII = [
  278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278,
  556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,
  1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778,
  667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,
  333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556,
  556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584,
];

const HELVETICA_BOLD_ASCII = [
  278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278,
  556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611,
  975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778,
  667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556,
  333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611,
  611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584,
];

/**
 * WinAnsi 128..255. Index 0 is byte 0x80. A hole in the encoding (0x81,
 * 0x8D, 0x8F, 0x90, 0x9D) has no glyph and no width; it is never emitted
 * because `encoding.js` will not map anything to it.
 */
const HELVETICA_HIGH = [
  556, 0, 222, 556, 333, 1000, 556, 556, 333, 1000, 667, 333, 1000, 0, 611, 0,
  0, 222, 222, 333, 333, 350, 556, 1000, 333, 1000, 500, 333, 944, 0, 500, 667,
  278, 333, 556, 556, 556, 556, 260, 556, 333, 737, 370, 556, 584, 333, 737, 333,
  400, 584, 333, 333, 333, 556, 537, 278, 333, 333, 365, 556, 834, 834, 834, 611,
  667, 667, 667, 667, 667, 667, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278,
  722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611,
  556, 556, 556, 556, 556, 556, 889, 500, 556, 556, 556, 556, 278, 278, 278, 278,
  556, 556, 556, 556, 556, 556, 556, 584, 611, 556, 556, 556, 556, 500, 556, 500,
];

const HELVETICA_BOLD_HIGH = [
  556, 0, 278, 556, 500, 1000, 556, 556, 333, 1000, 667, 333, 1000, 0, 611, 0,
  0, 278, 278, 500, 500, 350, 556, 1000, 333, 1000, 556, 333, 889, 0, 500, 667,
  278, 333, 556, 556, 556, 556, 280, 556, 333, 737, 370, 556, 584, 333, 737, 333,
  400, 584, 333, 333, 333, 611, 556, 278, 333, 333, 365, 556, 834, 834, 834, 611,
  722, 722, 722, 722, 722, 722, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278,
  722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611,
  556, 556, 556, 556, 556, 556, 889, 556, 556, 556, 556, 556, 278, 278, 278, 278,
  611, 611, 611, 611, 611, 611, 611, 584, 611, 611, 611, 611, 611, 556, 611, 556,
];

/**
 * The faces this package can draw with.
 *
 * Deliberately not all fourteen. Times and Symbol would each need another
 * transcribed table, and a font whose widths I cannot check is worse than a
 * font the package does not offer: the failure is silent, and shows up as a
 * line that overflows its column in the reader's viewer and nowhere else.
 */
const FONTS = {
  Helvetica: { base: 'Helvetica', ascii: HELVETICA_ASCII, high: HELVETICA_HIGH },
  'Helvetica-Bold': { base: 'Helvetica-Bold', ascii: HELVETICA_BOLD_ASCII, high: HELVETICA_BOLD_HIGH },
  // Sheared, not redrawn: identical metrics to the roman.
  'Helvetica-Oblique': { base: 'Helvetica-Oblique', ascii: HELVETICA_ASCII, high: HELVETICA_HIGH },
  'Helvetica-BoldOblique': { base: 'Helvetica-BoldOblique', ascii: HELVETICA_BOLD_ASCII, high: HELVETICA_BOLD_HIGH },
  // Every glyph 600/1000 - which is what "monospace" means.
  Courier: { base: 'Courier', fixed: 600 },
  'Courier-Bold': { base: 'Courier-Bold', fixed: 600 },
};

function fontOrThrow(name) {
  const font = FONTS[name];
  if (!font) {
    throw new Error(
      `${name} is not one of the fonts this package has widths for `
      + `(${Object.keys(FONTS).join(', ')}). Adding one means adding its width table, not just its name.`
    );
  }
  return font;
}

/** Advance width of one WinAnsi byte, in 1000ths of the em. */
function byteWidth(fontName, byte) {
  const font = fontOrThrow(fontName);
  if (font.fixed) return font.fixed;
  if (byte >= 32 && byte <= 126) return font.ascii[byte - 32];
  if (byte >= 128 && byte <= 255) return font.high[byte - 128] || font.ascii[63 - 32];
  // Control bytes never reach here - the encoder replaces them - but a
  // width of zero would silently shorten a line, so answer with something.
  return font.ascii[0];
}

/**
 * Width of already-encoded bytes, in points at `size`.
 *
 * Takes bytes rather than a string on purpose: the encoding decides which
 * glyph each character became, and measuring the string would measure the
 * character we wanted rather than the one that will be drawn.
 */
function widthOfBytes(bytes, fontName, size) {
  let total = 0;
  for (const byte of bytes) total += byteWidth(fontName, byte);
  return (total * size) / 1000;
}

/**
 * Line spacing for a size. 1.2 is the typographic default and the one a
 * reader expects from an invoice; the sheet may override it per block.
 */
const lineHeight = (size) => size * 1.2;

/** Cap height, for putting a baseline where a box wants text centred. */
const capHeight = (size) => size * 0.717;

module.exports = {
  FONTS, byteWidth, widthOfBytes, lineHeight, capHeight, fontOrThrow,
  HELVETICA_ASCII, HELVETICA_BOLD_ASCII, HELVETICA_HIGH, HELVETICA_BOLD_HIGH,
};
