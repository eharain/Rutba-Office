'use strict';

/**
 * Text into WinAnsi bytes.
 *
 * A PDF string is bytes, not characters. With a base-14 font and
 * `/WinAnsiEncoding`, byte 0xA3 is a pound sign and byte 0x92 is a right
 * single quote - so the encoder, not the caller, decides what a character
 * becomes, and `metrics.js` measures the bytes it produced rather than the
 * string it was given. Measuring the string would measure the character we
 * wanted instead of the glyph that will actually be drawn.
 *
 * WinAnsi is Latin-1 with a different 0x80-0x9F block. Those 27 codes are
 * where the typography lives - curly quotes, dashes, the euro, the ellipsis
 * - which is precisely the range a name or an address pasted out of Word
 * arrives in. Getting them wrong is not exotic: it is the common case.
 *
 * == What happens to a character with no glyph ==========================
 *
 * It is transliterated if there is an honest ASCII equivalent, and otherwise
 * becomes '?'. Not dropped. A dropped character shortens a name silently -
 * "Muller" for "Müller" reads as correct and is not - whereas a '?' is
 * visibly a substitution and someone will fix the data. This encoding cannot
 * write Greek, Cyrillic or CJK at all; the fix for those is an embedded font
 * with a real Unicode CMap, which is a much bigger machine and one to build
 * when somebody actually needs it rather than to fake now.
 */

/** Unicode above Latin-1 that WinAnsi keeps, in 0x80-0x9F. */
const HIGH_BLOCK = {
  '€': 0x80, '‚': 0x82, 'ƒ': 0x83, '„': 0x84,
  '…': 0x85, '†': 0x86, '‡': 0x87, 'ˆ': 0x88,
  '‰': 0x89, 'Š': 0x8A, '‹': 0x8B, 'Œ': 0x8C,
  'Ž': 0x8E, '‘': 0x91, '’': 0x92, '“': 0x93,
  '”': 0x94, '•': 0x95, '–': 0x96, '—': 0x97,
  '˜': 0x98, '™': 0x99, 'š': 0x9A, '›': 0x9B,
  'œ': 0x9C, 'ž': 0x9E, 'Ÿ': 0x9F,
};

/**
 * Characters worth spelling rather than replacing. Everything here reads
 * correctly as its substitute; anything that would read as a DIFFERENT word
 * belongs in the '?' case instead.
 *
 * A non-breaking space is deliberately NOT here. WinAnsi has one at 0xA0,
 * and it earns its keep: this package owns line breaking and the wrapper
 * splits on ordinary spaces only, so an author who wrote a hard space to
 * keep "INV 0001" on one line gets what they asked for. Flattening it to
 * 0x20 would throw that away for nothing.
 */
const TRANSLITERATE = {
  '−': '-', '‐': '-', '‑': '-', '‒': '-', '―': '-',
  ' ': ' ', ' ': ' ', ' ': ' ', '​': '',
  '⁄': '/', '′': "'", '″': '"',
  'Ł': 'L', 'ł': 'l', 'Đ': 'D', 'đ': 'd',
  '№': 'No.', '℃': 'degC',
  // Soft hyphen: invisible unless a break lands on it, and nothing here
  // implements discretionary breaks, so it must not print as a hyphen.
  '­': '',
};

const TAB_WIDTH = 4;

/**
 * Encode a string to WinAnsi bytes.
 *
 * Newlines and tabs are the caller's business, not the encoder's: a PDF
 * content stream has no concept of either, so they must have been resolved
 * into separate lines and spacing before text reaches here. Rather than emit
 * a control byte no font can draw, a tab becomes spaces and a newline
 * becomes a space - visible evidence that something upstream skipped a step.
 */
function encode(value) {
  const text = value === null || value === undefined ? '' : String(value);
  const bytes = [];
  for (const ch of text) {
    const swap = TRANSLITERATE[ch];
    if (swap !== undefined) {
      for (const s of swap) bytes.push(s.charCodeAt(0));
      continue;
    }
    if (ch === '\t') { for (let i = 0; i < TAB_WIDTH; i++) bytes.push(32); continue; }
    if (ch === '\n' || ch === '\r') { bytes.push(32); continue; }
    const code = ch.codePointAt(0);
    if (code >= 32 && code <= 126) { bytes.push(code); continue; }
    if (code >= 0xA0 && code <= 0xFF) { bytes.push(code); continue; }
    const high = HIGH_BLOCK[ch];
    if (high !== undefined) { bytes.push(high); continue; }
    bytes.push(0x3F); // '?'
  }
  return bytes;
}

/**
 * A PDF literal string: `(...)` with backslash, parens and high bytes
 * escaped.
 *
 * The parens matter more than they look. An unescaped ')' in a customer name
 * - "Acme (UK) Ltd" - closes the string early and the rest of the page
 * becomes syntax. The file then either fails to open or opens with content
 * missing, and which of the two you get depends on the viewer.
 */
function literal(bytes) {
  let out = '(';
  for (const b of bytes) {
    if (b === 0x28 || b === 0x29 || b === 0x5C) out += '\\' + String.fromCharCode(b);
    else if (b < 32 || b > 126) out += '\\' + b.toString(8).padStart(3, '0');
    else out += String.fromCharCode(b);
  }
  return out + ')';
}

/** Encode and escape in one step - what a content stream actually wants. */
const pdfString = (value) => literal(encode(value));

module.exports = { encode, literal, pdfString, HIGH_BLOCK, TRANSLITERATE };
