// Bullets that live in symbol fonts.
//
// Word and PowerPoint write a bullet as a character in Wingdings or Symbol —
// U+F0A7 in Wingdings is a small square, U+F0B7 in Symbol is a round bullet —
// and a renderer without those fonts shows the private-use code point as a
// box, or the ASCII letter the glyph sits on ("§", "v", "ü"). This is the map
// from the code point in each font to the Unicode character that looks the
// same in any font. Anything not listed comes back as a round bullet, which
// is what the character almost always means.

const WINGDINGS = {
  0x6c: '●', 0x6d: '❍', 0x6e: '■', 0x6f: '□', 0x70: '◻', 0x71: '❑', 0x72: '❒', 0x73: '◆', 0x74: '❖',
  0x75: '◆', 0x76: '❖', 0x77: '◆', 0x9f: '•', 0xa7: '▪', 0xa8: '□', 0xb7: '●', 0xd8: '➢', 0xdc: '➔',
  0xe0: '⇨', 0xf0: '⇨', 0xfc: '✓', 0xfe: '☒', 0x21: '✏', 0x22: '✂', 0x28: '☎', 0x2a: '✉',
  0x3c: '⌧', 0x4a: '☺', 0x4c: '☹', 0x51: '✈', 0x52: '☼',
};
const SYMBOL = { 0xb7: '•', 0xa7: '♣', 0xa8: '♦', 0xa9: '♥', 0xaa: '♠', 0xd8: '¬', 0xde: '⇒', 0xdb: '⇔', 0xe0: '◊' };
const WEBDINGS = { 0x3d: '●', 0x67: '■', 0x69: '◆', 0x6e: '■', 0x72: '☺' };

const TABLES = { wingdings: WINGDINGS, 'wingdings 2': WINGDINGS, 'wingdings 3': WINGDINGS, symbol: SYMBOL, webdings: WEBDINGS };

/**
 * The Unicode character to draw for a bullet stored in a symbol font.
 *
 * @param {string} char the stored character — private-use (U+F0xx) or plain
 * @param {string} [font] the font it was stored for
 * @returns {string} something every font can draw
 */
export function bulletGlyph(char, font) {
  if (!char) return '•';
  const cp = char.codePointAt(0);
  const code = cp >= 0xf000 && cp <= 0xf0ff ? cp - 0xf000 : cp;
  const table = TABLES[String(font || '').trim().toLowerCase()];
  if (table) return table[code] || '•';
  // No symbol font named: a private-use code point is still a symbol-font
  // bullet (Wingdings is the one that gets written without a font), a
  // plain character is itself.
  if (cp >= 0xf000 && cp <= 0xf0ff) return code === 0xb7 ? '•' : WINGDINGS[code] || SYMBOL[code] || '•';
  return char;
}
