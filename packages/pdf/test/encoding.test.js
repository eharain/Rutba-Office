'use strict';

const test = require('node:test');
const assert = require('node:assert');
const { encode, literal, pdfString } = require('../src/encoding');

const bytes = (s) => encode(s);

test('ASCII is itself', () => {
  assert.deepStrictEqual(bytes('AZaz09'), [65, 90, 97, 122, 48, 57]);
});

test('Latin-1 accents survive, because customer names contain them', () => {
  assert.deepStrictEqual(bytes('Müller'), [77, 0xFC, 108, 108, 101, 114]);
  assert.deepStrictEqual(bytes('Ståle Ångström'), [83, 116, 0xE5, 108, 101, 32, 0xC5, 110, 103, 115, 116, 114, 0xF6, 109]);
});

test('the 0x80 block - the characters Word actually produces - is mapped', () => {
  // Curly quotes and an em dash are what a pasted description arrives as.
  assert.deepStrictEqual(bytes('‘a’ “b” – — …'),
    [0x91, 97, 0x92, 32, 0x93, 98, 0x94, 32, 0x96, 32, 0x97, 32, 0x85]);
  assert.deepStrictEqual(bytes('€10'), [0x80, 49, 48]);
});

test('currency symbols a finance document needs', () => {
  assert.deepStrictEqual(bytes('£$€¥¢'), [0xA3, 0x24, 0x80, 0xA5, 0xA2]);
});

test('a character with no glyph becomes a visible ?, never nothing', () => {
  // Dropping it would turn a name into a different, plausible name.
  assert.deepStrictEqual(bytes('张'), [0x3F]);
  assert.deepStrictEqual(bytes('Ivan Ы'), [73, 118, 97, 110, 32, 0x3F]);
  assert.strictEqual(bytes('日本語').length, 3);
});

test('characters with an honest ASCII equivalent are spelled, not questioned', () => {
  assert.deepStrictEqual(bytes('a−b'), [97, 45, 98], 'minus sign becomes a hyphen');
  assert.deepStrictEqual(bytes('a b'), [97, 32, 98], 'thin space becomes a space');
  assert.deepStrictEqual(bytes('№ 7'), [78, 111, 46, 32, 55], 'numero becomes No.');
  assert.deepStrictEqual(bytes('soft­hyphen').length, 10, 'a soft hyphen prints as nothing');

  // A hard space is NOT one of these: WinAnsi has it, and it is how an
  // author keeps two words on one line when we own the line breaking.
  assert.deepStrictEqual(bytes('a b'), [97, 0xA0, 98]);
});

test('control characters cannot reach the page as controls', () => {
  assert.deepStrictEqual(bytes('a\nb'), [97, 32, 98]);
  assert.deepStrictEqual(bytes('a\tb'), [97, 32, 32, 32, 32, 98]);
  // A control byte typed as an ESCAPE, not as itself: a literal one is
  // invisible in a diff, which is what verify-encoding exists to catch.
  for (const b of bytes('a\u0007b\u001bc')) assert.ok(b >= 32, 'no byte below space');
});

test('parens in a company name do not end the string early', () => {
  // "Acme (UK) Ltd" unescaped closes the PDF string and the rest of the
  // page becomes syntax. This is the bug that makes a file unopenable.
  const s = pdfString('Acme (UK) Ltd \\ Co.');
  assert.strictEqual(s, '(Acme \\(UK\\) Ltd \\\\ Co.)');
  const opens = (s.match(/(?<!\\)\(/g) || []).length;
  const closes = (s.match(/(?<!\\)\)/g) || []).length;
  assert.strictEqual(opens, 1);
  assert.strictEqual(closes, 1);
});

test('high bytes are escaped octally, so the file stays 7-bit', () => {
  const s = literal(encode('Müller'));
  assert.strictEqual(s, '(M\\374ller)');
  for (const ch of s) assert.ok(ch.charCodeAt(0) < 128, 'nothing above ASCII survives into the file');
});

test('nothing is not a crash', () => {
  assert.deepStrictEqual(bytes(null), []);
  assert.deepStrictEqual(bytes(undefined), []);
  assert.strictEqual(pdfString(''), '()');
  assert.deepStrictEqual(bytes(0), [48]);
});

// ── which markets this engine can actually invoice ────────────────────────
//
// The market survey in docs/todo/books-program/MARKETS.md rests on a claim
// about this file: that Latin-script languages draw correctly and Arabic
// and Hebrew do not. That is a MARKET boundary, not a typography detail —
// it is what puts Saudi Arabia and Egypt behind a document-engine problem
// rather than an XML one, and what keeps South America and anglophone
// Africa reachable. So it is asserted here rather than left to be
// rediscovered.

/** Characters lost to substitution: '?' bytes the source did not contain. */
function lost(text) {
  return [...encode(text)].filter((b) => b === 0x3f).length
       - [...text].filter((c) => c === '?').length;
}

test('the Latin-script languages of the surveyed markets draw intact', () => {
  const drawable = {
    spanish: 'Factura electrónica — señor, año, jalapeño',
    portuguese: 'Nota Fiscal Eletrônica — informações, São Paulo, ação',
    afrikaans: 'Belastingfaktuur — dié, sê, oë',
    isizulu: 'Ikhasi lentela — isibalo esiphelele',
    french: 'Facture électronique — société à responsabilité limitée',
    swahili: 'Ankara ya kodi — malipo kamili',
  };
  for (const [language, text] of Object.entries(drawable)) {
    assert.strictEqual(lost(text), 0, `${language} lost characters to substitution`);
  }
});

test('Arabic and Hebrew substitute rather than draw, visibly', () => {
  // Not a defect — a documented boundary. The point of asserting it is that
  // the failure is a VISIBLE '?' rather than a silently shortened string: a
  // dropped character reads as correct and is not.
  for (const text of ['فاتورة ضريبية', 'חשבונית מס']) {
    assert.ok(lost(text) > 0, 'a script with no glyph must substitute');
    assert.strictEqual(encode(text).length, [...text].length, 'one substitute per character, nothing dropped');
  }
});
