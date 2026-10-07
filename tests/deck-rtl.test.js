// Presentations: right-to-left paragraphs — an Arabic or Hebrew paragraph's
// `<a:pPr rtl="1">` read and written back, and drawn from the right: the
// line's start at the right margin (less its indent) and the words running
// leftwards from there, the bullet at the right of them, the alignment the
// side it names.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const BOX = { x: 100, y: 100, w: 400, h: 120 };
const deckWith = (paragraphs) => Deck.open(buildPptx({ slides: [{ layout: 'blank', textBoxes: [{ ...BOX, paragraphs }] }] }));
const box = (deck) => deck.slide(0).shapes.find((s) => s.text);
/** The <text> elements that carry words, with their x, direction and words. */
const texts = (svg) => [...svg.matchAll(/<text\b([^>]*)>([\s\S]*?)<\/text>/g)].map((m) => ({
  x: Number(/\bx="([\d.]+)"/.exec(m[1])?.[1]),
  rtl: /\bdirection="rtl"/.test(m[1]),
  words: m[2].replace(/<[^>]+>/g, ''),
}));

test('a paragraph right to left is read from rtl="1" and written back with it', () => {
  const deck = deckWith([{ rtl: true, align: 'right', runs: [{ text: 'مرحبا بالعالم' }] }, { runs: [{ text: 'Left to right' }] }]);
  const [first, second] = box(deck).text.paragraphs;
  assert.equal(first.rtl, true);
  assert.equal(first.align, 'right');
  assert.equal(second.rtl, undefined);
  // A round trip through setText, as an edit makes one, keeps it.
  deck.setText(0, box(deck).id, box(deck).text.paragraphs);
  const xml = OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part);
  assert.match(xml, /<a:pPr algn="r" rtl="1"\/>/);
  assert.equal((xml.match(/rtl="1"/g) || []).length, 1, 'only the Arabic paragraph');
  const turnedBack = box(deck).text.paragraphs.map((p, i) => (i === 0 ? { ...p, rtl: false } : p));
  deck.setText(0, box(deck).id, turnedBack);
  assert.match(OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part), /rtl="0"/, 'stated left to right');
});

test('right to left a paragraph is drawn from the right edge, leftwards; left to right from the left', () => {
  const deck = deckWith([{ rtl: true, align: 'right', runs: [{ text: 'مرحبا' }] }, { runs: [{ text: 'Hello' }] }]);
  const svg = renderSlide(deck.slide(0), { width: 960 });
  const words = texts(svg).filter((t) => t.words === 'مرحبا' || t.words === 'Hello');
  const arabic = words.find((t) => t.words === 'مرحبا');
  const latin = words.find((t) => t.words === 'Hello');
  assert.ok(arabic && latin, 'both drawn');
  assert.equal(arabic.rtl, true, 'drawn right to left');
  assert.equal(latin.rtl, false);
  // Drawn in the slide's own pixels; the box's right inset is 7.2 px.
  const right = BOX.x + BOX.w - 7.2;
  assert.ok(Math.abs(arabic.x - right) < 2, `the Arabic line starts at the box's right margin (${arabic.x} by ${right})`);
  assert.ok(latin.x < BOX.x + 20, 'the English one at its left');
});

test('right to left a bullet stands at the right of its words, and a left-aligned line ends at the left margin', () => {
  const deck = deckWith([
    { rtl: true, align: 'right', runs: [{ text: 'البند الأول' }] },
    { rtl: true, align: 'left', runs: [{ text: 'يسار' }] },
  ]);
  const shape = box(deck);
  const paragraphs = shape.text.paragraphs.map((p, i) => (i === 0 ? { ...p, bullet: { type: 'char', char: '•' } } : p));
  deck.setText(0, shape.id, paragraphs);
  const svg = renderSlide(deck.slide(0), { width: 960 });
  const all = texts(svg);
  const bullet = all.find((t) => t.words === '•');
  const first = all.find((t) => t.words === 'البند الأول');
  const left = all.find((t) => t.words === 'يسار');
  assert.ok(bullet && first && left);
  assert.ok(bullet.x > first.x, `the bullet (${bullet.x}) at the right of its words (${first.x})`);
  assert.ok(left.x < BOX.x + BOX.w / 2, 'aligned left, its start — the right end of its words — is near the left edge');
});
