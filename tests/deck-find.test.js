// Find and replace — the deck engine's half: every shape the words are on,
// across the deck, and the words replaced run by run so each keeps its look,
// a field's text left alone.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';

const DECK = buildPptx({ title: 'Find', slides: [
  { layout: 'title', title: 'Rutba Office', body: 'A free office suite' },
  { layout: 'obj', title: 'Second slide', body: ['One office', 'Two'] },
] });
const words = (s) => (s?.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join('')).join(' ');

test('findText lists every shape the words are on, with counts, case-insensitively unless asked', () => {
  const deck = Deck.open(DECK);
  const hits = deck.findText('office');
  assert.deepEqual(hits.map((h) => [h.slide, h.count]), [[0, 1], [0, 1], [1, 1]]);
  assert.equal(hits[0].text, 'Rutba Office');
  assert.deepEqual(deck.findText('Office', { matchCase: true }).map((h) => [h.slide, h.count]), [[0, 1]]);
  assert.deepEqual(deck.findText(''), []);
  assert.deepEqual(deck.findText('nowhere'), []);
});

test('replaceText rewrites the words on every shape, keeps each run\'s look, leaves a field alone, and says how many', () => {
  const deck = Deck.open(DECK);
  deck.setFooter(0, { slideNumber: true, footer: 'One office' });
  assert.equal(deck.replaceText('office', 'suite'), 4, 'title, body, the second slide and the footer');
  assert.deepEqual(deck.slide(0).shapes.map(words).filter((t) => /suite/.test(t)).sort(), ['A free suite suite', 'One suite', 'Rutba suite'].sort());
  assert.equal(words(deck.slide(1).shapes.find((s) => /One/.test(words(s)))), 'One suite Two');
  assert.equal(deck.findText('office').length, 0);

  // A run keeps its look, and the slide number field its text.
  const id = deck.addTextBox(1, { x: 10, y: 10, w: 300, h: 60, paragraphs: [{ runs: [{ text: 'Bold ', bold: true }, { text: 'One', italic: true, size: 24 }] }] });
  assert.equal(deck.replaceText('one', '1'), 3, 'the footer too, and the body\'s One and the box\'s, not the number 1');
  const box = deck.slide(1).shapes.find((s) => String(s.id) === String(id));
  assert.deepEqual(box.text.paragraphs[0].runs.map((r) => [r.text, r.bold ?? null, r.italic ?? null]), [['Bold ', true, null], ['1', null, true]]);
  const num = deck.slide(0).shapes.find((s) => s.placeholder?.type === 'sldNum');
  assert.equal(words(num), '1', 'the field says the slide\'s place still');
  assert.equal(deck.replaceText('1', 'x'), 3, 'the box, the body and the footer: the field is not text to replace');
  assert.equal(words(num), '1');
  assert.equal(Deck.open(deck.save()).findText('suite').length, 4, 'through a save');
});
