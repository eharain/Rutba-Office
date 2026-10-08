// Find and Replace across a deck — edges the first tests leave alone: a
// capital that lowers to two characters, a match across two runs in a shape
// and in a table cell, and what a save keeps.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';

const DECK = buildPptx({ title: 'Find', slides: [
  { layout: 'title', title: 'Rutba Office', body: 'A free office suite' },
  { layout: 'obj', title: 'Second slide', body: ['One office', 'Two'] },
] });
const runsOf = (deck, slide, id) => deck.slide(slide).shapes.find((s) => String(s.id) === String(id)).text.paragraphs.map((p) => p.runs.map((r) => r.text));

test('a hit after a dotted capital I is replaced where it sits, not two characters early', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTextBox(0, { x: 10, y: 10, w: 300, h: 60, paragraphs: [{ runs: [{ text: 'İstanbul trip' }] }] });
  const hit = deck.find('stanbul').find((h) => String(h.shape) === String(id));
  assert.equal(hit.offset, 1, 'the offset is into the words as written');
  assert.equal(hit.text, 'stanbul');
  assert.equal(deck.replaceAll('stanbul', 'X'), 1);
  assert.deepEqual(runsOf(Deck.open(deck.save()), 0, id), [['İX trip']]);
});

test('a match across two runs is skipped by find and by replaceAll alike, in a shape and in a table cell, and survives a save untouched', () => {
  const deck = Deck.open(DECK);
  const box = deck.addTextBox(1, { x: 10, y: 10, w: 300, h: 60, paragraphs: [{ runs: [{ text: 'Of', bold: true }, { text: 'fice', italic: true }, { text: ' office' }] }] });
  const table = deck.addTable(1, { rows: 1, cols: 2, cells: [['Of', 'office']] });
  const before = deck.find('office');
  const mine = before.filter((h) => String(h.shape) === String(box) || String(h.shape) === String(table));
  assert.deepEqual(mine.map((h) => [String(h.shape) === String(box) ? 'box' : 'table', h.run ?? null, h.col]), [['box', 2, null], ['table', 0, 1]], 'only the whole-in-one-run ones are listed');
  const replaced = deck.replaceAll('office', 'suite');
  assert.equal(replaced, before.length, 'what find listed is what replaceAll rewrote');
  assert.equal(deck.find('office').length, 0, 'nothing listed is left');

  const again = Deck.open(deck.save());
  assert.deepEqual(runsOf(again, 1, box), [['Of', 'fice', ' suite']]);
  const cells = again.slide(1).shapes.find((s) => String(s.id) === String(table)).table.rows[0].cells;
  assert.deepEqual(cells.map((c) => c.text.paragraphs[0].runs.map((r) => r.text).join('')), ['Of', 'suite']);
  const kept = runsOf(again, 1, box)[0];
  assert.equal(kept.slice(0, 2).join(''), 'Office', 'the straddling word is whole, split as it was');
});

test('a case-sensitive replaceAll leaves the other case alone and counts only what it changed', () => {
  const deck = Deck.open(DECK);
  assert.equal(deck.replaceAll('Office', 'Suite', { matchCase: true }), 1);
  assert.equal(deck.find('office', { matchCase: true }).length, 2);
  assert.equal(deck.find('Suite', { matchCase: true }).length, 1);
});
