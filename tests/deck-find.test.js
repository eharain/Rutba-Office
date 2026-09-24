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

// `find` / `replace` / `replaceAll` are the pane's own API: one entry per
// occurrence, precise enough to walk hit by hit and rewrite exactly one.
test('find lists every occurrence in slide, shape and run order', () => {
  const deck = Deck.open(DECK);
  const hits = deck.find('office');
  assert.deepEqual(hits.map((h) => h.slide), [0, 0, 1]);
  assert.equal(hits[0].text, 'Office');
  assert.deepEqual(hits.map((h) => h.length), [6, 6, 6]);
  // Each hit knows exactly where it sits, not just which shape it is on.
  for (const h of hits) {
    assert.equal(typeof h.paragraph, 'number');
    assert.equal(typeof h.run, 'number');
    assert.equal(typeof h.offset, 'number');
  }
});

test('find is case-insensitive unless asked, and case-sensitive drops what does not match', () => {
  const deck = Deck.open(DECK);
  assert.equal(deck.find('OFFICE').length, 3, 'case-insensitive by default');
  assert.deepEqual(deck.find('Office', { matchCase: true }).map((h) => h.slide), [0]);
  assert.equal(deck.find('office', { matchCase: true }).length, 2, 'the two lower-case ones, on each slide\'s body');
});

test('find answers nothing for a word the deck does not hold', () => {
  const deck = Deck.open(DECK);
  assert.deepEqual(deck.find('nowhere'), []);
  assert.deepEqual(deck.find(''), []);
});

test('replace rewrites one hit and keeps its run\'s look, leaving the rest of the words alone', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTextBox(1, { x: 10, y: 10, w: 300, h: 60, paragraphs: [{ runs: [{ text: 'Bold ', bold: true }, { text: 'Office hours', italic: true, size: 24 }] }] });
  const hit = deck.find('Office').find((h) => String(h.shape) === String(id));
  assert.ok(hit, 'a hit in the new box');
  assert.equal(deck.replace(hit, 'Suite'), true);
  const box = Deck.open(deck.save()).slide(1).shapes.find((s) => String(s.id) === String(id));
  const runs = box.text.paragraphs[0].runs;
  assert.deepEqual(runs.map((r) => r.text), ['Bold ', 'Suite hours']);
  assert.equal(runs[0].bold, true, 'the first run is untouched');
  assert.equal(runs[1].italic, true, 'the edited run keeps its own look');
  assert.equal(runs[1].size, 24);
});

test('replaceAll rewrites every hit and says how many, without one replacement\'s length throwing off the next', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTextBox(0, { x: 10, y: 10, w: 300, h: 60, paragraphs: [{ runs: [{ text: 'office office office' }] }] });
  const count = deck.replaceAll('office', 'suite', { matchCase: false });
  assert.equal(count, 6, 'the title, both bodies, the second slide and three in the new box');
  const reopened = Deck.open(deck.save());
  const box = reopened.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(box.text.paragraphs[0].runs[0].text, 'suite suite suite', 'none of the three hits in one run tripped over the others\' offsets');
  assert.equal(deck.find('office').length, 0);
});

test('a table cell\'s words are found and replaced too, carrying the row and column of the cell', () => {
  const deck = Deck.open(DECK);
  const id = deck.addTable(1, { rows: 2, cols: 2, cells: [['Office', 'Two'], ['Three', 'Office suite']] });
  const hits = deck.find('office');
  const cellHits = hits.filter((h) => String(h.shape) === String(id));
  assert.equal(cellHits.length, 2);
  assert.deepEqual(cellHits.map((h) => [h.row, h.col]).sort(), [[0, 0], [1, 1]]);

  assert.equal(deck.replace(cellHits[0], 'Suite'), true);
  const reopened = Deck.open(deck.save());
  const table = reopened.slide(1).shapes.find((s) => String(s.id) === String(id)).table;
  const words = (r, c) => (table.rows[r].cells[c].text?.paragraphs || []).map((p) => p.runs.map((run) => run.text).join('')).join('');
  assert.equal(words(0, 0), 'Suite');
  assert.equal(words(1, 1), 'Office suite', 'the other cell was not touched');
});

test('find, replace and replaceAll all survive a save', () => {
  const deck = Deck.open(DECK);
  deck.replaceAll('office', 'suite', { matchCase: false });
  const saved = Deck.open(deck.save());
  assert.equal(saved.find('office').length, 0);
  assert.ok(saved.find('suite').length >= 3);
  const hit = saved.find('suite')[0];
  assert.equal(saved.replace(hit, 'Suite'), true);
});
