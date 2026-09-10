// The Layers and Designs panes — the deck engine's half.
//
// The drawing order is the spTree's order; hiding is cNvPr's hidden flag; a
// name is cNvPr's name; the deck's layouts are its slideLayout parts, each
// with the boxes its placeholders make; a slide moves to another layout by
// its relationship. All of it survives a save and a reopen.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';

const DECK = buildPptx({
  title: 'Layers',
  slides: [
    { layout: 'title', title: 'A deck', body: 'with two slides' },
    { layout: 'obj', title: 'Second', body: ['One', 'Two'] },
  ],
});

const ids = (deck, i = 0) => deck.slide(i).shapes.map((s) => s.id);

test('reorderShape moves a shape among its siblings and nothing else', () => {
  const deck = Deck.open(DECK);
  // A third shape, so front / backward / back each move something.
  deck.addTextBox(0, { x: 40, y: 40, w: 200, h: 60, paragraphs: [{ runs: [{ text: 'Note' }] }] });
  const before = ids(deck);
  assert.equal(before.length, 3, 'three shapes on the title slide');
  const [bottom, top] = [before[0], before[before.length - 1]];

  assert.equal(deck.reorderShape(0, bottom, 'front'), true);
  assert.deepEqual(ids(deck).slice(-1), [bottom], 'the bottom shape is now on top');
  assert.equal(ids(deck).length, before.length, 'nothing lost');

  assert.equal(deck.reorderShape(0, bottom, 'backward'), true);
  assert.equal(ids(deck)[before.length - 2], bottom, 'one step back');
  assert.equal(deck.reorderShape(0, bottom, 'back'), true);
  assert.deepEqual(ids(deck), before, 'and back where it started');
  assert.equal(deck.reorderShape(0, bottom, 'back'), false, 'already at the back: nothing to do');
  assert.equal(deck.reorderShape(0, top, 'forward'), false, 'already at the front');
  assert.throws(() => deck.reorderShape(0, '999', 'front'), /not found/);

  const reopened = Deck.open(deck.save());
  assert.deepEqual(ids(reopened), before, 'the order is in the file');
});

test('setShapeHidden and renameShape write cNvPr, and the scene reads them back', () => {
  const deck = Deck.open(DECK);
  const [id] = ids(deck);
  assert.equal(deck.slide(0).shapes[0].hidden, false);
  assert.equal(deck.setShapeHidden(0, id, true), true);
  assert.equal(deck.slide(0).shapes[0].hidden, true, 'hidden');
  assert.equal(deck.setShapeHidden(0, id, false), true);
  assert.equal(deck.slide(0).shapes[0].hidden, false, 'shown again');
  assert.equal(deck.setShapeHidden(0, id, false), false, 'already shown: nothing to write');

  assert.equal(deck.renameShape(0, id, 'Hero & title'), true);
  const reopened = Deck.open(deck.save());
  assert.equal(reopened.slide(0).shapes[0].name, 'Hero & title', 'the name, unescaped, after a reopen');
});

test('layouts lists the deck’s layouts with their placeholders, and applyLayout moves a slide', () => {
  const deck = Deck.open(DECK);
  const layouts = deck.layoutList();
  assert.deepEqual(layouts.map((l) => l.name), ['Title Slide', 'Title and Content']);
  assert.deepEqual(layouts.map((l) => l.type), ['title', 'obj']);
  for (const l of layouts) {
    assert.ok(l.placeholders.length >= 2, `${l.name} has placeholders`);
    for (const p of l.placeholders) assert.ok(p.geometry.w > 0 && p.geometry.h > 0, `${l.name}: ${p.type} has a box`);
  }

  assert.equal(deck.layoutOf(1), layouts[1].part, 'the second slide is on Title and Content');
  assert.equal(deck.applyLayout(1, layouts[0].part), true);
  assert.equal(deck.layoutOf(1), layouts[0].part, 'and now on Title Slide');
  assert.equal(deck.slide(1).layout, layouts[0].part, 'the scene says so too — its cache is keyed on XML this did not touch');
  assert.throws(() => deck.applyLayout(1, 'ppt/slideLayouts/slideLayout9.xml'), /no layout/);

  const reopened = Deck.open(deck.save());
  assert.equal(reopened.layoutOf(1), layouts[0].part, 'in the file');
  assert.ok(reopened.slide(1).shapes.length >= 1, 'the slide still draws');
});

test('insertSlide can name the layout it wants', () => {
  const deck = Deck.open(DECK);
  const layouts = deck.layoutList();
  const at = deck.insertSlide(0, { layout: 'title', layoutPart: layouts[0].part, title: 'Inserted' });
  assert.equal(at, 1);
  assert.equal(deck.layoutOf(1), layouts[0].part, 'on the named layout, not the neighbour’s');
  const reopened = Deck.open(deck.save());
  assert.equal(reopened.slideCount, 3);
  assert.equal(reopened.layoutOf(1), layouts[0].part);
});
