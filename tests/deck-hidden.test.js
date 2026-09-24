// Hide Slide — the engine's half: PowerPoint marks a slide left out of the
// show as show="0" on the slide part's own root tag, and writes no such
// attribute at all once the slide is shown again. Everything else about the
// slide's XML, and every other slide's, is untouched.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';

const DECK = buildPptx({
  title: 'Hidden',
  slides: [
    { layout: 'title', title: 'Opening', body: 'A deck with a hidden slide' },
    { layout: 'obj', title: 'The hidden one', body: ['Skip me'] },
    { layout: 'obj', title: 'Closing', body: ['Thanks'] },
  ],
});

test('a slide is hidden, shown again, and only its root tag changes', () => {
  const deck = Deck.open(DECK);
  const part0 = deck.slideParts[0].part;
  const part1 = deck.slideParts[1].part;
  const part2 = deck.slideParts[2].part;
  const before0 = deck.pkg.text(part0);
  const before1 = deck.pkg.text(part1);
  const before2 = deck.pkg.text(part2);

  assert.equal(deck.isSlideHidden(1), false);
  assert.equal(deck.setSlideHidden(1, true), true, 'a real change reports true');
  assert.ok(deck.dirty);

  assert.equal(deck.isSlideHidden(1), true);
  assert.deepEqual(deck.outline().map((o) => o.hidden), [false, true, false]);
  assert.equal(deck.slide(0).hidden, false);
  assert.equal(deck.slide(1).hidden, true);
  assert.equal(deck.slide(2).hidden, false);

  // Only the root tag changed: show="0" added, nothing else in the XML moved.
  const after1 = deck.pkg.text(part1);
  assert.match(after1, /<p:sld\b[^>]*\sshow="0"/);
  const withoutShow = after1.replace(/<p:sld\b[^>]*>/, (m) => m.replace(/\s+show="0"/, ''));
  assert.equal(withoutShow, before1, 'nothing but the show attribute changed');

  // The other slides are untouched, byte for byte.
  assert.equal(deck.pkg.text(part0), before0);
  assert.equal(deck.pkg.text(part2), before2);

  // Survives a save and reopen.
  const reopened = Deck.open(deck.save());
  assert.equal(reopened.isSlideHidden(0), false);
  assert.equal(reopened.isSlideHidden(1), true);
  assert.equal(reopened.isSlideHidden(2), false);

  // Hiding an already-hidden slide is not a change.
  assert.equal(deck.setSlideHidden(1, true), false);

  // Un-hiding takes the attribute off entirely — PowerPoint writes none, not show="1".
  assert.equal(deck.setSlideHidden(1, false), true);
  assert.equal(deck.isSlideHidden(1), false);
  const after1shown = deck.pkg.text(part1);
  assert.doesNotMatch(after1shown, /\sshow=/);
  assert.equal(after1shown, before1, 'back to exactly what it was');
  assert.equal(deck.setSlideHidden(1, false), false, 'already shown is not a change');
});

test('a slide whose root already carries show="1" is hidden as show="0", not left at show="1"', () => {
  const deck = Deck.open(DECK);
  const part = deck.slideParts[1].part;
  const xml = deck.pkg.text(part);
  deck.pkg.write_(part, Buffer.from(xml.replace(/<p:sld\b/, '<p:sld show="1"'), 'utf8'));

  assert.equal(deck.setSlideHidden(1, true), true);
  const after = deck.pkg.text(part);
  assert.match(after, /<p:sld\b[^>]*\sshow="0"/);
  assert.doesNotMatch(after, /show="1"/);
});

test('duplicating a hidden slide duplicates the hiding too', () => {
  const deck = Deck.open(DECK);
  deck.setSlideHidden(1, true);
  const newIndex = deck.duplicateSlide(1);
  assert.equal(deck.isSlideHidden(newIndex), true);
  assert.equal(deck.outline()[newIndex].hidden, true);
});

test('a bad index throws', () => {
  const deck = Deck.open(DECK);
  assert.throws(() => deck.isSlideHidden(9), /no slide/);
  assert.throws(() => deck.setSlideHidden(9, true), /no slide/);
});
