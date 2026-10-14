// Review → Compare in Presentations: a reviewed copy merged in as changes,
// words put in and taken out, a shape added, deleted, moved, resized and
// restyled, a slide added, deleted and moved, each accepted or rejected,
// one at a time, a slide's at once or all, whether accepted read from the
// deck itself so Undo leaves the list true, and the deck saved with what
// was accepted and opened again as PowerPoint would open it.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { compareDecks, beginReview, reviewState, decideChanges, wordDiff } from '@rutba/presentation/compare';

const original = () => buildPptx({ title: 'Plan', slides: [
  { layout: 'title', title: 'Plan for 2027', body: 'Draft' },
  { layout: 'obj', title: 'Costs', body: ['Rent 100', 'Staff 200'] },
  { layout: 'obj', title: 'Risks', body: ['Late supplies'] },
  { layout: 'obj', title: 'Extra', body: ['To be cut'] },
] });

/** The reviewer's copy: words changed, a title moved and resized, a box added, the last slide deleted, the third moved first. */
function reviewedCopy(bytes) {
  const d = Deck.open(bytes);
  const shapes = (i) => d.slide(i).shapes;
  const body = shapes(1).find((s) => s.placeholder?.type !== 'title' && s.text);
  d.setText(1, body.id, [{ runs: [{ text: 'Rent 100' }] }, { runs: [{ text: 'Staff 250 from March' }] }]);
  const title = shapes(0).find((s) => s.placeholder?.type === 'ctrTitle' || s.placeholder?.type === 'title');
  const g = title.geometry;
  d.setGeometry(0, title.id, { x: g.x + 40, y: g.y + 20, w: g.w - 100, h: g.h });
  d.addTextBox(2, { x: 100, y: 400, w: 300, h: 60, paragraphs: [{ runs: [{ text: 'Ask the board' }] }], name: 'Reviewer note' });
  d.removeSlide(3);
  d.moveSlide(2, 0);
  return Deck.open(d.save());
}

test('what was put in and taken out, word by word', () => {
  assert.deepEqual(wordDiff('Staff 200', 'Staff 250 from March'), { inserted: ['250 from March'], deleted: ['200'] });
  assert.deepEqual(wordDiff('a b c', 'a c'), { inserted: [], deleted: ['b'] });
});

test('Compare finds each kind of change PowerPoint lists', () => {
  const bytes = original();
  const changes = compareDecks(Deck.open(bytes), reviewedCopy(bytes));
  const kinds = changes.map((c) => c.kind).sort();
  assert.deepEqual([...new Set(kinds)].sort(), ['moved', 'resized', 'shapeAdded', 'slideDeleted', 'slideMoved', 'text'].sort());
  const text = changes.find((c) => c.kind === 'text');
  assert.deepEqual(text.detail.inserted, ['250 from March']);
  assert.deepEqual(text.detail.deleted, ['200']);
  assert.match(changes.find((c) => c.kind === 'shapeAdded').name, /^Reviewer note/);
});

test('each change is accepted and rejected on its own, a slide\'s at once and all at once, and Undo leaves the list true', () => {
  const bytes = original();
  const deck = Deck.open(bytes);
  const review = beginReview(deck, reviewedCopy(bytes), { DeckClass: Deck });
  const id = (kind) => review.changes.find((c) => c.kind === kind).id;
  let state = reviewState(deck, review);
  assert.ok(state.every((c) => !c.accepted), 'nothing is accepted to begin with');

  state = decideChanges(deck, review, [id('text')], true);
  assert.equal(state.find((c) => c.kind === 'text').accepted, true);
  const costs = deck.slide(1).shapes.find((s) => s.text && /Staff/.test(JSON.stringify(s.text)));
  assert.match(JSON.stringify(costs.text), /Staff 250 from March/);

  state = decideChanges(deck, review, [id('text')], false);
  assert.equal(state.find((c) => c.kind === 'text').accepted, false);
  assert.match(JSON.stringify(deck.slide(1).shapes.map((s) => s.text)), /Staff 200/, 'rejected: the words as they were');

  // The first slide's changes at once: its title moved and resized.
  const first = review.changes.filter((c) => c.scope === 'shape' && c.slideId === String(deck.slideParts[0].id)).map((c) => c.id);
  assert.equal(first.length, 2);
  state = decideChanges(deck, review, first, true);
  assert.ok(state.filter((c) => first.includes(c.id)).every((c) => c.accepted));
  deck.undo();
  assert.ok(reviewState(deck, review).filter((c) => first.includes(c.id)).every((c) => !c.accepted), 'Undo puts them back, and the list says so');
  deck.redo();
  assert.ok(reviewState(deck, review).filter((c) => first.includes(c.id)).every((c) => c.accepted));

  // All at once.
  state = decideChanges(deck, review, review.changes.map((c) => c.id), true);
  assert.ok(state.every((c) => c.accepted), JSON.stringify(state.filter((c) => !c.accepted).map((c) => c.kind)));
  assert.equal(deck.slideCount, 3, 'the deleted slide is gone');
  const titles = deck.slideParts.map((_, i) => deck.slide(i).shapes.find((s) => /title/i.test(s.placeholder?.type || ''))).map((s) => JSON.stringify(s?.text || ''));
  assert.match(titles[0], /Risks/, 'the moved slide is first');
  assert.ok(deck.slide(0).shapes.some((s) => /^Reviewer note/.test(s.name || '')), 'the reviewer\'s box is on the moved slide');

  // Saved and opened again: the accepted deck, as any reader sees it.
  const again = Deck.open(deck.save());
  assert.equal(again.slideCount, 3);
  assert.ok(again.slide(0).shapes.some((s) => /^Reviewer note/.test(s.name || '')));
  assert.match(JSON.stringify(again.slide(2).shapes.map((s) => s.text)), /Staff 250 from March/);
});

test('rejecting everything leaves the deck as it was', () => {
  const bytes = original();
  const deck = Deck.open(bytes);
  const before = deck.slideParts.map((s) => deck.pkg.text(s.part));
  const review = beginReview(deck, reviewedCopy(bytes), { DeckClass: Deck });
  decideChanges(deck, review, review.changes.map((c) => c.id), true);
  decideChanges(deck, review, review.changes.map((c) => c.id), false);
  assert.equal(deck.slideCount, 4);
  assert.deepEqual(deck.slideParts.map((s) => deck.pkg.text(s.part)), before);
});
