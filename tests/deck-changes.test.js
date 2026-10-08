// Review → Show Changes, the engine's half: a deck's fingerprint as it was
// last seen here, and what the file holds now compared with it — slides
// added, removed and moved, and on a slide the shapes added and removed, the
// words changed, a shape moved, or only its formatting.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, deckFingerprint, compareFingerprints } from '@rutba/presentation';

const deck = () => Deck.open(buildPptx({ title: 'Plan', slides: [
  { layout: 'title', title: 'The plan' },
  { layout: 'obj', title: 'Steps', body: ['Plan', 'Build'] },
  { layout: 'obj', title: 'Costs', body: ['Small'] },
] }));

test('a fingerprint names each slide by its id and title, and each shape by its box and words', () => {
  const d = deck();
  const fp = deckFingerprint(d);
  assert.equal(fp.slides.length, 3);
  assert.deepEqual(fp.slides.map((s) => s.title), ['The plan', 'Steps', 'Costs']);
  assert.ok(fp.slides.every((s) => s.id && s.hash && s.shapes.length));
  assert.equal(JSON.stringify(fp).length < 4000, true, 'small enough to keep for every deck');
  assert.deepEqual(compareFingerprints(fp, deckFingerprint(Deck.open(d.save()))), [], 'the same file, saved and opened, has no changes');
});

test('slides added, removed and moved, and on a slide the words, the boxes, the shapes and the rest', () => {
  const d = deck();
  const before = deckFingerprint(d);
  const steps = d.slide(1).shapes.find((s) => (s.text?.paragraphs || []).some((p) => p.runs?.some((r) => r.text === 'Plan')));
  d.setText(1, steps.id, [{ runs: [{ text: 'Plan' }] }, { runs: [{ text: 'Build' }] }, { runs: [{ text: 'Ship' }] }]);
  d.addShape(1, { preset: 'rect', x: 50, y: 50, w: 100, h: 50, name: 'Badge' });
  d.moveSlide(2, 0);
  d.insertSlide(d.slideCount - 1, {});
  const now = deckFingerprint(d);
  const changes = compareFingerprints(before, now);
  const byTitle = Object.fromEntries(changes.map((c) => [c.title, c]));
  assert.equal(byTitle.Steps.kind, 'changed');
  assert.ok(byTitle.Steps.details.some((x) => /^words of .* changed/.test(x)), JSON.stringify(byTitle.Steps.details));
  assert.ok(byTitle.Steps.details.includes('Badge added'));
  assert.equal(byTitle.Costs.kind, 'moved');
  assert.deepEqual(byTitle.Costs.details, ['moved from slide 3']);
  assert.equal(changes.filter((c) => c.kind === 'added').length, 1, 'the new slide');
  assert.equal(changes.find((c) => c.kind === 'added').index, 3);
});

test('a slide removed is listed last, with where it was', () => {
  const d = deck();
  const before = deckFingerprint(d);
  d.removeSlide(1);
  const changes = compareFingerprints(before, deckFingerprint(d));
  assert.deepEqual(changes.map((c) => [c.kind, c.title, c.was]), [['removed', 'Steps', 1]]);
});
