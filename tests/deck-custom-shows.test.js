// Slide Show → Custom Slide Show: named shows of a deck's slides in an order
// of their own, kept where PowerPoint keeps them — the presentation's
// p:custShowLst, each slide by the presentation's relationship to it — read
// back after a save, and a deleted slide taken out of the shows that had it.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';

const deck = () => Deck.open(buildPptx({ slides: [1, 2, 3, 4, 5].map((n) => ({ layout: 'obj', title: `Slide ${n}` })) }));

test('a deck with no custom shows has none; written shows come back after a save, in their own order', () => {
  const d = deck();
  assert.deepEqual(d.customShows(), []);
  d.setCustomShows([{ name: 'Short version', slides: [0, 3, 1] }, { name: 'Board & staff', slides: [4] }]);
  const back = Deck.open(d.save());
  assert.deepEqual(back.customShows(), [{ id: 0, name: 'Short version', slides: [0, 3, 1] }, { id: 1, name: 'Board & staff', slides: [4] }]);
  const pres = back.pkg.text('ppt/presentation.xml');
  assert.match(pres, /<p:custShowLst><p:custShow name="Short version" id="0"><p:sldLst><p:sld r:id="rId\d+"\/><p:sld r:id="rId\d+"\/><p:sld r:id="rId\d+"\/><\/p:sldLst><\/p:custShow><p:custShow name="Board &amp; staff" id="1">/);
  assert.ok(pres.indexOf('<p:custShowLst>') < pres.indexOf('<p:defaultTextStyle'), 'ahead of the default text style, as the schema orders it');
});

test('names must be given and differ; an empty list takes the shows out', () => {
  const d = deck();
  assert.throws(() => d.setCustomShows([{ name: ' ', slides: [0] }]), /needs a name/);
  assert.throws(() => d.setCustomShows([{ name: 'A', slides: [0] }, { name: 'a', slides: [1] }]), /already a custom show/);
  assert.throws(() => d.setCustomShows([{ name: 'A', slides: [9] }]), /no slide/);
  d.setCustomShows([{ name: 'A', slides: [0] }]);
  d.setCustomShows([]);
  assert.doesNotMatch(d.pkg.text('ppt/presentation.xml'), /custShow/);
});

test('a slide deleted goes out of the shows that played it, and the rest keep their slides', () => {
  const d = deck();
  d.setCustomShows([{ name: 'Middle', slides: [1, 2, 3] }]);
  d.removeSlide(2);
  assert.deepEqual(d.customShows()[0].slides, [1, 2], 'the old fourth slide is the third now');
  assert.equal(d.customShows()[0].name, 'Middle');
});
