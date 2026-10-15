// Redact and Inspect for a deck: the marked words taken out of the slides,
// the speaker notes, the comments and their authors, the alt text and the
// deck's properties, the package read back with none of them in it; and the
// inspector finding and removing comments, notes, personal details,
// properties, invisible and off-slide objects, as PowerPoint's does.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml';
import { redactPackage, partsHolding } from '@rutba/ooxml/redact';
import { inspectPackage, removeFromPackage } from '@rutba/ooxml/inspect';

/** A deck that names Ayesha Khan on a slide, in its notes, a comment, alt text and its properties; one hidden box and one off the slide. */
export function secretDeck() {
  const d = Deck.open(buildPptx({ title: 'Offer for Ayesha Khan', author: 'Bilal Ahmed', slides: [
    { layout: 'title', title: 'Offer', body: 'For Ayesha Khan, ayesha.khan@example.pk' },
    { layout: 'obj', title: 'Terms', body: ['Call +92 300 1234567'] },
  ] }));
  d.setNotes(0, 'Remember: Ayesha Khan prefers mornings.');
  d.addComment(1, { text: 'Check with Ayesha Khan', author: 'Sana Iqbal' });
  const box = d.addTextBox(1, { x: 50, y: 50, w: 200, h: 40, paragraphs: [{ runs: [{ text: 'Draft only' }] }], name: 'Hidden note' });
  d.setShapeHidden(1, box, true);
  d.addTextBox(1, { x: 5000, y: 50, w: 200, h: 40, paragraphs: [{ runs: [{ text: 'Parked outside' }] }], name: 'Parked' });
  const pic = d.slide(0).shapes[0];
  const pkg = OoxmlPackage.read(d.save());
  const s1 = 'ppt/slides/slide1.xml';
  pkg.write_(s1, pkg.text(s1).replace(/(<p:cNvPr\b[^>]*id="\d+"[^>]*name="[^"]*")/, '$1 descr="Photo of Ayesha Khan"'));
  return { bytes: pkg.write(), pic };
}

test('Redact takes the words out of slides, notes, comments, alt text and properties, and the deck still opens', () => {
  const pkg = OoxmlPackage.read(secretDeck().bytes);
  const scan = redactPackage(pkg, { marks: [{ text: 'Ayesha Khan' }, { pattern: 'email' }, { pattern: 'phone' }], apply: false });
  const kinds = new Set(scan.parts.map((p) => p.kind));
  for (const k of ['slides', 'notes', 'comments', 'properties']) assert.ok(kinds.has(k), 'found in ' + k + ': ' + [...kinds]);
  redactPackage(pkg, { marks: [{ text: 'Ayesha Khan' }, { pattern: 'email' }, { pattern: 'phone' }] });
  const out = Buffer.from(pkg.write());
  assert.deepEqual(partsHolding(out, ['Ayesha Khan', 'ayesha.khan@example.pk', '300 1234567']), []);
  const deck = Deck.open(out);
  assert.match(JSON.stringify(deck.slide(0).shapes.map((s) => s.text)), /For ███████████, ██████████████████████/);
  assert.equal(deck.slideCount, 2);
});

test('Inspect Presentation finds and removes what PowerPoint\'s inspector does', () => {
  const pkg = OoxmlPackage.read(secretDeck().bytes);
  const found = Object.fromEntries(inspectPackage(pkg, 'presentation').map((k) => [k.id, k]));
  assert.equal(found.comments.found, 1);
  assert.equal(found.notes.found, 1);
  assert.equal(found.invisible.items.length, 1);
  assert.match(found.invisible.items[0], /^Hidden note/);
  assert.equal(found.offSlide.items.length, 1);
  assert.match(found.offSlide.items[0], /^Parked/);
  assert.ok(found.personal.items.includes('Bilal Ahmed') && found.personal.items.includes('Sana Iqbal'));
  removeFromPackage(pkg, 'presentation', ['comments', 'notes', 'invisible', 'offSlide', 'personal', 'properties']);
  const after = Object.fromEntries(inspectPackage(pkg, 'presentation').map((k) => [k.id, k.found]));
  for (const id of ['comments', 'notes', 'invisible', 'offSlide', 'personal']) assert.equal(after[id], 0, id);
  const deck = Deck.open(pkg.write());
  assert.equal(deck.comments().length, 0);
  assert.ok(!deck.slide(1).shapes.some((s) => /^(?:Hidden note|Parked)/.test(s.name || '')));
  assert.equal(deck.slideCount, 2);
});
