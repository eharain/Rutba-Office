/**
 * Presentations: New Slide on one of the deck's own layouts. The new slide
 * holds that layout's placeholders, each by the layout's own index, so each
 * sits where the layout puts it: a Section Header whose text is placeholder
 * 13, a Title Only with no body at all, Two Content with both columns. Date,
 * footer, slide number and picture placeholders are left to the layout.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Deck } from '@rutba/presentation';

const here = path.dirname(fileURLToPath(import.meta.url));
const PITCH = fs.readFileSync(path.join(here, '..', 'packages', 'templates', 'files', 'deck-pitch.pptx'));
const SHOWCASE = fs.readFileSync(path.join(here, 'fixtures', 'rich', 'showcase.pptx'));

const placeholders = (deck, i) => deck.slide(i).shapes.filter((s) => s.placeholder).map((s) => ({
  type: s.placeholder.type || null, idx: s.placeholder.idx ?? null, text: (s.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join('')).join('|'), geometry: s.geometry,
}));

test('New Slide on a layout gives the slide that layout\'s placeholders, each where the layout puts it', () => {
  const deck = Deck.open(PITCH);
  const layouts = deck.layoutList();
  const section = layouts.find((l) => l.name === 'Section Header');
  const at = deck.insertSlide(0, { layout: 'obj', layoutPart: section.part, fromLayout: true, title: 'New slide', body: ['Point one'] });
  const got = placeholders(deck, at);
  assert.deepEqual(got.map((p) => [p.type, String(p.idx ?? ''), p.text]), [['title', '', 'New slide'], ['body', '13', 'Point one']]);
  const laid = section.placeholders.find((p) => p.type === 'body').geometry;
  assert.deepEqual(got[1].geometry && [got[1].geometry.x, got[1].geometry.y, got[1].geometry.w], [laid.x, laid.y, laid.w], 'the body where the layout has it');

  const only = layouts.find((l) => l.name === 'Title Only');
  const t = deck.insertSlide(at, { layout: 'obj', layoutPart: only.part, fromLayout: true, title: 'New slide', body: ['Point one'] });
  assert.deepEqual(placeholders(deck, t).map((p) => p.type), ['title'], 'no body a Title Only layout does not have');

  const caption = layouts.find((l) => l.name === 'Picture with Caption');
  const c = deck.insertSlide(t, { layout: 'obj', layoutPart: caption.part, fromLayout: true, title: 'New slide', body: ['Point one'] });
  assert.deepEqual(placeholders(deck, c).map((p) => [p.type, String(p.idx ?? '')]), [['title', ''], ['body', '14']], 'the picture placeholder left to the layout');

  const reopened = Deck.open(deck.save());
  assert.equal(reopened.slideCount, deck.slideCount);
  assert.equal(reopened.layoutOf(at), section.part);
});

test('Two Content gives both columns, the body in the first and the second empty; the footer placeholders stay on the layout', () => {
  const deck = Deck.open(SHOWCASE);
  const two = deck.layoutList().find((l) => l.name === 'Two Content');
  const at = deck.insertSlide(0, { layout: 'obj', layoutPart: two.part, fromLayout: true, title: 'New slide', body: ['Point one'] });
  const got = placeholders(deck, at);
  assert.deepEqual(got.map((p) => [String(p.idx ?? ''), p.text]), [['', 'New slide'], ['1', 'Point one'], ['2', '']]);
  assert.ok(got[1].geometry.x < got[2].geometry.x, 'side by side, as the layout has them');
  assert.ok(!/type="(?:dt|ftr|sldNum)"/.test(deck.pkg.text(deck.slideParts[at].part)));
});
