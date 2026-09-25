// Review → Check Accessibility in Presentation: a picture with no
// description (written on `p:cNvPr descr`), a slide with no title, two with
// the same one, objects layered out of reading order, pale text on a shape's
// fill, and a table whose first row is not a header.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { checkAccessibility, describeDeck, setDeckAltText, setDeckTableHeader, setDeckSlideTitle, setDeckTextColour } from '@rutba/proofing';
import { gradientPng } from '../apps/desktop/main/sample-picture.js';

function fixture() {
  const deck = Deck.open(buildPptx({ title: 'Checks', slides: [
    { layout: 'title', title: 'Quarterly review', body: 'Spring' },
    { layout: 'obj', title: 'Results', body: ['Orders up', 'Returns down'] },
    { layout: 'obj', title: 'Results', body: ['More results'] },
    { layout: 'obj', title: '', body: ['A slide with no title'] },
  ] }));
  // Slide 2: a picture layered before the title though it sits below it.
  // Its description is its file's name, which is what a picture inserted by
  // name carries and which says nothing to a person who cannot see it.
  const picture = deck.addPicture(1, { data: gradientPng(20, 20, [30, 30, 200], [200, 30, 30]), contentType: 'image/png', name: 'IMG_2041.png', x: 600, y: 300, w: 200, h: 150 }).id;
  deck.reorderShape(1, picture, 'back');
  // Slide 1: pale words on a white box; slide 3: a table with no header row.
  const box = deck.addTextBox(0, { x: 100, y: 420, w: 400, h: 60, paragraphs: [{ runs: [{ text: 'Faint footnote', color: '#DDDDDD' }] }], name: 'Faint' });
  deck.setShapeStyle(0, box, { fill: { type: 'solid', color: '#FFFFFF' } });
  const table = deck.addTable(2, { rows: 2, cols: 2, x: 100, y: 300, w: 400, h: 80 });
  const part = deck.slideParts[2].part;
  deck.pkg.write_(part, Buffer.from(deck.pkg.text(part).replace('<a:tblPr firstRow="1"', '<a:tblPr'), 'utf8'));
  return { deck, picture: String(picture), box: String(box), table: String(table) };
}

test('a deck is described slide by slide and flagged', () => {
  const { deck, picture, box, table } = fixture();
  const { issues } = checkAccessibility(describeDeck(deck));
  const by = (rule) => issues.filter((i) => i.rule === rule);
  assert.deepEqual(by('altText').map((i) => i.where), [{ slide: 1, shape: picture }]);
  assert.deepEqual(by('slideTitle').map((i) => i.where.slide), [3]);
  assert.ok(by('slideTitle')[0].fixes[0], 'the empty title placeholder can take one');
  assert.deepEqual(by('duplicateTitle').map((i) => i.where.slide), [1, 2]);
  assert.deepEqual(by('readingOrder').map((i) => i.where.slide), [1], 'the picture is read before the title');
  assert.deepEqual(by('contrast').map((i) => i.where), [{ slide: 0, shape: box }]);
  assert.deepEqual(by('tableHeader').map((i) => i.where), [{ slide: 2, shape: table }]);
});

test('the fixes: a description and decorative on p:cNvPr, a header row, a title, a readable colour', () => {
  const { deck, picture, box, table } = fixture();
  setDeckAltText(deck, { slide: 1, shape: picture, descr: 'Orders by month' });
  assert.match(deck.pkg.text(deck.slideParts[1].part), /<p:cNvPr [^>]*descr="Orders by month"/);
  setDeckAltText(deck, { slide: 1, shape: picture, decorative: true });
  assert.match(deck.pkg.text(deck.slideParts[1].part), /<adec:decorative [^>]*val="1"\/>/);

  setDeckTableHeader(deck, { slide: 2, shape: table });
  assert.match(deck.pkg.text(deck.slideParts[2].part), /<a:tblPr firstRow="1"/);

  const titleIssue = checkAccessibility(describeDeck(deck)).issues.find((i) => i.rule === 'slideTitle');
  setDeckSlideTitle(deck, { ...titleIssue.fixes[0].target, title: 'Next steps' });
  setDeckTextColour(deck, { slide: 0, shape: box, colour: '#000000' });

  const after = checkAccessibility(describeDeck(deck)).issues.map((i) => i.rule);
  for (const rule of ['altText', 'tableHeader', 'slideTitle', 'contrast']) assert.ok(!after.includes(rule), `${rule} is fixed: ${after.join(', ')}`);
  const reopened = Deck.open(deck.save());
  const described = describeDeck(reopened);
  assert.equal(described.slides[3].title, 'Next steps', 'the saved deck has the title');
  assert.equal(described.objects.find((o) => o.where.shape === picture).decorative, true);
});
