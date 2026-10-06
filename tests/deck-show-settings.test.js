// Set Up Slide Show, kept in the deck where PowerPoint keeps it.
//
// ppt/presProps.xml's p:showPr carries who the show is for, whether it loops
// and plays its narration and animations, whether it keeps to its timings,
// which slides it shows and the pen's colour. A deck without the part answers
// PowerPoint's defaults; writing makes the part, and keeps the rest of one
// that is already there.

import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';

const deck = (n = 4) => Deck.open(buildPptx({ title: 'Show', slides: Array.from({ length: n }, (_, i) => ({ layout: 'title', title: `Slide ${i + 1}` })) }));

test('a deck with no settings answers PowerPoint\'s defaults', () => {
  assert.deepEqual(deck().showSettings(), { type: 'present', loop: false, narration: true, animation: true, useTimings: true, range: null, pen: null });
});

test('written settings come back after a save, in a part made with its relationship and content type', () => {
  const d = deck();
  d.setShowSettings({ loop: true, animation: false, useTimings: false, range: { from: 2, to: 3 }, pen: '#00aa55' });
  const back = Deck.open(d.save());
  assert.deepEqual(back.showSettings(), { type: 'present', loop: true, narration: true, animation: false, useTimings: false, range: { from: 2, to: 3 }, pen: '#00AA55' });
  assert.equal(back.pkg.contentTypeOf('ppt/presProps.xml'), 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml');
  assert.match(back.pkg.text('ppt/_rels/presentation.xml.rels'), /relationships\/presProps" Target="presProps\.xml"/);
  assert.match(back.pkg.text('ppt/presProps.xml'), /<p:showPr loop="1" showNarration="1" showAnimation="0" useTimings="0"><p:present\/><p:sldRg st="2" end="3"\/><p:penClr><a:srgbClr val="00AA55"\/><\/p:penClr><\/p:showPr>/);
});

test('a kiosk always loops, a range is kept inside the deck, and every slide is written as all of them', () => {
  const d = deck(4);
  assert.equal(d.setShowSettings({ type: 'kiosk', loop: false }).loop, true, 'a kiosk loops whatever was asked');
  assert.deepEqual(d.setShowSettings({ type: 'present', range: { from: 3, to: 99 } }).range, { from: 3, to: 4 }, 'clamped to the deck');
  assert.equal(d.setShowSettings({ range: { from: 1, to: 4 } }).range, null, 'the whole deck is "all"');
  assert.throws(() => d.setShowSettings({ type: 'broadcast' }), /unknown show type/);
});

test('the rest of a presProps part already there is kept as it was', () => {
  const d = deck();
  const kept = '<p:extLst><p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="0"/></p:ext></p:extLst>';
  d.setShowSettings({ loop: true });
  d.pkg.write_('ppt/presProps.xml', Buffer.from(d.pkg.text('ppt/presProps.xml').replace('</p:presentationPr>', () => kept + '</p:presentationPr>'), 'utf8'));
  d.setShowSettings({ loop: false, pen: '#112233' });
  const xml = d.pkg.text('ppt/presProps.xml');
  assert.ok(xml.includes(kept), 'the extension list survives');
  assert.equal((xml.match(/<p:showPr\b/g) || []).length, 1, 'one showPr, rebuilt in place');
  assert.ok(xml.indexOf('<p:showPr') < xml.indexOf('<p:extLst'), 'and still ahead of the extensions, as the schema orders them');
});
