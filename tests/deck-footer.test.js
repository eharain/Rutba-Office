// Insert → Header & Footer — the deck engine's half.
//
// The footer's words, the slide number and the date are the three footer
// placeholders PowerPoint writes; the number and an automatic date are
// fields. With no layout to place them they go along the bottom where
// PowerPoint's own templates put them; the number says the slide's place.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({ title: 'Footer', slides: [{ layout: 'obj', title: 'One', body: ['a'] }, { layout: 'obj', title: 'Two', body: ['b'] }] });
const ofType = (deck, slide, type) => deck.slide(slide).shapes.find((s) => s.placeholder?.type === type) || null;
const words = (s) => (s?.text?.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join('')).join('|');

test('a footer, a slide number and a date are written as placeholders along the bottom, the number following the slide, and taken off again', () => {
  const deck = Deck.open(DECK);
  assert.equal(deck.setFooter(1, { footer: 'Rutba Office', slideNumber: true, date: { auto: true } }), true);
  const ftr = ofType(deck, 1, 'ftr');
  const num = ofType(deck, 1, 'sldNum');
  const dt = ofType(deck, 1, 'dt');
  assert.ok(ftr && num && dt, 'three placeholders');
  assert.equal(words(ftr), 'Rutba Office');
  assert.equal(words(num), '2', 'the number is the slide\'s place');
  assert.equal(num.text.paragraphs[0].runs[0].field, 'slidenum');
  assert.equal(dt.text.paragraphs[0].runs[0].field, 'datetime1');
  assert.match(words(dt), /^\d{2}\/\d{2}\/\d{4}$/, 'an automatic date shows today');
  for (const s of [ftr, num, dt]) {
    assert.ok(s.geometry && s.geometry.y > deck.size.height * 0.85, `${s.placeholder.type} sits along the bottom`);
  }
  assert.ok(dt.geometry.x < ftr.geometry.x && ftr.geometry.x < num.geometry.x, 'date, footer, number from left to right');
  assert.equal(ofType(deck, 0, 'ftr'), null, 'the other slide is untouched');

  // The XML, as PowerPoint writes it.
  const xml = OoxmlPackage.read(deck.save()).text(deck.slideParts[1].part);
  assert.match(xml, /<p:nvPr><p:ph type="ftr" sz="quarter" idx="11"\/><\/p:nvPr>/);
  assert.match(xml, /<p:ph type="sldNum" sz="quarter" idx="12"\/>/);
  assert.match(xml, /<a:fld id="\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}" type="slidenum"><a:rPr lang="en-US"\/><a:t>2<\/a:t><\/a:fld>/);
  assert.match(xml, /<a:fld id="\{[0-9A-F-]{36}\}" type="datetime1">/);
  assert.match(xml, /<a:pPr algn="r"\/><a:fld/, 'the number is right-aligned in its own band');

  // The number follows the slide.
  deck.moveSlide(1, 0);
  assert.equal(words(ofType(deck, 0, 'sldNum')), '1', 'moved to the front, it says 1');
  assert.equal(ofType(deck, 1, 'sldNum'), null);

  // New words replace the old in the same placeholder; blanks take it off.
  deck.setFooter(0, { footer: 'Changed' });
  assert.equal(words(ofType(deck, 0, 'ftr')), 'Changed');
  assert.equal(deck.slide(0).shapes.filter((s) => s.placeholder?.type === 'ftr').length, 1, 'still one footer');
  assert.equal(words(ofType(deck, 0, 'sldNum')), '1', 'the number was left alone');
  deck.setFooter(0, { footer: '', slideNumber: false, date: false });
  assert.equal(ofType(deck, 0, 'ftr'), null);
  assert.equal(ofType(deck, 0, 'sldNum'), null);
  assert.equal(ofType(deck, 0, 'dt'), null);
  assert.equal(deck.slide(0).shapes.length, 2, 'the title and body remain');
});

test('a fixed date is plain words, and a footer put on every slide by the service op reaches each one', () => {
  const deck = Deck.open(DECK);
  deck.setFooter(0, { date: { text: 'Spring 2026' } });
  const dt = ofType(deck, 0, 'dt');
  assert.equal(words(dt), 'Spring 2026');
  assert.equal(dt.text.paragraphs[0].runs[0].field, undefined, 'not a field');
  for (let i = 0; i < deck.slideCount; i++) deck.setFooter(i, { slideNumber: true });
  assert.equal(words(ofType(deck, 0, 'sldNum')), '1');
  assert.equal(words(ofType(deck, 1, 'sldNum')), '2');
  const reopened = Deck.open(deck.save());
  assert.equal(words(ofType(reopened, 1, 'sldNum')), '2', 'through a save');
  assert.equal(words(ofType(reopened, 0, 'dt')), 'Spring 2026');
});
