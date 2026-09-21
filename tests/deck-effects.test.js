// Shape Effects — the deck engine's half: an outer shadow written after the
// outline in the shape's own properties, read back in pixels and degrees,
// kept through a fill write, and turned off with the empty effect list
// PowerPoint writes to silence a style's shadow.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({ title: 'Effects', slides: [{ layout: 'obj', title: 'A shape', body: ['One'] }] });

test('a shadow is written after the outline, read back, kept under a fill write, and taken off with an empty effect list', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'roundRect', x: 100, y: 100, w: 300, h: 150 });
  const shape = () => deck.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(shape().effects, null, 'a fresh shape states no effects');

  assert.equal(deck.setShapeStyle(0, id, { effects: { shadow: { dist: 3, dir: 45, blur: 4, color: '#000000', alpha: 0.4 } } }), true);
  const sh = shape().effects.shadow;
  assert.equal(Math.round(sh.dir), 45);
  assert.equal(Math.round(sh.distPx), 4, '3 pt is 4 px');
  assert.equal(Math.round(sh.blurPx * 10) / 10, 5.3, '4 pt of blur');
  assert.equal(sh.color, '#000000');
  assert.equal(sh.alpha, 0.4);

  const spPrOf = () => {
    const xml = OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part);
    const sp = xml.slice(xml.indexOf(`<p:cNvPr id="${id}"`));
    return /<p:spPr>[\s\S]*?<\/p:spPr>/.exec(sp)[0];
  };
  assert.match(spPrOf(), /<\/a:ln><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"\/><\/a:srgbClr><\/a:outerShdw><\/a:effectLst><\/p:spPr>/, 'after the outline, as PowerPoint writes it');

  deck.setShapeStyle(0, id, { fill: '#ff0000' });
  assert.equal(Math.round(shape().effects.shadow.dir), 45, 'a fill write keeps the shadow');
  assert.equal(shape().fill.color, '#ff0000');

  deck.setShapeStyle(0, id, { effects: { shadow: { dist: 0, dir: 0, blur: 6, color: '#1F4E79', alpha: 0.45 } } });
  assert.equal(shape().effects.shadow.distPx, 0, 'a shadow all round has no offset');
  assert.equal(shape().effects.shadow.color, '#1f4e79');

  deck.setShapeStyle(0, id, { effects: 'none' });
  assert.deepEqual(shape().effects, { shadow: null }, 'an empty effect list, so a style shadow stays off too');
  assert.match(spPrOf(), /<\/a:ln><a:effectLst\/><\/p:spPr>/);
  assert.equal(Deck.open(deck.save()).slide(0).shapes.find((s) => String(s.id) === String(id)).effects.shadow, null, 'through a save');
});
