// A shape's fill beyond solid colour — gradient, picture, transparency —
// and its effects beyond a shadow — glow, soft edges, reflection — written
// by `setShapeStyle`, read back by the scene, and kept through a save.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({ title: 'Fill and effects', slides: [{ layout: 'obj', title: 'A shape', body: ['One'] }] });

// A tiny valid PNG (1×1, red) — enough bytes for the engine, which never decodes them.
const PNG_1PX = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  'base64'
);

const shapeOf = (deck, id) => deck.slide(0).shapes.find((s) => String(s.id) === String(id));
const spXmlOf = (deck, id) => {
  const xml = OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part);
  const m = new RegExp(`<p:cNvPr id="${id}"[\\s\\S]*?</p:sp>`).exec(xml);
  return m ? m[0] : '';
};

test('a solid fill can carry transparency, written as an alpha under the colour', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 200, h: 100 });
  assert.equal(deck.setShapeStyle(0, id, { fill: { color: '#3366CC', alpha: 0.4 } }), true);
  const fill = shapeOf(deck, id).fill;
  assert.equal(fill.type, 'solid');
  assert.equal(fill.color, '#3366cc');
  assert.equal(Math.round(fill.alpha * 100) / 100, 0.4);
  assert.match(spXmlOf(deck, id), /<a:solidFill><a:srgbClr val="3366CC"><a:alpha val="40000"\/><\/a:srgbClr><\/a:solidFill>/);
});

test('a gradient fill writes every stop and the angle, and reads all of them back', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 200, h: 100 });
  assert.equal(deck.setShapeStyle(0, id, {
    fill: { gradient: { angle: 45, stops: [{ pos: 0, color: '#FFFFFF' }, { pos: 0.5, color: '#FF0000', alpha: 0.5 }, { pos: 1, color: '#000000' }] } },
  }), true);
  const fill = shapeOf(deck, id).fill;
  assert.equal(fill.type, 'gradient');
  assert.equal(fill.stops.length, 3);
  assert.deepEqual(fill.stops.map((s) => s.offset), [0, 0.5, 1]);
  assert.equal(fill.stops[1].color, '#ff0000');
  assert.equal(fill.stops[1].alpha, 0.5);
  assert.equal(fill.angle, 45);
  assert.match(spXmlOf(deck, id), /<a:gradFill[^>]*><a:gsLst><a:gs pos="0">[\s\S]*<a:gs pos="50000">[\s\S]*<a:gs pos="100000">[\s\S]*<\/a:gsLst><a:lin ang="2700000"/);
});

test('a gradient preset builds a light or a dark variation of one colour', () => {
  const deck = Deck.open(DECK);
  const light = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 100, h: 100 });
  deck.setShapeStyle(0, light, { fill: { gradient: { preset: 'light', color: { scheme: 'accent1' } } } });
  const lightFill = shapeOf(deck, light).fill;
  assert.equal(lightFill.stops.length, 2);
  // The light preset lightens toward the second stop — its colour should not equal the first's.
  assert.notEqual(lightFill.stops[0].color, lightFill.stops[1].color);

  const dark = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 100, h: 100 });
  deck.setShapeStyle(0, dark, { fill: { gradient: { preset: 'dark', color: '#4472C4' } } });
  const darkFill = shapeOf(deck, dark).fill;
  assert.equal(darkFill.stops.length, 2);
  assert.notEqual(darkFill.stops[0].color, darkFill.stops[1].color);
});

test('a picture fill embeds a media part and a relationship, and clips to the shape', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'ellipse', x: 0, y: 0, w: 200, h: 150 });
  assert.equal(deck.setShapeStyle(0, id, { fill: { picture: { data: PNG_1PX, contentType: 'image/png', name: 'dot' } } }), true);
  const fill = shapeOf(deck, id).fill;
  assert.equal(fill.type, 'picture');
  assert.ok(fill.embed);
  assert.equal(fill.tile, false);
  assert.ok(fill.source?.part?.startsWith('ppt/media/image'), `media part: ${fill.source?.part}`);
  assert.deepEqual(deck.media(fill.source.part), PNG_1PX);
  assert.match(spXmlOf(deck, id), /<a:blipFill><a:blip r:embed="rId\d+"\/><a:stretch><a:fillRect\/><\/a:stretch><\/a:blipFill>/);

  const reopened = Deck.open(deck.save());
  const rf = shapeOf(reopened, id).fill;
  assert.equal(rf.type, 'picture');
  assert.deepEqual(reopened.media(rf.source.part), PNG_1PX);
});

test('a tiled picture fill writes a:tile, and clearing the fill takes the blipFill off', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 200, h: 150 });
  deck.setShapeStyle(0, id, { fill: { picture: { data: PNG_1PX, contentType: 'image/png', tile: true } } });
  assert.equal(shapeOf(deck, id).fill.tile, true);
  assert.match(spXmlOf(deck, id), /<a:blipFill><a:blip r:embed="rId\d+"\/><a:tile\/><\/a:blipFill>/);

  deck.setShapeStyle(0, id, { fill: 'none' });
  assert.equal(shapeOf(deck, id).fill.type, 'none');
  assert.ok(!/<a:blipFill>/.test(spXmlOf(deck, id)), 'the blipFill is gone');
});

test('a picture fill\'s tile can be changed without giving the bytes again — it keeps the blip it already had', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 200, h: 150 });
  deck.setShapeStyle(0, id, { fill: { picture: { data: PNG_1PX, contentType: 'image/png' } } });
  const embed = shapeOf(deck, id).fill.embed;
  assert.equal(deck.setShapeStyle(0, id, { fill: { picture: { tile: true } } }), true);
  const after = shapeOf(deck, id).fill;
  assert.equal(after.tile, true);
  assert.equal(after.embed, embed, 'the same relationship — no picture was re-embedded');
  assert.throws(() => deck.setShapeStyle(0, deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 10, h: 10 }), { fill: { picture: { tile: true } } }), /no picture fill/);
});

test('replacing a picture fill with another does not disturb the shape — the first media part is simply left unreferenced', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 200, h: 150 });
  deck.setShapeStyle(0, id, { fill: { picture: { data: PNG_1PX, contentType: 'image/png' } } });
  const first = shapeOf(deck, id).fill.source.part;
  deck.setShapeStyle(0, id, { fill: { picture: { data: PNG_1PX, contentType: 'image/png' } } });
  const second = shapeOf(deck, id).fill.source.part;
  assert.notEqual(first, second, 'a fresh media part, not a reused one');
  assert.deepEqual(deck.media(first), PNG_1PX, 'the first part is still in the package, just unreferenced');
  assert.deepEqual(deck.media(second), PNG_1PX);
});

test('glow is written before the shadow, in points converted to EMU, and reads back', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 200, h: 100 });
  assert.equal(deck.setShapeStyle(0, id, { effects: { glow: { radius: 8, color: '#00B0F0', alpha: 0.75 } } }), true);
  const glow = shapeOf(deck, id).effects.glow;
  assert.equal(glow.color, '#00b0f0');
  assert.equal(glow.alpha, 0.75);
  assert.equal(Math.round(glow.radiusPt), 8);
  assert.match(spXmlOf(deck, id), /<a:effectLst><a:glow rad="101600">/);
});

test('soft edges write a radius in points, converted to EMU', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 200, h: 100 });
  assert.equal(deck.setShapeStyle(0, id, { effects: { softEdge: { radius: 5 } } }), true);
  assert.equal(Math.round(shapeOf(deck, id).effects.softEdge.radiusPt), 5);
  assert.match(spXmlOf(deck, id), /<a:effectLst><a:softEdge rad="63500"\/><\/a:effectLst>/);
});

test('a reflection preset round-trips through its own numbers — tight, half and full read back as themselves', () => {
  const deck = Deck.open(DECK);
  for (const kind of ['tight', 'half', 'full']) {
    const id = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 100, h: 60 });
    deck.setShapeStyle(0, id, { effects: { reflection: kind } });
    assert.equal(shapeOf(deck, id).effects.reflection, kind, `${kind} round-trips`);
  }
});

test('glow, shadow, reflection and soft edge combine in one effectLst, in schema order, and "none" clears all four at once', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 200, h: 100 });
  deck.setShapeStyle(0, id, {
    effects: {
      glow: { radius: 5, color: '#FFC000' },
      shadow: { dist: 3, dir: 45, blur: 4, color: '#000000', alpha: 0.4 },
      reflection: 'half',
      softEdge: { radius: 2.5 },
    },
  });
  const effects = shapeOf(deck, id).effects;
  assert.ok(effects.glow && effects.shadow && effects.reflection && effects.softEdge, JSON.stringify(effects));
  const xml = spXmlOf(deck, id);
  const order = ['a:glow', 'a:outerShdw', 'a:reflection', 'a:softEdge'].map((tag) => xml.indexOf(`<${tag}`));
  assert.ok(order.every((n) => n >= 0), `all four present: ${JSON.stringify(order)}`);
  assert.ok(order[0] < order[1] && order[1] < order[2] && order[2] < order[3], `schema order: ${JSON.stringify(order)}`);

  assert.equal(deck.setShapeStyle(0, id, { effects: 'none' }), true);
  assert.deepEqual(shapeOf(deck, id).effects, { shadow: null });
  assert.match(spXmlOf(deck, id), /<a:effectLst\/>/);

  const reopened = Deck.open(deck.save());
  assert.deepEqual(shapeOf(reopened, id).effects, { shadow: null }, 'cleared through a save too');
});

test('setting one effect keeps the whole list explicit, not merged — asking for a shadow alone turns an existing glow off, the way solid fill and outline already behave', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 200, h: 100 });
  deck.setShapeStyle(0, id, { effects: { glow: { radius: 8, color: '#FFC000' } } });
  assert.ok(shapeOf(deck, id).effects.glow);
  deck.setShapeStyle(0, id, { effects: { shadow: { dist: 3, dir: 45, blur: 4, color: '#000000', alpha: 0.4 } } });
  const effects = shapeOf(deck, id).effects;
  assert.ok(effects.shadow, 'the shadow is there');
  assert.equal(effects.glow, undefined, 'the glow is gone: only what is named stays on');
});
