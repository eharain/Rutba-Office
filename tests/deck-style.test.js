// The Format pane — the deck engine's half: a shape's fill and outline are
// written into its own spPr, in schema order, and read back by the scene.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({ title: 'Style', slides: [{ layout: 'obj', title: 'A shape', body: ['One'] }] });

test('setShapeStyle writes a solid fill and an outline, and the scene reads them', () => {
  const deck = Deck.open(DECK);
  const id = deck.addShape(0, { preset: 'roundRect', x: 100, y: 100, w: 300, h: 150 });
  const shape = () => deck.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(deck.setShapeStyle(0, id, { fill: '#ff0000' }), true);
  assert.deepEqual({ type: shape().fill.type, color: shape().fill.color }, { type: 'solid', color: '#ff0000' });
  assert.ok(shape().line && shape().line.type !== 'none' && shape().line.color, 'the outline addShape gave it survives a fill-only write');

  assert.equal(deck.setShapeStyle(0, id, { line: { color: '#00ff00', width: 3, dash: 'dash' } }), true);
  assert.equal(shape().fill.color, '#ff0000', 'a line-only write keeps the fill');
  assert.equal(shape().line.color, '#00ff00');
  assert.equal(shape().line.width, 3);
  assert.equal(shape().line.dash, 'dash');

  // Schema order inside spPr: geometry, then fill, then line.
  const xml = OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part);
  const sp = xml.slice(xml.indexOf(`<p:cNvPr id="${id}"`));
  const spPr = /<p:spPr>[\s\S]*?<\/p:spPr>/.exec(sp)[0];
  const order = ['<a:prstGeom', '<a:solidFill', '<a:ln'].map((t) => spPr.indexOf(t));
  assert.ok(order[0] < order[1] && order[1] < order[2], `geometry, fill, line in that order: ${order.join(',')}`);
  assert.equal((spPr.match(/<a:solidFill>/g) || []).length, 2, 'one fill for the shape, one inside the line — no leftovers');

  assert.equal(deck.setShapeStyle(0, id, { fill: 'none', line: 'none' }), true);
  assert.equal(shape().fill.type, 'none');
  assert.equal(shape().line.type, 'none');

  assert.equal(deck.setShapeStyle(0, id, { fill: { scheme: 'accent2' } }), true);
  assert.equal(shape().fill.type, 'solid');
  assert.match(shape().fill.color, /^#[0-9A-Fa-f]{6}$/, 'a scheme colour resolves through the theme');

  const reopened = Deck.open(deck.save());
  const again = reopened.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(again.fill.color, shape().fill.color, 'in the file');
  assert.equal(again.line.type, 'none');
});

test('setShapeStyle fills a placeholder that had no fill of its own, and refuses a frame', () => {
  const deck = Deck.open(DECK);
  const title = deck.slide(0).shapes[0];
  assert.equal(deck.setShapeStyle(0, title.id, { fill: '#123456' }), true);
  assert.equal(deck.slide(0).shapes[0].fill.color, '#123456');
  assert.throws(() => deck.setShapeStyle(0, '999', { fill: '#000000' }), /not found/);
});
