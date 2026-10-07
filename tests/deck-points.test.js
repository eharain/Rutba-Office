// Presentations: Edit Points — a shape's outline as points, moved, added
// and taken away, and kept in the file as the shape's own path.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { presetCommands, parsePath, pathData, anchors, moveAnchor, deleteAnchor, insertAnchor, nearestSegment, fit, custGeomXml } from '@rutba/presentation/points';

const blank = () => Deck.open(buildPptx({ title: 'Points', slides: [{ layout: 'blank' }] }));

test('a preset\'s outline as points: a rectangle\'s four corners, an ellipse\'s four quarters', () => {
  assert.equal(pathData(presetCommands('rect', 100, 50)), 'M0 0 L100 0 L100 50 L0 50 Z');
  assert.deepEqual(anchors(presetCommands('rect', 100, 50)).map((a) => [a.x, a.y]), [[0, 0], [100, 0], [100, 50], [0, 50]]);
  const ellipse = presetCommands('ellipse', 100, 100);
  assert.deepEqual(anchors(ellipse).map((a) => [Math.round(a.x), Math.round(a.y)]), [[50, 0], [100, 50], [50, 100], [0, 50]], 'its start shown once, though its last curve comes back to it');
  assert.equal(anchors(presetCommands('line', 80, 0)).length, 2);
  assert.equal(anchors(presetCommands('star5', 100, 100)).length, 10);
});

test('a path read from a file — curves, arcs and all — becomes points in the shape\'s box', () => {
  assert.equal(pathData(parsePath('M0 0 L10 0 L10 10 Z', { w: 10, h: 10, width: 200, height: 100 })), 'M0 0 L200 0 L200 100 Z');
  assert.equal(pathData(parsePath('M0 0 Q50 100 100 0', { w: 100, h: 100 })), 'M0 0 C33.33 66.67 66.67 66.67 100 0', 'a quadratic curve as a cubic one');
  const arc = parsePath('M0 50 A50 50 0 0 1 100 50', { w: 100, h: 100 });
  assert.equal(arc.filter((c) => c.op === 'C').length, 2, 'a half circle as two quarter curves');
  assert.deepEqual(arc[arc.length - 1].pts[2].map(Math.round), [100, 50]);
});

test('moving a point bends the curves either side of it; a figure\'s start moves with the curve that returns to it', () => {
  const rect = presetCommands('rect', 100, 50);
  assert.equal(pathData(moveAnchor(rect, 1, 130, -10)), 'M0 0 L130 -10 L100 50 L0 50 Z');
  const ellipse = presetCommands('ellipse', 100, 100);
  const pulled = moveAnchor(ellipse, 0, 50, -20);
  assert.deepEqual(pulled[0].pts[0], [50, -20]);
  assert.deepEqual(pulled[4].pts[2], [50, -20], 'the last quarter still closes on it');
  assert.equal(pulled[1].pts[0][1], ellipse[1].pts[0][1] - 20, 'the control point leading out moved with it');
  assert.equal(pulled[4].pts[1][1], ellipse[4].pts[1][1] - 20, 'and the one leading in');
});

test('a point taken out joins its neighbours; one added splits a line or a curve in two', () => {
  const rect = presetCommands('rect', 100, 50);
  assert.equal(pathData(deleteAnchor(rect, 2)), 'M0 0 L100 0 L0 50 Z');
  assert.equal(pathData(deleteAnchor(rect, 0)), 'M100 0 L100 50 L0 50 Z', 'the next point starts the figure');
  assert.equal(deleteAnchor(presetCommands('triangle', 10, 10), 1), null, 'a closed figure keeps three');
  assert.equal(pathData(insertAnchor(rect, 1)), 'M0 0 L50 0 L100 0 L100 50 L0 50 Z');
  assert.equal(pathData(insertAnchor(rect, 4)), 'M0 0 L100 0 L100 50 L0 50 L0 25 Z', 'on the line that closes it');
  const curve = insertAnchor(presetCommands('ellipse', 100, 100), 1);
  assert.equal(anchors(curve).length, 5);
  assert.equal(nearestSegment(rect, 50, 1), 1);
  assert.equal(nearestSegment(rect, 1, 25), 4);
});

test('after a point leaves the box, the box is the path\'s', () => {
  const f = fit(moveAnchor(presetCommands('rect', 100, 50), 1, 130, -10));
  assert.deepEqual([f.dx, f.dy, f.w, f.h], [0, -10, 130, 60]);
  assert.equal(pathData(f.commands), 'M0 10 L130 0 L100 60 L0 60 Z');
  // A curve's box is what it draws, not where its control points sit.
  const e = fit(presetCommands('ellipse', 100, 100));
  assert.deepEqual([Math.round(e.w), Math.round(e.h)], [100, 100]);
});

test('kept in the file as the shape\'s own path, drawn and read back the same', () => {
  const deck = blank();
  const id = deck.addShape(0, { preset: 'rect', x: 100, y: 100, w: 200, h: 100, name: 'Box' });
  const f = fit(moveAnchor(presetCommands('rect', 200, 100), 1, 260, -20));
  deck.setGeometry(0, id, { x: 100 + f.dx, y: 100 + f.dy, w: f.w, h: f.h });
  deck.setShapePath(0, id, { commands: f.commands, w: f.w, h: f.h });
  const xml = deck.pkg.text('ppt/slides/slide1.xml');
  assert.doesNotMatch(xml, /<a:prstGeom prst="rect"/, 'no preset now');
  assert.match(xml, /<a:custGeom><a:avLst\/><a:gdLst\/><a:ahLst\/><a:cxnLst\/><a:rect l="l" t="t" r="r" b="b"\/><a:pathLst><a:path w="2476500" h="1143000"><a:moveTo><a:pt x="0" y="190500"\/><\/a:moveTo>/);
  const shape = Deck.open(deck.save()).slide(0).shapes.find((s) => s.name === 'Box');
  assert.equal(shape.preset, 'custom');
  assert.deepEqual([Math.round(shape.geometry.x), Math.round(shape.geometry.y), Math.round(shape.geometry.w), Math.round(shape.geometry.h)], [100, 80, 260, 120]);
  const again = parsePath(shape.path.d, { w: shape.path.w, h: shape.path.h, width: shape.geometry.w, height: shape.geometry.h });
  assert.deepEqual(anchors(again).map((a) => [Math.round(a.x), Math.round(a.y)]), [[0, 20], [260, 0], [200, 120], [0, 120]], 'the points as they were left');
  assert.match(renderSlide(deck.slide(0), { width: 1280 }), /<path d="M100 100L360 80L300 200L100 200Z"/, 'drawn where its points are');
  // An outline only, for a line.
  assert.match(custGeomXml(presetCommands('line', 10, 0), 10, 1, { filled: false }), /<a:path w="95250" h="9525" fill="none">/);
  assert.throws(() => deck.setShapePath(0, id, { commands: [], w: 1, h: 1 }), /starts with a point/);
});
