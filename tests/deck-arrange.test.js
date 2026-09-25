// Home → Arrange — the deck engine's half: Align (to the slide or to the
// selection's own bounds), Distribute, Rotate, Flip, and Group/Ungroup.
//
// A group is one `p:grpSp` whose child coordinate window (chOff/chExt)
// starts equal to its own bounding box, so grouping never moves or resizes
// a member; only a later resize of the group scales them, because setGeometry
// on the group's own id writes off/ext and leaves chOff/chExt alone.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { renderSlide } from '@rutba/presentation';
import { OoxmlPackage } from '@rutba/ooxml/package';

const DECK = buildPptx({ title: 'Arrange', slides: [{ layout: 'obj', title: 'Shapes', body: ['One'] }] });

const shapeOf = (deck, id) => deck.slide(0).shapes.find((s) => String(s.id) === String(id));
const near = (a, b, tol = 0.6) => Math.abs(a - b) <= tol;

test('alignShapes lines a single shape up to the slide, on every edge', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 300, y: 300, w: 200, h: 80 });

  assert.equal(deck.alignShapes(0, [a], 'left'), true);
  assert.ok(near(shapeOf(deck, a).geometry.x, 0));
  assert.equal(deck.alignShapes(0, [a], 'right'), true);
  assert.ok(near(shapeOf(deck, a).geometry.x, deck.size.width - 200));
  assert.equal(deck.alignShapes(0, [a], 'center'), true);
  assert.ok(near(shapeOf(deck, a).geometry.x, (deck.size.width - 200) / 2));
  assert.equal(deck.alignShapes(0, [a], 'top'), true);
  assert.ok(near(shapeOf(deck, a).geometry.y, 0));
  assert.equal(deck.alignShapes(0, [a], 'bottom'), true);
  assert.ok(near(shapeOf(deck, a).geometry.y, deck.size.height - 80));
  assert.equal(deck.alignShapes(0, [a], 'middle'), true);
  assert.ok(near(shapeOf(deck, a).geometry.y, (deck.size.height - 80) / 2));
  // Already there: nothing to write.
  assert.equal(deck.alignShapes(0, [a], 'middle'), false);

  // The ribbon always sends its own Align to Slide/Selected Objects toggle
  // along, and that toggle defaults to "Selected Objects" — a single shape
  // must still go to the slide even when `to: 'selection'` arrives explicitly,
  // since there is no selection of more than one to align it to.
  const deck2 = Deck.open(DECK);
  const b = deck2.addShape(0, { preset: 'rect', x: 300, y: 300, w: 200, h: 80 });
  assert.equal(deck2.alignShapes(0, [b], 'left', { to: 'selection' }), true);
  assert.ok(near(shapeOf(deck2, b).geometry.x, 0), 'one shape aligns to the slide regardless of the toggle');
});

test('alignShapes lines several shapes up to their own combined bounds by default, and to the slide when asked', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 100, y: 50, w: 200, h: 80 });
  const b = deck.addShape(0, { preset: 'rect', x: 500, y: 300, w: 100, h: 150 });
  const c = deck.addShape(0, { preset: 'rect', x: 300, y: 500, w: 60, h: 60 });

  assert.equal(deck.alignShapes(0, [a, b, c], 'left'), true);
  const leftmost = Math.min(100, 500, 300);
  for (const id of [a, b, c]) assert.ok(near(shapeOf(deck, id).geometry.x, leftmost), `${id} at the selection's left`);

  // Align to Slide: every shape goes to the slide's own edge instead.
  const deck2 = Deck.open(DECK);
  const a2 = deck2.addShape(0, { preset: 'rect', x: 100, y: 50, w: 200, h: 80 });
  const b2 = deck2.addShape(0, { preset: 'rect', x: 500, y: 300, w: 100, h: 150 });
  assert.equal(deck2.alignShapes(0, [a2, b2], 'right', { to: 'slide' }), true);
  assert.ok(near(shapeOf(deck2, a2).geometry.x, deck2.size.width - 200));
  assert.ok(near(shapeOf(deck2, b2).geometry.x, deck2.size.width - 100));
});

test('alignShapes lines a rotated shape up by its visual bounds, not its unrotated box', () => {
  const deck = Deck.open(DECK);
  // A 100×40 rectangle turned 90° looks 40 wide, 100 tall.
  const a = deck.addShape(0, { preset: 'rect', x: 200, y: 200, w: 100, h: 40 });
  deck.setGeometry(0, a, { x: 200, y: 200, w: 100, h: 40, rot: 90 });
  assert.equal(deck.alignShapes(0, [a], 'left'), true);
  const box = shapeOf(deck, a).geometry;
  // Its visual left edge — centre minus half the rotated width (40) — sits at 0.
  const cx = box.x + box.w / 2;
  assert.ok(near(cx - box.h / 2, 0), `visual left edge at 0, got centre ${cx} height ${box.h}`);
});

test('distributeShapes spaces three or more shapes with equal gaps, between the outer two by default or across the slide when asked', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 0, y: 100, w: 100, h: 50 });
  const b = deck.addShape(0, { preset: 'rect', x: 200, y: 100, w: 40, h: 50 });
  const c = deck.addShape(0, { preset: 'rect', x: 900, y: 100, w: 100, h: 50 });
  assert.equal(deck.distributeShapes(0, [a, b, c], 'horizontal'), true);
  // The outer two never move (they define the span); the middle one gets an equal gap either side.
  assert.ok(near(shapeOf(deck, a).geometry.x, 0));
  assert.ok(near(shapeOf(deck, c).geometry.x, 900));
  const gapBefore = shapeOf(deck, b).geometry.x - (shapeOf(deck, a).geometry.x + 100);
  const gapAfter = 900 - (shapeOf(deck, b).geometry.x + 40);
  assert.ok(near(gapBefore, gapAfter, 1), `equal gaps: ${gapBefore} vs ${gapAfter}`);

  assert.throws(() => deck.distributeShapes(0, [a, b], 'horizontal'), /three or more/);

  // Across the slide: the first shape's own left edge moves to 0 and the
  // last's right edge to the slide's own width, not to where they were.
  const deck2 = Deck.open(DECK);
  const a2 = deck2.addShape(0, { preset: 'rect', x: 40, y: 100, w: 100, h: 50 });
  const b2 = deck2.addShape(0, { preset: 'rect', x: 400, y: 100, w: 40, h: 50 });
  const c2 = deck2.addShape(0, { preset: 'rect', x: 700, y: 100, w: 60, h: 50 });
  assert.equal(deck2.distributeShapes(0, [a2, b2, c2], 'horizontal', { to: 'slide' }), true);
  assert.ok(near(shapeOf(deck2, a2).geometry.x, 0));
  assert.ok(near(shapeOf(deck2, c2).geometry.x + 60, deck2.size.width));
});

test('rotateShapes turns each shape about its own centre, wrapping at 360°, and each shape turns independently', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 100, y: 100, w: 200, h: 80 });
  const b = deck.addShape(0, { preset: 'rect', x: 600, y: 100, w: 60, h: 60 });
  assert.equal(deck.rotateShapes(0, [a, b], 90), true);
  assert.equal(shapeOf(deck, a).geometry.rot, 90);
  assert.equal(shapeOf(deck, b).geometry.rot, 90);
  // Each shape's own box stayed put — a rigid group rotation would have swung b around a's centre too.
  assert.ok(near(shapeOf(deck, b).geometry.x + 30, 630), 'b rotated about its own centre, not the selection’s');
  assert.equal(deck.rotateShapes(0, [a], -90), true);
  assert.equal(shapeOf(deck, a).geometry.rot, 0);
  assert.equal(deck.rotateShapes(0, [a], -15), true);
  assert.equal(shapeOf(deck, a).geometry.rot, 345, 'wraps to stay in 0..360');
});

test('flipShapes mirrors a shape about its own centre, leaves its rotation exactly as it was, and survives a save', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 100, y: 100, w: 200, h: 80 });
  deck.setGeometry(0, a, { x: 100, y: 100, w: 200, h: 80, rot: 30 });
  assert.equal(deck.flipShapes(0, [a], 'horizontal'), true);
  assert.equal(shapeOf(deck, a).geometry.flipH, true);
  assert.equal(shapeOf(deck, a).geometry.flipV, false);
  assert.equal(shapeOf(deck, a).geometry.rot, 30, 'the rotation is untouched');
  assert.equal(deck.flipShapes(0, [a], 'vertical'), true);
  assert.equal(shapeOf(deck, a).geometry.flipV, true);
  assert.equal(deck.flipShapes(0, [a], 'horizontal'), true);
  assert.equal(shapeOf(deck, a).geometry.flipH, false, 'flipping again takes it off');

  const reopened = Deck.open(deck.save());
  const g = shapeOf(reopened, a).geometry;
  assert.equal(g.flipH, false);
  assert.equal(g.flipV, true);
  assert.equal(g.rot, 30);
});

test('groupShapes gathers shapes into one p:grpSp whose chOff/chExt starts equal to its own bounding box', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 100, y: 100, w: 200, h: 80 });
  const b = deck.addShape(0, { preset: 'ellipse', x: 400, y: 250, w: 100, h: 100 });
  const groupId = deck.groupShapes(0, [a, b]);
  assert.ok(groupId);

  const group = shapeOf(deck, groupId);
  assert.equal(group.kind, 'group');
  assert.ok(near(group.geometry.x, 100) && near(group.geometry.y, 100));
  assert.ok(near(group.geometry.w, 400) && near(group.geometry.h, 250), `bbox is the union: ${JSON.stringify(group.geometry)}`);

  // The members carry the group's id, and their own absolute geometry is
  // exactly what it was before grouping — grouping alone moves nothing.
  // (Shape ids are strings on the scene, the way every shape's `id` is;
  // `groupShapes` hands back the same id as the number it minted it from.)
  assert.equal(shapeOf(deck, a).groupId, String(groupId));
  assert.equal(shapeOf(deck, b).groupId, String(groupId));
  assert.ok(near(shapeOf(deck, a).geometry.x, 100) && near(shapeOf(deck, a).geometry.y, 100));
  assert.ok(near(shapeOf(deck, b).geometry.x, 400) && near(shapeOf(deck, b).geometry.y, 250));

  const xml = OoxmlPackage.read(deck.save()).text(deck.slideParts[0].part);
  const grpXml = new RegExp(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${groupId}"[\\s\\S]*?<\\/p:grpSpPr>`).exec(xml)[0];
  assert.match(grpXml, /<a:chOff x="\d+" y="\d+"\/><a:chExt cx="\d+" cy="\d+"\/>/);
  const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(grpXml);
  const chOff = /<a:chOff x="(\d+)" y="(\d+)"\/>/.exec(grpXml);
  assert.deepEqual([off[1], off[2]], [chOff[1], chOff[2]], 'chOff starts equal to off, so the first read scales 1:1');
});

test('grouping takes an existing group as a member, and a shape that is not at the top level refuses', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 0, y: 0, w: 100, h: 100 });
  const b = deck.addShape(0, { preset: 'rect', x: 200, y: 0, w: 100, h: 100 });
  const inner = deck.groupShapes(0, [a, b]);
  assert.throws(() => deck.alignShapes(0, [a], 'left'), /inside a group/);

  const c = deck.addShape(0, { preset: 'rect', x: 500, y: 500, w: 60, h: 60 });
  const outer = deck.groupShapes(0, [inner, c]);
  assert.equal(shapeOf(deck, inner).groupId, String(outer), 'the inner group nests inside the outer one');
  assert.equal(shapeOf(deck, a).groupId, String(inner), 'a member of the inner group still names the inner one, not the outer');
});

test('moving or resizing a group carries its members with it, scaling them from the resize', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 100, y: 100, w: 100, h: 100 });
  const b = deck.addShape(0, { preset: 'rect', x: 300, y: 100, w: 100, h: 100 });
  const groupId = deck.groupShapes(0, [a, b]);
  const g0 = shapeOf(deck, groupId).geometry; // { x:100, y:100, w:300, h:100 }

  // A plain move: every member shifts by the same amount.
  deck.setGeometry(0, groupId, { x: g0.x + 50, y: g0.y + 20, w: g0.w, h: g0.h });
  assert.ok(near(shapeOf(deck, a).geometry.x, 150) && near(shapeOf(deck, a).geometry.y, 120));
  assert.ok(near(shapeOf(deck, b).geometry.x, 350) && near(shapeOf(deck, b).geometry.y, 120));

  // Doubling the group's width doubles each member's own width and its
  // offset from the group's own left edge — a stretch, not a shear.
  const g1 = shapeOf(deck, groupId).geometry;
  deck.setGeometry(0, groupId, { x: g1.x, y: g1.y, w: g1.w * 2, h: g1.h });
  assert.ok(near(shapeOf(deck, a).geometry.w, 200), `a stretched: ${shapeOf(deck, a).geometry.w}`);
  assert.ok(near(shapeOf(deck, b).geometry.w, 200), `b stretched: ${shapeOf(deck, b).geometry.w}`);
  assert.ok(near(shapeOf(deck, b).geometry.x, g1.x + 400), `b's offset from the group doubled too: ${shapeOf(deck, b).geometry.x}`);
});

test('ungroupShape puts members back at their true slide coordinates, matching where the group left them — round trip fidelity through a move, a resize, a rotation and a flip', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 100, y: 100, w: 100, h: 60 });
  const b = deck.addShape(0, { preset: 'ellipse', x: 300, y: 200, w: 80, h: 80 });
  const groupId = deck.groupShapes(0, [a, b]);

  // Move, resize, rotate and flip the group — everything ungroup has to undo correctly.
  const g0 = shapeOf(deck, groupId).geometry;
  deck.setGeometry(0, groupId, { x: g0.x + 40, y: g0.y - 20, w: g0.w * 1.5, h: g0.h * 0.8 });
  deck.setGeometry(0, groupId, { ...shapeOf(deck, groupId).geometry, flipH: true });
  const beforeUngroup = { a: { ...shapeOf(deck, a).geometry }, b: { ...shapeOf(deck, b).geometry } };

  const members = deck.ungroupShape(0, groupId);
  assert.deepEqual(members.map(String).sort(), [a, b].map(String).sort());
  assert.equal(shapeOf(deck, groupId), undefined, 'the group is gone');
  assert.equal(shapeOf(deck, a).groupId, null);

  for (const [id, before] of [[a, beforeUngroup.a], [b, beforeUngroup.b]]) {
    const after = shapeOf(deck, id).geometry;
    assert.ok(near(after.x, before.x, 1) && near(after.y, before.y, 1) && near(after.w, before.w, 1) && near(after.h, before.h, 1),
      `shape ${id}: ${JSON.stringify(before)} → ${JSON.stringify(after)}`);
  }

  const reopened = Deck.open(deck.save());
  assert.equal(reopened.slide(0).shapes.some((s) => s.kind === 'group'), false, 'no group in the saved file');
  const ra = reopened.slide(0).shapes.find((s) => String(s.id) === String(a));
  assert.ok(near(ra.geometry.x, beforeUngroup.a.x, 1), 'the ungrouped position is in the file');
});

test('a group renders as its members, not as an extra box of its own', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 100, y: 100, w: 200, h: 80 });
  const b = deck.addShape(0, { preset: 'rect', x: 400, y: 100, w: 100, h: 80 });
  deck.groupShapes(0, [a, b]);
  const svg = renderSlide(deck.slide(0), { width: 1280 });
  // Exactly the two members are drawn — the group's own frame draws nothing extra.
  assert.equal((svg.match(/<rect x="100"/g) || []).length, 1);
  assert.equal((svg.match(/<rect x="400"/g) || []).length, 1);
});

test('a flipped shape mirrors its outline on the stage without mirroring its words', () => {
  const deck = Deck.open(DECK);
  const a = deck.addShape(0, { preset: 'rect', x: 100, y: 100, w: 200, h: 80, text: 'Right way up' });
  deck.flipShapes(0, [a], 'horizontal');
  const svg = renderSlide(deck.slide(0), { width: 1280 });
  assert.match(svg, /scale\(-1 1\)/, 'the flipped shape mirrors its geometry');
  assert.match(svg, />Right way up</, 'its words are still drawn left-to-right, not mirrored');
});
