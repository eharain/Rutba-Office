// Design → Slide Size, the engine's half: the slides' size written to
// p:sldSz, and every slide, layout and master scaled to suit as PowerPoint
// does it: Ensure Fit by the smaller ratio, Maximize by the larger, both
// centred, the words and a table's rows and columns scaled with the shapes.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';

const WIDE = buildPptx({ title: 'Size', slides: [{ layout: 'obj', title: 'One', body: ['a', 'b'] }, { layout: 'title', title: 'Two' }] });
const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 0.6, `${msg}: ${a} against ${b}`);
const shapeOf = (deck, i, name) => deck.slide(i).shapes.find((s) => s.name === name);
/** A slide with a sized box and a 24 pt run of its own, so its geometry and size are the slide's, not inherited. */
function withOwnBox(bytes) {
  const deck = Deck.open(bytes);
  const part = deck.slideParts[0].part;
  const xml = deck.pkg.text(part).replace('</p:spTree>',
    '<p:sp><p:nvSpPr><p:cNvPr id="9" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1219200" y="685800"/><a:ext cx="2438400" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:spcBef><a:spcPts val="1200"/></a:spcBef></a:pPr><a:r><a:rPr lang="en-US" sz="2400"/><a:t>Box</a:t></a:r></a:p></p:txBody></p:sp>' +
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="10" name="Grid"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="740000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid><a:tr h="370000"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>y</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree>');
  deck.pkg.write_(part, Buffer.from(xml, 'utf8'));
  return deck.save();
}

test('widescreen to 4:3 with Ensure Fit: the size is written, each shape scales by three quarters and the column of it is centred top to bottom', () => {
  const deck = Deck.open(withOwnBox(WIDE));
  const before = shapeOf(deck, 0, 'Box').geometry;
  assert.equal(deck.setSlideSize(9144000, 6858000, { fit: 'ensure' }), true);
  assert.deepEqual([deck.size.cx, deck.size.cy], [9144000, 6858000]);
  assert.match(deck.pkg.text('ppt/presentation.xml'), /<p:sldSz cx="9144000" cy="6858000" type="screen4x3"\/>/);
  const after = shapeOf(deck, 0, 'Box').geometry;
  near(after.w, before.w * 0.75, 'width');
  near(after.h, before.h * 0.75, 'height');
  near(after.x, before.x * 0.75, 'left');
  near(after.y, before.y * 0.75 + (720 - 720 * 0.75) / 2, 'top, centred');
  const xml = deck.pkg.text(deck.slideParts[0].part);
  assert.match(xml, /sz="1800"/, '24 pt becomes 18 pt');
  assert.match(xml, /<a:spcPts val="900"\/>/, 'the spacing before scales too');
  assert.match(xml, /<a:gridCol w="1500000"\/><a:gridCol w="1500000"\/>/, 'the table columns');
  assert.match(xml, /<a:tr h="277500">/, 'and its row');
  assert.match(xml, /<p:grpSpPr><a:xfrm><a:off x="0" y="0"\/><a:ext cx="0" cy="0"\/><a:chOff x="0" y="0"\/><a:chExt cx="0" cy="0"\/><\/a:xfrm><\/p:grpSpPr>/, 'the shape tree\'s own transform is untouched');
});

test('a placeholder that takes its place from the layout moves with the layout, and stays on the slide', () => {
  const deck = Deck.open(WIDE);
  deck.setSlideSize(9144000, 6858000);
  for (const name of ['Title 1', 'Content 2']) {
    const g = shapeOf(deck, 0, name).geometry;
    assert.ok(g.x >= 0 && g.y >= 0 && g.x + g.w <= deck.size.width + 0.5 && g.y + g.h <= deck.size.height + 0.5, `${name} inside the slide: ${JSON.stringify(g)}`);
  }
});

test('Maximize keeps the larger ratio: going to 4:3 the shapes keep their size and the sides run off, centred', () => {
  const deck = Deck.open(withOwnBox(WIDE));
  const before = shapeOf(deck, 0, 'Box').geometry;
  deck.setSlideSize(9144000, 6858000, { fit: 'maximize' });
  const after = shapeOf(deck, 0, 'Box').geometry;
  near(after.w, before.w, 'width kept');
  near(after.x, before.x - (1280 - 960) / 2, 'moved left by half the width lost');
  assert.match(deck.pkg.text(deck.slideParts[0].part), /sz="2400"/, 'the words keep their size');
});

test('4:3 to widescreen: nothing grows, everything is centred side to side, and the size survives a save', () => {
  const deck = Deck.open(withOwnBox(WIDE));
  deck.setSlideSize(9144000, 6858000);
  const small = shapeOf(deck, 0, 'Box').geometry;
  deck.setSlideSize(12192000, 6858000);
  const wide = shapeOf(deck, 0, 'Box').geometry;
  near(wide.w, small.w, 'same width');
  near(wide.x, small.x + 160, 'centred: half of the 320 px gained');
  assert.match(deck.pkg.text('ppt/presentation.xml'), /<p:sldSz cx="12192000" cy="6858000"\/>/, 'no 4:3 type on a wide slide');
  const again = Deck.open(deck.save());
  assert.deepEqual([again.size.cx, again.size.cy], [12192000, 6858000]);
});

test('the same size changes nothing, a size beyond PowerPoint\'s limits is refused, and one undo step puts it all back', () => {
  const deck = Deck.open(WIDE);
  assert.equal(deck.setSlideSize(12192000, 6858000), false);
  assert.throws(() => deck.setSlideSize(100, 6858000), /between 2\.54 cm and 142\.24 cm/);
  const snap = deck.snapshot();
  deck.setSlideSize(9144000, 6858000);
  deck.pushUndo(snap);
  assert.equal(deck.undo(), true);
  assert.deepEqual([deck.size.cx, deck.size.cy], [12192000, 6858000]);
  assert.doesNotMatch(deck.pkg.text('ppt/presentation.xml'), /screen4x3/);
});
