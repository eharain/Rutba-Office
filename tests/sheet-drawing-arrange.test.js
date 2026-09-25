// Page Layout → Arrange for a sheet's drawings: Bring Forward, Send
// Backward, to Front and to Back (the drawing part's order is the layering);
// Align and Distribute; Rotate and Flip on the drawing's xfrm (`rot`,
// `flipH`, `flipV`); Group into an `xdr:grpSp` and Ungroup; the Selection
// Pane's hide, show and rename; and moving or resizing by hand, the anchor
// following the cells.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView, anchorSpans } from '@rutba/sheet-view';
import { gradientPng } from '../apps/desktop/main/sample-picture.js';

const BOOK = () => buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Item', 'Qty'], ['Pens', 12], ['Ink', 3], ['Paper', 7]] }] });

/** Three shapes and a picture, in that drawing order. */
function withDrawings() {
  const view = new SheetView(BOOK());
  const shapes = [[1, 4, 'rect'], [3, 6, 'ellipse'], [6, 3, 'rightArrow']];
  for (const [row, col, geometry] of shapes) {
    view.select(row, col);
    view.insertShape({ geometry });
  }
  view.select(10, 1);
  view.insertPicture({ name: 'logo.png', contentType: 'image/png', data: gradientPng(40, 30, [60, 160, 90], [200, 50, 50]), widthPx: 80, heightPx: 60 });
  return view;
}
const drawn = (view) => view.render().drawings;
const byName = (view, name) => (view.drawings.get(view.activeSheet) ?? []).find((d) => d.name === name);
const drawingXml = (view) => view.pkg.text(view.pkg.partNames().find((p) => /^xl\/drawings\/drawing\d+\.xml$/.test(p)));

test('Bring Forward, Send Backward, to Front and to Back move the anchor in the drawing part; the ids stay', () => {
  const view = withDrawings();
  const order = () => (view.drawings.get(view.activeSheet) ?? []).map((d) => d.name);
  assert.deepEqual(order(), ['Shape 2', 'Shape 3', 'Shape 4', 'logo.png']);
  const first = byName(view, 'Shape 2');
  view.reorderDrawings({ ids: [first.id], to: 'forward' });
  assert.deepEqual(order(), ['Shape 3', 'Shape 2', 'Shape 4', 'logo.png']);
  view.reorderDrawings({ ids: [first.id], to: 'front' });
  assert.deepEqual(order(), ['Shape 3', 'Shape 4', 'logo.png', 'Shape 2']);
  assert.equal(byName(view, 'Shape 2').id, first.id, 'a drawing keeps its id when its place changes');
  view.reorderDrawings({ ids: [byName(view, 'logo.png').id], to: 'back' });
  assert.deepEqual(order(), ['logo.png', 'Shape 3', 'Shape 4', 'Shape 2']);
  view.reorderDrawings({ ids: [byName(view, 'Shape 4').id], to: 'backward' });
  assert.deepEqual(order(), ['logo.png', 'Shape 4', 'Shape 3', 'Shape 2']);
  const saved = SheetView.open(view.save());
  assert.deepEqual((saved.drawings.get('Sheet1') ?? []).map((d) => d.name), ['logo.png', 'Shape 4', 'Shape 3', 'Shape 2'], 'the order is the file\'s');
  view.undo();
  view.undo();
  assert.deepEqual(order(), ['Shape 3', 'Shape 4', 'logo.png', 'Shape 2']);
});

test('Align and Distribute line up the picked drawings by their combined bounds', () => {
  const view = withDrawings();
  const ids = ['Shape 2', 'Shape 3', 'Shape 4'].map((n) => byName(view, n).id);
  view.alignDrawings({ ids, edge: 'left' });
  const lefts = drawn(view).filter((d) => ids.includes(d.id)).map((d) => d.x);
  assert.equal(new Set(lefts).size, 1, 'one left edge');
  view.alignDrawings({ ids, edge: 'bottom' });
  const bottoms = drawn(view).filter((d) => ids.includes(d.id)).map((d) => d.y + d.height);
  assert.ok(Math.max(...bottoms) - Math.min(...bottoms) <= 1, 'one bottom edge');
  view.undo();
  view.undo();
  view.distributeDrawings({ ids, axis: 'vertical' });
  const boxes = drawn(view).filter((d) => ids.includes(d.id)).sort((a, b) => a.y - b.y);
  const gap1 = boxes[1].y - (boxes[0].y + boxes[0].height);
  const gap2 = boxes[2].y - (boxes[1].y + boxes[1].height);
  assert.ok(Math.abs(gap1 - gap2) <= 2, `equal gaps ${gap1} ${gap2}`);
  assert.throws(() => view.alignDrawings({ ids: [ids[0]], edge: 'left' }), /two or more/);
  assert.throws(() => view.distributeDrawings({ ids: ids.slice(0, 2) }), /three or more/);
});

test('Rotate and Flip write rot, flipH and flipV on the xfrm, and the drawing turns', () => {
  const view = withDrawings();
  const arrow = byName(view, 'Shape 4');
  view.rotateDrawings({ ids: [arrow.id], by: 90 });
  assert.match(drawingXml(view), /<xdr:spPr><a:xfrm rot="5400000"><a:off x="\d+" y="\d+"\/><a:ext cx="\d+" cy="\d+"\/><\/a:xfrm><a:prstGeom prst="rightArrow">/);
  view.rotateDrawings({ ids: [arrow.id], flip: 'horizontal' });
  assert.match(drawingXml(view), /<a:xfrm rot="5400000" flipH="1">/);
  const d = drawn(view).find((x) => x.id === arrow.id);
  assert.deepEqual([d.rot, d.flipH, d.flipV], [90, true, false]);
  assert.match(d.svg, /rotate\(90 /, 'the shape turns in its drawing');
  assert.match(d.svg, /scale\(-1 1\)/, 'and is mirrored');
  const pic = byName(view, 'logo.png');
  view.rotateDrawings({ ids: [pic.id], by: -90 });
  assert.match(drawingXml(view), /<xdr:pic>[\s\S]*<xdr:spPr><a:xfrm rot="16200000">/, 'a picture turns too');
  view.rotateDrawings({ ids: [arrow.id], to: 30 });
  assert.match(drawingXml(view), /<a:xfrm rot="1800000" flipH="1">/, 'the rotation handle sets an angle');
  view.undo();
  assert.match(drawingXml(view), /<a:xfrm rot="5400000" flipH="1">/);
});

test('Group gathers drawings into one xdr:grpSp at the topmost member\'s place; Ungroup puts them back where they were', () => {
  const view = withDrawings();
  const before = drawn(view);
  const ids = ['Shape 2', 'Shape 3'].map((n) => byName(view, n).id);
  const boxOf = (id, list) => { const d = list.find((x) => x.id === id); return [d.x, d.y, d.width, d.height]; };
  const gid = view.groupDrawings({ ids });
  const xml = drawingXml(view);
  assert.equal(anchorSpans(xml).length, 3, 'two anchors became one');
  assert.match(xml, /<xdr:twoCellAnchor><xdr:from>[\s\S]*?<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="\d+" name="Group \d+"\/><xdr:cNvGrpSpPr\/><\/xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="\d+" y="\d+"\/><a:ext cx="\d+" cy="\d+"\/><a:chOff x="\d+" y="\d+"\/><a:chExt cx="\d+" cy="\d+"\/><\/a:xfrm><\/xdr:grpSpPr><xdr:sp\b/);
  const group = drawn(view).find((d) => d.kind === 'group');
  assert.equal(group.id, gid);
  assert.equal(group.members.length, 2, 'drawn member by member');
  assert.ok(group.members.every((m) => /<svg/.test(m.svg)));
  // Members drawn where they were: the group's box is the union.
  const b2 = boxOf(ids[0], before);
  assert.ok(Math.abs(group.x + group.members[0].x - b2[0]) <= 1 && Math.abs(group.y + group.members[0].y - b2[1]) <= 1);
  // Moved as one, then ungrouped: each member where the group took it.
  view.setDrawingBox({ id: gid, x: group.x + 40, y: group.y + 20, width: group.width, height: group.height });
  const back = view.ungroupDrawings({ ids: [gid] });
  assert.equal(back, 1);
  const after = drawn(view);
  assert.equal(after.filter((d) => d.kind === 'group').length, 0);
  const moved = boxOf(ids[0], after);
  assert.ok(Math.abs(moved[0] - (b2[0] + 40)) <= 2 && Math.abs(moved[1] - (b2[1] + 20)) <= 2, `member moved with its group: ${moved} from ${b2}`);
  // The file reads back the same.
  const saved = SheetView.open(view.save());
  assert.equal(saved.render().drawings.length, 4);
  assert.throws(() => view.groupDrawings({ ids: [ids[0]] }), /two or more/);
});

test('the Selection Pane hides, shows and renames; a hidden drawing is not drawn; a hand drag moves and resizes', () => {
  const view = withDrawings();
  const frame = view.render();
  assert.deepEqual(frame.objects.map((o) => o.name), ['Shape 2', 'Shape 3', 'Shape 4', 'logo.png']);
  const shape = byName(view, 'Shape 3');
  view.setDrawingHidden({ id: shape.id, hidden: true });
  assert.match(drawingXml(view), /<xdr:cNvPr id="\d+" name="Shape 3" hidden="1"\/>/);
  assert.equal(view.render().objects.find((o) => o.id === shape.id).hidden, true);
  assert.equal(view.render().drawings.find((d) => d.id === shape.id).hidden, true);
  view.renameDrawing({ id: shape.id, name: 'Callout' });
  assert.equal(view.render().objects.find((o) => o.id === shape.id).name, 'Callout');
  view.setAllDrawingsHidden({ hidden: false });
  assert.ok(view.render().objects.every((o) => !o.hidden), 'Show All');
  const pic = drawn(view).find((d) => d.kind === 'image');
  view.setDrawingBox({ id: pic.id, x: pic.x + 64, y: pic.y + 20, width: 120, height: 90 });
  const moved = drawn(view).find((d) => d.id === pic.id);
  assert.deepEqual([moved.x, moved.y, moved.width, moved.height], [pic.x + 64, pic.y + 20, 120, 90]);
  assert.match(drawingXml(view), /<xdr:oneCellAnchor><xdr:from><xdr:col>\d+<\/xdr:col>[\s\S]*?<xdr:ext cx="1143000" cy="857250"\/>/, 'a one-cell anchor keeps its extent');
  const two = ['Shape 2', 'Shape 4'].map((n) => drawn(view).find((d) => d.name === n));
  view.setDrawingBoxes({ boxes: two.map((d) => ({ id: d.id, x: d.x + 10, y: d.y + 10, width: d.width, height: d.height })) });
  const after = two.map((d) => drawn(view).find((x) => x.id === d.id));
  assert.deepEqual(after.map((d) => d.x), two.map((d) => d.x + 10), 'several move as one step');
  view.undo();
  assert.deepEqual(drawn(view).filter((d) => two.some((t) => t.id === d.id)).map((d) => d.x), two.map((d) => d.x));
});
