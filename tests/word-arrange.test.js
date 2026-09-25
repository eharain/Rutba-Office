// Arrange for a document's drawings: behind and in front of the words, the
// order of the floating ones, Align and Distribute, Rotate and Flip, Group
// and Ungroup, and the Selection Pane's names and eye — each written as Word
// writes it and drawn (and printed) as it says.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { floatPlace } from '@rutba/doc-view/floats';
import { buildDocx } from '@rutba/ooxml';

const here = path.dirname(fileURLToPath(import.meta.url));
const PNG = fs.readFileSync(path.join(here, 'fixtures', 'rich', 'picture.png'));
const lorem = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn. ';

/** A page of words with a floating picture and two floating text boxes, all anchored in paragraph 1. */
function page() {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Top.' }, { text: lorem.repeat(6) }, { text: lorem.repeat(3) }] }));
  view.setSelection({ block: 0, offset: 0 });
  view.insertImage({ name: 'logo', contentType: 'image/png', data: PNG, widthPx: 120, heightPx: 80 });
  // The picture's paragraph is 1 now; the words are 2 and 3.
  const pic = view.drawings().find((d) => d.kind === 'picture').id;
  view.updateDrawings({ id: pic, wrap: 'square', h: { rel: 'column', offsetPx: 0 }, v: { rel: 'paragraph', offsetPx: 0 } });
  view.setSelection({ block: 2, offset: 0 });
  view.insertTextBox({ name: 'Box A', widthPx: 160, heightPx: 60, h: { rel: 'column', offsetPx: 300 }, v: { rel: 'paragraph', offsetPx: 0 }, paragraphs: [{ text: 'A' }] });
  view.setSelection({ block: 2, offset: 0 });
  view.insertTextBox({ name: 'Box B', widthPx: 160, heightPx: 60, h: { rel: 'column', offsetPx: 340 }, v: { rel: 'paragraph', offsetPx: 20 }, paragraphs: [{ text: 'B' }] });
  const ids = Object.fromEntries(view.drawings().map((d) => [d.name, d.id]));
  return { view, pic, a: ids['Box A'], b: ids['Box B'] };
}
const drawing = (view, id) => view.drawings().find((d) => d.id === id);
const xml = (view) => view.doc.doc.xml;

test('Behind Text and In Front of Text write wrapNone with behindDoc, as Word does', () => {
  const { view, pic, a } = page();
  view.updateDrawings([{ id: pic, wrap: 'behind' }, { id: a, wrap: 'front' }]);
  assert.equal(drawing(view, pic).wrap, 'none');
  assert.equal(drawing(view, pic).behind, true);
  assert.equal(drawing(view, a).behind, false);
  assert.match(xml(view), /<wp:anchor\b[^>]*behindDoc="1"[^>]*>[\s\S]*?<wp:wrapNone\/>[\s\S]*?name="logo"/);
  const reopened = openDocx(view.save());
  assert.equal(reopened.drawings().find((d) => d.id === pic).behind, true, 'in the file');
  // Back beside the words: square, and not behind.
  view.updateDrawings({ id: pic, wrap: 'square' });
  assert.equal(drawing(view, pic).wrap, 'square');
  assert.equal(drawing(view, pic).behind, false);
});

test('Bring Forward, Send Backward, Bring to Front and Send to Back reorder relativeHeight', () => {
  const { view, pic, a, b } = page();
  const order = () => view.drawings().filter((d) => d.anchored).sort((x, y) => x.relativeHeight - y.relativeHeight).map((d) => d.id);
  assert.deepEqual(order(), [pic, a, b], 'the order they went in');
  view.orderDrawings([pic], 'front');
  assert.deepEqual(order(), [a, b, pic]);
  view.orderDrawings([pic], 'backward');
  assert.deepEqual(order(), [a, pic, b]);
  view.orderDrawings([b], 'back');
  assert.deepEqual(order(), [b, a, pic]);
  view.orderDrawings([b], 'forward');
  assert.deepEqual(order(), [a, b, pic]);
  const heights = view.drawings().filter((d) => d.anchored).map((d) => d.relativeHeight);
  assert.equal(new Set(heights).size, 3, 'each its own height');
  assert.ok(heights.every((z) => z > 251658240 && z < 2 ** 32));
  // One undo takes one press back.
  view.undo();
  assert.deepEqual(order(), [b, a, pic]);
  // A drawing in the line has no order.
  view.updateDrawings({ id: pic, wrap: 'inline' });
  assert.throws(() => view.orderDrawings([pic], 'front'), /floating/);
});

test('drawings behind the words print under them and drawings in front print over them', () => {
  const { view, pic, a } = page();
  view.updateDrawings([{ id: pic, wrap: 'behind', h: { rel: 'column', offsetPx: 40 }, v: { rel: 'paragraph', offsetPx: 10 } }, { id: a, wrap: 'front' }]);
  const frags = view.pages.pages[0].fragments;
  const under = frags.find((f) => f.layer === 'behind');
  const over = frags.find((f) => f.layer === 'front');
  assert.equal(under.kind, 'overlay');
  assert.equal(under.xPx, 40);
  assert.equal(over.kind, 'overlaybox');
  assert.ok(!frags.some((f) => f.kind === 'images'), 'not drawn in the flow any more');
  // The words take no room for them: the paragraph after is not pushed down.
  const pdf = renderPdf(view, { created: '2026-09-25T00:00:00Z' }).buffer.toString('latin1');
  const stream = /stream\r?\n([\s\S]*?)endstream/.exec(pdf.slice(pdf.indexOf('/Contents')))?.[1] || pdf;
  const picture = stream.indexOf(' Do');
  const words = stream.indexOf('(Top.) Tj');
  const boxWords = stream.indexOf('(A) Tj');
  assert.ok(picture >= 0 && words >= 0 && boxWords >= 0, 'all three are printed');
  assert.ok(picture < words, 'the picture behind is drawn before the words');
  assert.ok(boxWords > stream.lastIndexOf('northern'), 'the box in front is drawn after them');
});

test('Align to the margin or the page writes Word\'s alignments; the page and the printout place them the same', () => {
  const { view, a } = page();
  view.updateDrawings({ id: a, h: { rel: 'margin', align: 'right' }, v: { rel: 'page', align: 'bottom' } });
  const d = drawing(view, a);
  assert.equal(d.hRel, 'margin');
  assert.equal(d.hAlign, 'right');
  assert.equal(d.vRel, 'page');
  assert.equal(d.vAlign, 'bottom');
  assert.match(xml(view), /<wp:positionH relativeFrom="margin"><wp:align>right<\/wp:align><\/wp:positionH><wp:positionV relativeFrom="page"><wp:align>bottom<\/wp:align><\/wp:positionV>/);
  const section = view.section;
  const g = { marginLeftPx: section.margins.left, marginRightPx: section.margins.right, marginTopPx: section.margins.top, marginBottomPx: section.margins.bottom, contentWidthPx: section.contentWidthPx, pageWidthPx: section.widthPx, pageHeightPx: section.heightPx, columnWidthPx: section.contentWidthPx };
  const place = floatPlace({ ...d, widthPx: 160, heightPx: 60 }, g);
  assert.equal(place.x, section.contentWidthPx - 160, 'against the right margin');
  assert.equal(place.yFrom, 'page');
  assert.equal(place.y, section.heightPx - 60, 'on the foot of the page');
  const box = view.pages.pages[0].fragments.find((f) => f.kind === 'floatbox' && f.paragraphs?.some((p) => (p.runs || []).some((r) => r.text === 'A')));
  assert.equal(box.xPx, place.x, 'printed where the page puts it');
  assert.equal(box.topPx, place.y - section.margins.top + (d.dist?.t || 0));
  // Centred on the page across, middle down.
  view.updateDrawings({ id: a, h: { rel: 'page', align: 'center' }, v: { rel: 'margin', align: 'center' } });
  const c = floatPlace({ ...drawing(view, a), widthPx: 160, heightPx: 60 }, g);
  assert.equal(c.x, (section.widthPx - 160) / 2 - section.margins.left);
});

test('Align Selected Objects and Distribute are one press: every drawing\'s place in one undo step', () => {
  const { view, pic, a, b } = page();
  view.updateDrawings([
    { id: pic, h: { rel: 'column', offsetPx: 20 }, v: { rel: 'paragraph', offsetPx: 0 } },
    { id: a, h: { rel: 'column', offsetPx: 20 }, v: { rel: 'paragraph', offsetPx: 0 } },
    { id: b, h: { rel: 'column', offsetPx: 20 }, v: { rel: 'paragraph', offsetPx: 0 } },
  ]);
  assert.deepEqual([pic, a, b].map((id) => drawing(view, id).hOffsetPx), [20, 20, 20]);
  view.undo();
  assert.deepEqual([pic, a, b].map((id) => drawing(view, id).hOffsetPx), [0, 300, 340], 'one undo puts all three back');
});

test('Rotate and Flip write the turn and the mirror on the drawing\'s own transform, and it prints turned', () => {
  const { view, pic, a } = page();
  view.updateDrawings([{ id: pic, rot: 90, flipH: true }, { id: a, rot: 270 }]);
  assert.equal(drawing(view, pic).rot, 90);
  assert.equal(drawing(view, pic).flipH, true);
  assert.equal(drawing(view, a).rot, 270);
  assert.match(xml(view), /<pic:spPr><a:xfrm rot="5400000" flipH="1">/);
  assert.match(xml(view), /<wps:spPr><a:xfrm rot="16200000">/);
  const image = view.render({ pages: false }).blocks[1].images[0];
  assert.equal(image.rot, 90, 'the page is told the turn');
  const pdf = renderPdf(view, { created: '2026-09-25T00:00:00Z' }).buffer.toString('latin1');
  assert.match(pdf, /q\n-?0 -?1 -?1 -?0 [\d.-]+ [\d.-]+ cm/, 'the picture is drawn through a turning matrix');
  view.updateDrawings({ id: pic, rot: 0, flipH: false });
  assert.doesNotMatch(xml(view), /<pic:spPr><a:xfrm rot=/);
});

test('Group gathers floating drawings into one wpg group as Word writes it; the box in it is still typed in', () => {
  const { view, a, b } = page();
  const blocksBefore = view.blocks.length;
  view.groupDrawings([a, b], { rects: [{ id: a, x: 300, y: 0, w: 160, h: 60 }, { id: b, x: 340, y: 20, w: 160, h: 60 }], place: { h: { rel: 'column', offsetPx: 300 }, v: { rel: 'paragraph', offsetPx: 0 } } });
  const groups = view.drawings().filter((d) => d.kind === 'group');
  assert.equal(groups.length, 1);
  const group = groups[0];
  assert.equal(group.widthPx, 200);
  assert.equal(group.heightPx, 80);
  assert.deepEqual(group.members.map((m) => m.name), ['Box A', 'Box B']);
  assert.equal(view.drawings().filter((d) => d.kind === 'textbox').length, 0, 'the boxes are in the group now');
  const x = xml(view);
  assert.match(x, /<a:graphicData uri="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr\/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"\/><a:ext cx="1905000" cy="762000"\/><a:chOff x="0" y="0"\/><a:chExt cx="1905000" cy="762000"\/><\/a:xfrm><\/wpg:grpSpPr><wps:wsp><wps:cNvPr id="\d+" name="Box A"\/>/);
  assert.match(x, /<wps:spPr><a:xfrm><a:off x="381000" y="190500"\/><a:ext cx="1524000" cy="571500"\/>/, 'Box B at its place in the group');
  // The boxes' words are still the edit space's.
  assert.equal(view.blocks.length, blocksBefore);
  const boxB = view.blocks.find((bl) => bl.box && bl.text === 'B');
  view.setSelection({ block: boxB.index, offset: 1 });
  view.insertText('ee');
  assert.equal(view.blocks[boxB.index].text, 'Bee');
  const frame = view.render({ pages: false });
  const drawn = frame.blocks.flatMap((bl) => bl.groups || [])[0];
  assert.equal(drawn.members.length, 2);
  assert.ok(drawn.members.every((m) => m.kind === 'textbox' && m.blocks?.length === 1), 'each member names its paragraphs');
  // It prints: the group's boxes and their words.
  const pdf = renderPdf(view, { created: '2026-09-25T00:00:00Z' }).buffer.toString('latin1');
  assert.ok(/\(Bee\) Tj/.test(pdf), 'the words of a box in a group are printed');
  // Survives a save.
  assert.equal(openDocx(view.save()).drawings().filter((d) => d.kind === 'group').length, 1);
});

test('Ungroup puts each member back as a drawing of its own, where it stood', () => {
  const { view, a, b } = page();
  view.groupDrawings([a, b], { rects: [{ id: a, x: 300, y: 0, w: 160, h: 60 }, { id: b, x: 340, y: 20, w: 160, h: 60 }], place: { h: { rel: 'column', offsetPx: 300 }, v: { rel: 'paragraph', offsetPx: 0 } } });
  const group = view.drawings().find((d) => d.kind === 'group');
  view.ungroupDrawing(group.id);
  const boxes = view.drawings().filter((d) => d.kind === 'textbox');
  assert.equal(boxes.length, 2);
  assert.deepEqual(boxes.map((d) => [d.name, Math.round(d.hOffsetPx), Math.round(d.vOffsetPx), Math.round(d.widthPx)]), [['Box A', 300, 0, 160], ['Box B', 340, 20, 160]]);
  assert.equal(new Set(view.drawings().map((d) => d.id)).size, view.drawings().length, 'every id still its own');
  assert.throws(() => view.ungroupDrawing(boxes[0].id), /not a group/);
  assert.throws(() => view.groupDrawings([boxes[0].id]), /two or more/);
});

test('the Selection Pane\'s eye hides a drawing from the page and the printout, and its name is written back', () => {
  const { view, a } = page();
  view.updateDrawings({ id: a, hidden: true, name: 'Quote' });
  const d = drawing(view, a);
  assert.equal(d.hidden, true);
  assert.equal(d.name, 'Quote');
  assert.match(xml(view), /<wp:docPr id="\d+" name="Quote" hidden="1"\/>/);
  const frags = view.pages.pages[0].fragments;
  assert.ok(!frags.some((f) => (f.paragraphs || []).some((p) => (p.runs || []).some((r) => r.text === 'A'))), 'not printed');
  view.updateDrawings({ id: a, hidden: false });
  assert.equal(drawing(view, a).hidden, false);
  assert.doesNotMatch(xml(view), /hidden="1"/);
});

test('a drawing dragged down the page takes its anchor with it to the paragraph it now stands by', () => {
  const { view, a } = page();
  assert.equal(drawing(view, a).block, 2);
  view.updateDrawings({ id: a, block: 3, h: { rel: 'column', offsetPx: 10 }, v: { rel: 'paragraph', offsetPx: 5 } });
  const d = drawing(view, a);
  assert.equal(d.block, 3);
  assert.equal(d.hOffsetPx, 10);
  assert.equal(view.blocks[2].text, lorem.repeat(6), 'the old anchor keeps its words');
  assert.equal(view.render({ pages: false }).blocks[3].textBoxes.length, 1, 'the new one holds the box');
  view.undo();
  assert.equal(drawing(view, a).block, 2, 'one undo moves it back');
});
