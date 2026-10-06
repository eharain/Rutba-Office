// Insert → SmartArt — the layouts' arithmetic, and each engine's half.
//
// A diagram is laid out from lines of text (a Tab in for a line under the
// one above) as shapes: filled nodes with their words, and the arrows,
// ring and lines between them. Presentations puts them in as one group of
// preset shapes named for the layout; Documents as one Word group of
// shapes holding their words; Worksheets as one group on the sheet. Each
// is a group of shapes Office opens, every word editable — not a SmartArt
// part.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { SMARTART_LAYOUTS, parseItems, itemsText, layoutSmartArt, fitSize, boundsOf, smartArtSvg } from '../apps/desktop/renderer/smartart.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { renderPdf } from '@rutba/doc-view/export/pdf';
import { buildDocx, drawingAnchorXml } from '@rutba/ooxml/build';

const BOX = { x: 100, y: 120, w: 1080, h: 480 };
const inside = (s, box) => s.x >= box.x - 1 && s.y >= box.y - 1 && s.x + s.w <= box.x + box.w + 1 && s.y + s.h <= box.y + box.h + 1;

test('lines of text are items, a Tab or two spaces in a level under the one above, and back again', () => {
  const items = parseItems('Plan\n\tResearch\n  Budget\n\n• Build\n\t\t\tDeep');
  assert.deepEqual(items, [
    { text: 'Plan', level: 0 }, { text: 'Research', level: 1 }, { text: 'Budget', level: 1 },
    { text: 'Build', level: 0 }, { text: 'Deep', level: 2 },
  ]);
  assert.equal(itemsText(items), 'Plan\n\tResearch\n\tBudget\nBuild\n\t\tDeep');
});

test('every layout lays its items out inside the box, a node for each, with its words at a size that fits', () => {
  const items = parseItems('Plan\nBuild\nTest\nShip');
  for (const { id } of SMARTART_LAYOUTS) {
    const shapes = layoutSmartArt(id, items, BOX);
    const nodes = shapes.filter((s) => s.role === 'node');
    assert.equal(nodes.length, 4, id);
    assert.deepEqual(nodes.map((n) => n.text), ['Plan', 'Build', 'Test', 'Ship'], id);
    for (const s of shapes) assert.ok(inside(s, BOX), `${id}: ${JSON.stringify(s)} outside the box`);
    for (const n of nodes) assert.ok(n.size >= 9 && n.size <= 32 && n.scheme === 'accent1' && n.line === '#FFFFFF', id);
  }
  const process = layoutSmartArt('process', items, BOX);
  assert.equal(process.filter((s) => s.role === 'arrow' && s.preset === 'rightArrow').length, 3, 'an arrow between each step');
  const cycle = layoutSmartArt('cycle', items, BOX);
  assert.equal(cycle[0].role, 'ring', 'the ring behind the circles');
  assert.ok(cycle.filter((s) => s.role === 'node').every((s) => s.preset === 'ellipse' && s.w === s.h));
  assert.ok(layoutSmartArt('chevron', items, BOX).filter((s) => s.role === 'node').every((s) => s.preset === 'chevron'));
  // Nodes in a row do not overlap one another (a chevron's nose sits in the next one's notch).
  const row = layoutSmartArt('process', items, BOX).filter((s) => s.role === 'node');
  for (let i = 1; i < row.length; i++) assert.ok(row[i].x >= row[i - 1].x + row[i - 1].w, 'process boxes apart');
});

test('a hierarchy puts each top item over the items under it, a line down, across and down to each', () => {
  const shapes = layoutSmartArt('hierarchy', parseItems('Lead\n\tFirst\n\tSecond\n\tThird\nOther'), BOX);
  const nodes = shapes.filter((s) => s.role === 'node');
  const lines = shapes.filter((s) => s.role === 'line');
  assert.deepEqual(nodes.map((n) => n.text), ['Lead', 'First', 'Second', 'Third', 'Other']);
  const [lead, first, , third, other] = nodes;
  assert.ok(lead.y + lead.h <= first.y, 'the lead above its items');
  assert.equal(other.y, lead.y, 'a top item with nothing under it on the top row');
  assert.ok(Math.abs(lead.x + lead.w / 2 - (first.x + third.x + third.w) / 2) <= 1, 'the lead centred over its items');
  assert.equal(lines.length, 5, 'one down, one across, three down');
  assert.ok(lines.every((l) => l.w === 0 || l.h === 0), 'straight lines');
  assert.equal(lines[1].w, Math.round(third.x + third.w / 2) - Math.round(first.x + first.w / 2));
  // Words that fit: long words at a smaller size, a small box a small size.
  assert.ok(fitSize('Internationalisation', 200, 100) < fitSize('Plan', 200, 100));
  assert.equal(fitSize('Plan', 400, 300), 32);
  assert.deepEqual(boundsOf(nodes.slice(0, 1)), { x: lead.x, y: lead.y, w: lead.w, h: lead.h });
  assert.match(smartArtSvg('hierarchy', [], {}), /^<svg[\s\S]*<line[\s\S]*<rect[\s\S]*<\/svg>$/);
});

test('Presentations puts a diagram in as one group of preset shapes named for its layout, the words white in the theme colour', () => {
  const deck = Deck.open(buildPptx({ title: 'SmartArt', slides: [{ layout: 'blank' }] }));
  const shapes = layoutSmartArt('hierarchy', parseItems('Lead\n\tFirst\n\tSecond'), BOX);
  const group = deck.addDiagram(0, { name: 'Hierarchy', shapes });
  const xml = deck.pkg.text('ppt/slides/slide1.xml');
  assert.match(xml, new RegExp(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${group}" name="Hierarchy ${group}"/>`));
  assert.equal((xml.match(/<a:prstGeom prst="roundRect">/g) || []).length, 3);
  assert.equal((xml.match(/<a:prstGeom prst="line">/g) || []).length, 4);
  assert.match(xml, /<a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"\/>/);
  assert.match(xml, /<a:rPr[^>]*sz="\d+00"[^>]*>(?:(?!<\/a:rPr>).)*<a:srgbClr val="FFFFFF"\/>(?:(?!<\/a:rPr>).)*<\/a:rPr><a:t xml:space="preserve">Lead<\/a:t>/);
  // A line straight down is a line no wider than nothing.
  assert.match(xml, /<a:ext cx="1" cy="\d+"\/><\/a:xfrm><a:prstGeom prst="line">/);
  const scene = deck.slide(0);
  const top = scene.shapes.find((s) => String(s.id) === String(group));
  assert.equal(top.kind, 'group');
  const members = scene.shapes.filter((s) => String(s.groupId) === String(group));
  assert.equal(members.length, 7);
  const words = (s) => (s.text?.paragraphs || []).map((p) => p.runs.map((r) => r.text).join('')).join('');
  assert.deepEqual(members.map(words).filter(Boolean), ['Lead', 'First', 'Second']);
  const svg = renderSlide(scene, { width: 1280 });
  assert.ok(['Lead', 'First', 'Second'].every((w) => svg.includes(`>${w}<`)), 'drawn with its words');
  assert.throws(() => deck.addShape(0, { preset: 'rect', w: 0, h: 10 }), /positive width/);
});

test('the service puts a diagram in by the addDiagram op, its group id back, in one step of undo', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'deck' });
  const shapes = layoutSmartArt('process', parseItems('One\nTwo\nThree'), BOX);
  const before = docs.model({ id, slide: 0 }).slide.shapes.length;
  const out = docs.apply({ id, ops: [{ op: 'addDiagram', slide: 0, name: 'Basic Process', shapes }] });
  const made = out.model.slide.shapes.find((s) => String(s.id) === String(out.opResult));
  assert.equal(made.kind, 'group');
  assert.equal(made.name, `Basic Process ${out.opResult}`);
  assert.equal(out.model.slide.shapes.length, before + 1 + shapes.length);
  docs.undo({ id, slide: 0 });
  assert.equal(docs.model({ id, slide: 0 }).slide.shapes.length, before);
});

test('Documents puts a diagram in as one Word group of shapes holding their words, in a paragraph of its own, wrapped top and bottom', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'The stages follow.' }, { text: 'The end.' }] }));
  view.setSelection({ block: 0, offset: 0 });
  const shapes = layoutSmartArt('chevron', parseItems('Plan\nBuild\nShip'), { x: 0, y: 0, w: 600, h: 336 });
  view.insertDiagram({ name: 'Basic Chevron Process', shapes });
  const id = view.lastDrawing;
  const xml = view.doc.doc.xml;
  assert.match(xml, new RegExp(`<wp:docPr id="${id}" name="Basic Chevron Process ${id}"/>`));
  assert.match(xml, /<wp:positionH relativeFrom="column"><wp:align>center<\/wp:align><\/wp:positionH>/);
  assert.match(xml, /<wp:wrapTopAndBottom\/>/);
  assert.equal((xml.match(/<a:prstGeom prst="chevron">/g) || []).length, 3);
  assert.match(xml, /<wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"\/><w:jc w:val="center"\/>/);
  // The group's box is the shapes' own: no band above a layout centred in a taller box.
  const group = view.drawings().find((d) => d.id === id);
  assert.equal(group.kind, 'group');
  assert.ok(Math.abs(group.heightPx - shapes[0].h) <= 1, 'as tall as the chevrons');
  assert.match(xml, /<wps:spPr><a:xfrm><a:off x="0" y="0"\/>/);
  // Drawn: each chevron painted as its shape, its words over it as the edit space's own paragraphs.
  const block = view.render({ pages: false }).blocks.find((b) => b.groups);
  const members = block.groups[0].members;
  assert.deepEqual(members.map((m) => [m.kind, m.geom]), [['textbox', 'chevron'], ['textbox', 'chevron'], ['textbox', 'chevron']]);
  assert.ok(members.every((m) => /^data:image\/svg\+xml/.test(m.href) && m.blocks?.length === 1));
  assert.deepEqual(members.map((m) => view.block(m.blocks[0]).text), ['Plan', 'Build', 'Ship']);
  // Printed: the shapes and their words.
  const pdf = renderPdf(view, { created: '2026-10-07T00:00:00Z' }).buffer.toString('latin1');
  assert.ok(['Plan', 'Build', 'Ship'].every((w) => pdf.includes(`(${w})`)), 'the words printed');
  // Saved and read again, it is the same group.
  const again = openDocx(view.save());
  assert.equal(again.drawings().find((d) => d.kind === 'group')?.name, `Basic Chevron Process ${id}`);
});

test('a hierarchy\'s lines in a document are painted a stroke wider all round, so a line straight down still shows', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'doc' });
  const shapes = layoutSmartArt('hierarchy', parseItems('Lead\n\tA\n\tB'), { x: 0, y: 0, w: 600, h: 336 });
  const out = docs.apply({ id, ops: [{ op: 'insertDiagram', name: 'Hierarchy', shapes }] });
  assert.equal(typeof out.opResult, 'number');
  const group = docs.model({ id }).blocks.flatMap((b) => b.groups || []).find((g) => g.name === `Hierarchy ${out.opResult}`);
  const lines = group.members.filter((m) => m.kind === 'shape');
  assert.equal(lines.length, 4);
  const down = lines.find((m) => m.heightPx > m.widthPx);
  assert.equal(Math.round(down.widthPx), 6, 'a line no wider than nothing, painted 3 px either side');
  assert.match(down.href, /^data:image\/svg\+xml/);
  assert.throws(() => docs.apply({ id, ops: [{ op: 'insertDiagram', name: 'Hierarchy', shapes: [] }] }), /at least one shape/);
});

test('Worksheets puts a diagram in as one group of shapes at the selected cell, the words white in the theme colour', () => {
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id } = docs.new({ kind: 'sheet' });
  docs.apply({ id, ops: [{ op: 'select', row: 2, col: 2 }] });
  const shapes = layoutSmartArt('hierarchy', parseItems('Lead\n\tNorth\n\tSouth'), { x: 0, y: 0, w: 576, h: 324 });
  const out = docs.apply({ id, ops: [{ op: 'insertDiagram', name: 'Hierarchy', shapes }] });
  assert.equal(typeof out.opResult, 'number');
  const group = docs.model({ id }).drawings.find((d) => d.kind === 'group');
  assert.equal(group.name, `Hierarchy ${out.opResult}`);
  assert.deepEqual(group.anchor, { row: 2, col: 2 }, 'at the selected cell');
  assert.equal(group.members.length, 7);
  assert.ok(['Lead', 'North', 'South'].every((w) => group.members.some((m) => m.svg.includes(`>${w}<`))), 'every word drawn');
  // A line straight down is drawn in a box a stroke wider all round.
  const down = group.members.find((m) => m.height > 20 && m.width < 10);
  assert.equal(down.width, 7);
  assert.match(down.svg, /<line\b/);
});

test('the diagram writer puts the group\'s members in its own space, numbered on from its id', () => {
  const shapes = layoutSmartArt('process', parseItems('One\nTwo'), { x: 100, y: 50, w: 400, h: 200 });
  const xml = drawingAnchorXml({ kind: 'diagram', id: 5, name: 'Basic Process 5', shapes, from: { row: 1, col: 1 }, to: { row: 12, col: 8 } }, () => null);
  assert.match(xml, /^<xdr:twoCellAnchor editAs="oneCell"><xdr:from>/);
  assert.match(xml, /<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="5" name="Basic Process 5"\/><xdr:cNvGrpSpPr\/><\/xdr:nvGrpSpPr>/);
  assert.match(xml, /<a:xfrm><a:off x="952500" y="(\d+)"\/><a:ext cx="3810000" cy="(\d+)"\/><a:chOff x="952500" y="\1"\/><a:chExt cx="3810000" cy="\2"\/><\/a:xfrm>/);
  assert.deepEqual([...xml.matchAll(/<xdr:cNvPr id="(\d+)" name="([^"]+)"\/>/g)].map((m) => m[2]), ['Basic Process 5', 'Rectangle: Rounded Corners 6', 'Arrow: Right 7', 'Rectangle: Rounded Corners 8']);
  assert.match(xml, /<a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"\/><\/a:solidFill><\/a:ln>/);
  assert.match(xml, /<a:rPr lang="en-US" sz="\d+00"><a:solidFill><a:srgbClr val="FFFFFF"\/><\/a:solidFill><\/a:rPr><a:t>Two<\/a:t>/);
});
