// Insert → Text Box: a floating box of words, written as Word 2010 and later
// write one, typed in place like the body, moved and sized and formatted,
// and laid out beside the words on paper as on screen.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';

const lorem = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn. ';

/** A document of three paragraphs with a box anchored in the second. */
function withBox(spec = {}) {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'First paragraph.' }, { text: lorem.repeat(6) }, { text: 'Last paragraph.' }] }));
  view.setSelection({ block: 1, offset: 0 });
  view.insertTextBox({ widthPx: 200, heightPx: 90, paragraphs: [{ text: 'Box words' }, { text: 'Second line' }], ...spec });
  return view;
}
const mainXml = (view) => view.doc.doc.xml;
const boxOf = (view) => view.render({ pages: false }).blocks[1].textBoxes[0];

test('Insert → Text Box writes the box as Word does: an anchor, a wps shape with its words, the VML twin', () => {
  const view = withBox();
  const xml = mainXml(view);
  const run = /<w:r><mc:AlternateContent>[\s\S]*?<\/mc:AlternateContent><\/w:r>/.exec(xml)?.[0];
  assert.ok(run, 'the box is a run holding markup compatibility content');
  assert.match(run, /<mc:Choice Requires="wps"><w:drawing><wp:anchor\b[^>]*behindDoc="0"/);
  assert.match(run, /<wp:positionH relativeFrom="column"><wp:posOffset>0<\/wp:posOffset><\/wp:positionH>/);
  assert.match(run, /<wp:positionV relativeFrom="paragraph"><wp:posOffset>0<\/wp:posOffset><\/wp:positionV>/);
  assert.match(run, /<wp:extent cx="1905000" cy="857250"\/>/, '200 × 90 px in EMU');
  assert.match(run, /<wp:wrapSquare wrapText="bothSides"\/>/, 'wrapped square by default');
  assert.match(run, /<wps:wsp><wps:cNvSpPr txBox="1"\/>/);
  assert.match(run, /<wps:txbx><w:txbxContent><w:p>[\s\S]*Box words[\s\S]*Second line[\s\S]*<\/w:txbxContent><\/wps:txbx>/);
  assert.match(run, /<wps:bodyPr\b[^>]*lIns="91440" tIns="45720" rIns="91440" bIns="45720"[^>]*anchor="t"/, 'Word\'s own margins');
  assert.match(run, /<mc:Fallback><w:pict><v:rect\b[^>]*style="position:absolute;[^"]*width:150pt;height:67\.5pt[^"]*"[\s\S]*<v:textbox\b[^>]*><w:txbxContent>[\s\S]*Box words[\s\S]*<\/v:textbox><w10:wrap type="square"\/>/, 'the VML twin, the same box');
  const root = /<w:document\b[^>]*>/.exec(xml)[0];
  for (const prefix of ['mc', 'wp', 'wps', 'a', 'v', 'w10']) assert.match(root, new RegExp(`xmlns:${prefix}=`), `${prefix} is declared on the document`);
  const box = boxOf(view);
  assert.equal(box.kind, 'textbox');
  assert.equal(box.wrap, 'square');
  assert.equal(box.widthPx, 200);
  assert.equal(box.fill, '#FFFFFF');
  assert.equal(box.line, '#000000');
  assert.deepEqual(box.insets, { l: 9.6, t: 4.8, r: 9.6, b: 4.8 });
});

test('a box\'s words are the edit space\'s, after the body — the body\'s paragraphs keep their places', () => {
  const view = withBox();
  const blocks = view.blocks;
  assert.deepEqual(blocks.slice(0, 3).map((b) => b.container), [null, null, null], 'the body is where it was');
  assert.equal(blocks[1].text, lorem.repeat(6), 'the anchor\'s words are its own');
  const inBox = blocks.filter((b) => b.box);
  assert.deepEqual(inBox.map((b) => b.text), ['Box words', 'Second line']);
  assert.deepEqual(inBox.map((b) => b.index), [3, 4]);
  assert.ok(inBox.every((b) => /^x\d+$/.test(b.container) && b.box.anchor === 1));
  assert.deepEqual(boxOf(view).blocks, [3, 4], 'the anchor names its box\'s paragraphs');
  assert.equal(view.render({ pages: false }).flow.filter((f) => f.kind === 'paragraph').length, 3, 'the flow draws the body only');
  assert.deepEqual(view.selection.from, { block: 3, offset: 0 }, 'the new box\'s words are selected, so typing replaces them');
  assert.deepEqual(view.selection.to, { block: 4, offset: 'Second line'.length });
});

test('typing in a box edits its words, and the VML twin an older Word reads follows', () => {
  const view = withBox();
  view.setSelection({ block: 3, offset: 3 });
  view.insertText(' of ours');
  assert.equal(view.blocks[3].text, 'Box of ours words');
  const fallback = /<mc:Fallback>([\s\S]*?)<\/mc:Fallback>/.exec(mainXml(view))[1];
  assert.match(fallback, /Box of ours words/, 'the twin has the new words');
  assert.doesNotMatch(fallback, /<w:t[^>]*>Box words</);
  // Bold, and Enter, inside the box.
  view.setSelection({ block: 3, offset: 0 }, { block: 3, offset: 3 });
  view.toggleFormat('b');
  assert.match(/<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/.exec(mainXml(view))[1], /<w:b\/>[\s\S]*?Box/);
  view.setSelection({ block: 4, offset: 6 });
  view.splitParagraph();
  assert.deepEqual(view.blocks.filter((b) => b.box).map((b) => b.text), ['Box of ours words', 'Second', ' line'], 'a new paragraph in the box');
  assert.deepEqual(view.focus, { block: 5, offset: 0 });
  // Backspace at the top of the box stops at its wall.
  view.setSelection({ block: 3, offset: 0 });
  view.deleteBackward();
  assert.equal(view.blocks[2].text, 'Last paragraph.', 'the body is untouched');
  assert.equal(view.blocks[3].text, 'Box of ours words');
  // Ctrl+A in a box takes the box's words.
  view.selectAll();
  assert.equal(view.selection.from.block, 3);
  assert.equal(view.selection.to.block, 5);
});

test('the paragraph that anchors a box is typed in like any other, and keeps its box', () => {
  const view = withBox();
  view.setSelection({ block: 1, offset: 0 });
  view.insertText('Typed. ');
  assert.equal(view.blocks[1].text.startsWith('Typed. The northern'), true);
  assert.equal(view.blocks[1].structural, false);
  assert.equal(boxOf(view).blocks.length, 2, 'the box is still there');
  // Enter in the anchor: the box stays with the first half.
  view.setSelection({ block: 1, offset: 6 });
  view.splitParagraph();
  const frame = view.render({ pages: false });
  assert.equal(frame.blocks[1].text, 'Typed.');
  assert.equal(frame.blocks[1].textBoxes.length, 1);
  assert.equal(frame.blocks[2].textBoxes, undefined);
  assert.deepEqual(view.focus, { block: 2, offset: 0 }, 'the caret goes to the new paragraph, not into the box');
  assert.equal((mainXml(view).match(/<wps:txbx>/g) || []).length, 1, 'one box, not two');
});

test('a box survives a save and a reopen: its place, size, look and words', () => {
  const view = withBox({ fill: 'EEF3F8', line: '1F3864', lineWidthPx: 2, insets: { l: 16, t: 12, r: 16, b: 12 }, vAnchor: 'middle', h: { rel: 'column', offsetPx: 300 }, v: { rel: 'paragraph', offsetPx: 24 } });
  const again = openDocx(view.save());
  const box = again.render({ pages: false }).blocks[1].textBoxes[0];
  assert.equal(box.hRel, 'column');
  assert.equal(box.hOffsetPx, 300);
  assert.equal(box.vOffsetPx, 24);
  assert.equal(box.fill, '#EEF3F8');
  assert.equal(box.line, '#1F3864');
  assert.equal(box.lineWidthPx, 2);
  assert.deepEqual(box.insets, { l: 16, t: 12, r: 16, b: 12 });
  assert.equal(box.vAnchor, 'middle');
  assert.deepEqual(again.blocks.filter((b) => b.box).map((b) => b.text), ['Box words', 'Second line']);
  // The package is one Word opens: the document part parses as XML the way it did.
  const pkg = OoxmlPackage.read(view.save());
  assert.ok(pkg.text('word/document.xml').includes('<wps:txbx>'));
});

test('Shape Format writes the fill, the outline, the margins, the alignment and the direction; a drag writes the place and size', () => {
  const view = withBox();
  const id = boxOf(view).id;
  view.updateDrawings({ id, fill: null, line: { colour: '2B5FD9', widthPx: 3 }, insets: { l: 0, t: 0, r: 0, b: 0 }, vAnchor: 'bottom', vert: 'vert270' });
  let box = boxOf(view);
  assert.equal(box.fill, null, 'no fill');
  assert.equal(box.line, '#2B5FD9');
  assert.equal(box.lineWidthPx, 3);
  assert.deepEqual(box.insets, { l: 0, t: 0, r: 0, b: 0 });
  assert.equal(box.vAnchor, 'bottom');
  assert.equal(box.vert, 'vert270');
  assert.match(mainXml(view), /<wps:spPr><a:xfrm>[\s\S]*?<a:noFill\/><a:ln w="28575"><a:solidFill><a:srgbClr val="2B5FD9"\/>/);
  // Moved and sized by hand: the offsets and the extent (and the shape's own transform).
  view.updateDrawings({ id, h: { rel: 'column', offsetPx: 120 }, v: { rel: 'paragraph', offsetPx: 40 }, widthPx: 260, heightPx: 120 });
  box = boxOf(view);
  assert.equal(box.hOffsetPx, 120);
  assert.equal(box.vOffsetPx, 40);
  assert.equal(box.widthPx, 260);
  assert.equal(box.heightPx, 120);
  assert.match(mainXml(view), /<wps:spPr><a:xfrm><a:off x="0" y="0"\/><a:ext cx="2476500" cy="1143000"\/>/);
  assert.match(/<mc:Fallback>([\s\S]*?)<\/mc:Fallback>/.exec(mainXml(view))[1], /margin-left:90pt;margin-top:30pt;width:195pt;height:90pt/, 'the twin moved too');
  // One undo takes the drag back.
  view.undo();
  assert.equal(boxOf(view).hOffsetPx, 0);
});

test('a text box prints where it stands, the words beside it shorter, and its words in it', async () => {
  const { renderPdf } = await import('@rutba/doc-view/export/pdf');
  const view = withBox({ h: { rel: 'column', offsetPx: 380 }, v: { rel: 'paragraph', offsetPx: 20 }, widthPx: 220, heightPx: 120 });
  const page = view.pages.pages[0];
  const box = page.fragments.find((f) => f.kind === 'floatbox');
  assert.ok(box, 'placed as a float');
  assert.equal(box.side, 'right');
  assert.equal(box.xPx, 380, 'across where the anchor says');
  assert.ok(box.paragraphs.length >= 2);
  const lines = page.fragments.filter((f) => f.kind === 'paragraph' && f.paragraphIndex === 1).flatMap((f) => f.lines);
  const width = view.pages.contentWidthPx;
  assert.ok(lines.some((l) => l.widthPx != null && l.widthPx <= 380), 'the lines beside it stop short of it');
  assert.ok(lines.some((l) => l.widthPx == null || l.widthPx >= width - 1), 'and the ones below run the full width');
  const pdf = renderPdf(view, { created: '2026-09-25T00:00:00Z' }).buffer.toString('latin1');
  assert.ok(/Box words/.test(pdf) && /Second line/.test(pdf), 'the box\'s words are printed');
});

test('a box taken out takes its words out of the edit space, and Undo brings it back', () => {
  const view = withBox();
  const id = boxOf(view).id;
  view.setSelection({ block: 3, offset: 2 });
  view.removeDrawing(id);
  assert.equal(view.blocks.length, 3, 'only the body is left');
  assert.equal(view.blocks[1].text, lorem.repeat(6), 'the anchor keeps its words');
  assert.ok(view.focus.block <= 2, 'the caret is back in the body');
  assert.doesNotMatch(mainXml(view), /<wps:txbx>|<mc:AlternateContent>/);
  view.undo();
  assert.equal(view.blocks.filter((b) => b.box).length, 2);
});
