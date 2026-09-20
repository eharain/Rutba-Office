// Floating layout — the engine's half.
//
// A picture in a `wp:anchor` says where it floats and how the text treats
// it; the reader hands that to the page, the writer moves a picture between
// inline and floating, and both survive a save and a reopen.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { anchorLayout } from '@rutba/ooxml/document';
import { buildCoverPageDocument } from './fixtures/cover-page-document.js';
import { OoxmlPackage } from '@rutba/ooxml';

/**
 * A document whose first paragraph anchors a text box floating at the right,
 * wrapped square — a pull quote beside the words — built from the fixture
 * builder's body with the box's run spliced in after the paragraph's
 * properties and the drawing namespaces declared on the root.
 */
function withFloatingBox(bytes, { side = 'right', cx = 2743200, cy = 914400 } = {}) {
  const pkg = OoxmlPackage.read(bytes);
  const run =
    '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
      '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
        '<wp:simplePos x="0" y="0"/>' +
        `<wp:positionH relativeFrom="margin"><wp:align>${side}</wp:align></wp:positionH>` +
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
        `<wp:extent cx="${cx}" cy="${cy}"/>` +
        '<wp:wrapSquare wrapText="bothSides"/>' +
        '<wp:docPr id="7" name="Pull Quote"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
          '<wps:wsp><wps:cNvSpPr txBox="1"/>' +
            `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
              '<a:solidFill><a:srgbClr val="EEF3F8"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="1F2123"/></a:solidFill></a:ln>' +
            '</wps:spPr>' +
            '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Growth held through the autumn.</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
          '</wps:wsp>' +
        '</a:graphicData></a:graphic>' +
      '</wp:anchor>' +
    '</w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect style="width:216pt;height:72pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Growth held through the autumn.</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r>';
  const ns =
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
    ' xmlns:v="urn:schemas-microsoft-com:vml"';
  const xml = pkg.text('word/document.xml')
    .replace('<w:document ', '<w:document' + ns + ' ')
    .replace(/(<w:p>(?:<w:pPr>[\s\S]*?<\/w:pPr>)?)/, '$1' + run);
  pkg.write_('word/document.xml', Buffer.from(xml, 'utf8'));
  return pkg.write();
}

const here = path.dirname(fileURLToPath(import.meta.url));
const PNG = fs.readFileSync(path.join(here, 'fixtures', 'rich', 'picture.png'));

test('anchorLayout reads the wrap, the position and the distances of an anchor', () => {
  const inner =
    '<wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>95250</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="1905000" cy="1428750"/><wp:wrapSquare wrapText="left"/>' +
    '<wp:docPr id="3" name="Picture 3"/></wp:anchor>';
  const l = anchorLayout(inner);
  assert.equal(l.anchored, true);
  assert.equal(l.wrap, 'square');
  assert.equal(l.wrapSide, 'left');
  assert.equal(l.hAlign, 'right');
  assert.equal(l.hRel, 'margin');
  assert.equal(l.vRel, 'paragraph');
  assert.equal(l.vOffsetPx, 10);
  assert.deepEqual(l.dist, { l: 12, r: 12, t: 4.8, b: 4.8 });
  assert.equal(l.behind, false);
  assert.deepEqual(anchorLayout('<wp:inline><wp:extent cx="1" cy="1"/></wp:inline>'), { anchored: false });
  assert.equal(anchorLayout('<wp:anchor behindDoc="1"><wp:wrapNone/></wp:anchor>').behind, true);
});

test('the cover page fixture reads as a centred text box with text above and below', () => {
  const view = openDocx(buildCoverPageDocument());
  const boxes = view.render({ pages: false }).blocks.flatMap((b) => b.textBoxes || []);
  const floating = boxes.find((i) => i.anchored);
  assert.ok(floating, 'the cover page has an anchored text box');
  assert.equal(floating.wrap, 'topAndBottom');
  assert.equal(floating.hAlign, 'center');
  assert.equal(floating.vRel, 'page');
  assert.equal(floating.vOffsetPx, 192);
});

test('setImageLayout floats a picture, moves it, and puts it back in the line', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'A logo beside these words.' }, { text: 'After.' }] }));
  view.setSelection({ block: 0, offset: 0 });
  view.insertImage({ name: 'logo', contentType: 'image/png', data: PNG, widthPx: 120, heightPx: 80 });
  const pictureBlock = view.render({ pages: false }).blocks.findIndex((b) => (b.images || []).length);
  assert.ok(pictureBlock >= 0, 'the picture is in a paragraph');
  const image = () => view.render({ pages: false }).blocks[pictureBlock].images[0];
  assert.equal(image().anchored, false, 'inserted in the line');

  view.setImageLayout({ block: pictureBlock, image: 0, wrap: 'square', hAlign: 'right' });
  assert.equal(image().anchored, true);
  assert.equal(image().wrap, 'square');
  assert.equal(image().hAlign, 'right');
  assert.equal(image().widthPx, 120, 'the size is untouched');
  assert.equal(image().href.slice(0, 15), 'data:image/png;', 'the picture is untouched');

  view.setImageLayout({ block: pictureBlock, image: 0, wrap: 'topAndBottom', hAlign: 'center' });
  assert.equal(image().wrap, 'topAndBottom');
  assert.equal(image().hAlign, 'center');

  view.setImageLayout({ block: pictureBlock, image: 0, wrap: 'behind' });
  assert.equal(image().wrap, 'none');
  assert.equal(image().behind, true);

  const reopened = openDocx(view.save());
  const again = reopened.render({ pages: false }).blocks[pictureBlock].images[0];
  assert.equal(again.behind, true, 'in the file');

  reopened.setImageLayout({ block: pictureBlock, image: 0, wrap: 'inline' });
  const back = reopened.render({ pages: false }).blocks[pictureBlock].images[0];
  assert.equal(back.anchored, false, 'back in the line');
  assert.equal(back.widthPx, 120);
  assert.throws(() => reopened.setImageLayout({ block: pictureBlock, image: 4, wrap: 'square' }), /no picture 4/);
  assert.equal(reopened.canUndo, true, 'an undoable edit');
});

test('setImageSize writes the extent and the transform, and nothing else moves', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Sized.' }] }));
  view.setSelection({ block: 0, offset: 0 });
  view.insertImage({ name: 'logo', contentType: 'image/png', data: PNG, widthPx: 120, heightPx: 80 });
  view.setImageLayout({ block: 1, image: 0, wrap: 'square', hAlign: 'right' });
  view.setImageSize({ block: 1, image: 0, widthPx: 240, heightPx: 160 });
  const image = view.render({ pages: false }).blocks[1].images[0];
  assert.equal(image.widthPx, 240);
  assert.equal(image.heightPx, 160);
  assert.equal(image.wrap, 'square', 'the layout is untouched');
  assert.equal(image.hAlign, 'right');
  const reopened = openDocx(view.save()).render({ pages: false }).blocks[1].images[0];
  assert.equal(reopened.widthPx, 240, 'in the file');
  assert.throws(() => view.setImageSize({ block: 1, image: 0, widthPx: 0, heightPx: 10 }), /positive/);
});

test('a floating picture prints beside the words: the lines next to it are shorter, and it is drawn at its side', async () => {
  const { renderPdf } = await import('@rutba/doc-view/export/pdf');
  const lorem = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn. ';
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Before.' }, { text: lorem.repeat(8) }, { text: 'After.' }] }));
  view.setSelection({ block: 0, offset: 0 });
  view.insertImage({ name: 'logo', contentType: 'image/png', data: PNG, widthPx: 160, heightPx: 160 });
  const pictureBlock = view.render({ pages: false }).blocks.findIndex((b) => (b.images || []).length);
  view.setImageLayout({ block: pictureBlock, image: 0, wrap: 'square', hAlign: 'right' });

  const laid = view.pages;
  const first = laid.pages[0];
  const width = laid.contentWidthPx;
  const float = first.fragments.find((f) => f.kind === 'float');
  assert.ok(float, 'the picture is placed as a float on the page');
  assert.equal(float.side, 'right');
  assert.equal(float.widthPx, 160);
  assert.ok(first.fragments.every((f) => f.kind !== 'images'), 'it is not also drawn as a block under the words');
  // The picture goes into a paragraph of its own; the words that wrap round
  // it are the paragraphs after it, for as far down as it reaches.
  const paragraphs = first.fragments.filter((f) => f.kind === 'paragraph' && f.paragraphIndex > pictureBlock);
  const lines = paragraphs.flatMap((f) => f.lines.map((l) => ({ ...l, lineHeightPx: f.lineHeightPx })));
  const beside = lines.filter((l) => l.widthPx != null && l.widthPx < width - 160);
  const clear = lines.filter((l) => l.widthPx == null || l.widthPx >= width - 1);
  assert.ok(beside.length >= 3, `the lines beside the picture are shorter: ${beside.length} of ${lines.length}`);
  assert.ok(clear.length >= 1, `the lines below the picture run the full column: ${clear.length}`);
  const lastBeside = lines.lastIndexOf(beside[beside.length - 1]);
  const firstClear = lines.indexOf(clear[0]);
  assert.ok(firstClear > lastBeside, 'the shortened lines come first, the full ones after');
  assert.ok(beside.every((l) => (l.offsetPx || 0) === 0), 'beside a right float the lines still start at the margin');
  assert.ok(beside.every((l) => l.width <= l.widthPx), 'no line runs into the picture');
  const cleared = beside.reduce((h, l) => h + l.lineHeightPx, 0);
  assert.ok(cleared >= 160 + 6 - 2 * beside[0].lineHeightPx, `the picture's whole height is kept clear: ${cleared}px of lines`);

  // A left float: the lines beside it start further in.
  view.setImageLayout({ block: pictureBlock, image: 0, wrap: 'square', hAlign: 'left' });
  const leftPage = view.pages.pages[0];
  const leftLines = leftPage.fragments.filter((f) => f.kind === 'paragraph' && f.paragraphIndex > pictureBlock).flatMap((f) => f.lines);
  assert.equal(leftPage.fragments.find((f) => f.kind === 'float').side, 'left');
  assert.ok(leftLines.slice(0, 3).every((l) => l.offsetPx >= 160), `beside a left float the lines start past it: ${leftLines.map((l) => l.offsetPx).slice(0, 3)}`);

  // Centred, top-and-bottom: a block under the words, centred.
  view.setImageLayout({ block: pictureBlock, image: 0, wrap: 'topAndBottom', hAlign: 'center' });
  const centred = view.pages.pages[0].fragments.find((f) => f.kind === 'images');
  assert.ok(centred && centred.images[0].hAlign === 'center', 'a centred picture keeps its side as a block');

  view.setImageLayout({ block: pictureBlock, image: 0, wrap: 'square', hAlign: 'right' });
  const { buffer, pages } = renderPdf(view, { created: '2026-09-20T00:00:00Z' });
  assert.equal(pages, 1);
  const pdf = buffer.toString('latin1');
  assert.ok(pdf.includes('/Subtype /Image'), 'the picture is embedded in the PDF');
  assert.ok(/northern/.test(pdf), 'the words are printed');
});

test('a floating text box prints beside the words: the lines next to it are shorter, and it is drawn at its side', async () => {
  const { renderPdf } = await import('@rutba/doc-view/export/pdf');
  const lorem = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn. ';
  const bytes = withFloatingBox(buildDocx({ styles: true, paragraphs: [{ text: lorem.repeat(8) }, { text: 'After.' }] }));
  const view = openDocx(bytes);
  const box = view.render({ pages: false }).blocks[0].textBoxes?.[0];
  assert.ok(box, 'the paragraph carries the box');
  assert.equal(box.anchored, true);
  assert.equal(box.wrap, 'square');
  assert.equal(box.hAlign, 'right');

  const laid = view.pages;
  const first = laid.pages[0];
  const width = laid.contentWidthPx;
  const float = first.fragments.find((f) => f.kind === 'floatbox');
  assert.ok(float, 'the box is placed as a float on the page');
  assert.equal(float.side, 'right');
  assert.equal(float.widthPx, 288, 'three inches wide, as the file says');
  assert.ok(float.heightPx >= 96, 'at least as tall as the file says');
  assert.equal(float.paragraphs.length, 1);
  assert.ok(first.fragments.every((f) => f.kind !== 'textbox'), 'it is not also drawn as a block under the words');
  const lines = first.fragments.filter((f) => f.kind === 'paragraph' && f.paragraphIndex === 0).flatMap((f) => f.lines);
  const beside = lines.filter((l) => l.widthPx != null && l.widthPx < width - 288);
  const clear = lines.filter((l) => l.widthPx == null || l.widthPx >= width - 1);
  assert.ok(beside.length >= 3, `the lines beside the box are shorter: ${beside.length} of ${lines.length}`);
  assert.ok(clear.length >= 1, `the lines below it run the full column: ${clear.length}`);
  assert.ok(lines.indexOf(clear[0]) > lines.lastIndexOf(beside[beside.length - 1]), 'the shortened lines come first');
  assert.ok(beside.every((l) => l.width <= l.widthPx), 'no line runs into the box');

  // At the left, the lines beside it start past it.
  const leftView = openDocx(withFloatingBox(buildDocx({ styles: true, paragraphs: [{ text: lorem.repeat(8) }, { text: 'After.' }] }), { side: 'left' }));
  const leftPage = leftView.pages.pages[0];
  assert.equal(leftPage.fragments.find((f) => f.kind === 'floatbox').side, 'left');
  const leftLines = leftPage.fragments.filter((f) => f.kind === 'paragraph' && f.paragraphIndex === 0).flatMap((f) => f.lines);
  assert.ok(leftLines.slice(0, 3).every((l) => l.offsetPx >= 288), `beside a left box the lines start past it: ${leftLines.map((l) => l.offsetPx).slice(0, 3)}`);

  const { buffer, pages } = renderPdf(view, { created: '2026-09-20T00:00:00Z' });
  assert.equal(pages, 1);
  const pdf = buffer.toString('latin1');
  assert.ok(/northern/.test(pdf), 'the words are printed');
  assert.ok(/autumn/.test(pdf), 'and the box\'s words with them');
});
