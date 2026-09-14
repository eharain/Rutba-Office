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
