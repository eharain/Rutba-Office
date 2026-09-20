// A picture in a worksheet, written as Excel keeps one.
//
// The bytes become a media part; the sheet's drawing part gets a one-cell
// anchor at the active cell pointing at them through a relationship, at the
// size asked for; the content types know the extension. The frame draws it
// at once over the cells, a reopened file draws it too, and an undo takes the
// picture and its parts away.

import test from 'node:test';
import assert from 'node:assert/strict';
import { buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { gradientPng } from '../apps/desktop/main/sample-picture.js';

const BOOK = () => buildXlsx({ sheets: [{ name: 'Sales', rows: [['Item', 'Qty'], ['Pens', 12], ['Ink', 3]] }] });
const PNG = () => gradientPng(40, 30, [60, 160, 90], [200, 50, 50]);

test('a picture goes in at the active cell: media part, drawing anchor, relationship, content type, and the frame draws it', () => {
  const view = new SheetView(BOOK());
  view.select(5, 3);
  view.insertPicture({ name: 'logo.png', contentType: 'image/png', data: PNG(), widthPx: 120, heightPx: 90 });

  // The frame names a picture drawing "image", the drawing reader's word for the anchor.
  const drawn = (view.render().drawings || []).filter((d) => d.kind === 'image');
  assert.equal(drawn.length, 1, 'the frame carries the picture at once');
  assert.deepEqual([drawn[0].width, drawn[0].height], [120, 90], 'at its size');
  assert.match(drawn[0].svg || '', /<image\b/, 'drawn as an image');

  const saved = SheetView.open(view.save());
  assert.ok(saved.pkg.has('xl/media/image1.png'), 'the media part');
  assert.equal(saved.pkg.contentTypes().defaults.get('png'), 'image/png');
  const sheetPart = saved.workbook.partNameFor('Sales');
  const drawingRel = saved.pkg.rels(sheetPart).find((r) => r.Type.endsWith('/drawing'));
  assert.ok(drawingRel, 'the sheet points at a drawing part');
  const drawingPart = 'xl/' + drawingRel.Target.replace(/^\.\.\//, '');
  const xml = saved.pkg.text(drawingPart);
  assert.match(xml, /<xdr:oneCellAnchor><xdr:from><xdr:col>3<\/xdr:col><xdr:colOff>0<\/xdr:colOff><xdr:row>5<\/xdr:row><xdr:rowOff>0<\/xdr:rowOff><\/xdr:from><xdr:ext cx="1143000" cy="857250"\/>/, 'anchored at D6 at 120×90 px');
  assert.match(xml, /<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="\d+" name="logo.png"\/>/);
  const imageRel = saved.pkg.rels(drawingPart).find((r) => r.Type.endsWith('/image'));
  assert.ok(imageRel, 'the drawing part points at the media');
  assert.equal(imageRel.Target, '../media/image1.png');
  assert.match(xml, new RegExp('<a:blip r:embed="' + imageRel.Id + '"/>'));
  assert.equal((saved.render().drawings || []).filter((d) => d.kind === 'image').length, 1, 'the reopened file draws it');
});

test('a second picture takes the next media name; an undo takes a picture and its part away; a bad type is refused', () => {
  const view = new SheetView(BOOK());
  view.select(0, 0);
  view.insertPicture({ name: 'one', contentType: 'image/png', data: PNG(), widthPx: 40, heightPx: 30 });
  view.select(4, 1);
  view.insertPicture({ name: 'two', contentType: 'image/png', data: PNG(), widthPx: 40, heightPx: 30 });
  assert.ok(view.pkg.has('xl/media/image1.png') && view.pkg.has('xl/media/image2.png'));
  assert.equal((view.render().drawings || []).filter((d) => d.kind === 'image').length, 2);

  view.undo();
  assert.equal((view.render().drawings || []).filter((d) => d.kind === 'image').length, 1, 'the second picture is gone');
  assert.ok(!view.pkg.has('xl/media/image2.png'), 'and its part with it');
  view.undo();
  assert.equal((view.render().drawings || []).filter((d) => d.kind === 'image').length, 0);

  assert.throws(() => view.insertPicture({ contentType: 'image/tiff', data: PNG(), widthPx: 10, heightPx: 10 }), /unsupported image type/);
  assert.throws(() => view.insertPicture({ contentType: 'image/png', data: PNG(), widthPx: 0, heightPx: 10 }), /positive/);
});
