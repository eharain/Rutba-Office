// Edit in Image (docs/STUDIO.md I16): a picture in a document, a worksheet
// or a slide handed out as its bytes, and an edited picture put back in its
// place: the same drawing, where it was, its width kept and its height from
// the new picture's shape; in the picture's own part when nothing else uses
// it, in a new one when a copy shares it; undone in one step; and kept by the
// file once saved and opened again.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { SheetView } from '@rutba/sheet-view';
import { Deck, buildPptx } from '@rutba/presentation';
import { gradientPng } from '../apps/desktop/main/sample-picture.js';

const RED = () => gradientPng(40, 20, [220, 40, 40], [200, 30, 30]);
const BLUE = () => gradientPng(30, 30, [30, 60, 220], [20, 40, 200]);
const same = (a, b) => Buffer.compare(Buffer.from(a), Buffer.from(b)) === 0;

test('I16 Documents: the picture handed out, the edited one put back in place, its height from its shape, undone in one step and kept on save', () => {
  const view = openDocx(buildDocx({ paragraphs: [{ text: 'Before' }, { text: '' }, { text: 'After' }] }));
  view.setSelection({ block: 1, offset: 0 });
  view.insertImage({ name: 'logo', contentType: 'image/png', data: RED(), widthPx: 160, heightPx: 80 });
  const block = view.render({ pages: false }).blocks.findIndex((b) => b.images?.length);
  const src = view.pictureSource(block, 0);
  assert.equal(src.contentType, 'image/png');
  assert.ok(same(src.data, RED()), 'the picture\'s own bytes come out');
  view.replacePicture({ block, image: 0, data: BLUE(), contentType: 'image/png', widthPx: 160, heightPx: 160 });
  const img = view.render({ pages: false }).blocks[block].images[0];
  assert.ok(same(view.pictureSource(block, 0).data, BLUE()), 'the edited picture is in');
  assert.deepEqual([Math.round(img.widthPx), Math.round(img.heightPx)], [160, 160], 'width kept, height from the square picture');
  assert.equal(view.render({ pages: false }).blocks[block].images.length, 1, 'the same one picture, not a second');
  view.undo();
  assert.ok(same(view.pictureSource(block, 0).data, RED()), 'Undo brings the old picture back');
  view.redo();
  const reopened = openDocx(view.save());
  assert.ok(same(reopened.pictureSource(block, 0).data, BLUE()), 'saved and opened again, the edited picture is there');
});

test('I16 Documents: a PNG replaced by a JPEG gets a part of its own, named as a JPEG', () => {
  const view = openDocx(buildDocx({ paragraphs: [{ text: '' }, { text: '' }] }));
  view.setSelection({ block: 0, offset: 0 });
  view.insertImage({ name: 'a', contentType: 'image/png', data: RED(), widthPx: 40, heightPx: 20 });
  const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 16, 0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0, 0xff, 0xd9]);
  const block = view.render({ pages: false }).blocks.findIndex((b) => b.images?.length);
  view.replacePicture({ block, image: 0, data: jpeg, contentType: 'image/jpeg' });
  const src = view.pictureSource(block, 0);
  assert.equal(src.contentType, 'image/jpeg');
  assert.match(src.name, /\.jpg$/);
});

test('I16 Worksheets: the picture handed out by its frame id, the edited one put back on its anchor, its height from its shape, one undo step', () => {
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Sales', rows: [['Item', 'Qty'], ['Pens', 12]] }] }));
  view.select(3, 2);
  view.insertPicture({ name: 'Logo', contentType: 'image/png', data: RED(), widthPx: 120, heightPx: 60 });
  const pic = (view.render().drawings || []).find((d) => d.kind === 'image');
  view.undo();
  view.redo();
  const src = view.pictureSource({ id: pic.id });
  assert.ok(same(src.data, RED()), 'a picture inserted, undone and redone keeps its bytes (they were written back as text)');
  view.replacePicture({ id: pic.id, data: BLUE(), contentType: 'image/png', aspect: 1 });
  const after = (view.render().drawings || []).find((d) => d.kind === 'image');
  assert.equal(after.id, pic.id, 'the same picture');
  assert.deepEqual([after.x, after.y, after.width], [pic.x, pic.y, pic.width], 'where it was, as wide');
  assert.equal(after.height, after.width, 'as tall as the square picture');
  assert.ok(same(view.pictureSource({ id: pic.id }).data, BLUE()));
  view.undo();
  assert.ok(same(view.pictureSource({ id: pic.id }).data, RED()), 'Undo brings the old picture back');
  view.redo();
  const reopened = SheetView.open(view.save());
  const back = (reopened.render().drawings || []).find((d) => d.kind === 'image');
  assert.ok(same(reopened.pictureSource({ id: back.id }).data, BLUE()));
  assert.throws(() => view.pictureSource({ id: 'nothing' }), /no drawing/);
});

test('I16 Presentations: the picture handed out, the edited one put back where it was, its height from its shape, one undo step', () => {
  const d = Deck.open(buildPptx({ slides: [{ title: 'Our office' }] }));
  const { id } = d.addPicture(0, { data: RED(), contentType: 'image/png', name: 'Office', x: 100, y: 80, w: 400, h: 200 });
  assert.ok(same(d.pictureSource(0, id).data, RED()));
  const box = (deck) => {
    const xml = deck.pkg.text('ppt/slides/slide1.xml');
    const at = xml.indexOf(`<p:cNvPr id="${id}"`);
    const pic = xml.slice(xml.lastIndexOf('<p:pic>', at), xml.indexOf('</p:pic>', at));
    return /<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"/.exec(pic).slice(1).map(Number);
  };
  const before = box(d);
  const snap = d.snapshot();
  d.replacePicture(0, id, { data: BLUE(), contentType: 'image/png', aspect: 1 });
  d.pushUndo(snap);
  const after = box(d);
  assert.deepEqual(after.slice(0, 3), before.slice(0, 3), 'where it was, as wide');
  assert.equal(after[3], after[2], 'as tall as the square picture');
  assert.ok(same(d.pictureSource(0, id).data, BLUE()));
  d.undo();
  assert.ok(same(d.pictureSource(0, id).data, RED()), 'Undo brings the old picture back');
  d.redo();
  const reopened = Deck.open(d.save());
  assert.ok(same(reopened.pictureSource(0, id).data, BLUE()));
  assert.throws(() => d.pictureSource(0, 9999), /not a picture/);
});
