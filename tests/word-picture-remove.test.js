// Pictures — what the window's own file dialog hands over, and the way out.
//
// The dialog's read gives a Uint8Array; the engine used to take only a
// Buffer or base64 text, and wrote the digits of the array as the picture.
// Now any of the three is the file's own bytes. And a picture can be taken
// out again: its run, and its block when the block held nothing else.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx } from '@rutba/ooxml';
import { OoxmlPackage } from '@rutba/ooxml';

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');
const b64 = PNG.toString('base64');

test('a picture inserted as the byte array the file dialog hands over is written whole, taken out with its block, and undone', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'one' }, { text: 'two' }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertImage({ name: 'dot.png', contentType: 'image/png', data: new Uint8Array(PNG), widthPx: 40, heightPx: 40 });
  assert.equal(view.blocks.length, 3, 'a block of its own');
  const img = view.block(1).images[0];
  assert.ok(img && img.href.startsWith('data:image/png;base64,'), 'drawn from a data URL');
  assert.equal(img.href.split(',')[1], b64, 'the bytes, not a base64 reading of their digits');
  assert.ok(Buffer.from(OoxmlPackage.read(view.save()).read('word/media/rutba1.png')).equals(PNG), 'the media part is the file');

  view.removeImage({ block: 1, image: 0 });
  assert.equal(view.blocks.length, 2, 'the picture\'s own block went with it');
  assert.deepEqual(view.blocks.map((b) => b.text), ['one', 'two']);
  assert.equal(view.focus.block, 1, 'the caret on the block that took its place');
  view.undo();
  assert.equal(view.blocks.length, 3);
  assert.equal(view.block(1).images[0]?.href.split(',')[1], b64, 'one undo step brings it back');
});

test('bmp and webp are accepted with their own extensions, tiff refused, and a picture among words leaves the words', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'one' }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertImage({ name: 'a.webp', contentType: 'image/webp', data: PNG, widthPx: 10, heightPx: 10 });
  assert.ok(OoxmlPackage.read(view.save()).has('word/media/rutba1.webp'));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertImage({ name: 'b.bmp', contentType: 'image/bmp', data: b64, widthPx: 10, heightPx: 10 });
  assert.ok(OoxmlPackage.read(view.save()).has('word/media/rutba1.bmp'), 'base64 text still works, and each format counts its own parts');
  assert.throws(() => view.insertImage({ name: 'x.tif', contentType: 'image/tiff', data: PNG, widthPx: 10, heightPx: 10 }), /unsupported image type/);

  // The bmp's run moved into the paragraph of words: the picture out, the words stay.
  const doc = view.doc.doc;
  const words = doc.editParagraph(0);
  const run = /<w:r>[\s\S]*?<\/w:r>/.exec(doc.editParagraph(1).xml)[0];
  doc.xml = doc.xml.replace(words.xml, words.xml.replace(/<\/w:p>$/, run + '</w:p>'));
  view._invalidate();
  assert.equal(view.block(0).images.length, 1, 'a picture among the words');
  view.removeImage({ block: 0, image: 0 });
  assert.equal(view.block(0).text, 'one', 'the words stay');
  assert.equal(view.block(0).images.length, 0);
  assert.equal(view.blocks.length, 3, 'no block went');

  // The body's last paragraph, holding only a picture, is emptied rather than removed.
  const only = openDocx(buildDocx({ styles: true, paragraphs: [{ text: '' }] }));
  only.collapseTo({ block: 0, offset: 0 });
  only.insertImage({ name: 'c.png', contentType: 'image/png', data: PNG, widthPx: 10, heightPx: 10 });
  const empty = only.doc.doc.editParagraph(0);
  only.doc.doc._spliceBody(empty.start, empty.end, '');
  only._invalidate();
  assert.equal(only.blocks.length, 1, 'a body of one picture');
  only.removeImage({ block: 0, image: 0 });
  assert.equal(only.blocks.length, 1, 'the last paragraph stays');
  assert.equal(only.block(0).images.length, 0, 'emptied');
  assert.throws(() => only.removeImage({ block: 0, image: 0 }), /no picture/);
});
