// Insert > QR Code puts the code in as a picture whose alt text holds its
// text (and whose title says it is a QR Code, at which level and colour), so
// a double-click opens it again: in a document, a workbook and a deck, the
// alt text and title are written on insert, read back after a save and a
// reopen, and written again when the picture is replaced by a new code.
import test from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { Deck, buildPptx } from '@rutba/presentation';

/** A small PNG, a grey square, `n` pixels a side. */
function png(n = 8, grey = 0) {
  const raw = Buffer.alloc((n * 3 + 1) * n, grey);
  for (let y = 0; y < n; y++) raw[y * (n * 3 + 1)] = 0;
  const crc = (buf) => { let c = -1; for (const b of buf) { c ^= b; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; } return (c ^ -1) >>> 0; };
  const chunk = (type, data) => { const h = Buffer.alloc(8); h.writeUInt32BE(data.length); h.write(type, 4, 'latin1'); const t = Buffer.alloc(4); t.writeUInt32BE(crc(Buffer.concat([h.subarray(4), data]))); return Buffer.concat([h, data, t]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(n, 0); ihdr.writeUInt32BE(n, 4); ihdr[8] = 8; ihdr[9] = 2;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const ALT = { descr: 'https://rutba.io/office?a=1&b="2"', title: 'QR Code; M; #1F3864' };
const NEXT = { descr: 'Second <text> & more', title: 'QR Code; H; #000000' };

test('a document keeps a picture\'s alt text and title, and writes them again when the picture is replaced', () => {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Scan this:' }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertImage({ name: 'QR Code', contentType: 'image/png', data: png(), widthPx: 96, heightPx: 96, ...ALT });
  let again = openDocx(Buffer.from(view.save()));
  let image = again.block(1).images[0];
  assert.equal(image.descr, ALT.descr);
  assert.equal(image.title, ALT.title);
  again.replacePicture({ block: 1, image: 0, data: png(12, 40), contentType: 'image/png', widthPx: 128, heightPx: 128, alt: NEXT });
  again = openDocx(Buffer.from(again.save()));
  image = again.block(1).images[0];
  assert.equal(image.descr, NEXT.descr);
  assert.equal(image.title, NEXT.title);
  assert.equal(Math.round(image.widthPx), 128);
});

test('a workbook keeps a picture\'s alt text and title, in its frame for the window too, and writes them again on a replace', () => {
  const view = SheetView.open(Buffer.from(buildXlsx({ sheets: [{ name: 'Sheet1', rows: [['Scan this']] }] })), { viewportWidth: 800, viewportHeight: 600 });
  view.select(2, 1);
  view.insertPicture({ name: 'QR Code', contentType: 'image/png', data: png(), widthPx: 96, heightPx: 96, ...ALT });
  let again = SheetView.open(Buffer.from(view.save()), { viewportWidth: 800, viewportHeight: 600 });
  let picture = again.drawings.get('Sheet1').find((d) => d.kind === 'image');
  assert.equal(picture.descr, ALT.descr);
  assert.equal(picture.title, ALT.title);
  const inFrame = again.render().drawings.find((d) => d.kind === 'image');
  assert.equal(inFrame.descr, ALT.descr);
  again.replacePicture({ id: picture.id, data: png(12, 40), contentType: 'image/png', alt: NEXT });
  again = SheetView.open(Buffer.from(again.save()), { viewportWidth: 800, viewportHeight: 600 });
  picture = again.drawings.get('Sheet1').find((d) => d.kind === 'image');
  assert.equal(picture.descr, NEXT.descr);
  assert.equal(picture.title, NEXT.title);
});

test('a deck keeps a picture\'s alt text and title, and writes them again on a replace', () => {
  const deck = Deck.open(Buffer.from(buildPptx({ slides: [{ layout: 'obj', title: 'Scan this', body: '' }] })));
  const { id } = deck.addPicture(0, { data: png(), contentType: 'image/png', name: 'QR Code', x: 100, y: 100, w: 200, h: 200, ...ALT });
  let again = Deck.open(Buffer.from(deck.save()));
  let picture = again.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(picture.descr, ALT.descr);
  assert.equal(picture.title, ALT.title);
  again.replacePicture(0, id, { data: png(12, 40), contentType: 'image/png', alt: NEXT });
  again = Deck.open(Buffer.from(again.save()));
  picture = again.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(picture.descr, NEXT.descr);
  assert.equal(picture.title, NEXT.title);
});
