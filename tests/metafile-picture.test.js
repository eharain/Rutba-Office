// A picture kept as drawing calls, drawn in all three apps the same way.
//
// Word, Excel and PowerPoint keep a chart or a diagram pasted from another
// program as an EMF or a WMF. Documents and Worksheets gave such a picture
// no href and drew its empty frame; Presentations drew it only when it held
// nothing but bitmaps. Now one helper, drawablePicture, plays it as an SVG,
// and each app is handed that SVG: Documents and Worksheets through the
// drawer the app lends their readers, Presentations through its picture
// blobs, and a printed deck as data.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import zlib from 'node:zlib';
import { drawablePicture, pictureDataUri, setPictureDecoder, metafileKind } from '../packages/office-formats/src/picture.js';
import { sniff } from '../packages/office-formats/src/sniff.js';
import { APP_EXTENSIONS } from '../packages/office-formats/src/registry.js';
import { setPictureDrawer as lendSheets, toDataUri } from '@rutba/drawing/shapes';
import { setPictureDrawer as lendDocuments } from '@rutba/ooxml/document';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { Deck, buildPptx } from '@rutba/presentation';
import { deckPrintHtml } from '@rutba/presentation/print';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { gradientPng } from '../apps/desktop/main/sample-picture.js';
import { showcaseEmf, showcaseWmf, withMetafiles, EmfWriter } from './fixtures/metafiles.js';

const PNG = () => gradientPng(40, 30, [60, 160, 90], [200, 50, 50]);
const svgOf = (uri) => Buffer.from(String(uri).split(',')[1], 'base64').toString('utf8');

test('one helper: an EMF and a WMF as an SVG, a PNG as it is, damage as nothing', () => {
  const emf = drawablePicture(showcaseEmf(), 'word/media/image1.emf');
  assert.equal(emf.type, 'image/svg+xml');
  assert.match(emf.bytes.toString('utf8'), /^<svg [^>]*viewBox=/);
  assert.deepEqual([emf.width, emf.height], [480, 300]);
  const wmf = drawablePicture(showcaseWmf(), 'image2.wmf');
  assert.equal(wmf.type, 'image/svg+xml');
  assert.deepEqual([wmf.width, wmf.height], [288, 192]);
  // By its bytes, whatever its name says.
  assert.equal(drawablePicture(showcaseEmf(), 'misnamed.png').type, 'image/svg+xml');
  assert.equal(metafileKind(new Uint8Array(showcaseWmf()), ''), 'wmf', 'a Uint8Array is read too');
  const png = PNG();
  assert.equal(drawablePicture(png, 'a.png').bytes, png, 'a picture a window draws is handed back untouched');
  const bytes = showcaseEmf();
  assert.equal(drawablePicture(bytes, 'x.emf'), drawablePicture(bytes, 'x.emf'), 'played once for the same bytes');
  assert.equal(drawablePicture(Buffer.from('not a metafile at all, only words'), 'broken.emf'), null);
  assert.equal(drawablePicture(Buffer.alloc(0), 'empty.wmf'), null);
  assert.equal(drawablePicture(new EmfWriter(100, 100).build(), 'blank.emf'), null, 'a metafile that draws nothing is no picture');
  const emz = drawablePicture(zlib.gzipSync(showcaseEmf()), 'image3.emz');
  assert.equal(emz?.type, 'image/svg+xml', 'an .emz is the metafile gzipped');
  assert.match(pictureDataUri(showcaseWmf(), 'a.wmf'), /^data:image\/svg\+xml;base64,/);
});

test('a decoder lent for the kinds a window cannot draw is asked after the metafiles, and its throw is a null', () => {
  try {
    setPictureDecoder((b, name) => (name.endsWith('.tif') ? { type: 'image/png', bytes: PNG() } : null));
    assert.equal(drawablePicture(Buffer.from('II*\0 a tiff'), 'scan.tif').type, 'image/png');
    assert.equal(drawablePicture(showcaseEmf(), 'x.emf').type, 'image/svg+xml', 'a metafile is still drawn here');
    setPictureDecoder(() => { throw new Error('cannot'); });
    assert.equal(drawablePicture(Buffer.from('II*\0 another'), 'other.tif'), null);
  } finally {
    setPictureDecoder(null);
  }
  assert.equal(drawablePicture(Buffer.from('II*\0 a tiff'), 'scan.tif'), null, 'none lent, none drawn');
});

test('Pictures knows an .emf and a .wmf, by name and by their first bytes', () => {
  assert.ok(APP_EXTENSIONS.pictures.includes('emf') && APP_EXTENSIONS.pictures.includes('wmf'));
  assert.equal(sniff(showcaseEmf(), 'chart.emf').kind, 'emf');
  assert.equal(sniff(showcaseWmf(), 'clip.wmf').kind, 'wmf');
  assert.equal(sniff(showcaseEmf(), 'no-name').kind, 'emf', 'by the header record');
  assert.equal(sniff(showcaseWmf(), 'no-name').kind, 'wmf', 'by the placeable key');
  assert.equal(sniff(showcaseEmf(), 'chart.emf').app, 'pictures');
});

test('Worksheets: a lent drawer turns a metafile part into an SVG the frame draws; without one it has no href', () => {
  lendSheets(null);
  assert.equal(toDataUri(showcaseEmf(), 'xl/media/image1.emf'), null, 'nothing lent: no href rather than a broken one');
  lendSheets(drawablePicture);
  assert.match(toDataUri(showcaseEmf(), 'xl/media/image1.emf'), /^data:image\/svg\+xml;base64,/);
  assert.match(toDataUri(PNG(), 'xl/media/image2.png'), /^data:image\/png;base64,/, 'a PNG as before');
});

test('Worksheets: an EMF and a WMF in a workbook are drawn over the cells as SVG', async () => {
  lendSheets(drawablePicture);
  const view = new SheetView(buildXlsx({ sheets: [{ name: 'Sales', rows: [['Item', 'Qty'], ['Pens', 12]] }] }));
  view.select(3, 1);
  view.insertPicture({ name: 'chart', contentType: 'image/png', data: PNG(), widthPx: 240, heightPx: 150 });
  view.select(3, 6);
  view.insertPicture({ name: 'clip', contentType: 'image/png', data: PNG(), widthPx: 144, heightPx: 96 });
  const file = await withMetafiles(view.save(), [showcaseEmf(), showcaseWmf()]);
  const reopened = new SheetView(file);
  const drawn = (reopened.render().drawings || []).filter((d) => d.kind === 'image');
  assert.equal(drawn.length, 2);
  const inner = drawn.map((d) => svgOf(/href="(data:image\/svg\+xml;base64,[^"]+)"/.exec(d.svg)?.[1] || ''));
  assert.match(inner[0], /Quarterly sales/, 'the EMF, its title among its words');
  assert.match(inner[1], /Rutba drawing/, 'the WMF, its words too');
});

test('Documents: an EMF and a WMF in a document are drawn as SVG on the page', async () => {
  lendDocuments(drawablePicture);
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'A report' }, { text: 'Its chart.' }, { text: 'Its clip art.' }] }));
  view.setSelection({ block: 1, offset: 0 });
  view.insertImage({ name: 'chart', contentType: 'image/png', data: PNG(), widthPx: 480, heightPx: 300 });
  view.setSelection({ block: 3, offset: 0 });
  view.insertImage({ name: 'clip', contentType: 'image/png', data: PNG(), widthPx: 288, heightPx: 192 });
  const file = await withMetafiles(view.save(), [showcaseEmf(), showcaseWmf()]);
  const images = openDocx(file).render({ pages: false }).blocks.flatMap((b) => b.images || []);
  assert.equal(images.length, 2);
  for (const img of images) assert.match(img.href || '', /^data:image\/svg\+xml;base64,/);
  assert.match(svgOf(images[0].href), /Quarterly sales/);
  assert.match(svgOf(images[1].href), /Rutba drawing/);
});

test('Presentations: a slide\'s EMF and WMF are held as SVG, and a printed deck carries them as data', async () => {
  const d = Deck.open(buildPptx({ title: 'Metafiles', slides: [{ layout: 'blank' }] }));
  d.addPicture(0, { data: PNG(), contentType: 'image/png', name: 'Chart', x: 40, y: 40, w: 480, h: 300 });
  d.addPicture(0, { data: PNG(), contentType: 'image/png', name: 'Clip', x: 560, y: 40, w: 288, h: 192 });
  const bytes = await withMetafiles(d.save(), [showcaseEmf(), showcaseWmf()]);
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-metafile-')), 'metafiles.pptx');
  fs.writeFileSync(file, bytes);
  const held = [];
  let n = 0;
  const service = createDocumentService({
    holdBlob: (b, type, name) => { held.push({ type, name, svg: type === 'image/svg+xml' ? Buffer.from(b).toString('utf8') : '' }); const id = `t${++n}`; return { id, url: `blob://${id}`, size: 0 }; },
    releaseBlob: () => {},
  });
  const opened = service.open({ path: file, kind: 'deck' });
  service.model({ id: opened.id, slide: 0 });
  const svgs = held.filter((h) => h.type === 'image/svg+xml');
  assert.equal(svgs.length, 2, JSON.stringify(held.map((h) => [h.type, h.name])));
  assert.deepEqual(svgs.map((h) => h.name).sort(), ['image1.svg', 'image2.svg']);
  assert.ok(svgs.some((h) => /Quarterly sales/.test(h.svg)) && svgs.some((h) => /Rutba drawing/.test(h.svg)));
  const html = deckPrintHtml(Deck.open(bytes), {});
  assert.equal((html.match(/data:image\/svg\+xml;base64,/g) || []).length, 2, 'on paper as on the stage');
  fs.rmSync(path.dirname(file), { recursive: true, force: true });
});
