// The Image app's project file (.rimg, docs/STUDIO.md D2 and I17): a stored
// zip whose first entry says what it is, the layered document and every
// picture it is made of inside it, that comes back as it went, names a
// picture a layer asks for that is not there, refuses what is not one, and
// is registered as a kind the suite opens in Image.
import test from 'node:test';
import assert from 'node:assert/strict';
import { writeRimg, readRimg, rimgPath, rimgPaths, safeName, RIMG_MIME } from '../packages/studio/src/rimg.js';
import { unzip, zipStored } from '../packages/studio/src/rvid.js';
import { emptyImage, pictureLayer, groupLayers, normaliseImage } from '../packages/studio/src/layers.js';
import { sniff, kindFromExtension, KINDS } from '@rutba/office-formats/sniff';
import { APP_EXTENSIONS, WRITABLE, fileAssociations, saveFilters } from '@rutba/office-formats/registry';

const png = (n) => new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, n, n, n]);

function project() {
  let doc = emptyImage({ width: 800, height: 600 });
  doc = {
    ...doc,
    quick: { source: 'source/beach.jpg', ops: [{ op: 'rotate', degrees: 90 }], adjust: { brightness: 110 }, pixels: { exposure: 0.4 } },
    patches: [
      pictureLayer(doc, { url: 'rimg:bitmaps/background.png', width: 800, height: 600, label: 'Background', id: 'background' }),
      { id: 'image-1', type: 'image', url: 'rimg:bitmaps/image-1.png', fx: 0.1, fy: 0.1, fw: 0.3, shadow: false, label: 'Logo', opacity: 0.7, blend: 'multiply' },
      { id: 'wordart-1', type: 'wordart', text: 'SALE', preset: 'gold', fx: 0.5, fy: 0.3, sizeFrac: 0.1, rot: 12 },
      { id: 'adjust-1', type: 'adjust', preset: 'warm', locked: true },
    ],
  };
  return groupLayers(doc, ['image-1', 'wordart-1']).doc;
}

const files = () => ({
  'source/beach.jpg': new Uint8Array([0xff, 0xd8, 0xff, 1, 2, 3]),
  'bitmaps/background.png': png(1),
  'bitmaps/image-1.png': png(2),
});

test('I17: a project goes out as a stored zip, mimetype first, and comes back exactly with its pictures', async () => {
  const doc = project();
  const bytes = writeRimg(doc, files());
  assert.equal(String.fromCharCode(...bytes.subarray(0, 4)), 'PK\u0003\u0004');
  assert.equal(new TextDecoder().decode(bytes.subarray(38, 38 + RIMG_MIME.length)), RIMG_MIME, 'mimetype first, stored, readable in the head');
  const parts = await unzip(bytes);
  assert.deepEqual(Object.keys(parts), ['mimetype', 'doc.json', 'source/beach.jpg', 'bitmaps/background.png', 'bitmaps/image-1.png']);
  const back = await readRimg(bytes);
  assert.deepEqual(normaliseImage(back.doc), normaliseImage(doc), 'layers, groups, opacity, blend, lock and the quick edits round-trip');
  assert.deepEqual(back.missing, []);
  assert.deepEqual([...back.files['bitmaps/image-1.png']], [...png(2)], 'a picture comes back byte for byte');
  assert.deepEqual([...back.files['source/beach.jpg']], [...files()['source/beach.jpg']]);
  assert.equal(back.doc.rimg, undefined, 'the version mark is the file\'s, not the document\'s');
});

test('I17: a layer whose picture is not in the file is named; only pictures are kept inside', async () => {
  const doc = project();
  const some = files();
  delete some['bitmaps/image-1.png'];
  some['../escape.png'] = png(9);
  const bytes = writeRimg(doc, some);
  const back = await readRimg(bytes);
  assert.deepEqual(back.missing, ['bitmaps/image-1.png']);
  assert.ok(!Object.keys(await unzip(bytes)).includes('../escape.png'), 'nothing outside source/ and bitmaps/ is written');
  assert.deepEqual(rimgPaths(doc).sort(), ['bitmaps/background.png', 'bitmaps/image-1.png', 'source/beach.jpg']);
  assert.equal(rimgPath('rimg:bitmaps/a.png'), 'bitmaps/a.png');
  assert.equal(rimgPath('D:/pics/a.png'), null);
  assert.equal(safeName('C:\\Users\\me\\My <pic>.jpg'), 'My _pic_.jpg');
  assert.equal(safeName('...'), 'picture');
});

test('I17: what is not a project is refused in words, and a newer one asks for an update', async () => {
  await assert.rejects(readRimg(new Uint8Array([1, 2, 3, 4])), /not a Rutba Image project: it is not a zip/);
  const video = zipStored([{ name: 'mimetype', bytes: new TextEncoder().encode('application/vnd.rutba.video-project+zip') }]);
  await assert.rejects(readRimg(video), /not a Rutba Image project\./);
  const empty = zipStored([{ name: 'mimetype', bytes: new TextEncoder().encode(RIMG_MIME) }]);
  await assert.rejects(readRimg(empty), /no document in it/);
  const newer = writeRimg({ ...project(), rimg: 99 });
  const parts = await unzip(newer);
  const body = JSON.parse(new TextDecoder().decode(parts['doc.json']));
  assert.equal(body.rimg, 99);
  await assert.rejects(readRimg(newer), /newer Rutba Office/);
});

test('I17: .rimg is a kind the suite opens in Image, is offered by the Open and Save As dialogs, and is associated with the app', () => {
  assert.equal(kindFromExtension('poster.rimg'), 'rimg');
  assert.equal(KINDS.rimg.app, 'image');
  assert.equal(KINDS.rimg.mime, RIMG_MIME);
  assert.equal(APP_EXTENSIONS.image[0], 'rimg');
  assert.equal(WRITABLE.image[0].kind, 'rimg');
  assert.deepEqual(saveFilters('image')[0].extensions, ['rimg']);
  const assoc = fileAssociations().find((a) => a.ext === 'rimg');
  assert.equal(assoc.app, 'image');
  assert.equal(assoc.name, 'Rutba Image Project');
  const bytes = writeRimg(project(), files());
  assert.equal(sniff(bytes, 'poster.rimg').kind, 'rimg', 'a zip named .rimg is a project');
  assert.equal(sniff(bytes, 'renamed.zip').kind, 'rimg', 'and one renamed is still known by its first entry');
});
