// Image: selections kept in the .rimg project (docs/STUDIO.md 7.2, I11 and
// I17): a selection packs to a few bytes and unpacks to the same plane value
// for value, feathered edges included; anything that is not a packed
// selection is refused; and the project carries the selection on the canvas
// and each Adjusted selection layer's own (with the picture it adjusts),
// names them among what it needs, and says which are missing.
import test from 'node:test';
import assert from 'node:assert/strict';
import { rectSelection, ellipseSelection, featherSelection, emptySelection, selectAll, packSelection, unpackSelection } from '../packages/imaging/src/select.js';
import { writeRimg, readRimg, rimgPaths } from '../packages/studio/src/rimg.js';

const same = (a, b) => a.width === b.width && a.height === b.height && a.data.length === b.data.length && a.data.every((v, i) => v === b.data[i]);

test('a selection packs small and comes back value for value', () => {
  const rect = rectSelection(2000, 1500, { x: 100.5, y: 200.25, width: 900, height: 600 });
  const packed = packSelection(rect);
  assert.ok(packed.length < 20000, `${packed.length} bytes for a 2000 × 1500 rectangle`);
  assert.ok(same(unpackSelection(packed), rect));
  const soft = featherSelection(ellipseSelection(640, 480, { x: 80, y: 60, width: 400, height: 300 }), 12);
  const back = unpackSelection(packSelection(soft));
  assert.ok(same(back, soft), 'a feathered ellipse, every in-between value kept');
  assert.ok(soft.data.some((v) => v > 0 && v < 255), 'the fixture has soft edges');
  for (const s of [emptySelection(33, 17), selectAll(1, 1), selectAll(4096, 3)]) assert.ok(same(unpackSelection(packSelection(s)), s));
});

test('anything that is not a packed selection is refused rather than read wrong', () => {
  const packed = packSelection(rectSelection(50, 40, { x: 5, y: 5, width: 20, height: 10 }));
  assert.equal(unpackSelection(new Uint8Array(0)), null);
  assert.equal(unpackSelection(Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 1, 0, 0, 0, 1, 0, 0, 0, 1)), null, 'a PNG');
  assert.equal(unpackSelection(packed.subarray(0, packed.length - 3)), null, 'cut short');
  const longer = new Uint8Array(packed.length + 2);
  longer.set(packed);
  longer.set([0x7f, 9], packed.length);
  assert.equal(unpackSelection(longer), null, 'more pixels than its size');
  const future = packed.slice();
  future[4] = 2;
  assert.equal(unpackSelection(future), null, 'a later version');
});

test('the project carries the selection on the canvas and each layer\'s own, and lists them among what it needs', async () => {
  const doc = {
    v: 1, kind: 'image', size: { width: 64, height: 48 },
    patches: [
      { id: 'background', type: 'image', url: 'rimg:bitmaps/background.png' },
      { id: 'adjusted-1', type: 'image', kind: 'adjusted', url: 'rimg:bitmaps/adjusted-1.png', quick: { adjust: { grayscale: 100 } } },
    ],
    selection: 'rimg:selections/selection.rsel',
    scoped: 'adjusted-1',
    scopes: { 'adjusted-1': { selection: 'rimg:selections/adjusted-1.rsel', base: 'rimg:bitmaps/adjusted-1-under.png' } },
  };
  const active = ellipseSelection(64, 48, { x: 10, y: 8, width: 30, height: 20 });
  const own = rectSelection(64, 48, { x: 4, y: 4, width: 20, height: 12 });
  const png = Uint8Array.of(0x89, 0x50, 0x4e, 0x47);
  const files = {
    'bitmaps/background.png': png,
    'bitmaps/adjusted-1.png': png,
    'bitmaps/adjusted-1-under.png': png,
    'selections/selection.rsel': packSelection(active),
    'selections/adjusted-1.rsel': packSelection(own),
    'elsewhere/x.bin': png,
  };
  assert.deepEqual(rimgPaths(doc).sort(), ['bitmaps/adjusted-1-under.png', 'bitmaps/adjusted-1.png', 'bitmaps/background.png', 'selections/adjusted-1.rsel', 'selections/selection.rsel']);
  const read = await readRimg(writeRimg(doc, files));
  assert.deepEqual(read.missing, []);
  assert.equal(read.doc.selection, doc.selection);
  assert.equal(read.doc.scoped, 'adjusted-1');
  assert.deepEqual(read.doc.scopes, doc.scopes);
  assert.ok(same(unpackSelection(read.files['selections/selection.rsel']), active));
  assert.ok(same(unpackSelection(read.files['selections/adjusted-1.rsel']), own));
  assert.equal(read.files['elsewhere/x.bin'], undefined, 'only the project\'s own folders are kept');
  const lacking = { ...files };
  delete lacking['selections/adjusted-1.rsel'];
  assert.deepEqual((await readRimg(writeRimg(doc, lacking))).missing, ['selections/adjusted-1.rsel']);
});

test('a project from before selections were kept opens as it did', async () => {
  const doc = { v: 1, kind: 'image', size: { width: 8, height: 8 }, patches: [{ id: 'background', type: 'image', url: 'rimg:bitmaps/background.png' }] };
  const read = await readRimg(writeRimg(doc, { 'bitmaps/background.png': Uint8Array.of(1) }));
  assert.equal(read.doc.selection, undefined);
  assert.equal(read.doc.scopes, undefined);
  assert.deepEqual(read.missing, []);
});
