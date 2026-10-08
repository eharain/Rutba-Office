// Insert → 3D Models in Documents: the model goes in as a picture in a
// paragraph of its own, the .glb kept beside it from the picture's cNvPr
// with the view it is drawn at; a new view rewrites the picture in its own
// part, and one Undo brings the old picture back with the old view.
import test from 'node:test';
import assert from 'node:assert/strict';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';
import { cubeGlb, pngOf } from './fixtures/cube-glb.js';

const RED = pngOf([[[255, 0, 0, 255]]]);
const BLUE = pngOf([[[0, 0, 255, 255]]]);
const b64 = (buf) => Buffer.from(buf).toString('base64');

function withModel() {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Before' }, { text: 'After' }] }));
  view.collapseTo({ block: 0, offset: 0 });
  view.insertImage({ name: 'Cube', contentType: 'image/png', data: RED, widthPx: 200, heightPx: 200, model3d: { data: new Uint8Array(cubeGlb()), view: { yaw: 25, pitch: 15 } } });
  return view;
}

test('a 3D model goes in as a picture with the model kept beside it and the view it is drawn at', () => {
  const view = withModel();
  const img = view.block(1).images[0];
  assert.deepEqual(img.model3d, { view: { yaw: 25, pitch: 15, roll: 0 } });
  assert.equal(img.href.split(',')[1], b64(RED));
  const pkg = OoxmlPackage.read(view.save());
  assert.ok(Buffer.from(pkg.read('word/media/model1.glb')).equals(cubeGlb()));
  assert.match(pkg.text('[Content_Types].xml'), /model\/gltf-binary/);
  assert.match(pkg.text('word/document.xml'), /<pic:cNvPr id="\d+" name="Cube"><a:extLst><a:ext uri="\{5E2C9A41-7B3D-4F6A-9C18-3D0A6E1B2F77\}"><r3d:model xmlns:r3d="http:\/\/schemas\.rutba\.io\/office\/2026\/model3d" xmlns:r="[^"]+" r:embed="rId\d+" yaw="25" pitch="15" roll="0"\/><\/a:ext><\/a:extLst><\/pic:cNvPr>/);
  assert.ok(Buffer.from(view.model3dSource(1, 0).data).equals(cubeGlb()));
  assert.throws(() => view.insertImage({ name: 'x', contentType: 'image/png', data: RED, widthPx: 5, heightPx: 5, model3d: { data: Buffer.from('nope') } }), /binary glTF/);
});

test('a new view rewrites the picture in its own part; one Undo brings the old picture and view back', () => {
  const view = withModel();
  view.setModel3dView(1, 0, { png: BLUE, view: { yaw: -35, pitch: 25 } });
  let img = view.block(1).images[0];
  assert.deepEqual(img.model3d.view, { yaw: -35, pitch: 25, roll: 0 });
  assert.equal(img.href.split(',')[1], b64(BLUE));
  const pkg = OoxmlPackage.read(view.save());
  assert.equal(pkg.partNames().filter((n) => /^word\/media\/rutba\d+\.png$/.test(n)).length, 1, 'no second picture part');
  view.undo();
  img = view.block(1).images[0];
  assert.deepEqual(img.model3d.view, { yaw: 25, pitch: 15, roll: 0 });
  assert.equal(img.href.split(',')[1], b64(RED), 'the old picture back');
  view.redo();
  assert.equal(view.block(1).images[0].href.split(',')[1], b64(BLUE));
  assert.throws(() => view.setModel3dView(0, 0, { png: BLUE, view: {} }), /not a 3D model/);
});

test('the file opens again with the model and its view', () => {
  const view = withModel();
  view.setModel3dView(1, 0, { png: BLUE, view: { yaw: 90 } });
  const again = openDocx(view.save());
  assert.deepEqual(again.block(1).images[0].model3d.view, { yaw: 90, pitch: 0, roll: 0 });
  assert.ok(Buffer.from(again.model3dSource(1, 0).data).equals(cubeGlb()));
});
