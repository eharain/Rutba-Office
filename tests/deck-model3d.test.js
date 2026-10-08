/**
 * Insert → 3D Models in a deck: the model drawn as a picture every reader
 * shows, the .glb kept beside it from the picture's nvPr with the view it
 * is drawn at, read back for drawing again, and a new view's picture put in
 * its place — in its own part when nothing else uses it, a new one when a
 * duplicated slide shares it.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { cubeGlb, pngOf } from './fixtures/cube-glb.js';

const RED = pngOf([[[255, 0, 0, 255]]]);
const BLUE = pngOf([[[0, 0, 255, 255]]]);
const xmlOf = (d, i = 0) => d.pkg.text(d.slideParts[i].part);

function withModel() {
  const d = Deck.open(buildPptx({ title: 'Models', slides: [{ layout: 'title', title: 'A cube' }] }));
  const placed = d.addModel3d(0, { model: cubeGlb(), png: RED, view: { yaw: 25, pitch: 15 }, name: 'Cube', x: 100, y: 80, w: 300, h: 300 });
  return { d, placed };
}

test('a 3D model goes in as its picture, the .glb kept beside it with the view it is drawn at', () => {
  const { d, placed } = withModel();
  assert.equal(placed.model, 'ppt/media/model1.glb');
  assert.ok(d.pkg.has('ppt/media/model1.glb'));
  assert.match(d.pkg.text('[Content_Types].xml'), /<Default Extension="glb" ContentType="model\/gltf-binary"\/>/);
  const xml = xmlOf(d);
  assert.match(xml, /<p:pic><p:nvPicPr><p:cNvPr id="\d+" name="Cube" descr="Cube"\/><p:cNvPicPr><a:picLocks noChangeAspect="1"\/><\/p:cNvPicPr><p:nvPr><p:extLst><p:ext uri="\{5E2C9A41-7B3D-4F6A-9C18-3D0A6E1B2F77\}"><r3d:model xmlns:r3d="http:\/\/schemas\.rutba\.io\/office\/2026\/model3d" r:embed="rId\d+" yaw="25" pitch="15" roll="0"\/><\/p:ext><\/p:extLst><\/p:nvPr><\/p:nvPicPr><p:blipFill><a:blip r:embed="rId\d+"\/>/);
  const shape = d.slide(0).shapes.find((s) => String(s.id) === String(placed.id));
  assert.equal(shape.kind, 'picture');
  assert.deepEqual(shape.model3d, { view: { yaw: 25, pitch: 15, roll: 0 } });
  const src = d.model3dSource(0, placed.id);
  assert.ok(Buffer.from(src.data).equals(cubeGlb()));
  assert.deepEqual(src.view, { yaw: 25, pitch: 15, roll: 0 });
  assert.throws(() => d.addModel3d(0, { model: Buffer.from('not a model'), png: RED, w: 10, h: 10 }), /binary glTF/);
});

test('a new view replaces the picture in its own part and keeps the view; the file opens again with both', () => {
  const { d, placed } = withModel();
  d.setModel3dView(0, placed.id, { png: BLUE, view: { yaw: -35, pitch: 25 } });
  assert.ok(d.pkg.read(placed.part).equals(BLUE), 'the same part, rewritten');
  assert.match(xmlOf(d), /<r3d:model [^>]*yaw="-35" pitch="25" roll="0"\/>/);
  const again = Deck.open(d.save());
  const shape = again.slide(0).shapes.find((s) => String(s.id) === String(placed.id));
  assert.deepEqual(shape.model3d.view, { yaw: -35, pitch: 25, roll: 0 });
  assert.ok(Buffer.from(again.model3dSource(0, placed.id).data).equals(cubeGlb()));
  assert.throws(() => d.setModel3dView(0, 2, { png: BLUE, view: {} }), /not a 3D model/);
});

test('on a duplicated slide the picture is shared, so a new view there gets a picture of its own', () => {
  const { d, placed } = withModel();
  d.duplicateSlide(0);
  d.setModel3dView(1, placed.id, { png: BLUE, view: { yaw: 90 } });
  assert.ok(d.pkg.read(placed.part).equals(RED), 'the first slide\'s picture is as it was');
  const blip = /<a:blip r:embed="(rId\d+)"/.exec(xmlOf(d, 1))[1];
  const rels = d.pkg.rels(d.slideParts[1].part);
  const target = rels.find((r) => r.Id === blip).Target;
  assert.notEqual(target.split('/').pop(), placed.part.split('/').pop());
  assert.match(xmlOf(d, 1), /yaw="90"/);
  assert.match(xmlOf(d, 0), /yaw="25"/);
});

test('undoing a turn brings back the picture and the view it had', () => {
  const { d, placed } = withModel();
  const snap = d.snapshot();
  d.setModel3dView(0, placed.id, { png: BLUE, view: { yaw: 120 } });
  d.pushUndo(snap);
  assert.ok(d.media(placed.part).equals(BLUE));
  d.undo();
  assert.ok(d.media(placed.part).equals(RED), 'the old picture back');
  assert.match(xmlOf(d), /yaw="25" pitch="15"/);
});
