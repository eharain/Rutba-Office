/**
 * Insert → 3D Models, the engine's half: a glTF 2.0 read into triangles —
 * a .glb's chunks, a .gltf's buffers in data: addresses or in files beside
 * it, node transforms, materials and their pictures — and drawn into pixels
 * by the suite's own renderer, the same on every computer: the model fitted
 * to the picture, lit, its texture sampled, transparent round it.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { readModel, renderModel, fit, toGlb, gltfFiles, isGlb, MODEL_VIEWS } from '@rutba/imaging/model3d';
import { cubeGlb } from './fixtures/cube-glb.js';

const px = (pic, x, y) => Array.from(pic.data.slice((y * pic.width + x) * 4, (y * pic.width + x) * 4 + 4));

test('a .glb is read into its triangles, its materials and the sphere round it', () => {
  const m = readModel(cubeGlb());
  assert.equal(m.triangles, 12);
  assert.equal(m.meshes.length, 2);
  assert.deepEqual(m.meshes[0].material.color, [0.8, 0.05, 0.05, 1]);
  assert.deepEqual(m.bounds.min, [-1, -1, -1]);
  assert.ok(Math.abs(m.bounds.radius - Math.sqrt(3)) < 1e-6);
  assert.equal(readModel(cubeGlb({ texture: true })).images.length, 1);
});

test('the cube drawn from the front is its red face, lit, edge to edge; turned, it shows three faces with nothing round it', () => {
  const m = readModel(cubeGlb());
  const front = renderModel(m, { width: 120, height: 120, view: { yaw: 0, pitch: 0 } });
  const [r, g, b, a] = px(front, 60, 60);
  assert.equal(a, 255);
  assert.ok(r > 150 && g < 90 && b < 90 && r > g + 100, `red, lit: ${[r, g, b]}`);
  const turned = renderModel(m, { width: 120, height: 120, view: { yaw: 35, pitch: 25 } });
  assert.equal(px(turned, 0, 0)[3], 0, 'transparent round the model');
  assert.equal(px(turned, 60, 60)[3], 255);
  // From above the front and the left: the top (grey) above the middle, the left side (grey) at the left, the front (red) at the right.
  const top = px(turned, 60, 14); const left = px(turned, 22, 70); const face = px(turned, 90, 75);
  assert.ok(Math.abs(top[0] - top[1]) < 12 && top[3] === 255, `top grey: ${top}`);
  assert.ok(Math.abs(left[0] - left[1]) < 12 && left[3] === 255, `left grey: ${left}`);
  assert.ok(face[0] > face[1] + 80, `front red: ${face}`);
  assert.ok(top[0] > left[0], 'the top, facing the key light, is lighter than the left side');
});

test('a texture is sampled where its coordinates say, and fit gives the picture the model\'s own shape at a view', () => {
  const m = readModel(cubeGlb({ texture: true }));
  // The checks' four by four squares, decoded as the window would hand them in.
  const R = [220, 30, 30, 255]; const W = [255, 255, 255, 255];
  const rows = [[R, W, R, W], [W, R, W, R], [R, W, R, W], [W, R, W, R]];
  const textures = new Map([[0, { width: 4, height: 4, data: Uint8ClampedArray.from(rows.flat(2)) }]]);
  const pic = renderModel(m, { width: 200, height: 200, view: {}, textures, margin: 0.02 });
  const red = px(pic, 28, 28); const white = px(pic, 76, 28); const red2 = px(pic, 76, 76);
  assert.ok(red[0] > 150 && red[1] < 80, `red square: ${red}`);
  assert.ok(white[1] > 200 && white[2] > 200, `white square: ${white}`);
  assert.ok(red2[0] > 150 && red2[1] < 80, `the next row starts the other way: ${red2}`);
  const square = fit(m, { yaw: 0, pitch: 0 }, 400);
  assert.deepEqual(square, { width: 400, height: 400 });
  const tall = fit(m, { yaw: 0, pitch: 60 }, 400);
  assert.ok(tall.height === 400 && tall.width < 400, `tipped towards us the cube is taller than wide: ${JSON.stringify(tall)}`);
  assert.equal(MODEL_VIEWS.length, 10);
});

test('a .gltf with its buffer in a data: address or a file beside it is packed into one .glb that reads the same', () => {
  const glb = cubeGlb();
  const jl = glb.readUInt32LE(12);
  const json = JSON.parse(glb.subarray(20, 20 + jl).toString());
  const bin = glb.subarray(28 + jl);
  const inline = Buffer.from(JSON.stringify({ ...json, buffers: [{ byteLength: bin.length, uri: `data:application/octet-stream;base64,${bin.toString('base64')}` }] }));
  assert.equal(readModel(inline).triangles, 12);
  assert.equal(readModel(toGlb(inline)).triangles, 12);
  assert.ok(isGlb(toGlb(inline)));
  const beside = Buffer.from(JSON.stringify({ ...json, buffers: [{ byteLength: bin.length, uri: 'cube%20data.bin' }] }));
  assert.deepEqual(gltfFiles(beside), ['cube data.bin']);
  assert.throws(() => readModel(beside), /needs the file "cube%20data.bin" beside it/);
  const packed = toGlb(beside, { readFile: (u) => (u === 'cube data.bin' ? bin : null) });
  assert.equal(readModel(packed).triangles, 12);
  assert.deepEqual(gltfFiles(packed), []);
});

test('what is not a model, or is one this cannot read, says so in a sentence', () => {
  assert.throws(() => readModel(Buffer.from('hello')), /not a glTF file/);
  assert.throws(() => readModel(Buffer.from(JSON.stringify({ asset: { version: '1.0' } }))), /Only glTF 2\.0/);
  assert.throws(() => readModel(Buffer.from(JSON.stringify({ asset: { version: '2.0' }, extensionsRequired: ['KHR_draco_mesh_compression'] }))), /Draco/);
  assert.throws(() => readModel(Buffer.from(JSON.stringify({ asset: { version: '2.0' }, scenes: [{ nodes: [] }] }))), /nothing to draw/);
});
