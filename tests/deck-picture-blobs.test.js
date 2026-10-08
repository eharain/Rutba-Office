// A deck's pictures are held once: the blob store keeps the bytes it is given
// rather than a copy, and a picture whose bytes change (a 3D model turned)
// lets its old blob go.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';
import { hold, heldBlob } from '../packages/office-shell/src/electron/blobs.js';
import { createDocumentService } from '../apps/desktop/main/documents.js';
import { cubeGlb, pngOf } from './fixtures/cube-glb.js';

const square = (rgba) => pngOf(Array.from({ length: 4 }, () => Array.from({ length: 4 }, () => rgba)));

test('the blob store keeps the bytes it is given, not a copy', () => {
  const bytes = Buffer.from('a picture');
  const { id } = hold(bytes, 'image/png', 'a.png');
  assert.equal(heldBlob(id).bytes, bytes);
  const view = new Uint8Array([1, 2, 3, 4]).subarray(1, 3);
  const second = hold(view, 'image/png', 'b.png');
  assert.equal(heldBlob(second.id).bytes.buffer, view.buffer, 'a view, not a copy');
  assert.deepEqual([...heldBlob(second.id).bytes], [2, 3]);
});

test('a picture whose bytes change lets its old blob go', () => {
  const d = Deck.open(buildPptx({ title: 'Model', slides: [{ layout: 'blank' }] }));
  const placed = d.addModel3d(0, { model: cubeGlb(), png: square([200, 0, 0, 255]), view: { yaw: 25, pitch: 15 }, name: 'Cube', x: 100, y: 80, w: 300, h: 300 });
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-blobs-')), 'model.pptx');
  fs.writeFileSync(file, d.save());
  const held = new Set();
  const released = [];
  let n = 0;
  const service = createDocumentService({
    holdBlob: () => { const id = `t${++n}`; held.add(id); return { id, url: `blob://${id}`, size: 0 }; },
    releaseBlob: (id) => { released.push(id); held.delete(id); },
  });
  const opened = service.open({ path: file, kind: 'deck' });
  const shapeId = placed?.id ?? placed;
  const before = new Set(held);
  service.apply({ id: opened.id, ops: [{ op: 'setModel3dView', slide: 0, shape: shapeId, png: square([0, 0, 200, 255]), view: { yaw: 60, pitch: 10 } }] });
  service.model({ id: opened.id, slide: 0 });
  assert.ok(released.length >= 1 && released.every((id) => before.has(id)), `released ${released.join(', ')} of ${[...before].join(', ')}`);
});
