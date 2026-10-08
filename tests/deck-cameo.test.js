// Insert → Cameo, the engine's half: a shape the show fills with the camera,
// put at the slide's lower right in a light grey, marked by an extension of
// the suite's own; drawn with a camera in its middle where there is no
// camera to fill it; Camera Shape changes its outline; the file keeps it.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx } from '@rutba/presentation';
import { renderSlide } from '@rutba/presentation/render';

const fresh = () => Deck.open(buildPptx({ title: 'Talk', slides: [{ layout: 'title', title: 'Hello' }] }));

test('a cameo goes in at the slide\'s lower right, a grey rectangle marked as the camera\'s', () => {
  const d = fresh();
  const id = d.addCameo(0);
  const shape = d.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.equal(shape.cameo, true);
  assert.equal(shape.preset, 'rect');
  const W = d.size.width; const H = d.size.height;
  const g = shape.geometry;
  assert.ok(Math.abs(g.w - W * 0.3) < 2 && Math.abs(g.h - (W * 0.3 * 9) / 16) < 2, `a third of the width at 16:9: ${JSON.stringify(g)}`);
  assert.ok(Math.abs(g.x + g.w - W * 0.96) < 2 && Math.abs(g.y + g.h - H * 0.96) < 2, 'at the lower right, a margin in');
  assert.match(d.pkg.text(d.slideParts[0].part), /<p:nvPr><p:extLst><p:ext uri="\{8B1F3C27-5D64-4E0A-A9C2-71E4F0D35B19\}"><rcam:cameo xmlns:rcam="http:\/\/schemas\.rutba\.io\/office\/2026\/cameo"\/><\/p:ext><\/p:extLst><\/p:nvPr>/);
  assert.match(renderSlide(d.slide(0), {}), /<circle [^>]*fill="#d9d9d9"\/><\/g>/, 'a camera drawn in it');
  assert.throws(() => d.addCameo(0, { shape: 'star5' }), /rectangle, an oval or a rounded rectangle/);
});

test('Camera Shape makes it an oval or a rounded rectangle; it is refused for any other shape; the file keeps it', () => {
  const d = fresh();
  const id = d.addCameo(0, { x: 10, y: 10, w: 200, h: 200 });
  d.setCameoShape(0, id, 'ellipse');
  assert.equal(d.slide(0).shapes.find((s) => String(s.id) === String(id)).preset, 'ellipse');
  const other = d.addShape(0, { preset: 'rect', x: 300, y: 10, w: 50, h: 50 });
  assert.throws(() => d.setCameoShape(0, other, 'ellipse'), /not a cameo/);
  assert.throws(() => d.setCameoShape(0, id, 'heart'), /rectangle, an oval or a rounded rectangle/);
  const again = Deck.open(d.save());
  const kept = again.slide(0).shapes.find((s) => String(s.id) === String(id));
  assert.deepEqual([kept.cameo, kept.preset, kept.geometry.w], [true, 'ellipse', 200]);
});
