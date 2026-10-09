// Format Shape's pattern fills and 3-D — the engine's half.
//
// A pattern fill is one of DrawingML's fifty-four presets in a foreground
// and a background colour (`a:pattFill`), drawn as Office's eight-by-eight
// tile — it used to be drawn as its foreground alone. A shape's 3-D is a top
// bevel, a depth and the camera it is seen through (`a:sp3d`, `a:scene3d`):
// written as PowerPoint writes them, read back, and drawn — the bevel lit
// from the top left, a rotation as its flat projection with the depth
// stepped out behind.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { PATTERNS, patternDef, isPattern } from '@rutba/presentation/patterns';
import { BEVELS, CAMERAS, shape3dXml, cameraLook } from '@rutba/presentation/shape3d';

const fresh = () => {
  const deck = Deck.open(buildPptx({ title: '3-D', slides: [{ layout: 'blank' }] }));
  const id = deck.addShape(0, { preset: 'roundRect', x: 100, y: 100, w: 300, h: 200 });
  return { deck, id, shape: () => deck.slide(0).shapes.find((s) => String(s.id) === String(id)), xml: () => deck.pkg.text('ppt/slides/slide1.xml') };
};

test('the fifty-four presets each make a tile of their own, foreground on background', () => {
  assert.equal(PATTERNS.length, 54);
  assert.equal(new Set(PATTERNS.map(([n]) => n)).size, 54);
  const tiles = new Set();
  for (const [name] of PATTERNS) {
    const def = patternDef('t', name, '#000', '#fff');
    assert.match(def, /^<pattern id="t" patternUnits="userSpaceOnUse" width="8" height="8"><rect width="8" height="8" fill="#fff"\/><g fill="#000">/);
    tiles.add(def.replace(/id="t"/, ''));
  }
  assert.ok(tiles.size >= 50, 'the presets look different from one another');
  const dots = (name) => (patternDef('t', name).match(/width="(\d+)" height="1"/g) || []).map((m) => Number(/\d+/.exec(m)[0])).reduce((a, b) => a + b, 0);
  assert.ok(dots('pct10') < dots('pct50') && dots('pct50') < dots('pct90'), 'the percentages get darker');
  assert.equal(isPattern('dkUpDiag'), true);
  assert.equal(isPattern('nonsense'), false);
});

test('a pattern fill is written as PowerPoint writes one, read back, and drawn as its tile', () => {
  const { deck, id, shape, xml } = fresh();
  deck.setShapeStyle(0, id, { fill: { pattern: { preset: 'dkUpDiag', fg: '#C00000', bg: { scheme: 'bg1' } } } });
  assert.match(xml(), /<a:pattFill prst="dkUpDiag"><a:fgClr><a:srgbClr val="C00000"\/><\/a:fgClr><a:bgClr><a:schemeClr val="bg1"\/><\/a:bgClr><\/a:pattFill>/);
  const fill = shape().fill;
  assert.equal(fill.type, 'pattern');
  assert.equal(fill.preset, 'dkUpDiag');
  assert.equal(fill.color.toLowerCase(), '#c00000');
  const svg = renderSlide(deck.slide(0), { width: 1280 });
  assert.match(svg, /<pattern id="[^"]*pat\d+" patternUnits="userSpaceOnUse"[\s\S]*?<g fill="#c00000">/i);
  assert.match(svg, /fill="url\(#[^"]*pat\d+\)"/);
  assert.throws(() => deck.setShapeStyle(0, id, { fill: { pattern: { preset: 'nonsense' } } }), /not a pattern/);
});

test('a bevel, a depth and a rotation are written as PowerPoint writes them, read back and drawn', () => {
  const { deck, id, shape, xml } = fresh();
  deck.setShape3d(0, id, { bevel: { prst: 'angle', w: 6, h: 4 }, depth: 18, depthColor: '#404040', camera: 'isometricRightUp' });
  assert.match(xml(), /<a:ln\b[\s\S]*?<\/a:ln><a:scene3d><a:camera prst="isometricRightUp"\/><a:lightRig rig="threePt" dir="t"\/><\/a:scene3d><a:sp3d extrusionH="228600"><a:bevelT w="76200" h="50800" prst="angle"\/><a:extrusionClr><a:srgbClr val="404040"\/><\/a:extrusionClr><\/a:sp3d><\/p:spPr>/);
  assert.deepEqual(shape().shape3d, { bevel: { prst: 'angle', w: 6, h: 4 }, depth: 18, depthColor: '#404040', camera: 'isometricRightUp' });
  const svg = renderSlide(deck.slide(0), { width: 1280 });
  assert.match(svg, /<feDiffuseLighting[\s\S]*<feSpecularLighting/, 'the bevel lit');
  assert.match(svg, /<g transform="matrix\(0\.866 -0\.5 0 1 /, 'seen through the camera');
  assert.equal((svg.match(/fill="#404040"/g) || []).length, 24, 'the depth stepped out behind in its colour');
  // Back to flat and front on: nothing left in the file.
  deck.setShape3d(0, id, { bevel: null, depth: 0, camera: 'orthographicFront' });
  assert.ok(!/<a:(?:scene3d|sp3d)\b/.test(xml()));
  assert.equal(shape().shape3d, null);
  assert.throws(() => deck.setShape3d(0, id, { camera: 'sideways' }), /not a 3-D rotation/);
  assert.throws(() => deck.setShape3d(0, id, { bevel: { prst: 'wobbly' } }), /not a bevel/);
});

test('a bevel PowerPoint wrote is read at its default size, and the presets cover the galleries', () => {
  assert.equal(BEVELS.length, 12);
  assert.ok(CAMERAS.some(([p]) => p === 'perspectiveFront'));
  assert.deepEqual(cameraLook('somethingElse').matrix, [1, 0, 0, 1], 'an unknown camera is drawn front on');
  assert.equal(shape3dXml({}), '', 'flat and front on says nothing');
  const { deck, id } = fresh();
  const part = 'ppt/slides/slide1.xml';
  deck.pkg.write_(part, deck.pkg.text(part).replace('</a:ln></p:spPr>', '</a:ln><a:sp3d><a:bevelT/></a:sp3d></p:spPr>'));
  const reopened = Deck.open(deck.pkg.write());
  const read = reopened.slide(0).shapes.find((s) => String(s.id) === String(id)).shape3d;
  assert.deepEqual(read.bevel, { prst: 'circle', w: 6, h: 6 });
});

test('two 3-D changes sent one after the other both stand, the second keeping what the first set', async () => {
  // The window sends only what changed; the service fills the rest from the
  // deck as it is now. It used to send the whole from what it last saw, so
  // the camera chosen before the depth came back put the depth back to 0.
  const fs = await import('node:fs');
  const os = await import('node:os');
  const path = await import('node:path');
  const { createDocumentService } = await import('../apps/desktop/main/documents.js');
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rutba-3d-')), 'three.pptx');
  const { deck, id } = fresh();
  fs.writeFileSync(file, deck.save());
  const docs = createDocumentService({ holdBlob: () => ({ url: 'blob:x' }) });
  const { id: doc } = await docs.open({ path: file });
  const three = () => docs.model({ id: doc, slide: 0 }).slide.shapes.find((s) => String(s.id) === String(id)).shape3d;
  docs.apply({ id: doc, ops: [{ op: 'setShape3d', slide: 0, shape: id, patch: { bevel: { prst: 'circle', w: 6, h: 6 } } }] });
  docs.apply({ id: doc, ops: [{ op: 'setShape3d', slide: 0, shape: id, patch: { depth: 18 } }] });
  docs.apply({ id: doc, ops: [{ op: 'setShape3d', slide: 0, shape: id, patch: { camera: 'isometricRightUp' } }] });
  assert.deepEqual({ bevel: three().bevel, depth: three().depth, camera: three().camera }, { bevel: { prst: 'circle', w: 6, h: 6 }, depth: 18, camera: 'isometricRightUp' });
});
