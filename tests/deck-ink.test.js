// Draw → Pen, Pencil, Highlighter, Ruler and Ink to Shape — the engine's
// half and the window's arithmetic.
//
// A stroke is kept as a freeform shape named "Ink N": a smooth path
// through its points (quadratic curves through the midpoints) in the pen's
// colour and width, see-through for a highlighter, with round or square
// ends — and drawn so. A stroke begun along the Ruler's edge is laid on it;
// Ink to Shape knows a drawn rectangle, oval and triangle, and leaves a
// scribble as ink.
import test from 'node:test';
import assert from 'node:assert/strict';
import { Deck, buildPptx, renderSlide } from '@rutba/presentation';
import { recognise, alongRuler, strokeLook, isInk } from '../apps/desktop/renderer/apps/slides/ink-geometry.js';

const blank = () => Deck.open(buildPptx({ title: 'Ink', slides: [{ layout: 'blank' }] }));

test('a stroke is kept as an ink shape: a smooth path in the pen\'s colour, width and tip', () => {
  const deck = blank();
  const ids = deck.addInk(0, [
    { points: [[100, 100], [120, 110], [140, 130], [160, 135]], color: '#E81123', width: 1, alpha: 1, cap: 'rnd' },
    { points: [[200, 300], [400, 300]], color: '#FFFF00', width: 8, alpha: 0.5, cap: 'sq' },
  ]);
  assert.equal(ids.length, 2);
  const xml = deck.pkg.text('ppt/slides/slide1.xml');
  assert.match(xml, /<p:cNvPr id="\d+" name="Ink 1"\/>/);
  assert.match(xml, /<p:cNvPr id="\d+" name="Ink 2"\/>/);
  assert.match(xml, /<a:path w="\d+" h="\d+" fill="none" extrusionOk="0"><a:moveTo><a:pt x="0" y="0"\/><\/a:moveTo><a:quadBezTo>/);
  assert.match(xml, /<a:noFill\/><a:ln w="12700" cap="rnd"><a:solidFill><a:srgbClr val="E81123"\/><\/a:solidFill><a:round\/><\/a:ln>/);
  assert.match(xml, /<a:ln w="101600" cap="sq"><a:solidFill><a:srgbClr val="FFFF00"><a:alpha val="50000"\/><\/a:srgbClr>/);
  const shapes = deck.slide(0).shapes;
  assert.ok(shapes.every(isInk));
  const pen = shapes.find((s) => s.name === 'Ink 1');
  assert.deepEqual([Math.round(pen.geometry.x), Math.round(pen.geometry.y), Math.round(pen.geometry.w), Math.round(pen.geometry.h)], [100, 100, 60, 35]);
  assert.equal(pen.line.cap, 'round');
  const svg = renderSlide(deck.slide(0), { width: 1280 });
  assert.match(svg, /stroke="#e81123"[^>]*stroke-linecap="round" stroke-linejoin="round"/i);
  assert.match(svg, /stroke-linecap="square"[^>]*stroke-opacity="0\.5"/);
  // A second batch numbers on.
  deck.addInk(0, [{ points: [[10, 10]], color: '#000000', width: 1 }]);
  assert.match(deck.pkg.text('ppt/slides/slide1.xml'), /name="Ink 3"/);
});

test('each pen lays its own ink', () => {
  assert.deepEqual(strokeLook({ tool: 'pen', color: '#000000', width: 1 }), { color: '#000000', width: 1, alpha: 1, cap: 'rnd' });
  assert.deepEqual(strokeLook({ tool: 'highlighter', color: '#FFFF00', width: 8 }), { color: '#FFFF00', width: 8, alpha: 0.5, cap: 'sq' });
  assert.equal(strokeLook({ tool: 'pencil', color: '#404040', width: 1 }).alpha, 0.85);
});

/** Points round a shape's outline, a little shaky as a hand draws. */
const around = (corners, per = 20) => {
  const out = [];
  for (let i = 0; i < corners.length; i++) {
    const a = corners[i];
    const b = corners[(i + 1) % corners.length];
    for (let k = 0; k < per; k++) out.push([a[0] + ((b[0] - a[0]) * k) / per + Math.sin(k) * 1.2, a[1] + ((b[1] - a[1]) * k) / per + Math.cos(k) * 1.2]);
  }
  out.push([corners[0][0] + 2, corners[0][1] + 1]);
  return out;
};

test('Ink to Shape knows a rectangle, an oval and a triangle, and leaves a scribble as ink', () => {
  const rect = recognise(around([[100, 100], [300, 100], [300, 220], [100, 220]]));
  assert.equal(rect?.preset, 'rect');
  assert.ok(Math.abs(rect.w - 200) < 8 && Math.abs(rect.h - 120) < 8);
  const tri = recognise(around([[200, 100], [300, 260], [100, 260]], 30));
  assert.equal(tri?.preset, 'triangle');
  const oval = [];
  for (let i = 0; i <= 80; i++) oval.push([400 + Math.cos((i / 80) * Math.PI * 2) * 120, 300 + Math.sin((i / 80) * Math.PI * 2) * 70]);
  assert.equal(recognise(oval)?.preset, 'ellipse');
  const scribble = Array.from({ length: 60 }, (_, i) => [100 + i * 6, 100 + (i % 2) * 40]);
  assert.equal(recognise(scribble), null, 'an open zigzag stays ink');
  assert.equal(recognise([[0, 0], [5, 5]]), null);
});

test('a stroke begun along the Ruler\'s edge is laid on it, a stroke away from it is left alone', () => {
  const ruler = { x: 640, y: 360, angle: 0, depth: 64 };
  // The top edge is at y = 328; a wobbly stroke along it becomes a straight line on it.
  const line = alongRuler([[300, 330], [400, 335], [500, 326], [700, 331]], ruler);
  assert.deepEqual(line.map((p) => p.map((v) => Math.round(v))), [[300, 328], [700, 328]]);
  assert.equal(alongRuler([[300, 200], [500, 210]], ruler), null);
  const turned = alongRuler([[640, 328], [700, 340]], { ...ruler, angle: 90 });
  assert.equal(turned, null, 'away from a turned ruler\'s edge');
});
