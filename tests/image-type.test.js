// Image: typing on the canvas (docs/STUDIO.md 7.2, I8). The in-place text
// box stands over the layer in the layer's own face, size, weight and colour,
// read from the layer as stored or compiled; a font picked becomes a stack
// the canvas draws with and falls back by kind; a size typed in pixels is
// kept as a fraction of the canvas's width; and the engine's box for a text
// layer, which the text box follows as words are typed, stays centred where
// the layer is and grows a line at a time, kept even while it is empty.
import test from 'node:test';
import assert from 'node:assert/strict';
import { faceOf, stackOf, inkOf, sizeOf, sizeFracOf, isBold } from '../packages/studio/src/type.js';
import { applyLayerPatches, layerBounds } from '../packages/studio/src/index.js';
import '../packages/studio/src/textblock.js';

test('a font stack names its face, and a face picked becomes a stack that falls back by kind', () => {
  assert.equal(faceOf('"Georgia", Georgia, serif'), 'Georgia');
  assert.equal(faceOf("'Trebuchet MS', sans-serif"), 'Trebuchet MS');
  assert.equal(faceOf(''), 'Segoe UI');
  assert.equal(faceOf(undefined), 'Segoe UI');
  assert.equal(stackOf('Georgia'), '"Georgia", Georgia, serif');
  assert.equal(stackOf('Times New Roman'), '"Times New Roman", Georgia, serif');
  assert.equal(stackOf('Verdana'), '"Verdana", "Segoe UI", system-ui, sans-serif');
  assert.equal(faceOf(stackOf('Century Gothic')), 'Century Gothic', 'a face goes round');
  assert.equal(stackOf('Bad"Name'), '"BadName", "Segoe UI", system-ui, sans-serif', 'no quote breaks the stack');
});

test('the words\' colour: a solid fill, a gradient\'s first stop, white when there is none', () => {
  assert.equal(inkOf({ fill: { kind: 'solid', color: '#123456' } }), '#123456');
  assert.equal(inkOf({ fill: { kind: 'linear', stops: [{ at: 0, color: '#ff0000' }, { at: 1, color: '#0000ff' }] } }), '#ff0000');
  assert.equal(inkOf({ fill: { kind: 'linear', stops: [[0, '#00ff00'], [1, '#000000']] } }), '#00ff00');
  assert.equal(inkOf({}), '#ffffff');
  assert.equal(inkOf(null), '#ffffff');
});

test('a size in pixels is kept as a fraction of the canvas\'s width, and read back the same', () => {
  assert.equal(sizeOf({ sizeFrac: 0.06 }, 1000), 60);
  assert.equal(sizeOf({}, 1000), 35, 'the painter\'s own default');
  assert.equal(sizeOf({ sizeFrac: 0.001 }, 1000), 8, 'never below eight, as the painter sets it');
  for (const [px, W] of [[48, 1200], [13, 640], [200, 4000], [72, 1920]]) assert.equal(sizeOf({ sizeFrac: sizeFracOf(px, W) }, W), px);
  assert.equal(sizeFracOf(3, 1000), 0.008, 'a size below eight is eight');
  assert.ok(isBold({ weight: 700 }) && isBold({ weight: 800 }) && !isBold({ weight: 600 }) && !isBold({ weight: 500 }) && !isBold({}), 'semibold is not bold, as in Word');
});

/* ── the box the text box follows ─────────────────────────────────────── */

const CHAR = 10;
const ctx = {
  font: '',
  measureText: (s) => ({ width: String(s).length * CHAR }),
};
const plan = () => ({
  W: 1000, H: 800, duration: 5, fps: 30,
  theme: { bg: '#000', scrim: 'rgba(0,0,0,.7)', text: '#fff', dim: '#aaa', accent: '#ffc107' },
  layers: [], images: [], videos: {}, assets: {}, context: {},
});
const boxOf = (patch) => {
  const p = plan();
  applyLayerPatches(p, [{ id: 'tb', type: 'textblock', fx: 0.5, fy: 0.5, fw: 0.4, sizeFrac: 0.04, ...patch }]);
  const layer = p.layers.find((l) => l.id === 'tb');
  return layer ? layerBounds(ctx, p, layer) : null;
};

test('the layer\'s box is centred where the layer is and grows a line at a time as lines are typed', () => {
  const one = boxOf({ text: 'Hello' });
  const two = boxOf({ text: 'Hello\nworld' });
  const three = boxOf({ text: 'Hello\nworld\nagain' });
  assert.equal(one.w, 400);
  assert.ok(Math.abs(one.x + one.w / 2 - 500) < 1e-6 && Math.abs(one.y + one.h / 2 - 400) < 1e-6, 'centred on the layer\'s place');
  const lead = two.h - one.h;
  assert.ok(lead > 40 && Math.abs((three.h - two.h) - lead) < 1e-6, `a line's height each time (${lead})`);
  assert.ok(Math.abs(three.y + three.h / 2 - 400) < 1e-6, 'still centred, growing both ways');
  // Words wider than the box wrap onto another line, as the text box wraps them.
  const long = boxOf({ text: 'a '.repeat(60).trim() });
  assert.ok(long.h > one.h * 2, 'wrapped');
});

test('an empty layer being typed in keeps its box; left empty it draws nothing', () => {
  assert.equal(boxOf({ text: '' }), null, 'no words, no layer');
  const kept = boxOf({ text: '', keepEmpty: true });
  assert.ok(kept && kept.w === 400 && kept.h >= 0, JSON.stringify(kept));
});
