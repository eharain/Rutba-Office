// A scene's tall layout (packages/studio/src/tall.js): made from the wide one,
// the sources kept, stacked and cropped to fit, and a layout kept by hand.

import test from 'node:test';
import assert from 'node:assert/strict';
import { makeTallLayers, cropToFill, tallSize, tallStackOf, setTallLayers, tallLayersFor, normaliseTallLayouts, clipAspect } from '../packages/studio/src/tall.js';
import { normaliseStack, addItem, emptyStack } from '../packages/studio/src/stack.js';

const item = (id, kind, on = true, more = {}) => ({ id, kind, on, ref: id, name: '', live: true, box: { x: 0, y: 0, w: 1, h: 1 }, crop: null, shape: 'rect', mirror: false, text: '', subtitle: '', path: '', ...more });
const near = (a, b, eps = 0.01) => assert.ok(Math.abs(a - b) < eps, `${a} is near ${b}`);

test('the tall sizes', () => {
  assert.deepEqual(tallSize(720), { width: 720, height: 1280 });
  assert.deepEqual(tallSize(1080), { width: 1080, height: 1920 });
  assert.deepEqual(tallSize(999), { width: 1080, height: 1920 });
});

test('a crop to fill is cut about the middle of the clip', () => {
  const c = cropToFill(16 / 9, null, 9 / 16);
  near(c.w, (9 / 16) / (16 / 9));
  near(c.x, (1 - c.w) / 2);
  assert.equal(c.h, 1);
  assert.equal(cropToFill(16 / 9, null, 16 / 9), null, 'already that shape');
  const tall = cropToFill(9 / 16, null, 16 / 9);
  assert.equal(tall.w, 1);
  near(tall.y, (1 - tall.h) / 2);
});

test('one picture fills the tall frame, cropped', () => {
  const [l] = makeTallLayers([item('cam', 'camera')], { aspectOf: () => 16 / 9 });
  assert.deepEqual(l.box, { x: 0, y: 0, w: 1, h: 1 });
  near(clipAspect(16 / 9, l.crop), 9 / 16);
  assert.equal(l.shape, 'rect');
});

test('a screen and a camera: the screen whole across the top, the camera cropped to the rest', () => {
  const layers = makeTallLayers([item('c', 'camera', true, { box: { x: 0.7, y: 0.6, w: 0.25, h: 0.3 }, shape: 'circle' }), item('s', 'screen')], { aspectOf: () => 16 / 9 });
  const screen = layers.find((l) => l.id === 's');
  const camera = layers.find((l) => l.id === 'c');
  assert.equal(screen.box.y, 0);
  near(screen.box.h, (9 / 16) / (16 / 9));
  assert.equal(screen.crop, null, 'shown whole');
  near(camera.box.y, screen.box.h);
  near(camera.box.y + camera.box.h, 1);
  near(clipAspect(16 / 9, camera.crop), (9 / 16) / camera.box.h);
  assert.equal(camera.shape, 'rect');
});

test('things that are not shown stay out; titles and pictures go over the pictures', () => {
  const layers = makeTallLayers([
    item('t', 'title', true, { text: 'Hi' }),
    item('logo', 'picture'),
    item('hid', 'camera', false),
    item('s', 'screen'),
  ], { aspectOf: (it) => (it.kind === 'picture' ? 1 : 16 / 9) });
  assert.deepEqual(layers.map((l) => l.id), ['t', 'logo', 's']);
  assert.ok(layers[0].box.y > 0.8);
  assert.ok(layers[1].box.x > 0.7);
});

test('nothing shown is nothing', () => {
  assert.deepEqual(makeTallLayers([item('s', 'screen', false)]), []);
});

test('a tall stack: same items in tall places, hidden ones hidden, a web page drawn from its tall render', () => {
  const stack = { ...normaliseStack(null), items: [item('w', 'web'), item('c', 'camera', false)] };
  const tall = tallStackOf(stack, 'x', { aspectOf: () => 16 / 9 });
  assert.equal(tall.items[0].id, 'w');
  assert.equal(tall.items[0].tall, true);
  assert.equal(tall.items[1].on, false);
});

test('a layout changed by hand is kept under its scene, and only while it covers what is shown', () => {
  let stack = { ...emptyStack(), items: [item('s', 'screen')] };
  const mine = [{ id: 's', box: { x: 0, y: 0.25, w: 1, h: 0.5 }, crop: null, shape: 'rect' }];
  stack = setTallLayers(stack, 'scene-1', mine);
  assert.deepEqual(tallLayersFor(stack, 'scene-1', {}).map((l) => l.box), [mine[0].box]);
  assert.notDeepEqual(tallLayersFor(stack, 'other', {})[0].box, mine[0].box, 'another scene makes its own');
  const more = { ...stack, items: [...stack.items, item('c', 'camera')] };
  assert.equal(tallLayersFor(more, 'scene-1', {}).length, 2, 'a new source falls back to a made layout');
  assert.deepEqual(setTallLayers(stack, 'scene-1', null).tall, {});
});

test('tall layouts survive being stored; anything else is dropped', () => {
  const raw = { 'own-0': [{ id: 's', box: { x: 0, y: 0, w: 1, h: 0.5 }, crop: null, shape: 'nope' }, 5], bad: 'x' };
  const out = normaliseTallLayouts(raw);
  assert.deepEqual(Object.keys(out), ['own-0']);
  assert.equal(out['own-0'].length, 1);
  assert.equal(out['own-0'][0].shape, 'rect');
  const stored = normaliseStack({ ...emptyStack(), tall: raw });
  assert.deepEqual(stored.tall['own-0'][0].box, { x: 0, y: 0, w: 1, h: 0.5 });
});

test('a web page is a kind of source the stack keeps by its address', () => {
  const { stack, id } = addItem(emptyStack(), { kind: 'web', ref: 'https://shop.example/show/products/summer', name: 'Summer' });
  const it = stack.items.find((x) => x.id === id);
  assert.equal(it.kind, 'web');
  assert.equal(it.ref, 'https://shop.example/show/products/summer');
  assert.equal(it.mirror, false);
  assert.deepEqual(it.box, { x: 0, y: 0, w: 1, h: 1 });
  const again = normaliseStack({ ...stack });
  assert.equal(again.items[0].kind, 'web');
  assert.equal(again.items[0].sound, false, 'sound off until turned on');
});

test('a tall layout is kept under the scene, or under one key for a frame arranged by hand', async () => {
  const { tallKey } = await import('../packages/studio/src/tall.js');
  assert.equal(tallKey({ scene: 'screen-camera' }), 'screen-camera');
  assert.equal(tallKey({ scene: '' }), 'free');
});

test('the copy is wide unless Tall was chosen, and is kept with the setup', async () => {
  const { normaliseStudio } = await import('../packages/studio/src/compositor.js');
  const { SETUP_KEYS } = await import('../packages/studio/src/setup.js');
  assert.equal(normaliseStudio({}).liveCopyShape, 'wide');
  assert.equal(normaliseStudio({ liveCopyShape: 'tall' }).liveCopyShape, 'tall');
  assert.equal(normaliseStudio({ liveCopyShape: 'square' }).liveCopyShape, 'wide');
  assert.ok(SETUP_KEYS.includes('liveCopyShape'));
});
