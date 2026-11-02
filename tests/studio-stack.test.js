// The Recorder's studio as a stack of items (packages/studio/src/stack.js).

import test from 'node:test';
import assert from 'node:assert/strict';
import {
  emptyStack, normaliseStack, stackFromSettings, addItem, removeItem, duplicateItem, moveItem, nudgeItem, showItem,
  placeItem, placeBox, shapeItem, clipItem, frameToSource, itemAt, drawOrder, sceneOf, saveScene, applyScene,
  builtInScenes, recordKind, updateSound, PLACES,
} from '../packages/studio/src/stack.js';

const near = (a, b, what, eps = 0.002) => assert.ok(Math.abs(a - b) <= eps, `${what}: ${a} is not ${b}`);
const inside = (b, what) => assert.ok(b.x >= 0 && b.y >= 0 && b.x + b.w <= 1.0001 && b.y + b.h <= 1.0001, `${what} is inside the frame: ${JSON.stringify(b)}`);
const ids = (s) => s.items.map((it) => it.id);
const on = (s) => s.items.filter((it) => it.on).map((it) => it.id);

/** A studio with a screen at the bottom, a camera over it, and a second camera not shown. */
function studio() {
  let s = emptyStack();
  ({ stack: s } = addItem(s, { kind: 'screen', ref: 'screen:1', name: 'Screen 1' }));
  ({ stack: s } = addItem(s, { kind: 'camera', ref: 'cam-a', name: 'USB camera' }));
  ({ stack: s } = addItem(s, { kind: 'camera', ref: 'cam-b', name: 'Laptop camera', on: false }));
  return s;
}

test('anything stored comes back as a stack, with unknown kinds dropped and ids made unique', () => {
  const s = normaliseStack({ items: [{ kind: 'screen', id: 'a' }, { kind: 'nonsense' }, { kind: 'camera', id: 'a', box: { x: 5, y: -1, w: 3, h: 0 } }], sounds: [{ kind: 'mic' }, { kind: 'speaker' }], video: 'sideways' });
  assert.equal(s.items.length, 2);
  assert.notEqual(s.items[0].id, s.items[1].id);
  const b = s.items[1].box;
  inside(b, 'a box from nonsense');
  assert.equal(s.sounds.length, 1);
  assert.equal(s.video, 'on');
  assert.deepEqual(normaliseStack(null), emptyStack());
});

test('every source is added the same way: on top and shown; a screen fills, a camera goes in a corner over it', () => {
  const s = studio();
  assert.deepEqual(ids(s), ['camera-1', 'screen-1', 'camera-2'], 'shown first, top first; the one not shown after them');
  assert.deepEqual(on(s), ['camera-1', 'screen-1']);
  assert.deepEqual(s.items.find((it) => it.id === 'screen-1').box, { x: 0, y: 0, w: 1, h: 1 });
  const cam = s.items.find((it) => it.id === 'camera-1');
  assert.ok(cam.box.w < 0.5 && cam.box.x > 0.5 && cam.box.y > 0.5, 'the camera sits in the bottom right corner');
  assert.equal(cam.mirror, true);
  assert.deepEqual(drawOrder(s).map((it) => it.id), ['screen-1', 'camera-1'], 'drawn bottom first');
});

test('Place puts any item where it is named, inside the frame, a circle round on the frame', () => {
  for (const place of PLACES) {
    for (const shape of ['rect', 'circle']) {
      const b = placeBox(place, { aspect: 16 / 9, shape });
      inside(b, `${place} ${shape}`);
      if (shape === 'circle') near(b.w * 16, b.h * 9, `${place}: as wide as it is tall in pixels`);
    }
  }
  const left = placeBox('left');
  const right = placeBox('right');
  assert.ok(left.x + left.w <= right.x, 'left and right do not overlap');
  // A screen can be a circle on top of a camera that fills the frame (owner, 2 Nov).
  let s = studio();
  s = placeItem(s, 'camera-1', 'fill');
  s = shapeItem(s, 'screen-1', 'circle');
  s = placeItem(s, 'screen-1', 'bl');
  s = moveItem(s, 'screen-1', 0);
  assert.deepEqual(drawOrder(s).map((it) => it.id), ['camera-1', 'screen-1']);
  const scr = s.items[0];
  assert.equal(scr.shape, 'circle');
  near(scr.box.w * 16, scr.box.h * 9, 'the screen circle is round');
});

test('a clip works on a camera as on a screen, and the box takes the clipped picture\'s shape', () => {
  let s = studio();
  s = clipItem(s, 'camera-1', { x: 0.25, y: 0.25, w: 0.5, h: 0.25 });
  const cam = s.items.find((it) => it.id === 'camera-1');
  assert.deepEqual(cam.crop, { x: 0.25, y: 0.25, w: 0.5, h: 0.25 });
  // 16:9 clipped to half by a quarter is 32:9: the box is half as tall for the width.
  near((cam.box.w * 16) / (cam.box.h * 9), 32 / 9, 'the box is as wide as the clip');
  s = clipItem(s, 'camera-1', { x: 0, y: 0, w: 1, h: 1 });
  assert.equal(s.items.find((it) => it.id === 'camera-1').crop, null, 'the whole picture is no clip');
  // A point on the frame, through the clip and the mirror, to the source.
  s = clipItem(studio(), 'screen-1', { x: 0.5, y: 0.5, w: 0.5, h: 0.5 });
  const p = frameToSource(s.items.find((it) => it.id === 'screen-1'), 0.5, 0.5);
  near(p.x, 0.75, 'x through the clip');
  near(p.y, 0.75, 'y through the clip');
});

test('a point on the share is found on the source as it is drawn: filling a circle, inside a rectangle', () => {
  // A 16:9 camera in a square circle shows its middle 9/16 across: the circle's left edge is 7/32 in.
  const circle = { box: { x: 0.5, y: 0, w: 9 / 16, h: 1 }, crop: null, shape: 'circle', mirror: false };
  near(frameToSource(circle, 0.5, 0.5, { aspect: 16 / 9 }).x, 7 / 32, 'the left edge of the circle');
  near(frameToSource(circle, 0.5 + 9 / 32, 0.5, { aspect: 16 / 9 }).x, 0.5, 'its middle');
  near(frameToSource({ ...circle, mirror: true }, 0.5, 0.5, { aspect: 16 / 9 }).x, 25 / 32, 'mirrored, the other edge');
  // A 4:3 source in the whole 16:9 frame is drawn 3/4 wide, centred: the frame's 1/8 is its left edge.
  const rect = { box: { x: 0, y: 0, w: 1, h: 1 }, crop: null, shape: 'rect', mirror: false };
  near(frameToSource(rect, 0.125, 0.5, { aspect: 4 / 3 }).x, 0, 'the left edge of the picture');
  near(frameToSource(rect, 0.5, 0.5, { aspect: 4 / 3 }).x, 0.5, 'its middle');
  near(frameToSource(rect, 0.05, 0.5, { aspect: 4 / 3 }).x, 0, 'the bars beside it are its edge');
});

test('the list reorders the stack: dragged among the shown ones it is shown, among the rest hidden', () => {
  let s = studio();
  assert.deepEqual(ids(s), ['camera-1', 'screen-1', 'camera-2']);
  s = moveItem(s, 'screen-1', 0);
  assert.deepEqual(ids(s), ['screen-1', 'camera-1', 'camera-2']);
  s = moveItem(s, 'camera-2', 0);
  assert.deepEqual(on(s), ['camera-2', 'screen-1', 'camera-1'], 'dragged up among the shown, it is shown');
  s = moveItem(s, 'camera-1', 2);
  assert.deepEqual(on(s), ['camera-2', 'screen-1', 'camera-1'], 'dragged to the foot of the shown ones, it is at the bottom of the stack');
  s = nudgeItem(s, 'screen-1', -1);
  assert.deepEqual(ids(s), ['screen-1', 'camera-2', 'camera-1']);
  s = showItem(s, 'camera-1', true);
  assert.deepEqual(ids(s), ['camera-1', 'screen-1', 'camera-2'], 'the eye puts it on top');
  s = showItem(s, 'camera-1', false);
  assert.deepEqual(ids(s), ['screen-1', 'camera-2', 'camera-1'], 'hidden, it goes to the top of the rest');
  s = moveItem(s, 'screen-1', 2);
  assert.deepEqual(on(s), ['camera-2'], 'dragged among the ones not shown, it is hidden');
});

test('the item on top at a point is found, a circle only inside its round', () => {
  let s = studio();
  s = shapeItem(s, 'camera-1', 'circle');
  const cam = s.items.find((it) => it.id === 'camera-1');
  const mid = itemAt(s, cam.box.x + cam.box.w / 2, cam.box.y + cam.box.h / 2);
  assert.equal(mid.id, 'camera-1');
  assert.equal(itemAt(s, cam.box.x + 0.001, cam.box.y + 0.001).id, 'screen-1', 'the corner outside the circle is the screen under it');
  assert.equal(itemAt(s, 0.1, 0.1).id, 'screen-1');
});

test('a second copy of a source can be clipped apart from the first', () => {
  let s = studio();
  let copy;
  ({ stack: s, id: copy } = duplicateItem(s, 'screen-1'));
  s = clipItem(s, copy, { x: 0, y: 0, w: 0.3, h: 0.3 });
  assert.equal(s.items.find((it) => it.id === 'screen-1').crop, null);
  assert.ok(s.items.find((it) => it.id === copy).crop);
  assert.equal(s.items.find((it) => it.id === copy).ref, 'screen:1');
  s = removeItem(s, copy);
  assert.ok(!ids(s).includes(copy));
});

test('a scene keeps the shown items\' places and order, and puts them back', () => {
  let s = studio();
  s = saveScene(s, 'Talk');
  s = applyScene(s, builtInScenes(s).find((sc) => sc.key === 'camera'));
  assert.deepEqual(on(s), ['camera-1']);
  assert.deepEqual(s.items[0].box, { x: 0, y: 0, w: 1, h: 1 });
  s = applyScene(s, s.scenes[0]);
  assert.deepEqual(on(s), ['camera-1', 'screen-1']);
  assert.deepEqual(sceneOf(s, 'x').layers.map((l) => l.id), ['camera-1', 'screen-1']);
  const split = builtInScenes(s).find((sc) => sc.key === 'split');
  s = applyScene(s, split);
  const [a, b] = s.items.filter((it) => it.on);
  assert.ok(a.box.x > b.box.x + b.box.w - 0.001 || b.box.x > a.box.x + a.box.w - 0.001, 'side by side');
  assert.deepEqual(builtInScenes(normaliseStack({ items: [{ kind: 'camera', id: 'c' }] })).map((sc) => sc.key), ['camera']);
});

test('the Recorder\'s old settings come back as the stack they meant', () => {
  const screen = { id: 'window:7', name: 'Mail', kind: 'window' };
  const s = stackFromSettings({ scene: 'screen-camera', camera: true, corner: 'tl', shape: 'circle', insetSize: 'small', titleOn: true, title: 'Imogen', logo: 'C:/logo.png', logoCorner: 'tr', mic: true, system: true }, { screen });
  assert.deepEqual(s.items.map((it) => it.kind), ['picture', 'title', 'camera', 'window']);
  assert.deepEqual(drawOrder(s).map((it) => it.kind), ['window', 'camera', 'title', 'picture']);
  const cam = s.items[2];
  assert.equal(cam.shape, 'circle');
  assert.ok(cam.box.x < 0.2 && cam.box.y < 0.2, 'top left');
  assert.equal(s.items[1].text, 'Imogen');
  assert.deepEqual(s.sounds.map((x) => [x.kind, x.live]), [['mic', true], ['system', true]]);
  const brb = stackFromSettings({ scene: 'brb', brb: 'Back soon', camera: false }, { screen });
  assert.equal(brb.video, 'placeholder');
  assert.equal(brb.placeholderText, 'Back soon');
  const camOnly = stackFromSettings({ scene: 'camera', camera: true }, { screen });
  assert.deepEqual(drawOrder(camOnly).map((it) => it.kind), ['camera']);
});

test('Record makes a video, sound alone when the picture is off, and nothing with both off', () => {
  let s = emptyStack();
  assert.equal(recordKind(s), 'video');
  s = { ...s, video: 'placeholder', sound: 'music' };
  assert.equal(recordKind(s), 'video');
  s = { ...s, video: 'off', sound: 'on' };
  assert.equal(recordKind(s), 'sound');
  s = { ...s, sound: 'off' };
  assert.equal(recordKind(s), null);
});

test('a sound is muted and its level set without touching the others', () => {
  let s = normaliseStack({ sounds: [{ kind: 'mic', id: 'mic-1', live: true }, { kind: 'system', id: 'system-1', live: true }] });
  s = updateSound(s, 'mic-1', { muted: true, gain: 5 });
  assert.deepEqual(s.sounds.map((x) => [x.id, x.muted, x.gain]), [['mic-1', true, 2], ['system-1', false, 1]]);
});
