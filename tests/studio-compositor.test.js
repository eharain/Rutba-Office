// The Recorder's scenes as boxes on the frame (packages/studio/src/compositor.js).

import test from 'node:test';
import assert from 'node:assert/strict';
import { SCENES, normaliseStudio, STUDIO_DEFAULTS, containRect, coverCrop, sceneLayout, titleBand, logoBox } from '../packages/studio/src/compositor.js';

const near = (a, b, what, eps = 0.5) => assert.ok(Math.abs(a - b) <= eps, `${what}: ${a} is not ${b}`);
const inside = (box, W, H, what) => assert.ok(box.x >= -0.5 && box.y >= -0.5 && box.x + box.w <= W + 0.5 && box.y + box.h <= H + 0.5, `${what} is inside the frame: ${JSON.stringify(box)}`);

test('settings: anything stored comes back as settings the Recorder can use', () => {
  assert.deepEqual(normaliseStudio(null), { ...STUDIO_DEFAULTS });
  const s = normaliseStudio({ scene: 'nonsense', size: 4000, fps: 13, title: 'x'.repeat(500), corner: 'tl', titleOn: 1 });
  assert.equal(s.scene, 'screen-camera');
  assert.equal(s.size, 1080);
  assert.equal(s.fps, 30);
  assert.equal(s.title.length, 120);
  assert.equal(s.corner, 'tl');
  assert.equal(s.titleOn, true);
  for (const scene of SCENES) assert.equal(normaliseStudio({ scene }).scene, scene);
});

test('a picture fits a box whole, or fills it cropped from the middle', () => {
  const c = containRect(1920, 1200, { x: 0, y: 0, w: 1920, h: 1080 });
  near(c.h, 1080, 'height');
  near(c.w, 1728, 'width');
  near(c.x, 96, 'pillarboxed');
  const crop = coverCrop(1280, 720, { x: 0, y: 0, w: 400, h: 400 });
  near(crop.sw, 720, 'a square cut from a wide picture');
  near(crop.sx, 280, 'from its middle');
  near(crop.sh, 720, 'the whole height');
});

test('screen and camera: the screen fills the frame and the camera sits in its corner, shaped and mirrored as set', () => {
  const W = 1920; const H = 1080;
  const l = sceneLayout('screen-camera', { W, H, screen: { w: 2560, h: 1440 }, camera: { w: 1280, h: 720 }, settings: { corner: 'tl', shape: 'circle', mirror: true } });
  assert.equal(l.layers.length, 2);
  assert.equal(l.layers[0].source, 'screen');
  assert.deepEqual(l.layers[0].box, { x: 0, y: 0, w: W, h: H });
  const cam = l.layers[1];
  assert.equal(cam.source, 'camera');
  assert.equal(cam.shape, 'circle');
  assert.equal(cam.mirror, true);
  assert.ok(cam.box.x < W / 4 && cam.box.y < H / 4, 'top left');
  near(cam.box.w, cam.box.h, 'a circle is square');
  inside(cam.box, W, H, 'the camera');
});

test('each scene shows what there is: no camera, no screen, a card', () => {
  const W = 1280; const H = 720;
  assert.deepEqual(sceneLayout('screen-camera', { W, H, screen: { w: 1920, h: 1080 } }).layers.map((x) => x.source), ['screen']);
  assert.deepEqual(sceneLayout('screen-camera', { W, H, camera: { w: 1280, h: 720 } }).layers.map((x) => x.source), ['camera']);
  assert.deepEqual(sceneLayout('camera', { W, H, screen: { w: 1920, h: 1080 } }).layers.map((x) => x.source), ['screen'], 'no camera: the screen');
  assert.deepEqual(sceneLayout('screen', { W, H, camera: { w: 1, h: 1 } }).layers.map((x) => x.source), ['camera'], 'no screen: the camera');
  const brb = sceneLayout('brb', { W, H, screen: { w: 1, h: 1 }, settings: { brb: 'Back in five' } });
  assert.equal(brb.layers.length, 0);
  assert.equal(brb.card.text, 'Back in five');
});

test('split: the screen on the left two thirds, the camera beside it, neither over the other', () => {
  const W = 1920; const H = 1080;
  const l = sceneLayout('split', { W, H, screen: { w: 1920, h: 1080 }, camera: { w: 1280, h: 720 } });
  const [scr, cam] = l.layers;
  assert.equal(scr.source, 'screen');
  assert.equal(cam.source, 'camera');
  assert.ok(scr.box.x + scr.box.w <= cam.box.x, 'side by side');
  assert.ok(scr.box.w > cam.box.w * 1.8, 'the screen is the larger');
  inside(scr.box, W, H, 'the screen');
  inside(cam.box, W, H, 'the camera');
  assert.equal(l.background, 'studio');
});

test('the title band sits low on the left, and a logo keeps its shape in its corner', () => {
  const W = 1920; const H = 1080;
  const b = titleBand({ W, H });
  inside(b.band, W, H, 'the band');
  assert.ok(b.band.y > H * 0.7, 'low');
  assert.ok(b.title.size > b.subtitle.size, 'the title larger than the line under it');
  const logo = logoBox({ W, H, w: 400, h: 100, corner: 'tr' });
  near(logo.w / logo.h, 4, 'its shape');
  assert.ok(logo.w <= W * 0.1 + 1, 'at most a tenth of the width');
  assert.ok(logo.x > W / 2 && logo.y < H / 2, 'top right');
  inside(logo, W, H, 'the logo');
});
