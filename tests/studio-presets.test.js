// Export presets by destination (docs/STUDIO.md X11): YouTube at 1080p
// 16:9, LinkedIn and X square or 16:9, Reels, Shorts and TikTok at 9:16, and
// a small MP4 for email; each the frame's shape, size, rate and bitrates, and
// the project made ready for it with its captions moved for the shape.
import test from 'node:test';
import assert from 'node:assert/strict';
import { EXPORT_PRESETS, presetById, presetFrame, presetBytes, presetWarnings, presetDoc } from '../packages/studio/src/presets.js';
import { placeCaptions, isCaption, CAPTION_PLACE } from '../packages/studio/src/captions.js';
import { emptyProject } from '../packages/studio/src/project.js';
import { buildPlan } from '../packages/studio/src/index.js';

test('every destination the plan names is there', () => {
  assert.deepEqual(EXPORT_PRESETS.map((p) => p.id), ['project', 'youtube', 'social-wide', 'social-square', 'vertical', 'email']);
  assert.equal(presetById('nope').id, 'project');
});

test('frames: 1080p 16:9 for YouTube, 1080 square, 1080 by 1920 upright, and a 480 line email in the project\'s own shape', () => {
  assert.deepEqual(presetFrame('youtube', 'vertical'), { aspect: 'landscape', width: 1920, height: 1080 });
  assert.deepEqual(presetFrame('social-square'), { aspect: 'square', width: 1080, height: 1080 });
  assert.deepEqual(presetFrame('vertical'), { aspect: 'vertical', width: 1080, height: 1920 });
  assert.deepEqual(presetFrame('email', 'landscape'), { aspect: 'landscape', width: 854, height: 480 });
  assert.deepEqual(presetFrame('email', 'vertical'), { aspect: 'vertical', width: 480, height: 854 });
  assert.deepEqual(presetFrame('project', 'square'), { aspect: 'square', width: 1080, height: 1080 });
});

test('sizes and warnings: an email\'s minute is a few megabytes, a long one is said to be too big, a long upright one too long', () => {
  const minute = presetBytes('email', 60);
  assert.ok(minute > 6e6 && minute < 9e6, `${minute}`);
  assert.equal(presetBytes('project', 60), null);
  assert.deepEqual(presetWarnings('email', 60), []);
  assert.deepEqual(presetWarnings('email', 200), ['big']);
  assert.deepEqual(presetWarnings('vertical', 200), ['long']);
  assert.deepEqual(presetWarnings('youtube', 2000), []);
});

test('the project made ready: its options carry the frame, rate and bitrates, the plan is that size, and captions move for an upright frame', () => {
  const doc = placeCaptions(emptyProject(), [{ start: 0, end: 2, text: 'hello' }]).doc;
  assert.equal(presetDoc(doc, 'project'), doc, 'the project as it is');
  const tall = presetDoc(doc, 'vertical');
  assert.equal(tall.options.aspect, 'vertical');
  assert.deepEqual(tall.options.frame, { width: 1080, height: 1920 });
  assert.equal(tall.options.bitrate, 8_000_000);
  assert.equal(tall.patches.find(isCaption).fy, CAPTION_PLACE.vertical.fy);
  assert.equal(doc.patches.find(isCaption).fy, CAPTION_PLACE.landscape.fy, 'the project itself untouched');
  let seen = null;
  presetDoc(doc, 'social-square', { relayout: (d, aspect) => { seen = aspect; return d; } });
  assert.equal(seen, 'square', 'a change of shape lays out again');
  seen = null;
  presetDoc(doc, 'youtube', { relayout: (d, aspect) => { seen = aspect; return d; } });
  assert.equal(seen, null, 'the same shape is left alone');
  const small = presetDoc(doc, 'email');
  const canvas = { width: 0, height: 0, getContext: () => new Proxy({}, { get: (o, k) => (k === 'measureText' ? () => ({ width: 10 }) : () => {}), set: () => true }) };
  const plan = buildPlan({ canvas, images: [], options: small.options, layerPatches: small.patches, videos: {}, assets: {} });
  assert.equal(plan.W, 854);
  assert.equal(plan.H, 480);
  assert.equal(plan.fps, 24);
  assert.equal(plan.bitrate, 900_000);
  const cap = plan.layers.find((l) => l.text === 'hello');
  assert.equal(cap.y, Math.round(480 * CAPTION_PLACE.landscape.fy * 1000) / 1000, 'fractions of the small frame');
});
