// Export presets by destination (docs/STUDIO.md X11): YouTube at 1080p
// 16:9, LinkedIn and X square or 16:9, Reels, Shorts and TikTok at 9:16, and
// a small MP4 for email; each the frame's shape, size, rate and bitrates, and
// the project made ready for it with its captions moved for the shape.
import test from 'node:test';
import assert from 'node:assert/strict';
import { EXPORT_PRESETS, presetById, presetFrame, presetBytes, presetWarnings, presetDoc, projectFrame, fitBitrates, fitPlan, encodeToFit, megabytes, bitrateWords, parseMegabytes, poorBelow, FIT_AIM, FIT_FLOOR, FIT_TRIES, QUALITY_BPP } from '../packages/studio/src/presets.js';
import { placeCaptions, isCaption, CAPTION_PLACE } from '../packages/studio/src/captions.js';
import { emptyProject } from '../packages/studio/src/project.js';
import { buildPlan } from '../packages/studio/src/index.js';

test('every destination the plan names is there', () => {
  assert.deepEqual(EXPORT_PRESETS.map((p) => p.id), ['project', 'youtube', 'compact', 'small', 'social-wide', 'social-square', 'vertical', 'email']);
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

// Compact and Small, and Fit under a size (the 2.23 list, item 1).

test('Compact is 1080p at 4 Mb/s and Small 720p at 2 Mb/s, in the project\'s own shape', () => {
  assert.equal(presetById('compact').bitrate, 4_000_000);
  assert.equal(presetById('small').bitrate, 2_000_000);
  assert.deepEqual(presetFrame('compact', 'landscape'), { aspect: 'landscape', width: 1920, height: 1080 });
  assert.deepEqual(presetFrame('small', 'landscape'), { aspect: 'landscape', width: 1280, height: 720 });
  assert.deepEqual(presetFrame('small', 'vertical'), { aspect: 'vertical', width: 720, height: 1280 });
  assert.deepEqual(presetFrame('small', 'square'), { aspect: 'square', width: 720, height: 720 });
});

test('the sums: a preset\'s size is its bitrates over its length with the file\'s two percent', () => {
  assert.equal(presetBytes('compact', 85.6), Math.round(((4_000_000 + 128000) / 8) * 85.6 * 1.02));
  assert.equal(presetBytes('small', 85.6), Math.round(((2_000_000 + 128000) / 8) * 85.6 * 1.02));
  assert.equal(presetBytes('compact', 85.6, { sound: false }), Math.round((4_000_000 / 8) * 85.6 * 1.02));
});

test('the project\'s own preset has a size once its Quality is known, High, Medium and Low in order', () => {
  const frame = { width: 1920, height: 1080 };
  const at = (quality) => presetBytes('project', 60, { quality, frame, fps: 30 });
  assert.equal(at('high'), Math.round(((1920 * 1080 * 30 * QUALITY_BPP.high + 128000) / 8) * 60 * 1.02));
  assert.ok(at('high') > at('medium') && at('medium') > at('low'));
  assert.equal(presetBytes('project', 60), null, 'not without them');
});

test('the project\'s frame is its own size where it has one', () => {
  assert.deepEqual(projectFrame({ options: { aspect: 'square' } }), { aspect: 'square', width: 1080, height: 1080 });
  assert.deepEqual(projectFrame({ options: { aspect: 'landscape', frame: { width: 1281, height: 721 } } }), { aspect: 'landscape', width: 1282, height: 722 });
  assert.deepEqual(projectFrame({ options: {} }), { aspect: 'landscape', width: 1920, height: 1080 });
});

test('fit: the bitrate comes from the length and the sound\'s share, aimed 8% under', () => {
  const s = fitBitrates({ bytes: 50_000_000, seconds: 85.6 });
  assert.equal(FIT_AIM, 0.92);
  assert.equal(s.total, Math.floor((50_000_000 * 0.92 * 8) / (85.6 * 1.02)));
  assert.equal(s.audio, 128000, 'the sound keeps its rate at this size');
  assert.equal(s.video, s.total - s.audio);
  const made = ((s.video + s.audio) / 8) * 85.6 * 1.02;
  assert.ok(made <= 50_000_000 * 0.92 + 1 && made > 45_900_000, String(made));
  assert.equal(fitBitrates({ bytes: 50_000_000, seconds: 0 }), null);
  assert.equal(fitBitrates({ bytes: 0, seconds: 10 }), null);
});

test('fit: with no sound the picture has it all; where the sound would take over a quarter it drops to 96 kb/s', () => {
  const none = fitBitrates({ bytes: 5_000_000, seconds: 60, sound: false });
  assert.equal(none.audio, 0);
  assert.equal(none.video, none.total);
  const tight = fitBitrates({ bytes: 5_000_000, seconds: 200 });
  assert.equal(tight.audio, 96000);
  const email = fitBitrates({ bytes: 5_000_000, seconds: 200, audioBitrate: 96000 });
  assert.equal(email.audio, 96000, 'never above the destination\'s own');
});

test('fit plan: a size that binds lowers the picture, and says what it will come to, under the limit', () => {
  const f = fitPlan('compact', 85.6, { bytes: 40_000_000 });
  assert.equal(f.limit, 40_000_000);
  assert.ok(f.video < 4_000_000 && !f.capped);
  assert.ok(f.estimate <= 40_000_000 * 0.92 + 1 && f.estimate > 36_000_000, String(f.estimate));
});

test('fit plan: a size the preset is already under leaves the preset as it is, and says so', () => {
  const f = fitPlan('compact', 20, { bytes: 287_000_000 });
  assert.equal(f.capped, true);
  assert.equal(f.video, 4_000_000);
  assert.equal(f.poor, false);
  assert.ok(f.estimate < 11_000_000);
  // The project's own: its Quality is the ceiling.
  const own = fitPlan('project', 20, { bytes: 287_000_000, quality: 'low', frame: { aspect: 'landscape', width: 1920, height: 1080 }, fps: 30 });
  assert.equal(own.video, Math.round(1920 * 1080 * 30 * QUALITY_BPP.low));
  assert.equal(own.capped, true);
});

test('fit plan: under about 1 Mb/s at 1080p it warns and offers 720, which holds', () => {
  const at1080 = fitPlan('compact', 60, { bytes: 5_000_000 });
  assert.equal(at1080.poor, true);
  assert.equal(at1080.suggest, 720);
  assert.ok(at1080.video < 1_000_000);
  const at720 = fitPlan('compact', 60, { bytes: 5_000_000, short: 720 });
  assert.deepEqual([at720.frame.width, at720.frame.height], [1280, 720]);
  assert.equal(at720.short, 720);
  assert.equal(at720.poor, false, '720 is enough for this rate');
  // Just above the line: no warning.
  const fine = fitPlan('compact', 60, { bytes: 9_000_000, sound: false });
  assert.ok(fine.video > 1_000_000);
  assert.equal(fine.poor, false);
  assert.equal(poorBelow({ width: 1920, height: 1080 }), 1_000_000);
});

test('fit plan: a tall project keeps its shape at 720, and a thin 720 suggests 480', () => {
  const tall = fitPlan('vertical', 60, { bytes: 5_000_000, short: 720 });
  assert.deepEqual([tall.frame.width, tall.frame.height], [720, 1280]);
  const thin = fitPlan('small', 120, { bytes: 5_000_000 });
  assert.equal(thin.poor, true);
  assert.equal(thin.suggest, 480);
  assert.equal(fitPlan('email', 400, { bytes: 3_000_000, aspect: 'landscape' }).suggest, null, 'nothing smaller than 480');
});

test('fit plan: a size that leaves nothing for the picture says so', () => {
  const f = fitPlan('compact', 600, { bytes: 1_000_000 });
  assert.equal(f.tooSmall, true);
  assert.ok(f.video < FIT_FLOOR);
  assert.equal(f.poor, false, 'too small, not merely poor');
  assert.equal(fitPlan('compact', 0, { bytes: 1_000_000 }), null);
});

test('fit in the project: the bitrates reach the plan, on the project\'s own preset too, and the frame only when asked', () => {
  const doc = emptyProject();
  const canvas = { width: 0, height: 0, getContext: () => new Proxy({}, { get: (o, k) => (k === 'measureText' ? () => ({ width: 10 }) : () => {}), set: () => true }) };
  const planOf = (d) => buildPlan({ canvas, images: [], options: d.options, layerPatches: d.patches, videos: {}, assets: {} });
  const own = presetDoc(doc, 'project', { fit: { limit: 5e6, video: 1_500_000, audio: 96000, short: 0, frame: null } });
  assert.equal(own.options.bitrate, 1_500_000);
  assert.equal(own.options.audioBitrate, 96000);
  assert.equal(own.options.frame, undefined, 'the project\'s own frame');
  assert.equal(planOf(own).bitrate, 1_500_000);
  const lower = presetDoc(doc, 'project', { fit: { limit: 5e6, video: 800_000, audio: 96000, short: 720, frame: { aspect: 'landscape', width: 1280, height: 720 } } });
  assert.deepEqual([planOf(lower).W, planOf(lower).H], [1280, 720]);
  assert.equal(doc.options.bitrate, undefined, 'the project itself untouched');
  const compact = presetDoc(doc, 'compact', { fit: { limit: 5e6, video: 2_000_000, audio: 128000, short: 720, frame: { aspect: 'landscape', width: 1280, height: 720 } } });
  assert.deepEqual(compact.options.frame, { width: 1280, height: 720 });
  assert.equal(planOf(compact).bitrate, 2_000_000);
  assert.deepEqual(presetDoc(doc, 'compact').options.frame, { width: 1920, height: 1080 }, 'without a fit, the preset');
});

test('fit: a pass under the limit is done in one', async () => {
  const seen = [];
  const r = await encodeToFit({ fit: { limit: 5_000_000, video: 3_000_000, audio: 128000 }, encode: async (v) => { seen.push(v); return { size: 4_400_000 }; } });
  assert.deepEqual(seen, [3_000_000]);
  assert.deepEqual(r.tries, [{ video: 3_000_000, size: 4_400_000 }]);
  assert.equal(r.over, false);
});

test('fit: a file still over is encoded again, lower, and says so first', async () => {
  const seen = [];
  const told = [];
  let size = 5_600_000;
  const r = await encodeToFit({
    fit: { limit: 5_000_000, video: 3_000_000, audio: 128000 },
    encode: async (v) => { seen.push(v); const s = size; size = 4_500_000; return { size: s }; },
    onAgain: (x) => told.push(x),
  });
  assert.equal(seen.length, 2);
  assert.ok(seen[1] < seen[0] * 0.9 + 1, 'lower');
  assert.equal(told.length, 1);
  assert.deepEqual([told[0].attempt, told[0].size, told[0].limit, told[0].video], [1, 5_600_000, 5_000_000, seen[1]]);
  assert.equal(r.tries.length, 2);
  assert.equal(r.over, false);
  assert.equal(r.out.size, 4_500_000);
});

test('fit: an encoder that will go no lower is not asked a third time', async () => {
  const seen = [];
  const r = await encodeToFit({ fit: { limit: 5_000_000, video: 3_000_000, audio: 128000 }, encode: async (v) => { seen.push(v); return { size: 9_000_000 - seen.length * 20_000 }; } });
  assert.equal(seen.length, 2, 'the second go made it no smaller');
  assert.equal(r.over, true);
});

test('fit: after three goes what there is is handed over, marked still too big; the floor stops it sooner', async () => {
  let size = 9_000_000;
  const r = await encodeToFit({ fit: { limit: 5_000_000, video: 3_000_000, audio: 128000 }, encode: async () => { size -= 1_000_000; return { size }; } });
  assert.equal(FIT_TRIES, 3);
  assert.equal(r.tries.length, 3);
  assert.equal(r.over, true);
  for (let i = 1; i < r.tries.length; i++) assert.ok(r.tries[i].video < r.tries[i - 1].video);
  const low = await encodeToFit({ fit: { limit: 1_000_000, video: FIT_FLOOR, audio: 0 }, encode: async () => ({ size: 2_000_000 }) });
  assert.equal(low.tries.length, 1);
  assert.equal(low.over, true);
  // `first` is the checks' way of making the first pass come out over.
  const seen = [];
  await encodeToFit({ fit: { limit: 5_000_000, video: 3_000_000, audio: 0 }, first: 9_000_000, encode: async (v) => { seen.push(v); return { size: 100 }; } });
  assert.deepEqual(seen, [9_000_000]);
});

test('the words for sizes and rates, and a size typed', () => {
  assert.equal(megabytes(45_990_000), '46 MB');
  assert.equal(megabytes(4_600_000), '4.6 MB');
  assert.equal(megabytes(850_000), '850 KB');
  assert.equal(bitrateWords(4_400_000), '4.4 Mb/s');
  assert.equal(bitrateWords(850_000), '850 kb/s');
  assert.equal(parseMegabytes('50'), 50_000_000);
  assert.equal(parseMegabytes(' 12,5 '), 12_500_000);
  for (const bad of ['', 'abc', '0', '-3']) assert.equal(parseMegabytes(bad), 0, bad);
});
