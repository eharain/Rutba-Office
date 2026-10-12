// Captions (docs/STUDIO.md V23, X8): SubRip and WebVTT read and written, a
// caption track on Video's timeline as text layers, one clip a cue, drawn by
// the engine's text painter: a cue's two lines as two lines, right to left
// for Urdu and Arabic, in the frame's own place for captions.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  parseTime, formatTime, parseSrt, parseVtt, parseCaptions, writeSrt, writeVtt,
  placeCaptions, captionsOf, captionLane, layoutCaptions, isCaption, CAPTION_PLACE,
} from '../packages/studio/src/captions.js';
import { emptyProject, addMedia, laneClips, planPatches, splitAt, clipsOf } from '../packages/studio/src/project.js';
import { buildPlan, paintFrame, layerBounds, fontCss } from '../packages/studio/src/index.js';

const SRT = `1
00:00:00,500 --> 00:00:02,000
Hello, and welcome.

2
00:00:02,250 --> 00:00:04,800
<i>This is</i> the second line
and its second row.

3
00:01:02,003 --> 00:01:03,5
السلام علیکم
`;

const VTT = `WEBVTT - an introduction
Kind: captions

NOTE a note the reader skips

STYLE
::cue { color: yellow }

intro-1
00:00.500 --> 00:02.000 align:center line:90%
<v Ayesha>Hello &amp; welcome.</v>

00:02.250 --> 00:04.800
नमस्ते, <b>दुनिया</b>
<00:03.000>second row
`;

test('times: SubRip and WebVTT times as seconds, and back', () => {
  assert.equal(parseTime('00:00:02,250'), 2.25);
  assert.equal(parseTime('01:02:03.456'), 3723.456);
  assert.equal(parseTime('02:03.5'), 123.5, 'WebVTT without hours, thousandths short');
  assert.ok(Number.isNaN(parseTime('soon')));
  assert.equal(formatTime(3723.456), '01:02:03,456');
  assert.equal(formatTime(2.25, '.'), '00:00:02.250');
  assert.equal(formatTime(-4), '00:00:00,000');
});

test('SubRip: every cue with its times and words, its markup dropped and its two lines kept', () => {
  const cues = parseSrt(SRT);
  assert.equal(cues.length, 3);
  assert.deepEqual(cues[0], { start: 0.5, end: 2, text: 'Hello, and welcome.' });
  assert.equal(cues[1].text, 'This is the second line\nand its second row.');
  assert.equal(cues[2].start, 62.003);
  assert.equal(cues[2].end, 63.5);
  assert.equal(cues[2].text, 'السلام علیکم');
  // Windows line ends and a byte order mark read the same.
  assert.deepEqual(parseSrt('\ufeff' + SRT.replace(/\n/g, '\r\n')), cues);
  // A cue with no words, or ending before it starts, is not a cue.
  assert.equal(parseSrt('1\n00:00:01,000 --> 00:00:02,000\n\n\n2\n00:00:05,000 --> 00:00:03,000\nbackwards\n').length, 0);
});

test('WebVTT: the header, notes and styles skipped, voices, tags, entities and inline times taken out', () => {
  const cues = parseVtt(VTT);
  assert.equal(cues.length, 2);
  assert.deepEqual(cues[0], { start: 0.5, end: 2, text: 'Hello & welcome.' });
  assert.equal(cues[1].text, 'नमस्ते, दुनिया\nsecond row');
  // A header with a cue straight under it.
  assert.equal(parseVtt('WEBVTT\n00:01.000 --> 00:02.000\nNo blank line\n').length, 1);
});

test('either kind is told by its header, then its name', () => {
  assert.equal(parseCaptions(VTT).format, 'vtt');
  assert.equal(parseCaptions(SRT, 'talk.srt').format, 'srt');
  assert.equal(parseCaptions('00:01.000 --> 00:02.000\nhi\n', 'talk.vtt').format, 'vtt');
});

test('written back, each kind reads as the same cues, and WebVTT escapes what it reserves', () => {
  const cues = [{ start: 0.5, end: 2, text: 'A & B <tag>' }, { start: 2.25, end: 4.8, text: 'دو سطریں\nsecond row' }];
  const srt = writeSrt(cues);
  assert.match(srt, /^1\r\n00:00:00,500 --> 00:00:02,000\r\nA & B <tag>\r\n/);
  assert.deepEqual(parseSrt(srt).map((c) => c.text), ['A & B', 'دو سطریں\nsecond row'], 'SubRip has no escapes: a tag-like text reads as a tag');
  const vtt = writeVtt(cues);
  assert.match(vtt, /^WEBVTT\n\n00:00:00\.500 --> 00:00:02\.000\nA &amp; B &lt;tag&gt;\n/);
  assert.deepEqual(parseVtt(vtt), cues);
});

test('on the timeline: a cue a clip, on a lane above every picture, low and centred on the scrim, overlapping cues a lane apart', () => {
  let doc = emptyProject();
  doc = addMedia(doc, { url: 'a.webm', name: 'a', kind: 'video', duration: 10, hasAudio: true }).doc;
  doc = { ...doc, patches: doc.patches.map((p) => (p.type === 'video' ? p : p)) };
  const cues = [...parseSrt(SRT).slice(0, 2), { start: 1.5, end: 3, text: 'over the first' }];
  const { doc: out, ids } = placeCaptions(doc, cues, { at: 1 });
  assert.equal(ids.length, 3);
  const caps = clipsOf(out).filter(isCaption);
  assert.equal(caps.length, 3);
  assert.equal(captionLane(out), 1, 'the first lane above the video');
  assert.deepEqual(caps.find((c) => c.text === 'Hello, and welcome.').timing, { start: 1.5, end: 3 });
  assert.equal(caps.find((c) => c.text === 'over the first').lane, 2, 'an overlapping cue goes a lane up');
  const first = caps[0];
  assert.equal(first.align, 'center');
  assert.equal(first.baseline, 'bottom');
  assert.equal(first.fy, CAPTION_PLACE.landscape.fy);
  assert.equal(first.bg, true);
  assert.equal(laneClips(out, 'v', 1).length, 2);
  // Read back as cues, in time order, the offset kept.
  assert.deepEqual(captionsOf(out).map((c) => c.start), [1.5, 2.5, 3.25]);
  // Imported again, the old captions go.
  const again = placeCaptions(out, [{ start: 0, end: 1, text: 'only' }]).doc;
  assert.equal(clipsOf(again).filter(isCaption).length, 1);
  assert.equal(clipsOf(again).filter(isCaption)[0].lane, 1, 'on the same lane as the ones it replaced');
  // A caption is a clip: S splits it in two, both captions.
  const split = splitAt(out, 2, [first.id]).doc;
  assert.equal(clipsOf(split).filter(isCaption).length, 4);
  // Laid out for a tall frame, higher and bigger.
  const tall = layoutCaptions(out, 'vertical');
  assert.equal(clipsOf(tall).filter(isCaption)[0].fy, CAPTION_PLACE.vertical.fy);
});

/** Enough of a 2D canvas for a plan to be built and painted in Node. */
function canvasStub() {
  const calls = [];
  const ctx = new Proxy({ calls, measureText: (s) => ({ width: String(s).length * 10 }), createLinearGradient: () => ({ addColorStop() {} }) }, {
    get: (o, k) => (k in o ? o[k] : (...args) => { calls.push([k, ...args]); }),
    set: (o, k, v) => { if (k !== 'calls') calls.push([`set:${String(k)}`, v]); o[k] = v; return true; },
  });
  return { width: 0, height: 0, getContext: () => ctx, ctx };
}

test('drawn: a cue of two lines is two lines grown up from its foot, on one pill, and Urdu runs right to left', () => {
  const doc = placeCaptions(emptyProject(), [{ start: 0, end: 3, text: 'first row\nsecond row' }, { start: 3, end: 6, text: 'اردو میں سرخی' }]).doc;
  const canvas = canvasStub();
  const plan = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos: {}, assets: {} });
  const cap = plan.layers.find((l) => l.text === 'first row\nsecond row');
  paintFrame(canvas.ctx, plan, 1);
  const texts = canvas.ctx.calls.filter((k) => k[0] === 'fillText');
  assert.deepEqual(texts.map((k) => k[1]), ['first row', 'second row']);
  const step = Math.round(cap.sizePx * 1.25);
  assert.equal(texts[1][3] - texts[0][3], step, 'a line apart');
  assert.equal(texts[1][3], cap.y, 'the last line on the layer\'s foot');
  assert.equal(canvas.ctx.calls.filter((k) => k[0] === 'roundRect' || k[0] === 'arcTo').length > 0 || canvas.ctx.calls.some((k) => k[0] === 'fill'), true, 'on a pill');
  const b = layerBounds(canvas.ctx, plan, cap);
  assert.ok(b.h > cap.sizePx * 2, 'its box holds both lines');
  canvas.ctx.calls.length = 0;
  paintFrame(canvas.ctx, plan, 4);
  const urdu = plan.layers.find((l) => l.text === 'اردو میں سرخی');
  assert.equal(urdu.rtl, true);
  assert.ok(canvas.ctx.calls.some((k) => k[0] === 'set:direction' && k[1] === 'rtl'), 'the canvas set right to left');
});

test('a text layer in a family of its own keeps the stack after it for the letters that family lacks', () => {
  assert.equal(fontCss(700, 40, 'Georgia').startsWith('700 40px "Georgia", "Segoe UI"'), true);
  assert.equal(fontCss(600, 20, ''), '600 20px "Segoe UI", system-ui, Roboto, "Helvetica Neue", Arial, sans-serif');
  assert.equal(fontCss(600, 20, 'Bad"Name').includes('"BadName"'), true);
  const canvas = canvasStub();
  const plan = buildPlan({ canvas, images: [], options: { aspect: 'landscape', bare: true }, layerPatches: [{ id: 't1', type: 'text', text: 'Hi', fontFamily: 'Georgia', sizeFrac: 0.05, timing: { start: 0, end: 2 } }], videos: {}, assets: {} });
  assert.match(plan.layers.find((l) => l.id === 't1').font, /"Georgia"/);
});
