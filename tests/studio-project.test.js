// The Video app's project model (docs/STUDIO.md D1): the studio document as
// a timeline of clips with lanes and groups: adding media, split, trim,
// ripple, gaps, moving, copying, snapping, links, markers, transitions,
// titles; and what the renderer is handed: a bare plan as long as its
// clips, lanes in paint order, a lane and a group on every layer.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  emptyProject, normaliseProject, addMedia, splitAt, splitAll, splitClips, trimStart, trimEnd,
  removeClips, closeGap, gapAt, moveClip, copyClips, pasteClips, duplicateClips, snapTime,
  unlinkClips, linkClips, addMarker, nextMarker, previousMarker, removeMarker, setTransition,
  clearTransition, addTitle, setRange, clearRange, cutToRange, planPatches, videoKeys, projectMedia, projectDuration, laneClips,
  clipById, partnersOf, laneCount, clipsAt, previousOnLane, TITLE_PRESETS, animationName, TEXT_ANIMATIONS,
} from '../packages/studio/src/project.js';
import { buildPlan, layersEnd, paintFrame } from '../packages/studio/src/index.js';

const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 6, hasAudio: true, width: 320, height: 240 };
const B = { url: 'D:/clips/b.webm', name: 'b.webm', kind: 'video', duration: 4, hasAudio: true, width: 320, height: 240 };
const STILL = { url: 'D:/clips/still.png', name: 'still.png', kind: 'image', width: 640, height: 400 };
const TONE = { url: 'D:/clips/tone.wav', name: 'tone.wav', kind: 'audio', duration: 3 };

const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 0.002, `${msg}: ${a} is not ${b}`);
const timing = (doc, id) => clipById(doc, id).timing;

/** Two clips and a still, end to end. */
function three() {
  let doc = emptyProject();
  const a = addMedia(doc, A); doc = a.doc;
  const b = addMedia(doc, B); doc = b.doc;
  const s = addMedia(doc, STILL); doc = s.doc;
  return { doc, a: a.ids, b: b.ids, s: s.ids };
}

test('a project is a bare studio document; anything stored comes back as one', () => {
  const doc = emptyProject();
  assert.equal(doc.kind, 'video');
  assert.equal(doc.options.bare, true);
  assert.equal(doc.options.aspect, 'landscape');
  assert.deepEqual(doc.patches, []);
  const back = normaliseProject({ patches: [{ id: 'video-1', type: 'video', timing: { start: 0, end: 1 } }], markers: [{ id: 'm1', t: 2 }, { t: 'x' }], options: { bare: false } });
  assert.equal(back.options.bare, true, 'a project is always bare');
  assert.equal(back.markers.length, 1, 'a marker with no time is dropped');
  assert.equal(normaliseProject(null).kind, 'video');
});

test('V1: several files end to end - a video is its picture and its sound in one group, a still lasts five seconds', () => {
  const { doc, a, b, s } = three();
  assert.equal(a.length, 2, 'a clip with sound is two layers');
  const [av, as] = a.map((id) => clipById(doc, id));
  assert.equal(av.type, 'video');
  assert.equal(as.type, 'sound');
  assert.ok(av.group && av.group === as.group, 'picture and sound share a group');
  assert.notEqual(clipById(doc, b[0]).group, av.group, 'each clip its own group');
  assert.deepEqual(timing(doc, b[0]), { start: 6, end: 10 }, 'the second clip follows the first');
  assert.deepEqual(timing(doc, s[0]), { start: 10, end: 15 }, 'the still follows, for five seconds');
  assert.equal(clipById(doc, s[0]).fit, 'cover');
  assert.equal(projectDuration(doc), 15);
  assert.equal(doc.sources[A.url].duration, 6);
  const silent = addMedia(emptyProject(), { ...A, hasAudio: false });
  assert.equal(silent.ids.length, 1, 'a clip without sound has no sound layer');
  const music = addMedia(doc, TONE, { at: 0, lane: 1 });
  assert.equal(clipById(music.doc, music.ids[0]).type, 'sound');
  assert.equal(clipById(music.doc, music.ids[0]).lane, 1, 'a sound file on the lane asked for');
  const still = addMedia({ ...emptyProject(), settings: { stillSeconds: 2 } }, STILL);
  near(clipById(still.doc, still.ids[0]).timing.end, 2, 'the still length is a setting');
});

test('V2: split at the playhead - two halves, the second\'s offset moved on, its sound split with it in a new group', () => {
  const { doc, a } = three();
  const [vid, snd] = a;
  const out = splitAt(doc, 2.5);
  assert.equal(out.ids.length, 2, 'the picture and its sound');
  const first = clipById(out.doc, vid);
  const second = clipById(out.doc, out.ids.find((id) => clipById(out.doc, id).type === 'video'));
  assert.deepEqual(first.timing, { start: 0, end: 2.5 });
  assert.deepEqual(second.timing, { start: 2.5, end: 6 });
  near(second.offset, 2.5, 'the second half starts 2.5 s into the source');
  const secondSound = clipById(out.doc, out.ids.find((id) => clipById(out.doc, id).type === 'sound'));
  near(secondSound.offset, 2.5, 'and so does its sound');
  assert.equal(second.group, secondSound.group, 'the halves pair up');
  assert.notEqual(second.group, first.group, 'in a group of their own');
  assert.equal(clipById(out.doc, snd).timing.end, 2.5);
  assert.equal(projectDuration(out.doc), projectDuration(doc), 'a split changes nothing that plays');
  // Splitting again inside the second half rolls the offset on from there.
  const again = splitClips(out.doc, [second.id], 4);
  const third = clipById(again.doc, again.ids.find((id) => clipById(again.doc, id).type === 'video'));
  near(third.offset, 4, 'offset follows the clock');
  // Not at an edge.
  assert.equal(splitAt(doc, 0).ids.length, 0);
  assert.equal(splitAt(doc, 6).ids.length, 0, 'a cut at a clip boundary makes nothing');
});

test('V2: Split All Tracks cuts every lane at the playhead; plain Split prefers a selected clip, then the main lane', () => {
  let { doc } = three();
  const t = addTitle(doc, 'title', { at: 1, text: 'Hello' });
  doc = t.doc;
  const all = splitAll(doc, 2);
  assert.equal(all.ids.length, 3, 'picture, sound and title');
  const plain = splitAt(doc, 2);
  assert.ok(plain.ids.every((id) => clipById(plain.doc, id).type !== 'text'), 'S cuts the clip, not the title above it');
  const chosen = splitAt(doc, 2, t.ids);
  assert.equal(chosen.ids.length, 1);
  assert.equal(clipById(chosen.doc, chosen.ids[0]).type, 'text', 'a selected title is cut');
});

test('V3: trim start rolls the source offset, trim end stops at the source, the group trims together', () => {
  const { doc, a, b } = three();
  const [vid, snd] = a;
  const s = trimStart(doc, vid, 1.5);
  assert.deepEqual(timing(s, vid), { start: 1.5, end: 6 }, 'without ripple a gap opens');
  near(clipById(s, vid).offset, 1.5, 'offset rolls');
  near(clipById(s, snd).offset, 1.5, 'the sound trims with it');
  assert.deepEqual(timing(s, b[0]), { start: 6, end: 10 }, 'the next clip stays');
  const e = trimEnd(doc, b[0], 8);
  assert.deepEqual(timing(e, b[0]), { start: 6, end: 8 });
  const tooFar = trimEnd(doc, a[0], 9);
  near(timing(tooFar, a[0]).end, 6, 'a clip cannot run past its source');
  const back = trimStart(s, vid, 0);
  near(clipById(back, vid).offset, 0, 'a trim pulled back restores the start');
  const short = trimEnd(doc, a[0], -5);
  near(timing(short, a[0]).end - timing(short, a[0]).start, 0.1, 'never shorter than a tenth of a second');
});

test('V4: ripple trim and ripple delete close up the lane, the groups following; Close Gap closes a gap left behind', () => {
  const { doc, a, b, s } = three();
  const rs = trimStart(doc, a[0], 2, { ripple: true });
  assert.deepEqual(timing(rs, a[0]), { start: 0, end: 4 }, 'the clip keeps its place');
  near(clipById(rs, a[0]).offset, 2, 'and starts later in its source');
  assert.deepEqual(timing(rs, b[0]), { start: 4, end: 8 }, 'what follows moves up');
  assert.deepEqual(timing(rs, b[1]), { start: 4, end: 8 }, 'with its sound');
  assert.deepEqual(timing(rs, s[0]), { start: 8, end: 13 });
  const re = trimEnd(doc, a[0], 5, { ripple: true });
  assert.deepEqual(timing(re, b[0]), { start: 5, end: 9 });

  const gone = removeClips(doc, [b[0]], { ripple: true });
  assert.equal(clipById(gone, b[0]), null);
  assert.equal(clipById(gone, b[1]), null, 'its sound goes with it');
  assert.deepEqual(timing(gone, s[0]), { start: 6, end: 11 }, 'the still moves up');

  const left = removeClips(doc, [b[0]]);
  assert.deepEqual(timing(left, s[0]), { start: 10, end: 15 }, 'without ripple a gap stays');
  assert.deepEqual(gapAt(left, 'v', 0, 7), { start: 6, end: 10 });
  const closed = closeGap(left, 'v', 0, 7);
  assert.deepEqual(timing(closed, s[0]), { start: 6, end: 11 });
  assert.equal(gapAt(closed, 'v', 0, 7), null);
  assert.equal(closeGap(doc, 'v', 0, 3), doc, 'no gap, nothing to close');
});

test('V5: move to a free place, between lanes, with the group; copy, paste and duplicate make fresh clips and groups', () => {
  const { doc, a, b } = three();
  const onto = moveClip(doc, b[0], { start: 1 });
  const bt = timing(onto, b[0]);
  assert.ok(bt.end <= 0.001 + 0 || bt.start >= 5.999, `a clip never covers another on its lane (${bt.start})`);
  near(timing(onto, b[1]).start, bt.start, 'its sound moves with it');
  const up = moveClip(doc, b[0], { start: 2, lane: 1 });
  assert.equal(clipById(up, b[0]).lane, 1);
  near(timing(up, b[0]).start, 2, 'a lane of its own: the place asked for');
  assert.equal(clipById(up, b[1]).lane, 0, 'its sound keeps its own lane');
  assert.equal(laneCount(up).v, 2);

  const copied = copyClips(doc, [a[0]]);
  assert.equal(copied.clips.length, 2, 'a clip is copied with its sound');
  const pasted = pasteClips(doc, copied, 20);
  assert.equal(pasted.ids.length, 2);
  const pv = clipById(pasted.doc, pasted.ids[0]);
  assert.notEqual(pv.id, a[0]);
  assert.notEqual(pv.group, clipById(doc, a[0]).group, 'a pasted clip has a group of its own');
  assert.equal(clipById(pasted.doc, pasted.ids[1]).group, pv.group);
  near(pv.timing.start, 20, 'at the playhead');
  const over = pasteClips(doc, copied, 1);
  assert.ok(clipById(over.doc, over.ids[0]).timing.start >= 15 - 0.001, 'pasted over clips, it goes to the first free place');
  const dup = duplicateClips(doc, [a[0]]);
  assert.ok(clipById(dup.doc, dup.ids[0]).timing.start >= 6 - 0.001);
});

test('V5: snapping to clip edges, the playhead and markers, within the threshold only', () => {
  let { doc, b } = three();
  doc = addMarker(doc, 7.3, 'Beat');
  assert.deepEqual(snapTime(doc, 5.95, { threshold: 0.1 }), { t: 6, to: 6 });
  assert.deepEqual(snapTime(doc, 5.8, { threshold: 0.1 }), { t: 5.8, to: null });
  assert.equal(snapTime(doc, 7.26, { threshold: 0.1 }).to, 7.3, 'to a marker');
  assert.equal(snapTime(doc, 3.04, { threshold: 0.1, playhead: 3 }).to, 3, 'to the playhead');
  assert.equal(snapTime(doc, 6.02, { threshold: 0.1, exclude: [b[0], b[1]] }).to, 6, 'the end of the clip before still snaps');
  assert.equal(snapTime(doc, 0.9, { threshold: 0.2, beats: [1] }).to, 1, 'to a beat');
});

test('V6: detach audio unlinks a picture from its sound; link groups them again', () => {
  const { doc, a } = three();
  const free = unlinkClips(doc, [a[0]]);
  assert.equal(clipById(free, a[0]).group, undefined);
  assert.equal(clipById(free, a[1]).group, undefined);
  assert.deepEqual(partnersOf(free, a[0]), []);
  const moved = moveClip(free, a[1], { start: 20 });
  near(timing(moved, a[0]).start, 0, 'a detached sound moves on its own');
  const back = linkClips(free, [a[0], a[1]]);
  assert.equal(partnersOf(back, a[0])[0].id, a[1]);
});

test('V7: markers in order, next and previous from the playhead', () => {
  let doc = emptyProject();
  doc = addMarker(doc, 5, 'Two');
  doc = addMarker(doc, 2, 'One');
  assert.deepEqual(doc.markers.map((m) => m.name), ['One', 'Two']);
  assert.equal(nextMarker(doc, 0).name, 'One');
  assert.equal(nextMarker(doc, 2).name, 'Two', 'not the one under the playhead');
  assert.equal(nextMarker(doc, 5), null);
  assert.equal(previousMarker(doc, 5).name, 'One');
  assert.equal(removeMarker(doc, doc.markers[0].id).markers.length, 1);
});

test('V10: In and Out points; the project cut to them plays what it played there, from zero', () => {
  const { doc, a, b } = three();
  const r = setRange(doc, { in: 4 });
  assert.deepEqual(r.range, { in: 4, out: 15 }, 'the Out point defaults to the end');
  const r2 = setRange(r, { out: 8 });
  assert.deepEqual(r2.range, { in: 4, out: 8 });
  assert.deepEqual(setRange(r2, { in: 9 }).range, { in: 9, out: 15 }, 'an In after the Out moves the Out');
  assert.equal(clearRange(r2).range, null);
  const cut = cutToRange(addMarker(doc, 5, 'In range'), 4, 8);
  assert.equal(projectDuration(cut), 4);
  const first = cut.patches.find((p) => p.type === 'video' && p.url === A.url);
  assert.deepEqual(first.timing, { start: 0, end: 2 });
  near(first.offset, 4, 'the first clip starts where the range began in it');
  const pics = cut.patches.filter((p) => p.type === 'video');
  assert.equal(pics.length, 2);
  const second = pics.find((p) => p.url === B.url);
  assert.deepEqual(second.timing, { start: 2, end: 4 });
  near(second.offset, 0, 'the next clip from its own start');
  assert.equal(cut.patches.filter((p) => p.type === 'image').length, 0, 'what lies outside is gone');
  assert.deepEqual(cut.markers.map((m) => m.t), [1]);
  assert.equal(clipById(cut, b[0]).timing.end, 4);
});

test('V15: a crossfade overlaps the clips by its length, the later clips moving up, the sound crossfading too', () => {
  const { doc, a, b, s } = three();
  const x = setTransition(doc, b[0], { kind: 'fade', seconds: 1 });
  assert.deepEqual(timing(x, b[0]), { start: 5, end: 9 }, 'the incoming clip starts a second early');
  assert.deepEqual(timing(x, b[1]), { start: 5, end: 9 }, 'with its sound');
  assert.deepEqual(timing(x, s[0]), { start: 9, end: 14 }, 'and what follows moves up');
  assert.deepEqual(clipById(x, b[0]).enter, { kind: 'fade', seconds: 1 });
  assert.deepEqual(clipById(x, b[0]).transition, { kind: 'fade', seconds: 1 });
  assert.equal(clipById(x, b[1]).enter.seconds, 1, 'the sound fades in');
  assert.equal(clipById(x, a[1]).exit.seconds, 1, 'and the sound before fades out');
  assert.equal(previousOnLane(x, b[0]).id, a[0], 'the clip it fades from is still the one before');
  assert.equal(projectDuration(x), 14);
  // Replacing it does not move things twice; taking it off puts them back.
  const y = setTransition(x, b[0], { kind: 'wipe', seconds: 1 });
  assert.deepEqual(timing(y, b[0]), { start: 5, end: 9 });
  assert.equal(clipById(y, b[0]).enter.kind, 'wipe');
  const off = clearTransition(y, b[0]);
  assert.deepEqual(timing(off, b[0]), { start: 6, end: 10 });
  assert.equal(clipById(off, b[0]).transition, undefined);
  assert.equal(clipById(off, a[1]).exit.seconds, 0);
  // A dip to black does not overlap; a push moves the clip before out.
  const dip = setTransition(doc, b[0], { kind: 'dip', seconds: 1 });
  assert.deepEqual(timing(dip, b[0]), { start: 6, end: 10 });
  assert.deepEqual(clipById(dip, a[0]).exit, { kind: 'fade', seconds: 0.5 });
  const push = setTransition(doc, b[0], { kind: 'push', seconds: 1 });
  assert.equal(clipById(push, a[0]).exit.kind, 'push');
  // Held to half of the shorter clip; nothing across a gap or with nothing before.
  const long = setTransition(doc, b[0], { kind: 'fade', seconds: 9 });
  near(clipById(long, b[0]).transition.seconds, 2, 'half of the four-second clip');
  assert.equal(setTransition(doc, a[0], { kind: 'fade' }), doc, 'the first clip has nothing to fade from');
  const gapped = removeClips(doc, [b[0]]);
  assert.equal(clipById(setTransition(gapped, s[0], { kind: 'fade' }), s[0]).transition, undefined, 'not across a gap');
});

test('V21, V22: titles from the gallery on the lanes above the clips, with their animations', () => {
  const { doc } = three();
  const t = addTitle(doc, 'title', { at: 2, text: 'Hello' });
  const title = clipById(t.doc, t.ids[0]);
  assert.equal(title.type, 'text');
  assert.equal(title.lane, 1, 'above the clips');
  assert.deepEqual(title.timing, { start: 2, end: 6 });
  assert.equal(animationName(title.enter), 'fade');
  const lower = addTitle(t.doc, 'lowerThird', { at: 3, text: 'Name' });
  assert.equal(lower.ids.length, 2, 'a band and its words');
  const [band, words] = lower.ids.map((id) => clipById(lower.doc, id));
  assert.equal(band.type, 'shape');
  assert.equal(band.group, words.group, 'grouped');
  assert.equal(band.lane, 2, 'the first lane free for its whole length');
  assert.equal(words.lane, 3, 'its words above its band');
  assert.equal(animationName(words.enter), 'rise');
  assert.equal(animationName(TEXT_ANIMATIONS.type), 'type');
  assert.equal(animationName(TEXT_ANIMATIONS.pop), 'pop');
  for (const name of Object.keys(TITLE_PRESETS)) assert.ok(addTitle(doc, name, { text: 'x' }).ids.length >= 1, name);
});

test('lanes: the renderer gets pictures lane by lane, hidden lanes off, muted sound off, a lane\'s volume in', () => {
  let { doc, a, b } = three();
  doc = addTitle(doc, 'title', { at: 0, text: 'Over' }).doc;
  doc = setTransition(doc, b[0], { kind: 'fade', seconds: 1 });
  const patches = planPatches(doc);
  const pics = patches.filter((p) => p.type !== 'sound');
  assert.deepEqual(pics.map((p) => p.lane), [...pics.map((p) => p.lane)].sort((x, y) => x - y), 'lane order is paint order');
  const vids = pics.filter((p) => p.lane === 0);
  assert.ok(vids.findIndex((p) => p.id === b[0]) > vids.findIndex((p) => p.id === a[0]), 'the incoming clip paints over the one it fades from');
  const hidden = planPatches({ ...doc, lanes: { v: { 1: { hidden: true } }, a: { 0: { muted: true } } } });
  assert.equal(hidden.find((p) => p.type === 'text').visible, false);
  assert.ok(hidden.filter((p) => p.type === 'sound').every((p) => p.visible === false));
  const quiet = planPatches({ ...doc, lanes: { v: {}, a: { 0: { volume: 0.5 } } } });
  assert.equal(quiet.find((p) => p.type === 'sound').volume, 0.5);
  const blank = planPatches(addTitle(emptyProject(), 'title', { text: '' }).doc);
  assert.equal(blank[0].visible, false, 'a title with no words draws nothing');
});

test('two overlapping clips of one source get a decoder each; one after another share it', () => {
  const { doc } = three();
  const split = splitAt(doc, 3);
  const second = split.ids.find((id) => clipById(split.doc, id).type === 'video');
  assert.equal(new Set(videoKeys(split.doc).values()).size, 2, 'two sources, back to back: one key each');
  const faded = setTransition(split.doc, second, { kind: 'fade', seconds: 1 });
  const keys = videoKeys(faded);
  assert.equal(keys.get(second), `${A.url}#1`, 'the overlapping half has a decoder of its own');
  const media = projectMedia(faded);
  assert.deepEqual(media.videos.map((v) => v.key).sort(), [A.url, `${A.url}#1`, B.url].sort());
  assert.ok(media.videos.every((v) => v.url === A.url || v.url === B.url));
  assert.deepEqual(media.images, [STILL.url]);
  assert.deepEqual(media.sounds.sort(), [A.url, B.url].sort());
});

/** Enough of a 2D canvas for a plan to be built and painted in Node. */
function canvasStub() {
  const calls = [];
  const ctx = new Proxy({ calls, measureText: (s) => ({ width: String(s).length * 10 }), createLinearGradient: () => ({ addColorStop() {} }) }, {
    get: (o, k) => (k in o ? o[k] : (...args) => { calls.push([k, ...args]); }),
    set: (o, k, v) => { o[k] = v; return true; },
  });
  return { width: 0, height: 0, getContext: () => ctx, ctx };
}

test('the renderer: a bare plan runs as long as its clips, compiles nothing of its own, and every layer carries its lane and group', () => {
  let { doc, a, b } = three();
  doc = addTitle(doc, 'title', { at: 0, text: 'Over' }).doc;
  const canvas = canvasStub();
  const el = { readyState: 4, paused: true, currentTime: 0 };
  const entry = { el, img: el, width: 320, height: 240, duration: 6 };
  const plan = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches(doc), videos: { [A.url]: entry, [B.url]: entry }, assets: {} });
  assert.equal(plan.duration, 15);
  assert.equal(layersEnd([]), 0.5, 'an empty project still has a clock');
  assert.ok(!plan.layers.some((l) => ['caption', 'edges', 'gradient', 'progress'].includes(l.type)), 'no slideshow chrome');
  const v = plan.layers.find((l) => l.id === a[0]);
  assert.equal(v.lane, 0);
  assert.equal(v.group, clipById(doc, a[1]).group);
  assert.equal(v.fit, 'cover');
  assert.equal(plan.layers.find((l) => l.type === 'text').lane, 1);
  // Painting at a time inside the first clip draws its element, and the
  // picture's turn and flip happen about the frame's centre.
  paintFrame(canvas.ctx, plan, 1);
  assert.ok(canvas.ctx.calls.some((c) => c[0] === 'drawImage' && c[1] === el));
  const turned = buildPlan({ canvas, images: [], options: doc.options, layerPatches: planPatches({ ...doc, patches: doc.patches.map((p) => (p.id === b[0] ? { ...p, turn: 90, flipX: true, fit: 'contain' } : p)) }), videos: { [A.url]: entry, [B.url]: entry }, assets: {} });
  const tb = turned.layers.find((l) => l.id === b[0]);
  assert.equal(tb.turn, 90);
  assert.equal(tb.flipX, true);
  canvas.ctx.calls.length = 0;
  paintFrame(canvas.ctx, turned, 7);
  assert.ok(canvas.ctx.calls.some((c) => c[0] === 'rotate'), 'turned');
  assert.ok(canvas.ctx.calls.some((c) => c[0] === 'scale' && c[1] === -1), 'flipped');
  // clipsAt reads the timeline at an instant, topmost picture first.
  assert.equal(clipsAt(doc, 1)[0].type, 'text');
  assert.equal(laneClips(doc, 'v', 0).length, 3);
});
