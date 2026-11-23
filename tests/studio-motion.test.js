// Video's transitions, animations and movement over time (D-2026-11-22-14) in
// the project model: the transitions between two clips and what each asks of
// the renderer, an animation in and out for every kind of layer, and the
// points a clip moves between, kept as the renderer's own keys on the layer.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  emptyProject, addMedia, addTitle, addGraphic, splitClips, clipById, lengthOf, planPatches,
  TRANSITIONS, setTransition, clearTransition, transitionAfter,
  ANIMATIONS, TEXT_ANIMATIONS, animationName, animationsFor, animationEnvelope, animationMost, animationLasting,
  motionProps, motionTimes, isMoving, motionPointAt, motionValue, clipAtTime, changeAtTime, patchAtTime,
  addMotionPoint, removeMotionPoint, moveMotionPoint, motionEase, setMotionEase, clearMotion, MOTION_EASES,
} from '../packages/studio/src/project.js';

const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 6, hasAudio: true, width: 320, height: 240 };
const B = { url: 'D:/clips/b.webm', name: 'b.webm', kind: 'video', duration: 4, hasAudio: true, width: 320, height: 240 };
const STILL = { url: 'D:/clips/still.png', name: 'still.png', kind: 'image', width: 640, height: 400 };
const TONE = { url: 'D:/clips/tone.wav', name: 'tone.wav', kind: 'audio', duration: 3 };

const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 0.002, `${msg}: ${a} is not ${b}`);

function two() {
  let doc = emptyProject();
  const a = addMedia(doc, A); doc = a.doc;
  const b = addMedia(doc, B); doc = b.doc;
  return { doc, a: a.ids, b: b.ids };
}

/** A project with one shape on a lane of its own, five seconds long, and its id. */
function withShape(patch = {}) {
  const r = addGraphic(emptyProject(), { type: 'shape', geometry: 'rect', fx: 0.3, fy: 0.4, fw: 0.2, fh: 0.1, ...patch }, { at: 2, seconds: 5 });
  return { doc: r.doc, id: r.ids[0] };
}

// ── transitions ──────────────────────────────────────────────────────────────

test('the six transitions there were are still the first six, and every one is offered once', () => {
  assert.deepEqual(TRANSITIONS.slice(0, 6), ['fade', 'dip', 'wipe', 'slide', 'push', 'zoom']);
  assert.equal(new Set(TRANSITIONS).size, TRANSITIONS.length);
  for (const kind of ['wipe-left', 'wipe-down', 'wipe-up', 'slide-right', 'slide-up', 'slide-down', 'push-right', 'push-up', 'push-down', 'iris', 'blur']) assert.ok(TRANSITIONS.includes(kind), kind);
});

test('every transition but the dip overlaps the two clips by its length, and taking it off opens them again', () => {
  for (const kind of TRANSITIONS) {
    const { doc, a, b } = two();
    const x = setTransition(doc, b[0], { kind, seconds: 1 });
    const incoming = clipById(x, b[0]);
    assert.deepEqual(incoming.transition, { kind, seconds: 1 }, kind);
    near(incoming.timing.start, kind === 'dip' ? 6 : 5, `${kind}: where the incoming clip starts`);
    assert.ok(incoming.enter.seconds > 0 && incoming.enter.kind !== 'none', `${kind}: the incoming clip has a way in`);
    // The sound of each crossfades with its picture.
    assert.equal(clipById(x, b[1]).enter.kind, 'fade', `${kind}: the incoming sound fades in`);
    assert.equal(clipById(x, a[1]).exit.kind, 'fade', `${kind}: the outgoing sound fades out`);
    const back = clearTransition(x, b[0]);
    assert.deepEqual(clipById(back, b[0]).timing, { start: 6, end: 10 }, `${kind}: back where it was`);
    assert.equal(clipById(back, b[0]).enter.kind, 'none', kind);
    assert.equal(clipById(back, a[0]).exit.kind, 'none', kind);
    assert.equal(clipById(back, b[0]).transition, undefined, kind);
  }
});

test('each new transition asks the renderer for the way it is named: the incoming clip in, and for a push the outgoing one out the same way', () => {
  const kinds = (kind) => {
    const { doc, a, b } = two();
    const x = setTransition(doc, b[0], { kind, seconds: 1 });
    return [clipById(x, b[0]).enter.kind, clipById(x, a[0]).exit.kind];
  };
  assert.deepEqual(kinds('wipe-left'), ['wipe-left', 'none']);
  assert.deepEqual(kinds('wipe-down'), ['wipe-down', 'none']);
  assert.deepEqual(kinds('wipe-up'), ['wipe-up', 'none']);
  assert.deepEqual(kinds('slide-right'), ['slide-right', 'none']);
  assert.deepEqual(kinds('slide-up'), ['slide-up', 'none']);
  assert.deepEqual(kinds('slide-down'), ['slide-down', 'none']);
  assert.deepEqual(kinds('push'), ['push', 'push'], 'the push there was is as it was');
  assert.deepEqual(kinds('push-right'), ['slide-right', 'slide-right']);
  assert.deepEqual(kinds('push-up'), ['slide-up', 'slide-up']);
  assert.deepEqual(kinds('push-down'), ['slide-down', 'slide-down']);
  assert.deepEqual(kinds('iris'), ['iris', 'none']);
  assert.deepEqual(kinds('blur'), ['blur-through', 'blur-through']);
  assert.deepEqual(kinds('dip'), ['fade', 'fade'], 'the dip there was is as it was');
});

test('a transition the list does not have changes nothing, and two sounds of their own still take the crossfade alone', () => {
  const { doc, b } = two();
  assert.equal(setTransition(doc, b[0], { kind: 'tumble', seconds: 1 }), doc);
  let s = emptyProject();
  const one = addMedia(s, TONE); s = one.doc;
  const other = addMedia(s, TONE); s = other.doc;
  assert.equal(clipById(setTransition(s, other.ids[0], { kind: 'iris', seconds: 1 }), other.ids[0]).transition, undefined);
  assert.equal(clipById(setTransition(s, other.ids[0], { kind: 'fade', seconds: 1 }), other.ids[0]).transition.kind, 'fade');
});

test('the clip after a clip with a transition from it is found, for that transition is the first clip\'s way out', () => {
  const { doc, a, b } = two();
  assert.equal(transitionAfter(doc, a[0]), null);
  const x = setTransition(doc, b[0], { kind: 'push-up', seconds: 1 });
  assert.equal(transitionAfter(x, a[0]).id, b[0]);
  assert.equal(transitionAfter(x, b[0]), null);
});

// ── animations ───────────────────────────────────────────────────────────────

test('every animation is named back from its envelope, and the titles\' list is the same list', () => {
  assert.equal(TEXT_ANIMATIONS, ANIMATIONS);
  for (const name of Object.keys(ANIMATIONS)) assert.equal(animationName(ANIMATIONS[name]), name, name);
  assert.equal(animationName(null), 'none');
  assert.equal(animationName({ kind: 'slide-up', seconds: 0.6, dist: 60 }), 'rise', 'a title made before reads as it did');
  assert.equal(animationName({ kind: 'wipe-up', seconds: 1 }), 'wipe');
  assert.equal(animationName({ kind: 'push', seconds: 1 }), 'left');
});

test('every kind of layer but a sound, a pointer and the fade at the two ends can be animated in and out; typing on is a title\'s way in alone', () => {
  let doc = emptyProject();
  const v = addMedia(doc, A); doc = v.doc;
  const s = addMedia(doc, STILL); doc = s.doc;
  const title = addTitle(doc, 'title', { at: 0, text: 'Hello' }); doc = title.doc;
  const video = clipById(doc, v.ids[0]);
  const sound = clipById(doc, v.ids[1]);
  const words = title.ids.map((id) => clipById(doc, id)).find((p) => p.type === 'text');
  assert.deepEqual(animationsFor(sound, 'in'), []);
  assert.deepEqual(animationsFor({ type: 'edges', timing: { start: 0, end: 5 } }), []);
  assert.deepEqual(animationsFor({ type: 'pointer', timing: { start: 0, end: 5 } }), []);
  for (const clip of [video, clipById(doc, s.ids[0]), { type: 'shape', timing: { start: 0, end: 5 } }, { type: 'chart', timing: { start: 0, end: 5 } }, { type: 'adjust', timing: { start: 0, end: 5 } }]) {
    const ins = animationsFor(clip, 'in');
    assert.equal(ins[0], 'none', clip.type);
    assert.ok(!ins.includes('type'), `${clip.type}: typing on is not offered`);
    for (const name of ['fade', 'rise', 'drop', 'pop', 'grow', 'spin', 'left', 'right', 'up', 'down', 'wipe', 'iris', 'blur']) assert.ok(ins.includes(name), `${clip.type}: ${name}`);
    assert.deepEqual(animationsFor(clip, 'out'), ins, clip.type);
  }
  assert.ok(animationsFor(words, 'in').includes('type'));
  assert.ok(!animationsFor(words, 'out').includes('type'));
});

test('an animation keeps its own length unless one is asked for, held to half the clip; typing on and none have none', () => {
  const clip = { type: 'shape', timing: { start: 0, end: 1 } };
  assert.deepEqual(animationEnvelope('grow'), { kind: 'grow', seconds: 0.5 });
  assert.deepEqual(animationEnvelope('rise', 1.2), { kind: 'slide-up', seconds: 1.2, dist: 60 });
  assert.equal(animationEnvelope('left', 3, clip).seconds, 0.5);
  assert.equal(animationEnvelope('fade', 0, clip).seconds, 0.1);
  assert.deepEqual(animationEnvelope('type', 2), { kind: 'type-on', seconds: 0 });
  assert.deepEqual(animationEnvelope('none', 2), { kind: 'none', seconds: 0 });
  assert.deepEqual(animationEnvelope('whirl'), { kind: 'none', seconds: 0 });
  assert.notEqual(animationEnvelope('fade'), ANIMATIONS.fade, 'the list itself is never handed out to be changed');
});

// ── movement over time ───────────────────────────────────────────────────────

test('what a point holds follows the kind of layer', () => {
  let doc = emptyProject();
  const v = addMedia(doc, A); doc = v.doc;
  const s = addMedia(doc, STILL); doc = s.doc;
  const title = addTitle(doc, 'title', { at: 0, text: 'Hello' }); doc = title.doc;
  const video = clipById(doc, v.ids[0]);
  const words = title.ids.map((id) => clipById(doc, id)).find((p) => p.type === 'text');
  assert.deepEqual(motionProps(video), ['zoom', 'panX', 'panY', 'opacity'], 'a clip that fills the frame is pushed in and panned');
  assert.deepEqual(motionProps({ ...video, autoZoom: [{ start: 0, end: 1 }] }), ['opacity'], 'the zooms on clicks keep the zoom');
  assert.deepEqual(motionProps({ ...video, fx: 0.6, fy: 0.6, fw: 0.3, fh: 0.3 }), ['fx', 'fy', 'fw', 'fh', 'rot', 'opacity'], 'a picture in picture moves in its box');
  assert.deepEqual(motionProps(clipById(doc, s.ids[0])), ['zoom', 'panX', 'panY', 'opacity']);
  assert.deepEqual(motionProps(words), ['fx', 'fy', 'sizeFrac', 'rot', 'opacity']);
  assert.deepEqual(motionProps({ type: 'qr', timing: { start: 0, end: 5 }, fx: 0.5, fy: 0.5, fw: 0.2 }), ['fx', 'fy', 'opacity'], 'a QR code keeps its size and stands upright');
  assert.deepEqual(motionProps({ type: 'adjust', timing: { start: 0, end: 5 } }), ['opacity']);
  assert.deepEqual(motionProps(clipById(doc, v.ids[1])), [], 'a sound does not move');
  assert.deepEqual(motionProps({ ...words, caption: true }), [], 'nor a caption');
});

test('a point is added at the playhead holding the clip as it looks there, once', () => {
  const { doc, id } = withShape();
  assert.equal(isMoving(clipById(doc, id)), false);
  const x = addMotionPoint(doc, id, 3);
  const clip = clipById(x, id);
  assert.deepEqual(motionTimes(clip), [1], 'counted from the clip\'s start');
  assert.deepEqual(clip.keys, { fx: [{ t: 1, v: 0.3 }], fy: [{ t: 1, v: 0.4 }], fw: [{ t: 1, v: 0.2 }], fh: [{ t: 1, v: 0.1 }], rot: [{ t: 1, v: 0 }], opacity: [{ t: 1, v: 1 }] });
  assert.equal(addMotionPoint(x, id, 3.02), x, 'a point is not added on a point');
  assert.equal(addMotionPoint(doc, id, 1), doc, 'nor outside the clip');
  assert.equal(addMotionPoint(doc, id, 7.5), doc);
  assert.equal(motionPointAt(clip, 1.03), 1);
  assert.equal(motionPointAt(clip, 1.2), null);
});

test('a change made at another time sets the next point, and the clip is carried between the two', () => {
  const { doc, id } = withShape();
  let x = addMotionPoint(doc, id, 2);
  x = patchAtTime(x, id, 6, { fx: 0.7, fy: 0.4 });
  const clip = clipById(x, id);
  assert.deepEqual(motionTimes(clip), [0, 4]);
  assert.equal(clip.fx, 0.3, 'the clip\'s own place is left as it was');
  near(motionValue(clip, 'fx', 0), 0.3, 'at the first point');
  near(motionValue(clip, 'fx', 2), 0.5, 'half way');
  near(motionValue(clip, 'fx', 4), 0.7, 'at the second point');
  near(motionValue(clip, 'fx', 9), 0.7, 'held after the last point');
  near(motionValue(clip, 'fw', 2), 0.2, 'what did not change is the same at both');
  assert.equal(clip.keys.fw.length, 2, 'every point holds everything the clip moves by');
  // A change on a point changes that point, and makes no other.
  x = patchAtTime(x, id, 6.01, { rot: 30 });
  assert.deepEqual(motionTimes(clipById(x, id)), [0, 4]);
  assert.equal(clipById(x, id).keys.rot[1].v, 30);
  // What does not move is the clip's own, as before.
  x = patchAtTime(x, id, 4, { color: '#ff0000' });
  assert.equal(clipById(x, id).color, '#ff0000');
  assert.deepEqual(motionTimes(clipById(x, id)), [0, 4]);
  // A clip with no points takes a change as it is.
  const still = patchAtTime(doc, id, 4, { fx: 0.9 });
  assert.equal(clipById(still, id).fx, 0.9);
  assert.equal(clipById(still, id).keys, undefined);
});

test('the numbers shown for a moving clip are the ones at the playhead', () => {
  const { doc, id } = withShape();
  let x = addMotionPoint(doc, id, 2);
  x = patchAtTime(x, id, 6, { fx: 0.7, opacity: 0 });
  const at = clipAtTime(clipById(x, id), 2);
  near(at.fx, 0.5, 'across');
  near(at.opacity, 0.5, 'opacity');
  assert.equal(at.fy, 0.4);
  const still = clipById(doc, id);
  assert.equal(clipAtTime(still, 2), still, 'a clip with no points is itself');
});

test('a point is moved along the clip, held inside it, and replaces a point it lands on', () => {
  const { doc, id } = withShape();
  let x = addMotionPoint(doc, id, 2);
  x = patchAtTime(x, id, 4, { fx: 0.5 });
  x = patchAtTime(x, id, 6, { fx: 0.7 });
  assert.deepEqual(motionTimes(clipById(x, id)), [0, 2, 4]);
  const moved = moveMotionPoint(x, id, 2, 3);
  assert.deepEqual(motionTimes(clipById(moved, id)), [0, 3, 4]);
  near(motionValue(clipById(moved, id), 'fx', 3), 0.5, 'it carries what it held');
  const out = moveMotionPoint(x, id, 4, 99);
  assert.deepEqual(motionTimes(clipById(out, id)), [0, 2, 5], 'held at the clip\'s end');
  const onto = moveMotionPoint(x, id, 2, 4.01);
  assert.deepEqual(motionTimes(clipById(onto, id)), [0, 4.01]);
  near(motionValue(clipById(onto, id), 'fx', 4.01), 0.5, 'the point moved there is the one kept');
  assert.equal(moveMotionPoint(x, id, 1, 3), x, 'there is no point at one second');
});

test('how a move arrives is kept on the point it arrives at, for everything the clip moves by', () => {
  const { doc, id } = withShape();
  let x = addMotionPoint(doc, id, 2);
  x = patchAtTime(x, id, 6, { fx: 0.7 });
  assert.equal(motionEase(clipById(x, id), 4), '');
  x = setMotionEase(x, id, 4, 'in-out');
  const clip = clipById(x, id);
  assert.equal(motionEase(clip, 4), 'in-out');
  for (const prop of motionProps(clip)) assert.equal(clip.keys[prop][1].ease, 'in-out', prop);
  near(motionValue(clip, 'fx', 1), 0.3 + 0.4 * (0.25 * 0.25 * (3 - 0.5)), 'slow at first');
  near(motionValue(clip, 'fx', 2), 0.5, 'half way at the middle');
  x = setMotionEase(x, id, 4, '');
  assert.equal(clipById(x, id).keys.fx[1].ease, undefined);
  assert.equal(setMotionEase(x, id, 4, 'bounce'), x);
  assert.deepEqual(MOTION_EASES, ['', 'in', 'out', 'in-out']);
});

test('taking the last point off leaves the clip as that point had it; taking all movement off leaves it as it looks at the playhead', () => {
  const { doc, id } = withShape();
  let x = addMotionPoint(doc, id, 2);
  x = patchAtTime(x, id, 6, { fx: 0.7 });
  const one = removeMotionPoint(x, id, 0);
  assert.deepEqual(motionTimes(clipById(one, id)), [4]);
  const none = removeMotionPoint(one, id, 4);
  assert.equal(clipById(none, id).keys, undefined);
  assert.equal(clipById(none, id).fx, 0.7);
  assert.equal(removeMotionPoint(x, id, 1), x, 'there is no point there');
  const cleared = clearMotion(x, id, 4);
  assert.equal(clipById(cleared, id).keys, undefined);
  near(clipById(cleared, id).fx, 0.5, 'as it looked at four seconds');
  assert.equal(clearMotion(doc, id, 4), doc);
});

test('a clip that fills the frame is pushed in between two points, and the zooms on clicks keep their own keys', () => {
  let doc = emptyProject();
  const v = addMedia(doc, { ...A, hasAudio: false }); doc = v.doc;
  const id = v.ids[0];
  let x = addMotionPoint(doc, id, 0);
  x = patchAtTime(x, id, 6, { zoom: 1.5, panX: 0.2 });
  const clip = clipById(x, id);
  assert.deepEqual(clip.keys.zoom, [{ t: 0, v: 1 }, { t: 6, v: 1.5 }]);
  assert.deepEqual(clip.keys.panX, [{ t: 0, v: 0.5 }, { t: 6, v: 0.2 }]);
  assert.equal(clip.zoom, undefined, 'the clip itself is not zoomed');
  // A clip with zooms on clicks: its zoom keys are theirs, and movement leaves them alone.
  const zoomed = { ...doc, patches: doc.patches.map((p) => (p.id === id ? { ...p, autoZoom: [{ start: 1, end: 2 }], keys: { zoom: [{ t: 1, v: 1 }, { t: 1.5, v: 2 }] } } : p)) };
  assert.equal(isMoving(clipById(zoomed, id)), false);
  let y = addMotionPoint(zoomed, id, 0);
  y = patchAtTime(y, id, 3, { opacity: 0.2 });
  assert.deepEqual(clipById(y, id).keys.zoom, [{ t: 1, v: 1 }, { t: 1.5, v: 2 }]);
  const cleared = clearMotion(y, id, 0);
  assert.deepEqual(clipById(cleared, id).keys, { zoom: [{ t: 1, v: 1 }, { t: 1.5, v: 2 }] });
});

test('a still made a picture of its own size, or made to fill the frame again, drops the points of what it no longer moves by', () => {
  let doc = emptyProject();
  const s = addMedia(doc, STILL); doc = s.doc;
  const id = s.ids[0];
  let x = addMotionPoint(doc, id, 0);
  x = patchAtTime(x, id, 4, { zoom: 2 });
  // Made a picture of its own size: the push in goes, its opacity points stay.
  x = patchAtTime(x, id, 4, { fit: undefined, fx: 0.3, fy: 0.25, fw: 0.4, rot: 0 });
  let clip = clipById(x, id);
  assert.equal(clip.keys.zoom, undefined);
  assert.equal(clip.fw, 0.4);
  assert.deepEqual(motionTimes(clip), [0, 4]);
  assert.ok(clip.keys.fx && clip.keys.fw && clip.keys.opacity, 'the points now hold its place and size');
  // Made to fill the frame again: its place and size go from the points.
  x = patchAtTime(x, id, 4, { fit: 'cover', fx: undefined, fy: undefined, fw: undefined, rot: 0 });
  clip = clipById(x, id);
  assert.equal(clip.keys.fx, undefined);
  assert.equal(clip.keys.fw, undefined);
  assert.equal(clip.fit, 'cover');
});

test('movement goes with the clip: moved along the timeline, split, and handed to the renderer as its keys', () => {
  const { doc, id } = withShape();
  let x = addMotionPoint(doc, id, 2);
  x = patchAtTime(x, id, 6, { fx: 0.7 });
  // The renderer is handed the keys as they are kept.
  const patch = planPatches(x).find((p) => p.id === id);
  assert.deepEqual(patch.keys.fx, [{ t: 0, v: 0.3 }, { t: 4, v: 0.7 }]);
  // Split at four seconds: the second half starts where the first had got to.
  const split = splitClips(x, [id], 4).doc;
  const halves = split.patches.filter((p) => p.type === 'shape').sort((a, b) => a.timing.start - b.timing.start);
  assert.equal(halves.length, 2);
  near(motionValue(halves[0], 'fx', lengthOf(halves[0])), 0.5, 'the first half ends half way');
  near(motionValue(halves[1], 'fx', 0), 0.5, 'the second half starts there');
  near(motionValue(halves[1], 'fx', 2), 0.7, 'and arrives where the whole clip did');
});

test('a change to a moving clip is one clip\'s: the clip handed in is not changed', () => {
  const { doc, id } = withShape();
  const x = addMotionPoint(doc, id, 2);
  const before = JSON.stringify(clipById(x, id));
  const next = changeAtTime(clipById(x, id), 3, { fx: 0.9 });
  assert.equal(JSON.stringify(clipById(x, id)), before);
  assert.deepEqual(motionTimes(next), [0, 3]);
});

test('a length typed for an animation is held to half the clip, as one chosen from the list is', () => {
  let doc = emptyProject();
  const a = addMedia(doc, A); doc = a.doc; // six seconds
  const clip = clipById(doc, a.ids[0]);
  assert.equal(animationMost(clip), 3);
  // No animation yet: a length typed makes it a fade, of no more than half the clip.
  assert.deepEqual(animationLasting(clip, 'in', 1.5), { ...ANIMATIONS.fade, seconds: 1.5 });
  assert.equal(animationLasting(clip, 'in', 900).seconds, 3, 'a length with no end is held to half the clip');
  assert.equal(animationLasting(clip, 'out', 3.01).seconds, 3);
  // The kind already chosen is kept, at the typed length.
  const rising = { ...clip, enter: animationEnvelope('rise', null, clip), exit: animationEnvelope('left', null, clip) };
  assert.equal(animationName(animationLasting(rising, 'in', 2)), 'rise');
  assert.equal(animationLasting(rising, 'in', 2).seconds, 2);
  assert.equal(animationName(animationLasting(rising, 'out', 99)), 'left');
  assert.equal(animationLasting(rising, 'out', 99).seconds, 3);
  // No length, or what is not a length, is no animation.
  for (const none of [0, -1, NaN, '', null, undefined]) assert.deepEqual(animationLasting(rising, 'in', none), { kind: 'none', seconds: 0 }, String(none));
  // The two ends at their most meet in the middle and never cross.
  assert.ok(animationLasting(clip, 'in', 900).seconds + animationLasting(clip, 'out', 900).seconds <= lengthOf(clip));
  // A very short clip still has a length an animation can be given.
  assert.equal(animationMost({ ...clip, timing: { start: 0, end: 0.1 } }), 0.1);
});
