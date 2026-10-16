// Video's tracks (owner, 2026-10-16: "ability to split a layer into tow so
// multiple clips can be brought next to each other"): a track inserted, moved
// and split in two at the playhead, each carrying its state; solo leaving the
// other tracks of its kind out of the plan, so the export plays what the
// preview plays; and an edit that would change a clip on a locked track found.
import test from 'node:test';
import assert from 'node:assert/strict';
import { emptyProject, addMedia, insertLane, moveLane, splitLane, brokenLocks, onLockedLane, planPatches, laneClips, clipById, partnersOf, moveClip, removeClips } from '../packages/studio/src/project.js';

const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 6, hasAudio: true };
const B = { url: 'D:/clips/b.webm', name: 'b.webm', kind: 'video', duration: 4, hasAudio: true };
const STILL = { url: 'D:/clips/still.png', name: 'still.png', kind: 'image' };

function project() {
  let doc = emptyProject();
  const a = addMedia(doc, A); doc = a.doc;
  const b = addMedia(doc, B); doc = b.doc;
  const s = addMedia(doc, STILL, { lane: 1, at: 0 }); doc = s.doc;
  return { doc, a: a.ids[0], b: b.ids[0], still: s.ids[0] };
}

test('split a track at the playhead: the clips from there on move up onto a new track, the clip under it split, the sound left in step', () => {
  const { doc, a, b, still } = project();
  const out = splitLane(doc, 'v', 0, 3);
  const v1 = laneClips(out, 'v', 0);
  const v2 = laneClips(out, 'v', 1);
  assert.deepEqual(v1.map((p) => [p.id, p.timing.start, p.timing.end]), [[a, 0, 3]]);
  assert.deepEqual(v2.map((p) => p.timing.start), [3, 6], 'the second half of the first clip and the second clip');
  assert.equal(v2[1].id, b);
  assert.equal(clipById(out, still).lane, 2, 'the track that was above moves up one');
  for (const p of v2) near(partnersOf(out, p.id)[0].timing.start, p.timing.start, 'its sound with it');
  assert.equal(splitLane(doc, 'v', 0, 50), doc, 'nothing after the point: nothing changes');
});

const near = (x, y, msg) => assert.ok(Math.abs(x - y) < 0.002, `${msg}: ${x} is not ${y}`);

test('insert and move tracks: the clips and the track state go together', () => {
  let { doc, a, still } = project();
  doc = { ...doc, lanes: { v: { 1: { hidden: true } }, a: {} } };
  const inserted = insertLane(doc, 'v', 1);
  assert.equal(clipById(inserted, still).lane, 2);
  assert.deepEqual(inserted.lanes.v, { 2: { hidden: true } }, 'the hidden track is still the hidden one');
  const moved = moveLane(doc, 'v', 1, 0);
  assert.equal(clipById(moved, still).lane, 0, 'the still now draws under the clips');
  assert.equal(clipById(moved, a).lane, 1);
  assert.deepEqual(moved.lanes.v, { 0: { hidden: true } });
  assert.equal(moveLane(doc, 'v', 0, 0), doc);
  assert.equal(moveLane(doc, 'v', 0, 9).patches.find((p) => p.id === a).lane, 1, 'no further than the last track');
});

test('solo: only the soloed tracks of a section play, in the plan the preview and the export both draw', () => {
  const { doc, a, still } = project();
  const soloed = { ...doc, lanes: { v: { 1: { solo: true } }, a: {} } };
  const plan = planPatches(soloed);
  assert.equal(plan.find((p) => p.id === a).visible, false, 'V1 is left out');
  assert.equal(plan.find((p) => p.id === still).visible, true, 'V2 plays');
  assert.ok(plan.filter((p) => p.type === 'sound').every((p) => p.visible), 'a picture solo leaves the sound alone');
  const none = planPatches(doc);
  assert.ok(none.filter((p) => p.timing).every((p) => p.visible !== false));
});

test('locks: an edit that changes, removes or adds a clip on a locked track is found; renumbering whole tracks is not', () => {
  const { doc, a, b } = project();
  const locked = { ...doc, lanes: { v: { 0: { locked: true } }, a: {} } };
  assert.equal(onLockedLane(locked, clipById(locked, a)), true);
  assert.deepEqual(brokenLocks(locked, moveClip(locked, b, { start: 20 })), [b]);
  assert.ok(brokenLocks(locked, removeClips(locked, [a])).includes(a));
  const added = addMedia(locked, STILL, { lane: 0, at: 30 });
  assert.deepEqual(brokenLocks(locked, added.doc), [added.ids[0]]);
  assert.deepEqual(brokenLocks(locked, insertLane(locked, 'v', 0), { lanesMoved: true }), [], 'a track inserted under it carries its lock with it');
  assert.deepEqual(brokenLocks(doc, moveClip(doc, b, { start: 20 })), [], 'nothing locked, nothing found');
});
