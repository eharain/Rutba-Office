// Add Media's choice of where (owner, 2026-10-16: "the insert media should
// have choice to add layer or next to existing layer"): after the clips on a
// track, on a new track at the playhead, or at the playhead on a track with
// what follows moved on; a clip's own sound on a sound track free for it, and
// sound files one after another on one track rather than a track each.
import test from 'node:test';
import assert from 'node:assert/strict';
import { emptyProject, addMedia, placeMedia, musicLane, freeSoundLane, pushLater, clipById, laneClips, partnersOf } from '../packages/studio/src/project.js';

const A = { url: 'D:/clips/a.webm', name: 'a.webm', kind: 'video', duration: 6, hasAudio: true };
const B = { url: 'D:/clips/b.webm', name: 'b.webm', kind: 'video', duration: 4, hasAudio: true };
const STILL = { url: 'D:/clips/still.png', name: 'still.png', kind: 'image' };
const TONE = { url: 'D:/clips/tone.wav', name: 'tone.wav', kind: 'audio', duration: 3 };
const SONG = { url: 'D:/clips/song.mp3', name: 'song.mp3', kind: 'audio', duration: 8 };

const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 0.002, `${msg}: ${a} is not ${b}`);
const lanes = (doc, ids) => ids.map((id) => [clipById(doc, id).type, clipById(doc, id).lane ?? 0, clipById(doc, id).timing.start]);

function twoClips() {
  let doc = emptyProject();
  doc = addMedia(doc, A).doc;
  doc = addMedia(doc, B).doc;
  return doc;
}

test('at the end: pictures after the clips on the chosen track, sound files one after another on one track', () => {
  const doc = twoClips();
  const { doc: out, ids } = placeMedia(doc, [STILL, TONE, SONG], { where: 'end', lane: 0 });
  const still = clipById(out, ids[0]);
  near(still.timing.start, 10, 'the still follows the two clips');
  const [tone, song] = ids.slice(1).map((id) => clipById(out, id));
  assert.equal(tone.lane, song.lane, 'both sound files on one track');
  assert.notEqual(tone.lane, 0, 'not on the clips’ own sound track');
  near(song.timing.start, tone.timing.end, 'the second sound file after the first');
  // More music later goes after it on the same track.
  const again = placeMedia(out, [TONE], { where: 'end' });
  const more = clipById(again.doc, again.ids[0]);
  assert.equal(more.lane, tone.lane);
  near(more.timing.start, song.timing.end, 'after the music already there');
});

test('on a new track: above the picture tracks there are, from the playhead, with a new sound track for sound files', () => {
  const doc = twoClips();
  const { doc: out, ids } = placeMedia(doc, [STILL, B, TONE], { where: 'layer', at: 2 });
  const [still, video] = [clipById(out, ids[0]), clipById(out, ids[1])];
  assert.equal(still.lane, 1, 'a track above V1');
  assert.equal(video.lane, 1);
  near(still.timing.start, 2, 'from the playhead');
  near(video.timing.start, 7, 'the clip after the still');
  // The new clip's own sound does not cover the clips' sound on A1.
  const sound = partnersOf(out, video.id)[0];
  const a1 = laneClips(out, 'a', 0).filter((p) => p.id !== sound.id);
  assert.ok(a1.every((p) => sound.lane !== 0 || sound.timing.end <= p.timing.start || sound.timing.start >= p.timing.end), 'the sound covers nothing');
  assert.notEqual(sound.lane, 0, 'A1 is taken from 7 s on, so its sound goes lower');
  const tone = clipById(out, ids[ids.length - 1]);
  assert.ok(tone.lane >= 1, 'a sound track of its own');
  near(tone.timing.start, 2, 'from the playhead');
  // Nothing already there moved.
  near(laneClips(out, 'v', 0)[1].timing.start, 6, 'the second clip stays');
});

test('inserted at the playhead: the clip under it split, what follows moved on with its sound, the new clip in the room made', () => {
  const doc = twoClips();
  const { doc: out, ids } = placeMedia(doc, [STILL], { where: 'insert', lane: 0, at: 3 });
  const v1 = laneClips(out, 'v', 0);
  assert.deepEqual(v1.map((p) => [p.type, p.timing.start, p.timing.end]), [
    ['video', 0, 3], ['image', 3, 8], ['video', 8, 11], ['video', 11, 15],
  ]);
  assert.equal(v1[1].id, ids[0]);
  // Each moved picture still has its sound under it.
  for (const p of v1.filter((x) => x.type === 'video')) {
    const s = partnersOf(out, p.id)[0];
    near(s.timing.start, p.timing.start, `${p.id}'s sound with it`);
  }
});

test('the helpers: the music track, a free sound track and making room', () => {
  let doc = twoClips();
  assert.equal(musicLane(doc), 1, 'the clips’ sound uses A1, so music goes on A2');
  assert.equal(musicLane(emptyProject()), 0);
  doc = addMedia(doc, TONE, { lane: 2 }).doc;
  assert.equal(musicLane(doc), 2, 'where music already is');
  assert.equal(freeSoundLane(doc, 0, 1), 1, 'A1 has the clips’ sound at 0 s');
  assert.equal(freeSoundLane(doc, 20, 1), 0, 'nothing at 20 s');
  assert.equal(pushLater(doc, 'v', 0, 3, 0), doc, 'no room asked, nothing changed');
  assert.deepEqual(lanes(placeMedia(emptyProject(), [A], { where: 'end' }).doc, placeMedia(emptyProject(), [A], { where: 'end' }).ids), [['video', 0, 0], ['sound', 0, 0]]);
});
