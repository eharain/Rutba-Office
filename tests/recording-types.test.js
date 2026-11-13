// The types recordings are made in (gate entry L.0002, 14 Nov 2026): H.264 in
// MP4 as avc3 where the machine records it, so an encoder that starts again
// part way still leaves frames that decode; avc1 as the fall-back; WebM where
// MP4 is not recorded. A recorder whose constructor will not take the type it
// was said to support falls back within the file's own container, never to
// WebM in a file named .mp4.
import test from 'node:test';
import assert from 'node:assert/strict';
import { RECORDING_TYPES, supportedTypes, recordingType, typesFrom, openRecorder, takeType } from '../apps/desktop/renderer/recording/session.js';
import { pickMimeType } from '../packages/studio/src/index.js';

/** A MediaRecorder that says yes to `says` and whose constructor takes only `takes` (all it says yes to, unless given). */
function fakeRecorder({ says, takes = says, throwsOn = [] }) {
  return class Fake {
    static isTypeSupported(t) {
      if (throwsOn.includes(t)) throw new Error('asked badly');
      return says.includes(t);
    }
    constructor(stream, options) {
      if (!takes.includes(options.mimeType)) throw new DOMException(`not ${options.mimeType}`, 'NotSupportedError');
      this.stream = stream;
      this.options = options;
      this.mimeType = options.mimeType;
    }
  };
}

const ALL_MP4 = ['video/mp4;codecs=avc3', 'video/mp4;codecs=avc1', 'video/mp4', 'video/webm;codecs=vp9', 'video/webm'];

test('a slide is recorded as avc3 where the machine makes it, avc1 where it makes only that, then MP4, then WebM', () => {
  assert.equal(RECORDING_TYPES[0], 'video/mp4;codecs=avc3');
  assert.equal(recordingType(fakeRecorder({ says: ALL_MP4 })), 'video/mp4;codecs=avc3');
  assert.equal(recordingType(fakeRecorder({ says: ALL_MP4.slice(1) })), 'video/mp4;codecs=avc1');
  assert.equal(recordingType(fakeRecorder({ says: ['video/mp4', 'video/webm'] })), 'video/mp4');
  assert.equal(recordingType(fakeRecorder({ says: ['video/webm;codecs=vp9', 'video/webm'] })), 'video/webm;codecs=vp9');
  assert.equal(recordingType(fakeRecorder({ says: [] })), 'video/webm', 'asked of a recorder that makes nothing listed');
  assert.equal(recordingType(null), 'video/webm', 'no MediaRecorder at all');
});

test('a type the recorder throws on when asked is passed over, not the end of the list', () => {
  const R = fakeRecorder({ says: ALL_MP4, throwsOn: ['video/mp4;codecs=avc3'] });
  assert.deepEqual(supportedTypes(RECORDING_TYPES, R), ALL_MP4.slice(1));
  assert.equal(recordingType(R), 'video/mp4;codecs=avc1');
});

test('the fall-backs of a type are its own container, in order, from the type on', () => {
  assert.deepEqual(typesFrom('video/mp4;codecs=avc3'), ['video/mp4;codecs=avc3', 'video/mp4;codecs=avc1', 'video/mp4']);
  assert.deepEqual(typesFrom('video/mp4;codecs=avc1'), ['video/mp4;codecs=avc1', 'video/mp4']);
  assert.deepEqual(typesFrom('video/webm;codecs=vp9'), ['video/webm;codecs=vp9', 'video/webm']);
  // A type not on the list (the program's WebM, a take's) is tried alone.
  assert.deepEqual(typesFrom('video/webm;codecs=vp9,opus'), ['video/webm;codecs=vp9,opus']);
});

test('a recorder is made in avc3 when its constructor takes it, with the options asked for', () => {
  const R = fakeRecorder({ says: ALL_MP4 });
  const stream = { id: 's' };
  const { rec, type } = openRecorder(stream, recordingType(R), { videoBitsPerSecond: 8e6 }, R);
  assert.equal(type, 'video/mp4;codecs=avc3');
  assert.equal(rec.options.mimeType, 'video/mp4;codecs=avc3');
  assert.equal(rec.options.videoBitsPerSecond, 8e6);
  assert.equal(rec.stream, stream);
});

test('a constructor that refuses avc3, though it was said to be supported, falls back to avc1 and says so', () => {
  const R = fakeRecorder({ says: ALL_MP4, takes: ALL_MP4.slice(1) });
  const { rec, type } = openRecorder({}, recordingType(R), {}, R);
  assert.equal(type, 'video/mp4;codecs=avc1', 'the type it was made in, for the warm-up and the file\'s content type');
  assert.equal(rec.mimeType, 'video/mp4;codecs=avc1');
});

test('a fall-back the recorder does not say it makes is skipped', () => {
  const R = fakeRecorder({ says: ['video/mp4;codecs=avc3', 'video/mp4', 'video/webm'], takes: ['video/mp4', 'video/webm'] });
  assert.equal(openRecorder({}, 'video/mp4;codecs=avc3', {}, R).type, 'video/mp4');
});

test('a file chosen as MP4 is never recorded as WebM: with no MP4 type taken, the refusal is passed on', () => {
  const R = fakeRecorder({ says: ALL_MP4, takes: ['video/webm;codecs=vp9', 'video/webm'] });
  assert.throws(() => openRecorder({}, 'video/mp4;codecs=avc3', {}, R), (err) => err.name === 'NotSupportedError' && /video\/mp4/.test(err.message));
});

test('a take for Video stays WebM: the change is to MP4 recordings alone', () => {
  const R = fakeRecorder({ says: ['video/webm;codecs=vp9,opus', 'video/webm', ...ALL_MP4] });
  const original = globalThis.MediaRecorder;
  globalThis.MediaRecorder = R;
  try {
    assert.equal(takeType(), 'video/webm;codecs=vp9,opus');
  } finally {
    globalThis.MediaRecorder = original;
  }
});

test('Video\'s real-time export asks for avc3 first, with and without sound, and falls back to avc1, then WebM', () => {
  const every = ['video/mp4;codecs=avc3.42E01E', 'video/mp4;codecs=avc3', 'video/mp4;codecs=avc1.42E01E', 'video/mp4;codecs=avc1', 'video/mp4',
    'video/mp4;codecs=avc3.42E01E,mp4a.40.2', 'video/mp4;codecs=avc3,mp4a.40.2', 'video/mp4;codecs=avc1.42E01E,mp4a.40.2', 'video/mp4;codecs=avc1,mp4a.40.2',
    'video/webm;codecs=vp9', 'video/webm;codecs=vp9,opus', 'video/webm'];
  assert.equal(pickMimeType(false, fakeRecorder({ says: every })), 'video/mp4;codecs=avc3.42E01E');
  assert.equal(pickMimeType(true, fakeRecorder({ says: every })), 'video/mp4;codecs=avc3.42E01E,mp4a.40.2');
  const noAvc3 = every.filter((t) => !t.includes('avc3'));
  assert.equal(pickMimeType(false, fakeRecorder({ says: noAvc3 })), 'video/mp4;codecs=avc1.42E01E');
  assert.equal(pickMimeType(true, fakeRecorder({ says: noAvc3 })), 'video/mp4;codecs=avc1.42E01E,mp4a.40.2');
  assert.equal(pickMimeType(true, fakeRecorder({ says: ['video/webm;codecs=vp9,opus', 'video/webm'] })), 'video/webm;codecs=vp9,opus');
  assert.equal(pickMimeType(true, null), null);
});
