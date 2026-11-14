// The types recordings are made in (gate entry L.0002, 14 Nov 2026): H.264 in
// MP4 as avc3 where the machine records it, so an encoder that starts again
// part way still leaves frames that decode; avc1 as the fall-back; WebM where
// MP4 is not recorded. A recorder whose constructor will not take the type it
// was said to support falls back within the file's own container, never to
// WebM in a file named .mp4.
import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import { RECORDING_TYPES, supportedTypes, recordingType, typesFrom, openRecorder, openProgramRecorder, takeType } from '../apps/desktop/renderer/recording/session.js';
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

/* ── the callers (second review of L.0002, 14 Nov 2026) ─────────────────── */

const here = path.dirname(fileURLToPath(import.meta.url));

/** A window module bundled for node, the UI kit stood in for, as the kit brings React and a window with it. */
async function loadWindowModule(rel) {
  const out = await esbuild.build({
    entryPoints: [path.join(here, rel)],
    bundle: true,
    format: 'esm',
    platform: 'node',
    write: false,
    logLevel: 'silent',
    loader: { '.js': 'jsx' },
    jsx: 'automatic',
    plugins: [{
      name: 'ui-kit',
      setup(b) {
        b.onResolve({ filter: /^@rutba\/office-ui$/ }, () => ({ path: 'ui-kit', namespace: 'stub' }));
        b.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({ resolveDir: path.join(here, '../packages/office-ui/src'), contents: "export { t } from './messages.js'; export const Button = () => null; export const Dialog = () => null; export const Select = () => null;", loader: 'js' }));
      },
    }],
  });
  return import(`data:text/javascript;base64,${Buffer.from(out.outputFiles[0].text).toString('base64')}`);
}

/**
 * A MediaRecorder as fakeRecorder's that also records: every one made is
 * kept in `made` with the type it was made in ('' for one left to choose),
 * and it hands over a piece of data soon after it starts, as a real one does.
 */
function recordingFake({ says, takes = says, plainTakes = true }) {
  const made = [];
  class Fake extends EventTarget {
    static isTypeSupported(t) { return says.includes(t); }
    constructor(stream, options = {}) {
      super();
      const asked = options.mimeType ?? '';
      if (asked ? !takes.includes(asked) : !plainTakes) throw new DOMException(`not ${asked || 'as it comes'}`, 'NotSupportedError');
      this.stream = stream;
      this.mimeType = asked;
      this.state = 'inactive';
      made.push(this);
    }
    emit(name, extra = {}) {
      const e = Object.assign(new Event(name), extra);
      this[`on${name}`]?.(e);
      this.dispatchEvent(e);
    }
    start() {
      this.state = 'recording';
      setTimeout(() => { if (this.state === 'recording') this.emit('dataavailable', { data: new Blob([new Uint8Array(16)]) }); }, 5);
    }
    stop() {
      if (this.state === 'inactive') return;
      this.state = 'inactive';
      setTimeout(() => { this.emit('dataavailable', { data: new Blob([new Uint8Array(16)]) }); this.emit('stop'); }, 1);
    }
    pause() { this.state = 'paused'; }
    resume() { this.state = 'recording'; }
  }
  return { Fake, made };
}

/** Run `fn` with these globals in place, and put back what was there. */
async function withGlobals(set, fn) {
  const was = Object.fromEntries(Object.keys(set).map((k) => [k, Object.getOwnPropertyDescriptor(globalThis, k)]));
  for (const [k, v] of Object.entries(set)) Object.defineProperty(globalThis, k, { value: v, configurable: true, writable: true });
  try {
    return await fn();
  } finally {
    for (const [k, d] of Object.entries(was)) {
      if (d) Object.defineProperty(globalThis, k, d); else delete globalThis[k];
    }
  }
}

test('the program\'s recorder: an MP4 no MP4 type of which is taken is refused, never a recorder left to write WebM', () => {
  const { Fake, made } = recordingFake({ says: ALL_MP4, takes: ['video/webm;codecs=vp9', 'video/webm'] });
  assert.throws(() => openProgramRecorder({}, 'video/mp4;codecs=avc3', {}, Fake), (err) => err.name === 'NotSupportedError');
  assert.equal(made.length, 0, 'no recorder made at all');
  // A WebM type its constructor will not take with these options: the recorder as it comes, WebM too.
  const { Fake: W, made: m2 } = recordingFake({ says: ['video/webm;codecs=vp9,opus'], takes: [] });
  const { rec, type } = openProgramRecorder({}, 'video/webm;codecs=vp9,opus', { videoBitsPerSecond: 1 }, W);
  assert.equal(rec.mimeType, '', 'left to choose');
  assert.equal(type, 'video/webm;codecs=vp9,opus');
  assert.equal(m2.length, 1);
});

/** A Program as the Recorder window makes one, enough of it for record(). */
function programFor(Program) {
  const p = Object.create(Program.prototype);
  p.canvas = { width: 1280, height: 720 };
  p.fps = 30;
  p.settings = { quality: 'high' };
  p.stream = () => ({ stream: { id: 'program' }, heard: true });
  return p;
}
const shellFor = () => ({ fs: { append: async () => ({}) } });

test('a slide\'s recording in the Recorder: with no MP4 taken it fails in plain words, and no WebM recorder is made for its .mp4', async () => {
  const { Program } = await loadWindowModule('../apps/desktop/renderer/studio/program.js');
  const { Fake, made } = recordingFake({ says: ALL_MP4, takes: ['video/webm;codecs=vp9', 'video/webm'] });
  await withGlobals({ MediaRecorder: Fake }, async () => {
    await assert.rejects(programFor(Program).record({ shell: shellFor(), path: 'D:/films/slide.mp4', type: 'video/mp4;codecs=avc3' }), (err) => err.message === 'This system cannot record video.' && err.cause?.name === 'NotSupportedError');
  });
  assert.equal(made.length, 0, 'nothing recording into the .mp4');
});

test('a slide\'s recording in the Recorder falls back from avc3 to avc1 and keeps the type it was made in', async () => {
  const { Program } = await loadWindowModule('../apps/desktop/renderer/studio/program.js');
  const { Fake, made } = recordingFake({ says: ALL_MP4, takes: ALL_MP4.slice(1) });
  const r = await withGlobals({ MediaRecorder: Fake }, () => programFor(Program).record({ shell: shellFor(), path: 'D:/films/slide.mp4', type: 'video/mp4;codecs=avc3' }));
  assert.deepEqual(made.map((m) => m.mimeType), ['video/mp4;codecs=avc1']);
  assert.equal(r.type, 'video/mp4;codecs=avc1');
  made[0].stop();
});

test('the Recorder\'s own program (WebM) still gets a recorder as it comes when its options are refused', async () => {
  const { Program } = await loadWindowModule('../apps/desktop/renderer/studio/program.js');
  const { Fake, made } = recordingFake({ says: ['video/webm;codecs=vp9,opus'], takes: [] });
  const r = await withGlobals({ MediaRecorder: Fake }, () => programFor(Program).record({ shell: shellFor(), path: 'D:/films/program.webm', type: 'video/webm;codecs=vp9,opus' }));
  assert.deepEqual(made.map((m) => m.mimeType), ['']);
  assert.equal(r.type, 'video/webm;codecs=vp9,opus');
  made[0].stop();
});

/** The window's pieces Export to Video draws with, standing in: a canvas, its stream, sound and MediaStream. */
function windowFakes(Recorder) {
  const track = () => ({ stop() {}, requestFrame() {} });
  const ctx = new Proxy({}, { get: (o, k) => (k in o ? o[k] : () => {}), set: (o, k, v) => { o[k] = v; return true; } });
  const canvas = () => ({ width: 0, height: 0, getContext: () => ctx, captureStream: () => { const v = track(); return { getVideoTracks: () => [v], getTracks: () => [v] }; } });
  class AudioContext {
    createMediaStreamDestination() { return { stream: { getAudioTracks: () => [track()] } }; }
    async close() {}
  }
  class MediaStream { constructor(tracks = []) { this.tracks = tracks; } getTracks() { return this.tracks; } }
  return { MediaRecorder: Recorder, AudioContext, MediaStream, document: { createElement: (n) => (n === 'canvas' ? canvas() : {}) } };
}

test('Export to Video: the warm-up and the file\'s content type are the type openRecorder made the recorder in (avc3 refused, avc1)', async () => {
  const { renderVideo } = await loadWindowModule('../apps/desktop/renderer/apps/slides/export-video.js');
  const { Fake, made } = recordingFake({ says: ALL_MP4, takes: ALL_MP4.slice(1) });
  const out = await withGlobals(windowFakes(Fake), () => renderVideo([], { width: 1280, height: 720 }));
  // The deck's recorder first, then the warm-up on its own canvas, both avc1.
  assert.deepEqual(made.map((m) => m.mimeType), ['video/mp4;codecs=avc1', 'video/mp4;codecs=avc1']);
  assert.equal(out.contentType, 'video/mp4');
  assert.ok(out.bytes.length > 0);
});

test('Export to Video records avc3 where it is taken, its warm-up avc3 too', async () => {
  const { renderVideo } = await loadWindowModule('../apps/desktop/renderer/apps/slides/export-video.js');
  const { Fake, made } = recordingFake({ says: ALL_MP4 });
  const out = await withGlobals(windowFakes(Fake), () => renderVideo([], { width: 1280, height: 720 }));
  assert.deepEqual(made.map((m) => m.mimeType), ['video/mp4;codecs=avc3', 'video/mp4;codecs=avc3']);
  assert.equal(out.contentType, 'video/mp4');
});

test('a cameo\'s recording is made in the type openRecorder answers, and kept as that container', async () => {
  const { CameraRecorder } = await loadWindowModule('../apps/desktop/renderer/apps/slides/cameo.js');
  for (const [takes, want] of [[ALL_MP4, 'video/mp4;codecs=avc3'], [ALL_MP4.slice(1), 'video/mp4;codecs=avc1'], [['video/webm;codecs=vp9', 'video/webm'], 'video/webm;codecs=vp9']]) {
    const { Fake, made } = recordingFake({ says: takes, takes });
    // eslint-disable-next-line no-await-in-loop
    const parts = await withGlobals({ MediaRecorder: Fake }, async () => {
      const cam = new CameraRecorder(new Set([0]));
      cam.stream = { id: 'camera' };
      cam._begin(0);
      cam.current.recorder.stop();
      await Promise.all(cam.pending);
      return cam.parts;
    });
    assert.deepEqual(made.map((m) => m.mimeType), [want]);
    assert.equal(parts.length, 1);
    assert.equal(parts[0].type, want.split(';')[0]);
  }
});
