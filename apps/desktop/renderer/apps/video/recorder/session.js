// The recorder's media: what is being recorded and the files it goes to
// (docs/STUDIO.md section 8, decision D6).
//
// A rig holds the live sources while the recorder is open: the screen (a
// whole screen or a window, through the shell's capture sources), the camera
// (shared with Presentations' cameos), the microphone and, where the
// platform hears it, the computer's own sound. Every sound goes through one
// AudioContext, which gives the level meters somewhere to listen and lets
// the microphone and the computer's sound be mixed into one track.
//
// Record makes one file per picture, each with its own MediaRecorder, all
// started on the same tick: the screen with the computer's sound, the camera
// with the voice (without a camera, the screen's file takes both sounds,
// mixed). Each recorder hands its pieces over every second and they go
// straight to the file, a few megabytes a message (createChunkWriter), so a
// long take never sits in the window. Pause and resume hold every recorder
// at once, so the files stay on one clock.

import { createClock, createChunkWriter, capSize, videoBitrate, AUDIO_BITRATE } from '@rutba/studio/recorder';
import { openCameraPatiently, closeCamera, cameraTrouble } from './camera.js';
import { t } from '@rutba/office-ui';

/** The container a recording for a slide is made in: MP4 first, as PowerPoint plays it everywhere. */
export function recordingType() {
  const R = typeof MediaRecorder !== 'undefined' ? MediaRecorder : null;
  for (const k of ['video/mp4;codecs=avc1', 'video/mp4', 'video/webm;codecs=vp9', 'video/webm']) {
    if (R?.isTypeSupported?.(k)) return k;
  }
  return 'video/webm';
}

/**
 * The container a take for Video is made in: WebM, which a recorder writes
 * as a stream that plays to its last piece, so a take a crash cut short
 * still opens.
 */
export function takeType() {
  const R = typeof MediaRecorder !== 'undefined' ? MediaRecorder : null;
  for (const k of ['video/webm;codecs=vp9,opus', 'video/webm;codecs=vp8,opus', 'video/webm']) {
    if (R?.isTypeSupported?.(k)) return k;
  }
  return 'video/webm';
}

const stopAll = (stream) => { for (const k of stream?.getTracks() || []) { try { k.stop(); } catch { /* ended */ } } };

/** How loud a sound is just now, 0 to 1, read off its analyser. */
function levelOf(analyser, bins) {
  if (!analyser) return 0;
  analyser.getFloatTimeDomainData(bins);
  let sum = 0;
  for (let i = 0; i < bins.length; i++) sum += bins[i] * bins[i];
  // Speech rarely passes 0.3 RMS: scaled so a voice fills most of the meter.
  return Math.min(1, Math.sqrt(sum / bins.length) * 3.2);
}

export class RecorderRig {
  /** @param {{ shell, onChange?: () => void }} o */
  constructor({ shell, onChange = null }) {
    this.shell = shell;
    this.onChange = onChange;
    this.screen = { stream: null, source: null, error: null, key: '' };
    this.camera = { stream: null, error: null, key: '' };
    this.mic = { stream: null, error: null, key: '' };
    this.system = { error: null };
    this.ctx = null;
    this.nodes = {};
    this.bins = new Float32Array(1024);
    this.recording = null;
    this.disposed = false;
    this.seq = 0;
  }

  changed() { if (!this.disposed) this.onChange?.(); }

  audio() {
    if (!this.ctx) {
      const Ctx = window.AudioContext || window.webkitAudioContext;
      this.ctx = new Ctx();
    }
    if (this.ctx.state === 'suspended') this.ctx.resume().catch(() => {});
    return this.ctx;
  }

  /** A sound track into the graph, with an analyser for its meter. */
  plug(name, track) {
    this.unplug(name);
    if (!track) return;
    const ctx = this.audio();
    const src = ctx.createMediaStreamSource(new MediaStream([track]));
    const analyser = ctx.createAnalyser();
    analyser.fftSize = 2048;
    src.connect(analyser);
    this.nodes[name] = { src, analyser };
  }

  unplug(name) {
    const n = this.nodes[name];
    if (!n) return;
    try { n.src.disconnect(); } catch { /* already */ }
    delete this.nodes[name];
  }

  /** The meter's reading for 'mic' or 'system'. */
  level(name) {
    return levelOf(this.nodes[name]?.analyser, this.bins);
  }

  /**
   * The screen or the window to record, at the frame rate and size asked,
   * with the computer's own sound when `system` is on and the platform hears
   * it. A source already open the same way is kept.
   */
  async setScreen(source, { fps = 30, cap = 1080, system = false } = {}) {
    const key = source ? JSON.stringify([source.id, fps, cap, system]) : '';
    if (key === this.screen.key && (this.screen.stream || this.screen.error)) return;
    const mine = ++this.seq;
    stopAll(this.screen.stream);
    this.unplug('system');
    this.screen = { stream: null, source, error: null, key };
    this.system = { error: null };
    this.changed();
    if (!source) return;
    const size = source.display ? capSize(source.display.width, source.display.height, cap) : cap ? { width: Math.round((cap * 16) / 9), height: cap } : null;
    let stream = null;
    let error = null;
    if (system) {
      // The computer's own sound comes only with getDisplayMedia, which the
      // main process answers with this source and its loopback (R4).
      try {
        const { loopback } = await this.shell.capture.choose({ id: source.id, audio: true });
        stream = await navigator.mediaDevices.getDisplayMedia({
          video: { frameRate: { ideal: fps, max: fps }, ...(size ? { width: { max: size.width }, height: { max: size.height } } : {}) },
          audio: true,
        });
        if (!loopback || !stream.getAudioTracks().length) this.system = { error: t('This computer does not let its own sound be recorded.') };
      } catch {
        stream = null;
        this.system = { error: t('The computer\'s own sound could not be recorded.') };
      }
    }
    if (!stream) {
      try {
        stream = await navigator.mediaDevices.getUserMedia({
          audio: false,
          video: { mandatory: { chromeMediaSource: 'desktop', chromeMediaSourceId: source.id, maxFrameRate: fps, ...(size ? { maxWidth: size.width, maxHeight: size.height } : {}) } },
        });
      } catch (err) {
        error = t('{name} could not be recorded: {reason}', { name: source.name, reason: err?.message || String(err) });
      }
    }
    if (mine !== this.seq || this.disposed) { stopAll(stream); return; }
    this.screen = { stream, source, error, key };
    const sound = stream?.getAudioTracks()[0] || null;
    if (sound) this.plug('system', sound);
    const track = stream?.getVideoTracks()[0];
    if (track) track.addEventListener('ended', () => { if (this.screen.stream === stream) { this.screen = { ...this.screen, error: t('{name} closed or stopped sharing.', { name: source.name }) }; this.changed(); } });
    this.changed();
  }

  /** The camera on (which one, how tall a picture) or off. */
  async setCamera(on, { cameraId = '', cameraHeight = 720 } = {}) {
    const key = on ? JSON.stringify([cameraId, cameraHeight]) : '';
    if (key === this.camera.key && (this.camera.stream || this.camera.error || !on)) return;
    const old = this.camera.stream;
    this.camera = { stream: null, error: null, key };
    if (old) closeCamera(old);
    this.changed();
    if (!on) return;
    try {
      const stream = await openCameraPatiently({ deviceId: cameraId, height: cameraHeight });
      if (this.camera.key !== key || this.disposed) { closeCamera(stream); return; }
      this.camera = { stream, error: null, key };
    } catch (err) {
      if (this.camera.key === key) this.camera = { stream: null, error: cameraTrouble(err), key };
    }
    this.changed();
  }

  /** The microphone on (which one) or off. */
  async setMic(on, { micId = '' } = {}) {
    const key = on ? JSON.stringify([micId]) : '';
    if (key === this.mic.key && (this.mic.stream || this.mic.error || !on)) return;
    stopAll(this.mic.stream);
    this.unplug('mic');
    this.mic = { stream: null, error: null, key };
    this.changed();
    if (!on) return;
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: micId ? { deviceId: { exact: micId } } : true, video: false });
      if (this.mic.key !== key || this.disposed) { stopAll(stream); return; }
      this.mic = { stream, error: null, key };
      this.plug('mic', stream.getAudioTracks()[0]);
    } catch (err) {
      if (this.mic.key === key) this.mic = { stream: null, error: err?.name === 'NotAllowedError' ? t('The microphone was not allowed.') : t('No microphone was found.'), key };
    }
    this.changed();
  }

  /** The cameras and microphones there are, by name. */
  static async devices() {
    try {
      const all = await navigator.mediaDevices.enumerateDevices();
      return {
        cameras: all.filter((d) => d.kind === 'videoinput').map((d, i) => ({ id: d.deviceId, name: d.label || t('Camera {n}', { n: i + 1 }) })),
        mics: all.filter((d) => d.kind === 'audioinput' && d.deviceId !== 'default' && d.deviceId !== 'communications').map((d, i) => ({ id: d.deviceId, name: d.label || t('Microphone {n}', { n: i + 1 }) })),
      };
    } catch {
      return { cameras: [], mics: [] };
    }
  }

  /** A sound track for a file: the named sources mixed in the graph, or none. */
  mix(names) {
    const present = names.filter((n) => this.nodes[n]);
    if (!present.length) return null;
    const dest = this.audio().createMediaStreamDestination();
    for (const n of present) this.nodes[n].src.connect(dest);
    return dest.stream.getAudioTracks()[0] || null;
  }

  /** A silent sound track, for a file that records no sound. */
  silence() {
    const ctx = this.audio();
    const src = ctx.createConstantSource();
    src.offset.value = 0;
    const dest = ctx.createMediaStreamDestination();
    src.connect(dest);
    src.start();
    return dest.stream.getAudioTracks()[0] || null;
  }

  /**
   * Start recording what is open now. `file(role, ext)` makes the file a
   * take goes to and answers its path. Answers when every recorder has started.
   */
  async record({ settings, file, slide = false }) {
    if (!this.screen.stream) throw new Error(t('Choose a screen or a window to record.'));
    const withCamera = Boolean(settings.camera && this.camera.stream && !slide);
    const screenSound = withCamera ? this.mix(['system']) : this.mix(['mic', 'system']);
    const cameraSound = withCamera ? this.mix(['mic']) : null;
    const type = slide ? recordingType() : takeType();
    const contentType = type.split(';')[0];
    const ext = contentType === 'video/mp4' ? 'mp4' : 'webm';
    // A file with no sound to record still gets a silent track: a recorder
    // with a picture alone and a window that seldom changes kept the paused
    // stretch in the file, and stopped its clock at the last change.
    const plan = [{ role: 'screen', video: this.screen.stream.getVideoTracks()[0], sound: screenSound || this.silence(), heard: Boolean(screenSound) }];
    if (withCamera) plan.push({ role: 'camera', video: this.camera.stream.getVideoTracks()[0], sound: cameraSound || this.silence(), heard: Boolean(cameraSound) });
    const takes = [];
    for (const p of plan) {
      const path = await file(p.role, ext);
      const s = p.video.getSettings?.() || {};
      const fps = p.role === 'screen' ? settings.fps : 30;
      const bits = videoBitrate({ width: s.width || 1280, height: s.height || 720, fps, quality: settings.quality });
      const stream = new MediaStream([p.video, ...(p.sound ? [p.sound] : [])]);
      let rec;
      try {
        rec = new MediaRecorder(stream, { mimeType: type, videoBitsPerSecond: bits, audioBitsPerSecond: AUDIO_BITRATE });
      } catch {
        rec = new MediaRecorder(stream);
      }
      const writer = createChunkWriter({ append: (bytes) => this.shell.fs.append({ path, bytes }) });
      rec.ondataavailable = (e) => { if (e.data?.size) writer.push(e.data); };
      const stopped = new Promise((resolve) => rec.addEventListener('stop', resolve, { once: true }));
      takes.push({ role: p.role, path, contentType, ext, rec, writer, stopped, hasAudio: p.heard, width: s.width || null, height: s.height || null, bits });
    }
    const clock = createClock();
    // One tick for all: every recorder started before any piece is handed over.
    for (const k of takes) k.rec.start(1000);
    clock.start(performance.now());
    this.recording = { takes, clock, startedAt: Date.now() };
    this.changed();
    return this.recording;
  }

  get elapsed() { return this.recording ? this.recording.clock.elapsed(performance.now()) : 0; }

  get paused() { return Boolean(this.recording?.clock.paused); }

  pause() {
    const r = this.recording;
    if (!r || r.clock.paused) return;
    for (const k of r.takes) if (k.rec.state === 'recording') k.rec.pause();
    r.clock.pause(performance.now());
    this.changed();
  }

  resume() {
    const r = this.recording;
    if (!r || !r.clock.paused) return;
    for (const k of r.takes) if (k.rec.state === 'paused') k.rec.resume();
    r.clock.resume(performance.now());
    this.changed();
  }

  /** What has been recorded so far: the takes, as long as the clock says. */
  snapshot() {
    const r = this.recording;
    if (!r) return [];
    const seconds = this.elapsed / 1000;
    return r.takes.map((k) => ({ role: k.role, path: k.path, contentType: k.contentType, hasAudio: k.hasAudio, width: k.width, height: k.height, duration: seconds, written: k.writer.written }));
  }

  /** Stop: every recorder, every piece written. Answers the takes, or throws why one could not be kept. */
  async stop() {
    const r = this.recording;
    if (!r) return [];
    const seconds = this.elapsed / 1000;
    for (const k of r.takes) if (k.rec.state !== 'inactive') { try { k.rec.stop(); } catch { /* stopped */ } }
    await Promise.all(r.takes.map((k) => k.stopped));
    const out = [];
    let failed = null;
    for (const k of r.takes) {
      const done = await k.writer.done();
      if (done.failed) failed = done.failed;
      out.push({ role: k.role, path: k.path, contentType: k.contentType, ext: k.ext, hasAudio: k.hasAudio, width: k.width, height: k.height, duration: seconds, written: done.written });
    }
    this.recording = null;
    this.changed();
    if (failed) {
      for (const k of out) this.shell.fs.dropTemp({ path: k.path }).catch(() => {});
      throw new Error(t('The recording could not be kept: {reason}', { reason: failed.message || String(failed) }));
    }
    return out;
  }

  /** Stop and throw the files away. */
  async cancel() {
    const r = this.recording;
    this.recording = null;
    if (!r) return;
    for (const k of r.takes) if (k.rec.state !== 'inactive') { try { k.rec.stop(); } catch { /* stopped */ } }
    await Promise.all(r.takes.map((k) => k.stopped));
    for (const k of r.takes) {
      await k.writer.done();
      await this.shell.fs.dropTemp({ path: k.path }).catch(() => {});
    }
    this.changed();
  }

  /** Let every source go. */
  dispose() {
    this.disposed = true;
    stopAll(this.screen.stream);
    stopAll(this.mic.stream);
    if (this.camera.stream) closeCamera(this.camera.stream);
    this.screen = { stream: null, source: null, error: null, key: '' };
    this.camera = { stream: null, error: null, key: '' };
    this.mic = { stream: null, error: null, key: '' };
    for (const n of Object.keys(this.nodes)) this.unplug(n);
    try { this.ctx?.close(); } catch { /* closed */ }
    this.ctx = null;
  }
}
