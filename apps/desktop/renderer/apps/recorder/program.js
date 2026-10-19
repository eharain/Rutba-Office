// The Recorder's program: the one picture made live from the sources, and the
// file it is recorded to.
//
// A Program draws the scene on a canvas on a timer (not on animation frames,
// which stop while the window is minimised; windows are made with background
// throttling off, so the timer keeps time). The canvas's own stream is the
// picture recorded, with the microphone and the computer's sound mixed in the
// rig, and the same canvas is what a stream will be encoded from. The scene
// can change while it records: the next frame is drawn the new way.
//
// The recording is one WebM written in pieces as it goes, as the recorder's
// takes are, so a recording a crash cut short still plays.

import { sceneLayout, titleBand, logoBox, coverCrop, containRect, OUTPUT_SIZES } from '@rutba/studio/compositor';
import { createClock, createChunkWriter, videoBitrate, AUDIO_BITRATE } from '@rutba/studio/recorder';
import { t } from '@rutba/office-ui';

const FONT = '"Segoe UI", "Segoe UI Variable", system-ui, sans-serif';

/**
 * A stream as a video element the canvas can draw from, playing. It is kept
 * in the page, a pixel square and invisible: a video outside the page was
 * not always given the frames of a still window's capture, and the scene
 * then left the screen out.
 */
function sourceVideo(stream) {
  let shelf = document.getElementById('st-sources-shelf');
  if (!shelf) {
    shelf = document.createElement('div');
    shelf.id = 'st-sources-shelf';
    shelf.setAttribute('aria-hidden', 'true');
    shelf.style.cssText = 'position:fixed;left:0;top:0;width:1px;height:1px;overflow:hidden;opacity:0;pointer-events:none;';
    document.body.appendChild(shelf);
  }
  const v = document.createElement('video');
  v.muted = true;
  v.playsInline = true;
  v.autoplay = true;
  v.style.cssText = 'width:1px;height:1px;';
  v.srcObject = stream;
  shelf.appendChild(v);
  v.play().catch(() => {});
  return v;
}

function letGo(v) {
  if (!v) return;
  v.srcObject = null;
  v.remove();
}

/** The container the program is recorded in. */
export function programType() {
  const R = typeof MediaRecorder !== 'undefined' ? MediaRecorder : null;
  for (const k of ['video/webm;codecs=vp9,opus', 'video/webm;codecs=vp8,opus', 'video/webm']) {
    if (R?.isTypeSupported?.(k)) return k;
  }
  return 'video/webm';
}

function roundRect(ctx, x, y, w, h, r) {
  const k = Math.max(0, Math.min(r, w / 2, h / 2));
  ctx.beginPath();
  ctx.moveTo(x + k, y);
  ctx.arcTo(x + w, y, x + w, y + h, k);
  ctx.arcTo(x + w, y + h, x, y + h, k);
  ctx.arcTo(x, y + h, x, y, k);
  ctx.arcTo(x, y, x + w, y, k);
  ctx.closePath();
}

export class Program {
  /** @param {{ rig, canvas: HTMLCanvasElement, accent?: string }} o */
  constructor({ rig, canvas, accent = '#e03b4c' }) {
    this.rig = rig;
    this.canvas = canvas;
    this.ctx = canvas.getContext('2d');
    this.accent = accent;
    this.settings = {};
    this.videos = { screen: null, camera: null };
    this.streams = { screen: null, camera: null };
    this.logo = null;
    this.logoPath = '';
    this.timer = null;
    this.fps = 30;
    this.frames = 0;
    this.recording = null;
    this.listeners = new Set();
    this.track = null;
    // A source that is live but has given no frame: when it was opened, and what to do about it.
    this.opened = { screen: 0, camera: 0 };
    this.onStall = null;
  }

  /** The settings the next frame is drawn with (normalised by the window). */
  set(settings) {
    this.settings = settings;
    const size = OUTPUT_SIZES[settings.size] || OUTPUT_SIZES[1080];
    if (this.canvas.width !== size.width || this.canvas.height !== size.height) {
      // The size is fixed while it records: a file's frames keep one size.
      if (!this.recording) { this.canvas.width = size.width; this.canvas.height = size.height; }
    }
    if (settings.fps !== this.fps) { this.fps = settings.fps; if (this.timer) this.start(); }
  }

  /** The logo drawn in its corner, from a picture's bytes as a blob URL, or none. */
  setLogo(url) {
    if (!url) { this.logo = null; return; }
    const img = new Image();
    img.onload = () => { this.logo = img; };
    img.src = url;
  }

  /** Keeps the sources' video elements in step with the rig's streams. */
  follow() {
    for (const name of ['screen', 'camera']) {
      const stream = this.rig[name].stream || null;
      if (stream === this.streams[name]) continue;
      this.streams[name] = stream;
      letGo(this.videos[name]);
      this.videos[name] = stream ? sourceVideo(stream) : null;
      this.opened[name] = stream ? performance.now() : 0;
    }
    // Three seconds live without a single frame: the window is told, once per opening.
    for (const name of ['screen', 'camera']) {
      const v = this.videos[name];
      if (v && !v.videoWidth && this.opened[name] && performance.now() - this.opened[name] > 3000) {
        this.opened[name] = 0;
        this.onStall?.(name);
      }
    }
    // A source that has not started playing is asked again.
    for (const v of Object.values(this.videos)) if (v && v.paused) v.play().catch(() => {});
  }

  start() {
    clearInterval(this.timer);
    this.timer = setInterval(() => this.draw(), Math.round(1000 / this.fps));
  }

  stopDrawing() {
    clearInterval(this.timer);
    this.timer = null;
  }

  onFrame(fn) {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  /** One frame: the scene's background, its sources, then the title and the logo over them. */
  draw() {
    this.follow();
    const { ctx, canvas } = this;
    const W = canvas.width;
    const H = canvas.height;
    const s = this.settings;
    const size = (v) => (v && v.videoWidth ? { w: v.videoWidth, h: v.videoHeight } : null);
    const layout = sceneLayout(s.scene, { W, H, screen: size(this.videos.screen), camera: size(this.videos.camera), settings: s });

    // The ground.
    if (layout.background === 'card' || layout.background === 'studio') {
      const g = ctx.createLinearGradient(0, 0, W, H);
      g.addColorStop(0, '#1d2230');
      g.addColorStop(1, '#0d1018');
      ctx.fillStyle = g;
    } else {
      ctx.fillStyle = '#000';
    }
    ctx.fillRect(0, 0, W, H);

    for (const layer of layout.layers) {
      const v = this.videos[layer.source];
      if (!v || !v.videoWidth) continue;
      const { box } = layer;
      ctx.save();
      if (layer.shape === 'circle') {
        ctx.beginPath();
        ctx.ellipse(box.x + box.w / 2, box.y + box.h / 2, box.w / 2, box.h / 2, 0, 0, Math.PI * 2);
        ctx.clip();
      } else if (layer.shape === 'rounded') {
        roundRect(ctx, box.x, box.y, box.w, box.h, Math.min(box.w, box.h) * 0.06);
        ctx.clip();
      }
      if (layer.fit === 'cover') {
        const c = coverCrop(v.videoWidth, v.videoHeight, box);
        if (layer.mirror) {
          ctx.translate(box.x + box.w, box.y);
          ctx.scale(-1, 1);
          ctx.drawImage(v, c.sx, c.sy, c.sw, c.sh, 0, 0, box.w, box.h);
        } else {
          ctx.drawImage(v, c.sx, c.sy, c.sw, c.sh, box.x, box.y, box.w, box.h);
        }
      } else {
        const r = containRect(v.videoWidth, v.videoHeight, box);
        ctx.drawImage(v, r.x, r.y, r.w, r.h);
      }
      ctx.restore();
      // A ring round the camera over the screen, so it reads as a window onto the person.
      if (layer.source === 'camera' && layout.layers.length > 1) {
        ctx.save();
        ctx.lineWidth = Math.max(2, W * 0.002);
        ctx.strokeStyle = 'rgba(255,255,255,0.85)';
        if (layer.shape === 'circle') { ctx.beginPath(); ctx.ellipse(box.x + box.w / 2, box.y + box.h / 2, box.w / 2, box.h / 2, 0, 0, Math.PI * 2); ctx.stroke(); }
        else if (layer.shape === 'rounded') { roundRect(ctx, box.x, box.y, box.w, box.h, Math.min(box.w, box.h) * 0.06); ctx.stroke(); }
        else ctx.strokeRect(box.x, box.y, box.w, box.h);
        ctx.restore();
      }
    }

    // A card: its words in the middle, large.
    if (layout.card) {
      ctx.save();
      ctx.fillStyle = '#ffffff';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.font = `700 ${Math.round(H * 0.085)}px ${FONT}`;
      ctx.fillText(layout.card.text, W / 2, H * 0.47, W * 0.86);
      ctx.fillStyle = this.accent;
      ctx.fillRect(W / 2 - W * 0.05, H * 0.56, W * 0.1, Math.max(4, H * 0.006));
      ctx.restore();
    }

    // The title band, when it is on and has words.
    if (s.titleOn && (s.title || s.subtitle) && !layout.card) {
      const b = titleBand({ W, H });
      ctx.save();
      ctx.fillStyle = 'rgba(12,14,20,0.82)';
      roundRect(ctx, b.band.x, b.band.y, b.band.w, b.band.h, b.band.h * 0.12);
      ctx.fill();
      ctx.fillStyle = this.accent;
      ctx.fillRect(b.accent.x, b.accent.y, b.accent.w, b.accent.h);
      ctx.fillStyle = '#ffffff';
      ctx.textBaseline = 'alphabetic';
      ctx.font = `700 ${b.title.size}px ${FONT}`;
      // With no line under it, the title sits in the middle of its band.
      ctx.fillText(s.title || '', b.title.x, s.subtitle ? b.title.y : b.band.y + b.band.h * 0.5 + b.title.size * 0.36, b.band.w * 0.9);
      ctx.fillStyle = 'rgba(255,255,255,0.78)';
      ctx.font = `500 ${b.subtitle.size}px ${FONT}`;
      ctx.fillText(s.subtitle || '', b.subtitle.x, b.subtitle.y, b.band.w * 0.9);
      ctx.restore();
    }

    if (this.logo && this.logo.naturalWidth) {
      const r = logoBox({ W, H, w: this.logo.naturalWidth, h: this.logo.naturalHeight, corner: s.logoCorner });
      ctx.drawImage(this.logo, r.x, r.y, r.w, r.h);
    }

    this.frames += 1;
    // What was drawn, for the checks: how many frames, which scene, which sources.
    canvas.dataset.frames = String(this.frames);
    canvas.dataset.scene = s.scene || '';
    canvas.dataset.layers = layout.layers.map((l) => l.source).join(' ');
    { const v = this.videos.screen; const tr = this.streams.screen?.getVideoTracks()[0]; canvas.dataset.screen = v ? `${v.videoWidth}x${v.videoHeight} r${v.readyState} p${v.paused ? 1 : 0} ${tr?.readyState} m${tr?.muted ? 1 : 0}` : 'none'; }
    canvas.dataset.title = s.titleOn && (s.title || s.subtitle) && !layout.card ? '1' : '0';
    this.track?.requestFrame?.();
    for (const fn of this.listeners) fn(canvas);
  }

  /** The program as a stream: the canvas's picture, with the rig's sound mixed in. */
  stream({ sound = true } = {}) {
    if (!this.track || this.track.readyState === 'ended') {
      const s = this.canvas.captureStream(0);
      this.track = s.getVideoTracks()[0];
    }
    const tracks = [this.track];
    const mixed = sound ? this.rig.mix(['mic', 'system']) : null;
    tracks.push(mixed || this.rig.silence());
    return { stream: new MediaStream(tracks), heard: Boolean(mixed) };
  }

  /** Start recording the program to `path`; answers when it has begun. */
  async record({ shell, path }) {
    const type = programType();
    const { stream, heard } = this.stream();
    const bits = videoBitrate({ width: this.canvas.width, height: this.canvas.height, fps: this.fps, quality: this.settings.quality });
    let rec;
    try {
      rec = new MediaRecorder(stream, { mimeType: type, videoBitsPerSecond: bits, audioBitsPerSecond: AUDIO_BITRATE });
    } catch {
      rec = new MediaRecorder(stream);
    }
    const writer = createChunkWriter({ append: (bytes) => shell.fs.append({ path, bytes }) });
    rec.ondataavailable = (e) => { if (e.data?.size) writer.push(e.data); };
    const stopped = new Promise((resolve) => rec.addEventListener('stop', resolve, { once: true }));
    const clock = createClock();
    rec.start(1000);
    clock.start(performance.now());
    this.recording = { rec, writer, stopped, clock, path, heard, type };
    return this.recording;
  }

  get elapsed() { return this.recording ? this.recording.clock.elapsed(performance.now()) : 0; }

  get paused() { return Boolean(this.recording?.clock.paused); }

  pause() {
    const r = this.recording;
    if (!r || r.clock.paused) return;
    if (r.rec.state === 'recording') r.rec.pause();
    r.clock.pause(performance.now());
  }

  resume() {
    const r = this.recording;
    if (!r || !r.clock.paused) return;
    if (r.rec.state === 'paused') r.rec.resume();
    r.clock.resume(performance.now());
  }

  /** Stop: every piece written. Answers { path, written, seconds, heard }, or throws why it could not be kept. */
  async stop() {
    const r = this.recording;
    if (!r) return null;
    const seconds = this.elapsed / 1000;
    if (r.rec.state !== 'inactive') { try { r.rec.stop(); } catch { /* stopped */ } }
    await r.stopped;
    const done = await r.writer.done();
    this.recording = null;
    if (done.failed) throw new Error(t('The recording could not be kept: {reason}', { reason: done.failed.message || String(done.failed) }));
    return { path: r.path, written: done.written, seconds, heard: r.heard };
  }

  dispose() {
    this.stopDrawing();
    if (this.recording) { try { this.recording.rec.stop(); } catch { /* stopped */ } }
    for (const v of Object.values(this.videos)) letGo(v);
    this.videos = { screen: null, camera: null };
    try { this.track?.stop(); } catch { /* ended */ }
    this.listeners.clear();
  }
}
