// The Recorder's program: the one picture made live from the studio's stack,
// and the file it is recorded to.
//
// A Program draws the stack on a canvas on a timer (not on animation frames,
// which stop while the window is minimised; windows are made with background
// throttling off, so the timer keeps time): the shown items bottom first,
// each its clip of its source in its box and shape, a title as a band. The
// first line's Video switch draws black, or the placeholder, instead. The
// canvas's own stream is the picture recorded, with the studio's one sound
// output, and the same canvas is what a stream is encoded from. The stack can
// change while it records: the next frame is drawn the new way.
//
// The recording is one WebM written in pieces as it goes, as the recorder's
// takes are, so a recording a crash cut short still plays. With the picture
// off it is the sound alone, in a sound file.

import { OUTPUT_SIZES, coverCrop, containRect } from '@rutba/studio/compositor';
import { drawOrder } from '@rutba/studio/stack';
import { createClock, createChunkWriter, videoBitrate, AUDIO_BITRATE } from '@rutba/studio/recorder';
import { voiceType } from '../../recording/session.js';
import { t } from '@rutba/office-ui';

const FONT = '"Segoe UI", "Segoe UI Variable", system-ui, sans-serif';

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
  /** @param {{ media, canvas: HTMLCanvasElement, accent?: string }} o */
  constructor({ media, canvas, accent = '#e03b4c' }) {
    this.media = media;
    this.canvas = canvas;
    this.ctx = canvas.getContext('2d');
    this.accent = accent;
    this.settings = {};
    this.stack = null;
    this.timer = null;
    this.fps = 30;
    this.frames = 0;
    this.recording = null;
    this.listeners = new Set();
    this.track = null;
    this.onTick = null;
  }

  /** The settings the next frame is drawn with (size, frame rate, quality). */
  set(settings) {
    this.settings = settings;
    const size = OUTPUT_SIZES[settings.size] || OUTPUT_SIZES[1080];
    if (this.canvas.width !== size.width || this.canvas.height !== size.height) {
      // The size is fixed while it records: a file's frames keep one size.
      if (!this.recording) { this.canvas.width = size.width; this.canvas.height = size.height; }
    }
    if (settings.fps !== this.fps) { this.fps = settings.fps; if (this.timer) this.start(); }
  }

  /** The stack the next frame is drawn from. */
  setStack(stack) { this.stack = stack; }

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

  /** One item: its clip of its source, in its box, in its shape, mirrored when it is. */
  drawItem(it, W, H) {
    const { ctx } = this;
    const box = { x: it.box.x * W, y: it.box.y * H, w: it.box.w * W, h: it.box.h * H };
    if (it.kind === 'title') return this.drawTitle(it, box);
    const d = this.media.drawable(it);
    if (!d) return false;
    const sw = d.videoWidth || d.naturalWidth || d.width;
    const sh = d.videoHeight || d.naturalHeight || d.height;
    if (!sw || !sh) return false;
    const c = it.crop || { x: 0, y: 0, w: 1, h: 1 };
    const part = { sx: c.x * sw, sy: c.y * sh, sw: c.w * sw, sh: c.h * sh };
    ctx.save();
    if (it.shape === 'circle') {
      ctx.beginPath();
      ctx.ellipse(box.x + box.w / 2, box.y + box.h / 2, box.w / 2, box.h / 2, 0, 0, Math.PI * 2);
      ctx.clip();
    } else if (it.shape === 'rounded') {
      roundRect(ctx, box.x, box.y, box.w, box.h, Math.min(box.w, box.h) * 0.06);
      ctx.clip();
    }
    // A rectangle shows the whole of its clip; a circle or a rounded box is filled by it.
    let dest;
    let src = part;
    if (it.shape === 'rect') {
      dest = containRect(part.sw, part.sh, box);
    } else {
      const k = coverCrop(part.sw, part.sh, box);
      src = { sx: part.sx + k.sx, sy: part.sy + k.sy, sw: k.sw, sh: k.sh };
      dest = box;
    }
    if (it.mirror) {
      ctx.translate(dest.x * 2 + dest.w, 0);
      ctx.scale(-1, 1);
    }
    ctx.drawImage(d, src.sx, src.sy, src.sw, src.sh, dest.x, dest.y, dest.w, dest.h);
    ctx.restore();
    // A ring round a shaped picture over another, so it reads as a window onto it.
    if (it.shape !== 'rect' && !(it.box.w >= 0.999 && it.box.h >= 0.999)) {
      ctx.save();
      ctx.lineWidth = Math.max(2, W * 0.002);
      ctx.strokeStyle = 'rgba(255,255,255,0.85)';
      if (it.shape === 'circle') { ctx.beginPath(); ctx.ellipse(box.x + box.w / 2, box.y + box.h / 2, box.w / 2, box.h / 2, 0, 0, Math.PI * 2); ctx.stroke(); }
      else { roundRect(ctx, box.x, box.y, box.w, box.h, Math.min(box.w, box.h) * 0.06); ctx.stroke(); }
      ctx.restore();
    }
    return true;
  }

  /** A title: its band, an accent on its edge, its two lines. Answers whether there was anything to draw. */
  drawTitle(it, box) {
    if (!it.text && !it.subtitle) return false;
    const { ctx } = this;
    ctx.save();
    if (it.shape !== 'rect') { roundRect(ctx, box.x, box.y, box.w, box.h, it.shape === 'circle' ? box.h / 2 : box.h * 0.12); ctx.clip(); }
    ctx.fillStyle = 'rgba(12,14,20,0.82)';
    ctx.fillRect(box.x, box.y, box.w, box.h);
    ctx.fillStyle = this.accent;
    ctx.fillRect(box.x, box.y, Math.max(4, box.w * 0.017), box.h);
    const pad = box.x + Math.max(10, box.w * 0.048);
    const big = Math.round(box.h * 0.34);
    const small = Math.round(box.h * 0.22);
    ctx.fillStyle = '#ffffff';
    ctx.textBaseline = 'alphabetic';
    ctx.font = `700 ${big}px ${FONT}`;
    // With no line under it, the title sits in the middle of its band.
    ctx.fillText(it.text || '', pad, it.subtitle ? box.y + box.h * 0.46 : box.y + box.h * 0.5 + big * 0.36, box.w * 0.9);
    ctx.fillStyle = 'rgba(255,255,255,0.78)';
    ctx.font = `500 ${small}px ${FONT}`;
    ctx.fillText(it.subtitle || '', pad, box.y + box.h * 0.8, box.w * 0.9);
    ctx.restore();
    return true;
  }

  /** The placeholder: the person's picture over the frame, or a card with its words. */
  drawPlaceholder(stack, W, H) {
    const { ctx } = this;
    if (stack.placeholder) this.media.loadImage(stack.placeholder);
    const pic = stack.placeholder ? this.media.drawable({ kind: 'picture', path: stack.placeholder }) : null;
    if (pic) {
      const r = containRect(pic.naturalWidth, pic.naturalHeight, { x: 0, y: 0, w: W, h: H });
      ctx.fillStyle = '#000';
      ctx.fillRect(0, 0, W, H);
      ctx.drawImage(pic, r.x, r.y, r.w, r.h);
      return;
    }
    const g = ctx.createLinearGradient(0, 0, W, H);
    g.addColorStop(0, '#1d2230');
    g.addColorStop(1, '#0d1018');
    ctx.fillStyle = g;
    ctx.fillRect(0, 0, W, H);
    ctx.save();
    ctx.fillStyle = '#ffffff';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.font = `700 ${Math.round(H * 0.085)}px ${FONT}`;
    ctx.fillText(stack.placeholderText || t('Back soon'), W / 2, H * 0.47, W * 0.86);
    ctx.fillStyle = this.accent;
    ctx.fillRect(W / 2 - W * 0.05, H * 0.56, W * 0.1, Math.max(4, H * 0.006));
    ctx.restore();
  }

  /** One frame. */
  draw() {
    this.onTick?.();
    const { ctx, canvas } = this;
    const W = canvas.width;
    const H = canvas.height;
    const stack = this.stack;
    const drawn = [];
    if (!stack || stack.video === 'off') {
      ctx.fillStyle = '#000';
      ctx.fillRect(0, 0, W, H);
    } else if (stack.video === 'placeholder') {
      this.drawPlaceholder(stack, W, H);
    } else {
      ctx.fillStyle = '#000';
      ctx.fillRect(0, 0, W, H);
      for (const it of drawOrder(stack)) if (this.drawItem(it, W, H)) drawn.push(it);
    }
    this.frames += 1;
    // What was drawn, for the checks: how many frames, and the items drawn bottom first, by kind and by id.
    canvas.dataset.frames = String(this.frames);
    canvas.dataset.video = stack?.video || 'on';
    canvas.dataset.layers = drawn.map((it) => it.kind).join(' ');
    canvas.dataset.ids = drawn.map((it) => it.id).join(' ');
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
    tracks.push(sound ? this.media.mix() : this.media.silence());
    return { stream: new MediaStream(tracks), heard: sound && this.media.heard() };
  }

  /**
   * A still picture as PNG bytes: the program as it was last drawn, or one
   * item at its source's own size, through its clip, mirrored as it is
   * shown. Answers { bytes, width, height }, or null when there is nothing
   * to take yet.
   */
  async snapshot(item = null) {
    let canvas = this.canvas;
    if (item && item.kind !== 'title') {
      const d = this.media.drawable(item);
      const w = d?.videoWidth || d?.naturalWidth || d?.width;
      const h = d?.videoHeight || d?.naturalHeight || d?.height;
      if (!w || !h) return null;
      const c = item.crop || { x: 0, y: 0, w: 1, h: 1 };
      const sx = Math.round(c.x * w);
      const sy = Math.round(c.y * h);
      const sw = Math.max(1, Math.min(w - sx, Math.round(c.w * w)));
      const sh = Math.max(1, Math.min(h - sy, Math.round(c.h * h)));
      canvas = document.createElement('canvas');
      canvas.width = sw;
      canvas.height = sh;
      const ctx = canvas.getContext('2d');
      if (item.mirror) { ctx.translate(sw, 0); ctx.scale(-1, 1); }
      ctx.drawImage(d, sx, sy, sw, sh, 0, 0, sw, sh);
    } else if (!this.frames) {
      return null;
    }
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (!blob) return null;
    return { bytes: new Uint8Array(await blob.arrayBuffer()), width: canvas.width, height: canvas.height };
  }

  /** Start recording the program to `path` (picture and sound, or with `soundOnly` the sound alone); answers when it has begun. */
  async record({ shell, path, soundOnly = false }) {
    let rec;
    let heard;
    let type;
    if (soundOnly) {
      ({ type } = voiceType());
      heard = this.media.heard();
      const stream = new MediaStream([this.media.mix()]);
      try { rec = new MediaRecorder(stream, { mimeType: type, audioBitsPerSecond: AUDIO_BITRATE }); } catch { rec = new MediaRecorder(stream); }
    } else {
      type = programType();
      const made = this.stream();
      heard = made.heard;
      const bits = videoBitrate({ width: this.canvas.width, height: this.canvas.height, fps: this.fps, quality: this.settings.quality });
      try {
        rec = new MediaRecorder(made.stream, { mimeType: type, videoBitsPerSecond: bits, audioBitsPerSecond: AUDIO_BITRATE });
      } catch {
        rec = new MediaRecorder(made.stream);
      }
    }
    const writer = createChunkWriter({ append: (bytes) => shell.fs.append({ path, bytes }) });
    rec.ondataavailable = (e) => { if (e.data?.size) writer.push(e.data); };
    const stopped = new Promise((resolve) => rec.addEventListener('stop', resolve, { once: true }));
    const clock = createClock();
    rec.start(1000);
    clock.start(performance.now());
    this.recording = { rec, writer, stopped, clock, path, heard, type, soundOnly };
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
    return { path: r.path, written: done.written, seconds, heard: r.heard, soundOnly: r.soundOnly };
  }

  dispose() {
    this.stopDrawing();
    if (this.recording) { try { this.recording.rec.stop(); } catch { /* stopped */ } }
    try { this.track?.stop(); } catch { /* ended */ }
    this.listeners.clear();
  }
}
