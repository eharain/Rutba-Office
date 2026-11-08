// The studio's media: every screen, window and camera the stack names, open
// while one of its items is started, and every sound, through one graph.
//
// A source is captured once however many items show it (a screen whole and a
// clip of it are two items of one capture). An item stopped lets its source
// go when no other item still holds it, and keeps the last picture it gave as
// a still, drawn in its place (owner, 2026-11-02: "should have stop, start
// toggle buttons as well").
//
// Sound goes through one AudioContext: each microphone and the computer's
// sound into its own gain (its level and its mute) and analyser (its meter),
// all into one output that recordings and streams take their sound from. The
// output is made once, so a recording keeps its sound while sounds are
// started, stopped and muted under it. With Sound off nothing reaches it; with
// Music, the placeholder music does instead of the sounds.

import { openCameraPatiently, closeCamera, cameraTrouble } from '../recording/camera.js';
import { capSize } from '@rutba/studio/recorder';
import { t } from '@rutba/office-ui';

const stopAll = (stream) => { for (const k of stream?.getTracks() || []) { try { k.stop(); } catch { /* ended */ } } };

/** Where the playing video elements are kept: in the page, a pixel, unseen (a video outside the page was not always given a still window's frames). */
function shelf() {
  let el = document.getElementById('st-sources-shelf');
  if (!el) {
    el = document.createElement('div');
    el.id = 'st-sources-shelf';
    el.setAttribute('aria-hidden', 'true');
    el.style.cssText = 'position:fixed;left:0;top:0;width:1px;height:1px;overflow:hidden;opacity:0;pointer-events:none;';
    document.body.appendChild(el);
  }
  return el;
}

function playing(stream) {
  const v = document.createElement('video');
  v.muted = true;
  v.playsInline = true;
  v.autoplay = true;
  v.style.cssText = 'width:1px;height:1px;';
  v.srcObject = stream;
  shelf().appendChild(v);
  v.play().catch(() => {});
  return v;
}

function letGo(v) {
  if (!v) return;
  v.srcObject = null;
  v.remove();
}

/** A video's frame as it is now, on a canvas of its own size, or null. */
function grab(v) {
  if (!v?.videoWidth) return null;
  const c = document.createElement('canvas');
  c.width = v.videoWidth;
  c.height = v.videoHeight;
  c.getContext('2d').drawImage(v, 0, 0);
  return c;
}

/** How loud a sound is just now, 0 to 1. */
function levelOf(analyser, bins) {
  if (!analyser) return 0;
  analyser.getFloatTimeDomainData(bins);
  let sum = 0;
  for (let i = 0; i < bins.length; i++) sum += bins[i] * bins[i];
  return Math.min(1, Math.sqrt(sum / bins.length) * 3.2);
}

/** The key one capture is shared by: its kind and what it is of. */
export const sourceKey = (item) => `${item.kind === 'window' ? 'screen' : item.kind}:${item.ref}`;

export class StudioMedia {
  /** @param {{ shell, onChange?: () => void }} o */
  constructor({ shell, onChange = null }) {
    this.shell = shell;
    this.onChange = onChange;
    this.disposed = false;
    // key -> { stream, video, error, opening, opts }
    this.captures = new Map();
    // key -> canvas: the last picture of a source let go
    this.stills = new Map();
    // path -> { img, url }
    this.images = new Map();
    // sound id -> { stream, src, gain, analyser, key, error }
    this.sounds = new Map();
    this.ctx = null;
    this.out = null;
    this.bins = new Float32Array(1024);
    this.music = null;
    this.mode = 'on';
    this.list = [];
  }

  changed() { if (!this.disposed) this.onChange?.(); }

  /* ── pictures ──────────────────────────────────────────────────────── */

  /**
   * Opens what the stack's started items need and lets go of the rest.
   * `sources` is the capture list (shell.capture.sources()), for a screen or
   * window's size and stand-in.
   */
  sync(stack, sources, { fps = 30, size = 1080 } = {}) {
    this.list = sources || this.list;
    const want = new Map();
    for (const it of stack.items) {
      if (it.kind === 'picture' && it.path) this.loadImage(it.path);
      if (!it.live || !['screen', 'window', 'camera'].includes(it.kind)) continue;
      if (it.kind !== 'camera' && !it.ref) continue;
      want.set(sourceKey(it), it);
    }
    for (const [key, cap] of this.captures) {
      if (want.has(key)) continue;
      const still = grab(cap.video);
      if (still) this.stills.set(key, still);
      this.close(key);
    }
    for (const [key, it] of want) {
      const opts = JSON.stringify([fps, size]);
      const cap = this.captures.get(key);
      if (cap && cap.opts === opts) continue;
      if (cap) this.close(key);
      this.open(key, it, { fps, size, opts });
    }
  }

  async open(key, item, { fps, size, opts }) {
    const cap = { stream: null, video: null, error: null, opening: true, opts };
    this.captures.set(key, cap);
    this.changed();
    let stream = null;
    try {
      if (item.kind === 'camera') {
        stream = await openCameraPatiently({ deviceId: item.ref, height: Math.min(1080, size) });
      } else {
        const src = this.list.find((s) => s.id === item.ref);
        const cut = src?.display ? capSize(src.display.width, src.display.height, size) : { width: Math.round((size * 16) / 9), height: size };
        stream = await navigator.mediaDevices.getUserMedia({
          audio: false,
          // `capture`: what a source is recorded from when it is not itself (a check run's stand-in screen).
          video: { mandatory: { chromeMediaSource: 'desktop', chromeMediaSourceId: src?.capture || item.ref, maxFrameRate: fps, maxWidth: cut.width, maxHeight: cut.height } },
        });
      }
    } catch (err) {
      if (this.captures.get(key) !== cap) return;
      cap.opening = false;
      cap.error = item.kind === 'camera' ? cameraTrouble(err) : t('{name} could not be recorded: {reason}', { name: item.name || t('The screen'), reason: err?.message || String(err) });
      this.changed();
      return;
    }
    if (this.captures.get(key) !== cap || this.disposed) {
      if (item.kind === 'camera') closeCamera(stream); else stopAll(stream);
      return;
    }
    cap.stream = stream;
    cap.camera = item.kind === 'camera';
    cap.video = playing(stream);
    cap.opening = false;
    cap.since = performance.now();
    this.stills.delete(key);
    const track = stream.getVideoTracks()[0];
    track?.addEventListener('ended', () => {
      if (this.captures.get(key) !== cap) return;
      cap.error = t('{name} closed or stopped sharing.', { name: item.name || t('The screen') });
      this.changed();
    });
    cap.video.addEventListener('loadeddata', () => this.changed(), { once: true });
    this.changed();
  }

  close(key) {
    const cap = this.captures.get(key);
    if (!cap) return;
    this.captures.delete(key);
    letGo(cap.video);
    if (cap.stream) { if (cap.camera) closeCamera(cap.stream); else stopAll(cap.stream); }
    this.changed();
  }

  /** A source that is live but has given no frame for three seconds, opened again. */
  revive(stack, opts) {
    for (const [key, cap] of this.captures) {
      if (cap.video && !cap.video.videoWidth && cap.since && performance.now() - cap.since > 3000 && !cap.retried) {
        cap.retried = true;
        const it = stack.items.find((x) => sourceKey(x) === key);
        if (it) { this.close(key); this.open(key, it, { ...opts, opts: cap.opts }).then(() => { const c = this.captures.get(key); if (c) c.retried = true; }); }
      } else if (cap.video?.paused) {
        cap.video.play().catch(() => {});
      }
    }
  }

  /** What an item is drawn from: a playing video, a still, a picture, or null. */
  drawable(item) {
    if (item.kind === 'picture') { const im = this.images.get(item.path)?.img; return im?.naturalWidth ? im : null; }
    const key = sourceKey(item);
    const v = this.captures.get(key)?.video;
    if (v?.videoWidth) return v;
    return this.stills.get(key) || null;
  }

  /** The picture's own size, for Place and the clip, or null. */
  sizeOf(item) {
    const d = this.drawable(item);
    if (!d) return null;
    const w = d.videoWidth || d.naturalWidth || d.width;
    const h = d.videoHeight || d.naturalHeight || d.height;
    return w && h ? { w, h } : null;
  }

  /** The live stream of an item's source, for its preview in the list, or null. */
  streamOf(item) { return this.captures.get(sourceKey(item))?.stream || null; }

  /** A still of an item's source, as a data URL for its preview when stopped, or null. */
  stillOf(item) {
    const c = this.stills.get(sourceKey(item));
    if (!c) return null;
    if (!c.dataUrl) {
      const small = document.createElement('canvas');
      const k = Math.min(1, 320 / c.width);
      small.width = Math.round(c.width * k);
      small.height = Math.round(c.height * k);
      small.getContext('2d').drawImage(c, 0, 0, small.width, small.height);
      c.dataUrl = small.toDataURL('image/jpeg', 0.8);
    }
    return c.dataUrl;
  }

  stateOf(item) {
    if (item.kind === 'title') return 'ready';
    if (item.kind === 'picture') return this.images.get(item.path)?.img?.naturalWidth ? 'ready' : 'opening';
    const cap = this.captures.get(sourceKey(item));
    if (!cap) return this.stills.has(sourceKey(item)) ? 'stopped' : 'off';
    if (cap.error) return 'error';
    if (cap.opening || !cap.video?.videoWidth) return 'opening';
    return 'live';
  }

  errorOf(item) { return this.captures.get(sourceKey(item))?.error || null; }

  /** A picture file read and kept, once per path. */
  loadImage(path) {
    if (this.images.has(path)) return;
    const entry = { img: null, url: null };
    this.images.set(path, entry);
    this.shell.fs.read({ path }).then(({ bytes }) => {
      if (this.disposed) return;
      entry.url = URL.createObjectURL(new Blob([bytes]));
      const img = new Image();
      img.onload = () => { entry.img = img; this.changed(); };
      img.src = entry.url;
    }).catch(() => { entry.error = true; this.changed(); });
  }

  imageUrl(path) { return this.images.get(path)?.url || null; }

  /* ── sound ─────────────────────────────────────────────────────────── */

  audio() {
    if (!this.ctx) {
      const Ctx = window.AudioContext || window.webkitAudioContext;
      this.ctx = new Ctx();
      this.out = this.ctx.createMediaStreamDestination();
      // A source that is always there, at nothing: the output keeps a live track with no sound in it.
      const hush = this.ctx.createConstantSource();
      hush.offset.value = 0;
      hush.connect(this.out);
      hush.start();
    }
    if (this.ctx.state === 'suspended') this.ctx.resume().catch(() => {});
    return this.ctx;
  }

  /**
   * The sounds as the stack has them: each started one open, its level and
   * mute set; and what reaches the output, by the first line's Sound switch.
   */
  async syncSounds(stack) {
    const ctx = this.audio();
    this.mode = stack.sound;
    for (const [id, n] of this.sounds) {
      const s = stack.sounds.find((x) => x.id === id);
      if (!s || !s.live || n.key !== JSON.stringify([s.kind, s.ref])) this.unplug(id);
    }
    for (const s of stack.sounds) {
      if (!s.live) continue;
      let n = this.sounds.get(s.id);
      if (!n) {
        n = { key: JSON.stringify([s.kind, s.ref]), stream: null, error: null, gain: ctx.createGain(), analyser: ctx.createAnalyser(), src: null, opening: true };
        n.analyser.fftSize = 2048;
        n.gain.connect(n.analyser);
        this.sounds.set(s.id, n);
        this.openSound(s, n);
      }
      n.gain.gain.value = s.muted ? 0 : s.gain;
      this.route(n);
    }
    await this.syncMusic(stack);
    this.changed();
  }

  route(n) {
    try { n.gain.disconnect(this.out); } catch { /* not joined */ }
    if (this.mode === 'on') n.gain.connect(this.out);
  }

  async openSound(s, n) {
    let stream = null;
    let error = null;
    try {
      if (s.kind === 'mic') {
        stream = await navigator.mediaDevices.getUserMedia({ audio: s.ref ? { deviceId: { exact: s.ref } } : true, video: false });
      } else {
        // The computer's own sound comes only with getDisplayMedia, which the
        // main process answers with a screen and its loopback.
        const screen = this.list.find((x) => x.kind === 'screen') || this.list[0];
        if (!screen) throw new Error('no screen');
        const { loopback } = await this.shell.capture.choose({ id: screen.id, audio: true });
        stream = await navigator.mediaDevices.getDisplayMedia({ video: { width: { max: 64 }, height: { max: 64 }, frameRate: { max: 1 } }, audio: true });
        for (const v of stream.getVideoTracks()) v.stop();
        if (!loopback || !stream.getAudioTracks().length) { stopAll(stream); stream = null; error = t('This computer does not let its own sound be recorded.'); }
      }
    } catch (err) {
      error = s.kind === 'mic' ? (err?.name === 'NotAllowedError' ? t('The microphone was not allowed.') : t('No microphone was found.')) : t('The computer\'s own sound could not be recorded.');
    }
    if (this.sounds.get(s.id) !== n || this.disposed) { stopAll(stream); return; }
    n.opening = false;
    n.error = error;
    n.stream = stream;
    const track = stream?.getAudioTracks()[0];
    if (track) {
      n.src = this.ctx.createMediaStreamSource(new MediaStream([track]));
      n.src.connect(n.gain);
    }
    this.changed();
  }

  unplug(id) {
    const n = this.sounds.get(id);
    if (!n) return;
    this.sounds.delete(id);
    try { n.src?.disconnect(); } catch { /* already */ }
    try { n.gain.disconnect(); } catch { /* already */ }
    stopAll(n.stream);
  }

  soundState(id) {
    const n = this.sounds.get(id);
    if (!n) return 'off';
    if (n.error) return 'error';
    return n.opening ? 'opening' : 'live';
  }

  soundError(id) { return this.sounds.get(id)?.error || null; }

  /** A meter's reading, 0 to 1: a sound's id, or 'out' for what is recorded. */
  level(id) { return levelOf(this.sounds.get(id)?.analyser, this.bins); }

  /**
   * Music while Sound is on Music: the person's own file, played round, or a
   * quiet chord made here when there is none.
   */
  async syncMusic(stack) {
    const want = stack.sound === 'music' ? (stack.music || 'built-in') : null;
    if (this.music?.what === want) return;
    this.stopMusic();
    if (!want) return;
    const ctx = this.audio();
    const gain = ctx.createGain();
    gain.gain.value = 0.35;
    gain.connect(this.out);
    const m = { what: want, gain, nodes: [], el: null, url: null };
    this.music = m;
    if (want === 'built-in') {
      // A soft major chord, slowly breathing, so the stream is not silent.
      const lfo = ctx.createOscillator();
      lfo.frequency.value = 0.1;
      const depth = ctx.createGain();
      depth.gain.value = 0.04;
      lfo.connect(depth);
      depth.connect(gain.gain);
      for (const f of [220, 277.18, 329.63, 440]) {
        const o = ctx.createOscillator();
        o.type = 'sine';
        o.frequency.value = f;
        const g = ctx.createGain();
        g.gain.value = 0.06;
        o.connect(g);
        g.connect(gain);
        o.start();
        m.nodes.push(o, g);
      }
      lfo.start();
      m.nodes.push(lfo, depth);
      return;
    }
    try {
      const { bytes } = await this.shell.fs.read({ path: want });
      if (this.music !== m) return;
      m.url = URL.createObjectURL(new Blob([bytes]));
      const el = new Audio(m.url);
      el.loop = true;
      m.el = el;
      const src = ctx.createMediaElementSource(el);
      src.connect(gain);
      m.nodes.push(src);
      await el.play();
    } catch { /* no music: the output stays quiet */ }
  }

  stopMusic() {
    const m = this.music;
    this.music = null;
    if (!m) return;
    try { m.el?.pause(); } catch { /* stopped */ }
    for (const n of m.nodes) { try { n.stop?.(); } catch { /* stopped */ } try { n.disconnect(); } catch { /* already */ } }
    try { m.gain.disconnect(); } catch { /* already */ }
    if (m.url) URL.revokeObjectURL(m.url);
  }

  /** Whether anything is heard: a started sound with Sound on, or music. */
  heard() {
    if (this.mode === 'music') return true;
    if (this.mode !== 'on') return false;
    for (const n of this.sounds.values()) if (n.src && n.gain.gain.value > 0) return true;
    return false;
  }

  /** The sound recordings and streams take: the one output (the argument is the old rig's, and is not read). */
  mix() {
    this.audio();
    return this.out.stream.getAudioTracks()[0] || null;
  }

  silence() {
    const ctx = this.audio();
    const src = ctx.createConstantSource();
    src.offset.value = 0;
    const dest = ctx.createMediaStreamDestination();
    src.connect(dest);
    src.start();
    return dest.stream.getAudioTracks()[0] || null;
  }

  /** The microphones there are, and the cameras, by name. */
  static async devices() {
    try {
      const all = await navigator.mediaDevices.enumerateDevices();
      return {
        // Before the first permission the devices have no names, and nothing can be said to be missing.
        labelled: all.every((d) => Boolean(d.label)),
        cameras: all.filter((d) => d.kind === 'videoinput').map((d, i) => ({ id: d.deviceId, name: d.label || t('Camera {n}', { n: i + 1 }) })),
        mics: all.filter((d) => d.kind === 'audioinput' && d.deviceId !== 'default' && d.deviceId !== 'communications').map((d, i) => ({ id: d.deviceId, name: d.label || t('Microphone {n}', { n: i + 1 }) })),
      };
    } catch {
      return { cameras: [], mics: [], labelled: false };
    }
  }

  dispose() {
    this.disposed = true;
    for (const key of [...this.captures.keys()]) this.close(key);
    for (const id of [...this.sounds.keys()]) this.unplug(id);
    this.stopMusic();
    for (const e of this.images.values()) if (e.url) URL.revokeObjectURL(e.url);
    this.images.clear();
    try { this.ctx?.close(); } catch { /* closed */ }
    this.ctx = null;
  }
}
