// The Video window's media: how the engine reads a file, what each source
// is, and the pictures the timeline draws of it.
//
// The engine has no transport of its own (decision D4 in docs/STUDIO.md): it
// is given one here, once, that reads a path through the rutba://file
// protocol and answers a Blob. A Blob, never the protocol's URL, so the
// canvas that draws it is never tainted and captureStream keeps working.
// A source is read once and its Blob shared: the preview's decoder, an
// export's own decoders, the sound decode and the thumbnails all come from
// the same bytes.
//
// Thumbnails and waveforms are made once per source, in the background, at a
// fixed rate (a frame a second, two hundred peaks a second), and the timeline
// draws from those at whatever zoom it is at.

import { configureMediaFetch, loadVideo, loadImage, loadAudioTrack, releaseVideos } from '@rutba/studio/host';

/** A path as the window reads it, over rutba://file. */
export const fileUrl = (p) => `rutba://file/${btoa(unescape(encodeURIComponent(p))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')}`;

/** A clip's decoder key names its source and, after a #, which of its decoders. */
export const sourceOfKey = (key) => String(key).replace(/#\d+$/, '');

const VIDEO_EXT = /\.(mp4|m4v|mov|webm|mkv|ogv)$/i;
const AUDIO_EXT = /\.(mp3|m4a|aac|wav|flac|ogg|oga|opus|weba)$/i;
const IMAGE_EXT = /\.(png|jpe?g|webp|gif|bmp|avif)$/i;

/** What a file is to the timeline, by its name: 'video', 'audio', 'image', 'project' or null. */
export function mediaKind(path) {
  const p = String(path || '');
  if (/\.rvid$/i.test(p)) return 'project';
  if (VIDEO_EXT.test(p)) return 'video';
  if (AUDIO_EXT.test(p)) return 'audio';
  if (IMAGE_EXT.test(p)) return 'image';
  return null;
}

/** The extensions Insert → Media offers. */
export const MEDIA_EXTENSIONS = ['mp4', 'm4v', 'mov', 'webm', 'mkv', 'ogv', 'mp3', 'm4a', 'aac', 'wav', 'flac', 'ogg', 'opus', 'png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'avif'];

/**
 * The window's media library: Blobs by path, decoded entries by key, sound
 * by path, thumbnails and peaks by path. One per window.
 */
export class MediaLibrary {
  constructor() {
    this.blobs = new Map();
    this.videos = new Map(); // decoder key -> loadVideo entry
    this.images = new Map(); // path -> loadImage entry
    this.sounds = new Map(); // path -> AudioBuffer | null (no sound in it)
    this.thumbs = new Map(); // path -> { step, frames: [canvas] }
    this.peaks = new Map(); // path -> { rate, data: Float32Array }
    this.failed = new Map(); // path -> message
    this.listeners = new Set();
    this.pending = new Map();
    configureMediaFetch((url, opts) => this.blob(url, opts));
  }

  /** Something new to draw: a decoder ready, a thumbnail made. */
  onChange(fn) {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  changed() {
    for (const fn of this.listeners) fn();
  }

  /** A source's bytes, read once. */
  blob(url, { signal } = {}) {
    const path = sourceOfKey(url);
    if (!this.blobs.has(path)) {
      const read = fetch(fileUrl(path), { signal }).then((res) => {
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return res.blob();
      });
      read.catch(() => this.blobs.delete(path));
      this.blobs.set(path, read);
    }
    return this.blobs.get(path);
  }

  /** Run `fn` once per key, however many ask at once. */
  once(key, fn) {
    if (!this.pending.has(key)) {
      const p = Promise.resolve().then(fn);
      this.pending.set(key, p);
      p.finally(() => this.pending.delete(key)).catch(() => {});
    }
    return this.pending.get(key);
  }

  /** A decoder for a clip key (a source, or one of its extra decoders). */
  async video(key) {
    if (this.videos.has(key)) return this.videos.get(key);
    return this.once(`v:${key}`, async () => {
      const entry = await loadVideo(key);
      entry.duration = await settledDuration(entry.el);
      entry.el.addEventListener('seeked', () => this.changed());
      this.videos.set(key, entry);
      this.changed();
      return entry;
    });
  }

  async image(path) {
    if (this.images.has(path)) return this.images.get(path);
    return this.once(`i:${path}`, async () => {
      const entry = await loadImage(path);
      this.images.set(path, entry);
      this.changed();
      return entry;
    });
  }

  /** A source's sound, decoded, or null when it has none. */
  async sound(path) {
    if (this.sounds.has(path)) return this.sounds.get(path);
    return this.once(`s:${path}`, async () => {
      let buffer = null;
      try {
        buffer = await loadAudioTrack(path);
      } catch {
        buffer = null;
      }
      this.sounds.set(path, buffer);
      if (buffer) this.peaks.set(path, peaksOf(buffer));
      this.changed();
      return buffer;
    });
  }

  /**
   * What a file is, for the timeline: { url, name, kind, duration, hasAudio,
   * width, height }. Decodes it on the way, so the clip can draw at once.
   */
  async describe(path, { name, size } = {}) {
    const kind = mediaKind(path);
    if (kind === 'image') {
      const entry = await this.image(path);
      this.thumbsForImage(path, entry);
      return { url: path, name, size, kind, width: entry.width, height: entry.height };
    }
    if (kind === 'audio') {
      const buffer = await this.sound(path);
      if (!buffer) throw new Error('no sound');
      return { url: path, name, size, kind, duration: buffer.duration, hasAudio: true };
    }
    const entry = await this.video(path);
    const buffer = await this.sound(path);
    this.thumbsForVideo(path);
    return { url: path, name, size, kind: 'video', duration: entry.duration || buffer?.duration || 0, hasAudio: Boolean(buffer), width: entry.width, height: entry.height };
  }

  /** Everything a project's clips draw from, loaded; the paths that would not load, with why. */
  async loadProject(media) {
    const failures = [];
    const tasks = [
      ...media.videos.map(({ key, url }) => this.video(key).then(() => this.thumbsForVideo(url)).catch((err) => failures.push({ url, error: err.message }))),
      ...media.images.map((url) => this.image(url).then((e) => this.thumbsForImage(url, e)).catch((err) => failures.push({ url, error: err.message }))),
      ...media.sounds.map((url) => this.sound(url)),
    ];
    await Promise.all(tasks);
    const seen = new Set();
    return failures.filter((f) => (seen.has(f.url) ? false : seen.add(f.url)));
  }

  /** A still's thumbnail: the picture itself, small. */
  thumbsForImage(path, entry) {
    if (this.thumbs.has(path) || !entry?.img) return;
    const c = document.createElement('canvas');
    const h = 72;
    c.width = Math.max(1, Math.round((h * entry.width) / Math.max(1, entry.height)));
    c.height = h;
    c.getContext('2d').drawImage(entry.img, 0, 0, c.width, c.height);
    this.thumbs.set(path, { step: Infinity, frames: [c] });
    this.changed();
  }

  /** A clip's thumbnails: a frame a second (at most 120), from a decoder of their own. */
  thumbsForVideo(path) {
    if (this.thumbs.has(path)) return;
    const record = { step: 1, frames: [] };
    this.thumbs.set(path, record);
    this.once(`t:${path}`, async () => {
      const entry = await loadVideo(path);
      try {
        const el = entry.el;
        const duration = await settledDuration(el);
        const count = Math.max(1, Math.min(120, Math.ceil(duration || 1)));
        record.step = Math.max(0.25, (duration || 1) / count);
        const h = 72;
        const w = Math.max(1, Math.round((h * (entry.width || 16)) / Math.max(1, entry.height || 9)));
        for (let i = 0; i < count; i++) {
          const at = Math.min(Math.max(0, (duration || 1) - 0.05), i * record.step + 0.02);
          await seekTo(el, at);
          const c = document.createElement('canvas');
          c.width = w;
          c.height = h;
          c.getContext('2d').drawImage(el, 0, 0, w, h);
          record.frames[i] = c;
          if (i % 4 === 3 || i === count - 1) this.changed();
        }
      } finally {
        releaseVideos([entry]);
      }
    }).catch(() => {});
  }

  /** Let every decoder go. */
  dispose() {
    releaseVideos([...this.videos.values()]);
    this.videos.clear();
    this.listeners.clear();
  }
}

/** Two hundred peaks a second of a buffer's first channel (and second, when it has one). */
export function peaksOf(buffer, rate = 200) {
  const a = buffer.getChannelData(0);
  const b = buffer.numberOfChannels > 1 ? buffer.getChannelData(1) : null;
  const per = Math.max(1, Math.floor(buffer.sampleRate / rate));
  const data = new Float32Array(Math.ceil(a.length / per));
  for (let i = 0; i < data.length; i++) {
    let m = 0;
    const end = Math.min(a.length, (i + 1) * per);
    for (let j = i * per; j < end; j += 4) {
      const v = Math.abs(a[j]);
      if (v > m) m = v;
      if (b) {
        const w = Math.abs(b[j]);
        if (w > m) m = w;
      }
    }
    data[i] = m;
  }
  return { rate: buffer.sampleRate / per, data };
}

/** Seek an element and wait for the frame. */
export function seekTo(el, t) {
  return new Promise((resolve) => {
    const done = () => {
      el.removeEventListener('seeked', done);
      clearTimeout(timer);
      resolve();
    };
    const timer = setTimeout(done, 1500);
    el.addEventListener('seeked', done);
    try {
      el.currentTime = t;
    } catch {
      done();
    }
  });
}

/**
 * The length of what an element plays. A recorder's WebM states none, and
 * Chromium reports Infinity until it has been asked for the end: so it is
 * asked, and then put back to the start.
 */
export async function settledDuration(el) {
  if (Number.isFinite(el.duration) && el.duration > 0) return el.duration;
  await new Promise((resolve) => {
    const finish = () => {
      el.removeEventListener('durationchange', check);
      clearTimeout(timer);
      resolve();
    };
    const check = () => {
      if (Number.isFinite(el.duration) && el.duration > 0) finish();
    };
    const timer = setTimeout(finish, 4000);
    el.addEventListener('durationchange', check);
    try {
      el.currentTime = 1e7;
    } catch {
      finish();
    }
  });
  const d = Number.isFinite(el.duration) ? el.duration : 0;
  await seekTo(el, 0);
  return d;
}
