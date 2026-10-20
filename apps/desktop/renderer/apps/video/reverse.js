// Reverse: a copy of a clip's stretch of its source written backward, as an
// MP4 of our own (@rutba/media/export), its sound reversed with it. A
// second's frames at a time are decoded forward, held as pictures, and
// handed to the encoder last first, so the decoder never starts again for
// every frame and no more than a second is held at once.

import { encodeMovie } from '@rutba/media/export';
import { FrameReader } from '@rutba/media/frames';
import { loadVideo } from '@rutba/studio/host';
import { lengthOf, speedOf } from '@rutba/studio/project';
import { SeekFrames } from './export.js';

const WINDOW = 1;

/** A source's samples from `from` for `seconds`, last first, as an AudioBuffer. */
export function reversedSound(buffer, from, seconds) {
  const rate = buffer.sampleRate;
  const start = Math.max(0, Math.round(from * rate));
  const length = Math.max(1, Math.min(buffer.length - start, Math.round(seconds * rate)));
  const out = new AudioBuffer({ length, sampleRate: rate, numberOfChannels: buffer.numberOfChannels });
  for (let c = 0; c < buffer.numberOfChannels; c++) {
    const src = buffer.getChannelData(c).subarray(start, start + length);
    const dst = out.getChannelData(c);
    for (let i = 0; i < length; i++) dst[i] = src[length - 1 - i];
  }
  return out;
}

/**
 * Write `clip`'s stretch of its source backward to `target`. `media` is the
 * window's library; `onProgress(0..1)`; `signal` cancels. Answers { size,
 * duration, width, height, hasAudio }.
 */
export async function writeReversed(clip, media, shell, target, { fps = 30, onProgress, signal } = {}) {
  const from = Math.max(0, clip.offset || 0);
  const span = lengthOf(clip) * speedOf(clip);
  const bytes = new Uint8Array(await (await media.blob(clip.url)).arrayBuffer());
  const reader = (await FrameReader.open(bytes).catch(() => null)) || new SeekFrames(await loadVideo(clip.url));
  const width = reader.width || 1280;
  const height = reader.height || 720;
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  const frames = Math.max(1, Math.round(span * fps));
  const per = Math.max(1, Math.round(WINDOW * fps));
  // Frame i of the copy is the source at from + span - (i + 0.5) / fps.
  const sourceAt = (i) => Math.max(from, from + span - (i + 0.5) / fps);
  let held = new Map();
  const drop = () => { for (const b of held.values()) b.close?.(); held = new Map(); };
  const fill = async (i) => {
    drop();
    const last = Math.min(frames - 1, i + per - 1);
    // Forward through the source: the window's earliest frame first.
    for (let k = last; k >= i; k--) {
      const frame = await reader.frameAt(sourceAt(k));
      if (frame) held.set(k, await createImageBitmap(frame));
    }
  };
  const sound = await media.sound(clip.url).catch(() => null);
  let temp = null;
  let kept = false;
  try {
    temp = (await shell.fs.temp({ ext: '.mp4' })).path;
    const sink = {
      write: (b) => shell.fs.append({ path: temp, bytes: b }),
      patch: (offset, b) => shell.fs.patch({ path: temp, offset, bytes: b }),
    };
    const paint = async (i) => {
      if (!held.has(i)) await fill(i);
      const picture = held.get(i);
      ctx.fillStyle = '#000';
      ctx.fillRect(0, 0, width, height);
      if (picture) ctx.drawImage(picture, 0, 0, width, height);
      return canvas;
    };
    const out = await encodeMovie({
      sink, width, height, fps, duration: frames / fps, bitrate: Math.max(4e6, Math.round(width * height * fps * 0.12)),
      paint, sound: sound ? reversedSound(sound, from, span) : null, onProgress, signal,
    });
    await shell.fs.keepTemp({ path: temp, to: target });
    kept = true;
    return { size: out.size, duration: frames / fps, width, height, hasAudio: Boolean(sound) };
  } finally {
    drop();
    try { reader.close(); } catch { /* already closed */ }
    if (temp && !kept) shell.fs.dropTemp({ path: temp }).catch(() => {});
  }
}
