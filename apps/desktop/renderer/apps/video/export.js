// Export: the project written as a video file, with its sound.
//
// Where this machine has a WebCodecs encoder (H.264, or VP9 where H.264 is
// missing), the export is our own: each frame is painted by the engine as
// soon as the encoder will take it, every clip's frame decoded from its file
// for that instant (@rutba/media/frames), the sound mixed offline, and the
// samples written into an MP4 of our own making (@rutba/media/mp4-write)
// streamed to a scratch file, moved to where it was asked to go when it is
// done. It runs as fast as the machine paints and encodes, on the graphics
// hardware where there is an encoder there, and says how fast.
//
// Where there is none, the engine's renderVideo plays the plan in real time
// into a canvas stream and a MediaRecorder, as before: MP4 where the machine
// records it, WebM otherwise, taking as long as the project runs.

import React, { useEffect, useState } from 'react';
import { Button, Dialog, Field, Select, Progress, formatBytes, t } from '@rutba/office-ui';
import { timecode } from '@rutba/media/timeline';
import { chooseVideo, chooseAudio, encodeMovie } from '@rutba/media/export';
import { FrameReader } from '@rutba/media/frames';
import { buildPlan, paintFrame, renderVideo, soundLayers, clipFromSoundLayer, pickMimeType, extensionFor, loadVideo, releaseVideos, videoTimesAt, mixAudioOffline } from '@rutba/studio/host';
import { planPatches, projectMedia, projectDuration, cutToRange } from '@rutba/studio/project';
import { EXPORT_PRESETS, presetById, presetFrame, presetBytes, presetWarnings, presetDoc } from '@rutba/studio/presets';

/** Each destination by name, as the Export box offers them. */
export const PRESET_NAMES = () => ({
  project: t('This project as it is'),
  youtube: t('YouTube, 16:9 at 1080p'),
  'social-wide': t('LinkedIn and X, 16:9'),
  'social-square': t('LinkedIn and X, square'),
  vertical: t('Reels, Shorts and TikTok, 9:16'),
  email: t('Email, a small MP4'),
});

/** A short word for a destination, put after the project's name in the file's. */
export const PRESET_TAGS = () => ({
  youtube: t('YouTube'),
  'social-wide': t('LinkedIn and X'),
  'social-square': t('Square'),
  vertical: t('Vertical'),
  email: t('Email'),
});

/** What MediaRecorder writes here, when there is no encoder of our own: { mime, ext, label, fast: false }. */
function recorderFormat() {
  const mime = pickMimeType(true);
  if (!mime) return null;
  const ext = extensionFor(mime);
  return { mime, ext, fast: false, label: ext === 'mp4' ? t('MP4 video (H.264 and AAC)') : t('WebM video (VP9 and Opus)') };
}

let known;
let asking = null;

/**
 * What this machine will write: { mime, ext, label, fast } or null. Asked of
 * the encoders once (they answer asynchronously); `knownFormat` gives the
 * answer once there is one.
 */
export function exportFormat() {
  if (!asking) {
    asking = (async () => {
      const video = await chooseVideo({ width: 1920, height: 1080, fps: 30 }).catch(() => null);
      if (!video) return recorderFormat();
      const audio = await chooseAudio({}).catch(() => null);
      const v = video.codec.startsWith('avc1') ? 'H.264' : 'VP9';
      const a = audio ? (audio.codec === 'opus' ? 'Opus' : 'AAC') : null;
      return {
        mime: 'video/mp4',
        ext: 'mp4',
        fast: true,
        hardware: video.hardware,
        label: a ? t('MP4 video ({video} and {audio})', { video: v, audio: a }) : t('MP4 video ({video})', { video: v }),
      };
    })().then((f) => { known = f; return f; }, () => { known = recorderFormat(); return known; });
  }
  return asking;
}

/** The format already found, or undefined while the encoders are still being asked. */
export const knownFormat = () => known;

/** A clip's frames by seeking a <video>: for a file the decoder will not take. */
class SeekFrames {
  constructor(entry) {
    this.entry = entry;
    this.width = entry.width;
    this.height = entry.height;
  }

  async frameAt(t) {
    const el = this.entry.el;
    const want = Math.max(0, Math.min(t, (el.duration || t) - 0.001));
    if (!this.ready || Math.abs(el.currentTime - want) > 0.0005) {
      await new Promise((resolve) => {
        const done = () => { el.removeEventListener('seeked', done); clearTimeout(timer); resolve(); };
        const timer = setTimeout(done, 3000);
        el.addEventListener('seeked', done);
        try { el.currentTime = want; } catch { done(); }
      });
      this.ready = true;
    }
    return el;
  }

  close() {
    releaseVideos([this.entry]);
  }
}

/** The project's stretch, laid out for where it is going, and what it needs loaded. */
async function prepare(doc, media, { range, preset, relayout }) {
  const source = presetDoc(range ? cutToRange(doc, range.in, range.out) : doc, preset, { relayout });
  const want = projectMedia(source);
  const assets = {};
  for (const url of want.images) {
    const entry = media.images.get(url) || (await media.image(url).catch(() => null));
    if (entry) assets[url] = entry;
  }
  return { source, want, assets };
}

async function soundClips(plan, media) {
  const clips = [];
  for (const layer of soundLayers(plan)) {
    const buffer = await media.sound(layer.url).catch(() => null);
    if (buffer) clips.push(clipFromSoundLayer(layer, buffer, plan, { volume: 1, fadeIn: 0, fadeOut: 0 }));
  }
  return clips;
}

/**
 * Our own export: painted, encoded and written as an MP4 to `target`
 * through the shell, faster than the project plays. Answers { size,
 * duration, seconds, frames, video: { codec, hardware }, audio, fallbacks }.
 */
async function exportMp4(doc, media, shell, target, { quality, range, preset, relayout, onProgress, signal, hardware = true }) {
  const { source, want, assets } = await prepare(doc, media, { range, preset, relayout });
  const readers = {};
  const videos = {};
  let fallbacks = 0;
  let temp = null;
  let kept = false;
  try {
    const bytesByUrl = new Map();
    for (const { key, url } of want.videos) {
      try {
        if (!bytesByUrl.has(url)) bytesByUrl.set(url, new Uint8Array(await (await media.blob(url)).arrayBuffer()));
        let reader = await FrameReader.open(bytesByUrl.get(url));
        if (!reader) {
          reader = new SeekFrames(await loadVideo(key));
          fallbacks++;
        }
        readers[key] = reader;
        videos[key] = { frame: null, width: reader.width, height: reader.height, url };
      } catch {
        /* a source that will not load is a clip that draws nothing */
      }
    }
    bytesByUrl.clear();
    const canvas = document.createElement('canvas');
    const plan = buildPlan({ canvas, images: [], options: { ...source.options, quality }, layerPatches: planPatches(source), videos, assets });
    const ctx = canvas.getContext('2d');
    const clips = await soundClips(plan, media);
    const sound = clips.length ? await mixAudioOffline({ clips }, plan) : null;
    if (signal?.aborted) throw new DOMException('Export cancelled', 'AbortError');

    temp = (await shell.fs.temp({ ext: '.mp4' })).path;
    const sink = {
      write: (bytes) => shell.fs.append({ path: temp, bytes }),
      patch: (offset, bytes) => shell.fs.patch({ path: temp, offset, bytes }),
    };
    let decoding = 0;
    const paint = async (i, at) => {
      const t = performance.now();
      const times = videoTimesAt(plan, at);
      for (const [key, entry] of Object.entries(videos)) {
        const s = times.get(key);
        entry.frame = s === undefined ? null : await readers[key].frameAt(s);
      }
      decoding += performance.now() - t;
      paintFrame(ctx, plan, at);
      return canvas;
    };
    const audioBitrate = Number(plan.opts?.audioBitrate) > 0 ? Math.round(Number(plan.opts.audioBitrate)) : 128000;
    const out = await encodeMovie({ sink, width: plan.W, height: plan.H, fps: plan.fps, duration: plan.duration, bitrate: plan.bitrate, audioBitrate, paint, sound, onProgress, signal, hardware });
    await shell.fs.keepTemp({ path: temp, to: target });
    kept = true;
    // How each source was read: decoded frame by frame, or sought in a <video> where the decoder would not take it.
    const sources = Object.entries(readers).map(([key, r]) => ({ key: key.split(/[\\/]/).pop(), how: r instanceof SeekFrames ? 'seek' : 'frames', decoded: r.decoded ?? null, restarts: r.restarts ?? null, failed: r.failed ? String(r.failed.message || r.failed) : null }));
    return { ...out, spent: { ...out.spent, decode: Math.round(decoding) }, fallbacks, sources };
  } finally {
    for (const r of Object.values(readers)) {
      try { r.close(); } catch { /* already closed */ }
    }
    if (temp && !kept) shell.fs.dropTemp({ path: temp }).catch(() => {});
  }
}

/** The real-time export through MediaRecorder: bytes, for the window to write. */
async function recordProject(doc, media, { quality, range, preset, relayout, onProgress, signal }) {
  const { source, want, assets } = await prepare(doc, media, { range, preset, relayout });
  const canvas = document.createElement('canvas');
  // A canvas that belongs to the page, off to one side, so the engine's
  // capture is of a live document.
  canvas.className = 'vd-export-canvas';
  document.body.appendChild(canvas);
  const videos = {};
  try {
    for (const { key } of want.videos) {
      try {
        videos[key] = await loadVideo(key);
      } catch {
        /* a source that will not load is a clip that draws nothing */
      }
    }
    const plan = buildPlan({ canvas, images: [], options: { ...source.options, quality }, layerPatches: planPatches(source), videos, assets });
    const clips = await soundClips(plan, media);
    const out = await renderVideo({ canvas, plan, audio: clips.length ? { clips } : null, onProgress, signal });
    return { bytes: new Uint8Array(await out.blob.arrayBuffer()), duration: out.duration };
  } finally {
    releaseVideos(Object.values(videos));
    canvas.remove();
  }
}

/**
 * Export a project to `target`: our own MP4 where `format.fast`, a
 * recording otherwise. `media` is the window's library (its Blobs are
 * reused); `onProgress(0..1)`; `signal` cancels. Answers { size, seconds,
 * fast, hardware, codec }.
 */
export async function exportProject(doc, media, shell, target, { format, ...options }) {
  const started = performance.now();
  if (format?.fast) {
    let out;
    try {
      out = await exportMp4(doc, media, shell, target, options);
    } catch (err) {
      // A graphics encoder that gave up part way: once more on the processor.
      if (err?.name === 'AbortError' || options.signal?.aborted) throw err;
      out = await exportMp4(doc, media, shell, target, { ...options, hardware: false });
    }
    return { size: out.size, seconds: (performance.now() - started) / 1000, duration: out.duration, fast: true, hardware: out.video.hardware, codec: out.video.codec, audio: out.audio?.codec || null, fallbacks: out.fallbacks, spent: out.spent, sources: out.sources };
  }
  const out = await recordProject(doc, media, options);
  await shell.fs.write({ path: target, bytes: out.bytes });
  return { size: out.bytes.length, seconds: (performance.now() - started) / 1000, duration: out.duration, fast: false };
}

/** The format, once the encoders have answered. */
function useExportFormat() {
  const [format, setFormat] = useState(knownFormat);
  useEffect(() => {
    let live = true;
    if (format === undefined) exportFormat().then((f) => { if (live) setFormat(f); });
    return () => { live = false; };
  }, [format]);
  return format;
}

/** File → Export: the format this machine writes, the quality, the whole project or the In to Out range. */
export function ExportDialog({ doc, onClose, onExport, initialPreset = 'project' }) {
  const format = useExportFormat();
  const [quality, setQuality] = useState('high');
  const [which, setWhich] = useState(doc.range ? 'range' : 'all');
  const [preset, setPreset] = useState(presetById(initialPreset).id);
  const length = which === 'range' && doc.range ? doc.range.out - doc.range.in : projectDuration(doc);
  const names = PRESET_NAMES();
  const frame = presetFrame(preset, doc.options?.aspect);
  const p = presetById(preset);
  const bytes = presetBytes(p, length);
  const shape = { landscape: '16:9', square: '1:1', vertical: '9:16' }[frame.aspect];
  const warnings = presetWarnings(p, length);
  return (
    <Dialog
      title={t('Export video')}
      onClose={onClose}
      width={420}
      className="vd-export-dialog"
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary className="vd-export-go" label={t('Export')} disabled={!format || length <= 0} onClick={() => onExport({ quality, range: which === 'range' ? doc.range : null, format, preset })} />
        </>
      }
    >
      <Field label={t('Made for')}>
        <Select value={preset} className="rw-select vd-export-preset" onChange={(e) => setPreset(e.target.value)}>
          {EXPORT_PRESETS.map((x) => <option key={x.id} value={x.id}>{names[x.id]}</option>)}
        </Select>
        <div className="vd-export-frame" data-width={frame.width} data-height={frame.height}>
          <span className={`vd-export-shape s-${frame.aspect}`} aria-hidden="true" />
          <span>
            {t('{width} × {height}, {shape}, {fps} frames a second', { width: frame.width, height: frame.height, shape, fps: p.fps })}
            {bytes ? <><br />{t('About {size}', { size: formatBytes(bytes) })}</> : null}
          </span>
        </div>
        {warnings.includes('long') ? <div className="vd-export-warn">{t('Longer than {minutes} minutes: the place it is going may not take it.', { minutes: Math.round(p.maxSeconds / 60) })}</div> : null}
        {warnings.includes('big') ? <div className="vd-export-warn">{t('Larger than most email will carry. Export the In to Out range, or a shorter project.')}</div> : null}
      </Field>
      <Field label={t('Format')}>
        <div className="vd-export-format" data-fast={format?.fast ? '1' : '0'}>
          {format === undefined ? t('Finding this computer\'s encoders…') : format ? format.label : t('This system cannot record video.')}
        </div>
      </Field>
      {preset === 'project' ? (
        <Field label={t('Quality')}>
          <Select value={quality} onChange={(e) => setQuality(e.target.value)}>
            <option value="high">{t('High')}</option>
            <option value="medium">{t('Medium')}</option>
            <option value="low">{t('Low (smaller file)')}</option>
          </Select>
        </Field>
      ) : null}
      <Field label={t('What to export')}>
        <Select value={which} onChange={(e) => setWhich(e.target.value)}>
          <option value="all">{t('The whole project')}</option>
          <option value="range" disabled={!doc.range}>{t('From the In point to the Out point')}</option>
        </Select>
      </Field>
      <div className="rw-hint">
        {format?.fast
          ? (format.hardware
            ? t('Export writes the video as fast as this computer can encode it, on its graphics hardware: usually far less than the {length} the project runs.', { length: timecode(length) })
            : t('Export writes the video as fast as this computer can encode it: usually less than the {length} the project runs.', { length: timecode(length) }))
          : t('Export plays the project once as it records it, so it takes {length}.', { length: timecode(length) })}
      </div>
    </Dialog>
  );
}

/** The progress of an export, with Cancel: how far, and for our own export how fast and how long is left. */
export function ExportProgress({ progress, total, fast, since = null, from = 0, onCancel }) {
  const elapsed = since != null ? (performance.now() - since) / 1000 : 0;
  const done = progress * total;
  const speed = fast && elapsed > 1 && progress > from ? ((progress - from) * total) / elapsed : 0;
  const left = speed > 0 ? Math.max(0, (total - done) / speed) : null;
  return (
    <Dialog title={t('Exporting')} onClose={onCancel} width={380} className="vd-export-busy" actions={<Button label={t('Cancel')} onClick={onCancel} />}>
      <div className="vd-export-line">
        {fast ? t('Encoding {done} of {total}', { done: timecode(done), total: timecode(total) }) : t('Recording {done} of {total}', { done: timecode(done), total: timecode(total) })}
      </div>
      <Progress value={progress} />
      {fast ? (
        <div className="vd-export-speed" data-speed={speed ? speed.toFixed(2) : ''}>
          {speed ? t('{speed}× as fast as it plays, about {left} left', { speed: speed.toFixed(1), left: timecode(left) }) : t('Starting the encoder…')}
        </div>
      ) : null}
    </Dialog>
  );
}
