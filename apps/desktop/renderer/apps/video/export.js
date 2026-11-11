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
import { Button, Dialog, Field, Input, Select, Progress, t } from '@rutba/office-ui';
import { timecode } from '@rutba/media/timeline';
import { chooseVideo, chooseAudio, encodeMovie } from '@rutba/media/export';
import { FrameReader } from '@rutba/media/frames';
import { buildPlan, paintFrame, renderVideo, soundLayers, clipFromSoundLayer, pickMimeType, extensionFor, loadVideo, releaseVideos, videoTimesAt, mixAudioOffline } from '@rutba/studio/host';
import { planPatches, projectMedia, projectDuration, cutToRange } from '@rutba/studio/project';
import { EXPORT_PRESETS, FIT_SIZES, MEGABYTE, bitrateWords, encodeToFit, fitPlan, megabytes, parseMegabytes, presetById, presetFrame, presetBytes, presetWarnings, presetDoc, projectFrame } from '@rutba/studio/presets';

/** Each destination by name, as the Export box's cards show them. */
export const PRESET_NAMES = () => ({
  project: t('This project\'s settings'),
  youtube: t('YouTube'),
  compact: t('Compact'),
  small: t('Small'),
  'social-wide': t('Social, wide'),
  'social-square': t('Social, square'),
  vertical: t('Vertical'),
  email: t('Email'),
});

/** Where each destination goes, the card's tip. */
export const PRESET_TIPS = () => ({
  project: t('This project as it is'),
  youtube: t('YouTube, 16:9 at 1080p'),
  compact: t('Compact: 1080p at about 4 Mb/s, for a smaller file that still looks right'),
  small: t('Small: 720p at about 2 Mb/s, the size to send'),
  'social-wide': t('LinkedIn and X, 16:9'),
  'social-square': t('LinkedIn and X, square'),
  vertical: t('Reels, Shorts and TikTok, 9:16'),
  email: t('Email, a small MP4'),
});

/** A short word for a destination, put after the project's name in the file's. */
export const PRESET_TAGS = () => ({
  youtube: t('YouTube'),
  compact: t('Compact'),
  small: t('Small'),
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
export class SeekFrames {
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
async function prepare(doc, media, { range, preset, relayout, fit }) {
  const source = presetDoc(range ? cutToRange(doc, range.in, range.out) : doc, preset, { relayout, fit });
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
 * through the shell, faster than the project plays. With `fit` (what
 * fitPlan answered) the file written is read back and, if it is over the
 * limit, encoded again lower, `onAgain` saying so first; the sources are
 * read again from their start, so the second pass costs one more encode and
 * no more. Answers { size, duration, seconds, frames, video: { codec,
 * hardware }, audio, fallbacks, fit }.
 */
async function exportMp4(doc, media, shell, target, { quality, range, preset, relayout, fit = null, onAgain = null, onProgress, onStall = null, signal, hardware = true }) {
  const { source, want, assets } = await prepare(doc, media, { range, preset, relayout, fit });
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
        let reader = await FrameReader.open(bytesByUrl.get(url), { hardware });
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
    // One pass to a fresh scratch file: answers the movie, its size as read back from the file written.
    const pass = async (bitrate) => {
      if (temp) await shell.fs.dropTemp({ path: temp }).catch(() => {});
      temp = (await shell.fs.temp({ ext: '.mp4' })).path;
      const file = temp;
      const sink = {
        write: (bytes) => shell.fs.append({ path: file, bytes }),
        patch: (offset, bytes) => shell.fs.patch({ path: file, offset, bytes }),
      };
      const movie = await encodeMovie({ sink, width: plan.W, height: plan.H, fps: plan.fps, duration: plan.duration, bitrate, audioBitrate, paint, sound, onProgress, onStall, signal, hardware });
      const written = await shell.fs.stat({ path: file }).catch(() => null);
      return { ...movie, size: written?.size ?? movie.size };
    };
    let fitted = null;
    let out;
    if (fit) {
      // The checks set this to make the first pass come out over, to prove the second.
      const first = Math.round(fit.video * (Number(globalThis.__rutbaCheckFitOver) || 1));
      fitted = await encodeToFit({ fit, encode: pass, onAgain, first });
      out = fitted.out;
    } else {
      out = await pass(plan.bitrate);
    }
    await shell.fs.keepTemp({ path: temp, to: target });
    kept = true;
    // How each source was read: decoded frame by frame, or sought in a <video> where the decoder would not take it.
    const sources = Object.entries(readers).map(([key, r]) => ({ key: key.split(/[\\/]/).pop(), how: r instanceof SeekFrames ? 'seek' : 'frames', decoded: r.decoded ?? null, restarts: r.restarts ?? null, failed: r.failed ? String(r.failed.message || r.failed) : null }));
    return { ...out, spent: { ...out.spent, decode: Math.round(decoding) }, fallbacks, sources, fit: fitted ? { limit: fit.limit, tries: fitted.tries, over: fitted.over } : null };
  } finally {
    for (const r of Object.values(readers)) {
      try { r.close(); } catch { /* already closed */ }
    }
    if (temp && !kept) shell.fs.dropTemp({ path: temp }).catch(() => {});
  }
}

/** The real-time export through MediaRecorder: bytes, for the window to write. */
async function recordProject(doc, media, { quality, range, preset, relayout, fit = null, onProgress, signal }) {
  const { source, want, assets } = await prepare(doc, media, { range, preset, relayout, fit });
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
      // A graphics encoder that gave up part way, or stalled (sent nothing for
      // fifteen seconds): once more on the processor, and the box says why.
      if (err?.name === 'AbortError' || options.signal?.aborted) throw err;
      options.onSoftware?.({ stalled: Boolean(err?.stalled), reason: err?.message || String(err) });
      out = await exportMp4(doc, media, shell, target, { ...options, hardware: false });
    }
    return { size: out.size, seconds: (performance.now() - started) / 1000, duration: out.duration, fast: true, hardware: out.video.hardware, codec: out.video.codec, audio: out.audio?.codec || null, fallbacks: out.fallbacks, spent: out.spent, sources: out.sources, fit: out.fit };
  }
  // A recording plays through in real time, so a second go takes as long as the first.
  let fitted = null;
  let out;
  if (options.fit) {
    fitted = await encodeToFit({
      fit: options.fit,
      onAgain: options.onAgain,
      encode: async (video) => {
        const made = await recordProject(doc, media, { ...options, fit: { ...options.fit, video } });
        return { ...made, size: made.bytes.length };
      },
    });
    out = fitted.out;
  } else {
    out = await recordProject(doc, media, options);
  }
  await shell.fs.write({ path: target, bytes: out.bytes });
  return { size: out.bytes.length, seconds: (performance.now() - started) / 1000, duration: out.duration, fast: false, fit: fitted ? { limit: options.fit.limit, tries: fitted.tries, over: fitted.over } : null };
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

const SHAPE_WORDS = () => ({ landscape: t('wide'), square: t('square'), vertical: t('tall') });
const QUALITY_WORDS = () => ({ high: t('High'), medium: t('Medium'), low: t('Low') });

/** The warning triangle beside a picture too thin to watch. */
const WarnMark = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" aria-hidden="true">
    <path d="M12 3l10 18H2z M12 10v5M12 17.5v.5" fill="none" stroke="currentColor" strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
  </svg>
);

/** The little mark on what is new in this release. */
const NewMark = () => <i className="vd-export-new">{t('NEW')}</i>;

/**
 * The destinations as cards: each with its size for this project, its frame
 * and what it costs. A radio group: the arrows move through it, one Tab
 * stop, and the choice is on `data-value` for whoever reads it.
 */
function PresetCards({ doc, preset, onChoose, length, quality, sound, fit }) {
  const names = PRESET_NAMES();
  const tips = PRESET_TIPS();
  const shapes = SHAPE_WORDS();
  const ownFrame = projectFrame(doc);
  const ownFps = Number(doc.options?.fps) || 30;
  const move = (e, at) => {
    const step = { ArrowRight: 1, ArrowDown: 2, ArrowLeft: -1, ArrowUp: -2 }[e.key];
    if (!step) return;
    e.preventDefault();
    const next = EXPORT_PRESETS[(at + step + EXPORT_PRESETS.length) % EXPORT_PRESETS.length];
    onChoose(next.id);
    e.currentTarget.parentElement.querySelector(`[data-preset="${next.id}"]`)?.focus();
  };
  return (
    <div className="vd-export-cards vd-export-preset" role="radiogroup" aria-label={t('Made for')} data-value={preset}>
      {EXPORT_PRESETS.map((x, i) => {
        const own = x.id === 'project';
        const lowered = x.id === preset && fit?.short ? fit.frame : null;
        const frame = lowered || (own ? ownFrame : presetFrame(x, doc.options?.aspect));
        const fps = own ? ownFps : x.fps;
        const bytes = lowered
          ? fit.estimate
          : presetBytes(x, length, own ? { quality, frame, fps, sound } : { sound });
        const on = x.id === preset;
        const fresh = x.id === 'compact' || x.id === 'small';
        const line = t('{width} × {height}, {shape}, {fps} frames a second', { width: frame.width, height: frame.height, shape: shapes[frame.aspect], fps });
        const extra = fresh && !lowered ? t('about {rate}', { rate: bitrateWords(x.bitrate) }) : own ? QUALITY_WORDS()[quality] : null;
        return (
          <button
            key={x.id}
            type="button"
            role="radio"
            aria-checked={on}
            tabIndex={on ? 0 : -1}
            className="vd-export-card"
            data-preset={x.id}
            data-tip={tips[x.id]}
            onClick={() => onChoose(x.id)}
            onKeyDown={(e) => move(e, i)}
          >
            <span className="vd-export-card-top">
              <span className="vd-export-card-name">{names[x.id]}{fresh ? <NewMark /> : null}</span>
              <span className="vd-export-card-size" data-bytes={bytes || 0}>{bytes ? t('About {size}', { size: megabytes(bytes) }) : ''}</span>
            </span>
            <small className={`vd-export-card-detail${on ? ' vd-export-frame' : ''}`} data-width={frame.width} data-height={frame.height}>
              {extra ? `${line} · ${extra}` : line}
            </small>
          </button>
        );
      })}
    </div>
  );
}

/** Fit under a size: the chips, a size typed, what it comes to, and the warning when the picture would be too thin. */
function FitBox({ on, setOn, size, setSize, typed, setTyped, fit, onUseShort, onKeepFrame, length }) {
  const typing = typed.trim() !== '';
  const seconds = Math.max(0, length);
  return (
    <div className="vd-export-fit">
      <label className="vd-export-fit-on">
        <input type="checkbox" className="vd-export-fit-box" checked={on} onChange={(e) => setOn(e.target.checked)} />
        {t('Fit under a size')}<NewMark />
      </label>
      {on ? (
        <>
          <div className="vd-export-chips" role="group" aria-label={t('Fit under')}>
            {FIT_SIZES.map((mb) => (
              <button
                key={mb}
                type="button"
                className="vd-export-chip"
                data-mb={mb}
                aria-pressed={!typing && size === mb}
                onClick={() => { setSize(mb); setTyped(''); }}
              >
                {t('{n} MB', { n: mb })}
              </button>
            ))}
            <label className="vd-export-typed">
              {t('or')}
              <Input className="rw-input vd-export-size" aria-label={t('Size in MB')} placeholder={t('MB')} inputMode="decimal" value={typed} onChange={(e) => setTyped(e.target.value)} />
            </label>
          </div>
          {!fit ? (
            <div className="vd-export-aim">{typing ? t('Type a size in MB, like 50.') : seconds <= 0 ? t('There is nothing to export yet.') : null}</div>
          ) : fit.tooSmall ? (
            <div className="vd-export-too-low vd-export-too-small">
              <WarnMark />
              <span>{t('{size} leaves no room for the picture over {length}. Ask for a larger size, or export a shorter stretch.', { size: megabytes(fit.limit), length: timecode(seconds) })}</span>
            </div>
          ) : (
            <>
              <div className="vd-export-aim" data-video={fit.video} data-audio={fit.audio} data-capped={fit.capped ? '1' : '0'}>
                {fit.capped
                  ? t('This choice already comes to about {size}, under {limit}, so it stays as it is. The written file is still checked.', { size: megabytes(fit.estimate), limit: megabytes(fit.limit) })
                  : t('Aims for about {size}, at {rate}, so the file comes in under {limit}. The written file is checked; if it is still over, it is made again a little lower and you are told.', { size: megabytes(fit.estimate), rate: bitrateWords(fit.video), limit: megabytes(fit.limit) })}
                {fit.short ? <> <button type="button" className="vd-export-link" onClick={onKeepFrame}>{t('Keep the full size')}</button></> : null}
              </div>
              {fit.poor ? (
                <div className="vd-export-too-low" role="alert">
                  <WarnMark />
                  <span>
                    {t('At {limit} this would get {rate}, too little for a clear picture at {lines}.', { limit: megabytes(fit.limit), rate: bitrateWords(fit.video), lines: Math.min(fit.frame.width, fit.frame.height) })}
                    {' '}
                    {fit.suggest
                      ? <button type="button" className="vd-export-link vd-export-use-short" onClick={() => onUseShort(fit.suggest)}>{t('Use {lines} instead', { lines: fit.suggest })}</button>
                      : t('Ask for a larger size, or export a shorter stretch.')}
                  </span>
                </div>
              ) : null}
            </>
          )}
        </>
      ) : null}
    </div>
  );
}

/** File → Export: where it is going, each choice's size, Fit under a size, the format and how much of the project. */
export function ExportDialog({ doc, onClose, onExport, initialPreset = 'project' }) {
  const format = useExportFormat();
  const [quality, setQuality] = useState('high');
  const [which, setWhich] = useState(doc.range ? 'range' : 'all');
  const [preset, setPreset] = useState(presetById(initialPreset).id);
  const [fitOn, setFitOn] = useState(false);
  const [fitSize, setFitSize] = useState(FIT_SIZES[0]);
  const [typed, setTyped] = useState('');
  const [short, setShort] = useState(0);
  const length = which === 'range' && doc.range ? doc.range.out - doc.range.in : projectDuration(doc);
  const p = presetById(preset);
  const own = p.id === 'project';
  const sound = doc.patches.some((x) => x.type === 'sound' && x.visible !== false);
  const limit = typed.trim() !== '' ? parseMegabytes(typed) : fitSize * MEGABYTE;
  const fit = fitOn && limit > 0 && length > 0
    ? fitPlan(p, length, { bytes: limit, quality, frame: own ? projectFrame(doc) : presetFrame(p, doc.options?.aspect), fps: own ? Number(doc.options?.fps) || 30 : p.fps, aspect: doc.options?.aspect || 'landscape', sound, short })
    : null;
  const choose = (id) => { setPreset(id); setShort(0); };
  const warnings = presetWarnings(p, length);
  const cannot = !format || length <= 0 || (fitOn && (!fit || fit.tooSmall));
  return (
    <Dialog
      title={t('Export video')}
      onClose={onClose}
      width={720}
      className="vd-export-dialog"
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button
            primary
            className="vd-export-go"
            label={t('Export')}
            disabled={cannot}
            onClick={() => onExport({ quality, range: which === 'range' ? doc.range : null, format, preset, fit: fit ? { limit: fit.limit, video: fit.video, audio: fit.audio, short: fit.short, frame: fit.frame } : null })}
          />
        </>
      }
    >
      <div className="vd-export-section">
        <span className="vd-export-label">{t('Made for')}</span>
        <PresetCards doc={doc} preset={preset} onChoose={choose} length={length} quality={quality} sound={sound} fit={fit} />
        {warnings.includes('long') ? <div className="vd-export-warn">{t('Longer than {minutes} minutes: the place it is going may not take it.', { minutes: Math.round(p.maxSeconds / 60) })}</div> : null}
        {warnings.includes('big') ? <div className="vd-export-warn">{t('Larger than most email will carry. Export the In to Out range, or a shorter project.')}</div> : null}
      </div>
      <FitBox on={fitOn} setOn={setFitOn} size={fitSize} setSize={setFitSize} typed={typed} setTyped={setTyped} fit={fit} onUseShort={setShort} onKeepFrame={() => setShort(0)} length={length} />
      <div className={`vd-export-row${own ? ' three' : ''}`}>
        {own ? (
          <Field label={t('Quality')}>
            <Select value={quality} className="rw-select vd-export-quality" onChange={(e) => setQuality(e.target.value)}>
              <option value="high">{t('High')}</option>
              <option value="medium">{t('Medium')}</option>
              <option value="low">{t('Low (smaller file)')}</option>
            </Select>
          </Field>
        ) : null}
        <Field label={t('Format')}>
          <div className="vd-export-format" data-fast={format?.fast ? '1' : '0'}>
            {format === undefined ? t('Finding this computer\'s encoders…') : format ? format.label : t('This system cannot record video.')}
          </div>
        </Field>
        <Field label={t('What to export')}>
          <Select value={which} onChange={(e) => setWhich(e.target.value)}>
            <option value="all">{t('The whole project')}</option>
            <option value="range" disabled={!doc.range}>{t('From the In point to the Out point')}</option>
          </Select>
        </Field>
      </div>
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
export function ExportProgress({ progress, total, fast, since = null, from = 0, note = null, stalled = null, software = false, onCancel }) {
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
      {note ? <div className="vd-export-again">{note}</div> : null}
      {fast && stalled ? (
        <div className="vd-export-stalled" data-stalled={stalled}>{software
            ? t('Stalled: the encoder has sent nothing for {seconds} seconds. After 15 the export stops.', { seconds: stalled })
            : t('Stalled: the encoder has sent nothing for {seconds} seconds. After 15 it starts again on the processor.', { seconds: stalled })}</div>
      ) : fast ? (
        <div className="vd-export-speed" data-speed={speed ? speed.toFixed(2) : ''}>
          {speed ? t('{speed}× as fast as it plays, about {left} left', { speed: speed.toFixed(1), left: timecode(left) }) : t('Starting the encoder…')}
        </div>
      ) : null}
    </Dialog>
  );
}
