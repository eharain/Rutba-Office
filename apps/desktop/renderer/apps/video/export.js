// Export: the project played into a video file, with its sound.
//
// The engine's renderVideo plays the plan in real time into a canvas stream
// and a MediaRecorder, the sound layers mixed into the same recording (their
// volumes, fades and crossfades as the preview plays them). The file is MP4
// (H.264 and AAC) where this machine records it and WebM (VP9 and Opus)
// where it does not, as pickMimeType decides, and it is written through the
// shell. An export uses decoders of its own, so the preview is left as it
// was; it takes as long as the project runs, and says so.

import React, { useState } from 'react';
import { Button, Dialog, Field, Select, Progress, formatBytes, t } from '@rutba/office-ui';
import { timecode } from '@rutba/media/timeline';
import { buildPlan, renderVideo, soundLayers, clipFromSoundLayer, pickMimeType, extensionFor, loadVideo, releaseVideos } from '@rutba/studio/host';
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

/** What this machine will write: { mime, ext, label }. */
export function exportFormat() {
  const mime = pickMimeType(true);
  if (!mime) return null;
  const ext = extensionFor(mime);
  return { mime, ext, label: ext === 'mp4' ? t('MP4 video (H.264 and AAC)') : t('WebM video (VP9 and Opus)') };
}

/**
 * Render a project to bytes. `media` is the window's library (its Blobs are
 * reused); `onProgress(0..1)`; `signal` cancels. Answers { bytes, mimeType,
 * extension, duration, hasAudio }.
 */
export async function renderProject(doc, media, { quality = 'high', range = null, preset = 'project', relayout = null, onProgress, signal } = {}) {
  // The stretch asked for, then made ready for where it is going: its frame,
  // rate and bitrates, its captions and its scenes laid out for that shape.
  const source = presetDoc(range ? cutToRange(doc, range.in, range.out) : doc, preset, { relayout });
  const canvas = document.createElement('canvas');
  // A canvas that belongs to the page, off to one side, so the engine's
  // capture is of a live document.
  canvas.className = 'vd-export-canvas';
  document.body.appendChild(canvas);
  const want = projectMedia(source);
  const videos = {};
  try {
    for (const { key } of want.videos) {
      try {
        videos[key] = await loadVideo(key);
      } catch {
        /* a source that will not load is a clip that draws nothing */
      }
    }
    const assets = {};
    for (const url of want.images) {
      const entry = media.images.get(url) || (await media.image(url).catch(() => null));
      if (entry) assets[url] = entry;
    }
    const plan = buildPlan({ canvas, images: [], options: { ...source.options, quality }, layerPatches: planPatches(source), videos, assets });
    const clips = [];
    for (const layer of soundLayers(plan)) {
      const buffer = await media.sound(layer.url).catch(() => null);
      if (buffer) clips.push(clipFromSoundLayer(layer, buffer, plan, { volume: 1, fadeIn: 0, fadeOut: 0 }));
    }
    const out = await renderVideo({ canvas, plan, audio: clips.length ? { clips } : null, onProgress, signal });
    return { bytes: new Uint8Array(await out.blob.arrayBuffer()), mimeType: out.mimeType, extension: out.extension, duration: out.duration, hasAudio: out.hasAudio };
  } finally {
    releaseVideos(Object.values(videos));
    canvas.remove();
  }
}

/** File → Export: the format this machine writes, the quality, the whole project or the In to Out range. */
export function ExportDialog({ doc, onClose, onExport, initialPreset = 'project' }) {
  const format = exportFormat();
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
        <div className="vd-export-format">{format ? format.label : t('This system cannot record video.')}</div>
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
      <div className="rw-hint">{t('Export plays the project once as it records it, so it takes {length}.', { length: timecode(length) })}</div>
    </Dialog>
  );
}

/** The progress of an export, with Cancel. */
export function ExportProgress({ progress, total, onCancel }) {
  return (
    <Dialog title={t('Exporting')} onClose={onCancel} width={380} className="vd-export-busy" actions={<Button label={t('Cancel')} onClick={onCancel} />}>
      <div className="vd-export-line">{t('Recording {done} of {total}', { done: timecode(progress * total), total: timecode(total) })}</div>
      <Progress value={progress} />
    </Dialog>
  );
}
