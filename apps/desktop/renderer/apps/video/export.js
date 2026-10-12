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
import { Button, Dialog, Field, Select, Progress, t } from '@rutba/office-ui';
import { timecode } from '@rutba/media/timeline';
import { buildPlan, renderVideo, soundLayers, clipFromSoundLayer, pickMimeType, extensionFor, loadVideo, releaseVideos } from '@rutba/studio/host';
import { planPatches, projectMedia, projectDuration, cutToRange } from '@rutba/studio/project';

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
export async function renderProject(doc, media, { quality = 'high', range = null, onProgress, signal } = {}) {
  const source = range ? cutToRange(doc, range.in, range.out) : doc;
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
export function ExportDialog({ doc, onClose, onExport }) {
  const format = exportFormat();
  const [quality, setQuality] = useState('high');
  const [which, setWhich] = useState(doc.range ? 'range' : 'all');
  const length = which === 'range' && doc.range ? doc.range.out - doc.range.in : projectDuration(doc);
  return (
    <Dialog
      title={t('Export video')}
      onClose={onClose}
      width={420}
      className="vd-export-dialog"
      actions={
        <>
          <Button label={t('Cancel')} onClick={onClose} />
          <Button primary className="vd-export-go" label={t('Export')} disabled={!format || length <= 0} onClick={() => onExport({ quality, range: which === 'range' ? doc.range : null, format })} />
        </>
      }
    >
      <Field label={t('Format')}>
        <div className="vd-export-format">{format ? format.label : t('This system cannot record video.')}</div>
      </Field>
      <Field label={t('Quality')}>
        <Select value={quality} onChange={(e) => setQuality(e.target.value)}>
          <option value="high">{t('High')}</option>
          <option value="medium">{t('Medium')}</option>
          <option value="low">{t('Low (smaller file)')}</option>
        </Select>
      </Field>
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
