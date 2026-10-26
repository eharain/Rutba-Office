// File → Export → Create a Video, and Record → Export to Video: the deck
// played into a video file — each slide for its recorded timing (or the
// seconds chosen), a slide with a transition faded in from the one before,
// and each slide's narration heard as it comes in — at Full HD, HD or
// Standard, as MP4 where this machine makes one and WebM where it cannot.
// Animations are shown finished, and an equation drawn as math is left out
// of the frame: what a picture of the slide can hold. A cameo the camera was
// recorded into plays that recording in its shape, as the show plays it, and
// a narration's speaker (hidden in the show) is not drawn.

import React, { useRef, useState } from 'react';
import { Button, Dialog, Select, t } from '@rutba/office-ui';
import { isNarration } from '@rutba/presentation/narration';
import { recordingType } from './screen-record.js';
import { withoutShapes } from './svg-shapes.js';

export const QUALITIES = [['1920', t('Full HD (1080p)')], ['1280', t('HD (720p)')], ['852', t('Standard (480p)')]];

/**
 * A slide's drawing with its pictures carried inside it, and nothing a canvas
 * may not hold. `cache` keeps each picture's data once for the whole video, so
 * a logo on every slide is fetched and encoded once, not once a slide.
 */
async function standalone(svg, cache = new Map()) {
  const urls = [...new Set([...svg.matchAll(/(?:href|xlink:href)="(rutba:[^"]+|blob:[^"]+)"/g)].map((m) => m[1]))];
  let out = svg.replace(/<foreignObject\b[\s\S]*?<\/foreignObject>/g, '');
  for (const url of urls) {
    try {
      if (!cache.has(url)) {
        const blob = await (await fetch(url)).blob();
        cache.set(url, String(await new Promise((resolve) => { const r = new FileReader(); r.onload = () => resolve(r.result); r.readAsDataURL(blob); })));
      }
      out = out.split(url).join(cache.get(url));
    } catch { /* a picture that will not load is left out of the frame */ }
  }
  return out;
}

/** A cameo's recording, ready to play: fetched whole, so drawing it keeps the canvas the window's own. */
async function clipOf(c) {
  const blob = await (await fetch(c.url)).blob();
  const video = document.createElement('video');
  video.muted = true;
  video.playsInline = true;
  video.preload = 'auto';
  video.src = URL.createObjectURL(blob);
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('the recording did not load')), 8000);
    video.addEventListener('loadeddata', () => { clearTimeout(timer); resolve(); }, { once: true });
    video.addEventListener('error', () => { clearTimeout(timer); reject(new Error('the recording did not load')); }, { once: true });
  });
  return { ...c, video };
}

/** One cameo's recording drawn in its shape (a rectangle, an oval or a rounded one), filling it as the show's does. */
function drawClip(g, clip, scale) {
  const { video, geometry: b, preset, rot } = clip;
  const vw = video.videoWidth || 16;
  const vh = video.videoHeight || 9;
  const x = b.x * scale;
  const y = b.y * scale;
  const w = b.w * scale;
  const h = b.h * scale;
  g.save();
  g.translate(x + w / 2, y + h / 2);
  if (rot) g.rotate((rot * Math.PI) / 180);
  g.beginPath();
  if (preset === 'ellipse') g.ellipse(0, 0, w / 2, h / 2, 0, 0, Math.PI * 2);
  else if (preset === 'roundRect' && g.roundRect) g.roundRect(-w / 2, -h / 2, w, h, Math.min(w, h) * 0.16);
  else g.rect(-w / 2, -h / 2, w, h);
  g.clip();
  const k = Math.max(w / vw, h / vh);
  g.drawImage(video, -(vw * k) / 2, -(vh * k) / 2, vw * k, vh * k);
  g.restore();
}

const imageOf = (svg) => new Promise((resolve, reject) => {
  const img = new Image();
  img.onload = () => resolve(img);
  img.onerror = () => reject(new Error(t('a slide could not be drawn')));
  img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
});

/**
 * The deck as video bytes. `slides` are `{ svg, seconds, fade, narration, clips }`:
 * the drawing, how long it stays, whether it fades in, its narration's
 * URL, and the cameo recordings played over it ({ url, geometry, preset,
 * rot } in the slide's own units, `scale` turning them to the canvas's),
 * played into a canvas of `width`×`height` and recorded. `onProgress(i)`
 * as each slide comes up; `cancelled()` stops it early.
 */
export async function renderVideo(slides, { width, height, scale = 1, onProgress = () => {}, cancelled = () => false }) {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const g = canvas.getContext('2d');
  g.fillStyle = '#000';
  g.fillRect(0, 0, width, height);
  const stream = canvas.captureStream(30);
  const track = stream.getVideoTracks()[0];
  const audio = new AudioContext();
  const sink = audio.createMediaStreamDestination();
  const tracks = [track, ...sink.stream.getAudioTracks()];
  const type = recordingType();
  let rec = null;
  // Whatever stops the video — the end, Cancel, a slide that will not draw
  // — the recorder, the canvas's track and the sound are all let go.
  try {
    rec = new MediaRecorder(new MediaStream(tracks), { mimeType: type });
    const parts = [];
    rec.ondataavailable = (e) => { if (e.data?.size) parts.push(e.data); };
    const done = new Promise((resolve) => { rec.onstop = resolve; });
    // A slide's picture is made while the one before it plays: three at most
    // at once, not every slide of a long deck decoded before the first frame.
    const pictures = new Map();
    const imageAt = (i) => {
      if (i >= slides.length || (i > 0 && cancelled())) return null;
      const made = standalone(slides[i].svg, pictures).then(imageOf);
      made.catch(() => {}); // waited on below; one left behind by Cancel is not an error
      return made;
    };
    let previous = null;
    let current = await imageAt(0);
    let next = imageAt(1);
    // A slide's cameo recordings, loaded while the slide before it plays.
    const clipsAt = (i) => {
      if (i >= slides.length) return Promise.resolve([]);
      return Promise.all((slides[i].clips || []).map((c) => clipOf(c).catch(() => null))).then((list) => list.filter(Boolean));
    };
    let playing = await clipsAt(0);
    let nextClips = clipsAt(1);
    let first = true;
    const draw = (img, alpha = 1, clips = []) => {
      g.globalAlpha = alpha;
      g.drawImage(img, 0, 0, width, height);
      for (const c of clips) drawClip(g, c, scale);
      g.globalAlpha = 1;
      track.requestFrame?.();
    };
    const letGo = (clips) => { for (const c of clips) { c.video.pause(); URL.revokeObjectURL(c.video.src); } };
    draw(current, 1, playing);
    rec.start(500);
    const frame = () => new Promise((r) => setTimeout(r, 1000 / 30));
    for (let i = 0; i < slides.length && !cancelled(); i++) {
      if (i > 0) { previous = current; current = await next; next = imageAt(i + 1); }
      onProgress(i);
      const s = slides[i];
      // The slide's cameo recordings, each from its start as the slide comes in;
      // one that will not load leaves the cameo's own picture.
      const before = first ? [] : playing;
      if (!first) { playing = await nextClips; nextClips = clipsAt(i + 1); }
      first = false;
      for (const c of playing) c.video.play().catch(() => {});
      if (s.narration) {
        try {
          const buf = await audio.decodeAudioData(await (await fetch(s.narration)).arrayBuffer());
          const src = audio.createBufferSource();
          src.buffer = buf;
          src.connect(sink);
          src.start();
        } catch { /* a narration that will not play is left out */ }
      }
      const start = performance.now();
      const fadeMs = previous && s.fade ? 500 : 0;
      while (!cancelled()) {
        const t = performance.now() - start;
        if (t >= s.seconds * 1000) break;
        if (t < fadeMs) { draw(previous, 1, before); draw(current, t / fadeMs, playing); } else draw(current, 1, playing);
        await frame();
      }
      letGo(before);
    }
    letGo(playing);
    nextClips.then(letGo).catch(() => {});
    rec.stop();
    await done;
    const contentType = type.split(';')[0];
    return { bytes: new Uint8Array(await new Blob(parts, { type: contentType }).arrayBuffer()), contentType };
  } finally {
    if (rec && rec.state !== 'inactive') { try { rec.stop(); } catch { /* stopped already */ } }
    tracks.forEach((t) => t.stop());
    await audio.close().catch(() => {});
  }
}

/** The Export to Video box: quality, timings and narrations or seconds per slide, then the file, with progress. */
export function ExportVideoDialog({ shell, doc, model, onClose, toast }) {
  const [quality, setQuality] = useState('1280');
  const [useTimings, setUseTimings] = useState(true);
  const [seconds, setSeconds] = useState('5');
  const [busy, setBusy] = useState(null);
  const stop = useRef(false);
  const size = model?.size || { width: 1280, height: 720 };
  const go = async () => {
    const width = Number(quality);
    const height = Math.round((width * size.height) / size.width / 2) * 2;
    const ext = recordingType().startsWith('video/mp4') ? 'mp4' : 'webm';
    const base = (doc?.name || t('Presentation')).replace(/\.[^.]+$/, '');
    const target = await shell.dialog.save({ title: t('Export to Video'), defaultPath: `${base}.${ext}`, filters: [{ name: ext === 'mp4' ? t('MPEG-4 Video') : t('WebM Video'), extensions: [ext] }] });
    if (!target) return;
    stop.current = false;
    setBusy({ at: 0, of: model.count });
    try {
      const slides = [];
      for (const o of model.outline || []) {
        if (o.hidden) continue;
        const m = await shell.doc.model({ id: doc.id, slide: o.index, width });
        const s = m?.slide;
        if (!s) continue;
        const shapes = s.shapes || [];
        const narration = useTimings ? shapes.find((x) => isNarration(x) && x.media?.url)?.media.url || null : null;
        const timed = useTimings ? s.transition?.advanceAfter : null;
        // The speaker is hidden in the show, so not drawn; a cameo's recording plays in its place.
        const svg = withoutShapes(s.svg, shapes.filter(isNarration).map((x) => x.id));
        const clips = shapes.filter((x) => x.cameoRecording && !x.hidden && x.media?.url && x.geometry)
          .map((x) => ({ url: x.media.url, geometry: x.geometry, preset: x.preset, rot: x.geometry.rot || 0 }));
        slides.push({ svg, seconds: Math.max(0.5, timed ?? (Number(seconds) || 5)), fade: Boolean(s.transition && s.transition.type && s.transition.type !== 'none'), narration, clips });
      }
      const video = await renderVideo(slides, { width, height, scale: width / size.width, onProgress: (i) => setBusy({ at: i + 1, of: slides.length }), cancelled: () => stop.current });
      if (stop.current) { setBusy(null); return; }
      await shell.fs.write({ path: target, bytes: video.bytes });
      toast?.(t('Video saved — {name}', { name: target.split(/[\\/]/).pop() }), { tone: 'good' });
      onClose();
    } catch (err) {
      setBusy(null);
      toast?.(t('The video could not be made: {reason}', { reason: err.message || err }), { ms: 5000 });
    }
  };
  return (
    <Dialog
      title={t('Export to Video')}
      width={440}
      onClose={busy ? () => { stop.current = true; } : onClose}
      actions={busy
        ? <Button label={t('Cancel')} onClick={() => { stop.current = true; }} />
        : <><Button label={t('Cancel')} onClick={onClose} /><Button primary label={t('Create Video')} className="sl-video-go" onClick={go} /></>}
    >
      {busy ? (
        <div className="sl-video-busy" role="status" aria-live="polite">{t('Making the video: slide {at} of {of}. It plays the deck through, so it takes as long as the show.', { at: busy.at, of: busy.of })}</div>
      ) : (
        <div className="sl-video">
          <label>{t('Quality')} <Select className="rw-select sl-video-quality" value={quality} onChange={(e) => setQuality(e.target.value)}>{QUALITIES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</Select></label>
          <label className="sl-video-check"><input type="checkbox" className="sl-video-timings" checked={useTimings} onChange={(e) => setUseTimings(e.target.checked)} /> {t('Use Recorded Timings and Narrations')}</label>
          <label>{t('Seconds spent on each slide')} <input type="number" min="1" max="60" className="rw-input sl-video-seconds" value={seconds} onChange={(e) => setSeconds(e.target.value)} /></label>
          <p className="sl-video-lead">{t('A slide with a timing of its own keeps it when timings are used; the rest stay up for the seconds above. Animations are shown finished.')}</p>
        </div>
      )}
    </Dialog>
  );
}

export const EXPORT_VIDEO_CSS = `
.sl-video { display: flex; flex-direction: column; gap: 10px; font-size: 13px; }
.sl-video label { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
.sl-video .sl-video-check { justify-content: flex-start; }
.sl-video input[type=number] { width: 72px; }
.sl-video-lead { margin: 0; font-size: 11.5px; color: var(--ink-3); }
.sl-video-busy { font-size: 13px; padding: 8px 0; }
`;
