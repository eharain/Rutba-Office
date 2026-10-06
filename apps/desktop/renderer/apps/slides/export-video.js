// File → Export → Create a Video, and Record → Export to Video: the deck
// played into a video file — each slide for its recorded timing (or the
// seconds chosen), a slide with a transition faded in from the one before,
// and each slide's narration heard as it comes in — at Full HD, HD or
// Standard, as MP4 where this machine makes one and WebM where it cannot.
// Animations are shown finished, and an equation drawn as math is left out
// of the frame: what a picture of the slide can hold.

import React, { useRef, useState } from 'react';
import { Button, Dialog, Select } from '@rutba/office-ui';
import { isNarration } from '@rutba/presentation/narration';
import { recordingType } from './screen-record.js';

export const QUALITIES = [['1920', 'Full HD (1080p)'], ['1280', 'HD (720p)'], ['852', 'Standard (480p)']];

/** A slide's drawing with its pictures carried inside it, and nothing a canvas may not hold. */
async function standalone(svg) {
  const urls = [...new Set([...svg.matchAll(/(?:href|xlink:href)="(rutba:[^"]+|blob:[^"]+)"/g)].map((m) => m[1]))];
  let out = svg.replace(/<foreignObject\b[\s\S]*?<\/foreignObject>/g, '');
  for (const url of urls) {
    try {
      const blob = await (await fetch(url)).blob();
      const data = await new Promise((resolve) => { const r = new FileReader(); r.onload = () => resolve(r.result); r.readAsDataURL(blob); });
      out = out.split(url).join(String(data));
    } catch { /* a picture that will not load is left out of the frame */ }
  }
  return out;
}

const imageOf = (svg) => new Promise((resolve, reject) => {
  const img = new Image();
  img.onload = () => resolve(img);
  img.onerror = () => reject(new Error('a slide could not be drawn'));
  img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
});

/**
 * The deck as video bytes. `slides` are `{ svg, seconds, fade, narration }`
 * — the drawing, how long it stays, whether it fades in, and its narration's
 * URL — played into a canvas of `width`×`height` and recorded. `onProgress(i)`
 * as each slide comes up; `cancelled()` stops it early.
 */
export async function renderVideo(slides, { width, height, onProgress = () => {}, cancelled = () => false }) {
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
  const rec = new MediaRecorder(new MediaStream(tracks), { mimeType: type });
  const parts = [];
  rec.ondataavailable = (e) => { if (e.data?.size) parts.push(e.data); };
  const done = new Promise((resolve) => { rec.onstop = resolve; });
  const images = [];
  for (const s of slides) images.push(await imageOf(await standalone(s.svg)));
  const draw = (img, alpha = 1) => { g.globalAlpha = alpha; g.drawImage(img, 0, 0, width, height); g.globalAlpha = 1; track.requestFrame?.(); };
  draw(images[0]);
  rec.start(500);
  const frame = () => new Promise((r) => setTimeout(r, 1000 / 30));
  for (let i = 0; i < slides.length && !cancelled(); i++) {
    onProgress(i);
    const s = slides[i];
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
    const fadeMs = i > 0 && s.fade ? 500 : 0;
    while (!cancelled()) {
      const t = performance.now() - start;
      if (t >= s.seconds * 1000) break;
      if (t < fadeMs) { draw(images[i - 1]); draw(images[i], t / fadeMs); } else draw(images[i]);
      await frame();
    }
  }
  rec.stop();
  await done;
  tracks.forEach((t) => t.stop());
  await audio.close().catch(() => {});
  const contentType = type.split(';')[0];
  return { bytes: new Uint8Array(await new Blob(parts, { type: contentType }).arrayBuffer()), contentType };
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
    const base = (doc?.name || 'Presentation').replace(/\.[^.]+$/, '');
    const target = await shell.dialog.save({ title: 'Export to Video', defaultPath: `${base}.${ext}`, filters: [{ name: ext === 'mp4' ? 'MPEG-4 Video' : 'WebM Video', extensions: [ext] }] });
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
        const narration = useTimings ? (s.shapes || []).find((x) => isNarration(x) && x.media?.url)?.media.url || null : null;
        const timed = useTimings ? s.transition?.advanceAfter : null;
        slides.push({ svg: s.svg, seconds: Math.max(0.5, timed ?? (Number(seconds) || 5)), fade: Boolean(s.transition && s.transition.type && s.transition.type !== 'none'), narration });
      }
      const video = await renderVideo(slides, { width, height, onProgress: (i) => setBusy({ at: i + 1, of: slides.length }), cancelled: () => stop.current });
      if (stop.current) { setBusy(null); return; }
      await shell.fs.write({ path: target, bytes: video.bytes });
      toast?.(`Video saved — ${target.split(/[\\/]/).pop()}`, { tone: 'good' });
      onClose();
    } catch (err) {
      setBusy(null);
      toast?.(`The video could not be made: ${err.message || err}`, { ms: 5000 });
    }
  };
  return (
    <Dialog
      title="Export to Video"
      width={440}
      onClose={busy ? () => { stop.current = true; } : onClose}
      actions={busy
        ? <Button label="Cancel" onClick={() => { stop.current = true; }} />
        : <><Button label="Cancel" onClick={onClose} /><Button primary label="Create Video" className="sl-video-go" onClick={go} /></>}
    >
      {busy ? (
        <div className="sl-video-busy">Making the video: slide {busy.at} of {busy.of}. It plays the deck through, so it takes as long as the show.</div>
      ) : (
        <div className="sl-video">
          <label>Quality <Select className="rw-select sl-video-quality" value={quality} onChange={(e) => setQuality(e.target.value)}>{QUALITIES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</Select></label>
          <label className="sl-video-check"><input type="checkbox" className="sl-video-timings" checked={useTimings} onChange={(e) => setUseTimings(e.target.checked)} /> Use Recorded Timings and Narrations</label>
          <label>Seconds spent on each slide <input type="number" min="1" max="60" className="rw-input sl-video-seconds" value={seconds} onChange={(e) => setSeconds(e.target.value)} /></label>
          <p className="sl-video-lead">A slide with a timing of its own keeps it when timings are used; the rest stay up for the seconds above. Animations are shown finished.</p>
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
