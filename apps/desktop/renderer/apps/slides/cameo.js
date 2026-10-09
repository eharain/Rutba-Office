// Insert → Cameo: the computer's camera, live, inside a shape on the slide.
//
// The camera is opened only when something shows a cameo — the stage with
// Camera Format → Preview on, or a show that reaches a slide with one — and
// closed when nothing does: one stream shared by all, counted, stopped with
// the last. The picture is mirrored, as a camera pointed at oneself is, and
// cropped to the cameo's shape. Without a camera, or when the person says
// no, the cameo's own drawing (a camera in a grey shape) is all there is.

import React, { useEffect, useRef, useState } from 'react';
import { t } from '@rutba/office-ui';
import { recordingType } from './screen-record.js';

let shared = null; // { stream, users } once opened
let opening = null;

/** Whether a stream still has its camera: a stream whose camera stopped is not handed out again. */
const alive = (stream) => Boolean(stream?.active && stream.getVideoTracks().some((t) => t.readyState === 'live'));

async function openCamera() {
  if (shared && !alive(shared.stream)) {
    for (const t of shared.stream.getTracks()) t.stop();
    shared = null;
  }
  if (shared) { shared.users += 1; return shared.stream; }
  if (!opening) {
    opening = navigator.mediaDevices.getUserMedia({ video: { width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false })
      .then((stream) => { shared = { stream, users: 0 }; return stream; })
      .finally(() => { opening = null; });
  }
  const stream = await opening;
  if (!shared || shared.stream !== stream) shared = { stream, users: 0 };
  shared.users += 1;
  return stream;
}

/** A camera that is there but busy, or slow to start, is asked for again rather than given up on. */
const BUSY = ['NotReadableError', 'AbortError', 'TrackStartError'];

/**
 * The camera, asked for again while it is busy — let go a moment ago by the
 * stage's preview, held by another program, or slow to start — waiting a
 * little longer each time, for about nine seconds, as `useCamera` does.
 */
async function openCameraPatiently() {
  for (let tries = 1; ; tries += 1) {
    try {
      return await openCamera();
    } catch (err) {
      if (!BUSY.includes(err?.name) || tries > 5) throw err;
      await new Promise((resolve) => { setTimeout(resolve, 300 * 2 ** (tries - 1)); });
    }
  }
}

/** One holder of `stream` lets it go; the camera is closed with the last. A stream already replaced is left alone. */
function closeCamera(stream) {
  if (!shared || (stream && shared.stream !== stream)) return;
  shared.users = Math.max(0, shared.users - 1);
  if (!shared.users) {
    for (const t of shared.stream.getTracks()) t.stop();
    shared = null;
  }
}

/**
 * The camera's stream while `on`, or null — and why not, when it could not
 * be opened. A camera that stops (unplugged, taken by another program, or
 * handed back already stopped) is asked for again, waiting a little longer
 * each time, for about nine seconds: a camera another program held for a
 * moment is let go by then.
 */
export function useCamera(on) {
  const [state, setState] = useState({ stream: null, error: null });
  useEffect(() => {
    if (!on) return undefined;
    let live = true;
    let held = null;
    let tries = 0;
    let timer = null;
    const take = () => {
      openCamera().then((stream) => {
        if (!live) { closeCamera(stream); return; }
        held = stream;
        setState({ stream, error: null });
        const track = stream.getVideoTracks()[0];
        const ended = () => {
          track?.removeEventListener('ended', ended);
          if (!live || held !== stream) return;
          closeCamera(stream);
          held = null;
          if (++tries <= 5) timer = setTimeout(take, 300 * 2 ** (tries - 1));
          else setState({ stream: null, error: t('The camera stopped.') });
        };
        if (alive(stream)) track?.addEventListener('ended', ended);
        else ended();
      }).catch((err) => {
        if (!live) return;
        // A camera that is there but busy — let go a moment ago by the stage's
        // preview, or held by another program — is asked for again, as one
        // that stops is; only no camera, or one not allowed, is the answer.
        if (BUSY.includes(err?.name) && ++tries <= 5) {
          timer = setTimeout(take, 300 * 2 ** (tries - 1));
          return;
        }
        setState({ stream: null, error: err?.name === 'NotAllowedError' ? t('The camera was not allowed.') : err?.name === 'NotReadableError' ? t('The camera is in use by another program.') : t('No camera was found.') });
      });
    };
    take();
    return () => {
      live = false;
      clearTimeout(timer);
      if (held) closeCamera(held);
      held = null;
      setState({ stream: null, error: null });
    };
  }, [on]);
  return state;
}

const RADIUS = { ellipse: '50%', roundRect: '16%' };

/**
 * The live picture over one cameo. `box` places it — in pixels on the stage,
 * or in percentages of the slide in the show (`percent`).
 */
export function CameoVideo({ shape, stream, size = null, percent = false }) {
  const ref = useRef(null);
  useEffect(() => {
    const v = ref.current;
    if (!v || !stream) return;
    if (v.srcObject !== stream) v.srcObject = stream;
    v.play?.().catch(() => {});
  }, [stream]);
  const g = shape.geometry;
  if (!g || !stream) return null;
  const W = size?.width || 1;
  const H = size?.height || 1;
  const box = percent
    ? { left: `${(100 * g.x) / W}%`, top: `${(100 * g.y) / H}%`, width: `${(100 * g.w) / W}%`, height: `${(100 * g.h) / H}%` }
    : { left: g.x, top: g.y, width: g.w, height: g.h };
  return (
    <video
      ref={ref}
      className="sl-cameo-video"
      data-shape={shape.id}
      muted
      playsInline
      autoPlay
      style={{ ...box, borderRadius: RADIUS[shape.preset] || 0, transform: `${g.rot ? `rotate(${g.rot}deg) ` : ''}scaleX(-1)` }}
    />
  );
}

/**
 * Record with the camera: the camera recorded one slide at a time —
 * `mark(slide)` closes the slide that was up and starts on `slide`, if it
 * has a cameo; `stop()` answers each slide's `{ slide, data, type, ms }`.
 * A camera that stops part way through a slide keeps what it took, and is
 * asked for again (three times) to go on with the slide; a slide shown
 * twice keeps its last showing, as PowerPoint records over it.
 */
export class CameraRecorder {
  static async open(slides) {
    const rec = new CameraRecorder(new Set(slides));
    await rec._take();
    return rec;
  }
  constructor(slides) {
    this.slides = slides;
    this.stream = null;
    this.parts = [];
    this.pending = [];
    this.current = null;
    this.paused = false;
    this.closed = false;
    this.showing = 0;
    this.at = null;
    this.tries = 0;
  }
  async _take() {
    const stream = await openCameraPatiently();
    if (this.closed) { closeCamera(stream); return; }
    this.stream = stream;
    const track = stream.getVideoTracks()[0];
    if (alive(stream)) track?.addEventListener('ended', () => this._lost(stream), { once: true });
    else this._lost(stream);
  }
  /** The camera stopped: the slide's recording so far is kept, and the camera asked for again to go on with it. */
  _lost(stream) {
    if (this.closed || this.stream !== stream) return;
    closeCamera(stream);
    this.stream = null;
    this._close();
    if (++this.tries > 3) return;
    setTimeout(() => {
      this._take().then(() => {
        if (!this.closed && this.stream && !this.current && this.slides.has(this.at)) this._begin(this.at);
      }).catch(() => {});
    }, 300 * this.tries);
  }
  _begin(slide) {
    const chunks = [];
    // MP4 where the computer can make it, as PowerPoint plays it; WebM where not.
    const type = recordingType();
    const recorder = new MediaRecorder(this.stream, { mimeType: type });
    const c = { slide, showing: this.showing, recorder, at: Date.now(), held: 0, heldAt: null };
    recorder.ondataavailable = (e) => { if (e.data?.size) chunks.push(e.data); };
    // However it stops — closed here, or by the camera stopping — what it took is kept.
    c.done = new Promise((resolve) => {
      recorder.addEventListener('stop', async () => {
        try {
          const now = Date.now();
          const ms = now - c.at - c.held - (c.heldAt != null ? now - c.heldAt : 0);
          const kind = type.split(';')[0];
          const data = new Uint8Array(await new Blob(chunks, { type: kind }).arrayBuffer());
          if (data.length) this.parts.push({ slide, showing: c.showing, data, type: kind, ms });
        } finally {
          resolve();
        }
      }, { once: true });
    });
    this.pending.push(c.done);
    recorder.start(500);
    if (this.paused) { recorder.pause(); c.heldAt = Date.now(); }
    this.current = c;
  }
  _close() {
    const c = this.current;
    this.current = null;
    if (c && c.recorder.state !== 'inactive') {
      try { c.recorder.stop(); } catch { /* it has stopped already */ }
    }
  }
  mark(slide) {
    this._close();
    this.showing += 1;
    this.at = slide;
    if (this.slides.has(slide) && this.stream) this._begin(slide);
  }
  /** Record → Pause and back: the camera waits with the voice. */
  pause() {
    this.paused = true;
    const c = this.current;
    if (c?.recorder.state === 'recording') { c.recorder.pause(); c.heldAt = Date.now(); }
  }
  resume() {
    this.paused = false;
    const c = this.current;
    if (c?.recorder.state === 'paused') {
      c.recorder.resume();
      if (c.heldAt != null) c.held += Date.now() - c.heldAt;
      c.heldAt = null;
    }
  }
  async stop() {
    this.closed = true;
    this._close();
    await Promise.all(this.pending);
    if (this.stream) closeCamera(this.stream);
    this.stream = null;
    // Each slide's last showing, and of that the longest stretch the camera kept going.
    const best = new Map();
    for (const p of this.parts) {
      const b = best.get(p.slide);
      if (!b || p.showing > b.showing || (p.showing === b.showing && p.ms > b.ms)) best.set(p.slide, p);
    }
    return [...best.values()].map(({ slide, data, type, ms }) => ({ slide, data, type, ms }));
  }
}

/** Every cameo on a slide, live, for as long as this is shown — but where Record put the camera's recording in its place. */
export function CameoLayer({ shapes = [], size = null, percent = false, on = true }) {
  const recorded = shapes.some((s) => s.cameoRecording);
  const cameos = recorded ? [] : shapes.filter((s) => s.cameo && !s.hidden);
  const { stream } = useCamera(on && cameos.length > 0);
  if (!stream) return null;
  return cameos.map((s) => <CameoVideo key={s.id} shape={s} stream={stream} size={size} percent={percent} />);
}

export const CAMEO_CSS = `
.sl-cameo-video { position: absolute; object-fit: cover; pointer-events: none; z-index: 2; background: #000; }
`;
