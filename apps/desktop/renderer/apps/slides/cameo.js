// Insert → Cameo: the computer's camera, live, inside a shape on the slide.
//
// The camera is opened only when something shows a cameo — the stage with
// Camera Format → Preview on, or a show that reaches a slide with one — and
// closed when nothing does: one stream shared by all, counted, stopped with
// the last. The picture is mirrored, as a camera pointed at oneself is, and
// cropped to the cameo's shape. Without a camera, or when the person says
// no, the cameo's own drawing (a camera in a grey shape) is all there is.

import React, { useEffect, useRef, useState } from 'react';

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
 * handed back already stopped) is asked for again, three times.
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
          if (++tries <= 3) timer = setTimeout(take, 300);
          else setState({ stream: null, error: 'The camera stopped.' });
        };
        if (alive(stream)) track?.addEventListener('ended', ended);
        else ended();
      }).catch((err) => { if (live) setState({ stream: null, error: err?.name === 'NotAllowedError' ? 'The camera was not allowed.' : 'No camera was found.' }); });
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

/** Every cameo on a slide, live, for as long as this is shown. */
export function CameoLayer({ shapes = [], size = null, percent = false, on = true }) {
  const cameos = shapes.filter((s) => s.cameo && !s.hidden);
  const { stream } = useCamera(on && cameos.length > 0);
  if (!stream) return null;
  return cameos.map((s) => <CameoVideo key={s.id} shape={s} stream={stream} size={size} percent={percent} />);
}

export const CAMEO_CSS = `
.sl-cameo-video { position: absolute; object-fit: cover; pointer-events: none; z-index: 2; background: #000; }
`;
