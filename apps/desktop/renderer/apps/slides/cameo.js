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

async function openCamera() {
  if (shared) { shared.users += 1; return shared.stream; }
  if (!opening) {
    opening = navigator.mediaDevices.getUserMedia({ video: { width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false })
      .then((stream) => { shared = { stream, users: 0 }; return stream; })
      .finally(() => { opening = null; });
  }
  const stream = await opening;
  shared.users += 1;
  return stream;
}

function closeCamera() {
  if (!shared) return;
  shared.users -= 1;
  if (shared.users <= 0) {
    for (const t of shared.stream.getTracks()) t.stop();
    shared = null;
  }
}

/** The camera's stream while `on`, or null — and why not, when it could not be opened. */
export function useCamera(on) {
  const [state, setState] = useState({ stream: null, error: null });
  useEffect(() => {
    if (!on) return undefined;
    let live = true;
    let held = false;
    openCamera().then((stream) => {
      held = true;
      if (live) setState({ stream, error: null });
      else closeCamera();
    }).catch((err) => { if (live) setState({ stream: null, error: err?.name === 'NotAllowedError' ? 'The camera was not allowed.' : 'No camera was found.' }); });
    return () => {
      live = false;
      if (held) closeCamera();
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
