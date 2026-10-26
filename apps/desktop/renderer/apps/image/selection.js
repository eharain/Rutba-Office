// The selection on the canvas (docs/STUDIO.md 7.2, I11): its outline drawn
// as marching ants, black and white dashes that crawl, over the picture at
// the window's scale. The outline is worked out once per selection in the
// image worker (tool-jobs.js) and only drawn here.

import React, { useEffect, useRef, useState } from 'react';
import { toolOff } from '../../imaging/work.js';

/** The outline of a selection, worked out off the window's thread: a Float32Array of segments, or null. */
export function useOutline(selection) {
  const [segments, setSegments] = useState(null);
  useEffect(() => {
    if (!selection) { setSegments(null); return undefined; }
    let alive = true;
    const copy = new Uint8Array(selection.data);
    toolOff({ op: 'outline', width: selection.width, height: selection.height, mask: copy.buffer })
      .then(({ segments: buf }) => { if (alive) setSegments(new Float32Array(buf)); })
      .catch(() => { if (alive) setSegments(null); });
    return () => { alive = false; };
  }, [selection]);
  return segments;
}

/**
 * The ants: a canvas over the picture (`width` × `height` CSS pixels for a
 * picture `size` pixels big) on which the outline is drawn twice, white
 * under black dashes, the dashes moved a step a few times a second.
 */
export function SelectionAnts({ selection, size, width, height, bounds }) {
  const ref = useRef(null);
  const segments = useOutline(selection);
  useEffect(() => {
    const c = ref.current;
    if (!c || !segments) return undefined;
    const dpr = window.devicePixelRatio || 1;
    c.width = Math.max(1, Math.round(width * dpr));
    c.height = Math.max(1, Math.round(height * dpr));
    const k = (width * dpr) / size.width;
    const path = new Path2D();
    for (let i = 0; i < segments.length; i += 4) {
      path.moveTo(segments[i] * k + 0.5, segments[i + 1] * k + 0.5);
      path.lineTo(segments[i + 2] * k + 0.5, segments[i + 3] * k + 0.5);
    }
    const g = c.getContext('2d');
    let step = 0;
    const draw = () => {
      g.clearRect(0, 0, c.width, c.height);
      g.lineWidth = Math.max(1, dpr);
      g.setLineDash([]);
      g.strokeStyle = '#ffffff';
      g.stroke(path);
      g.setLineDash([4 * dpr, 4 * dpr]);
      g.lineDashOffset = -step;
      g.strokeStyle = '#000000';
      g.stroke(path);
      step = (step + dpr) % (8 * dpr);
    };
    draw();
    const timer = setInterval(draw, 110);
    return () => clearInterval(timer);
  }, [segments, size.width, width, height]);
  if (!selection) return null;
  return (
    <canvas
      ref={ref}
      className="im-ants"
      style={{ width, height }}
      data-selection={bounds ? `${bounds.x},${bounds.y},${bounds.width},${bounds.height}` : ''}
      data-outlined={segments ? String(segments.length / 4) : ''}
    />
  );
}

export const SELECTION_CSS = `
.im-ants { position: absolute; left: 0; top: 0; pointer-events: none; z-index: 3; }
`;
