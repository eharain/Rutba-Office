// Edit Points, on the stage.
//
// The shape's outline drawn over it, each point a small square to drag; a
// right-click on a point deletes it, on the outline adds one where it was
// pressed. The outline follows the pointer as the point moves; the engine
// is told once, on release, so a drag is one undo step. The overlay turns
// and flips with the shape, so a point is dragged where it is seen.

import React from 'react';
import { anchors, moveAnchor, deleteAnchor, insertAnchor, nearestSegment, pathData } from '@rutba/presentation/points';

/**
 * @param {object} p
 * @param {object} p.shape the shape whose points are edited (its geometry)
 * @param {Array} p.commands the outline, in the shape's box
 * @param {number|null} p.picked the point last moved
 * @param {number} p.scale the stage's zoom
 * @param {object} p.menu the window's context menu
 * @param {(commands: Array, picked?: number) => void} p.onChange while dragging
 * @param {(commands: Array) => void} p.onCommit when a change is made
 * @param {() => void} p.onExit leave Edit Points
 */
export function PointsOverlay({ shape, commands, picked, scale, menu, onChange, onCommit, onExit }) {
  const g = shape.geometry;
  const size = 8 / scale;
  const rot = ((g.rot || 0) * Math.PI) / 180;
  const transform = `rotate(${g.rot || 0}deg)${g.flipH ? ' scaleX(-1)' : ''}${g.flipV ? ' scaleY(-1)' : ''}`;

  /** A pointer position in the shape's own box, through the stage's zoom and the shape's turn. */
  const local = (e) => {
    const svg = e.currentTarget.ownerSVGElement || e.currentTarget;
    const p = svg.createSVGPoint();
    p.x = e.clientX;
    p.y = e.clientY;
    return p.matrixTransform(svg.getScreenCTM().inverse());
  };

  const drag = (e, index) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const start = anchors(commands).find((a) => a.index === index);
    if (!start) return;
    const x0 = e.clientX;
    const y0 = e.clientY;
    let current = commands;
    const move = (ev) => {
      // The pointer's travel, unscaled, then turned back into the shape's frame.
      let dx = (ev.clientX - x0) / scale;
      let dy = (ev.clientY - y0) / scale;
      [dx, dy] = [dx * Math.cos(-rot) - dy * Math.sin(-rot), dx * Math.sin(-rot) + dy * Math.cos(-rot)];
      if (g.flipH) dx = -dx;
      if (g.flipV) dy = -dy;
      current = moveAnchor(commands, index, start.x + dx, start.y + dy);
      onChange(current, index);
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      if (current !== commands) onCommit(current);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  return (
    <svg
      className="sl-points"
      style={{ left: g.x, top: g.y, width: g.w, height: g.h, transform }}
      viewBox={`0 0 ${g.w} ${g.h}`}
      overflow="visible"
    >
      <path
        className="sl-points-path"
        d={pathData(commands)}
        style={{ strokeWidth: 1.5 / scale }}
        onContextMenu={(e) => {
          e.preventDefault();
          e.stopPropagation();
          const at = local(e);
          const segment = nearestSegment(commands, at.x, at.y);
          menu.open(e, [
            { label: 'Add Point', disabled: segment < 0, run: () => onCommit(insertAnchor(commands, segment)) },
            '-',
            { label: 'Exit Edit Points', run: onExit },
          ]);
        }}
      >
        <title>Right-click to add a point</title>
      </path>
      {anchors(commands).map((a) => (
        <rect
          key={a.index}
          className={`sl-point${picked === a.index ? ' on' : ''}`}
          data-point={a.index}
          x={a.x - size / 2}
          y={a.y - size / 2}
          width={size}
          height={size}
          style={{ strokeWidth: 1.25 / scale }}
          onMouseDown={(e) => drag(e, a.index)}
          onContextMenu={(e) => {
            e.preventDefault();
            e.stopPropagation();
            const next = deleteAnchor(commands, a.index);
            menu.open(e, [
              { label: 'Delete Point', disabled: !next, title: next ? undefined : 'A figure needs the points it has left', run: () => next && onCommit(next) },
              '-',
              { label: 'Exit Edit Points', run: onExit },
            ]);
          }}
        >
          <title>Drag to move this point — right-click to delete it</title>
        </rect>
      ))}
    </svg>
  );
}

export const POINTS_CSS = `
.sl-points { position: absolute; z-index: 7; overflow: visible; pointer-events: none; transform-origin: 50% 50%; }
.sl-points-path { fill: none; stroke: #1a1c20; stroke-dasharray: 4 3; pointer-events: stroke; cursor: crosshair; }
.sl-point { fill: #1a1c20; stroke: #fff; pointer-events: all; cursor: move; }
.sl-point.on { fill: var(--accent); }
`;
