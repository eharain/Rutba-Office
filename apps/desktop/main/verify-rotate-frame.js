// What the four rotate-frame blocks share (rotate-frame-word, -sheet, -deck,
// -image): the pointer driven for real through the window, and the turned
// frame's geometry to judge what is drawn against.
//
// Each block turns an inserted object to 30, 90 and 135 degrees and reads
// the frame back from the window: its four corner handles must stand on the
// object's turned corners, a click at the object's centre must take hold of
// it, a click in the empty corner of its old upright box must not, and a
// corner drag must resize it along its own sides.

import fs from 'node:fs';
import path from 'node:path';
import { frameCorners, frameBounds, framePoint, insideFrame, turnPoint } from '@rutba/office-ui/turned-box';

export { frameCorners, frameBounds, framePoint, insideFrame, turnPoint };

/** Real pointer input on a window: a click and a drag, in window pixels. */
export function pointer(win, wait) {
  const wc = win.webContents;
  const at = (p) => ({ x: Math.round(p.x), y: Math.round(p.y) });
  return {
    async click(p, { modifiers = [], count = 1 } = {}) {
      const { x, y } = at(p);
      wc.sendInputEvent({ type: 'mouseMove', x, y, modifiers });
      for (let i = 1; i <= count; i++) {
        wc.sendInputEvent({ type: 'mouseDown', x, y, button: 'left', clickCount: i, modifiers });
        wc.sendInputEvent({ type: 'mouseUp', x, y, button: 'left', clickCount: i, modifiers });
      }
      await wait(250);
    },
    async drag(from, to, { modifiers = [], steps = 8 } = {}) {
      const a = at(from);
      const b = at(to);
      wc.sendInputEvent({ type: 'mouseMove', x: a.x, y: a.y, modifiers });
      wc.sendInputEvent({ type: 'mouseDown', x: a.x, y: a.y, button: 'left', clickCount: 1, modifiers });
      await wait(60);
      for (let i = 1; i <= steps; i++) {
        wc.sendInputEvent({ type: 'mouseMove', x: Math.round(a.x + ((b.x - a.x) * i) / steps), y: Math.round(a.y + ((b.y - a.y) * i) / steps), button: 'left', modifiers });
        await wait(35);
      }
      wc.sendInputEvent({ type: 'mouseUp', x: b.x, y: b.y, button: 'left', clickCount: 1, modifiers });
      await wait(350);
    },
  };
}

/** How far apart two lists of points are at worst, in pixels. */
export function worstGap(a, b) {
  if (!a || !b || a.length !== b.length || a.some((p) => !p)) return Infinity;
  return Math.max(...a.map((p, i) => Math.hypot(p.x - b[i].x, p.y - b[i].y)));
}

/**
 * A point inside the upright box, near one of its corners, that is well
 * outside the turned one: where a click must not take hold of the object.
 * Null when every corner of the upright box is inside the turned outline.
 */
export function emptyCorner(box, deg, inset = 6) {
  const pts = [
    { x: box.x + inset, y: box.y + inset }, { x: box.x + box.w - inset, y: box.y + inset },
    { x: box.x + box.w - inset, y: box.y + box.h - inset }, { x: box.x + inset, y: box.y + box.h - inset },
  ];
  return pts.find((p) => !insideFrame(box, deg, p.x, p.y, 8)) || null;
}

/** A capture of the window as it is painted now, into the run's capture folder. */
export async function shoot(win, wait, name) {
  const dir = process.env.RUTBA_VERIFY_CAPTURE;
  if (!dir || win.isDestroyed()) return;
  win.webContents.invalidate();
  await wait(700);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, name), (await win.webContents.capturePage()).toPNG());
}

/** The JavaScript, run in a window, that reads a handle's centre on screen: `{ x, y }` or null. */
export const centreOf = (selector) => `(() => { const r = document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect(); return r && r.width ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; })()`;
