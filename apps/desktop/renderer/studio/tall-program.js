// The tall picture: a second Program of the same sources, made while a tall
// stream is on. It draws when the wide one does, so the two stay in step and
// only one timer runs, and it follows the wide one's every change of shot.
//
// The canvas is kept in the page, unseen, so a check can read what was drawn.

import { Program } from './program.js';
import { tallKey, tallSize, tallStackOf } from '@rutba/studio/tall';

/** How often the layout is made again, for a source whose shape has just become known. */
const REFIT_MS = 1000;

/**
 * @param {object} o
 * @param {object} o.wide  the wide On air Program
 * @param {object} o.media  StudioMedia
 * @param {number} o.size  720 or 1080
 * @returns {object} a Program with attach() and detach()
 */
export function makeTallProgram({ wide, media, size }) {
  const canvas = document.createElement('canvas');
  const { width, height } = tallSize(size);
  canvas.width = width;
  canvas.height = height;
  canvas.className = 'st-tall-canvas';
  canvas.style.display = 'none';
  const tall = new Program({ media, canvas });

  const aspectOf = (it) => {
    const s = media.sizeOf(it);
    return s ? s.w / s.h : 16 / 9;
  };
  const follow = (stack, o) => {
    if (stack) tall.setStack(tallStackOf(stack, tallKey(stack), { aspectOf }), o);
  };

  let off = null;
  let last = 0;
  // Asked for by the tall streams and by a tall copy, each on its own.
  let users = 0;
  tall.inUse = () => users > 0;
  tall.attach = () => {
    users += 1;
    if (off) return;
    tall.settings = wide.settings;
    tall.fps = wide.fps;
    document.body.appendChild(canvas);
    wide.onStack = follow;
    follow(wide.stack);
    media.web?.setTall(true);
    off = wide.onFrame(() => {
      const now = performance.now();
      if (now - last > REFIT_MS) {
        last = now;
        follow(wide.stack);
      }
      tall.draw();
    });
  };
  tall.detach = () => {
    users = Math.max(0, users - 1);
    if (users || !off) return;
    off();
    off = null;
    wide.onStack = null;
    media.web?.setTall(false);
    canvas.remove();
  };
  return tall;
}
