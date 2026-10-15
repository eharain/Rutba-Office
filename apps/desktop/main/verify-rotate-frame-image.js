// Image: a turned layer's selection frame turns with it.
//
// A long red picture laid as a layer on a pale one. With the Move tool it
// is turned to 30, 90 and 135 degrees by its rotate grip (Shift snapping to
// 15). At each angle, in the window as it opens and in a smaller one (the
// picture is fitted to the window: Image has no zoom of its own), the
// frame's four corner grips stand on the layer's turned corners, the
// picture shows red just inside each of them, a click at the layer's
// middle takes hold of it and a click in the empty corner of its old
// upright box does not. Free Transform's corners stand where the plain
// selection's grips do. A corner drag scales it with the opposite corner
// staying put where it stands turned; Undo and Redo bring the frame back
// with the angle; right to left the frame stays on. Saved as a project and
// opened again, the angle and the frame are kept. Run alone with
// RUTBA_VERIFY_ONLY=rotate-frame-image.

import fs from 'node:fs';
import path from 'node:path';
import { imageDriver } from './verify-image-layers.js';
import { toolsDriver } from './verify-image-tools.js';
import { gradientPng } from './sample-picture.js';
import { pointer, emptyCorner, shoot, frameCorners, framePoint } from './verify-rotate-frame.js';

const W = 900;
const H = 600;
const RED = [210, 40, 40];
const isRed = (p) => p && p[0] > 170 && p[1] < 90 && p[2] < 90;

/** How far apart each expected point is from the nearest of the points got, at worst. */
function nearestGap(got, want) {
  if (!got || !got.length || got.some((p) => !p)) return Infinity;
  return Math.max(...want.map((w) => Math.min(...got.map((g) => Math.hypot(g.x - w.x, g.y - w.y)))));
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder; the block writes in a folder of its own
 */
export async function verifyRotateFrameImage(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'rotate-frame-image');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'ground.png');
  fs.writeFileSync(file, gradientPng(W, H, [236, 238, 242], [214, 222, 234]));
  const bar = path.join(root, 'bar.png');
  fs.writeFileSync(bar, gradientPng(360, 90, RED, RED));
  let win = await open('image', file);
  const sized = win.getSize();
  try {
    let d = imageDriver(win, { until, wait });
    let k = toolsDriver(d, { until, wait });
    let js = d.js;
    let mouse = pointer(win, wait);
    await until(async () => (await d.result()) === `${W}x${H}` && (await js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`)) > 0, 'the picture on the canvas', 15000);

    // The red bar, as a layer.
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [bar] });
    await d.tab('Insert');
    await d.press('[data-insert="picture"]');
    await until(async () => (await d.layers()) === 2, 'the bar as a layer', 8000).catch(() => {});
    await d.settled('the bar drawn', 6000);
    const id = (await k.rows())[0];
    await k.tool('move');
    await d.pickRow(id);

    const selected = () => js(`(document.querySelector('.im-stage-layers')?.dataset.selected || '').split(',').filter(Boolean)`);
    /** The frame on the screen: its upright box (middle and own size) and its turn, read from the drawn frame. */
    const frame = () => js(`(() => {
      const b = document.querySelector('.im-sel[data-sel=${JSON.stringify(id)}]');
      if (!b) return null;
      const r = b.getBoundingClientRect();
      const m = /rotate\\((-?[\\d.]+)deg\\)/.exec(b.style.transform || '');
      const w = b.offsetWidth; const h = b.offsetHeight;
      return { x: r.left + r.width / 2 - w / 2, y: r.top + r.height / 2 - h / 2, w, h, rot: m ? Number(m[1]) : 0 };
    })()`);
    const grips = () => js(`['nw', 'ne', 'se', 'sw'].map((n) => { const r = document.querySelector('[data-grip="' + n + '"]')?.getBoundingClientRect(); return r && r.width ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; })`);
    /** Screen pixels to the picture's, over the stage. */
    const toPicture = async (p) => {
      const r = await js(`(() => { const s = document.querySelector('.im-stage-layers')?.getBoundingClientRect(); return s ? { left: s.left, top: s.top, width: s.width } : null; })()`);
      const s = W / r.width;
      return [(p.x - r.left) * s, (p.y - r.top) * s];
    };

    /** The grips on the frame's turned corners, and red just inside each corner of the picture there. */
    const frameOn = async () => {
      let last = null;
      const ok = await until(async () => {
        const f = await frame();
        if (!f) return false;
        const want = frameCorners(f, f.rot);
        const got = await grips();
        const inside = [];
        for (const [fx, fy] of [[0.06, 0.15], [0.94, 0.15], [0.94, 0.85], [0.06, 0.85]]) inside.push(isRed(await k.px(...(await toPicture(framePoint(f, f.rot, fx, fy))))));
        last = { rot: f.rot, size: `${Math.round(f.w)}x${Math.round(f.h)}`, corners: +nearestGap(got, want).toFixed(1), redInside: inside.filter(Boolean).length };
        return last.corners <= 3 && last.redInside === 4;
      }, 'the frame on the layer', 5000).then(() => true).catch(() => false);
      return { ok, ...last };
    };
    /** A click at the middle takes hold; one in the old box's empty corner does not. */
    const clicks = async () => {
      const f = await frame();
      if (!f) return { ok: false, why: 'no frame' };
      const corner = emptyCorner(f, f.rot, 6);
      const cornerShows = corner ? await k.px(...(await toPicture(corner))) : null;
      await d.pickRow(id);
      await until(async () => (await selected()).includes(id), 'picked', 2000).catch(() => {});
      if (corner) await mouse.click(corner);
      const cornerLets = await until(async () => !(await selected()).includes(id), 'the corner to let go', 2500).then(() => true).catch(() => false);
      await mouse.click({ x: f.x + f.w / 2, y: f.y + f.h / 2 });
      const centreTakes = await until(async () => (await selected()).includes(id), 'the middle to take hold', 3000).then(() => true).catch(() => false);
      return { ok: Boolean(corner) && !isRed(cornerShows) && cornerLets && centreTakes, cornerLets, centreTakes, cornerShows };
    };
    /** Drag the rotate grip round the layer's middle to `deg`, Shift snapping to 15. */
    const turnTo = async (deg) => {
      await d.pickRow(id);
      await until(() => js(`Boolean(document.querySelector('[data-grip="rotate"]'))`), 'the rotate grip', 4000).catch(() => {});
      const f = await frame();
      const g = await js(`(() => { const r = document.querySelector('[data-grip="rotate"]')?.getBoundingClientRect(); return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; })()`);
      if (!f || !g) return false;
      const c = { x: f.x + f.w / 2, y: f.y + f.h / 2 };
      const r = Math.hypot(g.x - c.x, g.y - c.y);
      const a = (deg * Math.PI) / 180;
      await mouse.drag(g, { x: c.x + r * Math.sin(a), y: c.y - r * Math.cos(a) }, { modifiers: ['shift'], steps: 12 });
      return until(async () => Math.round(((((await frame())?.rot || 0) % 360) + 360) % 360) === deg, `the layer at ${deg}`, 5000).then(() => true).catch(() => false);
    };
    const atAngle = async (deg) => {
      const turned = await turnTo(deg);
      await d.settled('the turned layer', 4000);
      const big = await frameOn();
      await shoot(win, wait, `rotate-frame-image-${deg}.png`);
      const bigClicks = await clicks();
      // A smaller window: the picture fitted smaller, the frame still on the layer.
      win.setSize(Math.round(sized[0] * 0.72), Math.round(sized[1] * 0.78));
      await wait(500);
      await d.pickRow(id);
      const small = await frameOn();
      await shoot(win, wait, `rotate-frame-image-${deg}-small.png`);
      const smallClicks = await clicks();
      win.setSize(sized[0], sized[1]);
      await wait(500);
      return { ok: turned && big.ok && bigClicks.ok && small.ok && smallClicks.ok, turned, big, bigClicks, small, smallClicks };
    };

    for (const deg of [30, 90, 135]) {
      const r = await atAngle(deg);
      check(`rotate-frame-image: a layer turned to ${deg} by its rotate grip wears its frame and corner grips on its turned corners, in the window as it opens and in a smaller one; a click at its middle takes hold of it, one in the empty corner of its old upright box does not`,
        r.ok, JSON.stringify(r));
    }

    // Free Transform's corners stand where the plain selection's grips do.
    await d.pickRow(id);
    const plain = await grips();
    await k.tool('transform');
    await until(() => js(`document.querySelectorAll('.im-quad-handle').length === 4`), 'the Free Transform corners', 4000).catch(() => {});
    const quad = await js(`[...document.querySelectorAll('.im-quad-handle')].map((n) => { const r = n.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; })`);
    await shoot(win, wait, 'rotate-frame-image-free-transform.png');
    await k.tool('move');
    check('rotate-frame-image: Free Transform\'s four corners stand where the turned layer\'s plain selection puts its grips', nearestGap(quad, plain) <= 3, JSON.stringify({ gap: nearestGap(quad, plain), plain, quad }));

    // A corner drag scales it from the opposite corner, which stays where it stands turned.
    await d.pickRow(id);
    const f0 = await frame();
    const fixed = framePoint(f0, f0.rot, 0, 0);
    const se = (await grips())[2];
    const away = { x: se.x + (se.x - fixed.x) * 0.25, y: se.y + (se.y - fixed.y) * 0.25 };
    await mouse.drag(se, away, { steps: 10 });
    const grew = await until(async () => ((await frame())?.w || 0) > f0.w * 1.15, 'the layer scaled', 5000).then(() => true).catch(() => false);
    await d.settled('the scaled layer', 4000);
    const f1 = await frame();
    const kept = framePoint(f1, f1.rot, 0, 0);
    const moved = +Math.hypot(kept.x - fixed.x, kept.y - fixed.y).toFixed(1);
    const scaledFrame = await frameOn();
    await shoot(win, wait, 'rotate-frame-image-scaled.png');
    check('rotate-frame-image: a corner drag on the turned layer scales it with the opposite corner staying put where it stands turned, and the frame follows',
      grew && moved <= 3 && Math.round(f1.rot) === 135 && scaledFrame.ok, JSON.stringify({ before: { w: f0.w, h: f0.h }, after: { w: f1.w, h: f1.h }, moved, scaledFrame }));

    // Undo and Redo bring the frame back with the size and the angle.
    await k.undo();
    const undone = await until(async () => Math.abs(((await frame())?.w || 0) - f0.w) < 2, 'the scale undone', 5000).then(() => true).catch(() => false);
    const undoFrame = await frameOn();
    await k.undo();
    const unturned = await until(async () => Math.round((await frame())?.rot || 0) === 90, 'the last turn undone', 5000).then(() => true).catch(() => false);
    const unturnFrame = await frameOn();
    await k.redo();
    await k.redo();
    const redone = await until(async () => { const f = await frame(); return f && Math.round(f.rot) === 135 && Math.abs(f.w - f1.w) < 2; }, 'both again', 5000).then(() => true).catch(() => false);
    const redoFrame = await frameOn();
    check('rotate-frame-image: Undo takes the scale and then the turn back and Redo puts them again, the frame following each',
      undone && undoFrame.ok && unturned && unturnFrame.ok && redone && redoFrame.ok, JSON.stringify({ undone, undoFrame, unturned, unturnFrame, redone, redoFrame }));

    // Right to left, as the window is in Urdu or Arabic.
    await js(`document.documentElement.dir = 'rtl', 1`);
    await wait(500);
    await d.pickRow(id);
    const rtlFrame = await frameOn();
    await shoot(win, wait, 'rotate-frame-image-rtl.png');
    const rtlClicks = await clicks();
    await js(`document.documentElement.dir = 'ltr', 1`);
    await wait(400);
    check('rotate-frame-image: with the window right to left, the turned layer\'s frame stays on it and the clicks land by its turned outline',
      rtlFrame.ok && rtlClicks.ok, JSON.stringify({ rtlFrame, rtlClicks }));

    // Saved as a project and opened again: the angle and the frame are kept.
    const project = path.join(root, 'turned.rimg');
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: project });
    await d.tab('Home');
    await d.press('[data-role="save-project"]');
    const wrote = await until(() => fs.existsSync(project) && fs.statSync(project).size > 1000, 'the project', 8000).then(() => true).catch(() => false);
    const rotInFile = wrote ? /"rot":\s*135\b/.test(fs.readFileSync(project).toString('latin1')) : false;
    win = await open('image', project);
    d = imageDriver(win, { until, wait });
    k = toolsDriver(d, { until, wait });
    js = d.js;
    mouse = pointer(win, wait);
    await until(async () => (await d.layers()) === 2, 'the project open again', 15000).catch(() => {});
    await d.settled('the project drawn', 6000);
    await k.tool('move');
    await d.pickRow(id);
    const reopened = await frameOn();
    await shoot(win, wait, 'rotate-frame-image-reopened.png');
    const reopenedClicks = await clicks();
    check('rotate-frame-image: saved as a project and opened again, the layer keeps its angle and the frame is still on it',
      wrote && rotInFile && Math.round(reopened.rot || 0) === 135 && reopened.ok && reopenedClicks.ok, JSON.stringify({ wrote, rotInFile, reopened, reopenedClicks }));

    const complaints = await errorsIn(win);
    check('rotate-frame-image: the turned-frame checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('rotate-frame-image: the turned-frame checks ran', false, err.stack || err.message);
  } finally {
    // No dialog answer is left for the next block.
    if (Array.isArray(globalThis.__rutbaCheckDialogAnswers)) globalThis.__rutbaCheckDialogAnswers.length = 0;
  }
}
