// Image: a mask follows its layer (docs/STUDIO.md 7.2, I10 with I14), as
// Photoshop links them.
//
// Over the picture of flat colours the tools check uses, a picture of two
// halves (red on the left, yellow on the right) goes in as a layer; its left
// half is selected and made its mask, so only the red shows. Then, by hand:
// the layer scaled by its corner grip, with the mask growing with it (red
// where the mask was not before), Undo and Redo taking both back and putting
// both back; turned a quarter by its rotate grip, the red half turned with
// it (red where the unturned mask would hide); distorted by a free-transform
// corner pulled out, the red reaching past where the mask stood; and the
// project saved and opened again the same. Run alone with
// RUTBA_VERIFY_ONLY=image-mask-transform.

import fs from 'node:fs';
import path from 'node:path';
import { nativeImage } from 'electron';
import { imageDriver } from './verify-image-layers.js';
import { toolsPicture, toolsDriver } from './verify-image-tools.js';

const W = 600;
const H = 400;
const RED = [220, 40, 40];
const YELLOW = [240, 210, 40];
const near = (a, b, tol = 16) => b.every((v, i) => Math.abs(a[i] - v) <= tol);

/** 200 × 200: red on the left half, yellow on the right, as BGRA for nativeImage. */
function halves() {
  const S = 200;
  const buf = Buffer.alloc(S * S * 4);
  for (let y = 0; y < S; y++) {
    for (let x = 0; x < S; x++) {
      const c = x < S / 2 ? RED : YELLOW;
      const i = (y * S + x) * 4;
      buf[i] = c[2];
      buf[i + 1] = c[1];
      buf[i + 2] = c[0];
      buf[i + 3] = 255;
    }
  }
  return nativeImage.createFromBitmap(buf, { width: S, height: S });
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyImageMaskTransform(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'image-mask-transform');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'ground.png');
  fs.writeFileSync(file, toolsPicture().toPNG());
  const halvesFile = path.join(root, 'halves.png');
  fs.writeFileSync(halvesFile, halves().toPNG());
  const win = await open('image', file);
  const d = imageDriver(win, { until, wait });
  const k = toolsDriver(d, { until, wait });
  const { js } = d;
  await until(async () => (await d.result()) === `${W}x${H}` && (await js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`)) > 0, 'the picture on the canvas', 15000);

  // The picture of two halves, as a layer.
  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [halvesFile] });
  await d.tab('Insert');
  await d.press('[data-insert="picture"]');
  await until(async () => (await d.layers()) === 2, 'the picture as a layer', 8000).catch(() => {});
  await d.settled('the picture drawn', 6000);
  const id = (await k.rows())[0];
  await k.tool('move');
  await d.pickRow(id);
  /** The chosen layer's box on the canvas (unturned), in canvas pixels, from its frame on the stage. */
  const boxOf = () => js(`(() => {
    const s = document.querySelector('.im-stage-layers');
    const b = document.querySelector('.im-sel[data-sel=${JSON.stringify(id)}]');
    if (!s || !b) return null;
    const r = s.getBoundingClientRect();
    const q = b.getBoundingClientRect();
    const k = ${W} / r.width;
    return { x: (q.left - r.left) * k, y: (q.top - r.top) * k, w: q.width * k, h: q.height * k };
  })()`);
  await until(async () => Boolean(await boxOf()), 'the layer\'s frame', 4000).catch(() => {});
  const b0 = await boxOf();
  if (!b0) {
    check('image mask transform: a picture goes in as a layer and can be chosen', false, JSON.stringify({ rows: await k.rows() }));
    return;
  }

  // Its left half selected and made its mask.
  await k.tool('marquee');
  await k.gesture([[b0.x, b0.y], [b0.x + b0.w / 4, b0.y + b0.h / 2], [b0.x + b0.w / 2, b0.y + b0.h]]);
  await until(async () => /^\d+,\d+,\d+,\d+$/.test(await k.selectionBox()), 'the half selected', 4000).catch(() => {});
  await d.pickRow(id);
  await until(() => js(`Boolean(document.querySelector('[data-role="mask-selection"]:not(:disabled)'))`), 'From Selection', 4000).catch(() => {});
  await d.press('[data-role="mask-selection"]');
  await until(() => js(`Boolean(document.querySelector('.im-layer[data-layer=${JSON.stringify(id)}] canvas.im-layer-thumb.mask'))`), 'the mask', 5000).catch(() => {});
  await js(`document.querySelector('[data-role="deselect"]')?.click(), 1`);
  await until(async () => (await k.selectionBox()) === '', 'no selection', 3000).catch(() => {});
  await d.settled('the mask drawn', 5000);
  const at = (b, fx, fy) => [b.x + b.w * fx, b.y + b.h * fy];
  const leftShows = await k.waitPx(...at(b0, 0.25, 0.5), (p) => near(p, [...RED, 255]), 'the red half', 6000);
  const rightHidden = !near(await k.px(...at(b0, 0.75, 0.5)), YELLOW, 30);
  check('image mask transform: a picture of two halves on a layer, its left half selected and made its mask: only the red half shows',
    leftShows && rightHidden, JSON.stringify({ box: b0, left: await k.px(...at(b0, 0.25, 0.5)), right: await k.px(...at(b0, 0.75, 0.5)) }));

  // Scaled by its corner grip: the mask grows with it.
  await k.tool('move');
  await d.pickRow(id);
  await until(() => js(`Boolean(document.querySelector('[data-grip="se"]'))`), 'the corner grip', 4000).catch(() => {});
  /** A drag on the stage from a grip (or a canvas point) to a canvas point, as the mouse makes it. */
  const dragStage = (from, to, { shift = false } = {}) => js(`(() => {
    const s = document.querySelector('.im-stage-layers');
    const r = s.getBoundingClientRect();
    const k = r.width / ${W};
    const g = typeof ${JSON.stringify(from)} === 'string' ? document.querySelector('[data-grip="' + ${JSON.stringify(from)} + '"]') : null;
    const gr = g ? g.getBoundingClientRect() : null;
    const start = gr ? { x: gr.left + gr.width / 2, y: gr.top + gr.height / 2 } : { x: r.left + ${JSON.stringify(from)}[0] * k, y: r.top + ${JSON.stringify(from)}[1] * k };
    const end = { x: r.left + ${JSON.stringify(to)}[0] * k, y: r.top + ${JSON.stringify(to)}[1] * k };
    const ev = (p, extra) => ({ clientX: p.x, clientY: p.y, bubbles: true, cancelable: true, button: 0, buttons: 1, detail: 1, shiftKey: ${shift}, ...extra });
    s.dispatchEvent(new MouseEvent('mousedown', ev(start)));
    for (let i = 1; i <= 8; i++) window.dispatchEvent(new MouseEvent('mousemove', ev({ x: start.x + (end.x - start.x) * i / 8, y: start.y + (end.y - start.y) * i / 8 })));
    window.dispatchEvent(new MouseEvent('mouseup', ev(end, { buttons: 0 })));
    return 'dragged';
  })()`);
  await dragStage('se', [b0.x + b0.w * 1.5, b0.y + b0.h * 1.5]);
  await until(async () => { const b = await boxOf(); return b && b.w > b0.w * 1.3; }, 'the layer scaled', 5000).catch(() => {});
  const b1 = await boxOf();
  // Red at 40 % across the bigger layer: past where the mask stood before, so only a mask that grew shows it.
  const grew = b1 ? await k.waitPx(...at(b1, 0.4, 0.85), (p) => near(p, [...RED, 255]), 'the red half grown', 6000) : false;
  const stillHidden = b1 ? !near(await k.px(...at(b1, 0.75, 0.5)), YELLOW, 30) : false;
  await d.capture('image-mask-transform-1-scaled.png');
  check('image mask transform: scaled by its corner grip, the layer takes its mask with it: red shows where the mask was not before, the yellow half still hidden',
    Boolean(b1) && grew && stillHidden, JSON.stringify({ was: b0, now: b1, grown: b1 && await k.px(...at(b1, 0.4, 0.85)), right: b1 && await k.px(...at(b1, 0.75, 0.5)) }));

  // Undo takes the scale and the mask's growing back together; Redo puts both back.
  await k.undo();
  const undone = await until(async () => { const b = await boxOf(); return b && Math.abs(b.w - b0.w) < 2; }, 'the scale undone', 5000).then(() => true).catch(() => false);
  const undoneHidden = await k.waitPx(...at(b0, 0.75, 0.5), (p) => !near(p, YELLOW, 30), 'the yellow hidden again', 4000);
  const undoneRed = await k.waitPx(...at(b0, 0.25, 0.5), (p) => near(p, [...RED, 255]), 'the red half as it was', 4000);
  await k.redo();
  const redone = await until(async () => { const b = await boxOf(); return b && Math.abs(b.w - b1.w) < 2; }, 'the scale again', 5000).then(() => true).catch(() => false);
  const redoneRed = b1 ? await k.waitPx(...at(b1, 0.4, 0.85), (p) => near(p, [...RED, 255]), 'the grown red again', 5000) : false;
  check('image mask transform: Undo takes the scale and the mask back together, Redo puts both back',
    undone && undoneHidden && undoneRed && redone && redoneRed, JSON.stringify({ undone, undoneHidden, undoneRed, redone, redoneRed }));

  // Turned a quarter by its rotate grip (Shift snaps the angle): the red half turns with it.
  await d.pickRow(id);
  await until(() => js(`Boolean(document.querySelector('[data-grip="rotate"]'))`), 'the rotate grip', 4000).catch(() => {});
  const c = [b1.x + b1.w / 2, b1.y + b1.h / 2];
  await dragStage('rotate', [c[0] + b1.w * 0.9, c[1]], { shift: true });
  await until(() => js(`/rotate\\((-?9\\d|-?8\\d)/.test(document.querySelector('.im-sel[data-sel=${JSON.stringify(id)}]')?.style.transform || '')`), 'the layer turned', 5000).catch(() => {});
  await d.settled('the turned layer', 5000);
  const turn = await js(`document.querySelector('.im-sel[data-sel=${JSON.stringify(id)}]')?.style.transform || ''`);
  // Turned either way a quarter, the red half lies across the top or the bottom; at 75 % across it is red only if the mask turned too.
  const top = await k.px(c[0] + b1.w * 0.25, c[1] - b1.h * 0.25);
  const bottom = await k.px(c[0] + b1.w * 0.25, c[1] + b1.h * 0.25);
  const turned = (near(top, [...RED, 255]) && !near(bottom, YELLOW, 30) && !near(bottom, RED, 30)) || (near(bottom, [...RED, 255]) && !near(top, YELLOW, 30) && !near(top, RED, 30));
  await d.capture('image-mask-transform-2-turned.png');
  check('image mask transform: turned a quarter by its rotate grip, the mask turns with the layer: the red half lies across one side, right across it',
    /rotate\(/.test(turn) && turned, JSON.stringify({ turn, top, bottom }));
  const redTop = near(top, [...RED, 255]);

  // Free transform, Distort: a corner of the red side pulled out past the layer; the red follows it out.
  await k.tool('transform');
  await until(() => js(`Boolean(document.querySelector('[data-role="transform-quad"]'))`), 'the corner handles', 4000).catch(() => {});
  await js(`document.querySelector('[data-opt="transform"] button[data-value="distort"]')?.click(), 1`);
  await until(() => js(`document.querySelector('[data-opt="transform"] button[data-value="distort"]')?.getAttribute('aria-checked') === 'true'`), 'Distort', 3000).catch(() => {});
  // The turned square's corner on the red side, at the right: top right when red is on top, bottom right otherwise.
  const corner = redTop ? [b1.x + b1.w, b1.y] : [b1.x + b1.w, b1.y + b1.h];
  const pulled = redTop ? [Math.min(W - 4, corner[0] + 80), Math.max(4, corner[1] - 40)] : [Math.min(W - 4, corner[0] + 80), Math.min(H - 4, corner[1] + 40)];
  await k.gesture([corner, [(corner[0] + pulled[0]) / 2, (corner[1] + pulled[1]) / 2], pulled]);
  // Just past the old right edge, near the pulled corner's side: inside the bent layer's red, outside where the mask stood.
  const beyond = redTop ? [corner[0] + 20, corner[1] + 4] : [corner[0] + 20, corner[1] - 4];
  const reached = await k.waitPx(...beyond, (p) => near(p, [...RED, 255], 20), 'the red past the old edge', 6000);
  await d.capture('image-mask-transform-3-distorted.png');
  check('image mask transform: a free-transform corner pulled out (Distort) takes the mask through the same map: the red reaches past where the mask stood',
    reached, JSON.stringify({ corner, pulled, beyond: await k.px(...beyond) }));
  await k.undo();
  const backIn = await k.waitPx(...beyond, (p) => !near(p, RED, 30), 'the distort undone', 5000);
  await k.redo();
  await k.waitPx(...beyond, (p) => near(p, [...RED, 255], 20), 'the distort again', 5000);
  check('image mask transform: Undo takes the distort and its mask back, Redo puts them back', backIn, JSON.stringify({ back: backIn }));

  // Saved and opened again: the mask where it was taken.
  const spots = [beyond, [c[0], c[1] - b1.h * 0.3], [c[0], c[1] + b1.h * 0.3], at(b1, 0.1, 0.1), [30, 30], [300, 380]];
  const before = [];
  for (const s of spots) before.push(await k.px(...s));
  const project = path.join(root, 'masked.rimg');
  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: project });
  await k.tool('move');
  await d.tab('Home');
  await d.press('[data-role="save-project"]');
  const wrote = await until(() => fs.existsSync(project) && fs.statSync(project).size > 1000, 'the project', 8000).then(() => true).catch(() => false);
  const again = await open('image', project);
  const d2 = imageDriver(again, { until, wait });
  const k2 = toolsDriver(d2, { until, wait });
  await until(async () => (await d2.layers()) === 2, 'the project open again', 15000).catch(() => {});
  await until(async () => near(await k2.px(...spots[0]), before[0], 6), 'the project drawn again', 10000).catch(() => {});
  const after = [];
  for (const s of spots) after.push(await k2.px(...s));
  await d2.capture('image-mask-transform-4-reopened.png');
  const zip = wrote ? fs.readFileSync(project).toString('latin1') : '';
  check('image mask transform: saved and opened again, the layer and its mask are where the transforms left them',
    wrote && /"corners"/.test(zip) && after.every((p, i) => near(p, before[i], 8)), JSON.stringify({ before, after }));

  const complaints = [...(await errorsIn(win)), ...(await errorsIn(again))];
  check('image mask transform: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
}
