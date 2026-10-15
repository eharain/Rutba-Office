// Image: selections kept in the .rimg project (docs/STUDIO.md 7.2, I11 and
// I17).
//
// The picture of flat colours the tools check uses. A rectangle selected on
// the green block and Greyscale pressed make an Adjusted selection layer;
// the project is saved with that selection still on the canvas and opened
// again in a window of its own: the marching ants are back where they were,
// and Invert on the Filters tab changes the same layer, inside the
// selection only, with no new layer. That selection let go, an ellipse
// selected on the orange ground and the project saved again: opened once
// more, the ellipse is on the canvas, and the Adjusted selection layer's
// Load Selection (in the Layers panel) puts its own rectangle back and
// Invert pressed again changes that layer back. Run alone with
// RUTBA_VERIFY_ONLY=image-selection-save.

import fs from 'node:fs';
import path from 'node:path';
import { imageDriver } from './verify-image-layers.js';
import { toolsPicture, toolsDriver } from './verify-image-tools.js';

const W = 600;
const H = 400;
const ORANGE = [230, 130, 50];
const GREEN = [60, 170, 80];
const near = (a, b, tol = 14) => b.every((v, i) => Math.abs(a[i] - v) <= tol);
const grey = (p, tol = 8) => Math.max(p[0], p[1], p[2]) - Math.min(p[0], p[1], p[2]) <= tol;
const boxNear = (got, want, tol = 3) => {
  const g = String(got || '').split(',').map(Number);
  return g.length === 4 && want.every((v, i) => Math.abs(g[i] - v) <= tol);
};

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyImageSelectionSave(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'image-selection-save');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'flat.png');
  fs.writeFileSync(file, toolsPicture().toPNG());
  const win = await open('image', file);
  const d = imageDriver(win, { until, wait });
  const k = toolsDriver(d, { until, wait });
  await until(async () => (await d.result()) === `${W}x${H}` && (await d.js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`)) > 0, 'the picture on the canvas', 15000);

  // A rectangle on the green block, greyed on a layer of its own.
  await k.tool('marquee');
  await k.gesture([[60, 240], [110, 290], [160, 330]]);
  await until(async () => boxNear(await k.selectionBox(), [60, 240, 100, 90]), 'the rectangle selected', 5000).catch(() => {});
  await d.tab('Filters');
  await d.click('Greyscale');
  const greyed = await k.waitPx(110, 285, (p) => grey(p), 'the inside greyed', 8000);
  await d.settled('the greyed picture');
  const rows1 = await k.rows();
  const adjusted = rows1.find((r) => /^adjusted-/.test(r)) || null;
  check('image selection save: a rectangle selected and greyed makes an Adjusted selection layer, the selection still on the canvas',
    greyed && Boolean(adjusted) && boxNear(await k.selectionBox(), [60, 240, 100, 90]), JSON.stringify({ rows: rows1, box: await k.selectionBox() }));

  // Saved with the selection on the canvas.
  const project = path.join(root, 'kept.rimg');
  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: project });
  await k.tool('move');
  await d.tab('Home');
  await d.press('[data-role="save-project"]');
  const wrote = await until(() => fs.existsSync(project) && fs.statSync(project).size > 1000, 'the project file', 8000).then(() => true).catch(() => false);
  const zip = wrote ? fs.readFileSync(project).toString('latin1') : '';
  check('image selection save: the .rimg keeps the selection on the canvas and the layer\'s own, with the picture it adjusts',
    wrote && zip.includes('selections/selection.rsel') && zip.includes(`selections/${adjusted}.rsel`) && zip.includes(`bitmaps/${adjusted}-under.png`) && zip.includes('"scoped": "' + adjusted + '"'),
    wrote ? `${fs.statSync(project).size} bytes; ${(zip.match(/selections\/[^\0"]+?\.rsel/g) || []).filter((v, i, a) => a.indexOf(v) === i).join(', ')}` : 'not written');

  // Opened again: the ants where they were, and the Filters tab changes the same layer.
  const again = await open('image', project);
  const d2 = imageDriver(again, { until, wait });
  const k2 = toolsDriver(d2, { until, wait });
  await until(async () => (await d2.layers()) === rows1.length && (await d2.result()) === `${W}x${H}`, 'the project open again', 15000).catch(() => {});
  const antsBack = await until(async () => boxNear(await k2.selectionBox(), [60, 240, 100, 90]) && (await d2.js(`Number(document.querySelector('.im-ants')?.dataset.outlined || 0)`)) > 0, 'the ants back', 8000).then(() => true).catch(() => false);
  await k2.waitPx(110, 285, (p) => grey(p), 'the grey drawn again', 8000);
  await d2.capture('image-selection-save-1-reopened.png');
  check('image selection save: opened again, the selection is back on the canvas, its marching ants where they were',
    antsBack, JSON.stringify({ box: await k2.selectionBox(), rows: await k2.rows() }));

  const before = await k2.px(110, 285);
  await d2.tab('Filters');
  await d2.click('Invert');
  const inverted = await k2.waitPx(110, 285, (p) => grey(p) && Math.abs(p[0] - before[0]) > 12, 'the inside inverted', 8000);
  await d2.settled('the inverted picture');
  const rows2 = await k2.rows();
  const outside = await k2.px(200, 300);
  const ground = await k2.px(300, 300);
  await d2.capture('image-selection-save-2-changed-again.png');
  check('image selection save: Invert changes the same Adjusted selection layer, inside the selection only, with no new layer',
    inverted && JSON.stringify(rows2) === JSON.stringify(rows1) && near(outside, GREEN, 8) && near(ground, ORANGE, 8),
    JSON.stringify({ before, inside: await k2.px(110, 285), outside, ground, rows: rows2 }));

  // Another selection on the canvas, saved again.
  await d2.js(`document.querySelector('[data-role="deselect"]')?.click(), 1`);
  await until(async () => (await k2.selectionBox()) === '', 'no selection', 3000).catch(() => {});
  await k2.tool('ellipse');
  await k2.gesture([[300, 100], [360, 140], [420, 180]]);
  const ellipse = await until(async () => boxNear(await k2.selectionBox(), [300, 100, 120, 80], 3), 'the ellipse selected', 5000).then(() => true).catch(() => false);
  await k2.tool('move');
  await d2.tab('Home');
  const savedAt = fs.statSync(project).mtimeMs;
  await d2.press('[data-role="save-project"]');
  await until(() => fs.statSync(project).mtimeMs !== savedAt, 'the project saved again', 8000).catch(() => {});
  await wait(400);

  const third = await open('image', project);
  const d3 = imageDriver(third, { until, wait });
  const k3 = toolsDriver(d3, { until, wait });
  await until(async () => (await d3.layers()) === rows1.length, 'the project open a third time', 15000).catch(() => {});
  const ellipseBack = await until(async () => boxNear(await k3.selectionBox(), [300, 100, 120, 80], 3), 'the ellipse back', 8000).then(() => true).catch(() => false);
  check('image selection save: saved with another selection, it is that one that comes back on the canvas',
    ellipse && ellipseBack, JSON.stringify({ box: await k3.selectionBox() }));

  // The layer's own selection, loaded from the Layers panel.
  await d3.pickRow(adjusted);
  await until(() => d3.js(`Boolean(document.querySelector('[data-role="load-selection"]:not(:disabled)'))`), 'Load Selection', 5000).catch(() => {});
  const button = await d3.js(`(() => { const b = document.querySelector('[data-role="load-selection"]'); return b ? { label: b.textContent.trim(), tip: b.dataset.tip || b.title || '' } : null; })()`);
  await d3.capture('image-selection-save-3-load-selection.png');
  await d3.press('[data-role="load-selection"]');
  const loaded = await until(async () => boxNear(await k3.selectionBox(), [60, 240, 100, 90]), 'the layer\'s selection loaded', 5000).then(() => true).catch(() => false);
  const marked = await until(() => d3.js(`document.querySelector('[data-role="load-selection"]')?.disabled === true`), 'the button marked', 3000).then(() => true).catch(() => false);
  const inverted3 = await k3.px(110, 285);
  await d3.tab('Filters');
  await d3.click('Invert');
  const sepia = await k3.waitPx(110, 285, (p) => grey(p) && p[0] - inverted3[0] > 12, 'the inside back to its grey', 8000);
  await d3.settled('the picture changed again');
  const rows3 = await k3.rows();
  await d3.capture('image-selection-save-4-loaded-and-changed.png');
  check('image selection save: the Adjusted selection layer\'s Load Selection puts its own rectangle back, and Invert pressed again changes that layer back',
    Boolean(button && /^Load Selection/.test(button.tip)) && loaded && marked && sepia && JSON.stringify(rows3) === JSON.stringify(rows1) && near(await k3.px(200, 300), GREEN, 8),
    JSON.stringify({ button, box: await k3.selectionBox(), was: inverted3, inside: await k3.px(110, 285), rows: rows3 }));

  const complaints = [...(await errorsIn(win)), ...(await errorsIn(again)), ...(await errorsIn(third))];
  check('image selection save: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
}
