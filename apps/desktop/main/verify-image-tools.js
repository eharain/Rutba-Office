// Image: the advanced editor's tools (docs/STUDIO.md 7.2, I10 to I14),
// pressed in the window and read back from the canvas and the project saved.
// A picture of flat colours (an orange ground, a blue and a green block, and
// a white sign seen at an angle) is worked with the tools down the left: a
// rectangle selection limits Greyscale to its inside; the magic wand picks
// the blue block; a brush stroke lands on a new paint layer and Undo and Redo
// take it back and put it back; the clone stamp copies a patch of blue; the
// bucket fills the blue block; a mask made from a selection, feathered and
// painted black hides part of the Background (read in the alpha); the project
// is saved as .rimg and opens again the same; free transform shrinks the paint
// layer by its corner; perspective straightens the sign into an upright
// rectangle; and the tools stand on the right in Urdu. Run alone with
// RUTBA_VERIFY_ONLY=image-tools.
import fs from 'node:fs';
import path from 'node:path';
import { nativeImage } from 'electron';
import { imageDriver } from './verify-image-layers.js';

const W = 600;
const H = 400;
const ORANGE = [230, 130, 50];
const BLUE = [40, 110, 220];
const GREEN = [60, 170, 80];
const WHITE = [245, 245, 245];
/** The sign: a white four-sided shape, taller at the left, as a rectangle looks from its left. */
const SIGN = [[360, 60], [560, 100], [560, 300], [360, 340]];

const inside = (pts, x, y) => {
  let on = false;
  for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
    const [xi, yi] = pts[i];
    const [xj, yj] = pts[j];
    if ((yi > y) !== (yj > y) && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) on = !on;
  }
  return on;
};

/** The fixture, as BGRA for nativeImage. */
export function toolsPicture() {
  const buf = Buffer.alloc(W * H * 4);
  for (let y = 0; y < H; y++) {
    for (let x = 0; x < W; x++) {
      let c = ORANGE;
      if (x >= 40 && x < 240 && y >= 40 && y < 180) c = BLUE;
      else if (x >= 40 && x < 240 && y >= 220 && y < 360) c = GREEN;
      else if (inside(SIGN, x + 0.5, y + 0.5)) c = WHITE;
      const i = (y * W + x) * 4;
      buf[i] = c[2];
      buf[i + 1] = c[1];
      buf[i + 2] = c[0];
      buf[i + 3] = 255;
    }
  }
  return nativeImage.createFromBitmap(buf, { width: W, height: H });
}

const near = (a, b, tol = 14) => b.every((v, i) => Math.abs(a[i] - v) <= tol);
const grey = (p, tol = 8) => Math.max(p[0], p[1], p[2]) - Math.min(p[0], p[1], p[2]) <= tol;

/** Helpers over the image driver: the tools, pointer gestures on the tool stage, pixels by canvas position. */
export function toolsDriver(d, { until, wait }) {
  const { js } = d;
  const size = async () => {
    const r = await d.result();
    const [w, h] = String(r).split('x').map(Number);
    return { w: w || W, h: h || H };
  };
  const tool = async (id) => {
    await until(() => js(`Boolean(document.querySelector('.im-tool[data-tool-id="${id}"]:not(:disabled)'))`), `the ${id} tool`, 4000).catch(() => {});
    await js(`document.querySelector('.im-tool[data-tool-id="${id}"]')?.click(), 1`);
    return until(() => js(`document.querySelector('.im-tool[data-tool-id="${id}"]')?.getAttribute('aria-pressed') === 'true' && (${JSON.stringify(id)} === 'move' || Boolean(document.querySelector('.im-tool-stage[data-stage-tool="${id}"]')))`), `the ${id} tool in hand`, 4000).then(() => true).catch(() => false);
  };
  /** A press, moves and a release on the tool stage, at canvas pixels. */
  const gesture = async (points, { alt = false, pen = false } = {}) => {
    const { w } = await size();
    return js(`(() => {
      const el = document.querySelector('.im-tool-stage');
      if (!el) return 'no stage';
      const r = el.getBoundingClientRect();
      const k = r.width / ${w};
      const pts = ${JSON.stringify(points)};
      const at = (p) => ({ clientX: r.left + p[0] * k, clientY: r.top + p[1] * k, bubbles: true, cancelable: true, pointerId: 7, pointerType: ${JSON.stringify(pen ? 'pen' : 'mouse')}, pressure: 0.6, button: 0, buttons: 1, isPrimary: true, altKey: ${alt} });
      el.dispatchEvent(new PointerEvent('pointerdown', at(pts[0])));
      for (const p of pts.slice(1)) el.dispatchEvent(new PointerEvent('pointermove', at(p)));
      el.dispatchEvent(new PointerEvent('pointerup', at(pts[pts.length - 1])));
      return 'done';
    })()`);
  };
  /** The preview's average colour and alpha over a 5 × 5 patch at canvas pixel (x, y). */
  const px = async (x, y) => {
    const { w, h } = await size();
    return d.patchAt(x / w, y / h);
  };
  const selectionBox = () => js(`document.querySelector('.im-ants')?.dataset.selection || ''`);
  const rows = () => d.rows();
  const undo = () => js(`document.querySelector('.rw-quick .rw-btn[data-tip="Undo"]')?.click(), 1`);
  const redo = () => js(`document.querySelector('.rw-quick .rw-btn[data-tip="Redo"]')?.click(), 1`);
  const waitPx = (x, y, test, what, timeout = 6000) => until(async () => test(await px(x, y)), what, timeout).then(() => true).catch(() => false);
  return { size, tool, gesture, px, selectionBox, rows, undo, redo, waitPx, wait };
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyImageTools(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'image-tools');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'tools.png');
  fs.writeFileSync(file, toolsPicture().toPNG());
  const win = await open('image', file);
  const d = imageDriver(win, { until, wait });
  const k = toolsDriver(d, { until, wait });
  const { js, capture, settled, pickRow, press, setRange, setInput, tab } = d;
  await until(async () => (await d.result()) === `${W}x${H}` && (await js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`)) > 0, 'the picture on the canvas', 15000);

  // The layout: the tools in a column at the left of the picture, the options in a bar above it.
  const layout = await js(`(() => {
    const r = (s) => { const n = document.querySelector(s); if (!n) return null; const b = n.getBoundingClientRect(); return { left: b.left, top: b.top, right: b.right, bottom: b.bottom, width: b.width, height: b.height }; };
    return { tools: r('.im-tools'), options: r('.im-options'), stage: r('.im-stage'), count: document.querySelectorAll('.im-tool').length, tips: [...document.querySelectorAll('.im-tool')].map((b) => b.dataset.tip.split(':')[0]) };
  })()`);
  const picked = await k.tool('marquee');
  await capture('image-tools-1-layout.png');
  check('image tools: fourteen tools stand in a column at the left of the picture, the chosen tool\'s options in a bar above it',
    layout.count === 14 && layout.tools && layout.stage && layout.tools.right <= layout.stage.left + 1 && layout.options.bottom <= layout.stage.top + 1 && layout.tools.height > 300 && picked
      && (await d.layers()) >= 1,
    JSON.stringify(layout));

  // I11: a rectangle selection, then Greyscale changes only its inside, on a layer of its own.
  await k.gesture([[60, 240], [110, 290], [160, 330]]);
  const boxed = await until(async () => /^\d+,\d+,\d+,\d+$/.test(await k.selectionBox()), 'the marching ants', 5000).then(() => true).catch(() => false);
  const sel1 = (await k.selectionBox()).split(',').map(Number);
  const layersBefore = await d.layers();
  await tab('Filters');
  await d.click('Greyscale');
  const greyed = await k.waitPx(110, 285, (p) => grey(p), 'the inside greyed', 8000);
  await settled('the greyed picture');
  const inner = await k.px(110, 285);
  const outer = await k.px(200, 300);
  const ground = await k.px(300, 300);
  await capture('image-tools-2-selection-greyscale.png');
  check('image tools: a rectangle selection (marching ants) keeps Greyscale to its inside, on a layer of its own; outside it the picture keeps its colour',
    boxed && Math.abs(sel1[0] - 60) <= 2 && Math.abs(sel1[1] - 240) <= 2 && Math.abs(sel1[2] - 100) <= 3 && Math.abs(sel1[3] - 90) <= 3
      && greyed && grey(inner) && near(outer, GREEN, 6) && near(ground, ORANGE, 6) && (await d.layers()) === layersBefore + 1,
    JSON.stringify({ sel1, inner, outer, ground, layersBefore, layers: await d.layers() }));

  // I11: the magic wand picks the flat blue block, edge to edge.
  await k.tool('wand');
  await k.gesture([[100, 100]]);
  const wand = await until(async () => (await k.selectionBox()) === '40,40,200,140', 'the wand\'s selection', 6000).then(() => true).catch(() => false);
  await settled('the ants drawn', 3000);
  await capture('image-tools-3-wand.png');
  check('image tools: the magic wand picks the flat blue block exactly, 200 × 140 at 40, 40', wand, await k.selectionBox());

  // The Modify menu in the options bar: Grow by 4 px widens the selection by 4 on each side.
  await js(`document.querySelector('[data-role="modify-selection"]')?.click(), 1`);
  await until(() => js(`document.querySelectorAll('.rw-menu button').length === 3`), 'the Modify menu', 3000).catch(() => {});
  const menuItems = await js(`[...document.querySelectorAll('.rw-menu button')].map((b) => b.textContent.trim())`);
  await capture('image-tools-3b-modify.png');
  await js(`[...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === 'Grow')?.click(), 1`);
  const grown = await until(async () => (await k.selectionBox()) === '36,36,208,148', 'the selection grown', 5000).then(() => true).catch(() => false);
  check('image tools: Modify, Grow by 4 px in the options bar widens the wand\'s selection by 4 on every side; the menu holds Feather, Grow and Shrink',
    grown && JSON.stringify(menuItems) === JSON.stringify(['Feather', 'Grow', 'Shrink']) && !(await js(`Boolean(document.querySelector('.rw-menu'))`)),
    JSON.stringify({ menuItems, box: await k.selectionBox() }));

  // I12: a brush stroke lands on a new paint layer; Undo takes it off and Redo puts it back.
  await js(`document.querySelector('[data-role="deselect"]')?.click(), 1`);
  await until(async () => (await k.selectionBox()) === '', 'no selection', 3000).catch(() => {});
  await pickRow('background');
  await k.tool('brush');
  await setRange('[data-role="opt-size"]', 30);
  await k.gesture([[300, 20], [430, 20], [560, 20]]);
  const painted = await k.waitPx(430, 20, (p) => near(p, [0xe0, 0x39, 0x3e, 255], 16), 'the brush stroke', 6000);
  const paintRows = await k.rows();
  const paintId = paintRows.find((r) => /^paint-/.test(r)) || null;
  await capture('image-tools-4-brush.png');
  await k.undo();
  const undone = await k.waitPx(430, 20, (p) => near(p, ORANGE, 10), 'the stroke undone', 5000);
  await k.redo();
  const redone = await k.waitPx(430, 20, (p) => near(p, [0xe0, 0x39, 0x3e, 255], 16), 'the stroke again', 5000);
  check('image tools: a brush stroke paints red on a new paint layer above the Background; Undo takes the stroke off and Redo puts it back',
    painted && Boolean(paintId) && paintRows.indexOf(paintId) < paintRows.indexOf('background') && undone && redone && near(await k.px(430, 60), ORANGE, 8),
    JSON.stringify({ painted, paintRows, undone, redone, at: await k.px(430, 20) }));

  // I13: the clone stamp, Alt+clicked in the blue block, paints blue where it is dragged.
  await k.tool('clone');
  await k.gesture([[140, 110]], { alt: true });
  const sourced = await until(() => js(`/Copying from 140, 110/.test(document.querySelector('[data-role="clone-source"]')?.textContent || '')`), 'the clone source', 3000).then(() => true).catch(() => false);
  await k.gesture([[290, 200], [310, 200], [330, 200]]);
  const cloned = await k.waitPx(310, 200, (p) => near(p, BLUE, 16), 'the cloned patch', 6000);
  await capture('image-tools-5-clone.png');
  check('image tools: the clone stamp copies a patch of the blue block to where it is dragged, the same distance from where it was Alt+clicked',
    sourced && cloned && near(await k.px(310, 240), ORANGE, 8),
    JSON.stringify({ sourced, cloned, at: await k.px(310, 200), below: await k.px(310, 240) }));

  // I12: the bucket fills the blue block with the colour, and nothing round it.
  await k.tool('bucket');
  await setInput('[data-role="colour"]', '#7b49d6');
  await wait(150);
  await k.gesture([[100, 100]]);
  const filled = await k.waitPx(100, 100, (p) => near(p, [0x7b, 0x49, 0xd6, 255], 10), 'the bucket fill', 6000);
  const edge = await k.px(232, 172);
  await capture('image-tools-6-fill.png');
  check('image tools: the bucket fills the blue block with the chosen colour, to its edge and no further',
    filled && near(edge, [0x7b, 0x49, 0xd6, 255], 12) && near(await k.px(250, 100), ORANGE, 8) && near(await k.px(310, 200), BLUE, 16),
    JSON.stringify({ filled, centre: await k.px(100, 100), edge, outside: await k.px(250, 100) }));

  // I10: a mask from a selection, feathered, then painted black: part of the Background hidden (read in the alpha).
  await k.tool('marquee');
  await k.gesture([[300, 0], [450, 200], [600, 400]]);
  await until(async () => /^300,0,/.test(await k.selectionBox()), 'the right half selected', 4000).catch(() => {});
  await pickRow('background');
  await until(() => js(`Boolean(document.querySelector('[data-role="mask-selection"]:not(:disabled)'))`), 'From Selection', 4000).catch(() => {});
  await press('[data-role="mask-selection"]');
  const masked = await k.waitPx(200, 300, (p) => p[3] === 0, 'the left of the Background hidden', 6000);
  await js(`document.querySelector('[data-role="deselect"]')?.click(), 1`);
  await js(`document.querySelector('.im-layer[data-layer="background"] [data-role="target-mask"]')?.click(), 1`);
  await k.tool('brush');
  await setInput('[data-role="colour"]', '#000000');
  await setRange('[data-role="opt-size"]', 20);
  await wait(150);
  const note = await js(`document.querySelector('[data-role="paint-target"]')?.textContent || ''`);
  await k.gesture([[380, 380], [470, 380], [560, 380]]);
  const hidden = await k.waitPx(470, 380, (p) => p[3] === 0, 'the mask painted black', 6000);
  await pickRow('background');
  await until(() => js(`Boolean(document.querySelector('[data-role="mask-feather"]'))`), 'the mask\'s feather', 4000).catch(() => {});
  await setRange('[data-role="mask-feather"]', 40);
  const feathered = await until(async () => { const a = (await k.px(300, 300))[3]; return a > 40 && a < 215; }, 'the mask feathered', 6000).then(() => true).catch(() => false);
  const ramp = [await k.px(270, 300), await k.px(300, 300), await k.px(330, 300)].map((p) => p[3]);
  await capture('image-tools-7-mask.png');
  check('image tools: a mask made from a selection hides the Background outside it, black painted on the mask hides more, and a feather of 40 px ramps its edge',
    masked && feathered && ramp[0] < ramp[1] && ramp[1] < ramp[2] && hidden && (await k.px(450, 200))[3] === 255 && (await k.px(598, 200))[3] === 255
      && /mask/i.test(note) && (await js(`Boolean(document.querySelector('.im-layer[data-layer="background"] canvas.im-layer-thumb.mask'))`)),
    JSON.stringify({ masked, feathered, ramp, hidden, note, at: await k.px(470, 380) }));

  // A .rimg with a mask, a paint layer and an adjusted selection saved, and opened again the same.
  const project = path.join(root, 'tools.rimg');
  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: project });
  await k.tool('move');
  await tab('Home');
  await press('[data-role="save-project"]');
  const wrote = await until(() => fs.existsSync(project) && fs.statSync(project).size > 1000, 'the project file', 8000).then(() => true).catch(() => false);
  const spots = [[430, 20], [310, 200], [100, 100], [110, 285], [200, 300], [470, 380], [450, 200], [300, 300]];
  const before = { rows: await k.rows(), px: [] };
  for (const [x, y] of spots) before.px.push(await k.px(x, y));
  const zip = wrote ? fs.readFileSync(project).toString('latin1') : '';
  const again = await open('image', project);
  const d2 = imageDriver(again, { until, wait });
  const k2 = toolsDriver(d2, { until, wait });
  await until(async () => (await d2.layers()) === before.rows.length && (await d2.result()) === `${W}x${H}`, 'the project open again', 15000).catch(() => {});
  await until(async () => near(await k2.px(470, 380), before.px[5], 4) && near(await k2.px(430, 20), before.px[0], 4), 'the project drawn again', 10000).catch(() => {});
  const after = { rows: await k2.rows(), px: [] };
  for (const [x, y] of spots) after.px.push(await k2.px(x, y));
  await d2.capture('image-tools-8-reopened.png');
  check('image tools: a .rimg with a mask, a paint layer and an adjusted selection opens again with every layer, the mask, and the same pixels',
    wrote && /bitmaps\/background-mask-[^/]*\.png/.test(zip) && /bitmaps\/paint-1\.png/.test(zip) && JSON.stringify(after.rows) === JSON.stringify(before.rows)
      && after.px.every((p, i) => near(p, before.px[i], 6))
      && (await d2.js(`Boolean(document.querySelector('.im-layer[data-layer="background"] canvas.im-layer-thumb.mask'))`)),
    JSON.stringify({ wrote, before, after }));

  // I14: free transform, the paint layer's bottom-right corner dragged to the middle: the layer at half its size.
  await d2.pickRow(paintId);
  await k2.tool('transform');
  await until(() => d2.js(`Boolean(document.querySelector('[data-role="transform-quad"]'))`), 'the corner handles', 4000).catch(() => {});
  await k2.gesture([[600, 400], [450, 300], [300, 200]]);
  const shrunk = await k2.waitPx(215, 10, (p) => near(p, [0xe0, 0x39, 0x3e, 255], 24), 'the stroke at half size', 6000);
  const gone = near(await k2.px(500, 20), ORANGE, 10);
  await d2.capture('image-tools-9-transform.png');
  check('image tools: free transform scales the paint layer from its corner: dragged to the middle, its red stroke is drawn at half its size, half as far along',
    shrunk && gone, JSON.stringify({ shrunk, half: await k2.px(215, 10), was: await k2.px(500, 20) }));
  await k2.undo();
  await k2.waitPx(500, 20, (p) => near(p, [0xe0, 0x39, 0x3e, 255], 16), 'the transform undone', 5000);

  // I14: perspective, four pins on the sign's corners, straightened and cropped: the sign fills an upright rectangle.
  await d2.pickRow('background');
  await k2.tool('perspective');
  await until(() => d2.js(`document.querySelectorAll('.im-pin').length === 4`), 'the pins', 4000).catch(() => {});
  const pinAt = (i) => d2.js(`(() => { const c = document.querySelector('.im-pin[data-pin="${i}"]'); const el = document.querySelector('.im-tool-stage'); if (!c || !el) return null; const k = el.getBoundingClientRect().width / ${W}; return [Number(c.getAttribute('cx')) / k, Number(c.getAttribute('cy')) / k]; })()`);
  for (let i = 0; i < 4; i++) {
    const from = await pinAt(i);
    if (from) await k2.gesture([from, [(from[0] + SIGN[i][0]) / 2, (from[1] + SIGN[i][1]) / 2], SIGN[i]]);
    await wait(80);
  }
  const placed = [];
  for (let i = 0; i < 4; i++) placed.push((await pinAt(i))?.map(Math.round));
  await d2.capture('image-tools-10-pins.png');
  await d2.press('[data-role="straighten-crop"]');
  const cropped = await until(async () => /^20\dx24\d$/.test(await d2.result()), 'the straightened canvas', 6000).then(() => true).catch(() => false);
  await d2.settled('the straightened sign', 6000);
  const result = await d2.result();
  const [cw, ch] = result.split('x').map(Number);
  const putDown = await until(() => d2.js(`document.querySelector('.im-tool[data-tool-id="move"]')?.getAttribute('aria-pressed') === 'true' && !document.querySelector('.im-pin')`), 'the tool put down', 3000).then(() => true).catch(() => false);
  const corners = [];
  for (const [fx, fy] of [[0.07, 0.07], [0.93, 0.07], [0.93, 0.93], [0.07, 0.93], [0.5, 0.5]]) corners.push(await d2.patchAt(fx, fy));
  await d2.capture('image-tools-11-straightened.png');
  check('image tools: perspective maps the skewed sign to an upright rectangle: pins on its corners, Straighten and Crop leaves a 204 × 240 canvas white to its corners and puts the tool down',
    cropped && putDown && Math.abs(cw - 204) <= 2 && Math.abs(ch - 240) <= 2 && corners.every((p) => near(p, [...WHITE, 255], 14)) && placed.every((p, i) => p && Math.abs(p[0] - SIGN[i][0]) <= 2 && Math.abs(p[1] - SIGN[i][1]) <= 2),
    JSON.stringify({ result, putDown, placed, corners }));

  const errs = [...(await errorsIn(win)), ...(await errorsIn(again))];
  check('image tools: no errors in the windows', errs.length === 0, errs.slice(0, 3).join(' | '));
  again.destroy();

  // In Urdu: the tools on the right of the picture, the options bar in Urdu.
  const was = await js(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`);
  try {
    await js(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
    const ur = await open('image', project);
    const du = imageDriver(ur, { until, wait });
    const ku = toolsDriver(du, { until, wait });
    await until(async () => (await du.layers()) === before.rows.length, 'the project in Urdu', 15000).catch(() => {});
    await du.pickRow('background');
    await ku.tool('wand');
    await wait(300);
    const frame = await du.js(`(() => {
      const r = (s) => document.querySelector(s)?.getBoundingClientRect();
      const tools = r('.im-tools');
      const stage = r('.im-stage');
      return {
        dir: document.documentElement.dir,
        right: tools && stage ? tools.left >= stage.right - 1 : false,
        label: document.querySelector('.im-options-tool span')?.textContent || '',
        fits: (() => { const b = document.querySelector('.im-options'); return b ? b.scrollWidth <= b.clientWidth + 1 : false; })(),
        words: [...document.querySelectorAll('.im-options .rw-btn, .im-opt-seg button, .im-opt-label, .im-mask-props .im-prop-label')].map((n) => n.textContent.trim()).filter(Boolean),
      };
    })()`);
    await du.capture('image-tools-12-urdu.png');
    const english = frame.words.filter((w) => /^(New|Add|Subtract|Intersect|Tolerance|Modify|By|Feather|Grow|Shrink|Copy to Layer|Fill|Delete|Mask)$/.test(w));
    check('image tools: in Urdu the tools stand at the right of the picture and the options bar speaks Urdu, on one row that fits',
      frame.dir === 'rtl' && frame.right && frame.fits && frame.label && frame.label !== 'Magic Wand' && english.length === 0 && frame.words.length > 5,
      JSON.stringify({ ...frame, english }));
    ur.destroy();
  } finally {
    await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(was ?? null)} }).then(() => 1)`);
  }

  // The rest of the tools, on the picture afresh: the ellipse, the lasso and a gradient inside it, select
  // by colour, the eraser through a mask, and the healing brush over a blemish.
  const more = await open('image', file);
  const d3 = imageDriver(more, { until, wait });
  const k3 = toolsDriver(d3, { until, wait });
  await until(async () => (await d3.result()) === `${W}x${H}` && (await d3.js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`)) > 0, 'the picture again', 15000).catch(() => {});

  await k3.tool('ellipse');
  await k3.gesture([[40, 40], [140, 110], [240, 180]]);
  const oval = await until(async () => /^\d+,\d+,\d+,\d+$/.test(await k3.selectionBox()), 'the ellipse', 4000).then(() => true).catch(() => false);
  const ovalBox = (await k3.selectionBox()).split(',').map(Number);
  await k3.tool('colour');
  await k3.gesture([[300, 380]]);
  const everywhere = await until(async () => (await k3.selectionBox()) === '0,0,600,400', 'every orange pixel', 6000).then(() => true).catch(() => false);
  check('image tools: Ellipse Select draws an oval as wide as the drag; Select by Colour takes the orange in every part of the picture',
    oval && Math.abs(ovalBox[0] - 40) <= 2 && Math.abs(ovalBox[2] - 200) <= 3 && Math.abs(ovalBox[3] - 140) <= 3 && everywhere,
    JSON.stringify({ ovalBox, all: await k3.selectionBox() }));

  // The lasso round the green block, then a red gradient fading to clear across it, cut by the lasso.
  await k3.tool('lasso');
  await k3.gesture([[30, 210], [250, 210], [250, 370], [30, 370], [30, 212]]);
  await until(async () => /^3[01],21[01],/.test(await k3.selectionBox()), 'the lasso', 4000).catch(() => {});
  const lassoBox = await k3.selectionBox();
  await k3.tool('gradient');
  await d3.setInput('[data-role="colour"]', '#e0393e');
  await wait(120);
  await k3.gesture([[40, 290], [140, 290], [240, 290]]);
  const graded = await k3.waitPx(60, 290, (p) => p[0] > 190 && p[1] < 110, 'the gradient', 8000);
  const fade = [await k3.px(60, 290), await k3.px(140, 290), await k3.px(232, 290)];
  const past = await k3.px(300, 290);
  await d3.capture('image-tools-13-lasso-gradient.png');
  check('image tools: a gradient dragged across a lasso selection runs from red to clear over the green block and stops at the lasso\'s edge',
    graded && /^3[01],21[01],2[12]\d,1[56]\d$/.test(lassoBox) && fade[0][0] > fade[1][0] && fade[1][0] > fade[2][0] && near(fade[2], GREEN, 20) && near(past, ORANGE, 6),
    JSON.stringify({ lassoBox, fade, past }));
  await k3.tool('marquee');
  await d3.js(`document.querySelector('[data-role="deselect"]')?.click(), 1`);

  // The eraser on the Background: a mask is made for it, and the stroke hides what it passes over.
  await until(async () => (await k3.selectionBox()) === '', 'no selection', 3000).catch(() => {});
  await d3.pickRow('background');
  await k3.tool('eraser');
  await d3.setRange('[data-role="opt-size"]', 24);
  await k3.gesture([[300, 380], [450, 380], [590, 380]]);
  const erased = await k3.waitPx(450, 380, (p) => p[3] === 0, 'the Background rubbed out', 6000);
  const maskMade = await d3.js(`Boolean(document.querySelector('.im-layer[data-layer="background"] canvas.im-layer-thumb.mask'))`);
  check('image tools: the eraser on the Background hides what it passes over through a mask made for it, the quick edits kept',
    erased && maskMade && (await k3.px(450, 340))[3] === 255, JSON.stringify({ erased, maskMade, at: await k3.px(450, 380) }));

  // A dark blemish painted on the orange, then healed from clean orange below it.
  await d3.js(`document.querySelector('.im-layer[data-layer="background"] [data-role="target-pixels"]')?.click(), 1`);
  await k3.tool('brush');
  await d3.setInput('[data-role="colour"]', '#202020');
  await d3.setRange('[data-role="opt-size"]', 10);
  await wait(120);
  await k3.gesture([[300, 250], [302, 251]]);
  const spotted = await k3.waitPx(300, 250, (p) => p[0] < 120, 'the blemish', 5000);
  await k3.tool('heal');
  await d3.setRange('[data-role="opt-size"]', 36);
  await k3.gesture([[300, 300]], { alt: true });
  await until(() => d3.js(`/Copying from 300, 300/.test(document.querySelector('[data-role="clone-source"]')?.textContent || '')`), 'the heal source', 3000).catch(() => {});
  await k3.gesture([[290, 250], [300, 250], [310, 250]]);
  const healed = await k3.waitPx(300, 250, (p) => near(p, ORANGE, 28), 'the blemish healed', 8000);
  await d3.capture('image-tools-14-heal.png');
  check('image tools: the healing brush covers a dark blemish with the clean orange it was Alt+clicked on, matched to the orange round it',
    spotted && healed, JSON.stringify({ spotted, healed, at: await k3.px(300, 250), round: await k3.px(330, 250) }));

  const errs2 = await errorsIn(more);
  check('image tools: no errors in the window of the other tools', errs2.length === 0, errs2.slice(0, 3).join(' | '));
  more.destroy();
  // The first window stays: destroying the last window quits the app under the checks.
}
