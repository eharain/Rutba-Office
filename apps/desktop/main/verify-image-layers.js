// Image: the layered picture (docs/STUDIO.md 7.2, I7 to I9, I15 to I17),
// pressed in the window and read back from the canvas and the files saved.
// A picture of grain over a gradient takes a text and a shape as layers; the
// shape is sent to the back and dragged up the Layers panel again, hidden and
// shown, faded and blended, each proved by the pixels under it; an
// adjustment layer turns grey only what is under it; the canvas grows from a
// corner and is trimmed back; the project is saved as .rimg, copied for the
// launcher's Recover while unsaved, and opens again as it was; and the window
// reads right to left in Urdu. Edit in Image from Documents, Worksheets and
// Presentations puts the edited picture back where it was.
// Run alone with RUTBA_VERIFY_ONLY=image-layers.
import fs from 'node:fs';
import path from 'node:path';
import { BrowserWindow } from 'electron';
import { buildDocx, buildXlsx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { SheetView } from '@rutba/sheet-view';
import { Deck, buildPptx } from '@rutba/presentation';
import { noisyPicture } from './verify-image-edits.js';
import { gradientPng } from './sample-picture.js';

/** Helpers for one Image window. */
function driver(win, { until, wait }) {
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const capture = async (name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    wc.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
  };
  const result = () => js(`document.querySelector('.im-canvas-wrap')?.dataset.result || ''`);
  const layers = () => js(`Number(document.querySelector('.im-canvas-wrap')?.dataset.layers || 0)`);
  const rows = () => js(`[...document.querySelectorAll('.im-layer')].map((n) => n.dataset.layer)`);
  const drawnTurn = () => js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`);
  const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 1`);
  const click = async (label, scope = '.rw-groups') => {
    const find = `[...document.querySelectorAll(${JSON.stringify(`${scope} .rw-btn`)})].find((n) => (n.dataset.tip || n.textContent || '').trim().startsWith(${JSON.stringify(label)}) && !n.disabled)`;
    await until(() => js(`Boolean(${find})`), `the ${label} button`, 4000).catch(() => {});
    return js(`(() => { const b = ${find}; if (!b) return 'no button ' + ${JSON.stringify(label)}; b.click(); return 'clicked'; })()`);
  };
  const press = (selector) => js(`(() => { const b = document.querySelector(${JSON.stringify(selector)}); if (!b || b.disabled) return 'no ' + ${JSON.stringify(selector)}; b.click(); return 'clicked'; })()`);
  const setRange = (selector, value) => js(`(() => { const i = document.querySelector(${JSON.stringify(selector)}); if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(String(value))}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
  const setSelect = (selector, value) => js(`(() => { const s = document.querySelector(${JSON.stringify(selector)}); if (!s) return false; Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, ${JSON.stringify(String(value))}); s.dispatchEvent(new Event('change', { bubbles: true })); return true; })()`);
  const setInput = (selector, value) => js(`(() => { const i = document.querySelector(${JSON.stringify(selector)}); if (!i) return false; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, ${JSON.stringify(String(value))}); i.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
  // The average of a 7 × 7 patch of the preview, with its alpha, at a fraction of its width and height.
  const patchAt = (fx, fy) => js(`(() => { const c = document.querySelector('.im-canvas'); const x = Math.max(0, Math.min(c.width - 7, Math.floor(c.width * ${fx}) - 3)); const y = Math.max(0, Math.min(c.height - 7, Math.floor(c.height * ${fy}) - 3)); const d = c.getContext('2d').getImageData(x, y, 7, 7).data; const s = [0, 0, 0, 0]; for (let i = 0; i < d.length; i += 4) { s[0] += d[i]; s[1] += d[i + 1]; s[2] += d[i + 2]; s[3] += d[i + 3]; } return s.map((v) => Math.round(v / 49)); })()`);
  const settled = async (what, timeout = 8000) => {
    const before = await drawnTurn();
    await until(async () => (await drawnTurn()) > before, what, timeout).catch(() => {});
    let last = await drawnTurn();
    for (let i = 0; i < 10; i++) {
      await wait(160);
      const now = await drawnTurn();
      if (now === last) break;
      last = now;
    }
  };
  /** Pick a row in the Layers panel. */
  const pickRow = (id) => js(`(() => { const r = document.querySelector('.im-layer[data-layer=${JSON.stringify(id)}]'); if (!r) return 'no row'; r.click(); return 'picked'; })()`);
  /** Drag a row of the Layers panel onto another, above or below its middle. */
  const dragRow = (id, onto, where = 'above') => js(`(() => {
    const from = document.querySelector('.im-layer[data-layer=${JSON.stringify(id)}]');
    const to = document.querySelector('.im-layer[data-layer=${JSON.stringify(onto)}]');
    if (!from || !to) return 'no rows';
    const a = from.getBoundingClientRect();
    const b = to.getBoundingClientRect();
    const y = ${JSON.stringify(where)} === 'above' ? b.top + b.height * 0.2 : b.top + b.height * 0.8;
    const at = (cy) => ({ clientX: a.left + 60, clientY: cy, bubbles: true, button: 0 });
    from.dispatchEvent(new MouseEvent('mousedown', at(a.top + a.height / 2)));
    window.dispatchEvent(new MouseEvent('mousemove', at(a.top + a.height / 2 + 6)));
    window.dispatchEvent(new MouseEvent('mousemove', at(y)));
    window.dispatchEvent(new MouseEvent('mouseup', at(y)));
    return 'dragged';
  })()`);
  return { wc, js, capture, result, layers, rows, tab, click, press, setRange, setSelect, setInput, patchAt, settled, pickRow, dragRow };
}

const near = (a, b, tol = 14) => a.every((v, i) => Math.abs(v - b[i]) <= tol);
const grey = (p, tol = 6) => Math.max(p[0], p[1], p[2]) - Math.min(p[0], p[1], p[2]) <= tol;
const BLUE = [0x2f, 0x7d, 0xe1];

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyImageLayers(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc } = h;
  const root = path.join(dir, 'image-layers');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'scene.png');
  fs.writeFileSync(file, noisyPicture(900, 600, 23).toPNG());
  const win = await open('image', file);
  const d = driver(win, { until, wait });
  const { js, capture, result, layers, rows, tab, click, press, setRange, setSelect, setInput, patchAt, settled, pickRow, dragRow } = d;
  await until(async () => (await result()) === '900x600' && (await js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`)) > 0, 'the picture on the canvas', 15000);
  const photoCentre = await patchAt(0.5, 0.5);
  const photoCorner = await patchAt(0.1, 0.12);

  // I8: a text as a layer, and the picture turned into layers under it.
  await tab('Insert');
  await press('[data-insert="text"]');
  await until(async () => (await layers()) === 2, 'the text layer', 5000).catch(() => {});
  await settled('the text drawn', 8000);
  const afterText = await rows();
  const textId = afterText[0];
  check('image layers: Insert, Text puts words on a layer of their own over the picture, which becomes the Background layer',
    afterText.length === 2 && afterText[1] === 'background' && /^textblock-/.test(textId || ''),
    JSON.stringify({ rows: afterText }));

  // I8: a shape from the gallery, drawn by the engine in its tile.
  await press('[data-insert="shape"]');
  await until(() => js(`document.querySelectorAll('.im-gallery-item').length > 20`), 'the shapes gallery', 4000).catch(() => {});
  const tiles = await js(`[...document.querySelectorAll('.im-gallery-item canvas')].slice(0, 6).map((c) => { const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; let n = 0; for (let i = 3; i < d.length; i += 4) if (d[i] > 0) n++; return n; })`);
  await capture('image-layers-1-shapes.png');
  await press('.im-gallery-item[data-item="rect"]');
  await until(async () => (await layers()) === 3, 'the shape layer', 5000).catch(() => {});
  await settled('the shape drawn', 8000);
  const shapeId = (await rows())[0];
  const withShape = await patchAt(0.5, 0.5);
  check('image layers: Insert, Shapes offers tiles the engine draws, and a rectangle lands as a layer on top, in its colour over the picture',
    tiles.length === 6 && tiles.every((n) => n > 20) && /^shape-/.test(shapeId || '') && near(withShape, [...BLUE, 255], 10),
    JSON.stringify({ tiles, shapeId, withShape, photoCentre }));
  await capture('image-layers-2-panel.png');

  // I7: order. Send to Back puts the shape under the picture; dragged up the panel, it is on top again.
  await tab('Layers');
  await pickRow(shapeId);
  await click('Send to Back');
  await until(async () => (await rows()).at(-1) === shapeId, 'the shape at the back', 4000).catch(() => {});
  await settled('the shape behind the picture');
  const behind = await patchAt(0.5, 0.5);
  const order1 = await rows();
  await dragRow(shapeId, textId, 'above');
  await until(async () => (await rows())[0] === shapeId, 'the shape dragged to the top', 4000).catch(() => {});
  await settled('the shape on top again');
  const onTop = await patchAt(0.5, 0.5);
  check('image layers: Send to Back hides the shape under the picture; dragged above the text in the panel, it shows again',
    order1.at(-1) === shapeId && near(behind, photoCentre, 12) && (await rows())[0] === shapeId && near(onTop, [...BLUE, 255], 10),
    JSON.stringify({ order1, behind, onTop, now: await rows() }));

  // I7: hide and show.
  await js(`document.querySelector('.im-layer[data-layer="${shapeId}"] [data-role="eye"]')?.click(), 1`);
  await settled('the shape hidden');
  const hidden = await patchAt(0.5, 0.5);
  await js(`document.querySelector('.im-layer[data-layer="${shapeId}"] [data-role="eye"]')?.click(), 1`);
  await settled('the shape shown');
  const shown = await patchAt(0.5, 0.5);
  check('image layers: the eye hides a layer and shows it again', near(hidden, photoCentre, 12) && near(shown, [...BLUE, 255], 10), JSON.stringify({ hidden, shown }));

  // I7: opacity, then a blend mode, read in the pixels.
  await pickRow(shapeId);
  await until(() => js(`Boolean(document.querySelector('[data-role="opacity"]'))`), 'the opacity slider', 3000).catch(() => {});
  await setRange('[data-role="opacity"]', 50);
  await settled('the shape at half');
  const half = await patchAt(0.5, 0.5);
  const expectHalf = [0, 1, 2].map((i) => Math.round(photoCentre[i] * 0.5 + BLUE[i] * 0.5));
  await setRange('[data-role="opacity"]', 100);
  await settled('the shape whole');
  await setSelect('[data-role="blend"]', 'multiply');
  await settled('the shape multiplied');
  const multiplied = await patchAt(0.5, 0.5);
  const expectMul = [0, 1, 2].map((i) => Math.round((photoCentre[i] * BLUE[i]) / 255));
  await capture('image-layers-3-blend.png');
  check('image layers: opacity at 50% mixes the shape half and half with the picture under it, and Multiply darkens the picture by its colour',
    near(half.slice(0, 3), expectHalf, 10) && near(multiplied.slice(0, 3), expectMul, 10),
    JSON.stringify({ half, expectHalf, multiplied, expectMul }));
  await setSelect('[data-role="blend"]', 'normal');
  await settled('the shape normal');

  // I9: an adjustment layer greys everything under it; moved under the shape, the shape keeps its colour.
  await tab('Insert');
  await press('[data-insert="adjust"]');
  await until(() => js(`Boolean(document.querySelector('.im-gallery-item[data-item="mono"]'))`), 'the adjustments', 4000).catch(() => {});
  await press('.im-gallery-item[data-item="mono"]');
  await until(async () => (await layers()) === 4, 'the adjustment layer', 5000).catch(() => {});
  await settled('the picture grey', 10000);
  const adjustId = (await rows())[0];
  const greyCorner = await patchAt(0.1, 0.12);
  const greyShape = await patchAt(0.5, 0.5);
  await press('[data-role="layer-down"]');
  await until(async () => (await rows())[1] === adjustId, 'the adjustment under the shape', 4000).catch(() => {});
  await settled('the shape in colour again', 10000);
  const cornerUnder = await patchAt(0.1, 0.12);
  const shapeOver = await patchAt(0.5, 0.5);
  await capture('image-layers-4-adjustment.png');
  check('image layers: an adjustment layer (Mono) greys every layer under it and none above it',
    /^adjust-/.test(adjustId || '') && !grey(photoCorner, 20) && grey(greyCorner) && grey(greyShape) && grey(cornerUnder) && near(shapeOver, [...BLUE, 255], 10),
    JSON.stringify({ adjustId, photoCorner, greyCorner, greyShape, cornerUnder, shapeOver }));

  // I15: Canvas Size from the bottom-right anchor: the room added at the top and the left, clear.
  await tab('Home');
  await click('Canvas Size');
  await until(() => js(`Boolean(document.querySelector('[data-role="canvas-width"]'))`), 'the Canvas Size dialog', 4000).catch(() => {});
  await setInput('[data-role="canvas-width"]', 1100);
  await setInput('[data-role="canvas-height"]', 700);
  await js(`document.querySelector('[data-anchor="se"]')?.click(), 1`);
  await wait(150);
  await capture('image-layers-5-canvas.png');
  await press('[data-role="canvas-ok"]');
  await until(async () => (await result()) === '1100x700', 'the bigger canvas', 5000).catch(() => {});
  await settled('the bigger canvas drawn');
  const clearCorner = await patchAt(0.02, 0.03);
  const movedCorner = await patchAt((200 + 0.1 * 900) / 1100, (100 + 0.12 * 600) / 700);
  check('image layers: Canvas Size 1100 × 700 anchored bottom right adds clear room at the top and the left, the picture where it was in pixels',
    (await result()) === '1100x700' && clearCorner[3] === 0 && near(movedCorner.slice(0, 3), cornerUnder.slice(0, 3), 14),
    JSON.stringify({ result: await result(), clearCorner, movedCorner, cornerUnder }));

  // I15: Trim cuts the clear edges away again.
  await click('Trim');
  await until(async () => (await result()) === '900x600', 'the trimmed canvas', 5000).catch(() => {});
  await settled('the trimmed canvas drawn');
  const trimmedCorner = await patchAt(0.1, 0.12);
  check('image layers: Trim cuts the clear edges away, back to 900 × 600 with the picture where it was',
    (await result()) === '900x600' && near(trimmedCorner.slice(0, 3), cornerUnder.slice(0, 3), 14) && trimmedCorner[3] === 255,
    JSON.stringify({ result: await result(), trimmedCorner }));

  // I17: unsaved work is copied for the launcher's Recover.
  const copies = () => doc.recoverable().filter((e) => e.kind === 'image');
  const copied = await until(() => copies().length > 0, 'the recovery copy', 9000).then(() => true).catch(() => false);
  const copyEntry = copies()[0] || null;
  check('image layers: a layered picture with unsaved work is copied to the recovery folder as a .rimg the launcher offers back',
    copied && /^window-.+\.rimg$/.test(copyEntry?.file || ''), JSON.stringify(copyEntry));

  // I17: Save as .rimg, then open it again.
  const project = path.join(root, 'poster.rimg');
  globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: project });
  await press('[data-role="save-project"]');
  const wrote = await until(() => fs.existsSync(project) && fs.statSync(project).size > 1000, 'the project file', 8000).then(() => true).catch(() => false);
  const head = wrote ? fs.readFileSync(project).subarray(0, 100).toString('latin1') : '';
  const dropped = await until(() => copies().length === 0, 'the recovery copy to go', 5000).then(() => true).catch(() => false);
  const before = { rows: await rows(), centre: await patchAt(0.5, 0.5), corner: await patchAt(0.1, 0.12) };
  const again = await open('image', project);
  const d2 = driver(again, { until, wait });
  await until(async () => (await d2.layers()) === 4 && (await d2.result()) === '900x600', 'the project open again', 15000).catch(() => {});
  await until(async () => near((await d2.patchAt(0.1, 0.12)).slice(0, 3), before.corner.slice(0, 3), 10), 'the picture drawn again', 10000).catch(() => {});
  const after = { rows: await d2.rows(), centre: await d2.patchAt(0.5, 0.5), corner: await d2.patchAt(0.1, 0.12), title: again.getTitle() };
  await d2.capture('image-layers-6-reopened.png');
  check('image layers: Save writes a .rimg (a zip that names itself) and drops the recovery copy; it opens again with every layer, in order, drawn the same',
    wrote && head.startsWith('PK') && head.includes('application/vnd.rutba.image-project+zip') && dropped
      && JSON.stringify(after.rows) === JSON.stringify(before.rows) && near(after.centre, before.centre, 6) && near(after.corner, before.corner, 6),
    JSON.stringify({ wrote, dropped, before, after }));
  // The quick edits stay editable: the reopened picture still has its Background from the original.
  const quick = await d2.js(`document.querySelector('.im-layer[data-layer="background"]') ? 'background' : 'none'`);
  check('image layers: the reopened project still has the opened picture as its Background layer', quick === 'background', quick);

  // Undo takes back a step of the layered picture, Redo puts it back.
  await d2.pickRow(shapeId);
  await d2.press('[data-role="layer-delete"]');
  await until(async () => (await d2.layers()) === 3, 'the shape deleted', 4000).catch(() => {});
  await d2.js(`document.querySelector('.rw-quick .rw-btn[data-tip="Undo"]')?.click(), 1`);
  const undone = await until(async () => (await d2.layers()) === 4, 'the shape back', 4000).then(() => true).catch(() => false);
  await d2.js(`document.querySelector('.rw-quick .rw-btn[data-tip="Redo"]')?.click(), 1`);
  const redone = await until(async () => (await d2.layers()) === 3, 'the shape gone again', 4000).then(() => true).catch(() => false);
  check('image layers: Undo brings a deleted layer back and Redo deletes it again', undone && redone, JSON.stringify({ undone, redone }));

  const errs = [...(await errorsIn(win)), ...(await errorsIn(again))];
  check('image layers: no errors in the windows', errs.length === 0, errs.slice(0, 3).join(' | '));
  again.destroy();

  // In Urdu: the window's words in Urdu and its frame right to left, the layers panel with them.
  const was = await js(`window.rutbaOffice.store.get({ key: 'language', fallback: null })`);
  try {
    await js(`window.rutbaOffice.store.set({ key: 'language', value: 'ur' }).then(() => 1)`);
    const ur = await open('image', project);
    const du = driver(ur, { until, wait });
    await until(async () => (await du.layers()) === 4, 'the project in Urdu', 15000).catch(() => {});
    await du.pickRow(textId);
    await du.tab(await du.js(`[...document.querySelectorAll('.rw-tab')][2]?.textContent.trim() || ''`));
    await wait(300);
    const frame = await du.js(`({ dir: document.documentElement.dir, tabs: [...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim()), side: [...document.querySelectorAll('.im-side-tabs button')].map((b) => b.textContent.trim()), english: [...document.querySelectorAll('.rw-ribbon .rw-btn, .im-props .im-prop-label')].map((b) => b.textContent.trim()).filter((s) => /^(Bring Forward|Group|Opacity|Blend|Words)$/.test(s)) })`);
    await du.capture('image-layers-7-urdu.png');
    check('image layers: in Urdu the window, its Layers tab and its panel speak Urdu, right to left',
      frame.dir === 'rtl' && !frame.tabs.includes('Layers') && frame.side.length === 2 && !frame.side.includes('Layers') && frame.english.length === 0,
      JSON.stringify(frame));
    ur.destroy();
  } finally {
    await js(`window.rutbaOffice.store.set({ key: 'language', value: ${JSON.stringify(was ?? null)} }).then(() => 1)`);
  }
  // The first window stays: destroying the last window quits the app under the checks.
}

/**
 * Edit in Image from Documents, Worksheets and Presentations: Picture
 * Format → Edit in Image opens the picture in an Image window; an edit there
 * and Save and return put the edited picture back in the document, where it
 * was, as wide, its height from its new shape.
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc, sessionFor
 */
export async function verifyEditInImage(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc } = h;
  const root = path.join(dir, 'edit-in-image');
  fs.mkdirSync(root, { recursive: true });
  const source = async (id, op) => {
    const out = await doc.apply({ id, ops: [op] });
    return JSON.parse(out?.opResult || 'null');
  };
  const shot = async (win, name) => {
    if (!process.env.RUTBA_VERIFY_CAPTURE) return;
    win.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await win.webContents.capturePage()).toPNG());
  };
  /** The Image window a click opened, found as the window that was not there before. */
  const newImageWindow = async (before) => {
    let found = null;
    await until(() => (found = BrowserWindow.getAllWindows().find((w) => !before.has(w.id) && !w.isDestroyed())) != null, 'the Image window', 8000).catch(() => {});
    return found;
  };
  /** In the Image window: an edit, then Save and return. */
  const editThere = async (imageWin, name, edit) => {
    const di = driver(imageWin, { until, wait });
    await until(async () => (await di.result()) !== '' && (await di.js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`)) > 0, 'the picture in Image', 10000).catch(() => {});
    const title = imageWin.getTitle();
    const did = await edit(di);
    await di.settled('the edit drawn', 6000);
    await di.capture(name);
    await di.tab('Home');
    await until(() => di.js(`Boolean(document.querySelector('[data-role="put-back"]'))`), 'Save and return', 3000).catch(() => {});
    const pressed = await di.press('[data-role="put-back"]');
    return { title, did, pressed, errors: await errorsIn(imageWin) };
  };
  const formatTab = async (js) => {
    await until(() => js(`[...document.querySelectorAll('.rw-tab')].some((t) => t.textContent.trim() === 'Picture Format')`), 'the Picture Format tab', 4000).catch(() => {});
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Picture Format')?.click(), 1`);
    return until(() => js(`Boolean(document.querySelector('[data-role="edit-in-image"]:not(:disabled)'))`), 'Edit in Image', 4000).then(() => true).catch(() => false);
  };

  // Documents: the picture turned in Image, back twice as tall.
  {
    const file = path.join(root, 'report.docx');
    const view = openDocx(buildDocx({ paragraphs: [{ text: 'A picture to edit.' }, { text: '' }, { text: 'After it.' }] }));
    view.setSelection({ block: 1, offset: 0 });
    view.insertImage({ name: 'Logo', contentType: 'image/png', data: gradientPng(160, 80, [40, 120, 200], [230, 200, 60]), widthPx: 160, heightPx: 80 });
    fs.writeFileSync(file, view.save());
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = h.sessionFor('doc');
    const picture = () => {
      const blocks = doc.model({ id: session.id }).blocks || [];
      const b = blocks.findIndex((x) => x.images?.length);
      return b < 0 ? null : { block: b, ...blocks[b].images[0] };
    };
    await until(() => js(`Boolean(document.querySelector('.wd-image[data-drawing]'))`), 'the picture on the page', 8000).catch(() => {});
    const was = picture();
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await js(`document.querySelector('.wd-image[data-drawing]').click(), 1`);
    const ready = await formatTab(js);
    await shot(win, 'image-layers-8-word-picture-format.png');
    await js(`document.querySelector('[data-role="edit-in-image"]')?.click(), 1`);
    const imageWin = await newImageWindow(before);
    const there = imageWin ? await editThere(imageWin, 'image-layers-9-edit-from-word.png', async (di) => { await di.tab('Home'); return di.click('Right'); }) : null;
    const back = await until(() => { const p = picture(); return Boolean(p && was && p.href !== was.href); }, 'the picture back in the document', 10000).then(() => true).catch(() => false);
    const now = picture();
    check('edit in image: Documents, Picture Format → Edit in Image opens the picture in Image; turned there and saved, it is back in its place, as wide, twice as tall',
      ready && Boolean(imageWin) && back && /report\.docx/.test(there?.title || '') && Math.round(now.widthPx) === 160 && Math.abs(now.heightPx - 320) <= 2 && now.block === was.block,
      JSON.stringify({ ready, opened: Boolean(imageWin), there, back, was: was && { w: was.widthPx, h: was.heightPx, block: was.block }, now: now && { w: now.widthPx, h: now.heightPx, block: now.block } }));
    const errs = await errorsIn(win);
    check('edit in image: no errors in the Documents window', errs.length === 0, errs.slice(0, 3).join(' | '));
    imageWin?.destroy();
    // The document window stays open for app.exit (closing the last window quits the app).
  }

  // Worksheets: the picture flipped in Image, back on its anchor at its size.
  {
    const file = path.join(root, 'sales.xlsx');
    const view = new SheetView(buildXlsx({ sheets: [{ name: 'Sales', rows: [['Item', 'Qty'], ['Pens', 12], ['Ink', 3]] }] }));
    view.select(2, 4);
    view.insertPicture({ name: 'Logo', contentType: 'image/png', data: gradientPng(120, 60, [43, 95, 217], [15, 157, 88]), widthPx: 120, heightPx: 60 });
    fs.writeFileSync(file, view.save());
    const win = await open('sheets', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = h.sessionFor('sheet');
    await until(() => js(`Boolean(document.querySelector('.sh-drawing[data-name="Logo"]'))`), 'the picture on the sheet', 8000).catch(() => {});
    const drawing = () => (doc.model({ id: session.id }).drawings || []).find((d) => d.name === 'Logo');
    const was = drawing();
    const wasBytes = was ? (await source(session.id, { op: 'pictureSource', id: was.id }))?.data : null;
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await js(`(() => { const n = document.querySelector('.sh-drawing[data-name="Logo"]'); n.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); window.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, button: 0 })); return 1; })()`);
    const ready = await formatTab(js);
    await shot(win, 'image-layers-10-sheet-picture-format.png');
    await js(`document.querySelector('[data-role="edit-in-image"]')?.click(), 1`);
    const imageWin = await newImageWindow(before);
    const there = imageWin ? await editThere(imageWin, 'image-layers-11-edit-from-sheet.png', async (di) => { await di.tab('Home'); return di.click('Flip horizontal'); }) : null;
    let nowBytes = wasBytes;
    const back = await until(async () => { nowBytes = (await source(session.id, { op: 'pictureSource', id: was.id }))?.data; return Boolean(nowBytes && nowBytes !== wasBytes); }, 'the picture back on the sheet', 10000).then(() => true).catch(() => false);
    const now = drawing();
    check('edit in image: Worksheets, Picture Format → Edit in Image; flipped there and saved, the picture is back on its anchor, the same size',
      ready && Boolean(imageWin) && back && now && now.x === was.x && now.y === was.y && Math.abs(now.width - was.width) <= 1 && Math.abs(now.height - was.height) <= 1,
      JSON.stringify({ ready, opened: Boolean(imageWin), there, back, was: was && { x: was.x, y: was.y, w: was.width, h: was.height }, now: now && { x: now.x, y: now.y, w: now.width, h: now.height } }));
    imageWin?.destroy();
    // The document window stays open for app.exit (closing the last window quits the app).
  }

  // Presentations: the picture given a layer in Image, flattened on the way back.
  {
    const file = path.join(root, 'office.pptx');
    const deck = Deck.open(buildPptx({ slides: [{ title: 'Our office' }] }));
    const { id } = deck.addPicture(0, { data: gradientPng(200, 100, [200, 80, 60], [60, 80, 200]), contentType: 'image/png', name: 'Office', x: 300, y: 200, w: 400, h: 200 });
    fs.writeFileSync(file, deck.save());
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`Boolean(document.querySelector('.sl-hit[data-shape="${id}"]'))`), 'the picture on the slide', 8000).catch(() => {});
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const wasBytes = (await source(session.id, { op: 'pictureSource', slide: 0, shape: id }))?.data;
    const shape = () => { const m = doc.model({ id: session.id, slide: 0 }); return ((m.slide || m.slides?.[0])?.shapes || []).find((s) => String(s.id) === String(id) || s.name === 'Office'); };
    const was = shape();
    const before = new Set(BrowserWindow.getAllWindows().map((w) => w.id));
    await js(`document.querySelector('.sl-hit[data-shape="${id}"]')?.click(), 1`);
    const ready = await formatTab(js);
    await shot(win, 'image-layers-12-slide-picture-format.png');
    await js(`document.querySelector('[data-role="edit-in-image"]')?.click(), 1`);
    const imageWin = await newImageWindow(before);
    const there = imageWin ? await editThere(imageWin, 'image-layers-13-edit-from-slide.png', async (di) => {
      await di.tab('Insert');
      await di.press('[data-insert="shape"]');
      await until(() => di.js(`Boolean(document.querySelector('.im-gallery-item[data-item="ellipse"]'))`), 'the shapes', 4000).catch(() => {});
      return di.press('.im-gallery-item[data-item="ellipse"]');
    }) : null;
    let nowBytes = wasBytes;
    const back = await until(async () => { nowBytes = (await source(session.id, { op: 'pictureSource', slide: 0, shape: id }))?.data; return Boolean(nowBytes && nowBytes !== wasBytes); }, 'the picture back on the slide', 10000).then(() => true).catch(() => false);
    const now = shape();
    check('edit in image: Presentations, Picture Format → Edit in Image; a shape added on a layer there, saved, and the flattened picture is back where it was, the same size',
      ready && Boolean(imageWin) && back && now && JSON.stringify(now.geometry) === JSON.stringify(was.geometry),
      JSON.stringify({ ready, opened: Boolean(imageWin), there, back, was: was?.geometry, now: now?.geometry }));
    imageWin?.destroy();
    // The document window stays open for app.exit (closing the last window quits the app).
  }
}

export { driver as imageDriver };
