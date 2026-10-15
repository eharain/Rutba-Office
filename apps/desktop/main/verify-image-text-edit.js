// Image: typing on the canvas (docs/STUDIO.md 7.2, I8).
//
// Over the picture of flat colours the tools check uses: Insert, Text puts a
// text layer on the picture and starts typing in it at once, its words
// selected, in a text box standing exactly over the layer (the type bar
// above the picture, no dialog). Words typed replace them; the caret moves
// with End, a selection is made with Shift+Home and typed over, and a second
// line grows the box. The type bar's size and font and Ctrl+B restyle the
// layer while the caret stays in the words. Escape finishes and the engine
// draws the words on the picture where the box was; Undo takes the styles
// back a step each and the whole spell of typing back in one, Redo puts it
// all back. A double-click on the words opens them again with the caret at
// their end; Urdu typed there runs right to left. A new text whose words are
// all taken out goes when typing ends. Run alone with
// RUTBA_VERIFY_ONLY=image-text-edit.

import fs from 'node:fs';
import path from 'node:path';
import { imageDriver } from './verify-image-layers.js';
import { toolsPicture } from './verify-image-tools.js';

const W = 600;
const H = 400;

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args the run's scratch folder
 */
export async function verifyImageTextEdit(h, { dir }) {
  const { open, check, until, wait, errorsIn } = h;
  const root = path.join(dir, 'image-text-edit');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'flat.png');
  fs.writeFileSync(file, toolsPicture().toPNG());
  const win = await open('image', file);
  const wc = win.webContents;
  const d = imageDriver(win, { until, wait });
  const { js } = d;
  await until(async () => (await d.result()) === `${W}x${H}` && (await js(`Number(document.querySelector('.im-canvas')?.dataset.drawn || 0)`)) > 0, 'the picture on the canvas', 15000);

  const key = async (keyCode, modifiers = []) => {
    wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
    wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
    await wait(60);
  };
  const box = () => js(`(() => {
    const b = document.querySelector('.im-type-box');
    if (!b) return null;
    const r = b.getBoundingClientRect();
    const c = document.querySelector('.im-canvas').getBoundingClientRect();
    const cs = getComputedStyle(b);
    return { value: b.value, start: b.selectionStart, end: b.selectionEnd, focused: document.activeElement === b,
      x: (r.left + r.width / 2 - c.left) / c.width, y: (r.top + r.height / 2 - c.top) / c.height, h: r.height, w: r.width,
      font: cs.fontFamily, size: parseFloat(cs.fontSize), weight: cs.fontWeight, k: c.width / ${W}, rtl: b.matches(':dir(rtl)'), layer: b.dataset.layer };
  })()`);
  const panelText = () => js(`document.querySelector('[data-role="layer-text"]')?.value ?? null`);

  // Insert, Text: typing at once, the words selected, the type bar above the picture.
  await d.tab('Insert');
  await d.press('[data-insert="text"]');
  const opened = await until(async () => { const b = await box(); return b && b.focused && b.start === 0 && b.end === b.value.length && b.value.length > 0; }, 'the text box with its words selected', 6000).then(() => true).catch(() => false);
  const b0 = await box();
  const bar = await js(`(() => { const b = document.querySelector('[data-role="type-bar"]'); return b ? [...b.querySelectorAll('[data-role]')].map((n) => n.dataset.role) : null; })()`);
  const dialog = await js(`Boolean(document.querySelector('.rw-dialog'))`);
  await d.capture('image-text-edit-1-new.png');
  check('image text edit: Insert, Text starts typing on the picture at once: a text box over the new layer, its words selected, the type bar above, no dialog',
    opened && !dialog && Array.isArray(bar) && ['type-font', 'type-size', 'type-bold', 'type-italic', 'type-colour', 'type-done'].every((r) => bar.includes(r)) && Math.abs(b0.x - 0.5) < 0.03 && Math.abs(b0.y - 0.5) < 0.05,
    JSON.stringify({ box: b0, bar, dialog }));
  const id = b0?.layer;

  // Typed over; the caret moved with End; Shift+Home selects the line and typing replaces it; a second line grows the box.
  wc.insertText('Sale today');
  await until(async () => (await box())?.value === 'Sale today', 'the words typed', 4000).catch(() => {});
  await key('End');
  wc.insertText(' only');
  const appended = await until(async () => (await box())?.value === 'Sale today only', 'the words added at the end', 4000).then(() => true).catch(() => false);
  await key('Home', ['shift']);
  const selected = await until(async () => { const b = await box(); return b && b.start === 0 && b.end === b.value.length; }, 'the line selected', 3000).then(() => true).catch(() => false);
  wc.insertText('Big news');
  const oneLine = await until(async () => (await box())?.value === 'Big news', 'the selection typed over', 4000).then(() => true).catch(() => false);
  const h1 = (await box())?.h || 0;
  wc.insertText('\nLine two');
  const twoLines = await until(async () => { const b = await box(); return b?.value === 'Big news\nLine two' && b.h > h1 * 1.6; }, 'the second line', 5000).then(() => true).catch(() => false);
  const b1 = await box();
  const inPanel = await until(async () => (await panelText()) === 'Big news\nLine two', 'the layer\'s words', 4000).then(() => true).catch(() => false);
  check('image text edit: words typed on the picture go into the layer: typed over the selection, the caret moved with End, Shift+Home selects the line to type over, and a second line grows the box',
    appended && selected && oneLine && twoLines && inPanel, JSON.stringify({ appended, selected, oneLine, twoLines, inPanel, h1, box: b1 }));

  // The type bar: size and font; Ctrl+B in the words; the caret stays in them.
  await js(`(() => { const i = document.querySelector('[data-role="type-size"]'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(i, '64'); i.dispatchEvent(new Event('input', { bubbles: true })); return 1; })()`);
  const sized = await until(async () => { const b = await box(); return b && Math.abs(b.size - 64 * b.k) < 1.5; }, 'the words at 64 px', 4000).then(() => true).catch(() => false);
  await js(`document.querySelector('[data-role="type-size"]')?.blur(), 1`);
  await d.setSelect('[data-role="type-font"]', 'Georgia');
  const faced = await until(async () => /Georgia/.test((await box())?.font || ''), 'the words in Georgia', 4000).then(() => true).catch(() => false);
  await until(async () => (await box())?.focused, 'the focus back in the words', 3000).catch(() => {});
  await key('B', ['control']);
  const bolded = await until(async () => { const b = await box(); return b && Number(b.weight) >= 700 && (await js(`document.querySelector('[data-role="type-bold"]')?.getAttribute('aria-pressed')`)) === 'true'; }, 'bold', 4000).then(() => true).catch(() => false);
  const b2 = await box();
  await d.capture('image-text-edit-2-styled.png');
  check('image text edit: the type bar sets the size and the font and Ctrl+B makes the words bold, the box drawn so at once and the caret still in the words',
    sized && faced && bolded && b2?.focused && b2.value === 'Big news\nLine two', JSON.stringify({ sized, faced, bolded, box: b2 }));

  // Escape: the words drawn on the picture by the engine where the box was.
  const area = (b) => ({ x: Math.round(W * b.x - (b.w / b.k) / 2), y: Math.round(H * b.y - (b.h / b.k) / 2), w: Math.round(b.w / b.k), h: Math.round(b.h / b.k) });
  const where = area(b2);
  await key('Escape');
  const closed = await until(async () => !(await box()), 'typing finished', 4000).then(() => true).catch(() => false);
  await d.settled('the words drawn', 6000);
  const inked = await js(`(() => {
    const c = document.querySelector('.im-canvas');
    const k = c.width / ${W};
    const r = ${JSON.stringify(where)};
    const g = c.getContext('2d');
    const px = g.getImageData(Math.max(0, r.x * k), Math.max(0, r.y * k), Math.max(1, r.w * k), Math.max(1, r.h * k)).data;
    let white = 0;
    for (let i = 0; i < px.length; i += 4) if (px[i] > 235 && px[i + 1] > 235 && px[i + 2] > 235) white++;
    return { white, of: px.length / 4 };
  })()`);
  await d.capture('image-text-edit-3-drawn.png');
  check('image text edit: Escape finishes, and the engine draws the words on the picture where the box stood',
    closed && inked.white > inked.of * 0.04, JSON.stringify({ closed, where, inked }));

  // Undo: bold, the font and the size a step each, then the whole spell of typing in one; Redo puts it all back.
  await d.pickRow(id);
  const undo = () => js(`document.querySelector('.rw-quick .rw-btn[data-tip="Undo"]')?.click(), 1`);
  const redo = () => js(`document.querySelector('.rw-quick .rw-btn[data-tip="Redo"]')?.click(), 1`);
  for (let i = 0; i < 3; i++) { await undo(); await wait(250); }
  const stylesOff = await until(async () => (await panelText()) === 'Big news\nLine two', 'the words kept, the styles gone', 4000).then(() => true).catch(() => false);
  await undo();
  const typingOff = await until(async () => /^Your words/.test((await panelText()) || ''), 'the typing undone in one step', 4000).then(() => true).catch(() => false);
  for (let i = 0; i < 4; i++) { await redo(); await wait(250); }
  const back = await until(async () => (await panelText()) === 'Big news\nLine two', 'the typing redone', 4000).then(() => true).catch(() => false);
  check('image text edit: Undo takes the bold, the font and the size back a step each and the whole spell of typing back in one; Redo puts them back',
    stylesOff && typingOff && back, JSON.stringify({ stylesOff, typingOff, back, now: await panelText() }));

  // A double-click on the words opens them again, the caret at their end; Urdu typed there runs right to left.
  await d.settled('the words drawn again', 4000);
  await js(`(() => {
    const s = document.querySelector('.im-stage-layers');
    const c = document.querySelector('.im-canvas').getBoundingClientRect();
    const at = { clientX: c.left + c.width * ${b2.x}, clientY: c.top + c.height * ${b2.y}, bubbles: true, cancelable: true, button: 0 };
    s.dispatchEvent(new MouseEvent('mousedown', { ...at, detail: 1 }));
    window.dispatchEvent(new MouseEvent('mouseup', { ...at, detail: 1 }));
    s.dispatchEvent(new MouseEvent('mousedown', { ...at, detail: 2 }));
    window.dispatchEvent(new MouseEvent('mouseup', { ...at, detail: 2 }));
    return 1;
  })()`);
  const again = await until(async () => { const b = await box(); return b && b.focused && b.start === b.value.length && b.end === b.value.length && b.value === 'Big news\nLine two'; }, 'the words open again', 5000).then(() => true).catch(() => false);
  await js(`(() => { const b = document.querySelector('.im-type-box'); b.setSelectionRange(0, b.value.length); return 1; })()`);
  wc.insertText('آج سیل ہے');
  const urdu = await until(async () => { const b = await box(); return b?.value === 'آج سیل ہے' && b.rtl; }, 'Urdu right to left', 4000).then(() => true).catch(() => false);
  await d.capture('image-text-edit-4-urdu.png');
  await key('Escape');
  await until(async () => !(await box()), 'typing finished again', 4000).catch(() => {});
  check('image text edit: a double-click on the words opens them again with the caret at their end, and Urdu typed there runs right to left',
    again && urdu, JSON.stringify({ again, urdu, panel: await panelText() }));

  // A new text with every word taken out goes when typing ends.
  const layersBefore = await d.layers();
  await d.tab('Insert');
  await d.press('[data-insert="text"]');
  await until(async () => (await box())?.focused, 'a second text box', 5000).catch(() => {});
  await key('Delete');
  await until(async () => (await box())?.value === '', 'its words out', 3000).catch(() => {});
  await key('Escape');
  const gone = await until(async () => !(await box()) && (await d.layers()) === layersBefore, 'the empty text gone', 5000).then(() => true).catch(() => false);
  check('image text edit: a new text whose words are all taken out goes when typing ends', gone, JSON.stringify({ layersBefore, layers: await d.layers() }));

  const complaints = await errorsIn(win);
  check('image text edit: none of that reported an error', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
}
