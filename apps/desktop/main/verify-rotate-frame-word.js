// Documents: a turned drawing's frame turns with it.
//
// A page of words with a long picture in front of the text, a text box of
// words, a shape and a picture in line with the text. The picture is turned
// to 30 degrees by its rotate handle (Shift snapping to 15), to 90 by
// Layout → Rotate → More Rotation Options…, and to 135 by its rotate handle
// again. At each angle, at 100% and 150%, the frame's four corner handles
// stand on the picture's turned corners, a click at its centre picks it and
// a click in the empty corner of its old upright box does not. Undo and Redo
// bring the frame back with the angle, the arrows nudge it with its frame,
// and a corner drag resizes it along its own sides with the opposite corner
// staying put. The text box turned takes the caret in its words by a click;
// the shape and the picture in line with the text turn by Rotate Right 90°
// and keep their frames; right to left, the frame stays on. Saved and
// opened again, every angle is in the file (a:xfrm rot) and the frame still
// on the picture. Run alone with RUTBA_VERIFY_ONLY=rotate-frame-word.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx } from '@rutba/ooxml/build';
import { openDocx } from '@rutba/doc-view/backends/ooxml';
import { gradientPng } from './sample-picture.js';
import { pointer, worstGap, emptyCorner, shoot, centreOf, frameCorners } from './verify-rotate-frame.js';

const lorem = 'The northern region grew fastest in absolute terms, and the margin it opened in the spring held through the autumn. ';

/** The fixture: words, a long picture in front of them, a text box, a shape, and a picture in line. */
function makeFixture(file) {
  const view = openDocx(buildDocx({ styles: true, paragraphs: [{ text: 'Turned', style: 'Heading1' }, { text: lorem.repeat(5) }, { text: lorem.repeat(4) }, { text: 'In line: ' }, { text: lorem.repeat(2) }] }));
  // In line first, in the paragraph that says so: the picture and the shape.
  view.setSelection({ block: 3, offset: 9 });
  view.insertImage({ name: 'Inline', contentType: 'image/png', data: gradientPng(160, 60, [200, 60, 60], [250, 230, 120]), widthPx: 160, heightPx: 60 });
  view.insertShape({ preset: 'rightArrow', widthPx: 180, heightPx: 70, fill: '70AD47' });
  view.setSelection({ block: 1, offset: 0 });
  view.insertImage({ name: 'Banner', contentType: 'image/png', data: gradientPng(280, 80, [40, 120, 200], [230, 200, 60]), widthPx: 280, heightPx: 80 });
  const banner = view.drawings().find((d) => d.name === 'Banner').id;
  view.updateDrawings({ id: banner, wrap: 'none', behind: false, h: { rel: 'column', offsetPx: 170 }, v: { rel: 'paragraph', offsetPx: 60 } });
  const textBlock = view.drawings().find((d) => d.id === banner).block + 2;
  view.setSelection({ block: textBlock, offset: 0 });
  view.insertTextBox({ name: 'Note', widthPx: 220, heightPx: 70, fill: 'DEEBF7', h: { rel: 'column', offsetPx: 20 }, v: { rel: 'paragraph', offsetPx: 40 }, paragraphs: [{ text: 'Turned words' }] });
  fs.writeFileSync(file, view.save());
}

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn, doc, sessionFor
 * @param {{ dir: string }} args the run's scratch folder; the block writes in a folder of its own
 */
export async function verifyRotateFrameWord(h, { dir }) {
  const { open, check, until, wait, press, errorsIn, doc, sessionFor } = h;
  const root = path.join(dir, 'rotate-frame-word');
  fs.mkdirSync(root, { recursive: true });
  const file = path.join(root, 'turned.docx');
  try {
    makeFixture(file);
  } catch (err) {
    return check('rotate-frame-word: the fixture is written', false, err.stack || err.message);
  }
  let win = await open('word', file);
  try {
    let wc = win.webContents;
    let js = (code) => wc.executeJavaScript(code);
    let mouse = pointer(win, wait);
    const session = () => sessionFor('doc');
    const model = () => doc.model({ id: session().id });
    const drawing = (name) => (model().drawings || []).find((d) => d.name === name);
    const idOf = (name) => String(drawing(name)?.id);
    const sel = (name) => `.wd-drawing[data-drawing="${idOf(name)}"]`;
    const clickRibbon = (title) => js(`(() => {
      const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith(${JSON.stringify(title)}) && !n.disabled);
      if (!b) return 'no button ' + ${JSON.stringify(title)};
      b.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
      b.click();
      return 'clicked';
    })()`);
    const clickTab = (name) => js(`(() => { const t = [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)}); if (!t) return 'no tab'; t.click(); return 'tab'; })()`);
    const menuPick = async (button, label) => {
      await clickTab('Layout');
      await wait(150);
      const pressed = await clickRibbon(button);
      await until(() => js(`Boolean([...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)} && !b.disabled))`), `"${label}"`, 3000).catch(() => false);
      const picked = await js(`(() => { const b = [...document.querySelectorAll('.rw-menu button')].find((b) => b.textContent.trim() === ${JSON.stringify(label)}); if (!b || b.disabled) return 'missing'; b.click(); return 'clicked'; })()`);
      return pressed === 'clicked' && picked === 'clicked' ? 'clicked' : `${pressed}/${picked}`;
    };
    const pageZoom = () => js(`(() => { const p = document.querySelector('.wd-page'); return p ? p.getBoundingClientRect().width / p.offsetWidth : 0; })()`);
    const setZoom = async (z) => {
      await js(`(() => { const input = document.querySelector('.rw-zoom input[type="range"]'); if (!input) return 0; const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; set.call(input, ${JSON.stringify(String(z))}); input.dispatchEvent(new Event('input', { bubbles: true })); input.dispatchEvent(new Event('change', { bubbles: true })); return 1; })()`);
      return until(async () => Math.abs((await pageZoom()) - z) < 0.02, `the page at ${z * 100}%`, 4000).then(() => true).catch(() => false);
    };
    const bringIntoView = (name) => js(`(() => { document.querySelector(${JSON.stringify(sel(name))})?.scrollIntoView({ block: 'center', inline: 'center' }); return 1; })()`);
    /** The drawing's own upright box on screen, from what is drawn: its middle and its own size, scaled. */
    const boxOf = (name) => js(`(() => {
      const el = document.querySelector(${JSON.stringify(sel(name))});
      const p = document.querySelector('.wd-page');
      if (!el || !p) return null;
      const r = el.getBoundingClientRect();
      const z = p.getBoundingClientRect().width / p.offsetWidth;
      const w = el.offsetWidth * z; const h = el.offsetHeight * z;
      return { x: r.left + r.width / 2 - w / 2, y: r.top + r.height / 2 - h / 2, w, h };
    })()`);
    const centreOn = async (name) => { const b = await boxOf(name); return b ? { x: b.x + b.w / 2, y: b.y + b.h / 2 } : null; };
    const framed = () => js(`[...document.querySelectorAll('.wd-frame-sel')].map((f) => f.dataset.frame)`);
    const handles = () => js(`(() => ['nw', 'ne', 'se', 'sw'].map((n) => { const r = document.querySelector('.wd-frame-sel.main .wd-handle[data-handle="' + n + '"]')?.getBoundingClientRect(); return r && r.width ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null; }))()`);
    const letGo = async () => {
      await js(`document.querySelector('.wd-page')?.focus(), 1`);
      await press(wc, 'Escape');
      await until(async () => (await framed()).length === 0, 'nothing picked', 2000).catch(() => {});
    };
    const select = async (name) => {
      await bringIntoView(name);
      await wait(150);
      const at = await centreOn(name);
      if (at) await mouse.click(at);
      return until(async () => JSON.stringify(await framed()) === JSON.stringify([idOf(name)]), `${name} picked`, 3000).then(() => true).catch(() => false);
    };

    /** The frame's corner handles against the drawing's turned corners, waited for. */
    const frameOn = async (name) => {
      let last = null;
      const ok = await until(async () => {
        const b = await boxOf(name);
        const rot = drawing(name)?.rot || 0;
        if (!b) return false;
        const want = frameCorners(b, rot);
        const got = await handles();
        last = { rot, z: +(await pageZoom()).toFixed(2), corners: +worstGap(got, want).toFixed(1) };
        return last.corners <= 3;
      }, `the frame on ${name}`, 4000).then(() => true).catch(() => false);
      return { ok, ...last };
    };

    /** A click at the centre picks it; one in the old box's empty corner does not. */
    const clicks = async (name) => {
      await letGo();
      await bringIntoView(name);
      await wait(150);
      const b = await boxOf(name);
      const rot = drawing(name)?.rot || 0;
      await mouse.click({ x: b.x + b.w / 2, y: b.y + b.h / 2 });
      const centreTakes = await until(async () => (await framed()).includes(idOf(name)), 'the centre to pick', 3000).then(() => true).catch(() => false);
      await letGo();
      const corner = emptyCorner(b, rot, 5);
      const under = corner ? await js(`(() => { const n = document.elementFromPoint(${corner.x}, ${corner.y}); return n?.closest('.wd-drawing')?.dataset.drawing || 'words'; })()`) : 'no empty corner';
      if (corner) await mouse.click(corner);
      await wait(250);
      const cornerTakes = (await framed()).includes(idOf(name));
      return { ok: centreTakes && Boolean(corner) && !cornerTakes && under !== idOf(name), centreTakes, cornerTakes, under };
    };

    const atAngle = async (name, label) => {
      const results = [];
      for (const z of [1, 1.5]) {
        const zoomed = await setZoom(z);
        const held = await select(name);
        const frame = await frameOn(name);
        await shoot(win, wait, `rotate-frame-word-${label}${z === 1 ? '' : '-150'}.png`);
        const click = await clicks(name);
        results.push({ z, zoomed, held, frame, click });
      }
      await setZoom(1);
      return { ok: results.every((r) => r.zoomed && r.held && r.frame.ok && r.click.ok), results };
    };

    /** Drag the rotate handle round the drawing's centre to `deg`, Shift snapping to 15. */
    const turnTo = async (name, deg) => {
      await select(name);
      const centre = await centreOn(name);
      await until(() => js(centreOf('.wd-frame-sel.main .wd-rotate')), 'the rotate handle', 3000).catch(() => null);
      const grip = await js(centreOf('.wd-frame-sel.main .wd-rotate'));
      if (!grip || !centre) return false;
      const r = Math.hypot(grip.x - centre.x, grip.y - centre.y);
      const a = (deg * Math.PI) / 180;
      await mouse.drag(grip, { x: centre.x + r * Math.sin(a), y: centre.y - r * Math.cos(a) }, { modifiers: ['shift'], steps: 12 });
      return until(() => Math.round(drawing(name)?.rot || 0) === deg, `${name} at ${deg}`, 5000).then(() => true).catch(() => false);
    };

    await until(() => js(`document.querySelectorAll('.wd-drawing[data-drawing]').length >= 4`), 'the drawings', 10000);
    await setZoom(1);
    await wait(400);

    // 30, by the rotate handle.
    const turned30 = await turnTo('Banner', 30);
    const at30 = await atAngle('Banner', '30');
    check('rotate-frame-word: a picture turned to 30 by its rotate handle wears its frame, eight handles and rotate handle turned with it at 100% and 150%; a click at its centre picks it, one in the empty corner of its old upright box does not',
      turned30 && at30.ok, JSON.stringify({ rot: drawing('Banner')?.rot, ...at30 }));

    // A corner drag at 30: along its own sides, the opposite corner staying put.
    await select('Banner');
    const b0 = await boxOf('Banner');
    const d0 = { w: drawing('Banner').widthPx, h: drawing('Banner').heightPx };
    const fixed = frameCorners(b0, 30)[0];
    const se = (await handles())[2];
    const a30 = Math.PI / 6;
    // Word keeps a picture's proportions from a corner: 56 along its width is 16 along its height.
    const along = { x: 56 * Math.cos(a30) - 16 * Math.sin(a30), y: 56 * Math.sin(a30) + 16 * Math.cos(a30) };
    if (se) await mouse.drag(se, { x: se.x + along.x, y: se.y + along.y }, { steps: 8 });
    const resized = await until(() => Math.abs((drawing('Banner')?.widthPx || 0) - (d0.w + 56)) <= 3, 'the picture resized', 5000).then(() => true).catch(() => false);
    let moved = Infinity;
    await until(async () => {
      const b1 = await boxOf('Banner');
      if (!b1) return false;
      const kept = frameCorners(b1, drawing('Banner')?.rot || 0)[0];
      moved = +Math.hypot(kept.x - fixed.x, kept.y - fixed.y).toFixed(1);
      return moved <= 3;
    }, 'the opposite corner to stay put', 4000).catch(() => false);
    const reframed = await frameOn('Banner');
    await shoot(win, wait, 'rotate-frame-word-resized.png');
    check('rotate-frame-word: a corner drag on the turned picture resizes it along its own sides, its proportions kept and the opposite corner staying put on the page, and the frame follows',
      resized && moved <= 3 && drawing('Banner')?.rot === 30 && reframed.ok,
      JSON.stringify({ before: d0, after: { w: drawing('Banner')?.widthPx, h: drawing('Banner')?.heightPx }, moved, reframed }));

    // 90, by More Rotation Options…
    await select('Banner');
    const opened = await menuPick('Rotate', 'More Rotation Options…');
    await until(() => js(`Boolean(document.querySelector('.rotation-angle'))`), 'the Rotation dialog', 3000).catch(() => false);
    const shown = await js(`document.querySelector('.rotation-angle')?.value`);
    await js(`(() => { const i = document.querySelector('.rotation-angle'); if (!i) return 0; const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set; set.call(i, '90'); i.dispatchEvent(new Event('input', { bubbles: true })); document.querySelector('.rotation-ok')?.click(); return 1; })()`);
    const turned90 = await until(() => Math.round(drawing('Banner')?.rot || 0) === 90, 'the picture at 90', 5000).then(() => true).catch(() => false);
    const at90 = await atAngle('Banner', '90');
    check('rotate-frame-word: Layout → Rotate → More Rotation Options… shows the angle and turns the picture to the one typed; at 90 the frame stays on it at 100% and 150% and the clicks land by its turned outline',
      opened === 'clicked' && shown === '30' && turned90 && at90.ok, JSON.stringify({ opened, shown, rot: drawing('Banner')?.rot, ...at90 }));

    // 135, by the rotate handle again; then Undo and Redo.
    const turned135 = await turnTo('Banner', 135);
    const at135 = await atAngle('Banner', '135');
    check('rotate-frame-word: turned on to 135, the frame stays on the picture at 100% and 150% and the clicks land by its turned outline',
      turned135 && at135.ok, JSON.stringify({ rot: drawing('Banner')?.rot, ...at135 }));
    await select('Banner');
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 'z', { modifiers: ['control'] });
    const undone = await until(() => Math.round(drawing('Banner')?.rot || 0) === 90, 'undo to 90', 4000).then(() => true).catch(() => false);
    await select('Banner');
    const undoFrame = await frameOn('Banner');
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 'y', { modifiers: ['control'] });
    const redone = await until(() => Math.round(drawing('Banner')?.rot || 0) === 135, 'redo to 135', 4000).then(() => true).catch(() => false);
    await select('Banner');
    const redoFrame = await frameOn('Banner');
    check('rotate-frame-word: Undo takes the last turn back and Redo puts it again, the frame following each',
      undone && redone && undoFrame.ok && redoFrame.ok, JSON.stringify({ undone, undoFrame, redone, redoFrame }));

    // The arrows nudge it, and the frame comes along.
    const c0 = await centreOn('Banner');
    for (let i = 0; i < 3; i++) await press(wc, 'Right');
    const nudged = await until(async () => ((await centreOn('Banner'))?.x || 0) >= c0.x + 20, 'the nudge', 4000).then(() => true).catch(() => false);
    const nudgeFrame = await frameOn('Banner');
    check('rotate-frame-word: the arrows nudge the turned picture and its frame comes along', nudged && nudgeFrame.ok, JSON.stringify({ from: c0, to: await centreOn('Banner'), nudgeFrame }));

    // The text box, turned 30: its frame on it, and a click in its words puts the caret there.
    const boxTurned = await turnTo('Note', 30);
    const boxFrame = await frameOn('Note');
    await shoot(win, wait, 'rotate-frame-word-textbox.png');
    await letGo();
    const mid = await centreOn('Note');
    const word = await js(`(() => {
      const box = document.querySelector(${JSON.stringify(sel('Note'))});
      const w = [...box.querySelectorAll('.wd-block, p, span')].find((n) => /Turned words/.test(n.textContent || ''));
      const r = (w || box).getBoundingClientRect();
      return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
    })()`);
    await mouse.click(word || mid);
    const caretIn = await until(() => js(`(() => { const a = getSelection().anchorNode; const el = a && (a.nodeType === 1 ? a : a.parentElement); return Boolean(el?.closest(${JSON.stringify(sel('Note'))})); })()`), 'the caret in the words', 3000).then(() => true).catch(() => false);
    await shoot(win, wait, 'rotate-frame-word-textbox-typing.png');
    check('rotate-frame-word: a text box turned by its rotate handle wears its frame turned with it, and a click on its words puts the caret in them',
      boxTurned && boxFrame.ok && caretIn, JSON.stringify({ rot: drawing('Note')?.rot, boxFrame, caretIn }));
    await letGo();

    // The shape and the picture in line, by Rotate Right 90°.
    const shapeName = (model().drawings || []).find((d) => d.kind === 'shape')?.name;
    const inline = [];
    for (const name of ['Inline', shapeName].filter(Boolean)) {
      await select(name);
      const r = await menuPick('Rotate', 'Rotate Right 90°');
      const turned = await until(() => drawing(name)?.rot === 90, `${name} at 90`, 4000).then(() => true).catch(() => false);
      await select(name);
      const frame = await frameOn(name);
      await shoot(win, wait, `rotate-frame-word-${name === 'Inline' ? 'inline' : 'shape'}.png`);
      const click = await clicks(name);
      inline.push({ name, r, turned, frame, click, anchored: drawing(name)?.anchored });
    }
    check('rotate-frame-word: a shape and a picture in line with the text, turned by Rotate Right 90°, keep their frames on them and take clicks by their turned outlines',
      inline.length === 2 && inline.every((x) => x.r === 'clicked' && x.turned && x.frame.ok && x.click.ok), JSON.stringify(inline));

    // Right to left, as the window is in Urdu or Arabic.
    await js(`document.documentElement.dir = 'rtl', 1`);
    await wait(400);
    await select('Banner');
    const rtlFrame = await frameOn('Banner');
    await shoot(win, wait, 'rotate-frame-word-rtl.png');
    const rtlClicks = await clicks('Banner');
    await js(`document.documentElement.dir = 'ltr', 1`);
    await wait(300);
    check('rotate-frame-word: with the window right to left, the turned picture’s frame stays on it and the clicks land by its turned outline',
      rtlFrame.ok && rtlClicks.ok, JSON.stringify({ rtlFrame, rtlClicks }));

    // Saved and opened again.
    await letGo();
    await js(`document.querySelector('.wd-page')?.focus(), 1`);
    await press(wc, 's', { modifiers: ['control'] });
    const inFile = () => { try { return Object.fromEntries(openDocx(fs.readFileSync(file)).drawings().map((d) => [d.name, d.rot || 0])); } catch { return {}; } };
    const saved = await until(() => { const f = inFile(); return f.Banner === 135 && f.Note === 30 && f.Inline === 90 && (!shapeName || f[shapeName] === 90); }, 'the angles in the file', 8000).then(() => true).catch(() => false);
    const sid = session()?.id;
    const keep = path.join(root, 'keep.docx');
    fs.writeFileSync(keep, buildDocx({ paragraphs: [{ text: 'Keep' }] }));
    await open('word', keep);
    win.close();
    await until(() => win.isDestroyed(), 'the window closed', 4000).catch(() => {});
    if (sid && doc.sessions().some((s) => s.id === sid)) await doc.close({ id: sid }).catch(() => {});
    win = await open('word', file);
    wc = win.webContents;
    js = (code) => wc.executeJavaScript(code);
    mouse = pointer(win, wait);
    await until(() => js(`document.querySelectorAll('.wd-drawing[data-drawing]').length >= 4`), 'the drawings again', 10000).catch(() => {});
    await setZoom(1);
    const reopenedRot = drawing('Banner')?.rot;
    await select('Banner');
    const reopenedFrame = await frameOn('Banner');
    await shoot(win, wait, 'rotate-frame-word-reopened.png');
    const reopenedClicks = await clicks('Banner');
    check('rotate-frame-word: saved and opened again, each drawing keeps its angle and the frame is still on the turned picture',
      saved && reopenedRot === 135 && reopenedFrame.ok && reopenedClicks.ok, JSON.stringify({ saved, file: inFile(), reopenedRot, reopenedFrame, reopenedClicks }));

    const complaints = await errorsIn(win);
    check('rotate-frame-word: the turned-frame checks report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('rotate-frame-word: the turned-frame checks ran', false, err.stack || err.message);
  }
}
