// The panels beside the work resize, and keep their size. In Presentations
// (the slides), Mail (the folders), Calendar, Pictures (the details) and
// Video (the properties), the panel's grip, dragged, makes it wider; the width is kept
// for the next window of that app; the arrow keys move it a step; and a
// double-click puts it back. In Video the timeline's grips make the track
// names' column wider and a kind of track taller, kept the same way. Files
// are the block's own. Run alone with RUTBA_VERIFY_ONLY=panels.

import fs from 'node:fs';
import path from 'node:path';
import { buildPptx } from '@rutba/presentation';
import { gradientPng } from './sample-picture.js';

export async function verifyPanels({ open, check, until, wait }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  const own = path.join(dir, 'panels-fixtures');
  const pics = path.join(own, 'pictures');
  fs.mkdirSync(pics, { recursive: true });
  const pptx = path.join(own, 'panels.pptx');
  fs.writeFileSync(pptx, buildPptx({ title: 'Panels', slides: [{ layout: 'title', title: 'Panels' }, { layout: 'obj', title: 'Two', body: ['a'] }] }));
  const png = path.join(own, 'panels.png');
  fs.writeFileSync(png, gradientPng(960, 540, [20, 110, 140], [240, 190, 90]));
  fs.copyFileSync(png, path.join(pics, 'one.png'));
  const snap = async (win, name) => {
    if (!shots) return;
    win.webContents.invalidate();
    await wait(700);
    fs.writeFileSync(path.join(shots, name), (await win.webContents.capturePage()).toPNG());
  };
  const waitFor = async (fn, what, ms) => {
    let got = null;
    await until(async () => { got = await fn(); return Boolean(got); }, what, ms).catch(() => {});
    return got || null;
  };
  const drag = async (win, selector, dx, dy) => {
    const js = (code) => win.webContents.executeJavaScript(code);
    const at = await js(`(() => { const r = document.querySelector(${JSON.stringify(selector)})?.getBoundingClientRect(); return r && r.width > 0 ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : null; })()`);
    if (!at) return false;
    const wc = win.webContents;
    wc.sendInputEvent({ type: 'mouseMove', x: at.x, y: at.y });
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    for (const f of [0.25, 0.5, 0.75, 1]) {
      wc.sendInputEvent({ type: 'mouseMove', x: Math.round(at.x + dx * f), y: Math.round(at.y + dy * f), button: 'left' });
      await wait(40);
    }
    wc.sendInputEvent({ type: 'mouseUp', x: at.x + dx, y: at.y + dy, button: 'left', clickCount: 1 });
    return true;
  };

  // A window held open while the block closes and opens its own: the app ends when its last window closes.
  await open('home').catch(() => null);
  /** A panel's width on its grip, by the grip's place among the window's grips. */
  const PANELS = [
    { app: 'slides', file: pptx, ready: `document.querySelectorAll('.sl-thumb').length >= 2`, grip: '.rw-panel:not(.right) > .rw-panel-grip', sign: 1 },
    { app: 'mail', file: undefined, ready: `Boolean(document.querySelector('.rw-panel > .rw-panel-grip'))`, grip: '.rw-panel:not(.right) > .rw-panel-grip', sign: 1 },
    { app: 'calendar', file: undefined, ready: `Boolean(document.querySelector('.rw-panel > .rw-panel-grip'))`, grip: '.rw-panel > .rw-panel-grip', sign: 1 },
    // Pictures with a picture open: its Details panel, on the right.
    { app: 'pictures', file: path.join(pics, 'one.png'), ready: `Boolean(document.querySelector('.rw-panel.right > .rw-panel-grip'))`, grip: '.rw-panel.right > .rw-panel-grip', sign: -1 },
    { app: 'video', file: png, ready: `document.querySelectorAll('.vd-clip').length > 0 && Boolean(document.querySelector('.rw-panel.right > .rw-panel-grip'))`, grip: '.rw-panel.right > .rw-panel-grip', sign: -1 },
  ];
  for (const { app, file, ready, grip, sign } of PANELS) {
    try {
      let win = await open(app, file);
      win.setContentSize(1400, 820);
      let js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(ready), `the ${app} panel`, 15000);
      await wait(300);
      const width = () => js(`Math.round(document.querySelector(${JSON.stringify(grip)})?.parentElement.getBoundingClientRect().width || 0)`);
      // From the width as it starts.
      await js(`(() => { document.querySelector(${JSON.stringify(grip)})?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);
      const w0 = await waitFor(() => width(), 'the width', 3000);
      const dragged = await drag(win, grip, 80 * sign, 0);
      const w1 = await waitFor(() => width().then((w) => (Math.abs(w - (w0 + 80)) <= 3 ? w : null)), 'the panel wider', 3000);
      await js(`document.querySelector(${JSON.stringify(grip)})?.focus(), 1`);
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode: sign > 0 ? 'Right' : 'Left' });
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode: sign > 0 ? 'Right' : 'Left' });
      const w2 = await waitFor(() => width().then((w) => (Math.abs(w - ((w1 || 0) + 16)) <= 3 ? w : null)), 'a step wider', 3000);
      if (app === 'video') await snap(win, 'panels-video.png');
      if (app === 'slides') await snap(win, 'panels-slides.png');
      win.close();
      await until(() => win.isDestroyed(), 'the window closed', 5000).catch(() => {});
      win = await open(app, file);
      win.setContentSize(1400, 820);
      js = (code) => win.webContents.executeJavaScript(code);
      await until(() => js(ready), `the ${app} panel again`, 15000);
      const again = await waitFor(() => width().then((w) => (w2 && Math.abs(w - w2) <= 2 ? w : null)), 'the kept width', 4000);
      await js(`(() => { document.querySelector(${JSON.stringify(grip)})?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);
      const back = await waitFor(() => width().then((w) => (Math.abs(w - w0) <= 2 ? w : null)), 'the width as it starts', 3000);
      check(`panels: ${app}'s panel is made wider by its grip and a step wider by the arrow keys, keeps that width in its next window, and a double-click puts it back`,
        Boolean(dragged && w0 && w1 && w2 && again && back), JSON.stringify({ w0, w1, w2, again, back }));
    } catch (err) {
      check(`panels: the ${app} panel checks ran`, false, err.stack || err.message);
    }
  }

  // Video's timeline: the track names wider, the picture tracks taller.
  try {
    let win = await open('video', png);
    win.setContentSize(1400, 820);
    let js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.vd-clip').length > 0`), 'the clip', 20000);
    const sizes = () => js(`(() => { const head = document.querySelector('.vd-row[data-track="V1"] .vd-row-head')?.getBoundingClientRect(); const row = document.querySelector('.vd-row[data-track="V1"]')?.getBoundingClientRect(); return head && row ? { head: Math.round(head.width), row: Math.round(row.height) } : null; })()`);
    await js(`(() => { document.querySelector('.vd-head-grip')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); document.querySelector('.vd-row[data-track="V1"] .vd-lane-grip')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);
    const s0 = await waitFor(() => sizes().then((x) => (x && x.head === 120 && x.row === 60 ? x : null)), 'the sizes as they start', 3000);
    await drag(win, '.vd-head-grip', 50, 0);
    await drag(win, '.vd-row[data-track="V1"] .vd-lane-grip', 0, 30);
    const s1 = await waitFor(() => sizes().then((x) => (x && Math.abs(x.head - 170) <= 3 && Math.abs(x.row - 90) <= 3 ? x : null)), 'wider names and a taller track', 4000);
    const clip = await js(`Math.round(document.querySelector('.vd-row[data-track="V1"] .vd-clip')?.getBoundingClientRect().height || 0)`);
    await snap(win, 'panels-video-timeline.png');
    win.close();
    await until(() => win.isDestroyed(), 'the window closed', 5000).catch(() => {});
    win = await open('video', png);
    win.setContentSize(1400, 820);
    js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.vd-clip').length > 0`), 'the clip again', 20000);
    const kept = await waitFor(() => sizes().then((x) => (x && s1 && Math.abs(x.head - s1.head) <= 2 && Math.abs(x.row - s1.row) <= 2 ? x : null)), 'the sizes kept', 4000);
    await js(`(() => { document.querySelector('.vd-head-grip')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); document.querySelector('.vd-row[data-track="V1"] .vd-lane-grip')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);
    const back = await waitFor(() => sizes().then((x) => (x && x.head === 120 && x.row === 60 ? x : null)), 'the sizes back', 3000);
    check('panels: Video\'s timeline makes its track names wider and a kind of track taller by their grips (the clips taller with it), keeps both for the next window, and a double-click puts each back',
      Boolean(s0 && s1 && clip >= 80 && kept && back), JSON.stringify({ s0, s1, clip, kept, back }));
  } catch (err) {
    check('panels: the Video timeline checks ran', false, err.stack || err.message);
  }
}
