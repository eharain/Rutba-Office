// Video: the preview and the timeline share the window's height. At
// 1300x700 and 1680x1050 the preview has at least 45% of the editor's
// height. The divider between them drags (and moves with the arrow keys),
// the share is kept on this computer and comes back in the next window, a
// short timeline scrolls its lanes under a ruler that stays put, and a
// double-click on the divider puts the share back as it starts. Its picture
// is written in a folder of its own. Run alone with
// RUTBA_VERIFY_ONLY=video-split.

import fs from 'node:fs';
import path from 'node:path';
import { gradientPng } from './sample-picture.js';

const MEASURE = `(() => {
  const ed = document.querySelector('.vd-editor')?.getBoundingClientRect();
  const stage = document.querySelector('.vd-stage')?.getBoundingClientRect();
  const tl = document.querySelector('.vd-timeline')?.getBoundingClientRect();
  const split = document.querySelector('.vd-split');
  const scroll = document.querySelector('.vd-tl-scroll');
  const ruler = document.querySelector('.vd-ruler')?.getBoundingClientRect();
  return {
    editor: Math.round(ed?.height || 0),
    stage: Math.round(stage?.height || 0),
    timeline: Math.round(tl?.height || 0),
    share: Number(split?.getAttribute('aria-valuenow') || 0),
    scrolls: scroll ? scroll.scrollHeight > scroll.clientHeight + 1 && getComputedStyle(scroll).overflowY === 'auto' : false,
    rulerTop: ruler && tl ? Math.round(ruler.top - tl.top) : null,
  };
})()`;

export async function verifyVideoPane({ open, check, until, wait }, { dir }) {
  const shots = process.env.RUTBA_VERIFY_CAPTURE;
  const own = path.join(dir, 'video-split-fixtures');
  fs.mkdirSync(own, { recursive: true });
  const still = path.join(own, 'landscape.png');
  fs.writeFileSync(still, gradientPng(1280, 720, [20, 110, 140], [240, 190, 90]));
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
  const size = async (win, js, w, h) => {
    win.setContentSize(w, h);
    await until(() => js('[innerWidth, innerHeight]').then(([iw, ih]) => Math.abs(iw - w) <= 2 && Math.abs(ih - h) <= 2), `the window at ${w}x${h}`, 5000).catch(() => {});
    await wait(400);
  };

  let shareAfterDrag = null;
  try {
    const win = await open('video', still);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.vd-clip').length > 0 && Boolean(document.querySelector('.vd-split'))`), 'the picture on the timeline', 20000);
    // Whatever an earlier run kept, the checks start from the share as it starts.
    await js(`(() => { document.querySelector('.vd-split')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);

    for (const [w, h] of [[1300, 700], [1680, 1050]]) {
      await size(win, js, w, h);
      const m = await waitFor(() => js(MEASURE).then((x) => (x.editor > 0 && x.stage > 0 ? x : null)), `the editor at ${w}x${h}`, 4000);
      await snap(win, `video-split-${w}x${h}.png`);
      check(`video-split: at ${w}x${h} the preview has at least 45% of the editor's height, the timeline the rest less its transport bar`,
        Boolean(m && m.stage >= m.editor * 0.45 && m.timeline >= 120), JSON.stringify(m));
    }

    // Drag the divider up 140 px at 1300x700: the timeline grows, the preview gives way.
    await size(win, js, 1300, 700);
    const before = await js(MEASURE);
    const at = await js(`(() => { const r = document.querySelector('.vd-split').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }; })()`);
    const wc = win.webContents;
    wc.sendInputEvent({ type: 'mouseMove', x: at.x, y: at.y });
    wc.sendInputEvent({ type: 'mouseDown', x: at.x, y: at.y, button: 'left', clickCount: 1 });
    for (const f of [0.25, 0.5, 0.75, 1]) {
      wc.sendInputEvent({ type: 'mouseMove', x: at.x, y: Math.round(at.y - 140 * f), button: 'left' });
      await wait(40);
    }
    wc.sendInputEvent({ type: 'mouseUp', x: at.x, y: at.y - 140, button: 'left', clickCount: 1 });
    const dragged = await waitFor(() => js(MEASURE).then((x) => (x.timeline >= before.timeline + 100 ? x : null)), 'the timeline taller', 4000);
    const kept = await waitFor(() => js(`window.rutbaOffice.store.get({ key: 'video.timelineShare', fallback: null })`).then((v) => (Number.isFinite(v) && v > 0.4 ? v : null)), 'the share kept', 4000);
    shareAfterDrag = kept;
    await snap(win, 'video-split-dragged-1300x700.png');
    check('video-split: dragging the divider up gives the timeline the height the preview gives up, and the share is kept on this computer',
      Boolean(dragged && dragged.stage < before.stage - 100 && kept), JSON.stringify({ before, dragged, kept }));

    // The keyboard: Down gives the preview room back a step at a time.
    await js(`document.querySelector('.vd-split')?.focus(), 1`);
    const k0 = (await js(MEASURE)).timeline;
    wc.sendInputEvent({ type: 'keyDown', keyCode: 'Down' });
    wc.sendInputEvent({ type: 'keyUp', keyCode: 'Down' });
    const stepped = await waitFor(() => js(MEASURE).then((x) => (x.timeline <= k0 - 20 ? x : null)), 'a step down', 3000);
    check('video-split: the divider takes the keyboard: Down moves it a step towards the timeline', Boolean(stepped), `${k0} then ${stepped?.timeline}`);

    // A short timeline: the lanes scroll and the ruler stays at the top.
    for (let i = 0; i < 12; i++) {
      wc.sendInputEvent({ type: 'keyDown', keyCode: 'PageDown' });
      wc.sendInputEvent({ type: 'keyUp', keyCode: 'PageDown' });
    }
    const short = await waitFor(() => js(MEASURE).then((x) => (x.timeline <= 130 ? x : null)), 'the timeline at its least', 3000);
    await js(`(() => { const s = document.querySelector('.vd-tl-scroll'); s.scrollTop = s.scrollHeight; return 1; })()`);
    await wait(200);
    const scrolled = await js(MEASURE);
    await snap(win, 'video-split-short-1300x700.png');
    check('video-split: at its least the timeline keeps its ruler at the top and scrolls the lanes under it', Boolean(short && short.scrolls && scrolled.rulerTop !== null && scrolled.rulerTop <= 1), JSON.stringify({ short, rulerTop: scrolled.rulerTop }));

    // Back to the drag's share, for the next window to find.
    await js(`(async () => { await window.rutbaOffice.store.set({ key: 'video.timelineShare', value: ${JSON.stringify(shareAfterDrag ?? 0.5)} }); return 1; })()`);
    win.close();
    await until(() => win.isDestroyed(), 'the window closed', 5000).catch(() => {});
  } catch (err) {
    check('video-split: the divider checks ran', false, err.stack || err.message);
  }

  // A new window comes back with the share kept, and a double-click puts it back.
  try {
    const win = await open('video', still);
    const js = (code) => win.webContents.executeJavaScript(code);
    await until(() => js(`document.querySelectorAll('.vd-clip').length > 0`), 'the picture again', 20000);
    await size(win, js, 1300, 700);
    const want = Math.round((shareAfterDrag ?? 0.5) * 100);
    const back = await waitFor(() => js(MEASURE).then((x) => (Math.abs(x.share - want) <= 1 ? x : null)), 'the kept share', 4000);
    check('video-split: a new Video window opens with the divider where it was left', Boolean(back), `want ${want}; ${JSON.stringify(back || (await js(MEASURE)))}`);
    await js(`(() => { document.querySelector('.vd-split')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return 1; })()`);
    const reset = await waitFor(() => js(MEASURE).then((x) => (x.share === 36 && x.stage >= x.editor * 0.45 ? x : null)), 'the share as it starts', 4000);
    const stored = await js(`window.rutbaOffice.store.get({ key: 'video.timelineShare', fallback: null })`);
    check('video-split: a double-click on the divider puts the share back as it starts, and keeps that', Boolean(reset) && stored === 0.36, JSON.stringify({ reset, stored }));
  } catch (err) {
    check('video-split: the second window checks ran', false, err.stack || err.message);
  }
}
