// Presentations: Insert → Cameo, pressed on the ribbon — a cameo at the
// slide's lower right with the Camera Format tab; Preview fills it with the
// camera, live (in a check run, Chromium's own stand-in camera); Oval makes
// it round; the show fills it with the camera too, and leaving the show
// closes the camera; and the saved deck keeps it. Run alone with
// RUTBA_VERIFY_ONLY=cameo.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

export async function verifyDeckCameo({ open, check, until, wait, errorsIn, doc }, { dir }) {
  const file = path.join(dir, 'cameo.pptx');
  try {
    fs.writeFileSync(file, Deck.open(buildPptx({ title: 'Talk', slides: [{ layout: 'title', title: 'Hello from me' }, { layout: 'obj', title: 'Next' }] })).save());
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const model = () => doc.model({ id: session.id, slide: 0 }).slide;
    const cameo = () => (model().shapes || []).find((s) => s.cameo);
    const tabs = () => js(`[...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim())`);
    const playing = (scope) => js(`(() => { const v = document.querySelector(${JSON.stringify(scope)} + ' .sl-cameo-video'); return v ? { w: v.videoWidth, h: v.videoHeight, live: Boolean(v.srcObject?.active), radius: getComputedStyle(v).borderRadius } : null; })()`);

    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.sl-cameo-insert'))`), 'the Cameo button', 4000);
    await js(`document.querySelector('.sl-cameo-insert').click(), 1`);
    await until(() => Boolean(cameo()), 'the cameo', 6000).catch(() => {});
    await until(async () => (await tabs()).includes('Camera Format'), 'the Camera Format tab', 4000).catch(() => {});
    const W = doc.model({ id: session.id }).size?.width || 1280;
    const g = cameo()?.geometry;
    check('presentations: Insert → Cameo puts a cameo at the slide\'s lower right, with the Camera Format tab',
      Boolean(g) && g.x > W / 2 && (await tabs()).includes('Camera Format'), JSON.stringify({ g, tabs: await tabs() }));

    // Preview: the camera, live, on the stage.
    await js(`document.querySelector('.sl-cameo-preview')?.click(), 1`);
    const previewed = await until(async () => { const p = await playing('body'); return Boolean(p?.live) && p.w > 16; }, 'the camera on the stage', 14000).then(() => true).catch(() => false);
    const stage = await playing('body');
    check('presentations: Camera Format → Preview fills the cameo with the camera, live, mirrored as a camera pointed at oneself is',
      previewed && stage?.live, JSON.stringify(stage));

    // Oval.
    await js(`document.querySelector('.sl-cameo-shape[data-preset="ellipse"]')?.click(), 1`);
    await until(() => cameo()?.preset === 'ellipse', 'the oval', 5000).catch(() => {});
    await wait(300);
    const round = await playing('body');
    check('presentations: Camera Styles → Oval makes the cameo round, the camera cropped to it', cameo()?.preset === 'ellipse' && round?.radius === '50%', JSON.stringify({ preset: cameo()?.preset, round }));
    await js(`document.querySelector('.sl-cameo-preview')?.click(), 1`);
    await until(async () => !(await playing('body')), 'the preview off', 3000).catch(() => {});

    // The show: the camera in the cameo, closed again when the show ends.
    const startShow = () => js(`(async () => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Slide Show')?.click(); await new Promise((r) => setTimeout(r, 200)); [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === 'From Beginning')?.click(); return 1; })()`);
    const liveInShow = () => until(async () => { const p = await playing('.sl-present'); return Boolean(p?.live) && p.w > 16; }, 'the camera in the show', 14000).then(() => true).catch(() => false);
    await startShow();
    let inShow = await liveInShow();
    // A check run's stand-in camera, on a busy machine, sometimes will not
    // start again so soon after the preview let it go: the show is left and
    // started again, after a longer pause each time, and must then show the
    // camera live.
    for (const pause of [2000, 5000]) {
      if (inShow) break;
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
      await until(() => js(`!document.querySelector('.sl-present')`), 'the show closed', 4000).catch(() => {});
      await wait(pause);
      await startShow();
      inShow = await liveInShow();
    }
    const shown = await playing('.sl-present');
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    const closed = await until(() => js(`!document.querySelector('.sl-present') && !document.querySelector('.sl-cameo-video')`), 'the show closed', 4000).then(() => true).catch(() => false);
    check('presentations: the show fills the cameo with the camera, live, and leaving the show closes it', inShow && shown?.live && closed, JSON.stringify({ shown, closed }));

    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => { try { return /<rcam:cameo\b/.test(Deck.open(fs.readFileSync(file)).pkg.text('ppt/slides/slide1.xml')); } catch { return false; } }, 'the save', 6000).then(() => true).catch(() => false);
    const kept = saved ? Deck.open(fs.readFileSync(file)).slide(0).shapes.find((s) => s.cameo) : null;
    check('presentations: the saved deck keeps the cameo and its shape', Boolean(kept) && kept.preset === 'ellipse', JSON.stringify({ saved, preset: kept?.preset }));

    const complaints = await errorsIn(win);
    check('presentations: Cameo reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('presentations: the Cameo checks ran', false, err?.message || JSON.stringify(err));
  }
}
