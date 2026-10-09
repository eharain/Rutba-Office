// Presentations: Record with the camera — a deck whose first slide has a
// cameo is recorded across both slides (in a check run, Chromium's stand-in
// camera and microphone): the first slide gets the camera's recording in its
// cameo's place, the second (no cameo) gets none; in the show the recording
// plays with its slide; Reset to Cameo takes it off and the cameo is live
// again. Run alone with RUTBA_VERIFY_ONLY=cameorecord.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';

export async function verifyDeckCameoRecord({ open, check, until, wait, errorsIn, doc }, { dir }) {
  const file = path.join(dir, 'cameo-record.pptx');
  try {
    const deck = Deck.open(buildPptx({ title: 'Talk', slides: [{ layout: 'title', title: 'Hello' }, { layout: 'obj', title: 'Next' }] }));
    deck.addCameo(0, { shape: 'ellipse' });
    fs.writeFileSync(file, deck.save());
    const win = await open('slides', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const shapesOf = (i) => doc.model({ id: session.id, slide: i }).slide?.shapes || [];
    const tab = (name) => js(`(async () => { [...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(); await new Promise((r) => setTimeout(r, 200)); return 1; })()`);
    const ribbon = (label) => js(`(() => { const b = [...document.querySelectorAll('.rw-ribbon .rw-btn')].find((n) => n.textContent.trim() === ${JSON.stringify(label)}); if (!b) return 'no button'; b.click(); return 'clicked'; })()`);
    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);

    // Record from the beginning: a moment on each slide, then out.
    await tab('Record');
    await ribbon('From Beginning');
    await until(() => js(`Boolean(document.querySelector('.sl-present'))`), 'the show', 6000).catch(() => {});
    // Whether the window had to fall back to the voice alone, for the report.
    const refused = await js(`document.body.textContent.includes('The camera could not be opened')`);
    await wait(1500);
    await js(`document.querySelector('.sl-present')?.click(), 1`);
    await wait(1200);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await until(() => shapesOf(0).some((s) => s.cameoRecording), 'the recording in the cameo', 12000).catch(() => {});
    const rec = shapesOf(0).find((s) => s.cameoRecording);
    const cam = shapesOf(0).find((s) => s.cameo);
    check('presentations: Record with a cameo puts the camera\'s recording in the cameo\'s place on its slide, and none on a slide without a cameo',
      Boolean(rec) && rec.media?.kind === 'video' && Math.abs(rec.geometry.x - cam.geometry.x) < 1 && Math.abs(rec.geometry.w - cam.geometry.w) < 1 && !shapesOf(1).some((s) => s.cameoRecording),
      JSON.stringify({ rec: rec ? { kind: rec.media?.kind, g: rec.geometry } : null, cam: cam?.geometry, second: shapesOf(1).map((s) => s.name), refused }));

    // The show: the recording plays with its slide, and no live camera over it.
    await tab('Slide Show');
    await ribbon('From Beginning');
    await until(() => js(`Boolean(document.querySelector('.sl-present .sl-show-clip video'))`), 'the recording in the show', 6000).catch(() => {});
    const playing = await until(() => js(`(() => { const v = document.querySelector('.sl-present .sl-show-clip video'); return v ? !v.paused && v.currentTime > 0 : false; })()`), 'the recording playing', 6000).then(() => true).catch(() => false);
    const live = await js(`Boolean(document.querySelector('.sl-present .sl-cameo-video'))`);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Escape' });
    win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Escape' });
    await until(() => js(`!document.querySelector('.sl-present')`), 'the show over', 4000).catch(() => {});
    check('presentations: in the show the cameo\'s recording plays with its slide, the live camera not over it', playing && !live, JSON.stringify({ playing, live }));

    // Reset to Cameo.
    await tab('Record');
    await ribbon('Reset to Cameo');
    await until(() => !shapesOf(0).some((s) => s.cameoRecording), 'the recording off', 5000).catch(() => {});
    check('presentations: Reset to Cameo takes the recording off the slide, the cameo left', !shapesOf(0).some((s) => s.cameoRecording) && shapesOf(0).some((s) => s.cameo), JSON.stringify(shapesOf(0).map((s) => s.name)));

    const complaints = await errorsIn(win);
    check('presentations: recording into the cameo reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('presentations: the cameo recording checks ran', false, err?.message || JSON.stringify(err));
  }
}
