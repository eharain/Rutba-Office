// Presentations: Insert → 3D Models, pressed on the ribbon — a .glb (a cube
// whose front face carries a red and white checked picture) put on the
// slide as the picture the suite's own renderer draws of it, the 3D Model
// tab there while it is picked; 3D Model Views → Front turns it to face us,
// the checks drawn from its texture; the turn handle dragged across turns
// it, drawn small while it moves; Reset puts it back; and the saved deck
// keeps the model beside its picture. Run alone with RUTBA_VERIFY_ONLY=model3d.

import fs from 'node:fs';
import path from 'node:path';
import { Deck, buildPptx } from '@rutba/presentation';
import { cubeGlb } from '../../../tests/fixtures/cube-glb.js';

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn, doc
 */
export async function verifyDeckModel3d(h, { dir }) {
  const { open, check, until, wait, errorsIn, doc } = h;
  const file = path.join(dir, 'model3d.pptx');
  const glb = path.join(dir, 'checked cube.glb');
  try {
    fs.writeFileSync(file, Deck.open(buildPptx({ title: 'Models', slides: [{ layout: 'title', title: 'A cube' }] })).save());
    fs.writeFileSync(glb, cubeGlb({ texture: true }));
    const win = await open('slides', file);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const session = doc.sessions().filter((s) => s.kind === 'deck' && s.path === file).pop();
    const model = () => doc.model({ id: session.id, slide: 0 }).slide;
    const cube = () => (model().shapes || []).find((s) => s.model3d);
    const tabs = () => js(`[...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim())`);
    // The model's picture on the stage, read back: the colour at a fraction of its width and height.
    const colourAt = (fx, fy) => js(`(async () => {
      const el = document.querySelector('.sl-svg g[data-shape="${cube()?.id}"] image');
      if (!el) return null;
      const href = el.getAttribute('href') || el.getAttribute('xlink:href') || '';
      try {
        // Its bytes fetched and decoded here, so the canvas reads them back.
        const bitmap = await createImageBitmap(await (await fetch(href)).blob());
        const c = document.createElement('canvas');
        c.width = bitmap.width; c.height = bitmap.height;
        const ctx = c.getContext('2d');
        ctx.drawImage(bitmap, 0, 0);
        return Array.from(ctx.getImageData(Math.round(bitmap.width * ${fx}), Math.round(bitmap.height * ${fy}), 1, 1).data);
      } catch (e) {
        return { error: String(e), href: href.slice(0, 60) };
      }
    })()`);

    // Insert → 3D Models, the file answered as the dialog would.
    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.sl-model3d-insert'))`), 'the 3D Models button', 4000);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [glb] });
    await js(`document.querySelector('.sl-model3d-insert').click(), 1`);
    await until(() => Boolean(cube()), 'the model', 10000).catch(() => {});
    const placed = cube();
    await until(async () => (await tabs()).includes('3D Model'), 'the 3D Model tab', 4000).catch(() => {});
    const drawn = await until(() => js(`Boolean(document.querySelector('.sl-svg g[data-shape="${placed?.id}"] image'))`), 'the picture', 4000).then(() => true).catch(() => false);
    check('presentations: Insert → 3D Models puts a .glb on the slide as the picture the suite draws of it, named for its file, with the 3D Model tab',
      Boolean(placed) && placed.name === 'checked cube' && JSON.stringify(placed.model3d.view) === JSON.stringify({ yaw: 25, pitch: 15, roll: 0 }) && drawn && (await tabs()).includes('3D Model'),
      JSON.stringify({ name: placed?.name, view: placed?.model3d?.view, drawn, tabs: await tabs() }));

    // 3D Model Views → Front: the checked face towards us, drawn from its texture.
    await until(() => js(`Boolean(document.querySelector('.sl-model3d-view[data-view="front"]'))`), 'the views', 3000).catch(() => {});
    await js(`document.querySelector('.sl-model3d-view[data-view="front"]').click(), 1`);
    await until(() => cube()?.model3d?.view?.yaw === 0 && cube()?.model3d?.view?.pitch === 0, 'the front view', 8000).catch(() => {});
    await wait(400);
    const red = await colourAt(0.145, 0.145);
    const white = await colourAt(0.385, 0.145);
    check('presentations: 3D Model Views → Front turns the model to face us, its picture drawn again with the checks of its texture',
      cube()?.model3d?.view?.yaw === 0 && red && white && red[0] > 150 && red[1] < 90 && white[1] > 190 && white[2] > 190,
      JSON.stringify({ view: cube()?.model3d?.view, red, white }));

    // The turn handle, dragged across with the mouse: drawn small as it moves, turned where it is let go.
    const handle = await js(`(() => { const r = document.querySelector('.sl-model3d-handle')?.getBoundingClientRect(); return r ? { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } : null; })()`);
    let preview = false;
    if (handle) {
      win.focus();
      wc.sendInputEvent({ type: 'mouseMove', x: handle.x, y: handle.y });
      wc.sendInputEvent({ type: 'mouseDown', x: handle.x, y: handle.y, button: 'left', clickCount: 1 });
      await wait(300);
      for (let i = 1; i <= 8; i++) {
        wc.sendInputEvent({ type: 'mouseMove', x: handle.x + i * 10, y: handle.y, button: 'left' });
        await wait(60);
      }
      preview = await until(() => js(`Boolean(document.querySelector('.sl-model3d-preview'))`), 'the preview', 2000).then(() => true).catch(() => false);
      wc.sendInputEvent({ type: 'mouseUp', x: handle.x + 80, y: handle.y, button: 'left', clickCount: 1 });
    }
    await until(() => Math.abs((cube()?.model3d?.view?.yaw ?? 0) - 40) < 0.6, 'the turn', 8000).catch(() => {});
    const gone = await until(() => js(`!document.querySelector('.sl-model3d-preview')`), 'the preview gone', 3000).then(() => true).catch(() => false);
    check('presentations: dragging the 3D model\'s handle across turns it, drawn small while it moves and properly where it is let go',
      Boolean(handle) && preview && Math.abs((cube()?.model3d?.view?.yaw ?? 0) - 40) < 0.6 && gone,
      JSON.stringify({ handle, preview, view: cube()?.model3d?.view, gone }));

    // Reset 3D Model, then save.
    await js(`document.querySelector('.sl-model3d-reset')?.click(), 1`);
    await until(() => cube()?.model3d?.view?.yaw === 25, 'the reset', 8000).catch(() => {});
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => { try { return /<r3d:model [^>]*yaw="25"/.test(Deck.open(fs.readFileSync(file)).pkg.text('ppt/slides/slide1.xml')); } catch { return false; } }, 'the save', 8000).then(() => true).catch(() => false);
    const again = saved ? Deck.open(fs.readFileSync(file)) : null;
    const kept = again ? Buffer.from(again.model3dSource(0, placed.id).data).equals(cubeGlb({ texture: true })) : false;
    check('presentations: Reset 3D Model puts it back as it went in, and the saved deck keeps the model beside its picture',
      saved && kept && again.pkg.has('ppt/media/model1.glb'), JSON.stringify({ saved, kept }));

    const complaints = await errorsIn(win);
    check('presentations: 3D models report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('presentations: the 3D model checks ran', false, err?.message || JSON.stringify(err) || String(err));
  }
}
