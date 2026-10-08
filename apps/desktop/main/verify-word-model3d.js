// Documents: Insert → 3D Models, pressed on the ribbon — a .glb (a cube with
// a checked front face) put at the caret as the picture the suite draws of
// it, the 3D Model tab there while it is picked; Front turns it to face us,
// its checks drawn from its texture; Turn Right turns it on; Undo takes the
// turn back, picture and all; and the saved document keeps the model beside
// its picture. Run alone with RUTBA_VERIFY_ONLY=wordmodel3d.

import fs from 'node:fs';
import path from 'node:path';
import { buildDocx, OoxmlPackage } from '@rutba/ooxml';
import { cubeGlb } from '../../../tests/fixtures/cube-glb.js';

export async function verifyWordModel3d({ open, check, until, wait, errorsIn, doc, sessionFor }, { dir }) {
  const file = path.join(dir, 'model3d.docx');
  const glb = path.join(dir, 'word cube.glb');
  try {
    fs.writeFileSync(file, buildDocx({ styles: true, paragraphs: [{ text: 'A model of a cube follows.' }, { text: 'And the words go on.' }] }));
    fs.writeFileSync(glb, cubeGlb({ texture: true }));
    const win = await open('word', file);
    const js = (code) => win.webContents.executeJavaScript(code);
    const session = () => sessionFor('doc');
    const model = () => doc.model({ id: session().id });
    const found = () => {
      const blocks = model().blocks || [];
      for (let b = 0; b < blocks.length; b++) {
        const i = (blocks[b].images || []).findIndex((img) => img.model3d);
        if (i >= 0) return { block: b, image: i, img: blocks[b].images[i] };
      }
      return null;
    };
    const tabs = () => js(`[...document.querySelectorAll('.rw-tab')].map((t) => t.textContent.trim())`);
    const colourAt = (fx, fy) => {
      const href = found()?.img?.href;
      if (!href) return Promise.resolve(null);
      return js(`(async () => {
        try {
          const bitmap = await createImageBitmap(await (await fetch(${JSON.stringify(href)})).blob());
          const c = document.createElement('canvas');
          c.width = bitmap.width; c.height = bitmap.height;
          const ctx = c.getContext('2d');
          ctx.drawImage(bitmap, 0, 0);
          return Array.from(ctx.getImageData(Math.round(bitmap.width * ${fx}), Math.round(bitmap.height * ${fy}), 1, 1).data);
        } catch (e) { return { error: String(e) }; }
      })()`);
    };

    // Insert → 3D Models, the file answered as the dialog would.
    await until(() => js(`Boolean(document.querySelector('.rw-tab'))`), 'the ribbon', 8000);
    await js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === 'Insert')?.click(), 1`);
    await until(() => js(`Boolean(document.querySelector('.wd-model3d-insert'))`), 'the 3D Models button', 4000);
    globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'open', answer: [glb] });
    await js(`document.querySelector('.wd-model3d-insert').click(), 1`);
    await until(() => Boolean(found()), 'the model', 10000).catch(() => {});
    const placed = found();
    await until(async () => (await tabs()).includes('3D Model'), 'the 3D Model tab', 4000).catch(() => {});
    check('word: Insert → 3D Models puts a .glb at the caret as the picture the suite draws of it, with the 3D Model tab',
      Boolean(placed) && placed.img.name === 'word cube' && JSON.stringify(placed.img.model3d.view) === JSON.stringify({ yaw: 25, pitch: 15, roll: 0 }) && (await tabs()).includes('3D Model'),
      JSON.stringify({ name: placed?.img?.name, view: placed?.img?.model3d?.view, block: placed?.block, tabs: await tabs() }));

    // Front: the checked face towards us.
    await until(() => js(`Boolean(document.querySelector('.wd-model3d-view[data-view="front"]'))`), 'the views', 3000).catch(() => {});
    await js(`document.querySelector('.wd-model3d-view[data-view="front"]')?.click(), 1`);
    await until(() => found()?.img?.model3d?.view?.yaw === 0 && found()?.img?.model3d?.view?.pitch === 0, 'the front view', 8000).catch(() => {});
    const red = await colourAt(0.145, 0.145);
    const white = await colourAt(0.385, 0.145);
    check('word: 3D Model Views → Front turns the model to face us, its picture drawn again with the checks of its texture',
      found()?.img?.model3d?.view?.yaw === 0 && Array.isArray(red) && Array.isArray(white) && red[0] > 150 && red[1] < 90 && white[1] > 190,
      JSON.stringify({ view: found()?.img?.model3d?.view, red, white }));

    // Turn Right, then Undo.
    await js(`[...document.querySelectorAll('.rw-ribbon .rw-btn')].find((b) => b.textContent.trim() === 'Turn Right')?.click(), 1`);
    const turned = await until(() => found()?.img?.model3d?.view?.yaw === 15, 'the turn', 8000).then(() => true).catch(() => false);
    const turnedHref = found()?.img?.href;
    await js(`[...document.querySelectorAll('.rw-btn')].find((b) => (b.title || b.dataset.tip || '').startsWith('Undo'))?.click(), 1`);
    const undone = await until(() => found()?.img?.model3d?.view?.yaw === 0, 'the undo', 6000).then(() => true).catch(() => false);
    check('word: Turn Right turns the model fifteen degrees, and Undo takes the turn back, picture and all',
      turned && undone && found()?.img?.href !== turnedHref, JSON.stringify({ turned, undone, view: found()?.img?.model3d?.view }));

    // Saved.
    await js(`(() => { [...document.querySelectorAll('.rw-btn')].find((n) => (n.title || n.dataset.tip || '').startsWith('Save'))?.click(); return 1; })()`);
    const saved = await until(() => { try { const pkg = OoxmlPackage.read(fs.readFileSync(file)); return pkg.has('word/media/model1.glb') && /<r3d:model [^>]*yaw="0"/.test(pkg.text('word/document.xml')); } catch { return false; } }, 'the save', 8000).then(() => true).catch(() => false);
    const kept = saved && Buffer.from(OoxmlPackage.read(fs.readFileSync(file)).read('word/media/model1.glb')).equals(cubeGlb({ texture: true }));
    check('word: the saved document keeps the model beside its picture, at the view it was left at', saved && kept, JSON.stringify({ saved, kept }));

    const complaints = await errorsIn(win);
    check('word: 3D models report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('word: the 3D model checks ran', false, err?.message || JSON.stringify(err));
  }
}
