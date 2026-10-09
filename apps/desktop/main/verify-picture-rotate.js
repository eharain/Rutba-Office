// Pictures: Rotate the picture (Ctrl+R) turns the file a quarter clockwise
// and saves it — a JPEG by its EXIF orientation, its pixels untouched, a PNG
// redrawn — and the stage shows the picture turned. Run alone with
// RUTBA_VERIFY_ONLY=picturerotate.
import fs from 'node:fs';
import path from 'node:path';
import { nativeImage } from 'electron';
import { readExif } from '@rutba/imaging/exif';
import { probeImage } from '@rutba/imaging/probe';
import { gradientPng } from './sample-picture.js';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the folder is made
 */
export async function verifyPictureRotate({ open, check, until, wait, press, errorsIn }, { dir }) {
  const root = path.join(dir, 'rotate');
  const jpg = path.join(root, 'a-photo.jpg');
  const png = path.join(root, 'b-drawing.png');
  try {
    fs.mkdirSync(root, { recursive: true });
    fs.writeFileSync(jpg, nativeImage.createFromBuffer(gradientPng(320, 200, [30, 90, 160], [230, 200, 120])).toJPEG(85));
    fs.writeFileSync(png, gradientPng(320, 200, [160, 40, 60], [240, 230, 200]));
    const pixelsBefore = fs.readFileSync(jpg).subarray(-64);
    const win = await open('pictures', jpg);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const stage = () => js(`(() => { const i = document.querySelector('.pv-image:not(.hidden)'); return i && i.complete ? { w: i.naturalWidth, h: i.naturalHeight, alt: i.alt } : null; })()`);
    await until(async () => (await stage())?.w === 320, 'the photo', 8000);
    await js(`document.body.focus(), 1`);
    await press(wc, 'r', { modifiers: ['control'] });
    const turned = await until(() => readExif(fs.readFileSync(jpg)).orientation === 6, 'the photo turned in its file', 5000).then(() => true).catch(() => false);
    const lossless = turned && Buffer.compare(fs.readFileSync(jpg).subarray(-64), pixelsBefore) === 0;
    const shown = await until(async () => { const s = await stage(); return s && s.w === 200 && s.h === 320; }, 'the photo shown turned', 5000).then(() => true).catch(() => false);
    check('pictures: Rotate the picture turns a JPEG a quarter clockwise in its file by its EXIF orientation, its pixels untouched, and shows it turned',
      turned && lossless && shown, JSON.stringify({ turned, lossless, shown, stage: await stage() }));

    // The PNG next door: redrawn turned.
    await press(wc, 'ArrowRight');
    await until(async () => (await stage())?.alt === 'b-drawing.png', 'the drawing', 5000).catch(() => {});
    await js(`document.querySelector('.pv-rotate-file')?.click(), 1`);
    const redrawn = await until(() => { const p = probeImage(fs.readFileSync(png)); return p?.width === 200 && p.height === 320; }, 'the drawing turned in its file', 5000).then(() => true).catch(() => false);
    const pngShown = await until(async () => { const s = await stage(); return s && s.w === 200 && s.h === 320; }, 'the drawing shown turned', 5000).then(() => true).catch(() => false);
    check('pictures: Rotate the picture redraws a PNG turned a quarter clockwise and saves it, and shows it turned', redrawn && pngShown, JSON.stringify({ redrawn, pngShown }));
    const complaints = await errorsIn(win);
    check('pictures: rotating pictures reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
    void wait;
  } catch (err) {
    check('pictures: the rotate checks ran', false, err?.message || String(err));
  }
}
