// Remove location (2.21 D2) and the camera's details kept on more exports
// (D3): a phone's photograph, with its camera, lens, date and place, exported
// from Image as WebP with Remove location pressed keeps the camera and has no
// GPS; as JPEG with neither Privacy button pressed keeps the place too; and
// Pictures' Batch edit, Remove location only, writes PNGs that carry the
// camera as eXIf and no place. Run alone with RUTBA_VERIFY_ONLY=exif-location.
import fs from 'node:fs';
import path from 'node:path';
import { readMetadata } from '@rutba/imaging/exif';
import { probeImage } from '@rutba/imaging/probe';
import { noisyPicture } from './verify-image-edits.js';
import { withCamera } from './verify-exif-samples.js';

/** What a written file says of itself: its format, size, camera and place. */
function readBack(file) {
  if (!fs.existsSync(file)) return null;
  const bytes = fs.readFileSync(file);
  const meta = readMetadata(bytes);
  const info = probeImage(bytes) || {};
  return {
    container: meta.container,
    width: info.width,
    height: info.height,
    model: meta.exif.model || null,
    lens: meta.exif.lens || null,
    taken: meta.exif.dateTimeOriginal || null,
    orientation: meta.exif.orientation ?? null,
    place: meta.exif.latitude != null,
    gpsFields: ['lat', 'lon', 'latRef', 'lonRef'].filter((k) => meta.exif[k] != null),
    problem: meta.problem,
  };
}

/**
 * @param {object} h the harness: open, check, until, wait, errorsIn
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyExifLocation({ open, check, until, wait, errorsIn }, { dir }) {
  const root = path.join(dir, 'exif-location');
  try {
    fs.mkdirSync(root, { recursive: true });
    const photo = path.join(root, 'harbour.jpg');
    fs.writeFileSync(photo, withCamera(noisyPicture(600, 400, 3).toJPEG(90)));
    const original = readBack(photo);

    // ── Image: Export as WebP with Remove location; then JPEG keeping everything.
    const win = await open('image', photo);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    const tab = (name) => js(`[...document.querySelectorAll('.rw-tab')].find((t) => t.textContent.trim() === ${JSON.stringify(name)})?.click(), 1`);
    const pressed = (role) => js(`document.querySelector('[data-role="${role}"]')?.getAttribute('aria-pressed') === 'true'`);
    await until(() => js(`document.querySelector('.im-canvas-wrap')?.dataset.result === '600x400'`), 'the photograph on the canvas', 15000);
    await tab('Export');
    await until(() => js(`Boolean(document.querySelector('[data-role="strip-location"]'))`), 'the Export tab', 5000);
    await js(`document.querySelector('[data-format="webp"]')?.click(), 1`);
    const stripFirst = await pressed('strip');
    await js(`document.querySelector('[data-role="strip-location"]')?.click(), 1`);
    await until(() => pressed('strip-location'), 'Remove location pressed', 3000).catch(() => {});
    const both = { strip: await pressed('strip'), location: await pressed('strip-location') };
    const tip = await js(`document.querySelector('[data-role="strip-location"]')?.dataset.tip || ''`);
    await capture('exif-location-image.png');
    const exportTo = async (name) => {
      const target = path.join(root, name);
      globalThis.__rutbaCheckDialogAnswers?.push({ kind: 'save', answer: target });
      await js(`[...document.querySelectorAll('.rw-groups .rw-btn')].find((n) => (n.dataset.tip || n.textContent || '').trim().startsWith('Export…') && !n.disabled)?.click(), 1`);
      await until(() => fs.existsSync(target) && fs.statSync(target).size > 0, `${name} written`, 20000).catch(() => {});
      await wait(300);
      return readBack(target);
    };
    const webp = await exportTo('harbour-no-place.webp');
    check('image: Export with Remove location writes a WebP that keeps the camera, lens and date, upright, with no GPS at all; Strip all is let go',
      stripFirst === true && both.location && !both.strip && tip.startsWith('Remove location')
        && webp?.container === 'webp' && webp.model === 'Test Phone 7' && webp.lens === 'Wide 26mm' && webp.taken === '2026:10:27 09:15:00'
        && webp.orientation === 1 && !webp.place && webp.gpsFields.length === 0 && !webp.problem && original?.place,
      JSON.stringify({ stripFirst, both, tip, webp, original }));

    // Neither pressed: everything kept, the place too, in a JPEG.
    await js(`document.querySelector('[data-role="strip-location"]')?.click(), 1`);
    await until(async () => !(await pressed('strip-location')), 'Remove location let go', 3000).catch(() => {});
    await js(`document.querySelector('[data-format="jpeg"]')?.click(), 1`);
    await wait(200);
    const jpeg = await exportTo('harbour-all.jpg');
    check('image: Export with neither Privacy button pressed keeps the camera and the place in a JPEG',
      jpeg?.container === 'jpeg' && jpeg.model === 'Test Phone 7' && jpeg.place && jpeg.orientation === 1,
      JSON.stringify(jpeg));
    // The choice back as it was (Strip all), for the checks after this.
    await js(`document.querySelector('[data-role="strip"]')?.click(), 1`);
    await until(() => pressed('strip'), 'Strip all pressed again', 3000).catch(() => {});
    const imageComplaints = await errorsIn(win);
    check('image: Remove location reports nothing', imageComplaints.length === 0, imageComplaints.join(' | ') || 'nothing reported');

    // ── Pictures: Batch edit, Remove location only, as PNG.
    const pw = await open('pictures', root);
    const pc = pw.webContents;
    const pjs = (code) => pc.executeJavaScript(code);
    const pcapture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      pc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await pc.capturePage()).toPNG());
    };
    await until(() => pjs(`document.querySelectorAll('.pv-tile:not(.folder)').length >= 1`), 'the folder', 8000).catch(() => {});
    await pjs(`[...document.querySelectorAll('.rw-btn')].find((n) => (n.dataset.tip || '').startsWith('Batch edit') && !n.disabled)?.click(), 1`);
    await until(() => pjs(`Boolean(document.querySelector('[data-role="batch-location"]'))`), 'the Batch edit dialog', 5000);
    await pjs(`[...document.querySelectorAll('.im-batch-head .rw-btn')].find((n) => n.textContent.trim() === 'None')?.click(), 1`);
    await wait(100);
    await pjs(`document.querySelector('.im-batch-file[data-file="harbour.jpg"] input')?.click(), 1`);
    await pjs(`(() => {
      const s = document.querySelector('[data-role="batch-format"]');
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(s, 'png');
      s.dispatchEvent(new Event('change', { bubbles: true }));
      const p = document.querySelector('[data-role="batch-pattern"]');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(p, '{name} for sharing');
      p.dispatchEvent(new Event('input', { bubbles: true }));
      document.querySelector('[data-role="batch-location"]').click();
      return 1;
    })()`);
    await until(() => pjs(`document.querySelector('[data-role="batch-location"]').checked && !document.querySelector('[data-role="batch-strip"]').checked`), 'Remove location only chosen', 3000).catch(() => {});
    const radios = await pjs(`[...document.querySelectorAll('[name="batch-metadata"]')].map((n) => n.dataset.role + ':' + n.checked)`);
    await pcapture('exif-location-batch.png');
    await pjs(`document.querySelector('[data-role="batch-run"]').click(), 1`);
    await until(() => pjs(`Boolean(document.querySelector('.im-batch-run[data-state="done"]'))`), 'the batch to finish', 30000).catch(() => {});
    const png = readBack(path.join(root, 'harbour for sharing.png'));
    check('pictures: Batch edit with Remove location only writes a PNG carrying the camera, lens and date as eXIf, upright, and no GPS',
      radios.includes('batch-location:true') && png?.container === 'png' && png.model === 'Test Phone 7' && png.lens === 'Wide 26mm'
        && png.taken === '2026:10:27 09:15:00' && png.orientation === 1 && !png.place && png.gpsFields.length === 0 && !png.problem,
      JSON.stringify({ radios, png }));
    await pjs(`document.querySelector('[data-role="batch-close"]')?.click(), 1`);
    await until(() => pjs(`!document.querySelector('.im-batch')`), 'the dialog closed', 4000).catch(() => {});
    // The choices are remembered for the next batch: forgotten, so a later check starts from the defaults.
    await pjs(`(() => { try { localStorage.removeItem('pictures.batch'); } catch {} return 1; })()`);
    const complaints = await errorsIn(pw);
    check('pictures: Remove location in Batch edit reports nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('image: the Remove location checks ran', false, err?.message || String(err));
  }
}
