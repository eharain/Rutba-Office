// Windows' own stars and tags, read (2.21 D4): a photograph rated and
// tagged by Lightroom (XMP xmp:Rating and dc:subject) and one by Explorer's
// Properties (EXIF Rating and XPKeywords) show their stars and tags in
// Pictures' Details, under Pictures' own; Use takes them as Pictures' own.
// Run alone with RUTBA_VERIFY_ONLY=pictures-ratings.
import fs from 'node:fs';
import path from 'node:path';
import { noisyPicture } from './verify-image-edits.js';
import { withCamera, cameraBlock, xmpPacket } from './verify-exif-samples.js';

/**
 * @param {object} h the harness: open, check, until, wait, press, errorsIn
 * @param {{ dir: string }} args where the files are made
 */
export async function verifyPicturesRatings({ open, check, until, wait, press, errorsIn }, { dir }) {
  const root = path.join(dir, 'pictures-ratings');
  try {
    fs.mkdirSync(root, { recursive: true });
    const lightroom = path.join(root, 'a-lightroom.jpg');
    const explorer = path.join(root, 'b-explorer.jpg');
    fs.writeFileSync(lightroom, withCamera(noisyPicture(480, 320, 4).toJPEG(88), { xmp: xmpPacket({ rating: 4, keywords: ['Holiday', 'Beach'] }) }));
    fs.writeFileSync(explorer, withCamera(noisyPicture(480, 320, 8).toJPEG(88), { block: cameraBlock({ rating: 2, keywords: ['Family', 'Garden'] }) }));
    const win = await open('pictures', lightroom);
    const wc = win.webContents;
    const js = (code) => wc.executeJavaScript(code);
    const capture = async (name) => {
      if (!process.env.RUTBA_VERIFY_CAPTURE) return;
      wc.invalidate();
      await wait(700);
      fs.writeFileSync(path.join(process.env.RUTBA_VERIFY_CAPTURE, name), (await wc.capturePage()).toPNG());
    };
    const fileMarks = () => js(`(() => {
      const n = document.querySelector('[data-role="file-marks"]');
      if (!n) return null;
      return { rating: n.dataset.rating, tags: [...n.querySelectorAll('.pv-file-tag')].map((x) => x.textContent), stars: n.querySelectorAll('.pv-file-stars .on').length, use: Boolean(n.querySelector('[data-role="file-marks-use"]')), name: document.querySelector('.pv-detail-name')?.textContent };
    })()`);
    const own = () => js(`({ stars: document.querySelectorAll('.pv-stars .pv-star.on').length, tags: [...document.querySelectorAll('.pv-tags .pv-tag')].map((x) => x.dataset.tag) })`);

    const first = await until(async () => { const m = await fileMarks(); return m && m.name === 'a-lightroom.jpg' && m.rating === '4' ? m : null; }, 'the file\'s rating in Details', 8000).then(() => fileMarks()).catch(() => null);
    const ownBefore = await own();
    await capture('pictures-ratings-file.png');
    check('pictures: Details shows the rating and keywords Lightroom wrote into the file (XMP), beside Pictures\' own stars, read only',
      first?.stars === 4 && first.tags.join('|') === 'Holiday|Beach' && first.use && ownBefore.stars === 0 && ownBefore.tags.length === 0,
      JSON.stringify({ first, ownBefore }));

    // Use: the file's taken as Pictures' own; Use then has nothing left to take.
    await js(`document.querySelector('[data-role="file-marks-use"]')?.click(), 1`);
    const taken = await until(async () => { const o = await own(); return o.stars === 4 && o.tags.length === 2 ? o : null; }, 'the file\'s marks as Pictures\' own', 4000).then(() => own()).catch(() => null);
    const useGone = await until(async () => (await fileMarks())?.use === false, 'Use gone', 3000).then(() => true).catch(() => false);
    await capture('pictures-ratings-used.png');
    check('pictures: Use takes the file\'s rating and tags as Pictures\' own, and is then gone',
      taken?.tags.join('|') === 'Holiday|Beach' && useGone, JSON.stringify({ taken, useGone, after: await fileMarks() }));
    // Pictures' own let go again, so nothing is left behind for later checks.
    await js(`document.querySelector('.pv-star[data-star="4"]')?.click(), 1`);
    // One tag at a time: each removal draws the list again.
    for (let i = 0; i < 4 && (await own()).tags.length; i++) {
      await js(`document.querySelector('.pv-tags .pv-tag button')?.click(), 1`);
      await wait(200);
    }
    await until(async () => { const o = await own(); return o.stars === 0 && o.tags.length === 0; }, 'Pictures\' own cleared', 3000).catch(() => {});

    // The next picture: Explorer's EXIF Rating and XPKeywords.
    await js(`document.body.focus(), 1`);
    await press(wc, 'ArrowRight');
    const second = await until(async () => { const m = await fileMarks(); return m && m.name === 'b-explorer.jpg' && m.rating === '2' ? m : null; }, 'the second file\'s rating', 8000).then(() => fileMarks()).catch(() => null);
    await capture('pictures-ratings-explorer.png');
    check('pictures: Details shows the rating and tags Windows wrote into the file\'s EXIF (Rating, XPKeywords)',
      second?.stars === 2 && second.tags.join('|') === 'Family|Garden', JSON.stringify({ second, own: await own() }));
    const complaints = await errorsIn(win);
    check('pictures: the file\'s stars and tags report nothing', complaints.length === 0, complaints.join(' | ') || 'nothing reported');
  } catch (err) {
    check('pictures: the file rating checks ran', false, err?.message || String(err));
  }
}
